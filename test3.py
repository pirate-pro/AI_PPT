from copy import deepcopy
from io import BytesIO
import io, json, re
from typing import Dict, Any, List
import os
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.dml.color import RGBColor

app = FastAPI()

# --------------------- 占位符替换 --------------------- #
PATTERN = re.compile(r"\{\{(\w+)\}\}")

def replace_text_in_runs(para, mapping: Dict[str, Any]):
    """在保留格式的前提下替换文本（逐段处理）"""
    full_text = "".join(run.text for run in para.runs)
    if not PATTERN.search(full_text):
        return False

    new_text = replace_text(full_text, mapping)
    if new_text == full_text:
        return False

    # 关键改进：保留第一段格式，只替换文本内容
    if para.runs:
        # 保留第一个run的格式
        first_run = para.runs[0]
        font = first_run.font

        # 处理字体颜色
        if isinstance(font.color, RGBColor):
            color = font.color.rgb  # 正确访问 .rgb 属性
        else:
            color = None  # 如果是其他颜色类型，跳过

        # 清除所有run
        while para.runs:
            run = para.runs[0]
            if run._r.getparent() == para._p:
                para._p.remove(run._r)
            else:
                break

        # 用保留的格式创建新run
        new_run = para.add_run()
        new_run.text = new_text

        # 复制原始格式
        new_run.font.name = font.name
        new_run.font.size = font.size
        new_run.font.bold = font.bold
        new_run.font.italic = font.italic
        new_run.font.underline = font.underline
        if color:
            new_run.font.color.rgb = color  # 保留原始颜色

    else:
        para.text = new_text
    return True


def replace_text(text: str, mapping: Dict[str, Any]) -> str:
    def repl(m):
        key = m.group(1)
        m2 = re.match(r"^([A-Za-z_]+)(\d+)$", key)
        if m2:
            base, idx = m2.group(1), int(m2.group(2)) - 1
            if base in mapping and isinstance(mapping[base], list) and 0 <= idx < len(mapping[base]):
                return str(mapping[base][idx])
        return str(mapping.get(key, m.group(0)))

    return PATTERN.sub(repl, text)
def _process_group_shape(group, mapping: Dict[str, Any]):
    """递归处理组合形状，保留格式"""
    try:
        for shape in group.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    replace_text_in_runs(para, mapping)

            # 处理组合形状（类型值6）
            if shape.shape_type == 6:
                _process_group_shape(shape, mapping)
    except Exception as e:
        print(f"处理组合形状时出错: {e}")
def process_slide_text(slide, mapping: Dict[str, Any]):
    """处理幻灯片文本，保留原始格式"""
    print(f"\n===== 开始处理幻灯片文本 =====")

    try:
        shape_count = len(slide.shapes)
        print(f"当前形状数量: {shape_count}")
    except Exception as e:
        print(f"获取形状数量时出错: {e}")
        shape_count = 0

    # 刷新形状集合
    if shape_count == 0:
        print("警告：形状集合为空，尝试刷新...")
        try:
            slide.shapes._spTree = slide.element.xpath("./p:cSld/p:spTree")[0]
            shape_count = len(slide.shapes)
            print(f"刷新后形状数量: {shape_count}")
        except Exception as e:
            print(f"刷新形状集合失败: {e}")

    # 处理所有形状
    for i, shape in enumerate(slide.shapes):
        try:
            print(f"\n处理形状 {i + 1}/{len(slide.shapes)}")
            print(f"形状类型值: {shape.shape_type}, 是否有文本框: {shape.has_text_frame}")

            # 处理文本框（保留格式）
            if shape.has_text_frame:
                # 保留文本框的自动调整设置
                tf = shape.text_frame
                autofit = tf.word_wrap  # 保存自动换行设置
                for para in tf.paragraphs:
                    replace_text_in_runs(para, mapping)
                # 恢复自动调整设置
                tf.word_wrap = autofit
                tf.auto_size = tf.auto_size  # 确保自动调整字体大小功能有效

            # 处理组合形状
            if shape.shape_type == 6:
                _process_group_shape(shape, mapping)

        except Exception as e:
            print(f"处理形状 {i + 1} 时出错: {e}")

    print("===== 幻灯片文本处理结束 =====")

# --------------------- 复制工具函数 ------------------- #
def _copy_non_placeholder_shapes(src_cSld, dst_spTree):
    for shp in src_cSld.xpath("./p:spTree/*"):
        if shp.xpath(".//p:ph"):          # 占位符跳过
            continue
        dst_spTree.append(deepcopy(shp))


def _copy_bg(src_cSld, dst_cSld):
    # 尝试获取源幻灯片的背景
    bg = src_cSld.xpath("./p:bg")

    # 打印调试信息：源幻灯片是否包含背景
    print("源幻灯片背景存在:", bool(bg))

    if not bg:
        print("未找到背景，跳过复制操作")
        return

    # 打印调试信息：背景元素的内容
    print("找到的背景内容:", bg)

    # 删除目标幻灯片中的现有背景
    existing_bg = dst_cSld.xpath("./p:bg")
    if existing_bg:
        print("删除目标幻灯片中现有的背景")
        for old in existing_bg:
            dst_cSld.remove(old)
    else:
        print("目标幻灯片没有现有背景，跳过删除操作")

    # 复制背景到目标幻灯片
    dst_cSld.append(deepcopy(bg[0]))
    print("背景已复制到目标幻灯片")


def _fix_blip_rids(src_part, dst_slide):
    """
    确保 dst_slide 内所有 <a:blip> 引用的 rId 都存在；若缺失，则复制图片并更新 rId
    """
    # 查找目标幻灯片中的所有 <a:blip> 元素
    blips = dst_slide.element.xpath(".//a:blip")

    print(f"目标幻灯片中发现 {len(blips)} 张图片")

    if not blips:
        print("目标幻灯片中没有图片，跳过处理")
        return

    dst_rels = dst_slide.part._rels
    src_rels = src_part._rels

    # 遍历每个 <a:blip> 标签
    for i, blip in enumerate(blips):
        old_rid = blip.get(qn("r:embed"))

        # 打印调试信息：每个图片的 rId
        print(f"处理第 {i + 1} 张图片，rId 为: {old_rid}")

        # 如果 rId 为 None，生成一个新的唯一 rId
        if old_rid is None:
            old_rid = f"new_rId_{i}"
            print(f"图片 {i + 1} 的 rId 为 None，生成新 rId: {old_rid}")

        # 获取源幻灯片中的关系
        rel = src_rels.get(old_rid)

        if rel is None:
            print(f"未找到 rId {old_rid} 的关系，跳过")
            continue

        # 确保这是图像类型
        if rel.reltype != RT.IMAGE:
            print(f"rId {old_rid} 不是图像类型，跳过")
            continue

        # 获取图片数据
        img_blob = rel.target_part.blob
        if not img_blob:
            print(f"图片 {old_rid} 没有图像数据，跳过")
            continue

        print(f"找到图像 {old_rid}，正在复制到目标幻灯片")

        # 获取图片文件名（提取图片路径中的文件名）
        image_filename = os.path.basename(rel.target_part.partname)
        print(f"图片文件名: {image_filename}")

        # 复制图片并获取新的 rId
        _, new_rid = dst_slide.part.get_or_add_image_part(BytesIO(img_blob))

        # 更新目标幻灯片中的图片 rId
        blip.set(qn("r:embed"), new_rid)
        print(f"已更新 rId 为 {new_rid}，图像文件名: {image_filename}")

# --------------------- 幻灯片克隆 ------------------- #
def clone_slide(dst_prs: Presentation, src_slide):
    """三层克隆：幻灯片自身 + 版式 + 母版"""
    blank_layout = dst_prs.slide_layouts[6]        # 空白
    dst_slide = dst_prs.slides.add_slide(blank_layout)

    dst_cSld   = dst_slide.element.xpath("./p:cSld")[0]
    dst_spTree = dst_cSld.xpath("./p:spTree")[0]
    dst_cSld.remove(dst_spTree)                    # 去掉空白 spTree

    # ----- 幻灯片层 ----- #
    src_cSld = src_slide.element.xpath("./p:cSld")[0]
    dst_cSld.append(deepcopy(src_cSld.xpath("./p:spTree")[0]))
    _copy_bg(src_cSld, dst_cSld)

    # ----- 版式层 ----- #
    layout      = src_slide.slide_layout
    layout_cSld = layout.element.xpath("./p:cSld")[0]
    _copy_non_placeholder_shapes(layout_cSld, dst_cSld.xpath("./p:spTree")[0])
    _copy_bg(layout_cSld, dst_cSld)

    # ----- 母版层 ----- #
    master      = layout.slide_master
    master_cSld = master.element.xpath("./p:cSld")[0]
    _copy_non_placeholder_shapes(master_cSld, dst_cSld.xpath("./p:spTree")[0])
    _copy_bg(master_cSld, dst_cSld)

    # ----- 修正图片关系 ----- #
    _fix_blip_rids(src_slide.part, dst_slide)
    _fix_blip_rids(layout.part,    dst_slide)
    _fix_blip_rids(master.part,    dst_slide)

    # 返回新幻灯片索引
    slide_index = len(dst_prs.slides) - 1
    print(f"新幻灯片在目标演示文稿中的索引: {slide_index}")
    print("===== 幻灯片克隆结束 =====")
    return slide_index

# --------------------- FastAPI 入口 ------------------ #
@app.post("/merge_and_fill_slides")
async def merge_and_fill_slides(
    templates: List[UploadFile] = File(..., description="多个 PPT 模板"),
    json_data: str              = Form(..., description="合并与替换规则 JSON")
):
    try:
        # 1. 读取模板
        pres_list: List[Presentation] = []
        for tpl in templates:
            pres_list.append(Presentation(BytesIO(await tpl.read())))

        # 2. 解析 JSON
        cfg = json.loads(json_data)
        if "slides" not in cfg:
            raise HTTPException(400, "JSON 必须包含 slides 字段")

        # 3. 创建目标 PPT
        dest = Presentation()
        first = True

        # 4. 逐条配置合并
        for item in cfg["slides"]:
            t_idx = item["template_index"]
            s_idx = item["slide_idx"]
            repl  = item.get("replacements", {})

            if t_idx >= len(pres_list):
                raise HTTPException(400, f"模板索引 {t_idx} 超界")
            src_slide = pres_list[t_idx].slides[s_idx]

            # 同步尺寸
            if first:
                dest.slide_width  = pres_list[t_idx].slide_width
                dest.slide_height = pres_list[t_idx].slide_height
                first = False

            slide_index = clone_slide(dest, src_slide)
            # 获取新幻灯片并替换文本
            new_slide = dest.slides[slide_index]
            if repl:
                print(f"开始替换占位符，替换数据: {repl}")
                process_slide_text(new_slide, repl)
            else:
                print("无替换数据，跳过替换步骤")

        # 5. 输出文件
        buf = io.BytesIO()
        dest.save(buf)
        buf.seek(0)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": "attachment; filename=merged.pptx"}
        )

    except json.JSONDecodeError as e:
        raise HTTPException(400, f"JSON 解析失败：{e}")
    except Exception as e:
        raise HTTPException(500, f"处理失败：{e}")