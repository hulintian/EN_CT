#!/usr/bin/env python3

from __future__ import annotations

import copy
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path("/home/hlt/Coures_Master_S2/EN_CT")
SOURCE = ROOT / "hws/哈尔滨工程大学-张世琦-通用PPT模板.pptx"
OUTPUT = ROOT / "hws/英国都铎王朝文化发展的历史进程（1485–1603）.pptx"

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

R_NS = "{%s}" % NS["r"]

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def shape_name(shape: ET.Element) -> str:
    node = shape.find("./p:nvSpPr/p:cNvPr", NS)
    return node.get("name", "") if node is not None else ""


def find_shape(slide_root: ET.Element, name: str, occurrence: int = 0) -> ET.Element:
    matches = [shape for shape in slide_root.findall(".//p:sp", NS) if shape_name(shape) == name]
    if occurrence >= len(matches):
        raise ValueError(f"shape '{name}' occurrence {occurrence} not found")
    return matches[occurrence]


def reset_shape_text(shape: ET.Element, paragraphs: list[str]) -> None:
    tx_body = shape.find("./p:txBody", NS)
    if tx_body is None:
        raise ValueError(f"shape '{shape_name(shape)}' has no text body")

    body_pr = tx_body.find("./a:bodyPr", NS)
    lst_style = tx_body.find("./a:lstStyle", NS)
    template_paragraph = tx_body.find("./a:p", NS)

    p_pr_template = None
    end_pr_template = None
    run_pr_template = None
    if template_paragraph is not None:
        p_pr_template = template_paragraph.find("./a:pPr", NS)
        end_pr_template = template_paragraph.find("./a:endParaRPr", NS)
        run = template_paragraph.find("./a:r", NS)
        if run is not None:
            run_pr_template = run.find("./a:rPr", NS)

    for child in list(tx_body):
        if child.tag not in {
            "{%s}bodyPr" % NS["a"],
            "{%s}lstStyle" % NS["a"],
        }:
            tx_body.remove(child)

    if not paragraphs:
        paragraphs = [""]

    for text in paragraphs:
        paragraph = ET.SubElement(tx_body, "{%s}p" % NS["a"])
        if p_pr_template is not None:
            paragraph.append(copy.deepcopy(p_pr_template))
        if text:
            run = ET.SubElement(paragraph, "{%s}r" % NS["a"])
            if run_pr_template is not None:
                run.append(copy.deepcopy(run_pr_template))
            else:
                ET.SubElement(run, "{%s}rPr" % NS["a"])
            text_node = ET.SubElement(run, "{%s}t" % NS["a"])
            if text.startswith(" ") or text.endswith(" "):
                text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            text_node.text = text
        if end_pr_template is not None:
            paragraph.append(copy.deepcopy(end_pr_template))

    if body_pr is None:
        tx_body.insert(0, ET.Element("{%s}bodyPr" % NS["a"]))
    if lst_style is None:
        tx_body.insert(1, ET.Element("{%s}lstStyle" % NS["a"]))


def update_shape(slide_root: ET.Element, name: str, paragraphs: list[str], occurrence: int = 0) -> None:
    reset_shape_text(find_shape(slide_root, name, occurrence), paragraphs)


def write_xml(path: Path, root: ET.Element) -> None:
    ET.ElementTree(root).write(path, encoding="UTF-8", xml_declaration=True)


def reorder_slides(workdir: Path, slide_order: list[int]) -> None:
    presentation_path = workdir / "ppt/presentation.xml"
    rels_path = workdir / "ppt/_rels/presentation.xml.rels"

    rels_root = ET.parse(rels_path).getroot()
    rel_to_slide = {}
    for rel in rels_root.findall("./{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
        rel_id = rel.get("Id")
        target = rel.get("Target", "")
        match = re.search(r"slide(\d+)\.xml$", target)
        if match:
            rel_to_slide[rel_id] = int(match.group(1))

    presentation_root = ET.parse(presentation_path).getroot()
    slide_list = presentation_root.find("./p:sldIdLst", NS)
    if slide_list is None:
        raise ValueError("presentation.xml is missing p:sldIdLst")

    slide_map = {}
    for slide in list(slide_list):
        rel_id = slide.get(f"{R_NS}id")
        if rel_id in rel_to_slide:
            slide_map[rel_to_slide[rel_id]] = slide

    for slide in list(slide_list):
        slide_list.remove(slide)
    for slide_number in slide_order:
        slide_list.append(slide_map[slide_number])

    ext_list = presentation_root.find("./p:extLst", NS)
    if ext_list is not None:
        for ext in list(ext_list):
            if ext.find("./p14:sectionLst", NS) is not None:
                ext_list.remove(ext)
        if not list(ext_list):
            presentation_root.remove(ext_list)

    write_xml(presentation_path, presentation_root)


def update_slides(workdir: Path) -> None:
    slides = {}
    for slide_number in [2, 4, 7, 9, 10, 12, 15, 23, 25, 41, 47, 48]:
        path = workdir / f"ppt/slides/slide{slide_number}.xml"
        slides[slide_number] = ET.parse(path).getroot()

    update_shape(slides[2], "标题 3", ["英国都铎王朝文化发展", "历史进程（1485-1603）"])
    update_shape(slides[2], "文本框 23", ["时间范围"])
    update_shape(slides[2], "文本框 24", ["1485-1603"])
    update_shape(slides[2], "文本框 17", ["汇报主题"])
    update_shape(slides[2], "文本框 18", ["王权·宗教·人文主义"])

    update_shape(slides[4], "文本框 26", ["历史脉络总览"])
    update_shape(slides[4], "文本框 27", ["转型动力与制度变迁"])
    update_shape(slides[4], "文本框 28", ["伊丽莎白时期鼎盛"])
    update_shape(slides[4], "文本框 29", ["历史意义与结论"])

    update_shape(slides[7], "文本框 11", ["历史脉络总览"])
    update_shape(slides[7], "文本框 12", ["都铎王朝文化演进"])
    update_shape(slides[7], "文本框 15", ["PART01"])

    update_shape(slides[12], "标题 3", ["都铎王朝文化演进时间轴"])
    update_shape(slides[12], "文本框 4", ["01"])
    update_shape(slides[12], "文本框 32", ["1485  都铎王朝建立，秩序重建"])
    update_shape(slides[12], "文本框 39", ["1509  亨利八世即位，宫廷文化活跃"])
    update_shape(slides[12], "文本框 40", ["1534  英国国教会建立，文化结构重塑"])
    update_shape(slides[12], "文本框 41", ["1547  爱德华六世推进改革制度化"])
    update_shape(slides[12], "文本框 42", ["1553  玛丽一世复辟，文化短暂中断"])
    update_shape(slides[12], "文本框 43", ["1558  伊丽莎白即位，整合与繁荣开启"])
    update_shape(slides[12], "文本框 45", ["1603  民族文化成熟，近代框架形成"])
    for name in ["文本框 47", "文本框 48", "文本框 49", "文本框 53", "文本框 60", "文本框 64"]:
        update_shape(slides[12], name, [""])

    update_shape(slides[15], "文本框 10", ["02"])
    update_shape(slides[15], "文本框 11", ["王权、宗教与语言"])
    update_shape(slides[15], "文本框 12", ["文化重构"])
    update_shape(slides[15], "文本框 15", ["PART02"])

    update_shape(slides[9], "标题 3", ["文化转型的三大驱动力"])
    update_shape(slides[9], "文本框 4", ["02"])
    update_shape(slides[9], "文本框 23", ["CULTURAL TRANSFORMATION"])
    update_shape(slides[9], "文本框 22", ["政治秩序"])
    update_shape(slides[9], "文本框 24", ["王权重新集中", "社会逐步稳定", "文化发展获得基础"])
    update_shape(slides[9], "文本框 25", ["宗教改革"])
    update_shape(slides[9], "文本框 26", ["脱离罗马教廷", "打破教会垄断", "文化走向世俗化"])
    update_shape(slides[9], "文本框 27", ["语言与人文"])
    update_shape(slides[9], "文本框 28", ["古典学习传播", "英语地位上升", "民族文化开始成形"])

    update_shape(slides[10], "标题 3", ["前四位君主带来的连续转折"])
    update_shape(slides[10], "文本框 4", ["02"])
    update_shape(slides[10], "文本框 48", ["亨利七世"])
    update_shape(slides[10], "文本框 49", ["结束内战", "稳定秩序", "人文主义萌芽"])
    update_shape(slides[10], "文本框 53", ["亨利八世"])
    update_shape(slides[10], "文本框 54", ["宗教断裂", "修道院解散", "宫廷文化兴起"])
    update_shape(slides[10], "文本框 58", ["爱德华六世"])
    update_shape(slides[10], "文本框 59", ["《公祷书》推行", "英语进入宗教实践", "教育继续扩散"])
    update_shape(slides[10], "文本框 63", ["玛丽一世"])
    update_shape(slides[10], "文本框 64", ["天主教复辟", "新教学者流亡", "文化发展短暂中断"])

    update_shape(slides[25], "文本框 10", ["03"])
    update_shape(slides[25], "文本框 11", ["戏剧、英语与扩张"])
    update_shape(slides[25], "文本框 12", ["文化鼎盛"])
    update_shape(slides[25], "文本框 15", ["PART03"])

    update_shape(slides[23], "标题 2", ["伊丽莎白一世时期的文化鼎盛"])
    update_shape(slides[23], "文本框 6", ["03"])
    update_shape(slides[23], "矩形 3", ["鼎盛时期"])
    update_shape(
        slides[23],
        "矩形 47",
        [
            "政治与宗教逐步稳定，英国迎来都铎文化最成熟的阶段。",
            "戏剧成为最具影响力的大众文化形式，莎士比亚与马洛推动文学高峰。",
            "英语完成规范化与成熟化，成为文学、行政与思想表达的核心载体。",
            "航海与海外扩张拓宽了世界观，也增强了英国文化的开放性与民族性。",
        ],
    )

    update_shape(slides[41], "标题 4", ["都铎文化留下的长期遗产"])
    update_shape(slides[41], "文本框 82", ["都铎时期奠定了近代英国文化的基本框架。"])
    update_shape(slides[41], "文本框 34", ["01"])
    update_shape(slides[41], "文本框 35", ["02"])
    update_shape(slides[41], "文本框 36", ["03"])
    update_shape(slides[41], "文本框 37", ["04"])
    update_shape(slides[41], "文本框 38", ["英语核心化"])
    update_shape(slides[41], "文本框 40", ["教育社会化"])
    update_shape(slides[41], "文本框 42", ["世俗文化成熟"])
    update_shape(slides[41], "文本框 44", ["民族认同形成"])
    update_shape(slides[41], "文本框 49", ["英语从行政与宗教语言", "成长为文学与思想载体"])
    update_shape(slides[41], "文本框 50", ["学校增加与识字提升", "知识传播突破精英阶层"])
    update_shape(slides[41], "文本框 51", ["文学与戏剧走向独立", "文化重心逐步脱离教会"])
    update_shape(slides[41], "文本框 52", ["开放而有本土特色的文化", "奠定近代英国文化框架"])
    update_shape(slides[41], "文本框 41", [""])

    update_shape(slides[47], "文本框 137", ["结论"])
    update_shape(
        slides[47],
        "文本框 139",
        ["都铎文化不是线性繁荣，而是在稳定、断裂与整合中，完成了英国从中世纪传统向近代民族文化的跨越。"],
    )

    update_shape(slides[48], "标题 3", ["谢谢聆听"])
    update_shape(slides[48], "文本框 14", ["英国都铎王朝文化发展（1485-1603）"])

    for slide_number, slide_root in slides.items():
        write_xml(workdir / f"ppt/slides/slide{slide_number}.xml", slide_root)


def repackage(workdir: Path, output_file: Path) -> None:
    if output_file.exists():
        output_file.unlink()

    with zipfile.ZipFile(output_file, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(workdir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(workdir))


def main() -> None:
    slide_order = [2, 4, 7, 12, 15, 9, 10, 25, 23, 41, 47, 48]

    with tempfile.TemporaryDirectory() as temp_dir:
        workdir = Path(temp_dir) / "pptx"
        with zipfile.ZipFile(SOURCE) as archive:
            archive.extractall(workdir)

        reorder_slides(workdir, slide_order)
        update_slides(workdir)
        repackage(workdir, OUTPUT)

    print(f"Created {OUTPUT}")


if __name__ == "__main__":
    main()
