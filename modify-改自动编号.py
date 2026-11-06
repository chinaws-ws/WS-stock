import re
import xml.sax.saxutils as saxutils
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.opc.constants import RELATIONSHIP_TYPE as RT, CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.parts.numbering import NumberingPart

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 升级后的正则：支持 1. / (1) / 2.7.1 / 2.7.1.1 等格式
NUMBERING_PATTERN = re.compile(
    r"^\s*(?:[（(]\s*(\d+(?:\.\d+)*)\s*([)）])|(\d+(?:\.\d+)*)\s*([.、)）]|\s))"
)


def get_or_create_numbering_part(doc):
    """取现有 numbering part；文档没有则新建一个空的。"""
    try:
        return doc.part.numbering_part
    except (KeyError, NotImplementedError):
        pass
    # 去掉 XML 声明，避免 lxml 报错
    empty_xml = f'<w:numbering xmlns:w="{W_NS}"/>'
    part = NumberingPart(
        PackURI("/word/numbering.xml"),
        CT.WML_NUMBERING,
        parse_xml(empty_xml),
        doc.part.package,
    )
    doc.part.relate_to(part, RT.NUMBERING)
    return part


def add_custom_numbering(numbering_part, lvl_text, start_num):
    """动态创建一个单级自动编号定义，允许自定义前缀和起始值"""
    root = numbering_part.element
    existing_abs = [int(a.get(qn("w:abstractNumId"))) for a in root.findall(qn("w:abstractNum"))]
    existing_num = [int(n.get(qn("w:numId"))) for n in root.findall(qn("w:num"))]
    new_abs_id = (max(existing_abs) + 1) if existing_abs else 0
    new_num_id = (max(existing_num) + 1) if existing_num else 1

    # 转义特殊字符
    lvl_text_xml = saxutils.escape(lvl_text)

    abs_xml = f"""
    <w:abstractNum {nsdecls('w')} w:abstractNumId="{new_abs_id}">
      <w:multiLevelType w:val="singleLevel"/>
      <w:lvl w:ilvl="0">
        <w:start w:val="{start_num}"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="{lvl_text_xml}"/>
        <w:lvlJc w:val="left"/>
        <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      </w:lvl>
    </w:abstractNum>"""

    num_xml = f"""
    <w:num {nsdecls('w')} w:numId="{new_num_id}">
      <w:abstractNumId w:val="{new_abs_id}"/>
    </w:num>"""

    abs_el = parse_xml(abs_xml)
    num_el = parse_xml(num_xml)

    first_num = root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abs_el)
    else:
        root.append(abs_el)
    root.append(num_el)

    return new_num_id


def remove_prefix_chars(paragraph, n_chars):
    """从段落开头删除 n 个字符（跨 run 删除，保留格式）"""
    remaining = n_chars
    for run in paragraph.runs:
        if remaining <= 0:
            break
        text = run.text
        if len(text) <= remaining:
            remaining -= len(text)
            run.text = ""
        else:
            run.text = text[remaining:]
            remaining = 0


def set_numpr(paragraph, num_id, ilvl=0):
    """给段落设置 w:numPr，使其采用指定的自动编号。"""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = pPr.get_or_add_numPr()
    numPr.get_or_add_ilvl().val = ilvl
    numPr.get_or_add_numId().val = num_id


def convert_to_auto_numbering(input_path, output_path):
    doc = Document(input_path)
    numbering_part = get_or_create_numbering_part(doc)

    # 记录不同层级前缀的状态：key -> (num_id, 期待的下一个数字)
    prefix_state = {} 

    count = 0
    for para in doc.paragraphs:
        pPr = para._p.pPr
        if pPr is not None and pPr.numPr is not None:
            continue

        m = NUMBERING_PATTERN.match(para.text)
        if not m:
            continue

        # 提取数字和标点
        if m.group(1):
            num_str = m.group(1)
            end_mark = m.group(2)
            has_open_paren = True
        else:
            num_str = m.group(3)
            end_mark = m.group(4).strip()
            has_open_paren = False

        parts = num_str.split('.')
        current_num = int(parts[-1])

        # 拆分前缀，例如 2.7.1 -> prefix="2.7.", current_num=1
        if len(parts) == 1:
            prefix = ""
        else:
            prefix = ".".join(parts[:-1]) + "."

        # 生成编号格式，如 "2.7.%1."
        if has_open_paren:
            if end_mark == '）':
                lvl_text = f"（{prefix}%1）"
            else:
                lvl_text = f"({prefix}%1)"
        else:
            if end_mark == '' or end_mark == '.':
                lvl_text = f"{prefix}%1."
            elif end_mark == '、':
                lvl_text = f"{prefix}%1、"
            elif end_mark in [')', '）']:
                lvl_text = f"{prefix}%1{end_mark}"
            else:
                lvl_text = f"{prefix}%1."

        state_key = (prefix, lvl_text)

        # 判断是否延续上一个编号
        if state_key in prefix_state:
            num_id, next_exp = prefix_state[state_key]
            if current_num == next_exp:
                # 连贯：接着用原来的 numId
                prefix_state[state_key] = (num_id, next_exp + 1)
            else:
                # 不连贯（如跳过了 2.7.2 直接到 2.7.6）：新建一个序列
                num_id = add_custom_numbering(numbering_part, lvl_text, current_num)
                prefix_state[state_key] = (num_id, current_num + 1)
        else:
            # 首次遇到该前缀：新建序列
            num_id = add_custom_numbering(numbering_part, lvl_text, current_num)
            prefix_state[state_key] = (num_id, current_num + 1)

        remove_prefix_chars(para, m.end())
        set_numpr(para, num_id, ilvl=0)
        count += 1

    doc.save(output_path)
    print(f"成功处理！已将 {count} 个段落改为自动编号，输出文件：{output_path}")


if __name__ == "__main__":
    convert_to_auto_numbering(
        input_path="new_document.docx",
        output_path="new_document_auto_numbered.docx",
    )