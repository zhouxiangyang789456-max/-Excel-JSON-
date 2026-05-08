"""用纯 Python 标准库创建一个测试 Excel 文件（不需要 openpyxl）"""
import zipfile
import os

OUTPUT = r"D:\git\-Excel-JSON-\ExcelToJSON\Assets\Excel\Item.xlsx"
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

# [Content_Types].xml
CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

# _rels/.rels
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

# xl/_rels/workbook.xml.rels
WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""

# xl/workbook.xml
WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Weapon" sheetId="1" r:id="rId1"/>
    <sheet name="Armor" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"""

# ===== Sheet 1: Weapon =====
# 数据：id, name, attack, hp, skills, price, quality, desc, icon
WEAPON_DATA = [
    ("id", "name", "attack", "hp", "skills", "price", "quality", "desc", "icon"),
    ("int", "string", "int", "int", "int[]", "int", "enum:Quality", "string", "res:Sprite"),
    ("ID", "名称", "攻击力", "生命值", "技能ID列表", "价格", "品质", "描述", "图标"),
    ("1001", "新手剑", "12", "100", "[1,2]", "100", "1", "一把普通的练习剑", "Ui/Icon/Item_001"),
    ("1002", "铁剑", "25", "0", "[1,3,5]", "500", "2", "坚固的铁剑", "Ui/Icon/Item_002"),
    ("1003", "秘银剑", "45", "0", "[2,4,6]", "2000", "3", "轻巧而锋利的秘银剑", "Ui/Icon/Item_003"),
    ("1004", "火焰剑", "60", "0", "[3,7,8]", "5000", "4", "燃烧着不灭之焰的魔剑", "Ui/Icon/Item_004"),
    ("1005", "龙牙剑", "85", "0", "[5,6,9]", "12000", "5", "用龙牙锻造的传说之剑", "Ui/Icon/Item_005"),
]

# ===== Sheet 2: Armor =====
ARMOR_DATA = [
    ("id", "name", "defense", "hp_bonus", "weight", "price", "quality"),
    ("int", "string", "int", "int", "float", "int", "enum:Quality"),
    ("ID", "名称", "防御力", "生命加成", "重量", "价格", "品质"),
    ("2001", "皮甲", "8", "20", "2.5", "200", "1"),
    ("2002", "锁子甲", "18", "40", "8.0", "800", "2"),
    ("2003", "板甲", "35", "80", "15.0", "3000", "3"),
]

# ===== Shared Strings =====
# 收集所有字符串值用于 sharedStrings.xml
all_strings = []
string_index = {}

def add_string(s):
    if s not in string_index:
        string_index[s] = len(all_strings)
        all_strings.append(s)
    return string_index[s]

def build_sheet_xml(data, first_row, first_col, last_row, last_col):
    """构建 sheet XML"""
    rows_xml = []
    for ri, row in enumerate(data):
        cells_xml = []
        for ci, val in enumerate(row):
            ref_col = chr(ord('A') + ci)
            ref = f"{ref_col}{ri + 1}"

            # 数值类型
            try:
                if val.startswith("[") and val.endswith("]"):
                    # 数组 → 存为字符串
                    idx = add_string(val)
                    cells_xml.append(
                        f'<c r="{ref}" t="s"><v>{idx}</v></c>')
                elif val in ("true", "false", "是", "否"):
                    idx = add_string(val)
                    cells_xml.append(
                        f'<c r="{ref}" t="s"><v>{idx}</v></c>')
                else:
                    # 尝试解析为数字
                    v = float(val)
                    if v == int(v):
                        cells_xml.append(
                            f'<c r="{ref}"><v>{int(v)}</v></c>')
                    else:
                        cells_xml.append(
                            f'<c r="{ref}"><v>{v}</v></c>')
            except (ValueError, AttributeError):
                # 字符串
                idx = add_string(val)
                cells_xml.append(
                    f'<c r="{ref}" t="s"><v>{idx}</v></c>')

        rows_xml.append(f'<row r="{ri + 1}">{"".join(cells_xml)}</row>')

    return f'<sheetData>{"".join(rows_xml)}</sheetData>'

SHEET1_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {build_sheet_xml(WEAPON_DATA, 0, 0, len(WEAPON_DATA)-1, len(WEAPON_DATA[0])-1)}
</worksheet>"""

SHEET2_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  {build_sheet_xml(ARMOR_DATA, 0, 0, len(ARMOR_DATA)-1, len(ARMOR_DATA[0])-1)}
</worksheet>"""

# sharedStrings.xml
ss_items = []
for s in all_strings:
    escaped = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    ss_items.append(f'<si><t>{escaped}</t></si>')

SHARED_STRINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="{len(all_strings)}" uniqueCount="{len(all_strings)}">
  {"".join(ss_items)}
</sst>"""

# ===== 打包为 .xlsx =====
with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr('[Content_Types].xml', CONTENT_TYPES)
    zf.writestr('_rels/.rels', RELS)
    zf.writestr('xl/workbook.xml', WORKBOOK)
    zf.writestr('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
    zf.writestr('xl/worksheets/sheet1.xml', SHEET1_XML)
    zf.writestr('xl/worksheets/sheet2.xml', SHEET2_XML)
    zf.writestr('xl/sharedStrings.xml', SHARED_STRINGS)

print(f"Created: {OUTPUT}")
print(f"  Sheet 1: Weapon ({len(WEAPON_DATA)} rows)")
print(f"  Sheet 2: Armor  ({len(ARMOR_DATA)} rows)")
print(f"  Shared strings: {len(all_strings)}")
