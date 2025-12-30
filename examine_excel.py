import zipfile
import tempfile
import os
import sys
import lxml.etree as etree

def examine_excel(file_path):
    print(f"Examining Excel file: {file_path}")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        # 解压Excel文件
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        print("\n1. Checking file structure:")
        for root, dirs, files in os.walk(temp_dir):
            level = root.replace(temp_dir, '').count(os.sep)
            indent = ' ' * 2 * level
            print(f"{indent}{os.path.basename(root)}/")
            subindent = ' ' * 2 * (level + 1)
            for file in files:
                print(f"{subindent}{file}")
        
        # 检查cellimages.xml
        cellimages_path = os.path.join(temp_dir, 'xl', 'cellimages.xml')
        if os.path.exists(cellimages_path):
            print("\n2. Checking cellimages.xml:")
            with open(cellimages_path, 'r', encoding='utf-8') as f:
                content = f.read()
            print(content)
        
        # 检查工作表XML
        worksheet_path = os.path.join(temp_dir, 'xl', 'worksheets', 'sheet1.xml')
        if os.path.exists(worksheet_path):
            print("\n3. Checking sheet1.xml (B1 cell):")
            tree = etree.parse(worksheet_path)
            root = tree.getroot()
            
            # 找到sheetData和第一行
            sheet_data = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData')
            if sheet_data:
                row = sheet_data.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
                if row:
                    # 找到所有单元格
                    cells = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
                    if cells:
                        # 打印特定单元格（B1, C1, D1）
                        for cell_address in ['B1', 'C1', 'D1']:
                            for cell in cells:
                                if cell.get('r') == cell_address:
                                    print(f"\nCell {cell_address}:")
                                    # 打印单元格的XML
                                    cell_xml = etree.tostring(cell, encoding='unicode', pretty_print=True)
                                    print(cell_xml)
                                    break
        
        # 检查Content_Types.xml
        content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
        if os.path.exists(content_types_path):
            print("\n4. Checking [Content_Types].xml:")
            tree = etree.parse(content_types_path)
            root = tree.getroot()
            # 查看与cellimages相关的内容类型
            for override in root.findall('{http://schemas.openxmlformats.org/package/2006/content-types}Override'):
                if 'cellimage' in override.get('ContentType', ''):
                    print(etree.tostring(override, encoding='unicode', pretty_print=True))
        
        # 检查workbook.xml.rels
        workbook_rels_path = os.path.join(temp_dir, 'xl', '_rels', 'workbook.xml.rels')
        if os.path.exists(workbook_rels_path):
            print("\n5. Checking workbook.xml.rels:")
            tree = etree.parse(workbook_rels_path)
            root = tree.getroot()
            # 查看与cellimage相关的关系
            for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if 'cellImage' in rel.get('Type', ''):
                    print(etree.tostring(rel, encoding='unicode', pretty_print=True))
        
        # 检查cellimages.xml.rels
        cellimages_rels_path = os.path.join(temp_dir, 'xl', '_rels', 'cellimages.xml.rels')
        if os.path.exists(cellimages_rels_path):
            print("\n6. Checking cellimages.xml.rels:")
            with open(cellimages_rels_path, 'r', encoding='utf-8') as f:
                content = f.read()
            print(content)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python examine_excel.py <excel_file_path>")
        sys.exit(1)
    file_path = sys.argv[1]
    examine_excel(file_path)
