import os
import sys
import uuid
import shutil
import zipfile
import tempfile
from PIL import Image
import lxml.etree as etree

# 确保图片目录存在
IMG_DIR = './img'
if not os.path.exists(IMG_DIR):
    os.makedirs(IMG_DIR)
    print(f"创建了目录: {IMG_DIR}")
    print("请将图片文件放入该目录后重新运行脚本")
    sys.exit(1)

# 获取图片文件列表
image_files = []
for file in os.listdir(IMG_DIR):
    if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
        image_files.append(os.path.join(IMG_DIR, file))

if not image_files:
    print(f"目录 {IMG_DIR} 中没有找到图片文件")
    sys.exit(1)

# 步骤1: 使用XlsxWriter创建基础Excel文件
def create_base_excel():
    print("创建基础Excel文件...")
    import xlsxwriter
    
    # 创建临时Excel文件
    temp_excel = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    temp_excel.close()
    
    workbook = xlsxwriter.Workbook(temp_excel.name)
    worksheet = workbook.add_worksheet()
    
    # 设置列宽和行高
    for i in range(len(image_files)):
        column_letter = chr(ord('B') + i)
        worksheet.set_column(f'{column_letter}:{column_letter}', 30)
    worksheet.set_row(0, 150)
    
    # 先在单元格中添加占位内容
    for i in range(len(image_files)):
        column_letter = chr(ord('B') + i)
        cell_address = f'{column_letter}1'
        worksheet.write(cell_address, f"=DISPIMG(\"PLACEHOLDER\",1)")
        print(f"已准备单元格: {cell_address}")
    
    workbook.close()
    return temp_excel.name

# 步骤2: 解压Excel文件
def unzip_excel(excel_path, extract_dir):
    print(f"解压Excel文件到: {extract_dir}")
    with zipfile.ZipFile(excel_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

# 步骤3: 复制图片到Excel的media目录
def copy_image_to_excel(image_path, extract_dir):
    media_dir = os.path.join(extract_dir, 'xl', 'media')
    if not os.path.exists(media_dir):
        os.makedirs(media_dir)
    
    # 生成唯一的图片文件名
    image_ext = os.path.splitext(image_path)[1]
    image_name = f'image_{uuid.uuid4().hex}{image_ext}'
    dest_path = os.path.join(media_dir, image_name)
    
    # 复制图片
    shutil.copy(image_path, dest_path)
    return image_name, dest_path

# 步骤4: 修改Content_Types.xml
def update_content_types(extract_dir):
    content_types_path = os.path.join(extract_dir, '[Content_Types].xml')
    parser = etree.XMLParser(load_dtd=False, no_network=True, recover=True)
    tree = etree.parse(content_types_path, parser)
    root = tree.getroot()
    
    # 添加cellimages.xml的内容类型
    if root.find('{http://schemas.openxmlformats.org/package/2006/content-types}Override[@PartName="/xl/cellimages.xml"]') is None:
        xml_string = '''<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/>'''
        new_elem = etree.fromstring(xml_string, parser)
        root.append(new_elem)
    
    # 添加jpeg内容类型（如果需要）
    if root.find('{http://schemas.openxmlformats.org/package/2006/content-types}Default[@Extension="jpeg"]') is None:
        xml_string = '''<Default Extension="jpeg" ContentType="image/jpeg"/>'''
        new_elem = etree.fromstring(xml_string, parser)
        root.append(new_elem)
    
    tree.write(content_types_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    print("已更新Content_Types.xml")

# 步骤5: 修改workbook.xml.rels
def update_workbook_rels(extract_dir):
    workbook_rels_path = os.path.join(extract_dir, 'xl', '_rels', 'workbook.xml.rels')
    parser = etree.XMLParser(load_dtd=False, no_network=True, recover=True)
    tree = etree.parse(workbook_rels_path, parser)
    root = tree.getroot()
    
    # 检查是否已存在cellimages.xml的关系
    if root.find('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship[@Target="cellimages.xml"]') is None:
        # 生成新的rId
        r_ids = []
        for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            r_id = rel.get('Id')
            if r_id.startswith('rId'):
                try:
                    num = int(r_id[3:])
                    r_ids.append(num)
                except ValueError:
                    pass
        
        new_r_id = f'rId{max(r_ids) + 1}' if r_ids else 'rId1'
        
        # 添加关系
        xml_string = f'''<Relationship Id="{new_r_id}" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/>'''  # 保持与正常文件相同的类型
        new_elem = etree.fromstring(xml_string, parser)
        root.append(new_elem)
    
    tree.write(workbook_rels_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    print("已更新workbook.xml.rels")

# 步骤6: 创建或更新cellimages.xml
def update_cellimages(extract_dir, image_name, image_path):
    cellimages_path = os.path.join(extract_dir, 'xl', 'cellimages.xml')
    
    # 如果文件不存在，创建它
    if not os.path.exists(cellimages_path):
        # 创建带有正确命名空间的根元素
        root = etree.Element(
            '{http://www.wps.cn/officeDocument/2017/etCustomData}cellImages',
            nsmap={
                'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'etc': 'http://www.wps.cn/officeDocument/2017/etCustomData'
            }
        )
        tree = etree.ElementTree(root)
        tree.write(cellimages_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    
    # 解析现有文件
    parser = etree.XMLParser(load_dtd=False, no_network=True, recover=True)
    tree = etree.parse(cellimages_path, parser)
    root = tree.getroot()
    
    # 生成唯一ID
    image_id = f'ID_{uuid.uuid4().hex}'
    
    # 获取现有rId的最大值（从cellimages.xml.rels中获取，而不是cellimages.xml）
    r_ids = []
    cellimages_rels_path = os.path.join(extract_dir, 'xl', '_rels', 'cellimages.xml.rels')
    if os.path.exists(cellimages_rels_path):
        rels_tree = etree.parse(cellimages_rels_path, parser)
        rels_root = rels_tree.getroot()
        for rel in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            r_id = rel.get('Id')
            if r_id and r_id.startswith('rId'):
                try:
                    num = int(r_id[3:])
                    r_ids.append(num)
                except ValueError:
                    pass
    
    # 也检查cellimages.xml中的r:embed
    for img in root.findall('{http://www.wps.cn/officeDocument/2017/etCustomData}cellImage'):
        pic = img.find('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}pic')
        if pic is not None:
            blipFill = pic.find('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}blipFill')
            if blipFill is not None:
                blip = blipFill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                if blip is not None:
                    r_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if r_id and r_id.startswith('rId'):
                        try:
                            num = int(r_id[3:])
                            r_ids.append(num)
                        except ValueError:
                            pass
    
    new_r_id = f'rId{max(r_ids) + 1}' if r_ids else 'rId1'
    
    # 获取图片尺寸
    with Image.open(image_path) as img:
        width, height = img.size
    
    # 将像素转换为EMU（Excel内部单位，1像素=9525 EMU）
    width_emu = width * 9525
    height_emu = height * 9525
    
    # 添加新的cellImage元素，使用正确的WPS格式
    cell_image = etree.SubElement(
        root,
        '{http://www.wps.cn/officeDocument/2017/etCustomData}cellImage'
    )
    
    # 添加pic元素
    pic = etree.SubElement(
        cell_image,
        '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}pic'
    )
    
    # 添加nvPicPr元素
    nv_pic_pr = etree.SubElement(
        pic,
        '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}nvPicPr'
    )
    
    # 添加cNvPr元素
    c_nv_pr = etree.SubElement(
        nv_pic_pr,
        '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cNvPr',
        id='2',  # 这里可以考虑递增ID
        name=image_id
    )
    
    # 添加cNvPicPr元素
    c_nv_pic_pr = etree.SubElement(
        nv_pic_pr,
        '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cNvPicPr'
    )
    
    # 添加picLocks元素
    pic_locks = etree.SubElement(
        c_nv_pic_pr,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}picLocks',
        noChangeAspect='1'
    )
    
    # 添加blipFill元素
    blip_fill = etree.SubElement(
        pic,
        '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}blipFill'
    )
    
    # 添加blip元素
    blip = etree.SubElement(
        blip_fill,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}blip',
        **{'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed': new_r_id}
    )
    
    # 添加stretch元素
    stretch = etree.SubElement(
        blip_fill,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}stretch'
    )
    
    # 添加fillRect元素
    fill_rect = etree.SubElement(
        stretch,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect'
    )
    
    # 添加spPr元素
    sp_pr = etree.SubElement(
        pic,
        '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}spPr'
    )
    
    # 添加xfrm元素
    xfrm = etree.SubElement(
        sp_pr,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm'
    )
    
    # 添加off元素
    off = etree.SubElement(
        xfrm,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}off',
        x='0',
        y='0'
    )
    
    # 添加ext元素
    ext = etree.SubElement(
        xfrm,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}ext',
        cx=str(width_emu),
        cy=str(height_emu)
    )
    
    # 添加prstGeom元素
    prst_geom = etree.SubElement(
        sp_pr,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom',
        prst='rect'
    )
    
    # 添加avLst元素
    av_lst = etree.SubElement(
        prst_geom,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}avLst'
    )
    
    # 添加noFill元素
    no_fill = etree.SubElement(
        sp_pr,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}noFill'
    )
    
    # 添加ln元素
    ln = etree.SubElement(
        sp_pr,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}ln',
        w='9525'
    )
    
    # 添加ln的noFill元素
    ln_no_fill = etree.SubElement(
        ln,
        '{http://schemas.openxmlformats.org/drawingml/2006/main}noFill'
    )
    
    tree.write(cellimages_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    print(f"已更新cellimages.xml，添加图片ID: {image_id}, rId: {new_r_id}")
    return image_id, new_r_id

# 步骤7: 更新cellimages.xml.rels
def update_cellimages_rels(extract_dir, image_name, r_id):
    cellimages_rels_dir = os.path.join(extract_dir, 'xl', '_rels')
    if not os.path.exists(cellimages_rels_dir):
        os.makedirs(cellimages_rels_dir)
    
    cellimages_rels_path = os.path.join(cellimages_rels_dir, 'cellimages.xml.rels')
    
    # 如果文件不存在，创建它
    if not os.path.exists(cellimages_rels_path):
        root = etree.Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationships')
        tree = etree.ElementTree(root)
        tree.write(cellimages_rels_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    
    # 解析现有文件
    parser = etree.XMLParser(load_dtd=False, no_network=True, recover=True)
    tree = etree.parse(cellimages_rels_path, parser)
    root = tree.getroot()
    
    # 添加新的关系
    xml_string = f'''<Relationship Id="{r_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{image_name}"/>'''
    new_elem = etree.fromstring(xml_string, parser)
    root.append(new_elem)
    
    tree.write(cellimages_rels_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    print(f"已更新cellimages.xml.rels，添加图片: {image_name}")

# 步骤8: 更新工作表XML，替换占位符
def update_worksheet(extract_dir, image_id, column_index):
    worksheet_path = os.path.join(extract_dir, 'xl', 'worksheets', 'sheet1.xml')
    parser = etree.XMLParser(load_dtd=False, no_network=True, recover=True)
    tree = etree.parse(worksheet_path, parser)
    root = tree.getroot()
    
    # 找到sheetData元素
    sheet_data = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData')
    if sheet_data is None:
        return
    
    # 找到第一行
    row = sheet_data.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
    if row is None:
        return
    
    # 找到对应的单元格（B1, C1, D1等）
    column_letter = chr(ord('B') + column_index)
    cell_address = f'{column_letter}1'
    
    cells = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
    cell = None
    for c in cells:
        if c.get('r') == cell_address:
            cell = c
            break
    
    if cell is None:
        return
    
    # 确保单元格的r属性值正确
    cell.set('r', cell_address)
    
    # 为DISPIMG函数设置单元格类型为str（与正常文件保持一致）
    cell.set('t', 'str')
    
    # 更新f元素（公式）
    f_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f')
    if f_elem is not None:
        f_elem.text = f'_xlfn.DISPIMG("{image_id}",1)'
        # 设置公式属性
        f_elem.set('t', 'shared')
        f_elem.set('ref', f'{chr(ord("B") + column_index)}1')
    
    # 更新v元素（值）
    v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
    if v_elem is not None:
        # 对于DISPIMG函数，v元素应该是公式本身，与正常文件保持一致
        v_elem.text = f'=DISPIMG("{image_id}",1)'
    
    tree.write(worksheet_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    column_letter = chr(ord('B') + column_index)
    print(f"已更新工作表，在单元格 {column_letter}1 设置图片ID: {image_id}")

# 步骤9: 重新压缩Excel文件
def zip_excel(extract_dir, output_path):
    print(f"重新压缩Excel文件到: {output_path}")
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root_dir, _, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, extract_dir)
                zip_ref.write(file_path, arcname)

# 主函数
def main():
    # 创建临时目录
    with tempfile.TemporaryDirectory() as temp_dir:
        # 1. 创建基础Excel文件
        base_excel = create_base_excel()
        
        try:
            # 2. 解压Excel文件
            extract_dir = os.path.join(temp_dir, 'excel_extract')
            os.makedirs(extract_dir)
            unzip_excel(base_excel, extract_dir)
            
            # 3. 更新Content_Types.xml和workbook.xml.rels
            update_content_types(extract_dir)
            update_workbook_rels(extract_dir)
            
            # 4. 处理每个图片
            for i, image_file in enumerate(image_files):
                print(f"\n处理图片 {i+1}/{len(image_files)}: {image_file}")
                
                # 复制图片到Excel的media目录
                image_name, dest_path = copy_image_to_excel(image_file, extract_dir)
                
                # 更新cellimages.xml
                image_id, r_id = update_cellimages(extract_dir, image_name, dest_path)
                
                # 更新cellimages.xml.rels
                update_cellimages_rels(extract_dir, image_name, r_id)
                
                # 更新工作表
                update_worksheet(extract_dir, image_id, i)
            
            # 5. 重新压缩Excel文件
            output_excel = 'images_wps_embedded.xlsx'
            zip_excel(extract_dir, output_excel)
            
            print(f"\n\n✅ 成功生成WPS内嵌图片Excel文件: {output_excel}")
            print(f"共处理 {len(image_files)} 张图片")
            print(f"图片已内嵌到单元格 B1-{chr(ord('B') + len(image_files) - 1)}1")
            print(f"使用WPS打开文件，图片将真正内嵌在单元格中，无法移动")
            
        finally:
            # 清理临时文件
            if os.path.exists(base_excel):
                os.remove(base_excel)

if __name__ == "__main__":
    main()