import xml.dom.minidom
import zipfile
import os

#通过rar删除保护
def xlsx_remove_protections(zipin, zipout):
    books = set()
    sheets = set()

    content_types = zipin.read('[Content_Types].xml')
    dom = xml.dom.minidom.parseString(content_types)
    type_map = {'application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.sheet.main+xml': books,
                'application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.worksheet+xml': sheets}
    root = dom.documentElement
    for node in root.childNodes:
        if node.hasAttribute('ContentType') and \
                node.getAttribute('ContentType') in type_map:
            assert(node.nodeName == 'Override' and
                   node.hasAttribute('PartName'))
            part_name = node.getAttribute('PartName')
            assert(part_name.startswith('/'))
            part_name = part_name[1:]
            type_map[node.getAttribute('ContentType')].add(part_name)

    for zinfo in zipin.infolist():
        content = zipin.read(zinfo)
        if zinfo.filename in books:
            dom = xml.dom.minidom.parseString(content)
            root = dom.documentElement
            protections = root.getElementsByTagName('workbookProtection')
            for protection in protections:
                root.removeChild(protection)
            protections = root.getElementsByTagName('fileSharing')
            for protection in protections:
                root.removeChild(protection)
            content = dom.toxml(encoding=dom.encoding)
        if zinfo.filename in sheets:
            dom = xml.dom.minidom.parseString(content)
            root = dom.documentElement
            protections = root.getElementsByTagName('sheetProtection')
            for protection in protections:
                root.removeChild(protection)
            content = dom.toxml(encoding=dom.encoding)
        zipout.writestr(zipfile.ZipInfo(zinfo.filename, zinfo.date_time),
                        content, compress_type=zipfile.ZIP_DEFLATED)

def foreachExcel (path):
    for root, dirs, files in os.walk(path):

        # root 表示当前正在访问的文件夹路径
        # dirs 表示该文件夹下的子目录名list
        # files 表示该文件夹下的文件list

        # 遍历文件
        for f in files:
            if(f[-5:].__str__() == '.xlsx'):
                excelFile = os.path.join(path,f)
                print(excelFile)
                with zipfile.ZipFile(excelFile, 'r') as zipin, \
                        zipfile.ZipFile(excelFile, 'w') as zipout:
                    xlsx_remove_protections(zipin, zipout)


if __name__ == '__main__':
    filePath = input("输入文件夹目录：")
    foreachExcel(filePath)