
import docx
import imghdr
import re

# 读取文档
doc = docx.Document("your_word_doc.docx")

# 遍历文档中的段落
for para in doc.paragraphs:
    # 如果段落中有图片
    if para._p.xpath('.//w:drawing'):
        # 获取图片的数据
        image_data = para._p.xpath('.//w:drawing')[0].xpath('.//wp:docPr')[0].xpath('.//a:blip')[0].attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
        image_data = doc.part.related_parts[image_data]._blob
        # 测试图片类型并打印
        image_type = imghdr.what(None, image_data)
        print(f"Image type is {image_type}")
        # 获取图片下面的第一行文本
        image_name = re.search(r'\n(.*)\n', para.text).group(1)
        # 打印图片名称
        print(f"Image name is {image_name}")
