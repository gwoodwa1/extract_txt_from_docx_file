from lxml import etree
import zipfile

"""
Module that extracts text from MS Word document (.docx) file.

"""

namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
para = namespace + "p"
text = namespace + "t"


def get_txt_from_docx(filepath) -> str:
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(filepath)

    xml_content = document.read("word/document.xml")
    document.close()
    tree = etree.fromstring(
        xml_content, parser=etree.XMLParser(huge_tree=True, recover=True)
    )
    paragraphs = []
    for paragraph in tree.getiterator(para):
        texts = [node.text for node in paragraph.getiterator(text)]
        if texts:
            paragraphs.append("".join(texts))
    return "\n\n".join(paragraphs)
  
