import zipfile, os, tempfile
from lxml import etree

latin_to_georgian = {
    'a': 'ა', 'b': 'ბ', 'g': 'გ', 'd': 'დ', 'e': 'ე', 'v': 'ვ',
    'z': 'ზ', 'T': 'თ', 'i': 'ი', 'k': 'კ', 'l': 'ლ', 'm': 'მ',
    'n': 'ნ', 'o': 'ო', 'p': 'პ', 'J': 'ჟ', 'r': 'რ', 's': 'ს',
    't': 'ტ', 'u': 'უ', 'f': 'ფ', 'q': 'ქ', 'R': 'ღ', 'y': 'ყ',
    'S': 'შ', 'C': 'ჩ', 'c': 'ც', 'Z': 'ძ', 'w': 'წ', 'W': 'ჭ',
    'x': 'ხ', 'h': 'ჰ', 'j': 'ჯ'
}

def convert_text(text):
    return ''.join(latin_to_georgian.get(char, char) for char in text)

def process_xml_file(xml_path):
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(xml_path, parser)
    root = tree.getroot()
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for node in root.xpath('//w:t', namespaces=ns):
        if node.text:
            node.text = convert_text(node.text)
    tree.write(xml_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')

def convert_docx(input_path, output_path):
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        xml_targets = [
            'word/document.xml',
            *[f'word/header{i}.xml' for i in range(1, 4)],
            *[f'word/footer{i}.xml' for i in range(1, 4)],
            'word/footnotes.xml',
            'word/endnotes.xml'
        ]

        for target in xml_targets:
            full_path = os.path.join(tmpdir, target)
            if os.path.exists(full_path):
                process_xml_file(full_path)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
            for foldername, _, filenames in os.walk(tmpdir):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    arcname = os.path.relpath(filepath, tmpdir)
                    docx_zip.write(filepath, arcname)
