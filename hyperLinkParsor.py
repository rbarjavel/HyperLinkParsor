import sys
import glob
import zipfile
import os

import pandas as pd
from lxml import etree

namespacesList = [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
]


def getDocxAsXML(documentPath: str) -> str:
    with open(documentPath, 'rb') as f:
        zip = zipfile.ZipFile(f)
        xml = zip.read('word/document.xml')
    return xml


def getXlsxAsXML(documentPath: str) -> any:
    with open(documentPath, 'rb') as f:
        zip = zipfile.ZipFile(f)
        xml = []

        for i in range(1, 50):
            try:
                xml.append(zip.read(f'xl/worksheets/sheet{i}.xml'))
            except:
                break
    return xml


def getPptxAsXML(documentPath: str) -> any:
    with open(documentPath, 'rb') as f:
        zip = zipfile.ZipFile(f)
        xml = []

        for i in range(1, 50):
            try:
                xml.append(zip.read(f'ppt/slides/slide{i}.xml'))
            except:
                break
    return xml


def hyperLinkBaliseDocx(data: str, pathFile: str) -> bool:
    tree = etree.fromstring(data)
    if tree.xpath("//w:hyperlink", namespaces={
        'w': namespacesList[0]
    }):
        print(f'\t{pathFile} -> found')
        return True
    return False


def hyperLinkBaliseXlsx(data: any, pathFile: str) -> bool:
    for d in data:
        tree = etree.fromstring(d)
        if tree.xpath("//ns:hyperlink", namespaces={
            'ns': namespacesList[1]
        }):
            print(f'\t{pathFile} -> found')
            return True
    return False


def hyperLinkBalisePptx(data: any, pathFile: str) -> bool:
    for d in data:
        tree = etree.fromstring(d)
        if tree.xpath("//a:hlinkClick", namespaces={
            'a': namespacesList[2]
        }):
            print(f'\t{pathFile} -> found')
            return True
    return False


def process(rootPath: str) -> None:
    filesPaths = []
    print(f'Search for hyperlinks in {rootPath} files ...\n')

    # Searching in all .docx files
    files = glob.glob(f'{rootPath}/**/*.docx', recursive=True)
    print('- Docx:')
    for file in files:
        if hyperLinkBaliseDocx(getDocxAsXML(file), file):
            filesPaths.append(os.path.abspath(file))

    # Searching in all .xlsx files
    files = glob.glob(f'{rootPath}/**/*.xlsx', recursive=True)
    print('- Xlsx:')
    for file in files:
        if hyperLinkBaliseXlsx(getXlsxAsXML(file), file):
            filesPaths.append(os.path.abspath(file))

    # Searching in all .pptx files
    files = glob.glob(f'{rootPath}/**/*.pptx', recursive=True)
    print('- Pptx:')
    for file in files:
        if hyperLinkBalisePptx(getPptxAsXML(file), file):
            filesPaths.append(os.path.abspath(file))

    return filesPaths


if __name__ == '__main__':
    files = process(sys.argv[1])

    df = pd.DataFrame({'files': files})
    df.to_csv('./result.csv')

    print(files)
