import sys
import glob
import zipfile
import os
import re

import pandas as pd
from lxml import etree


def get_version(data: str, doctype: str) -> str:
    if doctype == 'docx':
        version = re.search(r"(?:xmlns:w=\")([^ \"]*)", str(data))
    if doctype == 'xlsx':
        version = re.search(r"(?:xmlns=\")([^ \"]*)", str(data))
    if doctype == 'pptx':
        version = re.search(r"(?:xmlns:a=\")([^ \"]*)", str(data))

    return version.group(1)


def get_docx_as_xml(documentPath: str) -> str:
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


def hyperLinkBaliseDocx(data: str, pathFile: str, version: str) -> bool:
    tree = etree.fromstring(data)
    if tree.xpath("//w:hyperlink", namespaces={
        'w': version
    }):
        print(f'\t{pathFile} -> found')
        return True
    return False


def hyperLinkBaliseXlsx(data: any, pathFile: str, version: str) -> bool:
    for d in data:
        tree = etree.fromstring(d)
        if tree.xpath("//ns:hyperlink", namespaces={
            'ns': version
        }):
            print(f'\t{pathFile} -> found')
            return True
    return False


def hyperLinkBalisePptx(data: any, pathFile: str, version: str) -> bool:
    for d in data:
        tree = etree.fromstring(d)
        if tree.xpath("//a:hlinkClick", namespaces={
            'a': version
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
        data = getDocxAsXML(file)
        if hyperLinkBaliseDocx(data, file, getVersion(data, 'docx')):
            filesPaths.append(os.path.abspath(file))

    # Searching in all .xlsx files
    files = glob.glob(f'{rootPath}/**/*.xlsx', recursive=True)
    print('- Xlsx:')
    for file in files:
        data = getXlsxAsXML(file)
        if hyperLinkBaliseXlsx(data, file, getVersion(data, 'xlsx')):
            filesPaths.append(os.path.abspath(file))

    # Searching in all .pptx files
    files = glob.glob(f'{rootPath}/**/*.pptx', recursive=True)
    print('- Pptx:')
    for file in files:
        data = getPptxAsXML(file)
        if hyperLinkBalisePptx(data, file, getVersion(data, 'pptx')):
            filesPaths.append(os.path.abspath(file))

    return filesPaths


if __name__ == '__main__':
    files = process(sys.argv[1])

    if files:
        df = pd.DataFrame({'files': files})
        df.to_csv('./result.csv')
