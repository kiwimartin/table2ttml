import pandas as pd
import xml.etree.ElementTree as ET

def xlsx_to_ttml(xlsx_file, ttml_file):

    df = pd.read_excel(xlsx_file)


    ttml = ET.Element('tt', attrib={
        'xmlns': 'http://www.w3.org/ns/ttml',
        'xmlns:tts': 'http://www.w3.org/ns/ttml#styling'
    })


    body = ET.SubElement(ttml, 'body')
    div = ET.SubElement(body, 'div')


    for index, row in df.iterrows():
        p = ET.SubElement(div, 'p', attrib={
            'begin': row['StartTime'],
            'end': row['EndTime']
        })
        p.text = str(row['Text'])


    tree = ET.ElementTree(ttml)
    tree.write(ttml_file, encoding='utf-8', xml_declaration=True)



xlsx_to_ttml('input.xlsx', 'output.ttml')