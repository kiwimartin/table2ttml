import pandas as pd
import argparse
import chardet
import xml.etree.ElementTree as ET

def xlsx_to_ttml(xlsx_file, ttml_file):

    df_timetable = pd.read_excel(xlsx_file, sheet_name='TIMETABLE')
    df_metadata = pd.read_excel(xlsx_file, sheet_name='METADATA')
    song_title = df_metadata.at[0, 'Title']

    ttml = ET.Element('tt', attrib={
        'xmlns': 'http://www.w3.org/ns/ttml',
        'xmlns:tts': 'http://www.w3.org/ns/ttml#styling'
    })

    head = ET.SubElement(ttml, 'head')
    metadata = ET.SubElement(head, 'metadata')
    title = ET.SubElement(metadata, 'title')
    title.text = song_title


    body = ET.SubElement(ttml, 'body')
    div = ET.SubElement(body, 'div')

    for index, row in df_timetable.iterrows():
        p = ET.SubElement(div, 'p', attrib={
            'begin': row['StartTime'],
            'end': row['EndTime']
        })
        p.text = str(row['Text'])

    tree = ET.ElementTree(ttml)
    ET.indent(tree, space="\t", level=0)
    tree.write(ttml_file, encoding='utf-8', xml_declaration=True)

def ttml_to_xlsx(ttml_file, xlsx_file):
    with open(ttml_file, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        encoding = result['encoding']

    tree = ET.ElementTree(ET.fromstring(raw_data.decode(encoding)))
    root = tree.getroot()

    namespaces = {'tt': 'http://www.w3.org/ns/ttml', 'ttm': 'http://www.w3.org/ns/ttml#metadata'}
    
    title_element = root.find('.//ttm:title', namespaces)
    song_title = title_element.text if title_element is not None else 'Unknown Title'

    rows = []
    for p in root.findall('.//{http://www.w3.org/ns/ttml}p'):
        start_time = p.attrib.get('begin')
        end_time = p.attrib.get('end')
        text = p.text
        rows.append({'StartTime': start_time, 'EndTime': end_time, 'Text': text})

    df_timetable = pd.DataFrame(rows)
    df_metadata = pd.DataFrame({'Title': [song_title]})

    with pd.ExcelWriter(xlsx_file) as writer:
        df_timetable.to_excel(writer, sheet_name='TIMETABLE', index=False)
        df_metadata.to_excel(writer, sheet_name='METADATA', index=False)



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert between XLSX and TTML files.')
    parser.add_argument('input_file', help='Input file path')
    parser.add_argument('output_file', help='Output file path')
    parser.add_argument('--reverse', action='store_true', help='Convert TTML to XLSX instead')

    args = parser.parse_args()

    if args.reverse:
        ttml_to_xlsx(args.input_file, args.output_file)
    else:
        xlsx_to_ttml(args.input_file, args.output_file)