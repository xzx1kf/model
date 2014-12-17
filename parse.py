import xml.etree.ElementTree as ET
tree = ET.parse('data.xml')
root = tree.getroot()

for child in root:
    print child.tag, child.attrib
