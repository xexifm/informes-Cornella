"""simula-excel.py - Reescriu un .xlsx com ho fa l'Excel quan l'obres i el deses.

Us: python3 simula-excel.py entrada.xlsx sortida.xlsx

Per que: el repas que et baixes del mapa es un ZIP SENSE comprimir i amb els
textos "inline" (dins de cada cel.la). Si l'obres i el deses amb l'Excel, torna
COMPRIMIT (deflate), amb els textos a xl/sharedStrings.xml i, sovint, amb la
fulla amb un altre nom intern. El lector del mapa (JavaScript) i el del programa
(PowerShell) han d'entendre totes dues versions, i aqui no hi ha Excel.

openpyxl NO serveix per simular-ho: la 3.1 desa els textos inline, com nosaltres.
Nomes biblioteca estandard.
"""
import re
import sys
import zipfile
from xml.sax.saxutils import escape

entrada, sortida = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(entrada)
parts = {i.filename: zin.read(i.filename).decode('utf-8') for i in zin.infolist()}

FULLA_VELLA = 'xl/worksheets/sheet1.xml'
FULLA_NOVA = 'xl/worksheets/fulla_desada.xml'   # l'Excel no sempre la diu sheet1
fulla = parts.pop(FULLA_VELLA)

compartits = []


def a_compartit(m):
    text = re.sub(r'<[^>]+>', '', m.group(2))
    text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')
    if text not in compartits:
        compartits.append(text)
    return '<c r="%s" t="s"><v>%d</v></c>' % (m.group(1), compartits.index(text))


fulla = re.sub(r'<c r="([A-Z]+[0-9]+)" t="inlineStr"><is>(.*?)</is></c>', a_compartit, fulla)
parts[FULLA_NOVA] = fulla
parts['xl/sharedStrings.xml'] = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">' % (len(compartits), len(compartits))
    + ''.join('<si><t xml:space="preserve">%s</t></si>' % escape(t) for t in compartits)
    + '</sst>')
parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace(
    'Target="worksheets/sheet1.xml"', 'Target="/xl/worksheets/fulla_desada.xml"').replace(
    '</Relationships>',
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
    '/xl/worksheets/sheet1.xml', '/xl/worksheets/fulla_desada.xml').replace(
    '</Types>',
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')

with zipfile.ZipFile(sortida, 'w', zipfile.ZIP_DEFLATED) as zout:
    for nom, contingut in parts.items():
        zout.writestr(nom, contingut.encode('utf-8'))
