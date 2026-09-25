import csv
import os
import sqlite3
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

DB_PATH = Path(os.environ['DB_PATH'])
OUT_DIR = Path(os.environ['OUT_DIR'])
OUT_DIR.mkdir(parents=True, exist_ok=True)
pdf_path = OUT_DIR / 'informe_coladas_madres_productos_20260907.pdf'
csv_path = OUT_DIR / 'informe_coladas_madres_productos_20260907.csv'

con = sqlite3.connect(f'file:{DB_PATH}?mode=ro', uri=True)
con.row_factory = sqlite3.Row
labels = con.execute('''SELECT id, tipo_producto, calidad, medida, longitud, peso, colada, lote,
    cantidad, fecha, carbono, silicio, manganeso, hornero, supervisor, operador_ccm,
    aprobado, tipo_b, no_conforme, observaciones, created_at
    FROM etiqueta_calidad ORDER BY colada, tipo_producto, medida, id''').fetchall()
certs = {row['numero_colada'].strip(): row for row in con.execute('''SELECT numero_colada,
    numero_certificado, fecha_certificacion, limite_fluencia, limite_resistencia,
    alargamiento, doblado_cumple FROM certificado_calidad''').fetchall()}
con.close()

def clean(value):
    return '-' if value is None or value == '' else str(value).strip()

def kind(value):
    return {'Palanquillas': 'Palanquillas', 'AP500S': 'AP500 S', 'VarillasLisas': 'Varillas Lisas'}.get(value, value)

groups = defaultdict(list)
for row in labels:
    groups[clean(row['colada'])].append(row)
mothers = {key: rows for key, rows in groups.items() if any(row['tipo_producto'] == 'Palanquillas' for row in rows)}
orphan = {key: rows for key, rows in groups.items() if not any(row['tipo_producto'] == 'Palanquillas' for row in rows)}

with csv_path.open('w', newline='', encoding='utf-8-sig') as handle:
    writer = csv.writer(handle, delimiter=';')
    writer.writerow(['Colada madre', 'Certificado', 'Tipo producto', 'Calidad', 'Medida (mm)', 'Longitud (m)', 'Peso (kg)', 'Lote', 'Cantidad', 'Fecha', 'Carbono %', 'Silicio %', 'Manganeso %', 'Hornero', 'Supervisor', 'Operador CCM', 'Aprobado', 'Tipo B', 'No conforme', 'Observaciones', 'Registrado'])
    for colada in sorted(groups):
        cert = certs.get(colada)
        for row in groups[colada]:
            writer.writerow([colada, cert['numero_certificado'] if cert else '-', kind(row['tipo_producto']), clean(row['calidad']), clean(row['medida']), clean(row['longitud']), clean(row['peso']), clean(row['lote']), clean(row['cantidad']), clean(row['fecha']), clean(row['carbono']), clean(row['silicio']), clean(row['manganeso']), clean(row['hornero']), clean(row['supervisor']), clean(row['operador_ccm']), 'Sí' if row['aprobado'] else 'No', 'Sí' if row['tipo_b'] else 'No', 'Sí' if row['no_conforme'] else 'No', clean(row['observaciones']), clean(row['created_at'])])

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='ReportTitle', parent=styles['Title'], alignment=TA_CENTER, fontSize=16, leading=19, textColor=colors.HexColor('#172b78')))
styles.add(ParagraphStyle(name='Section', parent=styles['Heading2'], fontSize=11, leading=14, textColor=colors.HexColor('#155e75'), spaceBefore=8, spaceAfter=5))
styles.add(ParagraphStyle(name='SmallReport', parent=styles['BodyText'], fontSize=7, leading=9))
styles.add(ParagraphStyle(name='TinyReport', parent=styles['BodyText'], fontSize=6.2, leading=7.5))

def paragraph(value, style='SmallReport'):
    return Paragraph(str(value).replace('&', '&amp;'), styles[style])

def table(data, widths):
    result = Table(data, colWidths=widths, repeatRows=1)
    result.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), .35, colors.HexColor('#b8c0c9')), ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#172b78')), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#eef6f8')]), ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3)]))
    return result

def count(rows, product):
    return sum(row['tipo_producto'] == product for row in rows)

elements = [Paragraph('INFORME DE COLADAS MADRES Y PRODUCTOS CORRESPONDIENTES', styles['ReportTitle']), Spacer(1, 4), paragraph(f'Generado: {datetime.now():%d/%m/%Y %H:%M}'), paragraph(f'Coladas madres: {len(mothers)} | Etiquetas: {len(labels)} | Coladas sin Palanquilla madre: {len(orphan)}'), Spacer(1, 8), Paragraph('RESUMEN POR COLADA MADRE', styles['Section'])]
summary = [[paragraph(x, 'TinyReport') for x in ['Colada madre', 'Certificado', 'Palanquillas', 'AP500 S', 'Varillas lisas', 'Productos asociados']]]
for colada in sorted(mothers):
    rows = mothers[colada]
    cert = certs.get(colada)
    summary.append([paragraph(x, 'TinyReport') for x in [colada, cert['numero_certificado'] if cert else '-', count(rows, 'Palanquillas'), count(rows, 'AP500S'), count(rows, 'VarillasLisas'), count(rows, 'AP500S') + count(rows, 'VarillasLisas')]])
elements.append(table(summary, [27*mm, 25*mm, 23*mm, 18*mm, 23*mm, 27*mm]))
for colada in sorted(mothers):
    rows = mothers[colada]
    cert = certs.get(colada)
    chemistry = next(row for row in rows if row['tipo_producto'] == 'Palanquillas')
    elements += [Spacer(1, 8), Paragraph(f'COLADA MADRE: {colada}', styles['Section']), paragraph(f"Certificado: {cert['numero_certificado'] if cert else '-'} | Fecha: {clean(cert['fecha_certificacion']) if cert else '-'} | LF: {clean(cert['limite_fluencia']) if cert else '-'} MPa | LR: {clean(cert['limite_resistencia']) if cert else '-'} MPa | Alargamiento: {clean(cert['alargamiento']) if cert else '-'}"), paragraph(f"Química: Calidad {clean(chemistry['calidad'])} | C {clean(chemistry['carbono'])}% | Si {clean(chemistry['silicio'])}% | Mn {clean(chemistry['manganeso'])}% | Hornero {clean(chemistry['hornero'])} | Supervisor {clean(chemistry['supervisor'])} | Operador CCM {clean(chemistry['operador_ccm'])}")]
    detail = [[paragraph(x, 'TinyReport') for x in ['Producto', 'Calidad', 'Medida', 'Longitud', 'Peso', 'Lote', 'Cantidad', 'Fecha', 'Estado']]]
    for row in rows:
        status = 'No conforme' if row['no_conforme'] else 'Tipo B' if row['tipo_b'] else 'Aprobado' if row['aprobado'] else '-'
        detail.append([paragraph(x, 'TinyReport') for x in [kind(row['tipo_producto']), clean(row['calidad']), clean(row['medida']), clean(row['longitud']), clean(row['peso']), clean(row['lote']), clean(row['cantidad']), clean(row['fecha']), status]])
    elements.append(table(detail, [25*mm, 24*mm, 17*mm, 20*mm, 18*mm, 22*mm, 16*mm, 21*mm, 22*mm]))
if orphan:
    elements += [Spacer(1, 10), Paragraph('REGISTROS SIN PALANQUILLA MADRE ASOCIADA', styles['Section']), paragraph('Se mantienen separados porque no tienen una Palanquilla con la misma colada exacta.'), table([[paragraph(x, 'TinyReport') for x in ['Colada', 'Productos', 'Registros']]] + [[paragraph(x, 'TinyReport') for x in [colada, ', '.join(sorted({kind(row['tipo_producto']) for row in rows})), len(rows)]] for colada, rows in sorted(orphan.items())], [55*mm, 75*mm, 35*mm])]

doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, rightMargin=12*mm, leftMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm)
doc.build(elements)
print(f'PDF={pdf_path}')
print(f'CSV={csv_path}')
print(f'MOTHERS={len(mothers)} ORPHAN_COLADAS={len(orphan)} LABELS={len(labels)}')
