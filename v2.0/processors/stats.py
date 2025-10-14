import pandas as pd
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet




def _read_any(path):
    if str(path).lower().endswith('.csv'):
        return pd.read_csv(path)
    return pd.read_excel(path)




def _save_any(df, output_path):
    if str(output_path).lower().endswith('.csv'):
        df.to_csv(output_path, index=False)
    else:
        df.to_excel(output_path, index=False)




def generate_stats(file_path, output_path, column=None):
    """
    Generate descriptive statistics.


    - If column is provided: compute count, mean, median, std, min, max, unique (where applicable)
    and save as a small table to output_path.
    - If column is None: compute df.describe(include='all').transpose() and save.


    Supports output_path .xlsx or .csv.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")


    df = _read_any(file_path)


    if column:
        if column not in df.columns:
            raise KeyError(f"Column '{column}' not found in input file.")
        ser = df[column]
        stats = {

        'count': ser.count(),
        'mean': ser.mean() if pd.api.types.is_numeric_dtype(ser) else None,
        'median': ser.median() if pd.api.types.is_numeric_dtype(ser) else None,
        'std': ser.std() if pd.api.types.is_numeric_dtype(ser) else None,
        'min': ser.min() if pd.api.types.is_numeric_dtype(ser) else None,
        'max': ser.max() if pd.api.types.is_numeric_dtype(ser) else None,
        'unique': ser.nunique()
        }
        out_df = pd.DataFrame(list(stats.items()), columns= ['metric', 'value'])
        _save_any(out_df, output_path)
        return output_path


# no column: full describe
    desc = df.describe(include='all').transpose()
# For non-numeric columns, include nunique
    if 'unique' not in desc.columns:
        desc['unique'] = df.nunique()
    _save_any(desc, output_path)
    return output_path




def export_pdf(file_path, output_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")


    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    # Path to DejaVuSans.ttf (must exist in your project folder)
    font_path = os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf")
    if not os.path.exists(font_path):
        # Try parent folder
        font_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "DejaVuSans.ttf")
    pdfmetrics.registerFont(TTFont('DejaVuSans', font_path))

    df = _read_any(file_path)
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()
    # Override font for all styles
    for style_name in styles.byName:
        styles[style_name].fontName = 'DejaVuSans'
    elements.append(Paragraph("Excel Toolkit - Report", styles['Heading1']))
    elements.append(Spacer(1, 8))

    # add first rows as table
    preview = df.head(30)
    data = [list(preview.columns.astype(str))]
    for _, row in preview.iterrows():
        data.append([str(x) for x in row.values])

    table = Table(data, hAlign='LEFT')
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'DejaVuSans'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    # basic stats for numeric columns
    num_cols = df.select_dtypes(include='number').columns
    if len(num_cols) > 0:
        elements.append(Paragraph('Numeric columns summary', styles['Heading2']))
        for col in num_cols:
            ser = df[col]
            stats = [
                [col, ''],
                ['metric', 'value'],
                ['count', str(ser.count())],
                ['mean', str(ser.mean())],
                ['std', str(ser.std())],
                ['min', str(ser.min())],
                ['max', str(ser.max())],
            ]
            tbl = Table(stats, hAlign='LEFT')
            tbl.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'DejaVuSans'),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.black)
            ]))
            elements.append(Paragraph(f'Column: {col}', styles['Normal']))
            elements.append(tbl)
            elements.append(Spacer(1, 8))

    doc.build(elements)
    return output_path