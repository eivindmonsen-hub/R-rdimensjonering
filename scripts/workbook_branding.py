"""Place the supplied AG logo inside existing headers without changing layout."""
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.writer.theme import theme_xml


LOGO = Path(__file__).resolve().parents[1] / "assets/branding/alfsen-gunderson-logo-white.png"


def brand_theme(xml):
    """Use the same profile for Excel's automatic chart and hyperlink colors."""
    namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
    root = ET.fromstring(xml)
    scheme = root.find(f".//{{{namespace}}}clrScheme")
    colors = {"dk1":"1A1C1B", "lt1":"FFFFFF", "dk2":"022B2A", "lt2":"F5FAF6",
              "accent1":"005149", "accent2":"40C1AC", "accent3":"FF4B1A",
              "accent4":"C0FFDE", "accent5":"022B2A", "accent6":"1A1C1B",
              "hlink":"005149", "folHlink":"005149"}
    for role, color in colors.items():
        value = scheme.find(f"{{{namespace}}}{role}")[0]
        value.tag = f"{{{namespace}}}srgbClr"
        value.attrib.clear()
        value.set("val", color)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def apply_branding(workbook):
    workbook.loaded_theme = brand_theme(workbook.loaded_theme or theme_xml.encode())
    for sheet in workbook:
        if sheet.sheet_state != "visible":
            continue
        title = next(r for r in sheet.merged_cells.ranges if r.min_col == 1 and r.min_row == 1 and r.max_row == 2)
        widths = [int(sheet.column_dimensions[get_column_letter(col)].width * 7 + 5) for col in range(1, title.max_col + 1)]
        logo = Image(str(LOGO))
        width = 210
        height = width * logo.height / logo.width
        x = sum(widths) - width - 16
        col = 0
        while x >= widths[col]:
            x -= widths[col]
            col += 1
        header_height = sum(sheet.row_dimensions[row].height or sheet.sheet_format.defaultRowHeight for row in (1,2)) * 96 / 72
        marker = AnchorMarker(col=col, colOff=pixels_to_EMU(x), row=0, rowOff=pixels_to_EMU((header_height-height)/2))
        logo.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(width), pixels_to_EMU(height)))
        sheet.add_image(logo)
        for chart in sheet._charts:
            for series in chart.series:
                series.graphicalProperties.solidFill = "005149"
                series.graphicalProperties.line.solidFill = "005149"
