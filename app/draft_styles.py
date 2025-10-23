"""Draft: header style for Excel."""
def get_header_format(wb):
    bold = wb.add_format({"bold": True})
    return wb.add_format({"bold": True, "align": "center"})
