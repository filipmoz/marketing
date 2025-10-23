"""Draft: bold format helper for Excel export."""
def bold_format(wb):
    bold = wb.add_format({"bold": True})
    return bold
