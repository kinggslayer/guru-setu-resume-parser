from sheets_utils import get_sheet

sheet = get_sheet()

sheet.append_row(
    ["TEST", "123", "resume.pdf"]
)

print("SUCCESS")