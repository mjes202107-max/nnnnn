import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment


def build_schedule(filename: str) -> None:
    # define the schedule data
    schedule = [
        {"時間": "09:00~09:30", "內容": "報到", "講者": "無"},
        {"時間": "09:30~09:40", "內容": "開場致詞", "講者": "王教授"},
        {"時間": "09:40~10:05", "內容": "邁向6G 的AI-RAN及O-RAN 趨勢介紹", "講者": "劉教授"},
        {"時間": "10:05~10:30", "內容": "下世代B5G/6G專網應用與未來趨勢", "講者": "陳教授"},
        {"時間": "10:30~10:50", "內容": "Break", "講者": "無"},
        {"時間": "10:50~11:20", "內容": "從O-RAN到AI-RAN 智慧通訊的節能應用", "講者": "教學團隊"},
        {"時間": "11:20~12:00", "內容": "O-RAN環境和各模組化功能介紹", "講者": "教學團隊"},
        {"時間": "12:00~13:30", "內容": "Lunch", "講者": "無"},
        {"時間": "13:30~14:00", "內容": "O-RAN 的市場應用案例", "講者": "教學團隊"},
        {"時間": "14:00~14:30", "內容": "O-RAN OSC環境建置教學", "講者": "教學團隊"},
        {"時間": "14:30~14:50", "內容": "Break", "講者": "無"},
        {"時間": "14:50~15:50", "內容": "O-RAN OSC第三方應用程式 xApps建置教學", "講者": "教學團隊"},
        {"時間": "15:50~16:30", "內容": "現場討論時間", "講者": "教學團隊"},
    ]

    df = pd.DataFrame(schedule)

    # write to excel
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    # post-process with openpyxl to apply merges/formatting
    wb = load_workbook(filename)
    ws = wb.active

    # center the header row
    for col in range(1, 4):
        ws.cell(row=1, column=col).alignment = Alignment(horizontal="center", vertical="center")

    max_row = ws.max_row

    # center all schedule cells (時間/內容/講者) and wrap text
    for row in range(2, max_row + 1):
        for col in range(1, 4):
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # adjust column widths based on content
    for col in range(1, 4):
        max_len = 0
        for row in range(1, max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        col_letter = ws.cell(row=1, column=col).column_letter
        ws.column_dimensions[col_letter].width = max_len + 2

    # merge cells for rows with special single entries
    for row in range(2, max_row + 1):
        content = ws.cell(row=row, column=2).value
        if content in ("報到", "Break", "Lunch"):
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
            ws.cell(row=row, column=2).alignment = Alignment(horizontal="center", vertical="center")


    wb.save(filename)


if __name__ == "__main__":
    build_schedule("nnnnn.xlsx")
