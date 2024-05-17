import time
import glob
import calendar
import openpyxl

from functions import get_activities, write_to_namelist, write_to_report


# 月の最終日を取得
def get_days_in_month(year, month) -> int:
    if month == 2:
        return 29 if calendar.isleap(year) else 28
    elif month in [4, 6, 9, 11]:
        return 30
    return 31
        

def main() -> None:
    all_member_count = 90   # マークシートの人数
    year = 2024             # 年
    month = 6               # 月
    days = get_days_in_month(year, month)
    filenames = glob.glob("../scanned/*.png")
    settings = [            # 書き込み設定
        {
            "start_row": 7,
            "start_column": 8,
        },
        {
            "start_row": 7 + all_member_count // 3,
            "start_column": 8,
        },
        {
            "start_row": 7 + all_member_count // 3 * 2,
            "start_column": 8,
        },
        {
            "start_row": 7,
            "start_column": 8 + days // 2,
        },
        {
            "start_row": 7 + all_member_count // 3,
            "start_column": 8 + days // 2,
        },
        {
            "start_row": 7 + all_member_count // 3 * 2,
            "start_column": 8 + days // 2,
        },
    ]
    
    with open(f"../excel/{month}月元ファイル.xlsx", "rb") as f:
        try:
            workbook = openpyxl.load_workbook(f)    

            for filename, setting in zip(filenames, settings):
                activities = get_activities(filename, days // 2, all_member_count // 3)
                write_to_namelist(workbook, activities, setting["start_row"], setting["start_column"])

            start = time.time()
            
            write_to_report(
                workbook,
                days,
                all_member_count,
                settings[0]["start_row"],
                settings[0]["start_column"]
            )
            report_filename = f"../excel/{month}月活動報告書_音楽研究部.xlsx"
            workbook.save(report_filename)
            print(f"`{report_filename}.xlsx` を作成しました")
            
            end = time.time()
            print("所要時間: {:.3f}秒".format(end - start))
        except Exception as e:
            print("エラーが発生しました")
            print(e)
        
        
if __name__ == "__main__":
    main()