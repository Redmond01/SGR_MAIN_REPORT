


from pathlib import Path
import shutil
import xlwings as xw
import os
import stat


# =========================================================
# PATHS
# =========================================================

source_workbook_path = Path(
    "../app/reports/server/lib/report/MAY Agodi obx.xlsx"
).resolve()

generated_report_path = Path(
    "../app/reports/server/lib/report/report.xlsx"
).resolve()


# =========================================================
# CREATE RUNTIME REPORT FILE
# =========================================================

if generated_report_path.exists():
    generated_report_path.unlink()

shutil.copyfile(
    source_workbook_path,
    generated_report_path
)

os.chmod(
    generated_report_path,
    stat.S_IWRITE
)

print(
    "RUNTIME REPORT CREATED:",
    generated_report_path
)


# =========================================================
# EXCEL
# =========================================================

app = xw.App(
    visible=False,
    add_book=False
)

app.display_alerts = False

app.screen_updating = False


try:

    wb = app.books.open(
    str(generated_report_path),
    read_only=False,
    editable=True,
    ignore_read_only_recommended=True
    )

    print(
        "READ ONLY:",
        wb.api.ReadOnly
    )

    print(
        "BEFORE CLONE:",
        wb.sheets.count
    )

    # =====================================================
    # CLONE LAST SHEET
    # =====================================================

    last_sheet = wb.sheets[-1]

    last_sheet.copy(
        after=last_sheet
    )

    cloned_sheet = wb.sheets[-1]

    cloned_sheet.name = (
        "AGODI"
    )

    print(
        "AFTER CLONE:",
        wb.sheets.count
    )

    print(
        "NEW SHEET:",
        cloned_sheet.name
    )

    # =====================================================
    # SAVE
    # =====================================================

    wb.save()

    print(
        "REPORT WORKBOOK SAVED"
    )

    wb.close()

finally:

    app.quit()