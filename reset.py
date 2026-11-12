from pathlib import Path
import json
import xlwings as xw


# =========================================================
# PATHS
# =========================================================

workbook_path = Path(
    "../app/reports/server/lib/report/report.xlsx"
).resolve()

schema_path = Path(
    "../app/reports/server/lib/reset_schema/agodi_schema.json"
).resolve()


# =========================================================
# LOAD SCHEMA
# =========================================================

with open(schema_path, "r") as file:
    schema = json.load(file)


# =========================================================
# HELPERS
# =========================================================

def build_range(
    column,
    start_row,
    end_row
):
    return (
        f"{column}{start_row}:"
        f"{column}{end_row}"
    )


# =========================================================
# RESET OPERATIONS
# =========================================================

def execute_copy_range(
    sheet,
    operation
):

    source = operation["source"]

    source_range = build_range(
        source["column"],
        source["startRow"],
        source["endRow"]
    )

    source_values = (
        sheet
        .range(source_range)
        .value
    )

    for target in operation["targets"]:

        target_range = build_range(
            target["column"],
            target["startRow"],
            target["endRow"]
        )

        sheet.range(
            target_range
        ).value = source_values

        print(
            "RESET:",
            source_range,
            "->",
            target_range
        )


def execute_clear_range(
    sheet,
    operation
):

    range_data = (
        operation["range"]
    )

    excel_range = (
        f'{range_data["startColumn"]}'
        f'{range_data["startRow"]}:'
        f'{range_data["endColumn"]}'
        f'{range_data["endRow"]}'
    )

    sheet.range(
        excel_range
    ).clear_contents()

    print(
        "CLEARED:",
        excel_range
    )


# =========================================================
# RESET FLOW
# =========================================================

def execute_reset_phase(
    sheet,
    flow
):

    for operation in flow["operations"]:

        operation_type = (
            operation["type"]
        )

        if (
            operation_type
            == "copy_range"
        ):

            execute_copy_range(
                sheet,
                operation
            )

        elif (
            operation_type
            == "clear_range"
        ):

            execute_clear_range(
                sheet,
                operation
            )


# =========================================================
# MAIN
# =========================================================

app = xw.App(
    visible=False,
    add_book=False
)

app.display_alerts = False

app.screen_updating = False


try:

    wb = app.books.open(
        str(workbook_path)
    )

    print(
        f"READONLY: {wb.api.ReadOnly}"
    )

    latest_sheet = wb.sheets[-1]

    print(
        "TARGET SHEET:",
        latest_sheet.name
    )

    reset_flow = next(
        flow
        for flow in schema["flows"]
        if flow["name"]
        == "PRE_UPDATE_FLOW"
    )

    execute_reset_phase(
        latest_sheet,
        reset_flow
    )

    wb.save()

    print(
        "\nRESET COMPLETED"
    )

    wb.close()

finally:

    app.quit()