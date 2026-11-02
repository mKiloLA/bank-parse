import csv
import os
import re
from collections import Counter, defaultdict
from datetime import date, timedelta
from typing import List, Optional, Sequence

import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .CategoryTargetHandler import CategoryTargetHandler
from .data.DataTypes import Bank, Category, category_name
from .data.Statement import Statement


class GenerateOutput:
    NAVY = "18324F"
    BLUE = "286090"
    TEAL = "218C87"
    LIGHT_BLUE = "EAF1F8"
    LIGHT_TEAL = "E7F4F2"
    LIGHT_RED = "FCE8E6"
    LIGHT_AMBER = "FFF4D6"
    GRID = "D7E0E9"
    TEXT = "25364A"
    MUTED = "64758A"
    CURRENCY_FORMAT = '$#,##0.00;[Red]($#,##0.00);$0.00'
    PERCENT_FORMAT = '0.0%;[Red](0.0%);0.0%'

    @staticmethod
    def output_filename(statements: Sequence[Statement]) -> str:
        month_counts = Counter(
            (transaction.date.year, transaction.date.month)
            for statement in statements
            for transaction in statement.transactions
        )
        if not month_counts:
            return "CombinedStatement.xlsx"
        year, month = max(
            month_counts, key=lambda item: (month_counts[item], item)
        )
        return f"CombinedStatement_{year:04d}-{month:02d}.xlsx"

    @staticmethod
    def default_output_path(statements: Sequence[Statement]) -> str:
        return os.path.join(
            GenerateOutput.default_output_directory(),
            GenerateOutput.output_filename(statements),
        )

    @staticmethod
    def default_output_directory() -> str:
        return os.path.expanduser("~/StatementParser")

    @staticmethod
    def _unique_sheet_title(base_name: str, existing_names: set[str]) -> str:
        safe_name = re.sub(r"[\[\]:*?/\\]", "_", base_name).strip(" '")
        safe_name = (safe_name or "Statement")[:31]
        candidate = safe_name
        suffix_number = 2
        while candidate.casefold() in existing_names:
            suffix = f" ({suffix_number})"
            candidate = f"{safe_name[:31 - len(suffix)]}{suffix}"
            suffix_number += 1
        existing_names.add(candidate.casefold())
        return candidate

    @staticmethod
    def _add_source_csv_sheets(
        workbook, source_files: Sequence[tuple[Bank, str]]
    ) -> None:
        existing_names = {name.casefold() for name in workbook.sheetnames}
        for bank, file_path in source_files:
            base_name = (
                bank.value
                if bank != Bank.NONE
                else os.path.splitext(os.path.basename(file_path))[0]
            )
            sheet = workbook.create_sheet(
                GenerateOutput._unique_sheet_title(base_name, existing_names)
            )
            sheet.sheet_view.showGridLines = False
            sheet.freeze_panes = "A2"
            with open(file_path, newline="", encoding="utf-8-sig") as file:
                sample = file.read(8192)
                file.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
                except csv.Error:
                    dialect = csv.excel

                for row_number, row in enumerate(csv.reader(file, dialect), start=1):
                    for column_number, value in enumerate(row, start=1):
                        cell = sheet.cell(
                            row=row_number, column=column_number, value=value
                        )
                        cell.data_type = "s"
                        if row_number == 1:
                            cell.font = Font(bold=True, color="FFFFFF")
                            cell.fill = PatternFill(
                                fill_type="solid", fgColor=GenerateOutput.NAVY
                            )
                if sheet.max_row:
                    sheet.auto_filter.ref = sheet.dimensions
                for column in range(1, sheet.max_column + 1):
                    letter = get_column_letter(column)
                    max_length = max(
                        (
                            len(str(sheet.cell(row, column).value or ""))
                            for row in range(1, sheet.max_row + 1)
                        ),
                        default=0,
                    )
                    sheet.column_dimensions[letter].width = min(
                        max(max_length + 2, 12), 42
                    )

    @staticmethod
    def _style_sheet(sheet) -> None:
        sheet.sheet_view.showGridLines = False
        sheet.sheet_view.zoomScale = 160
        sheet.sheet_view.zoomScaleNormal = 160
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.outlinePr.summaryBelow = True

    @staticmethod
    def _write_title(sheet, cell_range: str, title: str) -> None:
        min_col, min_row, max_col, _ = openpyxl.utils.range_boundaries(cell_range)
        fill = PatternFill(fill_type="solid", fgColor=GenerateOutput.NAVY)
        for column in range(min_col, max_col + 1):
            cell = sheet.cell(row=min_row, column=column)
            cell.fill = fill
        sheet.merge_cells(cell_range)
        cell = sheet.cell(row=min_row, column=min_col, value=title)
        cell.font = Font(name="Aptos Display", size=22, bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center")
        sheet.row_dimensions[min_row].height = 36

    @staticmethod
    def _write_section_table(
        sheet,
        title: str,
        headers: Sequence[str],
        rows: Sequence[Sequence],
        start_row: int,
        start_column: int,
        table_name: str,
    ) -> int:
        end_column = start_column + len(headers) - 1
        title_range = (
            f"{get_column_letter(start_column)}{start_row}:"
            f"{get_column_letter(end_column)}{start_row}"
        )
        for column in range(start_column, end_column + 1):
            cell = sheet.cell(row=start_row, column=column)
            cell.fill = PatternFill(fill_type="solid", fgColor=GenerateOutput.BLUE)
        sheet.merge_cells(title_range)
        title_cell = sheet.cell(row=start_row, column=start_column, value=title)
        title_cell.font = Font(name="Aptos", size=12, bold=True, color="FFFFFF")
        title_cell.alignment = Alignment(vertical="center")
        sheet.row_dimensions[start_row].height = 23

        header_row = start_row + 1
        for column, header in enumerate(headers, start=start_column):
            cell = sheet.cell(row=header_row, column=column, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor=GenerateOutput.NAVY)
            cell.alignment = Alignment(
                horizontal="left" if column == start_column else "center",
                vertical="center",
                wrap_text=True,
            )
            cell.border = Border(
                bottom=Side(style="thin", color=GenerateOutput.GRID)
            )
        sheet.row_dimensions[header_row].height = 30

        for row_offset, values in enumerate(rows, start=1):
            row_number = header_row + row_offset
            for offset, value in enumerate(values):
                column = start_column + offset
                cell = sheet.cell(row=row_number, column=column, value=value)
                cell.font = Font(name="Aptos", size=10, color=GenerateOutput.TEXT)
                cell.alignment = Alignment(
                    horizontal="left" if offset == 0 else "right",
                    vertical="center",
                )
                cell.border = Border(
                    bottom=Side(style="hair", color=GenerateOutput.GRID)
                )
                if row_offset % 2 == 0:
                    cell.fill = PatternFill(
                        fill_type="solid", fgColor="F5F8FB"
                    )
                header = headers[offset].casefold()
                if "share" in header or "used" in header or "rate" in header:
                    cell.number_format = GenerateOutput.PERCENT_FORMAT
                elif header == "month" and isinstance(value, date):
                    cell.number_format = "mmm yyyy"
                elif "transaction" in header and isinstance(value, (int, float)):
                    cell.number_format = "#,##0"
                elif any(
                    part in header
                    for part in (
                        "spend",
                        "budget",
                        "income",
                        "remaining",
                        "cash flow",
                        "variance",
                    )
                ):
                    cell.number_format = GenerateOutput.CURRENCY_FORMAT

        last_data_row = header_row + max(1, len(rows))
        table_ref = (
            f"{get_column_letter(start_column)}{header_row}:"
            f"{get_column_letter(end_column)}{last_data_row}"
        )
        table = Table(displayName=table_name, ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)
        return last_data_row

    @staticmethod
    def _write_kpi(
        sheet, start_column: int, end_column: int, label: str, value, number_format: str
    ) -> None:
        card_fill = PatternFill(fill_type="solid", fgColor=GenerateOutput.LIGHT_BLUE)
        for row in range(5, 8):
            for column in range(start_column, end_column + 1):
                sheet.cell(row=row, column=column).fill = card_fill
        label_range = (
            f"{get_column_letter(start_column)}5:"
            f"{get_column_letter(end_column)}5"
        )
        value_range = (
            f"{get_column_letter(start_column)}6:"
            f"{get_column_letter(end_column)}7"
        )
        sheet.merge_cells(label_range)
        sheet.merge_cells(value_range)
        label_cell = sheet.cell(row=5, column=start_column, value=label.upper())
        label_cell.font = Font(name="Aptos", size=9, bold=True, color=GenerateOutput.MUTED)
        label_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell = sheet.cell(row=6, column=start_column, value=value)
        value_cell.font = Font(
            name="Aptos Display", size=19, bold=True, color=GenerateOutput.NAVY
        )
        value_cell.number_format = number_format
        value_cell.alignment = Alignment(horizontal="center", vertical="center")

    @staticmethod
    def _add_excel_table(sheet, name: str, reference: str) -> None:
        table = Table(displayName=name, ref=reference)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)

    @staticmethod
    def create_excel(
        statements: List[Statement],
        output_path: str = "",
        source_files: Optional[Sequence[tuple[Bank, str]]] = None,
    ) -> None:
        if source_files is not None and len(source_files) != len(statements):
            raise ValueError("Provide one source CSV file for each statement.")
        if not output_path:
            output_path = GenerateOutput.default_output_path(statements)
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        transactions = [
            transaction
            for statement in statements
            for transaction in statement.transactions
        ]
        targets = CategoryTargetHandler()
        monthly_targets = targets.targets
        category_totals = defaultdict(float)
        monthly_category_totals = defaultdict(lambda: defaultdict(float))
        daily_category_totals = defaultdict(lambda: defaultdict(float))
        monthly_income_totals = defaultdict(float)
        merchant_totals = defaultdict(float)
        merchant_counts = Counter()
        expenses = []
        income = []
        needs_review = []
        for transaction in transactions:
            category = category_name(transaction.category)
            month_key = (transaction.date.year, transaction.date.month)
            amount = float(transaction.amount)
            if category == Category.NONE.value:
                needs_review.append(transaction)
            elif category == Category.IGNORE.value:
                continue
            elif category == Category.INCOME.value:
                income.append(transaction)
                monthly_income_totals[month_key] += amount
            else:
                expenses.append(transaction)
                category_totals[category] += amount
                monthly_category_totals[month_key][category] += amount
                daily_category_totals[transaction.date][category] += amount
                merchant = (transaction.title or transaction.description or "").strip()
                merchant = merchant or "Unknown merchant"
                merchant_totals[(merchant, category)] += amount
                merchant_counts[(merchant, category)] += 1

        dated_transactions = [transaction.date for transaction in transactions]
        month_keys = sorted(
            {
                (transaction.date.year, transaction.date.month)
                for transaction in transactions
            }
        )
        if month_keys:
            first_year, first_month = month_keys[0]
            last_year, last_month = month_keys[-1]
            months_covered = (
                (last_year - first_year) * 12 + last_month - first_month + 1
            )
            report_months = [
                (
                    first_year + (first_month - 1 + offset) // 12,
                    (first_month - 1 + offset) % 12 + 1,
                )
                for offset in range(months_covered)
            ]
            if (first_year, first_month) == (last_year, last_month):
                period_label = date(first_year, first_month, 1).strftime("%B %Y")
            else:
                period_label = (
                    f"{date(first_year, first_month, 1):%b %Y} - "
                    f"{date(last_year, last_month, 1):%b %Y} "
                    f"({months_covered} months)"
                )
        else:
            months_covered = 0
            report_months = []
            period_label = "No transaction dates"

        category_names = [
            category_name(category)
            for category in monthly_targets
            if category_name(category)
            not in (Category.INCOME.value, Category.IGNORE.value, Category.NONE.value)
        ]
        for category in sorted(category_totals):
            if category not in category_names:
                category_names.append(category)
        category_spending = {
            category: sum(
                abs(monthly_category_totals.get(month, {}).get(category, 0.0))
                for month in report_months
            )
            for category in category_names
        }
        category_names.sort(
            key=lambda category: (-category_spending[category], category.casefold())
        )
        total_spending = sum(category_spending.values())
        total_income = sum(
            abs(monthly_income_totals.get(month, 0.0)) for month in report_months
        )
        net_cash_flow = total_income - total_spending
        savings_rate = net_cash_flow / total_income if total_income > 0 else None
        total_period_budget = 0.0
        category_rows = []
        for category in category_names:
            monthly_target = float(monthly_targets.get(category, 0.0) or 0.0)
            period_budget = monthly_target * months_covered
            spent = category_spending[category]
            remaining = period_budget - spent if monthly_target > 0 else None
            budget_used = spent / period_budget if period_budget > 0 else None
            share = spent / total_spending if total_spending > 0 else 0.0
            category_rows.append(
                (
                    category,
                    monthly_target,
                    period_budget,
                    spent,
                    remaining,
                    budget_used,
                    share,
                )
            )
            total_period_budget += period_budget

        merchant_rows = []
        for (merchant, category), amount in sorted(
            merchant_totals.items(),
            key=lambda item: abs(item[1]),
            reverse=True,
        )[:10]:
            spent = abs(amount)
            merchant_rows.append(
                (
                    merchant,
                    category,
                    spent,
                    merchant_counts[(merchant, category)],
                    spent / total_spending if total_spending > 0 else 0.0,
                )
            )

        monthly_rows = []
        for year, month in report_months:
            key = (year, month)
            month_income = abs(monthly_income_totals.get(key, 0.0))
            month_spending = sum(
                abs(amount)
                for amount in monthly_category_totals.get(key, {}).values()
            )
            monthly_rows.append(
                (
                    date(year, month, 1),
                    month_income,
                    month_spending,
                    month_income - month_spending,
                )
            )
        daily_rows = []
        if dated_transactions:
            first_day = min(dated_transactions)
            last_day = max(dated_transactions)
            day_count = (last_day - first_day).days + 1
            crosses_year = first_day.year != last_day.year
            for offset in range(day_count):
                day = first_day + timedelta(days=offset)
                day_spending = sum(
                    abs(amount)
                    for amount in daily_category_totals.get(day, {}).values()
                )
                daily_rows.append((day, day_spending))
        else:
            crosses_year = False

        workbook = openpyxl.Workbook()
        dashboard = workbook.active
        dashboard.title = "Dashboard"
        summary = workbook.create_sheet("Summary")
        transactions_sheet = workbook.create_sheet("Expenses & Income")
        plotting_data = workbook.create_sheet("Plotting Data")
        workbook.properties.title = "Statement Parser Spending Dashboard"
        workbook.properties.subject = "Personal spending and monthly budget report"
        workbook.properties.creator = "Statement Parser"
        for sheet in (dashboard, summary, transactions_sheet, plotting_data):
            GenerateOutput._style_sheet(sheet)

        # Executive dashboard
        dashboard.column_dimensions["A"].width = 3
        for column in range(2, 14):
            dashboard.column_dimensions[get_column_letter(column)].width = 18
        GenerateOutput._write_title(
            dashboard, "B1:M1", "SPENDING AND CASH FLOW DASHBOARD"
        )
        dashboard.merge_cells("B2:M2")
        subtitle = dashboard["B2"]
        subtitle.value = (
            f"Reporting period: {period_label} | "
            "Spending excludes ignored and uncategorized transactions."
        )
        subtitle.font = Font(name="Aptos", size=11, color=GenerateOutput.MUTED)
        subtitle.alignment = Alignment(vertical="center")
        dashboard.row_dimensions[2].height = 23
        dashboard.merge_cells("B3:M3")
        review_amount = sum(abs(float(transaction.amount)) for transaction in needs_review)
        review_text = (
            f"Review needed: {len(needs_review)} uncategorized transaction(s), "
            f"{review_amount:,.2f} in activity."
            if needs_review
            else "Categorization status: no uncategorized transactions."
        )
        review_cell = dashboard["B3"]
        review_cell.value = review_text
        review_cell.fill = PatternFill(
            fill_type="solid",
            fgColor=GenerateOutput.LIGHT_AMBER if needs_review else GenerateOutput.LIGHT_TEAL,
        )
        review_cell.font = Font(name="Aptos", size=10, color=GenerateOutput.TEXT)
        review_cell.alignment = Alignment(vertical="center", indent=1)
        dashboard.row_dimensions[3].height = 24
        GenerateOutput._write_kpi(
            dashboard, 2, 4, "Total Income", total_income, GenerateOutput.CURRENCY_FORMAT
        )
        GenerateOutput._write_kpi(
            dashboard, 5, 7, "Total Spending", total_spending, GenerateOutput.CURRENCY_FORMAT
        )
        GenerateOutput._write_kpi(
            dashboard, 8, 10, "Net Cash Flow", net_cash_flow, GenerateOutput.CURRENCY_FORMAT
        )
        GenerateOutput._write_kpi(
            dashboard, 11, 13, "Savings Rate", savings_rate, GenerateOutput.PERCENT_FORMAT
        )
        dashboard.row_dimensions[5].height = 21
        dashboard.row_dimensions[6].height = 25
        dashboard.row_dimensions[7].height = 22

        dashboard_category_rows = [
            (
                category,
                spent,
                budget,
                remaining,
                budget_used,
                share,
            )
            for (
                category,
                _monthly_target,
                budget,
                spent,
                remaining,
                budget_used,
                share,
            ) in category_rows
        ]
        category_table_end = GenerateOutput._write_section_table(
            dashboard,
            "Spending and Budget by Category",
            (
                "Category",
                "Actual Spend",
                "Period Budget",
                "Remaining",
                "Budget Used",
                "Share",
            ),
            dashboard_category_rows,
            10,
            2,
            "DashboardCategoryTable",
        )
        for row_number, row in enumerate(category_rows, start=12):
            monthly_target, period_budget, spent = row[1], row[2], row[3]
            status_fill = (
                GenerateOutput.LIGHT_AMBER
                if monthly_target <= 0
                else (
                    GenerateOutput.LIGHT_RED
                    if spent > period_budget
                    else GenerateOutput.LIGHT_TEAL
                )
            )
            for column in (3, 5, 6):
                dashboard.cell(row=row_number, column=column).fill = PatternFill(
                    fill_type="solid", fgColor=status_fill
                )
        merchant_table_end = GenerateOutput._write_section_table(
            dashboard,
            "Top Merchants",
            ("Merchant", "Category", "Net Spend", "Transactions", "Share"),
            merchant_rows,
            10,
            8,
            "DashboardMerchantTable",
        )
        trend_table_start = max(category_table_end, merchant_table_end) + 30
        trend_table_end = GenerateOutput._write_section_table(
            dashboard,
            "Monthly Cash Flow",
            ("Month", "Income", "Spending", "Net Cash Flow"),
            monthly_rows,
            trend_table_start,
            8,
            "DashboardCashFlowTable",
        )

        # Chart data is kept on a hidden sheet so the dashboard stays uncluttered.
        plotting_data.append(("Category", "Actual Spend", "Period Budget"))
        for category, _monthly_target, budget, spent, *_ in category_rows:
            plotting_data.append((category, spent, budget))
        plotting_data.cell(row=1, column=5, value="Month")
        plotting_data.cell(row=1, column=6, value="Income")
        plotting_data.cell(row=1, column=7, value="Spending")
        plotting_data.cell(row=1, column=8, value="Net Cash Flow")
        for row_number, values in enumerate(monthly_rows, start=2):
            for column, value in enumerate(values, start=5):
                plotting_data.cell(row=row_number, column=column, value=value)
        plotting_data.cell(row=1, column=10, value="Day")
        plotting_data.cell(row=1, column=11, value="Spending")
        for row_number, (day, spending) in enumerate(daily_rows, start=2):
            date_label = day.strftime("%b %d, %Y" if crosses_year else "%b %d")
            plotting_data.cell(row=row_number, column=10, value=date_label)
            plotting_data.cell(row=row_number, column=11, value=spending)
        plotting_data.sheet_state = "hidden"

        chart_top = max(category_table_end, merchant_table_end) + 3
        if category_rows:
            budget_chart = BarChart()
            budget_chart.type = "bar"
            budget_chart.style = 10
            budget_chart.title = "Actual Spend vs. Period Budget"
            budget_chart.height = 12
            budget_chart.width = 18
            budget_chart.grouping = "clustered"
            budget_chart.add_data(
                Reference(
                    plotting_data,
                    min_col=2,
                    max_col=3,
                    min_row=1,
                    max_row=len(category_rows) + 1,
                ),
                titles_from_data=True,
            )
            budget_chart.set_categories(
                Reference(
                    plotting_data,
                    min_col=1,
                    min_row=2,
                    max_row=len(category_rows) + 1,
                )
            )
            budget_chart.legend.position = "r"
            budget_chart.visible_cells_only = False
            if len(budget_chart.series) >= 2:
                budget_chart.series[0].graphicalProperties.solidFill = GenerateOutput.TEAL
                budget_chart.series[1].graphicalProperties.solidFill = "AAB8C6"
            dashboard.add_chart(budget_chart, f"H{chart_top}")

        nonzero_categories = [
            row for row in category_rows if row[3] > 0
        ]
        if nonzero_categories:
            pie_chart = DoughnutChart()
            pie_chart.style = 10
            pie_chart.title = "Spending Mix"
            pie_chart.holeSize = 62
            pie_chart.height = 12
            pie_chart.width = 18
            pie_chart.add_data(
                Reference(
                    plotting_data,
                    min_col=2,
                    min_row=1,
                    max_row=len(nonzero_categories) + 1,
                ),
                titles_from_data=True,
            )
            pie_chart.set_categories(
                Reference(
                    plotting_data,
                    min_col=1,
                    min_row=2,
                    max_row=len(nonzero_categories) + 1,
                )
            )
            pie_chart.dataLabels = DataLabelList()
            pie_chart.dataLabels.showPercent = True
            pie_chart.dataLabels.showCatName = False
            pie_chart.dataLabels.showSerName = False
            pie_chart.dataLabels.showVal = False
            pie_chart.dataLabels.showLegendKey = False
            pie_chart.dataLabels.showLeaderLines = True
            pie_chart.legend.position = "r"
            pie_chart.visible_cells_only = False
            dashboard.add_chart(pie_chart, f"B{chart_top}")

        if monthly_rows:
            cash_flow_chart = BarChart()
            cash_flow_chart.type = "col"
            cash_flow_chart.style = 10
            cash_flow_chart.title = "Income, Spending and Net Cash Flow"
            cash_flow_chart.y_axis.numFmt = GenerateOutput.CURRENCY_FORMAT
            cash_flow_chart.height = 10
            cash_flow_chart.width = 18
            cash_flow_chart.grouping = "clustered"
            cash_flow_chart.add_data(
                Reference(
                    plotting_data,
                    min_col=6,
                    max_col=8,
                    min_row=1,
                    max_row=len(monthly_rows) + 1,
                ),
                titles_from_data=True,
            )
            cash_flow_chart.set_categories(
                Reference(
                    plotting_data,
                    min_col=5,
                    min_row=2,
                    max_row=len(monthly_rows) + 1,
                )
            )
            cash_flow_chart.legend.position = "r"
            cash_flow_chart.visible_cells_only = False
            for series, color in zip(
                cash_flow_chart.series,
                (GenerateOutput.TEAL, GenerateOutput.BLUE, "E5A94D"),
            ):
                series.graphicalProperties.solidFill = color
            dashboard.add_chart(cash_flow_chart, f"B{trend_table_start}")
        daily_chart_row = max(trend_table_start + 22, trend_table_end + 3)
        if daily_rows:
            daily_chart = LineChart()
            daily_chart.style = 13
            daily_chart.title = "Daily Spending"
            daily_chart.y_axis.numFmt = GenerateOutput.CURRENCY_FORMAT
            daily_chart.height = 10
            daily_chart.width = 24
            daily_chart.legend = None
            daily_chart.add_data(
                Reference(
                    plotting_data,
                    min_col=11,
                    min_row=1,
                    max_row=len(daily_rows) + 1,
                ),
                titles_from_data=True,
            )
            daily_chart.set_categories(
                Reference(
                    plotting_data,
                    min_col=10,
                    min_row=2,
                    max_row=len(daily_rows) + 1,
                )
            )
            label_interval = max(1, (len(daily_rows) + 9) // 10)
            daily_chart.x_axis.tickLblSkip = label_interval
            daily_chart.x_axis.tickMarkSkip = label_interval
            series = daily_chart.series[0]
            series.graphicalProperties.line.solidFill = GenerateOutput.TEAL
            series.graphicalProperties.line.width = 24000
            series.marker.symbol = "circle"
            series.marker.size = 4
            daily_chart.visible_cells_only = False
            dashboard.add_chart(daily_chart, f"B{daily_chart_row}")
        dashboard.freeze_panes = None
        dashboard.sheet_view.topLeftCell = "A1"
        if dashboard.sheet_view.selection:
            dashboard.sheet_view.selection[0].activeCell = "A1"
            dashboard.sheet_view.selection[0].sqref = "A1"
        dashboard.print_area = (
            f"B1:M{max(trend_table_end, daily_chart_row + 20)}"
        )

        # Category detail and income plan
        GenerateOutput._write_title(summary, "B1:M1", "CATEGORY AND INCOME SUMMARY")
        summary.merge_cells("B2:M2")
        summary["B2"] = (
            f"Period: {period_label} | Monthly targets are scaled across "
            f"{months_covered} month(s)."
        )
        summary["B2"].font = Font(name="Aptos", size=10, color=GenerateOutput.MUTED)
        summary_rows = [
            (
                category,
                monthly_target,
                period_budget,
                spent,
                remaining,
                budget_used,
                share,
            )
            for (
                category,
                monthly_target,
                period_budget,
                spent,
                remaining,
                budget_used,
                share,
            ) in category_rows
        ]
        summary_end = GenerateOutput._write_section_table(
            summary,
            "Category Performance",
            (
                "Category",
                "Monthly Target",
                "Period Budget",
                "Actual Spend",
                "Remaining",
                "Budget Used",
                "Share",
            ),
            summary_rows,
            3,
            2,
            "CategorySummary",
        )
        for row_number, row in enumerate(category_rows, start=5):
            monthly_target, period_budget, spent = row[1], row[2], row[3]
            status_fill = (
                GenerateOutput.LIGHT_AMBER
                if monthly_target <= 0
                else (
                    GenerateOutput.LIGHT_RED
                    if spent > period_budget
                    else GenerateOutput.LIGHT_TEAL
                )
            )
            for column in (5, 6, 7):
                summary.cell(row=row_number, column=column).fill = PatternFill(
                    fill_type="solid", fgColor=status_fill
                )
        total_row = summary_end + 1
        total_values = (
            "Total",
            sum(row[1] for row in category_rows),
            total_period_budget,
            total_spending,
            total_period_budget - total_spending,
            total_spending / total_period_budget if total_period_budget else None,
            1.0 if total_spending else 0.0,
        )
        for column, value in enumerate(total_values, start=2):
            cell = summary.cell(row=total_row, column=column, value=value)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor=GenerateOutput.NAVY)
            if column in (3, 4, 5, 6):
                cell.number_format = GenerateOutput.CURRENCY_FORMAT
            elif column in (7, 8):
                cell.number_format = GenerateOutput.PERCENT_FORMAT

        income_target = float(monthly_targets.get(Category.INCOME.value, 0.0) or 0.0)
        income_period_target = income_target * months_covered
        income_rows = [
            (
                income_target,
                income_period_target,
                total_income,
                total_income - income_period_target,
            )
        ]
        GenerateOutput._write_section_table(
            summary,
            "Income Plan",
            ("Monthly Target", "Period Target", "Actual Income", "Variance"),
            income_rows,
            3,
            10,
            "IncomePlan",
        )
        for column, width in {
            "A": 3,
            "B": 24,
            "C": 17,
            "D": 17,
            "E": 17,
            "F": 17,
            "G": 14,
            "H": 14,
            "I": 3,
            "J": 17,
            "K": 17,
            "L": 17,
            "M": 17,
        }.items():
            summary.column_dimensions[column].width = width
        summary.freeze_panes = "B5"
        summary.print_area = f"B1:M{max(total_row, 6)}"

        # Transaction detail remains available beside the dashboard.
        GenerateOutput._write_title(
            transactions_sheet, "B1:F1", "PURCHASES AND EXPENSES"
        )
        GenerateOutput._write_title(
            transactions_sheet, "H1:L1", "INCOME TRANSACTIONS"
        )
        expense_rows = []
        running_total = 0.0
        for transaction in sorted(expenses, key=lambda item: item.date):
            running_total += float(transaction.amount)
            expense_rows.append(
                (
                    transaction.title or transaction.description,
                    transaction.date,
                    category_name(transaction.category),
                    float(transaction.amount),
                    running_total,
                )
            )
        income_rows_detail = []
        running_total = 0.0
        for transaction in sorted(income, key=lambda item: item.date):
            running_total += float(transaction.amount)
            income_rows_detail.append(
                (
                    transaction.title or transaction.description,
                    transaction.date,
                    category_name(transaction.category),
                    float(transaction.amount),
                    running_total,
                )
            )
        transaction_headers = (
            "Title",
            "Date",
            "Category",
            "Amount",
            "Running Total",
        )
        for column_offset, header in enumerate(transaction_headers, start=2):
            cell = transactions_sheet.cell(row=2, column=column_offset, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor=GenerateOutput.NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for column_offset, header in enumerate(transaction_headers, start=8):
            cell = transactions_sheet.cell(row=2, column=column_offset, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor=GenerateOutput.NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row_number, values in enumerate(expense_rows, start=3):
            for column, value in enumerate(values, start=2):
                cell = transactions_sheet.cell(row=row_number, column=column, value=value)
                cell.alignment = Alignment(vertical="center")
                if column == 3:
                    cell.number_format = "mm/dd/yyyy"
                elif column in (5, 6):
                    cell.number_format = GenerateOutput.CURRENCY_FORMAT
        for row_number, values in enumerate(income_rows_detail, start=3):
            for column, value in enumerate(values, start=8):
                cell = transactions_sheet.cell(row=row_number, column=column, value=value)
                cell.alignment = Alignment(vertical="center")
                if column == 9:
                    cell.number_format = "mm/dd/yyyy"
                elif column in (11, 12):
                    cell.number_format = GenerateOutput.CURRENCY_FORMAT
        expense_table_last = max(2, len(expense_rows) + 2)
        income_table_last = max(2, len(income_rows_detail) + 2)
        GenerateOutput._add_excel_table(
            transactions_sheet, "ExpenseTransactions", f"B2:F{expense_table_last}"
        )
        GenerateOutput._add_excel_table(
            transactions_sheet, "IncomeTransactions", f"H2:L{income_table_last}"
        )
        for column, width in {
            "A": 3,
            "B": 32,
            "C": 13,
            "D": 22,
            "E": 16,
            "F": 18,
            "G": 3,
            "H": 32,
            "I": 13,
            "J": 22,
            "K": 16,
            "L": 18,
        }.items():
            transactions_sheet.column_dimensions[column].width = width
        transactions_sheet.freeze_panes = "B3"
        transactions_sheet.print_area = (
            f"B1:L{max(expense_table_last, income_table_last)}"
        )

        # Retain workbook chart values in a hidden source sheet.
        plotting_data.sheet_state = "hidden"
        if source_files:
            GenerateOutput._add_source_csv_sheets(workbook, source_files)
        workbook.active = 0
        workbook.save(output_path)

    @staticmethod
    def default_history_output_path(months: Sequence[date]) -> str:
        if months:
            period = f"{months[0]:%Y-%m}_to_{months[-1]:%Y-%m}"
        else:
            period = "no-dates"
        return os.path.join(
            GenerateOutput.default_output_directory(),
            f"HistoricalAnalysis_{period}.xlsx",
        )

    @staticmethod
    def create_history_excel(
        statements: Sequence[Statement],
        source_files: Sequence[str],
        analysis,
        output_path: str,
        duplicate_groups=(),
        excluded_occurrences: Optional[set[tuple[int, int]]] = None,
    ) -> None:
        if len(statements) != len(source_files):
            raise ValueError("Provide one source workbook for each statement.")
        if not analysis.months:
            raise ValueError("No transaction dates were found for historical analysis.")
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        workbook = openpyxl.Workbook()
        overview = workbook.active
        overview.title = "Overview"
        trends = workbook.create_sheet("Category Trends")
        subscriptions = workbook.create_sheet("Subscriptions")
        transactions_sheet = workbook.create_sheet("Transactions")
        plotting_data = workbook.create_sheet("Analysis Data")
        for sheet in (
            overview,
            trends,
            subscriptions,
            transactions_sheet,
            plotting_data,
        ):
            GenerateOutput._style_sheet(sheet)
        workbook.properties.title = "Statement Parser Historical Analysis"
        workbook.properties.subject = "Category trends and recurring transactions"
        workbook.properties.creator = "Statement Parser"

        month_label = (
            f"{analysis.months[0]:%b %Y}"
            if len(analysis.months) == 1
            else f"{analysis.months[0]:%b %Y} - {analysis.months[-1]:%b %Y}"
        )
        overview.column_dimensions["A"].width = 3
        for column in range(2, 17):
            overview.column_dimensions[get_column_letter(column)].width = 16
        GenerateOutput._write_title(
            overview, "B1:P1", "HISTORY AND TRENDS OVERVIEW"
        )
        overview.merge_cells("B2:P2")
        overview["B2"] = (
            f"Reporting period: {month_label} | "
            f"{len(analysis.months)} month(s) analyzed."
        )
        overview["B2"].font = Font(name="Aptos", size=11, color=GenerateOutput.MUTED)
        overview.merge_cells("B3:P3")
        overview["B3"] = (
            "Category trends use the monthly average and estimated change per month. "
            "Subscription candidates require the same merchant and amount in 3+ months."
        )
        overview["B3"].font = Font(name="Aptos", size=10, color=GenerateOutput.MUTED)
        average_spending = analysis.total_spending / len(analysis.months)
        monthly_subscriptions = sum(
            candidate.amount for candidate in analysis.subscriptions
        )
        for start_column, end_column, label, value, number_format in (
            (2, 4, "Total Income", analysis.total_income, GenerateOutput.CURRENCY_FORMAT),
            (5, 7, "Total Spending", analysis.total_spending, GenerateOutput.CURRENCY_FORMAT),
            (8, 10, "Average Monthly Spending", average_spending, GenerateOutput.CURRENCY_FORMAT),
            (11, 13, "Net Cash Flow", analysis.total_income - analysis.total_spending, GenerateOutput.CURRENCY_FORMAT),
            (14, 16, "Potential Monthly Subscriptions", monthly_subscriptions, GenerateOutput.CURRENCY_FORMAT),
        ):
            GenerateOutput._write_kpi(
                overview, start_column, end_column, label, value, number_format
            )
        overview.freeze_panes = None
        overview.sheet_view.topLeftCell = "A1"
        if overview.sheet_view.selection:
            overview.sheet_view.selection[0].activeCell = "A1"
            overview.sheet_view.selection[0].sqref = "A1"

        monthly_rows = [
            (
                month,
                analysis.monthly_income[index],
                analysis.monthly_spending[index],
                analysis.monthly_income[index] - analysis.monthly_spending[index],
            )
            for index, month in enumerate(analysis.months)
        ]
        monthly_end = GenerateOutput._write_section_table(
            overview,
            "Monthly Cash Flow",
            ("Month", "Income", "Spending", "Net Cash Flow"),
            monthly_rows,
            10,
            2,
            "HistoricalMonthlyCashFlow",
        )
        overview.column_dimensions["B"].width = 19
        overview.column_dimensions["C"].width = 18
        overview.column_dimensions["D"].width = 18
        overview.column_dimensions["E"].width = 18
        plotting_data.append(("Month", "Income", "Spending", "Net Cash Flow"))
        for row_number, values in enumerate(monthly_rows, start=2):
            for column, value in enumerate(values, start=1):
                plotting_data.cell(row=row_number, column=column, value=value)

        monthly_chart = BarChart()
        monthly_chart.type = "col"
        monthly_chart.style = 10
        monthly_chart.title = "Income, Spending and Net Cash Flow"
        monthly_chart.height = 10
        monthly_chart.width = 20
        monthly_chart.grouping = "clustered"
        monthly_chart.add_data(
            Reference(
                plotting_data,
                min_col=2,
                max_col=4,
                min_row=1,
                max_row=len(monthly_rows) + 1,
            ),
            titles_from_data=True,
        )
        monthly_chart.set_categories(
            Reference(
                plotting_data,
                min_col=1,
                min_row=2,
                max_row=len(monthly_rows) + 1,
            )
        )
        monthly_chart.legend.position = "r"
        monthly_chart.visible_cells_only = False
        for series, color in zip(
            monthly_chart.series,
            (GenerateOutput.TEAL, GenerateOutput.BLUE, "E5A94D"),
        ):
            series.graphicalProperties.solidFill = color
        overview.add_chart(monthly_chart, "H10")
        overview.print_area = f"B1:P{max(monthly_end, 31)}"

        # Category Trends sheet pairs each category's average and slope with its monthly history.
        GenerateOutput._write_title(trends, "B1:M1", "CATEGORY TRENDS")
        trends.merge_cells("B2:M2")
        trends["B2"] = (
            "Average and trend start at each category's first spending month; later "
            "zero-spend months remain included. Positive changes indicate increases; "
            "negative changes indicate decreases."
        )
        trends["B2"].font = Font(name="Aptos", size=10, color=GenerateOutput.MUTED)
        category_rows = []
        for item in analysis.category_history:
            if item.monthly_change is None:
                direction = "Need 2+ months"
            elif item.monthly_change > 0.01:
                direction = "Increasing"
            elif item.monthly_change < -0.01:
                direction = "Decreasing"
            else:
                direction = "Stable"
            category_rows.append(
                (
                    item.category,
                    sum(item.monthly_spending),
                    item.average_monthly_spend,
                    item.monthly_change,
                    direction,
                )
            )
        category_end = GenerateOutput._write_section_table(
            trends,
            "Total, Average and Monthly Change",
            (
                "Category",
                "Total Spent",
                "Average / Month Since First Spend",
                "Change / Month",
                "Direction",
            ),
            category_rows,
            3,
            2,
            "HistoricalCategoryTrends",
        )
        has_trend_chart = any(
            item.monthly_change is not None for item in analysis.category_history
        )
        trend_chart_height = max(8, min(18, len(category_rows) * 0.55))
        trend_chart_end_row = (
            3 + round(trend_chart_height * 1.9) if has_trend_chart else category_end
        )
        matrix_start_row = max(category_end + 3, trend_chart_end_row + 3)
        month_headers = tuple(month.strftime("%b %Y") for month in analysis.months)
        matrix_rows = [
            (item.category, *item.monthly_spending)
            for item in analysis.category_history
        ]
        matrix_end = GenerateOutput._write_section_table(
            trends,
            "Monthly Spending by Category",
            ("Category", *month_headers),
            matrix_rows,
            matrix_start_row,
            2,
            "HistoricalCategoryMonthlySpending",
        )
        for column in range(2, 3 + len(analysis.months)):
            trends.column_dimensions[get_column_letter(column)].width = 16
        trends.column_dimensions["B"].width = 24
        trends.column_dimensions["C"].width = 16
        trends.column_dimensions["D"].width = 30
        trends.column_dimensions["E"].width = 19
        trends.column_dimensions["F"].width = 18
        trends.freeze_panes = "C5"

        trend_data_start = 2
        plotting_data.cell(row=1, column=6, value="Category")
        plotting_data.cell(row=1, column=7, value="Increasing")
        plotting_data.cell(row=1, column=8, value="Decreasing")
        for row_number, item in enumerate(analysis.category_history, start=trend_data_start):
            plotting_data.cell(row=row_number, column=6, value=item.category)
            plotting_data.cell(
                row=row_number,
                column=7,
                value=max(item.monthly_change, 0)
                if item.monthly_change is not None
                else None,
            )
            plotting_data.cell(
                row=row_number,
                column=8,
                value=min(item.monthly_change, 0)
                if item.monthly_change is not None
                else None,
            )
        if has_trend_chart:
            trend_chart = BarChart()
            trend_chart.type = "bar"
            trend_chart.style = 10
            trend_chart.title = "Estimated Monthly Change by Category"
            trend_chart.height = trend_chart_height
            trend_chart.width = 18
            trend_chart.grouping = "stacked"
            trend_chart.overlap = 100
            trend_chart.add_data(
                Reference(
                    plotting_data,
                    min_col=7,
                    max_col=8,
                    min_row=1,
                    max_row=len(category_rows) + 1,
                ),
                titles_from_data=True,
            )
            trend_chart.set_categories(
                Reference(
                    plotting_data,
                    min_col=6,
                    min_row=2,
                    max_row=len(category_rows) + 1,
                )
            )
            trend_chart.legend.position = "r"
            trend_chart.visible_cells_only = False
            if len(trend_chart.series) >= 2:
                trend_chart.series[0].graphicalProperties.solidFill = GenerateOutput.TEAL
                trend_chart.series[1].graphicalProperties.solidFill = "D66B6B"
            trends.add_chart(trend_chart, "G3")
        trend_print_end_column = max(15 if has_trend_chart else 5, 2 + len(analysis.months))
        trends.print_area = (
            f"B1:{get_column_letter(trend_print_end_column)}{matrix_end}"
        )

        subscription_rows = [
            (
                item.merchant,
                item.amount,
                len(item.months),
                item.transaction_count,
                ", ".join(month.strftime("%b %Y") for month in item.months),
                item.first_date,
                item.last_date,
            )
            for item in analysis.subscriptions
        ]
        GenerateOutput._write_title(
            subscriptions, "B1:H1", "LIKELY SUBSCRIPTIONS"
        )
        subscriptions.merge_cells("B2:H2")
        subscriptions["B2"] = (
            "Candidates appear in at least 3 distinct months with the same merchant "
            "and amount. Review these suggestions before changing your budget."
        )
        subscriptions["B2"].font = Font(
            name="Aptos", size=10, color=GenerateOutput.MUTED
        )
        subscription_end = GenerateOutput._write_section_table(
            subscriptions,
            "Recurring Charges to Review",
            (
                "Merchant",
                "Amount",
                "Months",
                "Transactions",
                "Months Observed",
                "First Seen",
                "Last Seen",
            ),
            subscription_rows,
            4,
            2,
            "HistoricalSubscriptions",
        )
        for column, width in {
            "B": 32,
            "C": 16,
            "D": 12,
            "E": 15,
            "F": 40,
            "G": 16,
            "H": 16,
        }.items():
            subscriptions.column_dimensions[column].width = width
        for row_number in range(6, subscription_end + 1):
            subscriptions.cell(row=row_number, column=7).number_format = "mmm d, yyyy"
            subscriptions.cell(row=row_number, column=8).number_format = "mmm d, yyyy"
        subscriptions.freeze_panes = "B6"
        subscriptions.print_area = f"B1:H{max(subscription_end, 7)}"

        transaction_rows = []
        for statement_index, statement in enumerate(statements):
            source_path = source_files[statement_index]
            for transaction in statement.transactions:
                transaction_rows.append(
                    (
                        transaction.date,
                        date(transaction.date.year, transaction.date.month, 1),
                        transaction.title or transaction.description,
                        float(transaction.amount),
                        category_name(transaction.category),
                        os.path.basename(source_path),
                    )
                )
        GenerateOutput._write_title(
            transactions_sheet, "B1:G1", "CONSOLIDATED TRANSACTIONS"
        )
        transaction_end = GenerateOutput._write_section_table(
            transactions_sheet,
            "Transactions from Selected Reports",
            ("Date", "Month", "Merchant / Description", "Amount", "Category", "Source"),
            sorted(transaction_rows, key=lambda row: row[0]),
            3,
            2,
            "HistoricalTransactions",
        )
        for column, width in {
            "B": 14,
            "C": 14,
            "D": 34,
            "E": 16,
            "F": 22,
            "G": 32,
        }.items():
            transactions_sheet.column_dimensions[column].width = width
        for row_number in range(5, transaction_end + 1):
            transactions_sheet.cell(row=row_number, column=2).number_format = "mmm d, yyyy"
            transactions_sheet.cell(row=row_number, column=3).number_format = "mmm yyyy"
        transactions_sheet.freeze_panes = "D5"
        transactions_sheet.print_area = f"B1:G{max(transaction_end, 5)}"

        if duplicate_groups:
            duplicate_sheet = workbook.create_sheet("Duplicates")
            GenerateOutput._style_sheet(duplicate_sheet)
            GenerateOutput._write_title(
                duplicate_sheet, "B1:F1", "POTENTIAL DUPLICATES"
            )
            duplicate_rows = []
            automatically_excluded = set()
            if excluded_occurrences is None:
                for group in duplicate_groups:
                    first_source_key = group.occurrences[0].source_key
                    automatically_excluded.update(
                        (
                            occurrence.statement_index,
                            occurrence.transaction_index,
                        )
                        for occurrence in group.occurrences
                        if occurrence.source_key != first_source_key
                    )
            for group in duplicate_groups:
                for occurrence in group.occurrences:
                    transaction = occurrence.transaction
                    key = (
                        occurrence.statement_index,
                        occurrence.transaction_index,
                    )
                    is_excluded = (
                        key in excluded_occurrences
                        if excluded_occurrences is not None
                        else key in automatically_excluded
                    )
                    duplicate_rows.append(
                        (
                            transaction.date,
                            transaction.amount,
                            transaction.title or transaction.description,
                            os.path.basename(occurrence.source_path),
                            "Excluded from analysis"
                            if is_excluded
                            else "Included in analysis",
                        )
                    )
            duplicate_end = GenerateOutput._write_section_table(
                duplicate_sheet,
                "Matched Transactions and Analysis Decision",
                ("Date", "Amount", "Merchant / Description", "Source", "Action"),
                duplicate_rows,
                3,
                2,
                "HistoricalDuplicateRows",
            )
            duplicate_sheet.freeze_panes = "B5"
            duplicate_sheet.print_area = f"B1:F{max(duplicate_end, 5)}"

        plotting_data.sheet_state = "hidden"
        workbook.active = 0
        workbook.save(output_path)
