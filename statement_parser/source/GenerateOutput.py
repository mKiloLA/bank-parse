import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
from datetime import datetime
from typing import List
from openpyxl.chart import PieChart, Reference, LineChart
from openpyxl.styles import Font, PatternFill

from .data.Statement import Statement
from .data.DataTypes import Category
from .CategoryTargetHandler import CategoryTargetHandler

class GenerateOutput:
    @staticmethod
    def create_excel(statements: List[Statement], output_path: str = "output/expenses.xlsx") -> None:
        # Ensure output directory exists
        output_dir = os.path.dirname(output_path)
        if output_dir:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
        def _gradient_color(ratio: float, color_type: str) -> str:
            # color_type: 'green' or 'red'
            ratio = max(min(ratio, 1.0), -1.0)
            if color_type == 'red':
                r = 255
                g = int(254 - (254 - 0) * abs(ratio))
                b = 0
                return f'{r:02X}{g:02X}{b:02X}'
            else:
                r = int(255 - (255 - 0) * abs(ratio))
                g = 255
                b = int(224 - (224 - 0) * abs(ratio))
                return f'{r:02X}{g:02X}{b:02X}'

        def get_relative_color(spent: float, target: float, is_income: bool = False) -> str:
            # If target is zero, use yellow
            if target == 0:
                return 'FFFFE0'

            spent = abs(spent)

            ratio = (spent - target) / target
            # For income, reverse: below target is red, above is green
            if is_income:
                if spent < target:
                    return _gradient_color(ratio, 'red')
                else:
                    return _gradient_color(ratio, 'green')
            else:
                if spent <= target:
                    return _gradient_color(ratio, 'green')
                else:
                    return _gradient_color(ratio, 'red')

        def get_gradient_color(val: float) -> str:
            # Negative: red gradient, Positive: green gradient, Zero: yellow
            max_abs = 1000.0
            if val == 0:
                return 'FFFFE0'
            ratio = min(abs(val), max_abs) / max_abs
            if val < 0:
                return _gradient_color(ratio, 'red')
            else:
                return _gradient_color(ratio, 'green')

        wb = openpyxl.Workbook()
        if wb.active:
            ws_summary = wb.active
            ws_summary.title = "Summary"
        else:
            ws_summary = wb.create_sheet("Summary", 0)
        ws_expenses = wb.create_sheet("Expenses & Income", 1)
        ws_plotting_data = wb.create_sheet("Plotting Data", 2)

        # Expenses Table
        expenses_header = ["Title", "Date", "Category", "Amount", "Running Total"]
        expenses_data = [expenses_header]
        expenses = []
        for statement in statements:
            for t in statement.transactions:
                if t.is_purchase():
                    expenses.append(t)
        expenses.sort(key=lambda x: x.date)
        running_total = 0.0
        running_total_abs_values = []
        for t in expenses:
            running_total += t.amount
            expenses_data.append([
                t.title,
                t.date.strftime("%m/%d/%Y"),
                t.category.value,
                t.amount,
                running_total
            ])
            running_total_abs_values.append(abs(running_total))
        # Income Table
        income_header = ["Title", "Date", "Category", "Amount", "Running Total"]
        income_data = [income_header]
        income = []
        for statement in statements:
            for t in statement.transactions:
                if t.is_income():
                    income.append(t)
        income.sort(key=lambda x: x.date)
        running_total = 0.0
        for t in income:
            running_total += t.amount
            income_data.append([
                t.title,
                t.date.strftime("%m/%d/%Y"),
                t.category.value,
                t.amount,
                running_total
            ])
        # Write expenses data starting at B2
        for row_idx, row_data in enumerate(expenses_data, start=2):
            for col_idx, value in enumerate(row_data, start=2):
                cell = ws_expenses.cell(row=row_idx, column=col_idx, value=value)
                # Format Amount and Running Total columns as currency
                if row_idx > 2:
                    # Amount column (col_idx == 5)
                    if col_idx == 5:
                        cell.number_format = '$#,##0.00'
                        # Gradient color by value
                        try:
                            val = float(value)
                            fill_color = get_gradient_color(val)
                            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                        except:
                            pass
                    # Running Total column (col_idx == 6)
                    if col_idx == 6:
                        cell.number_format = '$#,##0.00'
        # Write income data starting one column apart from expenses
        income_start_col = 2 + len(expenses_header) + 1
        for row_idx, row_data in enumerate(income_data, start=2):
            for col_idx, value in enumerate(row_data, start=income_start_col):
                cell = ws_expenses.cell(row=row_idx, column=col_idx, value=value)
                # Format Amount and Running Total columns as currency
                if row_idx > 2:
                    # Amount column (col_idx == income_start_col + 3)
                    if col_idx == income_start_col + 3:
                        cell.number_format = '$#,##0.00'
                        # Gradient color by value
                        try:
                            val = float(value)
                            fill_color = get_gradient_color(val)
                            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                        except:
                            pass
                    # Running Total column (col_idx == income_start_col + 4)
                    if col_idx == income_start_col + 4:
                        cell.number_format = '$#,##0.00'
        # Auto-size columns for main sheet, skip column A
        for col in range(2, ws_expenses.max_column + 1):
            max_length = 0
            col_letter = get_column_letter(col)
            for cell in ws_expenses[col_letter]:
                try:
                    cell_length = len(str(cell.value)) if cell.value is not None else 0
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            ws_expenses.column_dimensions[col_letter].width = max_length + 2
        from openpyxl.styles import Alignment, Border, Side
        header_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        # Style header row for expenses (row 2)
        for col_idx in range(2, 2 + len(expenses_header)):
            cell = ws_expenses.cell(row=2, column=col_idx)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border
        # Style header row for income (row 2)
        for col_idx in range(income_start_col, income_start_col + len(income_header)):
            cell = ws_expenses.cell(row=2, column=col_idx)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border
        # Add border around outside of expenses table
        last_row_exp = 2 + len(expenses_data) - 1
        last_col_exp = 2 + len(expenses_header) - 1
        outside_border = Border(
            left=Side(style='medium'),
            right=Side(style='medium'),
            top=Side(style='medium'),
            bottom=Side(style='medium')
        )
        for row in range(2, last_row_exp + 1):
            for col in range(2, last_col_exp + 1):
                cell = ws_expenses.cell(row=row, column=col)
                if col == 2:
                    cell.border = cell.border.copy(left=outside_border.left)
                if col == last_col_exp:
                    cell.border = cell.border.copy(right=outside_border.right)
                if row == 2:
                    cell.border = cell.border.copy(top=outside_border.top)
                if row == last_row_exp:
                    cell.border = cell.border.copy(bottom=outside_border.bottom)
        # Add border around outside of income table
        last_row_inc = 2 + len(income_data) - 1
        last_col_inc = income_start_col + len(income_header) - 1
        for row in range(2, last_row_inc + 1):
            for col in range(income_start_col, last_col_inc + 1):
                cell = ws_expenses.cell(row=row, column=col)
                if col == income_start_col:
                    cell.border = cell.border.copy(left=outside_border.left)
                if col == last_col_inc:
                    cell.border = cell.border.copy(right=outside_border.right)
                if row == 2:
                    cell.border = cell.border.copy(top=outside_border.top)
                if row == last_row_inc:
                    cell.border = cell.border.copy(bottom=outside_border.bottom)


        # Category Summary Sheet
        # Prepare summary data
        target_handler = CategoryTargetHandler()
        summary_header = ["Category", "Target Amount", "Total Spent", "% of Total Income"]
        summary_data = [summary_header]
        category_totals = {}
        
        for statement in statements:
            for t in statement.transactions:
                if t.category not in [Category.NONE]:
                    category_totals[t.category.value] = category_totals.get(t.category.value, 0.0) + t.amount

        # Use all categories in target_handler for summary
        total_income = category_totals.get(Category.INCOME.value, 1)
        total_target = 0.0
        total_spent = 0.0
        total_percent_income = 0.0
        for cat in target_handler.targets:
            if cat == Category.INCOME.value:
                continue
            target_amt = target_handler.get_target(cat)
            total = abs(category_totals.get(cat, 0.0))
            percent_income = (total / total_income) if total_income > 0 else 0.0
            summary_data.append([cat, target_amt, total, percent_income]) # type: ignore
            total_target += target_amt
            total_spent += total
            total_percent_income += percent_income
        # Add Totals row
        summary_data.append(["Totals", total_target, total_spent, total_percent_income]) # type: ignore

        # Write summary data starting at B2
        totals_row_idx = 1 + len(summary_data)
        for row_idx, row_data in enumerate(summary_data, start=2):
            for col_idx, value in enumerate(row_data, start=2):
                # Write numbers as numbers, not strings
                if row_idx > 2 and col_idx in [3, 4, 5]:
                    value = float(value)
                cell = ws_summary.cell(row=row_idx, column=col_idx, value=value)
                # Bold the Totals row
                if row_idx == totals_row_idx:
                    cell.font = Font(bold=True)
                # Format target amount and total spent columns as currency
                if row_idx > 2 and col_idx == 3:
                    cell.number_format = '$#,##0.00'
                if row_idx > 2 and col_idx == 4:
                    cell.number_format = '$#,##0.00'
                    # Color code total spent relative to target
                    target_amt = float(row_data[1])
                    total_spent = float(row_data[2])
                    is_income = (row_data[0] == Category.INCOME.value) or (row_data[0] == Category.SAVINGS.value)
                    fill_color = get_relative_color(total_spent, target_amt, is_income)
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                if row_idx > 2 and col_idx == 5:
                    cell.number_format = '0.00%'
                    # Show as percent (divide by 100 for Excel percent format)
                    cell.value = value if value is not None else 0.0
        # Auto-size columns for Summary sheet, skip column A
        for col in range(2, ws_summary.max_column + 1):
            max_length = 0
            col_letter = get_column_letter(col)
            for cell in ws_summary[col_letter]:
                try:
                    cell_length = len(str(cell.value)) if cell.value is not None else 0
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            ws_summary.column_dimensions[col_letter].width = max_length + 2
        # Style header row (now at row 2)
        for col_idx in range(2, 2 + len(summary_header)):
            cell = ws_summary.cell(row=2, column=col_idx)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border
        # Add border around outside of summary table
        last_row = 2 + len(summary_data) - 1
        last_col = 2 + len(summary_header) - 1
        for row in range(2, last_row + 1):
            for col in range(2, last_col + 1):
                cell = ws_summary.cell(row=row, column=col)
                if col == 2:
                    cell.border = cell.border.copy(left=outside_border.left)
                if col == last_col:
                    cell.border = cell.border.copy(right=outside_border.right)
                if row == 2:
                    cell.border = cell.border.copy(top=outside_border.top)
                if row == last_row:
                    cell.border = cell.border.copy(bottom=outside_border.bottom)

        # Add Income Summary Table next to the Category Summary Table
        # Calculate Project Income (sum of targets for income categories)
        project_income = target_handler.get_target(Category.INCOME.value)
        # Calculate Total Income (actual sum from transactions)
        total_income = abs(category_totals.get(Category.INCOME.value, 0.0))
        difference = total_income - project_income

        # Table headers and values
        income_summary_header = ["Project Income", "Total Income", "Difference"]
        income_summary_values = [project_income, total_income, difference]

        # Place table to the right of Category Summary table
        # Find the first empty column after summary table
        income_table_start_col = last_col + 3  # leave one column gap
        income_table_start_row = 2

        # Write headers
        for col_idx, header in enumerate(income_summary_header, start=income_table_start_col):
            cell = ws_summary.cell(row=income_table_start_row, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")
            cell.border = header_border
        # Write values
        for col_idx, value in enumerate(income_summary_values, start=income_table_start_col):
            cell = ws_summary.cell(row=income_table_start_row + 1, column=col_idx, value=value)
            # Format as currency for Project Income and Total Income
            if col_idx in [income_table_start_col, income_table_start_col + 1]:
                cell.number_format = '$#,##0.00'
            # Color code Difference using gradient logic
            if col_idx == income_table_start_col + 2:
                cell.number_format = '$#,##0.00'
                fill_color = get_gradient_color(difference)
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border
        # Add border around outside of Income Summary table
        last_col_income = income_table_start_col + len(income_summary_header) - 1
        for row in range(income_table_start_row, income_table_start_row + 2):
            for col in range(income_table_start_col, last_col_income + 1):
                cell = ws_summary.cell(row=row, column=col)
                if col == income_table_start_col:
                    cell.border = cell.border.copy(left=outside_border.left)
                if col == last_col_income:
                    cell.border = cell.border.copy(right=outside_border.right)
                if row == income_table_start_row:
                    cell.border = cell.border.copy(top=outside_border.top)
                if row == income_table_start_row + 1:
                    cell.border = cell.border.copy(bottom=outside_border.bottom)

        # Add Total Income/Spent/Difference table beneath Income Summary table
        total_table_start_row = income_table_start_row + 3
        total_table_start_col = income_table_start_col
        total_table_header = ["Total Income", "Total Spent", "Difference"]
        total_spent_value = total_spent
        total_table_values = [total_income, total_spent_value, total_income - total_spent_value]

        # Write headers
        for col_idx, header in enumerate(total_table_header, start=total_table_start_col):
            cell = ws_summary.cell(row=total_table_start_row, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")
            cell.border = header_border
        # Write values
        for col_idx, value in enumerate(total_table_values, start=total_table_start_col):
            cell = ws_summary.cell(row=total_table_start_row + 1, column=col_idx, value=value)
            cell.number_format = '$#,##0.00'
            # Color code Difference using gradient logic
            if col_idx == total_table_start_col + 2:
                fill_color = get_gradient_color(total_table_values[2])
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border
        # Add border around outside of Total Income/Spent/Difference table
        last_col_total = total_table_start_col + len(total_table_header) - 1
        for row in range(total_table_start_row, total_table_start_row + 2):
            for col in range(total_table_start_col, last_col_total + 1):
                cell = ws_summary.cell(row=row, column=col)
                if col == total_table_start_col:
                    cell.border = cell.border.copy(left=outside_border.left)
                if col == last_col_total:
                    cell.border = cell.border.copy(right=outside_border.right)
                if row == total_table_start_row:
                    cell.border = cell.border.copy(top=outside_border.top)
                if row == total_table_start_row + 1:
                    cell.border = cell.border.copy(bottom=outside_border.bottom)
        

        # Add pie chart for category percentages (only purchases)
        first_data_row = 3  # B3 is first category
        last_data_row = 1 + len(summary_data) - 1  # Exclude Totals row
        category_range = Reference(ws_summary, min_col=2, min_row=first_data_row, max_row=last_data_row)
        total_range = Reference(ws_summary, min_col=4, min_row=first_data_row, max_row=last_data_row)
        pie = PieChart()
        pie.title = "Spending by Category (%)"
        pie.add_data(total_range, titles_from_data=False)
        pie.set_categories(category_range)
        if pie.legend is not None:
            pie.legend.position = 'r'  # Move legend to the right

        # Place pie chart below the table
        chart_row = last_row + 2
        chart_cell = f"B{chart_row}"
        ws_summary.add_chart(pie, chart_cell)

        # Add a column of absolute running total to the Plotting Data sheet
        ws_plotting_data.cell(row=1, column=1, value="Absolute Running Total")
        for idx, abs_val in enumerate(running_total_abs_values, start=2):
            ws_plotting_data.cell(row=idx, column=1, value=abs_val)

        # Add line graph to the right of the pie chart
        if expenses:
            # Use absolute running total from ws_plotting_data
            first_expense_row = 2  # Row 2 is first data row in ws_plotting_data
            last_expense_row = 1 + len(running_total_abs_values)
            abs_running_total_ref = Reference(ws_plotting_data, min_col=1, min_row=first_expense_row, max_row=last_expense_row)
            # Use dates from ws_expenses (column 3, C) for x-axis
            date_ref = Reference(ws_expenses, min_col=3, min_row=3, max_row=2 + len(expenses_data))
            line_chart = LineChart()
            line_chart.title = "Cumulative Spending Over Time (Absolute)"
            line_chart.style = 13
            line_chart.y_axis.title = "Total Spent ($)"
            line_chart.x_axis.title = "Date"
            line_chart.add_data(abs_running_total_ref, titles_from_data=False)
            line_chart.set_categories(date_ref)
            line_chart.x_axis.number_format = 'mm/dd/yyyy'
            line_chart.legend = None  # Remove legend
            # Place line chart to the right of pie chart
            line_chart_col = 2 + len(summary_header) + 1  # adjust as needed for spacing
            line_chart_cell = f"{get_column_letter(line_chart_col)}{chart_row}"
            ws_summary.add_chart(line_chart, line_chart_cell)

        # Write summary data starting at B2
        for row_idx, row_data in enumerate(summary_data, start=2):
            for col_idx, value in enumerate(row_data, start=2):
                cell = ws_summary.cell(row=row_idx, column=col_idx, value=value)
                # Format total spent column as currency
                if row_idx > 2 and col_idx == 3:
                    cell.number_format = '$#,##0.00'
        # Auto-size columns for Summary sheet, skip column A
        for col in range(2, ws_summary.max_column + 1):
            max_length = 0
            col_letter = get_column_letter(col)
            for cell in ws_summary[col_letter]:
                try:
                    cell_length = len(str(cell.value)) if cell.value is not None else 0
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            ws_summary.column_dimensions[col_letter].width = max_length + 2
        # Style header row (now at row 2)
        for col_idx in range(2, 2 + len(summary_header)):
            cell = ws_summary.cell(row=2, column=col_idx)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border
        # Add border around outside of summary table
        last_row = 2 + len(summary_data) - 1
        last_col = 2 + len(summary_header) - 1
        for row in range(2, last_row + 1):
            for col in range(2, last_col + 1):
                cell = ws_summary.cell(row=row, column=col)
                if col == 2:
                    cell.border = cell.border.copy(left=outside_border.left)
                if col == last_col:
                    cell.border = cell.border.copy(right=outside_border.right)
                if row == 2:
                    cell.border = cell.border.copy(top=outside_border.top)
                if row == last_row:
                    cell.border = cell.border.copy(bottom=outside_border.bottom)

        wb.save(output_path)
