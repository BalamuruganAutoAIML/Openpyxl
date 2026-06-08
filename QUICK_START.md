# OpenPyXL Quick Start Guide

## Installation

```bash
pip install openpyxl
```

## 🚀 5-Minute Quick Start

### 1. Create Your First Excel File

```python
from openpyxl import Workbook

# Create a new workbook
wb = Workbook()
ws = wb.active  # Get the active sheet

# Add data
ws['A1'] = "Hello"
ws['B1'] = "World"
ws['A2'] = 123
ws['B2'] = 456

# Save
wb.save('my_file.xlsx')
```

### 2. Read an Existing Excel File

```python
from openpyxl import load_workbook

# Load workbook
wb = load_workbook('my_file.xlsx')
ws = wb['Sheet']  # Get specific sheet

# Read data
value = ws['A1'].value  # Get single cell
```

### 3. Add Formatting

```python
from openpyxl.styles import Font, PatternFill, Alignment

cell = ws['A1']

# Font formatting
cell.font = Font(bold=True, size=14, color="FF0000")

# Background color
cell.fill = PatternFill(start_color="FFFF00", fill_type="solid")

# Alignment
cell.alignment = Alignment(horizontal="center", vertical="center")
```

## 📊 File Usage Guide

Open these files in Excel to see examples, then study the code that created them:

| File | Topic | Difficulty |
|------|-------|------------|
| **6_formatted_employees.xlsx** | Colors, fonts, borders | ⭐ Beginner |
| **7_multiple_sheets.xlsx** | Multiple sheets, formulas | ⭐⭐ Intermediate |
| **8_inventory_with_formulas.xlsx** | IF formulas, merged cells | ⭐⭐ Intermediate |
| **9_form_with_validation.xlsx** | Data validation, dropdowns | ⭐⭐ Intermediate |
| **10_charts_and_graphs.xlsx** | Charts, bar graphs, pie charts | ⭐⭐⭐ Advanced |

**CSV Files** (Import practice):
- `1_basic_data.csv` - Simple employee data
- `2_sales_data.csv` - Transaction data with dates
- `3_student_grades.csv` - Grade records
- `4_inventory_data.csv` - Inventory management
- `5_monthly_expenses.csv` - Expense tracking

## 🎯 Common Tasks

### Working with Multiple Sheets

```python
# Create new sheet
ws = wb.create_sheet("New Sheet")

# Access specific sheet
ws = wb["Sheet Name"]

# Remove sheet
del wb["Sheet Name"]

# Get all sheet names
names = wb.sheetnames
```

### Add Formulas

```python
# Simple formula
ws['D2'] = '=B2+C2'

# Function formula
ws['E2'] = '=SUM(A2:A10)'

# Cross-sheet formula
ws2['B1'] = '=Sheet1!A1'

# IF formula
ws['C1'] = '=IF(B1>100,"High","Low")'
```

### Data Validation

```python
from openpyxl.worksheet.datavalidation import DataValidation

# Dropdown list
dv = DataValidation(type="list", formula1='"Yes,No,Maybe"')
ws.add_data_validation(dv)
dv.add('B1')

# Numeric range (1-10)
dv = DataValidation(type="whole", operator="between", 
                    formula1="1", formula2="10")
ws.add_data_validation(dv)
dv.add('B1')
```

### Charts

```python
from openpyxl.chart import BarChart, Reference

# Create chart
chart = BarChart()
chart.title = "Sales"
chart.type = "col"  # Column chart

# Add data
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=10)
chart.add_data(data, titles_from_data=True)

# Add to sheet
ws.add_chart(chart, "E2")
```

## 📈 Learning Path

### Week 1: Basics
- [ ] Open `6_formatted_employees.xlsx` in Excel
- [ ] Read the code in `example_scripts.py` Example 1 & 2
- [ ] Create a simple workbook with formatted headers

### Week 2: Intermediate
- [ ] Study `7_multiple_sheets.xlsx`
- [ ] Follow Example 3, 4, 5 in `example_scripts.py`
- [ ] Create a multi-sheet workbook with formulas
- [ ] Add data validation to a form

### Week 3: Advanced
- [ ] Examine `10_charts_and_graphs.xlsx`
- [ ] Complete Example 6, 7, 8 in `example_scripts.py`
- [ ] Create a report with charts
- [ ] Convert CSV to formatted Excel

## ⚡ Useful Patterns

### Apply Formatting to Multiple Cells

```python
# Apply to entire row
for cell in ws[1]:
    cell.font = Font(bold=True)

# Apply to range
for row in ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=5):
    for cell in row:
        cell.alignment = Alignment(horizontal="center")
```

### Iterate Through Data

```python
# Iterate by row
for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
    print(row)  # Returns tuple of values

# Iterate specific cells
for row in ws['A1:C5']:
    for cell in row:
        print(cell.value)
```

### Merge Cells

```python
ws.merge_cells('A1:D1')  # Merge A1 to D1
```

### Number Formatting

```python
ws['A1'].number_format = '$#,##0.00'    # Currency
ws['A1'].number_format = '0%'           # Percentage
ws['A1'].number_format = '0.00'         # Decimal
ws['A1'].number_format = 'YYYY-MM-DD'   # Date
```

## 🔗 Useful Resources

- **Official Docs**: https://openpyxl.readthedocs.io/
- **Tutorial**: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- **API Reference**: https://openpyxl.readthedocs.io/en/stable/api.html

## 💡 Pro Tips

1. **Always save after modifications**
   ```python
   wb.save('output.xlsx')
   ```

2. **Use `values_only=True` when reading**
   ```python
   for row in ws.iter_rows(values_only=True):
       # row contains values, not cell objects
   ```

3. **Freeze header rows for easier navigation**
   ```python
   ws.freeze_panes = 'A2'  # Freeze first row
   ```

4. **Use meaningful variable names**
   ```python
   # Good
   sales_data_sheet = wb['Sales']
   employee_name_cell = ws['A2']
   
   # Bad
   ws1 = wb['Sales']
   c = ws['A2']
   ```

5. **Validate data before writing**
   ```python
   # Check before saving
   if ws['A1'].value is not None:
       wb.save('file.xlsx')
   ```

## 🐛 Common Issues & Solutions

**Issue**: "No module named 'openpyxl'"
- **Solution**: `pip install openpyxl`

**Issue**: Changes not saved
- **Solution**: Remember to call `wb.save(filename)`

**Issue**: Formula not calculating
- **Solution**: Reopen file in Excel to recalculate formulas

**Issue**: Merged cells not displaying correctly
- **Solution**: Ensure you're merging within the worksheet bounds

## 🎓 Next Steps

1. **Run the example scripts**
   ```bash
   python example_scripts.py
   ```

2. **Modify existing examples**
   - Change colors, fonts, data
   - Add more sheets or charts

3. **Build real projects**
   - Convert CSV to formatted Excel
   - Create an expense tracker
   - Build an inventory management system
   - Generate automated reports

4. **Explore advanced topics**
   - Conditional formatting
   - Named ranges
   - Pivot tables (limited support)
   - VBA macros (read-only)

Happy coding! 🚀
