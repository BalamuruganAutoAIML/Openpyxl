# OpenPyXL Learning Materials

This folder contains sample Excel and CSV files designed to help you learn the OpenPyXL library for Python. Each file demonstrates different features and use cases.

## 📁 File Structure

### CSV Files (Import Examples)
These files can be imported into Excel or converted to .xlsx format:

1. **1_basic_data.csv**
   - Basic employee information
   - Use case: Reading CSV, converting to Excel, working with simple tables
   - Topics: `openpyxl.utils`, CSV import, basic cell operations

2. **2_sales_data.csv**
   - Sales transactions with dates, products, and amounts
   - Use case: Working with dates, numbers, formulas
   - Topics: Date formatting, number formats, calculations

3. **3_student_grades.csv**
   - Student academic records with multiple scores
   - Use case: Data analysis, averaging, conditional formatting
   - Topics: Formulas (AVERAGE, SUM), conditional logic

4. **4_inventory_data.csv**
   - Inventory management with suppliers
   - Use case: Working with references, filtering
   - Topics: Data validation, lookups

5. **5_monthly_expenses.csv**
   - Expense tracking data
   - Use case: Categorization, status tracking
   - Topics: Data grouping, filtering, formatting

### Excel Files (Pre-built Examples)

6. **6_formatted_employees.xlsx**
   - Formatted employee data with styling
   - **Topics covered:**
     - Font styling (bold, color, size)
     - Cell fills and colors
     - Number formatting ($currency)
     - Column width adjustment
     - Borders (thin, thick)
     - Cell alignment (horizontal, vertical)
     - Freeze panes
   - **Learning code:**
     ```python
     from openpyxl import Workbook
     from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
     
     wb = Workbook()
     ws = wb.active
     
     # Add headers with styling
     header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
     header_font = Font(bold=True, color="FFFFFF", size=12)
     
     for cell in ws[1]:
         cell.fill = header_fill
         cell.font = header_font
     
     # Freeze header row
     ws.freeze_panes = "A2"
     ```

7. **7_multiple_sheets.xlsx**
   - Workbook with multiple sheets
   - **Topics covered:**
     - Creating multiple worksheets
     - Sheet naming and indexing
     - Cross-sheet formulas (=Sheet!CellRef)
     - Summary calculations
     - Working with different sheet objects
   - **Learning code:**
     ```python
     wb = Workbook()
     wb.remove(wb.active)  # Remove default sheet
     
     ws1 = wb.create_sheet("Sales", 0)
     ws2 = wb.create_sheet("Summary", 1)
     
     # Cross-sheet formula
     ws2["B2"] = "=SUM(Sales!E2:E6)"
     ```

8. **8_inventory_with_formulas.xlsx**
   - Inventory with IF formulas and conditional logic
   - **Topics covered:**
     - Cell formulas (IF, SUM, AVERAGE)
     - Formula references
     - Conditional logic in cells
     - Merged cells
     - Multiple sheets with relationships
   - **Learning code:**
     ```python
     # Add formula to check reorder status
     ws[f"H{row}"] = f'=IF(D{row}<E{row},"REORDER","OK")'
     
     # Merge cells
     ws.merge_cells('A1:H1')
     ```

9. **9_form_with_validation.xlsx**
   - Data entry form with validation rules
   - **Topics covered:**
     - Data validation (list dropdowns)
     - Data validation (numeric ranges)
     - Error messages
     - Form layout design
     - Cell borders for input fields
     - Sample data insertion
   - **Learning code:**
     ```python
     from openpyxl.worksheet.datavalidation import DataValidation
     
     # Dropdown validation
     dv = DataValidation(type="list", 
                        formula1='"Sales,Marketing,IT,HR,Finance"')
     ws.add_data_validation(dv)
     dv.add('B4')
     
     # Numeric range validation
     dv2 = DataValidation(type="whole", operator="between", 
                         formula1="1", formula2="5")
     ws.add_data_validation(dv2)
     ```

10. **10_charts_and_graphs.xlsx**
    - Sales data with bar and pie charts
    - **Topics covered:**
        - Creating charts (BarChart, PieChart, LineChart)
        - Chart titles and axis labels
        - Data references for charts
        - Adding charts to worksheets
        - Chart positioning
    - **Learning code:**
        ```python
        from openpyxl.chart import BarChart, PieChart, Reference
        
        # Create chart
        bar_chart = BarChart()
        bar_chart.type = "col"
        bar_chart.title = "Sales by Product"
        
        # Add data
        data_ref = Reference(ws, min_col=2, min_row=1, 
                            max_col=4, max_row=6)
        categories = Reference(ws, min_col=1, min_row=2, max_row=6)
        bar_chart.add_data(data_ref, titles_from_data=True)
        bar_chart.set_categories(categories)
        
        # Add to worksheet
        ws.add_chart(bar_chart, "A9")
        ```

## 🎯 Learning Path

### Beginner Level
1. Start with **6_formatted_employees.xlsx**
   - Learn basic styling and formatting
   - Understand cell formatting options

2. Convert **1_basic_data.csv** to Excel
   - Practice reading CSV files
   - Learn to write Excel files

### Intermediate Level
3. Work with **7_multiple_sheets.xlsx**
   - Create multi-sheet workbooks
   - Use cross-sheet formulas

4. Study **8_inventory_with_formulas.xlsx**
   - Learn conditional formulas
   - Practice cell merging

5. Explore **9_form_with_validation.xlsx**
   - Implement data validation
   - Create forms in Excel

### Advanced Level
6. Create **10_charts_and_graphs.xlsx**
   - Generate charts programmatically
   - Learn chart customization

## 🔧 Common OpenPyXL Operations Quick Reference

### Reading and Writing
```python
from openpyxl import Workbook, load_workbook

# Create new workbook
wb = Workbook()
ws = wb.active

# Load existing workbook
wb = load_workbook('file.xlsx')
ws = wb['Sheet Name']

# Save workbook
wb.save('output.xlsx')
```

### Cell Operations
```python
# Read cell value
value = ws['A1'].value
# or
value = ws.cell(row=1, column=1).value

# Write cell value
ws['A1'] = 'Hello'
ws.cell(row=1, column=1, value='Hello')

# Append row
ws.append(['A', 'B', 'C'])

# Iterate cells
for row in ws.iter_rows(min_row=1, max_row=5, min_col=1, max_col=3):
    for cell in row:
        print(cell.value)
```

### Styling
```python
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Font
cell.font = Font(name='Arial', size=12, bold=True, italic=True, color='FF0000')

# Fill (Background Color)
cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

# Alignment
cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Border
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
cell.border = thin_border

# Number Format
cell.number_format = '$#,##0.00'  # Currency
cell.number_format = '0%'          # Percentage
cell.number_format = 'YYYY-MM-DD'  # Date
```

### Sheet Operations
```python
# Create sheet
ws = wb.create_sheet('New Sheet', 0)  # Insert at position 0

# Rename sheet
ws.title = 'Renamed Sheet'

# Remove sheet
del wb['Sheet Name']

# Get all sheet names
sheet_names = wb.sheetnames

# Freeze panes
ws.freeze_panes = 'A2'  # Freeze above row 2

# Merge cells
ws.merge_cells('A1:C1')
```

### Formulas
```python
# Write formula
ws['D2'] = '=SUM(A2:C2)'
ws['E2'] = '=AVERAGE(A2:C2)'
ws['F2'] = f'=IF(D2>100,"High","Low")'

# Cross-sheet formula
ws2['B1'] = '=SUM(Sheet1!A1:A10)'
```

### Data Validation
```python
from openpyxl.worksheet.datavalidation import DataValidation

# List dropdown
dv = DataValidation(type="list", formula1='"Option1,Option2,Option3"')
ws.add_data_validation(dv)
dv.add('A1')

# Numeric range
dv = DataValidation(type="whole", operator="between", 
                    formula1="1", formula2="100")
ws.add_data_validation(dv)
dv.add('A1')

# Text length
dv = DataValidation(type="textLength", operator="lessThanOrEqual", 
                    formula1="50")
ws.add_data_validation(dv)
dv.add('A1')
```

### Charts
```python
from openpyxl.chart import BarChart, PieChart, LineChart, Reference

# Create chart
chart = BarChart()
chart.type = "col"  # Column chart
chart.title = "Sales"
chart.x_axis.title = "Month"
chart.y_axis.title = "Amount"

# Add data
data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=10)
cats = Reference(ws, min_col=1, min_row=2, max_row=10)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Add to worksheet
ws.add_chart(chart, "E2")
```

## 💡 Practice Exercises

1. **Create from CSV to Excel**
   - Load 1_basic_data.csv
   - Convert to Excel format
   - Add formatting

2. **Multi-Sheet Workbook**
   - Create a workbook with Sales, Inventory, and Summary sheets
   - Add cross-sheet formulas
   - Create summary statistics

3. **Data Validation Form**
   - Build a data entry form using 9_form_with_validation.xlsx as reference
   - Add multiple validation rules
   - Create linked dropdowns

4. **Generate Reports**
   - Read data from CSV
   - Create formatted Excel report
   - Add charts and summary statistics
   - Include professional styling

5. **Batch Processing**
   - Create multiple files from data
   - Apply consistent formatting
   - Save with automatic naming

## 📚 Resources

- Official OpenPyXL Documentation: https://openpyxl.readthedocs.io/
- PyPI Package: https://pypi.org/project/openpyxl/
- GitHub Repository: https://github.com/appleseedhq/openpyxl

## 🎓 Topics Covered by These Files

| Feature | File | Difficulty |
|---------|------|------------|
| Basic Reading/Writing | 6 | Beginner |
| Styling & Formatting | 6 | Beginner |
| Multiple Sheets | 7 | Intermediate |
| Formulas | 8 | Intermediate |
| Merging Cells | 8 | Intermediate |
| Data Validation | 9 | Intermediate |
| Charts & Graphs | 10 | Advanced |
| Forms & Layout | 9 | Advanced |

Happy Learning! 🚀
