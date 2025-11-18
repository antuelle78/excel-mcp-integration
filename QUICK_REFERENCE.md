# Quick Reference - Code Completion Issues

## 3 CRITICAL FUNCTIONS TO FIX

### 1. `get_excel_info()` - src/main.py:218-247
**Current Problem**: Only checks file existence, doesn't read data
```python
# CURRENT (BROKEN):
wb = Workbook()
wb.close()  # We would need openpyxl to read...
return {"filename": safe_filename, "exists": True, ...}

# SHOULD BE:
wb = load_workbook(safe_filename)
return {
    "filename": safe_filename,
    "sheets": wb.sheetnames,
    "sheet_count": len(wb.sheetnames),
    "dimensions": {sheet_name: ws.dimensions for sheet_name, ws in ...},
    "active_sheet": wb.active.title,
    ...
}
```

---

### 2. `format_excel_cells()` - src/main.py:333-436
**Current Problem**: Hardcoded to format A1:E10 always
```python
# CURRENT (BROKEN):
for row_idx in range(1, 11):  # HARDCODED 1-10
    for col_idx in range(1, 6):  # HARDCODED 1-5

# SHOULD BE:
# Parse "A1:C10" format
def parse_cell_range(cell_range):
    parts = cell_range.split(':')
    start = parts[0]  # "A1"
    end = parts[1]    # "C10"
    
    start_col = ord(start[0]) - ord('A') + 1
    start_row = int(re.search(r'\d+', start).group())
    end_col = ord(end[0]) - ord('A') + 1
    end_row = int(re.search(r'\d+', end).group())
    
    return start_row, end_row, start_col, end_col

start_row, end_row, start_col, end_col = parse_cell_range(cell_range)
for row_idx in range(start_row, end_row + 1):
    for col_idx in range(start_col, end_col + 1):
        # Apply formatting to ws.cell(row=row_idx, column=col_idx)
```

---

### 3. `create_excel_chart()` - src/main.py:250-330
**Current Problem**: Ignores data_range parameter
```python
# CURRENT (BROKEN):
data = Reference(ws, min_col=1, min_row=1, max_col=ws.max_column, max_row=ws.max_row)

# SHOULD BE:
# Use parse_cell_range() function (same as above)
start_row, end_row, start_col, end_col = parse_cell_range(data_range)
data = Reference(ws, min_col=start_col, min_row=start_row, 
                 max_col=end_col, max_row=end_row)
```

---

## OTHER ISSUES

### Missing Type Hints (23 functions)
Focus on Flask route handlers:
```python
# BEFORE:
def create_excel():
    """..."""
    return jsonify({...})

# AFTER:
def create_excel() -> Dict[str, Any]:
    """..."""
    return jsonify({...})
```

### Missing Docstrings (3 FileHandler methods)
```python
# BEFORE:
def do_GET(self):
    # code...

# AFTER:
def do_GET(self) -> None:
    """Handle HTTP GET requests for file downloads."""
    # code...
```

---

## QUICK TEST COMMAND
```bash
# Run all tests
python tests/integration_test.py

# Check syntax
python3 -m py_compile src/*.py tests/*.py

# Run specific test
python tests/test_new_tools.py
```

---

## FILES TO EDIT
- ✏️  `src/main.py` - Fix 3 functions (218, 250, 333)
- ✏️  `src/simple_web_wrapper.py` - Add type hints (7 functions)
- ✏️  `src/web_api_wrapper.py` - Add type hints (5+ functions)

---

## ESTIMATED FIX TIME
- `get_excel_info()`: 20 min
- `format_excel_cells()`: 40 min  
- `create_excel_chart()`: 30 min
- Type hints: 60 min
- **Total: 90-120 minutes to production-ready**
