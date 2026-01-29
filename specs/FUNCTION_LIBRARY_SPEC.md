# Excel Formula Function Library Specification

## Overview

This document specifies the implementation details for Excel functions. Each function includes signature, behavior, edge cases, and Excel compatibility notes.

## Function Categories

### 1. Mathematical Functions

#### SUM
**Signature**: `SUM(number1, [number2], ...)`

**Description**: Adds all numbers in a range

**Implementation**:
```python
def SUM(*args):
    """
    Sum all numeric values, ignoring text and errors
    
    Excel behavior:
    - Ignores text values
    - Ignores logical values in ranges
    - Includes logical values if passed directly: SUM(TRUE) = 1
    - Propagates errors
    - Empty cells = 0
    """
    total = 0
    for arg in args:
        if is_error(arg):
            return arg
        if is_range(arg):
            for cell in flatten_range(arg):
                if is_number(cell):
                    total += cell
        elif is_number(arg):
            total += arg
        elif is_boolean(arg):
            total += (1 if arg else 0)
    return total
```

**Test Cases**:
- `SUM(1,2,3)` → `6`
- `SUM(A1:A3)` where A1=1, A2="text", A3=3 → `4`
- `SUM(TRUE, FALSE)` → `1`
- `SUM(1, #DIV/0!)` → `#DIV/0!`

---

#### AVERAGE
**Signature**: `AVERAGE(number1, [number2], ...)`

**Description**: Returns the average of its arguments

**Implementation**:
```python
def AVERAGE(*args):
    """
    Average of numeric values
    
    Excel behavior:
    - Ignores text, logical values, empty cells in ranges
    - Includes logical values if passed directly
    - Returns #DIV/0! if no numeric values
    """
    total = 0
    count = 0
    
    for arg in args:
        if is_error(arg):
            return arg
        if is_range(arg):
            for cell in flatten_range(arg):
                if is_number(cell):
                    total += cell
                    count += 1
        elif is_number(arg):
            total += arg
            count += 1
        elif is_boolean(arg):
            total += (1 if arg else 0)
            count += 1
    
    if count == 0:
        return ExcelError("#DIV/0!")
    
    return total / count
```

---

#### ROUND
**Signature**: `ROUND(number, num_digits)`

**Description**: Rounds a number to specified digits

**Implementation**:
```python
def ROUND(number, num_digits):
    """
    Round number to num_digits decimal places
    
    Excel behavior:
    - Positive num_digits: decimal places
    - Negative num_digits: round to left of decimal
    - Uses "round half up" (banker's rounding in some versions)
    """
    if is_error(number) or is_error(num_digits):
        return first_error(number, num_digits)
    
    number = to_number(number)
    num_digits = int(to_number(num_digits))
    
    if is_error(number) or is_error(num_digits):
        return first_error(number, num_digits)
    
    # Excel uses round-half-up
    multiplier = 10 ** num_digits
    return math.floor(number * multiplier + 0.5) / multiplier
```

**Test Cases**:
- `ROUND(2.15, 1)` → `2.2`
- `ROUND(2.149, 1)` → `2.1`
- `ROUND(21.5, -1)` → `20`
- `ROUND(-50.55, -2)` → `-100`

---

### 2. Logical Functions

#### IF
**Signature**: `IF(logical_test, value_if_true, [value_if_false])`

**Description**: Returns one value if condition is TRUE, another if FALSE

**Implementation**:
```python
def IF(logical_test, value_if_true, value_if_false=False):
    """
    Conditional evaluation
    
    Excel behavior:
    - Evaluates logical_test first
    - Only evaluates the branch that will be returned (lazy evaluation)
    - If value_if_false omitted, returns FALSE
    - Coerces logical_test to boolean
    """
    if is_error(logical_test):
        return logical_test
    
    condition = to_boolean(logical_test)
    
    if is_error(condition):
        return condition
    
    # Lazy evaluation: only evaluate the branch we need
    if condition:
        return evaluate_lazy(value_if_true)
    else:
        return evaluate_lazy(value_if_false)
```

**Note**: This function requires lazy evaluation - arguments should be AST nodes, not pre-evaluated values.

---

#### AND
**Signature**: `AND(logical1, [logical2], ...)`

**Description**: Returns TRUE if all arguments are TRUE

**Implementation**:
```python
def AND(*args):
    """
    Logical AND
    
    Excel behavior:
    - Returns TRUE if all arguments are TRUE
    - Ignores text and empty cells in ranges
    - Coerces direct arguments to boolean
    - Returns #VALUE! if no logical values found
    """
    found_any = False
    
    for arg in args:
        if is_error(arg):
            return arg
        if is_range(arg):
            for cell in flatten_range(arg):
                if is_boolean(cell) or is_number(cell):
                    found_any = True
                    if not to_boolean(cell):
                        return False
        else:
            found_any = True
            if not to_boolean(arg):
                return False
    
    if not found_any:
        return ExcelError("#VALUE!")
    
    return True
```

---

### 3. Text Functions

#### CONCATENATE
**Signature**: `CONCATENATE(text1, [text2], ...)`

**Description**: Joins several text strings into one

**Implementation**:
```python
def CONCATENATE(*args):
    """
    Concatenate text strings
    
    Excel behavior:
    - Converts numbers to text
    - Converts booleans to "TRUE"/"FALSE"
    - Propagates errors
    - Does not accept ranges (returns #VALUE!)
    """
    result = []
    
    for arg in args:
        if is_error(arg):
            return arg
        if is_range(arg):
            return ExcelError("#VALUE!")
        result.append(to_text(arg))
    
    return ''.join(result)
```

**Modern Alternative**: `CONCAT` (accepts ranges), `TEXTJOIN` (with delimiter)

---

#### LEFT
**Signature**: `LEFT(text, [num_chars])`

**Description**: Returns leftmost characters from text

**Implementation**:
```python
def LEFT(text, num_chars=1):
    """
    Extract leftmost characters
    
    Excel behavior:
    - Default num_chars is 1
    - num_chars must be >= 0
    - If num_chars > length, returns entire text
    """
    if is_error(text) or is_error(num_chars):
        return first_error(text, num_chars)
    
    text = to_text(text)
    num_chars = int(to_number(num_chars))
    
    if num_chars < 0:
        return ExcelError("#VALUE!")
    
    return text[:num_chars]
```

---

### 4. Date/Time Functions

#### TODAY
**Signature**: `TODAY()`

**Description**: Returns current date (volatile)

**Implementation**:
```python
def TODAY():
    """
    Current date as Excel serial number
    
    Excel behavior:
    - Returns date only (no time component)
    - Serial number: days since 1900-01-01
    - Volatile: recalculates on every change
    - Excel has 1900 leap year bug (treats 1900 as leap year)
    """
    from datetime import date
    
    # Excel epoch: 1900-01-01 (but with leap year bug)
    excel_epoch = date(1899, 12, 30)  # Adjusted for bug
    today = date.today()
    
    delta = today - excel_epoch
    return delta.days
```

---

#### DATE
**Signature**: `DATE(year, month, day)`

**Description**: Returns serial number of a date

**Implementation**:
```python
def DATE(year, month, day):
    """
    Create date from components
    
    Excel behavior:
    - Year 0-1899: adds to 1900 (e.g., 108 = 2008)
    - Year 1900-9999: used as-is
    - Month > 12: rolls over to next year
    - Day > days in month: rolls over
    - Negative values: roll backwards
    """
    if is_error(year) or is_error(month) or is_error(day):
        return first_error(year, month, day)
    
    year = int(to_number(year))
    month = int(to_number(month))
    day = int(to_number(day))
    
    # Handle year 0-1899
    if 0 <= year <= 1899:
        year += 1900
    
    # Create date with rollover
    from datetime import date, timedelta
    
    # Start with first day of year/month
    d = date(year, 1, 1)
    
    # Add months
    d = add_months(d, month - 1)
    
    # Add days
    d = d + timedelta(days=day - 1)
    
    return date_to_serial(d)
```

---

### 5. Lookup Functions

#### VLOOKUP
**Signature**: `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

**Description**: Searches for value in first column and returns value from specified column

**Implementation**:
```python
def VLOOKUP(lookup_value, table_array, col_index_num, range_lookup=True):
    """
    Vertical lookup
    
    Excel behavior:
    - range_lookup=TRUE: approximate match (requires sorted data)
    - range_lookup=FALSE: exact match
    - Returns #N/A if not found
    - Returns #REF! if col_index_num > table width
    - Returns #VALUE! if col_index_num < 1
    """
    if is_error(lookup_value):
        return lookup_value
    
    if not is_range(table_array):
        return ExcelError("#VALUE!")
    
    col_index = int(to_number(col_index_num))
    
    if col_index < 1:
        return ExcelError("#VALUE!")
    
    table = get_range_as_2d_array(table_array)
    
    if col_index > len(table[0]):
        return ExcelError("#REF!")
    
    if range_lookup:
        # Approximate match (binary search on sorted data)
        return vlookup_approximate(lookup_value, table, col_index)
    else:
        # Exact match (linear search)
        for row in table:
            if compare_equal(row[0], lookup_value):
                return row[col_index - 1]
        return ExcelError("#N/A")
```

---

#### INDEX
**Signature**: `INDEX(array, row_num, [column_num])`

**Description**: Returns value at specified position in range

**Implementation**:
```python
def INDEX(array, row_num, column_num=None):
    """
    Return value at position
    
    Excel behavior:
    - row_num=0: return entire column
    - column_num=0: return entire row
    - 1-indexed
    - Returns #REF! if out of bounds
    """
    if not is_range(array):
        return ExcelError("#VALUE!")
    
    table = get_range_as_2d_array(array)
    rows = len(table)
    cols = len(table[0]) if rows > 0 else 0
    
    row_idx = int(to_number(row_num))
    
    # Single column range
    if cols == 1 and column_num is None:
        if row_idx < 1 or row_idx > rows:
            return ExcelError("#REF!")
        return table[row_idx - 1][0]
    
    # 2D range
    col_idx = 1 if column_num is None else int(to_number(column_num))
    
    if row_idx < 0 or row_idx > rows:
        return ExcelError("#REF!")
    if col_idx < 0 or col_idx > cols:
        return ExcelError("#REF!")
    
    # Return entire row
    if row_idx == 0:
        return [table[r][col_idx - 1] for r in range(rows)]
    
    # Return entire column
    if col_idx == 0:
        return table[row_idx - 1]
    
    return table[row_idx - 1][col_idx - 1]
```

---

#### MATCH
**Signature**: `MATCH(lookup_value, lookup_array, [match_type])`

**Description**: Returns relative position of item in array

**Implementation**:
```python
def MATCH(lookup_value, lookup_array, match_type=1):
    """
    Find position of value in array
    
    Excel behavior:
    - match_type=1: largest value <= lookup_value (requires sorted ascending)
    - match_type=0: exact match
    - match_type=-1: smallest value >= lookup_value (requires sorted descending)
    - Returns #N/A if not found
    - Returns 1-indexed position
    """
    if not is_range(lookup_array):
        return ExcelError("#VALUE!")
    
    array = flatten_range(lookup_array)
    match_type = int(to_number(match_type))
    
    if match_type == 0:
        # Exact match
        for i, val in enumerate(array):
            if compare_equal(val, lookup_value):
                return i + 1
        return ExcelError("#N/A")
    
    elif match_type == 1:
        # Largest value <= lookup_value
        best_idx = None
        for i, val in enumerate(array):
            if compare_lte(val, lookup_value):
                best_idx = i
            else:
                break
        if best_idx is None:
            return ExcelError("#N/A")
        return best_idx + 1
    
    elif match_type == -1:
        # Smallest value >= lookup_value
        for i, val in enumerate(array):
            if compare_gte(val, lookup_value):
                return i + 1
        return ExcelError("#N/A")
```

---

## Type Conversion Rules

### to_number()
```python
def to_number(value):
    """Convert value to number following Excel rules"""
    if is_number(value):
        return value
    if is_boolean(value):
        return 1 if value else 0
    if is_text(value):
        # Try to parse as number
        try:
            return float(value)
        except:
            return ExcelError("#VALUE!")
    if is_empty(value):
        return 0
    return ExcelError("#VALUE!")
```

### to_text()
```python
def to_text(value):
    """Convert value to text following Excel rules"""
    if is_text(value):
        return value
    if is_number(value):
        return format_number(value)
    if is_boolean(value):
        return "TRUE" if value else "FALSE"
    if is_error(value):
        return value  # Propagate error
    if is_empty(value):
        return ""
    return str(value)
```

### to_boolean()
```python
def to_boolean(value):
    """Convert value to boolean following Excel rules"""
    if is_boolean(value):
        return value
    if is_number(value):
        return value != 0
    if is_text(value):
        upper = value.upper()
        if upper == "TRUE":
            return True
        if upper == "FALSE":
            return False
        return ExcelError("#VALUE!")
    if is_empty(value):
        return False
    return ExcelError("#VALUE!")
```

## Comparison Rules

Excel has specific comparison rules:

```python
def compare_equal(a, b):
    """Excel equality comparison"""
    # Numbers
    if is_number(a) and is_number(b):
        return abs(a - b) < 1e-15  # Precision tolerance
    
    # Text (case-insensitive)
    if is_text(a) and is_text(b):
        return a.upper() == b.upper()
    
    # Boolean
    if is_boolean(a) and is_boolean(b):
        return a == b
    
    # Type mismatch
    return False
```

## Error Propagation

Most functions propagate errors:

```python
def first_error(*values):
    """Return first error in arguments"""
    for val in values:
        if is_error(val):
            return val
    return None
```

## Testing Strategy

### Test Each Function Against Excel

```python
def test_sum_basic():
    assert SUM(1, 2, 3) == 6

def test_sum_with_text():
    # In Excel: =SUM(A1:A3) where A1=1, A2="text", A3=3
    result = SUM(create_range([1, "text", 3]))
    assert result == 4

def test_sum_with_boolean():
    assert SUM(True, False) == 1

def test_vlookup_exact():
    table = create_range([
        ["A", 1],
        ["B", 2],
        ["C", 3]
    ])
    assert VLOOKUP("B", table, 2, False) == 2
    assert is_error(VLOOKUP("D", table, 2, False), "#N/A")
```

## Priority Order

Implement functions in this order:

### Phase 1 (Essential - Week 1)
1. SUM, AVERAGE, COUNT, MIN, MAX
2. IF, AND, OR, NOT
3. Basic arithmetic operators

### Phase 2 (Common - Week 2)
4. VLOOKUP, INDEX, MATCH
5. CONCATENATE, LEFT, RIGHT, MID, LEN
6. ROUND, ROUNDUP, ROUNDDOWN
7. TODAY, NOW, DATE

### Phase 3 (Advanced - Week 3)
8. SUMIF, COUNTIF, AVERAGEIF
9. IFERROR, IFNA
10. TEXT, VALUE
11. More date functions

### Phase 4 (Specialized - Week 4)
12. Statistical functions
13. Financial functions
14. Array functions

## References

- [Excel Function Reference](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188)
- [OpenFormula Specification](https://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2-part2.html)
