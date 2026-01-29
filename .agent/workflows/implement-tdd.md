---
description: Implement smallest change to increase test coverage
---

# Test-Driven Implementation Workflow

This workflow implements the smallest possible change to increase test coverage.

## Philosophy

- Make the smallest change that makes one more test pass
- Always run tests before and after
- Commit after each passing test
- Never skip tests

## Steps

### 1. Check Current Coverage

// turbo
```bash
pytest tests/ --cov=excel_engine --cov-report=term-missing
```

This shows which lines are not covered.

### 2. Find Next Failing Test

// turbo
```bash
pytest tests/ -x -v
```

The `-x` flag stops at first failure.

### 3. Identify Smallest Change

Look at the failing test and determine the minimal code needed:
- Is it a missing function?
- Is it a missing operator?
- Is it an edge case?
- Is it error handling?

### 4. Implement Minimal Fix

Write ONLY the code needed to make this ONE test pass.

Example:
```python
# If test is: test_sum_empty_range()
# And it fails because SUM doesn't handle empty ranges
# Add ONLY the empty range check:

def SUM(*args):
    total = 0
    for arg in args:
        if is_range(arg):
            cells = flatten_range(arg)
            if not cells:  # <- ONLY THIS LINE ADDED
                continue   # <- AND THIS LINE
            for cell in cells:
                if is_number(cell):
                    total += cell
    return total
```

### 5. Run Tests

// turbo
```bash
pytest tests/ -v
```

Verify:
- The failing test now passes
- No other tests broke
- Coverage increased

### 6. Check Coverage Increase

// turbo
```bash
pytest tests/ --cov=excel_engine --cov-report=term-missing
```

Confirm coverage went up.

### 7. Commit

```bash
git add .
git commit -m "feat: handle empty ranges in SUM (coverage: X% -> Y%)"
```

### 8. Repeat

Go back to step 1 and find the next failing test.

## Guidelines

### What is "Smallest Change"?

✅ **Good** (small changes):
- Add one if-statement for edge case
- Implement one operator
- Add one function
- Fix one type coercion rule

❌ **Bad** (too large):
- Implement entire function library at once
- Refactor multiple files
- Add features not tested yet

### When to Refactor

Only refactor when:
1. All tests pass
2. You see obvious duplication
3. Make refactor in separate commit

### Test Priority Order

1. **Parser tests** - Get formulas parsing
2. **Basic functions** - SUM, AVERAGE, COUNT
3. **Operators** - +, -, *, /, ^
4. **Logical functions** - IF, AND, OR
5. **Text functions** - CONCATENATE, LEFT, RIGHT
6. **Lookup functions** - VLOOKUP, INDEX, MATCH
7. **Advanced functions** - Everything else

## Metrics

Track these metrics:
- **Test coverage**: Should increase with each commit
- **Tests passing**: Should never decrease
- **Commit size**: Should be small (< 50 lines typically)

## Example Session

```bash
# 1. Check coverage
$ pytest --cov=excel_engine --cov-report=term
Coverage: 45%

# 2. Run tests
$ pytest -x -v
FAILED test_sum_with_text - Expected 6, got error

# 3. Implement fix (add text handling to SUM)
# ... edit code ...

# 4. Run tests again
$ pytest -v
All tests passed!

# 5. Check coverage
$ pytest --cov=excel_engine --cov-report=term
Coverage: 47%  # Increased by 2%

# 6. Commit
$ git commit -m "feat: SUM ignores text values (coverage: 45% -> 47%)"

# 7. Repeat
```

## Automation

Use this script to automate the workflow:

```bash
# scripts/tdd_loop.sh
while true; do
    pytest tests/ -x -v
    if [ $? -eq 0 ]; then
        echo "All tests pass! Add more tests or you're done."
        break
    fi
    echo "Test failed. Implement fix and press Enter..."
    read
done
```
