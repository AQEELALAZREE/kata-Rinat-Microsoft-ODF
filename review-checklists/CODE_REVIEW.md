# Parser Implementation Review Checklist

## Grammar & Syntax

- [ ] **Grammar completeness**: Does the grammar cover all Excel formula syntax?
  - [ ] Arithmetic operators: `+`, `-`, `*`, `/`, `^`
  - [ ] Comparison operators: `=`, `<>`, `<`, `>`, `<=`, `>=`
  - [ ] String concatenation: `&`
  - [ ] Parentheses for grouping
  - [ ] Cell references: `A1`, `$A$1`, `Sheet1!A1`
  - [ ] Range references: `A1:B10`
  - [ ] Function calls with arguments
  - [ ] String literals with escaped quotes
  - [ ] Boolean literals: `TRUE`, `FALSE`
  - [ ] Error literals: `#DIV/0!`, `#N/A`, etc.

- [ ] **Operator precedence**: Is precedence correct?
  - [ ] Parentheses (highest)
  - [ ] Range operator `:`
  - [ ] Percent `%`
  - [ ] Exponentiation `^`
  - [ ] Multiplication/Division `*`, `/`
  - [ ] Addition/Subtraction `+`, `-`
  - [ ] Concatenation `&`
  - [ ] Comparison operators (lowest)

- [ ] **Associativity**: Are operators left/right associative correctly?
  - [ ] Left: `+`, `-`, `*`, `/`, `&`
  - [ ] Right: `^`

## AST Construction

- [ ] **Node types**: Are all necessary AST node types defined?
  - [ ] NumberNode
  - [ ] StringNode
  - [ ] BooleanNode
  - [ ] ErrorNode
  - [ ] CellReferenceNode
  - [ ] RangeReferenceNode
  - [ ] BinaryOpNode
  - [ ] UnaryOpNode
  - [ ] FunctionCallNode

- [ ] **Node structure**: Do nodes contain all necessary information?
  - [ ] Cell references track absolute/relative flags
  - [ ] Sheet references preserved
  - [ ] Function names normalized (uppercase)

## Error Handling

- [ ] **Syntax errors**: Are syntax errors caught and reported clearly?
  - [ ] Missing closing parenthesis
  - [ ] Invalid cell reference
  - [ ] Unknown tokens
  - [ ] Empty formula

- [ ] **Error messages**: Are error messages helpful?
  - [ ] Include position information
  - [ ] Suggest corrections when possible
  - [ ] Don't crash on invalid input

## Testing

- [ ] **Unit tests**: Are there tests for:
  - [ ] Simple literals (numbers, strings, booleans)
  - [ ] Each operator type
  - [ ] Operator precedence
  - [ ] Cell references (all formats)
  - [ ] Range references
  - [ ] Function calls (0, 1, multiple arguments)
  - [ ] Nested functions
  - [ ] Complex formulas
  - [ ] Edge cases (empty, malformed)

- [ ] **Test coverage**: Is code coverage > 80%?

## Performance

- [ ] **Parsing speed**: Can parse typical formula in < 5ms?
- [ ] **Memory usage**: No memory leaks in repeated parsing?

## Code Quality

- [ ] **Documentation**: Are classes and methods documented?
- [ ] **Type hints**: Are type hints used (Python/TypeScript)?
- [ ] **Naming**: Are names clear and consistent?
- [ ] **No magic numbers**: Are constants named?

---

# Dependency Graph Review Checklist

## Graph Construction

- [ ] **Node creation**: Are all formula cells added as nodes?
- [ ] **Edge creation**: Are dependencies correctly identified?
  - [ ] Direct cell references
  - [ ] Range references (all cells in range)
  - [ ] Cross-sheet references
  - [ ] Named ranges

- [ ] **Dependency extraction**: Does it handle:
  - [ ] Simple references: `=A1`
  - [ ] Nested functions: `=SUM(IF(A1>0, B:B, C:C))`
  - [ ] Indirect references via functions

## Topological Sort

- [ ] **Algorithm correctness**: Does it produce valid calculation order?
- [ ] **Cycle detection**: Are circular references detected?
- [ ] **Cycle handling**: Are cycles handled appropriately?
  - [ ] Report error, or
  - [ ] Iterative calculation with convergence

## Incremental Updates

- [ ] **Affected cells**: Can identify cells affected by a change?
- [ ] **Minimal recalculation**: Only recalculates necessary cells?
- [ ] **Volatile functions**: Are volatile functions always recalculated?

## Testing

- [ ] **Test cases**:
  - [ ] Simple chain: A→B→C
  - [ ] Diamond: A→B,C→D
  - [ ] Multiple independent chains
  - [ ] Circular reference: A↔B
  - [ ] Self-reference: A→A
  - [ ] Complex cycles
  - [ ] Cross-sheet dependencies

- [ ] **Performance**: Can handle 10,000+ cell graph efficiently?

## Code Quality

- [ ] **Data structures**: Are appropriate data structures used?
  - [ ] Adjacency list for graph
  - [ ] Hash maps for O(1) lookups
- [ ] **Algorithms**: Are efficient algorithms used?
  - [ ] Kahn's algorithm or DFS for topological sort
  - [ ] Tarjan's for cycle detection

---

# Function Implementation Review Checklist

## Function Signature

- [ ] **Parameters**: Correct number and types?
- [ ] **Optional parameters**: Defaults match Excel?
- [ ] **Variadic arguments**: Handles variable argument count?

## Excel Compatibility

- [ ] **Behavior**: Matches Excel's behavior exactly?
  - [ ] Type coercion rules
  - [ ] Range handling
  - [ ] Empty cell handling
  - [ ] Error propagation

- [ ] **Edge cases**: Handles Excel's edge cases?
  - [ ] Division by zero
  - [ ] Empty ranges
  - [ ] Text in numeric functions
  - [ ] Booleans in calculations

## Type Handling

- [ ] **Type checking**: Validates argument types?
- [ ] **Type conversion**: Converts types per Excel rules?
  - [ ] Number to text
  - [ ] Text to number
  - [ ] Boolean to number (TRUE=1, FALSE=0)
  - [ ] Empty cell to 0 or ""

## Error Handling

- [ ] **Error propagation**: Propagates errors correctly?
- [ ] **Error generation**: Generates correct error types?
  - [ ] `#DIV/0!` for division by zero
  - [ ] `#VALUE!` for type errors
  - [ ] `#N/A` for lookup failures
  - [ ] `#REF!` for invalid references
  - [ ] `#NUM!` for numeric errors
  - [ ] `#NAME?` for unknown names

## Testing

- [ ] **Unit tests**: Tests for:
  - [ ] Basic functionality
  - [ ] Each parameter combination
  - [ ] Edge cases
  - [ ] Error cases
  - [ ] Type coercion

- [ ] **Excel comparison**: Results compared with actual Excel?

## Performance

- [ ] **Efficiency**: Is implementation efficient?
  - [ ] No unnecessary iterations
  - [ ] Appropriate algorithms
  - [ ] Early exit when possible

## Documentation

- [ ] **Docstring**: Includes:
  - [ ] Description
  - [ ] Parameters
  - [ ] Return value
  - [ ] Excel compatibility notes
  - [ ] Examples

---

# Calculation Engine Review Checklist

## Context Management

- [ ] **State tracking**: Maintains calculation state correctly?
  - [ ] Value cache
  - [ ] Current cell
  - [ ] Workbook reference

- [ ] **Cache management**: Cache invalidation works?

## Evaluation

- [ ] **AST evaluation**: Evaluates AST correctly?
- [ ] **Lazy evaluation**: Functions like IF evaluate lazily?
- [ ] **Recursion handling**: Handles deep nesting without stack overflow?

## Error Handling

- [ ] **Error propagation**: Errors propagate through formulas?
- [ ] **Error recovery**: Doesn't crash on errors?

## Testing

- [ ] **Integration tests**: Tests complete calculation flow?
- [ ] **Complex formulas**: Handles nested, complex formulas?
- [ ] **Large workbooks**: Performs well on large workbooks?

---

# Overall Code Review Checklist

## Architecture

- [ ] **Separation of concerns**: Components well-separated?
- [ ] **Modularity**: Code organized into logical modules?
- [ ] **Extensibility**: Easy to add new functions?

## Code Quality

- [ ] **Readability**: Code is clear and understandable?
- [ ] **Consistency**: Naming and style consistent?
- [ ] **DRY**: No significant code duplication?
- [ ] **Comments**: Complex logic explained?

## Testing

- [ ] **Coverage**: Overall coverage > 80%?
- [ ] **Test quality**: Tests are meaningful and thorough?
- [ ] **CI/CD**: Automated testing set up?

## Documentation

- [ ] **README**: Clear project overview?
- [ ] **API docs**: Public APIs documented?
- [ ] **Examples**: Usage examples provided?
- [ ] **Architecture docs**: Design documented?

## Performance

- [ ] **Benchmarks**: Performance benchmarks exist?
- [ ] **Profiling**: Performance bottlenecks identified?
- [ ] **Optimization**: Critical paths optimized?

## Security

- [ ] **Input validation**: Untrusted input validated?
- [ ] **Resource limits**: Protection against resource exhaustion?
- [ ] **Error messages**: Don't leak sensitive information?

---

# Pre-Merge Checklist

Before merging code:

- [ ] All tests pass
- [ ] Code coverage meets threshold
- [ ] No linting errors
- [ ] Documentation updated
- [ ] CHANGELOG updated (if applicable)
- [ ] Code reviewed by at least one other person
- [ ] Performance benchmarks run (if applicable)
- [ ] Excel compatibility tests pass
