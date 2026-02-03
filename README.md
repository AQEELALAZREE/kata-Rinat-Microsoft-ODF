# Excel Formula Engine Implementation

## Project Overview

This project implements a Microsoft Excel formula engine that can read Excel spreadsheets and recalculate all formulas with results matching Excel's behavior.

## Project Goals

1. **Accuracy**: Match Excel's calculation behavior as closely as possible
2. **Completeness**: Implement as many Excel functions as feasible
3. **Performance**: Efficiently handle large spreadsheets with complex dependencies
4. **Maintainability**: Clean, well-documented code architecture

## Architecture Documentation

See the `/docs` directory for detailed architectural specifications:

- `ARCHITECTURE.md` - Overall system architecture
- `TECHNOLOGY_STACK.md` - Recommended technology choices
- `IMPLEMENTATION_GUIDE.md` - Step-by-step implementation instructions
- `FORMULA_SPECIFICATIONS.md` - Detailed formula implementation specs
- `TESTING_STRATEGY.md` - Testing and validation approach

## Project Structure

```
/
├── docs/                    # Architecture and design documentation
├── specs/                   # Detailed component specifications
├── test-cases/             # Excel test files and expected results
│   └── specs/              # JSON-based spec files (RFC 004)
├── examples/               # Example usage and integration patterns
└── review-checklists/      # Code review guidelines
```

## Getting Started

1. Read `docs/ARCHITECTURE.md` to understand the overall design
2. Review `docs/TECHNOLOGY_STACK.md` for language/library recommendations
3. Follow `docs/IMPLEMENTATION_GUIDE.md` for step-by-step development
4. Use `review-checklists/` for code review guidance

## Implementation Phases

1. **Core Parser** - Parse Excel formulas into AST
2. **Dependency Graph** - Build calculation dependency graph
3. **Basic Functions** - Implement fundamental formulas
4. **Advanced Functions** - Add complex formulas
5. **Optimization** - Performance tuning and caching

## Success Criteria

- Parse and evaluate Excel formulas correctly
- Handle circular references appropriately
- Match Excel's calculation results within acceptable precision
- Process real-world Excel files efficiently
