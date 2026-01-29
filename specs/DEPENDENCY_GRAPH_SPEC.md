# Dependency Graph Specification

## Overview

The Dependency Graph Builder analyzes formulas to determine calculation order and detect circular references. It constructs a directed graph where nodes are cells and edges represent dependencies.

## Core Concepts

### Dependency Relationship
- **Cell A depends on Cell B** means: Cell A's formula references Cell B
- **Edge direction**: A → B (A points to B, meaning A needs B's value)
- **Calculation order**: Reverse topological order (calculate B before A)

### Graph Structure

```python
class DependencyGraph:
    def __init__(self):
        self.nodes: dict[CellAddress, CellNode] = {}
        self.edges: dict[CellAddress, set[CellAddress]] = {}
        self.reverse_edges: dict[CellAddress, set[CellAddress]] = {}
        self.calculation_order: list[CellAddress] = []
        self.circular_refs: list[list[CellAddress]] = []
```

### Cell Node

```python
class CellNode:
    def __init__(self, address: CellAddress):
        self.address = address
        self.formula: str | None = None
        self.ast: ASTNode | None = None
        self.value: any = None
        self.is_volatile: bool = False
        self.dependencies: set[CellAddress] = set()
        self.dependents: set[CellAddress] = set()
        self.in_degree: int = 0
```

## Building the Graph

### Algorithm

```python
def build_dependency_graph(workbook: Workbook) -> DependencyGraph:
    """
    Build dependency graph from workbook
    
    Steps:
    1. Create nodes for all cells with formulas
    2. Parse each formula to extract dependencies
    3. Add edges for each dependency
    4. Calculate in-degrees
    5. Detect circular references
    6. Compute topological order
    """
    graph = DependencyGraph()
    
    # Step 1: Create nodes
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            if cell.has_formula():
                node = CellNode(cell.address)
                node.formula = cell.formula
                graph.add_node(node)
    
    # Step 2-3: Parse formulas and add edges
    for node in graph.nodes.values():
        ast = parse_formula(node.formula)
        dependencies = extract_dependencies(ast)
        
        for dep in dependencies:
            graph.add_edge(node.address, dep)
            node.dependencies.add(dep)
        
        # Check if formula contains volatile functions
        node.is_volatile = contains_volatile_function(ast)
    
    # Step 4: Calculate in-degrees
    for node in graph.nodes.values():
        node.in_degree = len(node.dependencies)
    
    # Step 5: Detect circular references
    graph.circular_refs = detect_cycles(graph)
    
    # Step 6: Compute calculation order
    if not graph.circular_refs:
        graph.calculation_order = topological_sort(graph)
    else:
        graph.calculation_order = topological_sort_with_cycles(graph)
    
    return graph
```

### Extracting Dependencies

```python
def extract_dependencies(ast: ASTNode) -> set[CellAddress]:
    """
    Recursively extract all cell references from AST
    """
    dependencies = set()
    
    def visit(node: ASTNode):
        if isinstance(node, CellReferenceNode):
            dependencies.add(node.get_address())
        elif isinstance(node, RangeReferenceNode):
            # Add all cells in range
            for cell in node.get_all_cells():
                dependencies.add(cell)
        elif isinstance(node, FunctionCallNode):
            for arg in node.arguments:
                visit(arg)
        elif isinstance(node, BinaryOpNode):
            visit(node.left)
            visit(node.right)
        elif isinstance(node, UnaryOpNode):
            visit(node.operand)
    
    visit(ast)
    return dependencies
```

## Topological Sort

### Kahn's Algorithm (Preferred)

```python
def topological_sort(graph: DependencyGraph) -> list[CellAddress]:
    """
    Kahn's algorithm for topological sorting
    
    Returns cells in calculation order (dependencies first)
    """
    # Copy in-degrees
    in_degree = {addr: node.in_degree for addr, node in graph.nodes.items()}
    
    # Queue of nodes with no dependencies
    queue = deque([addr for addr, deg in in_degree.items() if deg == 0])
    
    result = []
    
    while queue:
        # Process node with no remaining dependencies
        current = queue.popleft()
        result.append(current)
        
        # Reduce in-degree of dependents
        for dependent in graph.reverse_edges.get(current, set()):
            in_degree[dependent] -= 1
            if in_degree[dependent] == 0:
                queue.append(dependent)
    
    # If not all nodes processed, there's a cycle
    if len(result) != len(graph.nodes):
        raise CircularReferenceError("Circular reference detected")
    
    return result
```

### DFS-Based Alternative

```python
def topological_sort_dfs(graph: DependencyGraph) -> list[CellAddress]:
    """
    DFS-based topological sort
    """
    visited = set()
    temp_mark = set()
    result = []
    
    def visit(node_addr: CellAddress):
        if node_addr in temp_mark:
            raise CircularReferenceError(f"Cycle detected at {node_addr}")
        if node_addr in visited:
            return
        
        temp_mark.add(node_addr)
        
        # Visit dependencies first
        for dep in graph.edges.get(node_addr, set()):
            visit(dep)
        
        temp_mark.remove(node_addr)
        visited.add(node_addr)
        result.append(node_addr)
    
    for node_addr in graph.nodes:
        if node_addr not in visited:
            visit(node_addr)
    
    return result
```

## Circular Reference Detection

### Detecting Cycles

```python
def detect_cycles(graph: DependencyGraph) -> list[list[CellAddress]]:
    """
    Detect all cycles in the dependency graph using Tarjan's algorithm
    
    Returns list of strongly connected components (cycles)
    """
    index_counter = [0]
    stack = []
    lowlinks = {}
    index = {}
    on_stack = set()
    cycles = []
    
    def strongconnect(node_addr: CellAddress):
        index[node_addr] = index_counter[0]
        lowlinks[node_addr] = index_counter[0]
        index_counter[0] += 1
        stack.append(node_addr)
        on_stack.add(node_addr)
        
        # Consider successors (dependencies)
        for dep in graph.edges.get(node_addr, set()):
            if dep not in index:
                strongconnect(dep)
                lowlinks[node_addr] = min(lowlinks[node_addr], lowlinks[dep])
            elif dep in on_stack:
                lowlinks[node_addr] = min(lowlinks[node_addr], index[dep])
        
        # If node is a root node, pop the stack
        if lowlinks[node_addr] == index[node_addr]:
            component = []
            while True:
                w = stack.pop()
                on_stack.remove(w)
                component.append(w)
                if w == node_addr:
                    break
            
            # Only add if it's a real cycle (more than 1 node)
            if len(component) > 1:
                cycles.append(component)
    
    for node_addr in graph.nodes:
        if node_addr not in index:
            strongconnect(node_addr)
    
    return cycles
```

### Handling Circular References

Excel supports iterative calculation for circular references:

```python
def handle_circular_reference(graph: DependencyGraph, 
                              max_iterations: int = 100,
                              tolerance: float = 0.001) -> dict[CellAddress, any]:
    """
    Iteratively calculate cells in circular reference
    
    Excel's approach:
    1. Initialize all cells in cycle to 0
    2. Calculate cycle repeatedly until convergence
    3. Stop after max_iterations or when change < tolerance
    """
    cycle_cells = set()
    for cycle in graph.circular_refs:
        cycle_cells.update(cycle)
    
    # Initialize cycle cells
    values = {cell: 0 for cell in cycle_cells}
    
    for iteration in range(max_iterations):
        old_values = values.copy()
        
        # Calculate each cell in cycle
        for cell in cycle_cells:
            node = graph.nodes[cell]
            values[cell] = evaluate_formula(node.ast, values)
        
        # Check convergence
        max_change = max(abs(values[cell] - old_values[cell]) 
                        for cell in cycle_cells)
        
        if max_change < tolerance:
            break
    
    return values
```

## Incremental Updates

When a cell changes, only recalculate affected cells:

```python
def get_affected_cells(graph: DependencyGraph, 
                       changed_cell: CellAddress) -> set[CellAddress]:
    """
    Get all cells that depend on changed_cell (directly or indirectly)
    
    Uses BFS to traverse dependents
    """
    affected = set()
    queue = deque([changed_cell])
    
    while queue:
        current = queue.popleft()
        
        for dependent in graph.reverse_edges.get(current, set()):
            if dependent not in affected:
                affected.add(dependent)
                queue.append(dependent)
    
    return affected

def incremental_recalculate(graph: DependencyGraph, 
                            changed_cells: set[CellAddress]) -> list[CellAddress]:
    """
    Determine minimal set of cells to recalculate
    
    Returns cells in calculation order
    """
    # Get all affected cells
    affected = set()
    for cell in changed_cells:
        affected.update(get_affected_cells(graph, cell))
    
    # Add volatile cells (always recalculate)
    for addr, node in graph.nodes.items():
        if node.is_volatile:
            affected.add(addr)
            affected.update(get_affected_cells(graph, addr))
    
    # Filter calculation order to only affected cells
    return [cell for cell in graph.calculation_order if cell in affected]
```

## Volatile Functions

Functions that always recalculate:

```python
VOLATILE_FUNCTIONS = {
    'NOW', 'TODAY', 'RAND', 'RANDBETWEEN',
    'OFFSET', 'INDIRECT', 'INFO'
}

def contains_volatile_function(ast: ASTNode) -> bool:
    """Check if AST contains any volatile function"""
    def visit(node: ASTNode) -> bool:
        if isinstance(node, FunctionCallNode):
            if node.function_name in VOLATILE_FUNCTIONS:
                return True
            return any(visit(arg) for arg in node.arguments)
        elif isinstance(node, BinaryOpNode):
            return visit(node.left) or visit(node.right)
        elif isinstance(node, UnaryOpNode):
            return visit(node.operand)
        return False
    
    return visit(ast)
```

## Example Scenarios

### Example 1: Simple Chain

```
A1: =10
B1: =A1 * 2
C1: =B1 + 5
```

**Graph**:
```
C1 → B1 → A1
```

**Calculation Order**: `[A1, B1, C1]`

### Example 2: Diamond Dependency

```
A1: =10
B1: =A1 * 2
C1: =A1 + 5
D1: =B1 + C1
```

**Graph**:
```
    D1
   /  \
  B1  C1
   \  /
    A1
```

**Calculation Order**: `[A1, B1, C1, D1]` or `[A1, C1, B1, D1]`

### Example 3: Circular Reference

```
A1: =B1 + 1
B1: =A1 + 1
```

**Graph**:
```
A1 ⇄ B1
```

**Handling**:
1. Detect cycle: `[A1, B1]`
2. Initialize: A1=0, B1=0
3. Iterate:
   - Iteration 1: A1=1, B1=1
   - Iteration 2: A1=2, B1=2
   - ... (diverges, needs max iterations)

## Performance Optimization

### Caching

```python
class DependencyGraph:
    def __init__(self):
        # ... existing fields ...
        self._calculation_order_cache: list[CellAddress] | None = None
        self._affected_cache: dict[CellAddress, set[CellAddress]] = {}
    
    def invalidate_cache(self):
        """Invalidate caches when graph changes"""
        self._calculation_order_cache = None
        self._affected_cache.clear()
```

### Parallel Calculation

Cells at the same dependency level can be calculated in parallel:

```python
def get_calculation_levels(graph: DependencyGraph) -> list[set[CellAddress]]:
    """
    Group cells by dependency level for parallel calculation
    
    Level 0: No dependencies
    Level 1: Depends only on Level 0
    Level 2: Depends on Level 0 or 1
    etc.
    """
    levels = []
    remaining = set(graph.nodes.keys())
    processed = set()
    
    while remaining:
        # Find cells with all dependencies processed
        current_level = set()
        for cell in remaining:
            deps = graph.edges.get(cell, set())
            if deps.issubset(processed):
                current_level.add(cell)
        
        if not current_level:
            # Circular reference
            break
        
        levels.append(current_level)
        processed.update(current_level)
        remaining -= current_level
    
    return levels
```

## Testing

### Test Cases

```python
def test_simple_dependency():
    # A1=10, B1=A1*2
    graph = build_test_graph({
        'A1': None,
        'B1': '=A1*2'
    })
    assert graph.calculation_order == ['A1', 'B1']

def test_diamond_dependency():
    graph = build_test_graph({
        'A1': None,
        'B1': '=A1*2',
        'C1': '=A1+5',
        'D1': '=B1+C1'
    })
    order = graph.calculation_order
    assert order.index('A1') < order.index('B1')
    assert order.index('A1') < order.index('C1')
    assert order.index('B1') < order.index('D1')
    assert order.index('C1') < order.index('D1')

def test_circular_reference_detection():
    graph = build_test_graph({
        'A1': '=B1+1',
        'B1': '=A1+1'
    })
    assert len(graph.circular_refs) == 1
    assert set(graph.circular_refs[0]) == {'A1', 'B1'}

def test_incremental_update():
    graph = build_test_graph({
        'A1': None,
        'B1': '=A1*2',
        'C1': '=B1+5',
        'D1': '=10'
    })
    affected = get_affected_cells(graph, 'A1')
    assert affected == {'B1', 'C1'}
```

## References

- [Topological Sorting Algorithms](https://en.wikipedia.org/wiki/Topological_sorting)
- [Tarjan's Strongly Connected Components](https://en.wikipedia.org/wiki/Tarjan%27s_strongly_connected_components_algorithm)
- [Excel Circular Reference Handling](https://support.microsoft.com/en-us/office/remove-or-allow-a-circular-reference-8540bd0f-6e97-4483-bcf7-1b49cd50d123)
