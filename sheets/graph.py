'''
Graph

This module holds the functionality for finding cell dependencies using
adjacency lists.

See the Workbook, Sheet, and Cell modules for implementation.

Classes:
- Graph

    Methods:
    - get_adjacency_list(object) -> Dist[T, List[T]]
    - transpose(object) -> None
    - get_strongly_connected_components(object) -> List[List[T]]
    - topological_sort(object) -> List[T]
    - get_reachable_nodes(object, Set[T]) -> Set[T]
    - subgraph_from_nodes(object, Set[T]) -> None

'''


from typing import Dict, List, Set, TypeVar


T = TypeVar('T')

class Graph:
    '''
    This class represents a generic graph used for cell dependencies in a
    spreadsheet.
    Stores the adjacency list representing the graph.
    '''

    def __init__(self, adjacency_list: Dict[T, List[T]]):
        '''
        Initialize a new Graph

        Arguments:
        - adjacency_list: Dict[T, List[T]] - dictionary representing the
            directed edges of the graph

        '''

        # Add keys that are in a dependency but not a key in the adjacency list
        for val in [val for lst in adjacency_list.values()
                for val in lst if val not in adjacency_list]:
            adjacency_list[val] = []

        self._adjacency_list = adjacency_list

    def get_adjacency_list(self) -> Dict[T, List[T]]:
        '''
        Get the adjacency list of the graph

        Returns:
        - Dict with the graph's adjacency list

        '''

        return self._adjacency_list

    def transpose(self) -> None:
        '''
        Transpose the graph in place

        '''

        transpose_adjacency_list = {}
        for key in self._adjacency_list:
            transpose_adjacency_list[key] = []

        for key, lst in self._adjacency_list.items():
            for val in lst:
                transpose_adjacency_list[val].append(key)
        self._adjacency_list = transpose_adjacency_list

    def get_strongly_connected_components(self) -> List[List[T]]:
        '''
        Calculate strongly connected components of the graph, following
        an iteritive version of Tarjan's algorithm

        Returns:
        - List of strongly connected components (list of nodes in the component)
            in graph

        '''

        scc = []
        stack = []
        lowlink = {}
        idxs = {}

        # Iterative
        def strongconnect(k):
            dfs_stack = [(k, True)]
            while dfs_stack:
                k, enter = dfs_stack.pop()
                if enter and k not in lowlink:
                    idx = len(lowlink)
                    idxs[k] = (idx, len(stack))
                    lowlink[k] = idx
                    stack.append(k)
                    dfs_stack.append((k, False))
                    for val in self._adjacency_list[k]:
                        if val not in lowlink:
                            dfs_stack.append((val, True))
                else:
                    for val in self._adjacency_list[k]:
                        lowlink[k] = min(lowlink[k], lowlink[val])
                    idx, stack_pos = idxs[k]
                    if lowlink[k] == idx:
                        component = stack[stack_pos:]
                        del stack[stack_pos:]
                        scc.append(component)
                        for item in component:
                            lowlink[item] = len(self._adjacency_list)

        for k in self._adjacency_list:
            if k not in lowlink:
                strongconnect(k)

        return scc

    def topological_sort(self) -> List[T]:
        '''
        Calculate a topological sort of the graph, only returns
        a valid topological sort for acyclic graphs

        Returns:
        - List of nodes in graph sorted in topological order

        '''

        visited = set()
        result = []

        for k in self._adjacency_list:
            stack = []

            if k not in visited:
                stack.append((k, True))
            while stack:
                k, enter = stack.pop()
                if enter and k not in visited:
                    visited.add(k)
                    stack.append((k, False))
                    for val in self._adjacency_list[k]:
                        if val not in visited:
                            stack.append((val, True))
                else:
                    result.append(k)

        result.reverse()
        return result

    def get_reachable_nodes(self, initial: Set[T]) -> Set[T]:
        '''
        Get reachable nodes of graph given a set of initial nodes

        Arguments:
        - initial: Set[T] - set of initial nodes

        Returns:
        - Set of reachable nodes in graph from initial nodes

        '''

        reachable = set(initial)
        for val in initial:
            stack = [val]

            while stack:
                head = stack.pop()

                if head in self._adjacency_list:
                    for val2 in self._adjacency_list[head]:
                        if val2 not in reachable:
                            stack.append(val2)
                            reachable.add(val2)

        return reachable

    def subgraph_from_nodes(self, nodes: Set[T]) -> None:
        '''
        Create a subgraph of the graph in place with the given set of nodes

        Arguments:
        - nodes: Set[T] - list of nodes in subgraph

        '''

        sub_adjacency_list = {}
        for k in self._adjacency_list:
            if k in nodes:
                sub_adjacency_list[k] = [v for v in
                                         self._adjacency_list[k] if v in nodes]
        self._adjacency_list = sub_adjacency_list
