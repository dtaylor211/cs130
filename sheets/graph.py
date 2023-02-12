from typing import Dict, List, Set, Optional, TypeVar

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
        for v in [v for l in adjacency_list.values() 
                for v in l if v not in adjacency_list]:
            adjacency_list[v] = []
        self._adjacency_list = adjacency_list

    def get_adjacency_list(self) -> Dict[T, List[T]]:
        '''
        Returns the graph's adjacency list.

        '''
        return self._adjacency_list

    def transpose(self) -> None:
        '''
        Transposes the graph in place.

        '''
        transpose_adjacency_list = {}
        for k in self._adjacency_list:
            transpose_adjacency_list[k] = []

        for k, l in self._adjacency_list.items():
            for v in l:
                transpose_adjacency_list[v].append(k)
        self._adjacency_list = transpose_adjacency_list
    
    def get_strongly_connected_components(self) -> List[List[T]]:
        '''
        Calculate strongly connected components of the graph.
        Follows an iteritive version of Tarjan's algorithm.

        Returns:
        - List of strongly connected components (list of nodes in the component)
            in graph

        '''
        scc = []
        stack = []
        lowlink = {}
            
        # Currently recursive, in future assignments need to make iterative
        def strongconnect(k):
            idx = len(lowlink)
            lowlink[k] = idx
            stack_pos = len(stack)
            stack.append(k)
            for v in self._adjacency_list[k]:
                if v not in lowlink:
                    strongconnect(v)
                lowlink[k] = min(lowlink[k], lowlink[v])
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
        Calculates a topological sort of the graph.
        Only returns a valid topological sort for acyclic graphs.

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
                (k, enter) = stack.pop()
                if enter and k not in visited:
                    visited.add(k)
                    stack.append((k, False))
                    for v in self._adjacency_list[k]:
                        if v not in visited:
                            stack.append((v, True))
                else:
                    result.append(k)
        result.reverse()
        return result

    def get_reachable_nodes(self, initial: Set[T]) -> Set[T]:
        '''
        Gets reachable nodes of graph given a set of initial nodes

        Arguments:
        - initial: Set[T] - set of initial nodes

        Returns:
        - Set of reachable nodes in graph from initial nodes

        '''
        reachable = set(initial)
        for v in initial:
            stack = [v]
            while stack:
                u = stack.pop()
                if u in self._adjacency_list:
                    for w in self._adjacency_list[u]:
                        if w not in reachable:
                            stack.append(w)
                            reachable.add(w)
        return reachable

    def subgraph_from_nodes(self, nodes: Set[T]) -> None:
        '''
        Creates a subgraph of the graph in place with the given set of nodes

        Arguments:
        - nodes: Set[T] - list of nodes in subgraph

        '''
        sub_adjacency_list = {}
        for k in self._adjacency_list:
            if k in nodes:
                sub_adjacency_list[k] = [v for v in 
                                         self._adjacency_list[k] if v in nodes]
        self._adjacency_list = sub_adjacency_list