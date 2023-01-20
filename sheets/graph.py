from typing import Dict, List, Set, Optional, TypeVar

T = TypeVar('T')

class Graph:
    '''
    This class represents a generic graph used for cell dependencies in a 
    spreadsheet.
    Stores the adjacency list representing the graph
    '''

    def __init__(self, adjacency_list: Dict[T,List[T]], transpose = None):
        '''
        Initialize a new Graph

        Arguments:
        - adjacency_list: Dict[T,List[T]] - dictionary representing the directed
        edges of the graph
        - transpose = None - optional input of transpose graph
        so we calculate the transpose on creation without infinite recursion

        '''

        # Add keys that are in a dependency but not a key in the adjacency list
        for v in [v for l in adjacency_list.values() 
                for v in l if v not in adjacency_list]:
            adjacency_list[v] = []
        self.adjacency_list = adjacency_list

        if transpose is None:
            transpose_adjacency_list = {}
            for k in adjacency_list:
                transpose_adjacency_list[k] = []

            for k, l in adjacency_list.items():
                for v in l:
                    transpose_adjacency_list[v].append(k)
            self.transpose = Graph(transpose_adjacency_list, self)
        else:
            self.transpose = transpose

    
    def get_strongly_connected_components(self) -> List[List[T]]:
        '''
        Calculate strongly connected components of the graph.
        Follows an iteritive version of Tarjan's algorithm.

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
            for v in self.adjacency_list[k]:
                if v not in lowlink:
                    strongconnect(v)
                lowlink[k] = min(lowlink[k], lowlink[v])
            if lowlink[k] == idx:
                component = [stack[stack_pos:]]
                del stack[stack_pos:]
                scc.append(component)
                for item in component:
                    lowlink[item] = len(self.adjacency_list)

        for k in self.adjacency_list:
            if k not in lowlink:
                strongconnect(k)
        
        return scc

    def topological_sort(self) -> List[T]:
        '''
        Calculates a topological sort of the graph.
        Follows the iterative implementation given in class.
        Doesn't handle graphs with cycles.

        '''
        visited = set()
        result = []
        for k in self.adjacency_list:
            stack = [(k, True)]
            while stack:
                (k, enter) = stack.pop()
                if enter and k not in visited:
                    visited.add(k)
                    stack.append((k, False))
                    for v in self.adjacency_list[k]:
                        if v not in visited:
                            stack.append((v, True))
                else: # Leaving the node
                    result.append(k)

    def get_reachable_nodes(self, initial: Set[T]):
        '''
        Gets reachable nodes of graph given a list of initial nodes

        Arguments:
        - initial: List[T] -  list of initial nodes

        '''
        reachable = set(initial)
        for v in initial:
            stack = [v]
            while stack:
                u = stack.pop()
                for w in self.adjacency_list[u]:
                    if w not in reachable:
                        stack.append(w)
                        reachable.add(w)
        return reachable

    def subgraph_from_nodes(self, nodes: Set[T]):
        '''
        Creates a subgraph of the graph with the given set of nodes

        Arguments:
        - nodes: Set[T] -  list of nodes in subgraph

        '''
        sub_adjacency_list = {}
        for k in nodes:
            sub_adjacency_list[k] = [v for v in self.adjacency_list[k] if v in nodes]
        return Graph(sub_adjacency_list)