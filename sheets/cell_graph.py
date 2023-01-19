from typing import List
from .cell import Cell

class CellGraph:
    '''
    This class represents graph of cell dependencies in a spreadsheet.
    Stores the adjacency list representing the graph
    '''

    def __init__(self, adjacency_list: Dict[Cell,List[Cell]]):
        '''
        Initialize a new CellGraph

        Arguments:
        - adjacency_list: Dict[Cell,List[Cell]] - dictionary with cell
        dependencies representing the directed edges of the graph

        '''

        # Add cells that are in a dependency but not a key in the adjacency list
        for v in [v for l in adjacency_list.values() 
                for v in l if v not in adjacency_list]:
            adjacency_list[v] = []

        self.adjacency_list = adjacency_list
    
    def get_strongly_connected_componenets(self) -> List[List[Cell]]:
        '''
        Calculate strongly connected componenets of the cell graph.
        Follows an iteritive version of Tarjan's algorithm.

        '''
        scc = []
        stack = []
        lowlink = {}
            
        # Currently recursive, in future assignments need to make iterative
        def strongconnect(cell):
            idx = len(lowlink)
            lowlink[cell] = idx
            stack_pos = len(stack)
            stack.append(cell)
            for child in self.adjacency_list[cell]:
                if child not in lowlink:
                    strongconnect(child)
                lowlink[cell] = min(lowlink[cell], lowlink[child])
            if lowlink[cell] == idx:
                component = [stack[stack_pos:]]
                del stack[stack_pos:]
                scc.append(component)
                for item in component:
                    lowlink[item] = len(self.adjacency_list)

        for cell in self.adjacency_list.keys():
            if cell not in lowlink:
                strongconnect(cell)
        
        return scc

    def topological_sort(self) -> List[Cell]:
        '''
        Calculates a topological sort of the cell graph.
        Follows the iterative implementation given in class.
        Doesn't handle graphs with cycles.

        '''
        visited = set()
        result = []
        for cell in self.adjacency_list.keys():
            stack = [(cell, True)]
            while stack:
                (cell, enter) = stack.pop()
                if enter and cell not in visited:
                    visited.add(cell)
                    stack.append((cell, False))
                    for n_cell in self.adjacency_list[cell]:
                        if n_cell not in visited:
                            stack.append((n_cell, True))
                else: # Leaving the node
                    result.append(cell)