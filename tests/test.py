# need to $ pip install pytest
import pytest
import context
from sheets.workbook import Workbook
from sheets.graph import Graph
from decimal import Decimal


wb = Workbook()
wb.new_sheet('Test')
wb.set_cell_contents('Test','A1','12')
wb.set_cell_contents('Test', 'A2', '13')
wb.set_cell_contents('Test','A4','=A1+A2')
wb.set_cell_contents('Test','A5','=1')
wb.set_cell_contents('Test','A6', '=A1+A4+A2+A5')
print(wb.get_cell_value('Test','A6'))
wb.set_cell_contents('Test','A1','=1')
print(wb.get_cell_value('Test','A6'))
wb.set_cell_contents('Test', 'A5', '=A6')
print(wb.get_cell_value('Test','A1'))
print(wb.get_cell_value('Test','A2'))
print(wb.get_cell_value('Test','A4'))
print(wb.get_cell_value('Test','A5'))
print(wb.get_cell_value('Test','A6'))
wb.new_sheet('Test1')
wb.set_cell_contents('Test1','A4','=Test!A1')
print(wb.get_cell_value('Test1','A4'))
wb.del_sheet('Test')
print(wb.get_cell_value('Test1','A4'))

# adj = {
#     1: [2, 3],
#     2: [1, 4],
#     3: [1, 4],
#     4: [2, 3, 5],
#     5: [4],
#     6: [7],
#     7: [6, 8],
#     8: [7]
# }
# g = Graph(adj)
# print(g.adjacency_list)
# print(g.transpose.adjacency_list)
# print(g.get_reachable_nodes([1]))
# print(g.subgraph_from_nodes(g.get_reachable_nodes([1])).adjacency_list)
# print(g.get_reachable_nodes([6]))
# print(g.subgraph_from_nodes(g.get_reachable_nodes([6])).adjacency_list)
# print(g.get_reachable_nodes([1,6]))
# print(g.subgraph_from_nodes(g.get_reachable_nodes([1,6])).adjacency_list)

# from lark import Lark
# import formula_evaluator

# evaluator = formula_evaluator.Evaluator()
# parser = Lark.open('formulas.lark', start='formula')
# tree = parser.parse("=A12")
# print(f'Tree Structure = \n{tree}\n')
# final = evaluator.transform(tree)
# print(f'Final Answer = \n{final}\n')