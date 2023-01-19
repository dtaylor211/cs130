from workbook import Workbook

wb = Workbook()
wb.new_sheet('Test')
print(wb.sheet_names)
print(wb.sheet_objects)
wb.set_cell_contents('Test','A1','12')
print(wb.get_cell_contents('Test','A1'))
wb.set_cell_contents('Test', 'A2', '13')
print(wb.get_cell_contents('Test', 'A2'))
wb.set_cell_contents('Test','A3','=1+1')
print(wb.get_cell_value('Test', 'A3'))
wb.set_cell_contents('Test','A4','=A1+A2')
print(wb.get_cell_value('Test', 'A4'))
wb.set_cell_contents('Test','A5',"'123")
print(wb.get_cell_value('Test', 'A5'))
print(wb.get_cell_contents('Test','A5'))
wb.set_cell_contents('Test','A6',"=A1+A5")
print(wb.get_cell_value('Test', 'A6'))
wb.set_cell_contents('Test', 'A7', "'true")
wb.set_cell_contents('Test', 'A8', '=A5&A7')
print(wb.get_cell_value('Test', 'A8'))
print(wb.get_cell_contents('Test','A8'))

# from lark import Lark
# import formula_evaluator

# evaluator = formula_evaluator.Evaluator()
# parser = Lark.open('formulas.lark', start='formula')
# tree = parser.parse("=A12")
# print(f'Tree Structure = \n{tree}\n')
# final = evaluator.transform(tree)
# print(f'Final Answer = \n{final}\n')