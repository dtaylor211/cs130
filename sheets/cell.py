class CellType:
    LITERAL: int = 1
    STRING: int = 2
    FORMULA: int = 3
    EMPTY: int = 4


class Cell:
    '''
    This class represents an individual cell in a spreadsheet.

    '''
    
    def __init__(self):
        self.contents = None
        self.value = None
        self.type: int = CellType.EMPTY

    def set_value(self, input_str: str):
        inp = input_str.strip()
        if inp != "":
            self.contents = inp
            if inp[0] == "'":
                self.type = CellType.STRING
                if len(inp) == 1:
                    self.value = ""
                else:
                    self.value = inp[1:]
            elif inp[0] == "=":
                self.type = CellType.FORMULA
                self.value = Formula(inp) #todo
                pass # evaluate formula here, maybe formula class
        else: 
            self.contents = None
            self.value = None
            self.type = CellType.EMPTY


