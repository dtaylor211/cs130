class CellType:
    pass


class Cell:
    '''
    This class represents an individual cell in a spreadsheet.

    '''

    def __init__(self):
        self.contents = None
        self.value = None
        self.type

    def set_value(self, input_str: str):
        inp = input_str.strip()
        if inp != "":
            self.contents = inp
            if inp[0] == "'":
                # want to set to StringType?
                if len(inp) == 1:
                    self.value = ""
                else:
                    self.value = inp[1:]
            elif inp[0] == "=":
                # todo
               
            # evaluate contents to set value 
        else: 
            self.contents = None
            self.value = None


