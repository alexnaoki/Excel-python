from pint import UnitRegistry

u = UnitRegistry()

adutora_dict = {'Adutora(mm)': [100, 150, 200, 250, 300, 350, 400, 500]}

class CheckList:
    def __init__(self):
        pass

    def teste(self, a):
        return 0

    def min_diametro(self, diametro):
        '''
        diametro (m)
        '''
        diametro = diametro*u.m

        for d in adutora_dict['Adutora(mm)']:
            d = d*u.mm
            if d >= diametro:
                return d
            else:
                pass

if __name__ == '__main__':
    # print(CheckList.min_diametro(0.2))
    c = CheckList()
    print(c.min_diametro(0.2))
