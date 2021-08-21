from openpyxl import load_workbook, Workbook

class Booth(object):
    def __init__(self,boothName, previousBooth=None,companyName="",isPower=False, isPremium = False):
        self.boothName = boothName
        self.company_name = "" #could replace this with company data type
        self.previousBooth= previousBooth
        self.isPower = isPower
        self.isPremium = isPremium
        if(self.isPremium):
            self.isPower = True
    def makePremium(self):
        self.isPremium = True
        self.isPower = True
    def show(self):
        if(self.isPremium):
            print(str(self.boothName) + 'is a premium booth')
        else:
            print(str(self.boothName) + 'is not a premium booth')


wb = load_workbook(filename = 'Booth Layout Spring 2020.xlsx')

startrow = 3
endrow = 21
startcol = 3
endcol = 58
sheet_ranges = wb['Version 6']
y = sheet_ranges.cell(column=3, row=5)
print(type(y).__name__)
print(y.value)
for col in range(startcol, endcol+1):
    for row in range(startrow, endrow + 1):
        # print(sheet_ranges.cell(column=col, row=row).value)
        cell = sheet_ranges.cell(column=col, row=row)
        # print(type(cell).__name__)
        if(str(cell.value) != 'None'):
            booth = Booth(cell.value)
            if str(cell.value)>='A' and str(cell.value) <='Z':
                booth.isPremium = True
            booth.show()
            
            
            
            
print("end of for loop")


# sheet_ranges.unmerge_cells('C5:D6')





x = Booth("A1")
x2 = Booth("A2",x,'Siemens',False)
x2.show()
x2.previousBooth.show()

