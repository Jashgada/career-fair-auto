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
        print(str(self.boothName))



wb = load_workbook(filename = 'CRC floor diagram 8x16 v1.xlsx')

startrow = 26
endrow = 73
startcol = 4
endcol = 30
sheet_ranges = wb['Courts 3-6 only']
boothArr = []

for col in range(startcol, endcol+1):
    for row in range(startrow, endrow + 1):
        # print(sheet_ranges.cell(column=col, row=row).value)
        cell = sheet_ranges.cell(column=col, row=row)
        # print(type(cell).__name__)
        if(str(cell.value) != 'None'):
            booth = Booth(cell.value)
            if len(booth.boothName)>=3 and booth.boothName[2]=='P':
                booth.isPremium = True
            boothArr = boothArr + [booth]
boothHash= {}
for b in boothArr:
    if not (b.boothName[1:] in boothHash):
        if b.boothName[-1] == 'P':
            if b.boothName[1:3] in boothHash:
                # Add booth number to 


            
            



# sheet_ranges.unmerge_cells('C5:D6')





x = Booth("A1")
x2 = Booth("A2",x,'Siemens',False)
x2.show()
x2.previousBooth.show()

