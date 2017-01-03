''' 
  @author TheMiken
'''

#!/usr/bin/env
import xlrd


class fileReader(object):
    def __init__(self, pathFile, date):
        self.pathFile = pathFile
        self.date = date
        try:
            self.book = xlrd.open_workbook(self.pathFile)
            print("[CORRECTO] Archivo leido")
        except Exception as e:
            print ("[ERROR] Ha ocurrido un error al abrir el archivo -> " + e)

    def search(self, sheet, fileItem):
        print self.book.sheet_names()
        searchState = False
        for x in self.book.sheet_names():
            if sheet == x:
                bookSheet = self.book.sheet_by_name(sheet)
                for rowidx in range(bookSheet.nrows):
                    row = bookSheet.row(rowidx)
                    for colidx, cell in enumerate(row):
                        if cell.value == fileItem :
                            print "[CORRECTO] Palabra base encontrada"
                            posItem = (rowidx,colidx)
                            searchState = True
        if(searchState == True):
            print posItem
            print(bookSheet.cell(posItem[0],posItem[1]).value)
            print bookSheet.nrows
            finalItems = []
            for x in range(posItem[0], bookSheet.nrows):
                finalItems.append(bookSheet.cell(x, posItem[1]).value)
            return finalItems

        else:
            print("[ERROR] No se encontro la palabra base")






if __name__ == "__main__":
    test = fileReader("Base 21 de diciembre.xlsx", "02/01/2017")
    print(test.search("NO", "Correo"))
