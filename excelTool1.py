import xlrd
import os
import json


class ManagerExcel:
  
    rootDir = "C:\\Users\\87793\Desktop\\config"

    rootJs = "C:\\Users\\87793\Desktop\\datajs"

    jsonPath = "C:\\Users\\87793\\Desktop\\datajson"

    def getListFile(self, root):
        return os.listdir(root)

    def parseOneExcel(self, file, jsPath, jsonPath):
        data = xlrd.open_workbook(file, encoding_override='utf-8')
        table = data.sheets()[0]
        rows = table.nrows
        cols = table.ncols
        tableDir1 = []
        listZD = table.row_values(1)
        listType = table.row_values(2)

        for i in range(rows-3):
            tableDir2 = {}
            for n in range(cols):
                if listType[n] == "numList":
                    if table.row_values(i + 3)[n] == '':
                        tableDir2.setdefault(listZD[n], [])
                    else:
                        listInt = []
                        if isinstance(table.row_values(i + 3)[n], float):
                            strData = str(table.row_values(i + 3)[n]).split(".")[0]
                            listInt = (strData.split("_"))
                        else:
                            listInt = (table.row_values(i + 3)[n].split("_"))
                        listResult = []
                        if len(listInt) > 0:
                            for m in range(len(listInt)):
                                listResult.append(int(listInt[m]))
                        tableDir2.setdefault(listZD[n], listResult)

                elif listType[n] == "num":
                    if table.row_values(i + 3)[n] == '':
                        tableDir2.setdefault(listZD[n], '')
                    else:
                        tableDir2.setdefault(listZD[n], int((str(table.row_values(i + 3)[n]).split("."))[0]))
                else:
                    tableDir2.setdefault(listZD[n], table.row_values(i+3)[n])
            tableDir1.append(tableDir2)

        if not os.path.exists(jsPath):
            file = open(jsPath[:-5] + '.js', 'w')
            file.write("let data = "+json.dumps(tableDir1) + '\n' + "module.exports = data;")
            file.close()

            if not os.path.exists(jsonPath):
                file1 = open(jsonPath[:-5] + '.json', 'w')
                file1.write(json.dumps(tableDir1))
                file1.close()
        return json.dumps(tableDir1)

    def parseAllFile(self, dirIn, dirOut, jsonP):
        listFile = self.getListFile(dirIn)
        for i in range(len(listFile)):
            if (dirIn + "/" + listFile[i]).find("~$") < 0:
                self.parseOneExcel(dirIn + "/" + listFile[i], dirOut + "/" + listFile[i], jsonP + "/" + listFile[i])


x = ManagerExcel()
x.parseAllFile(x.rootDir, x.rootJs, x.jsonPath)


