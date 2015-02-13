FileSelectFile, filename,,T:\2014 Login Book\Weekly Labscreen, Select File, Excel files (*.xlsm; *.xls)

X1 := ComObjGet(filename)

id3 := X1.Worksheets(1).Range("L4").Value
id4 := X1.Worksheets(1).Range("L5").Value
id5 := X1.Worksheets(1).Range("L6").Value
id6 := X1.Worksheets(1).Range("L7").Value
id7 := X1.Worksheets(1).Range("L8").Value
id8 := X1.Worksheets(1).Range("L9").Value
id9 := X1.Worksheets(1).Range("L10").Value
