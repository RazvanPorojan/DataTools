Set fso = CreateObject("Scripting.FileSystemObject")
Set inputFile = fso.OpenTextFile("iduri.csv")
Set jsonTemplate = fso.OpenTextFile("json_template.txt")
json = jsonTemplate.ReadAll()
i = 1000
Do While inputFile.AtEndOfStream <> True
    Set outputFile = fso.OpenTextFile("json" & i & ".json" , 2, True)
    i = i + 1
    arr = Split(inputFile.ReadLine, ",")
    id = arr(0)
    gcc = arr(1)
    ouput = Replace(json, "{id}" , id)
    ouput = Replace(ouput, "{gcc}" , gcc)
    outputFile.WriteLine ouput
Loop
Msgbox "Done"