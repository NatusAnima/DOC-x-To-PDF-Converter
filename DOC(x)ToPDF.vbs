Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False

inputFolder = "C:\Path\To\Your\Files"
outputFolder = "C:\Path\To\Save\PDFs"

Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set folder = fileSystem.GetFolder(inputFolder)

For Each file In folder.Files
    If LCase(fileSystem.GetExtensionName(file)) = "doc" Or LCase(fileSystem.GetExtensionName(file)) = "docx" Then
        Set doc = wordApp.Documents.Open(file.Path)
        pdfPath = outputFolder & "\" & fileSystem.GetBaseName(file) & ".pdf"
        doc.SaveAs2 pdfPath, 17
        doc.Close
    End If
Next

wordApp.Quit
