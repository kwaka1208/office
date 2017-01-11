Option Explicit

Const objWordExportAllDocument = 0 
Const objWordExportOptimizeForPrint = 0 
Const objWordExportDocumentContent = 0 
Const objWordExportFormatPDF = 17 
Const objWordExportCreateHeadingBookmarks = 1
'Const vbOKCancel = 1
'Const vbOK = 1

Dim objFs, objWord, filename, ext, path, pdfPath, outName, ret

Set objFs = CreateObject("Scripting.FileSystemObject")
Set objWord = CreateObject("Word.Application")

ret = MsgBox("ïœä∑ÇäJénÇµÇ‹Ç∑", vbOKCancel, "Word->PDFïœä∑")

If ret = vbOK Then
	ReDim args(-1)
	For Each filename in objFs.GetFolder(".").Files
		ext = objFs.GetExtensionName(filename)
		If (ext="doc" Or ext="docx") And Left(objFs.GetBaseName(filename),1) <> "~" Then
			Call ConvertPDF( fileName ) 
		End If
	Next
	MsgBox "ïœä∑ÇäÆóπÇµÇ‹ÇµÇΩ", vbInformation, "Word->PDFïœä∑"
End if
objWord.Quit(False) 
Set objWord = Nothing 
Set objFs = Nothing


' PDFïœä∑ä÷êî
Sub ConvertPDF( fileName ) 
    Dim objDoc 
    Dim pdf 
    Dim outName 
	path = objFs.GetAbsolutePathName(filename)
    Set objDoc = objWord.Documents.Open( path,,TRUE ) 
    outName = objFs.GetParentFolderName( path ) & "\" & _ 
                            objFs.GetBaseName( path ) & ".pdf"
    pdf = objWord.ActiveDocument.ExportAsFixedFormat ( _
                             outName, _
                             objWordExportFormatPDF, _ 
                             False, _ 
                             objWordExportOptimizeForPrint, _
                             objWordExportAllDocument,,, _
                             objWordExportDocumentContent, _
                             False, _ 
                             True, _
                             objWordExportCreateHeadingBookmarks) 
    objDoc.Close(False) 
    Set objDoc = Nothing
End Sub