
Sub ReplaceWordsForBook()
'
' ブック内の複数の文字列置換を一括で実行します。
'
	Dim sheet As Worksheet

	For Each sheet In ActiveWorkbook.Worksheets
		Call ReplaceWordsForSheet(sheet)
	Next

End Sub


Sub ReplaceWordsForSheet(ByRef sheet As Worksheet)
'
' シート内の複数の文字列置換を一括で実行します。
'
	Open "C:\tmp\list.csv" For Input As #1
	While Not EOF(1)
		Line Input #1, WordPair
		s = Split(WordPair, ",")
		' 取り出した文字列がペアでなかったら、その行は無視します。
		If (UBound(s) < 2) Then GoTo Continue

		sheet.Activate
		Cells.Select
		Selection.Replace What:=s(0), Replacement:=s(1), LookAt:=xlPart, _
		SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
		ReplaceFormat:=False
Continue:
	Wend
	Close #1
	Range("A1").Select
End Sub
