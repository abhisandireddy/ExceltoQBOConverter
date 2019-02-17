Sub ExceltoQBOConvert()
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile("strPath")
   
    Dim dateData As String
    Dim amount As String
    Dim TranType As String
    Dim fitID As String
    For i = 4 To LastRow
    
    TranType = "<TRNTYPE>" & (Range("E" & i).Value)
    dateData = "<DTPOSTED>" & (Range("A" & i).Value) & "120000"
    amount = "<TRNAMT>" & (Range("B" & i).Value)
    fitID = "<FITID>" & (Range("F" & i).Value)
    Name = "<NAME>" & (Range("C" & i).Value)
    Memo = "<MEMO>" & (Range("D" & i).Value)
    oFile.writeLine ("<STMTTRN>")
    oFile.writeLine (TranType)
    oFile.writeLine (dateData)
    oFile.writeLine (amount)
    oFile.writeLine (fitID)
    oFile.writeLine (Name)
    oFile.writeLine (Memo)
    oFile.writeLine ("</STMTTRN>")
    Next i
    
    
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Sub


