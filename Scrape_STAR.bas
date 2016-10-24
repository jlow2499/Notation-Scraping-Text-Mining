Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub Scrape_Star()
Dim CurrentHost As Object
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession
Dim irow As Long
Dim i As Long
Dim x As Variant
irow = 5
i = 6

  Do
  If Range("C" & irow).Value = "" Then
    Application.StatusBar = "Credit AR Add Complete"
    MsgBox "Add Complete"
    Exit Sub
    End If
    
If Range("B" & irow).Value = "DONE" Then
    Sleep 1
Else
''''''''''''''''''''''''''''

        If CurrentHost.GetText(0, 22, 52) = "ENTER SELECTION (.,FILE#,/,STATUS,-nnnnn,Tn,/R,HELP)" Then
        CurrentHost.Output Range("C" & irow).Value & ChrW$(13)
        Else
        Sleep 500
        CurrentHost.Output Range("C" & irow).Value & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 51) = "ENTER SELECTION, FILE#,HELP,W,V,LH,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "12" & ChrW$(13)
        Else
        Sleep 500
        CurrentHost.Output "12" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (<CR>,nnn,-,/)" Then
        CurrentHost.Output "4" & ChrW$(13)
        Else
        Sleep 500
        CurrentHost.Output "4" & ChrW$(13)
        End If
        
        Sleep 500
        
If Range("D" & irow).Value = 1 Then

Cells(irow, i) = CurrentHost.GetText(0, 15, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 16, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 17, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 18, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 19, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 20, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 21, 100)
        
        
        Sleep 500
        
Else

        
x = 1
Do Until x = Range("E" & irow).Value

        Cells(irow, i) = CurrentHost.GetText(0, 15, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 16, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 17, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 18, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 19, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 20, 100)
            i = i + 1
        Cells(irow, i) = CurrentHost.GetText(0, 21, 100)
        
        
        Sleep 500
                
        CurrentHost.Output ChrW$(13)
        
        x = x + 1
        
        Sleep 200
           
Loop

End If
       
        Sleep 200
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (<CR>,nnn,-,/)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 500
        CurrentHost.Output "/" & ChrW$(13)
        End If
                
        If CurrentHost.GetText(0, 22, 11) = "ENTER (n,/)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 500
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 51) = "ENTER SELECTION, FILE#,HELP,W,V,LH,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 500
        CurrentHost.Output "/" & ChrW$(13)
        End If
       
 '''''''''''''''''''''''''''
   
   Range("A" & irow).Value = "DONE"
   Sleep 200
   
End If
    x = 1
    i = 6
    irow = irow + 1
    Loop

End Sub

 Public Sub Analyze()
    Dim ws As Worksheet
    Dim iWarnColor As Integer
    Dim rng As Range, aCell As Range, bCell As Range
    Dim LR As Long
    Dim Str As String
    
  Set ws = ThisWorkbook.Sheets("Sheet1")
  
  Str = InputBox("What string are you looking for?")
   
 iWarnColor = xlThemeColorAccent2

    With ws
        LR = .Range("F" & .Rows.Count).End(xlUp).Row
        
        Set rng = .Range("F1:AHH" & LR)

        rng.Interior.ColorIndex = xlNone

        Set aCell = rng.Find(What:=Str, LookIn:=xlValues, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
                        

        If Not aCell Is Nothing Then
            Set bCell = aCell
            aCell.Interior.ColorIndex = iWarnColor
            Do
                Set aCell = rng.FindNext(After:=aCell)

                If Not aCell Is Nothing Then
                    If aCell.Address = bCell.Address Then Exit Do
                    aCell.Interior.ColorIndex = iWarnColor
                Else
                    Exit Do
                End If
            Loop
        End If
    End With

    
End Sub


Function GetCellColor(xlRange As Range)
    Dim indRow, indColumn As Long
    Dim arResults()
 
    Application.Volatile
 
    If xlRange Is Nothing Then
        Set xlRange = Application.ThisCell
    End If
 
    If xlRange.Count > 1 Then
      ReDim arResults(1 To xlRange.Rows.Count, 1 To xlRange.Columns.Count)
       For indRow = 1 To xlRange.Rows.Count
         For indColumn = 1 To xlRange.Columns.Count
           arResults(indRow, indColumn) = xlRange(indRow, indColumn).Interior.Color
         Next
       Next
     GetCellColor = arResults
    Else
     GetCellColor = xlRange.Interior.Color
    End If
End Function
 
Function GetCellFontColor(xlRange As Range)
    Dim indRow, indColumn As Long
    Dim arResults()
 
    Application.Volatile
 
    If xlRange Is Nothing Then
        Set xlRange = Application.ThisCell
    End If
 
    If xlRange.Count > 1 Then
      ReDim arResults(1 To xlRange.Rows.Count, 1 To xlRange.Columns.Count)
       For indRow = 1 To xlRange.Rows.Count
         For indColumn = 1 To xlRange.Columns.Count
           arResults(indRow, indColumn) = xlRange(indRow, indColumn).Font.Color
         Next
       Next
     GetCellFontColor = arResults
    Else
     GetCellFontColor = xlRange.Font.Color
    End If
 
End Function
 
Function CountCellsByColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Interior.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByColor = cntRes
End Function
 
Function SumCellsByColor(rData As Range, cellRefColor As Range)
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim sumRes
 
    Application.Volatile
    sumRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Interior.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Interior.Color Then
            sumRes = WorksheetFunction.Sum(cellCurrent, sumRes)
        End If
    Next cellCurrent
 
    SumCellsByColor = sumRes
End Function
 
Function CountCellsByFontColor(rData As Range, cellRefColor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cntRes As Long
 
    Application.Volatile
    cntRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Font.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Font.Color Then
            cntRes = cntRes + 1
        End If
    Next cellCurrent
 
    CountCellsByFontColor = cntRes
End Function
 
Function SumCellsByFontColor(rData As Range, cellRefColor As Range)
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim sumRes
 
    Application.Volatile
    sumRes = 0
    indRefColor = cellRefColor.Cells(1, 1).Font.Color
    For Each cellCurrent In rData
        If indRefColor = cellCurrent.Font.Color Then
            sumRes = WorksheetFunction.Sum(cellCurrent, sumRes)
        End If
    Next cellCurrent
 
    SumCellsByFontColor = sumRes
End Function

