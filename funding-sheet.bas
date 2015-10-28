Public firstNameRow As Integer
Public lastNameRow As Integer
Public escrowTaxRow As Integer
Public escrowInsRow As Integer
Public countyRow As Integer

Public accountColumn As Integer
Public nameColumn As Integer
Public countyColumn As Integer
Public taxColumn As Integer
Public insColumn As Integer

Public accountRow As Integer
Public WS As Worksheet


Public Sub FundingSheet()

Dim iter As Integer
Dim count As Integer
Dim firstName As String
Dim lastName As String
Dim fullName As String
Dim isItEmpty As Boolean

'For writing to the new spreadsheet'
accountColumn = 1
nameColumn = 2
countyColumn = 3
taxColumn = 4
insColumn = 5

'Declare the parsing variables'
isItEmpty = False
escrowTaxRow = 20
escrowInsRow = 21
iter = 4
firstNameRow = 3
lastNameRow = 2
accountRow = 6
MsgBox "Begin."
countyRow = 7
Dim escrowAmount As Integer

escrowAmount = CountEscrow("cinci")
escrowAmount = escrowAmount + CountEscrow("dayton")
escrowAmount = escrowAmount + CountEscrow("columbus")
escrowAmount = escrowAmount + CountEscrow("indianapolis")

MsgBox "New Escrow loans detected " & escrowAmount

'Add new page onto the worksheet to save our work'

Set WS = Sheets.Add

'Avoid internal global naming problems with events'
For i = 1 To 10000
    DoEvents
Next i

WS.Name = "Hello"

For i = 1 To 10000
    DoEvents
Next i

Dim currentRow As Integer
currentRow = 1

'Add accountnumbers'
currentRow = WriteAccountNumbers("cinci", currentRow)
currentRow = WriteAccountNumbers("dayton", currentRow)
currentRow = WriteAccountNumbers("columbus", currentRow)
currentRow = WriteAccountNumbers("indianapolis", currentRow)

Worksheets("Hello").Columns("A:E").Select
Selection.EntireColumn.AutoFit
'Worksheets("Hello").Columns("A:E").Select
Selection.EntireColumn.HorizontalAlignment = xlCenter
Selection.EntireColumn.Font.Size = 11
Selection.EntireColumn.Font.Name = "Calibri"

Worksheets("Hello").Columns("A").ColumnWidth = 11.43
Worksheets("Hello").Columns("B").ColumnWidth = 24.86
Worksheets("Hello").Columns("C").ColumnWidth = 17.86
Worksheets("Hello").Columns("D").ColumnWidth = 8.43
Worksheets("Hello").Columns("E").ColumnWidth = 8

currentRow = 1
currentRow = WriteNames("cinci", currentRow)
currentRow = WriteNames("dayton", currentRow)
currentRow = WriteNames("columbus", currentRow)
currentRow = WriteNames("indianapolis", currentRow)

currentRow = 1
currentRow = WriteCounty("cinci", currentRow)
currentRow = WriteCounty("dayton", currentRow)
currentRow = WriteCounty("columbus", currentRow)
currentRow = WriteCounty("indianapolis", currentRow)

currentRow = 1
currentRow = WriteCollected("cinci", currentRow)
currentRow = WriteCollected("dayton", currentRow)
currentRow = WriteCollected("columbus", currentRow)
currentRow = WriteCollected("indianapolis", currentRow)
End Sub

Function CountEscrow(x As String) As Integer

Dim counter As Integer
Dim isempty As Boolean
Dim iter As Integer

iter = 4
counter = 0
isempty = False


Do While isempty = False
    If Worksheets(x).Cells(firstNameRow, iter).Value <> "" Then
    
        If Worksheets(x).Cells(escrowTaxRow, iter).Value <> 0 Or _
        Worksheets(x).Cells(escrowInsRow, iter).Value <> 0 Then
        counter = counter + 1
        End If
        
        iter = iter + 1
    Else
        isempty = True
    End If
    
Loop

CountEscrow = counter


End Function

'Takes worksheetname and current rowNumber, returns the NEXT rownumber
Function WriteAccountNumbers(x As String, i As Integer) As Integer

Dim isempty As Boolean
Dim iter As Integer

iter = 4
isempty = False

    Do While isempty = False
    If Worksheets(x).Cells(firstNameRow, iter).Value <> "" Then
    
        If Worksheets(x).Cells(escrowTaxRow, iter).Value <> 0 Or _
        Worksheets(x).Cells(escrowInsRow, iter).Value <> 0 Then
           
        Worksheets("Hello").Cells(i, accountColumn).Value = _
        Worksheets(x).Cells(accountRow, iter).Value

        i = i + 1
        End If
        
        iter = iter + 1
    Else
        isempty = True
    End If
    
    Loop
    
    WriteAccountNumbers = i
End Function

Function WriteNames(x As String, i As Integer) As Integer

Dim iter As Integer
Dim isempty As Boolean
Dim firstName As String
Dim lastName As String

iter = 4
isempty = False

    Do While isempty = False
    If Worksheets(x).Cells(firstNameRow, iter).Value <> "" Then
    
        If Worksheets(x).Cells(escrowTaxRow, iter).Value <> 0 Or _
        Worksheets(x).Cells(escrowInsRow, iter).Value <> 0 Then
        
        firstName = Worksheets(x).Cells(firstNameRow, iter).Value
        lastName = Worksheets(x).Cells(lastNameRow, iter).Value
        
        Worksheets("Hello").Cells(i, nameColumn).Value = _
        UCase(lastName & " " & firstName)
        i = i + 1
        End If
        
        iter = iter + 1
    Else
        isempty = True
    End If
    
    Loop
    
    WriteNames = i

End Function

Function WriteCounty(x As String, i As Integer) As Integer

Dim iter As Integer
Dim isempty As Boolean
Dim countyString As String

iter = 4
isempty = False

    Do While isempty = False
    If Worksheets(x).Cells(firstNameRow, iter).Value <> "" Then
    
        If Worksheets(x).Cells(escrowTaxRow, iter).Value <> 0 Or _
        Worksheets(x).Cells(escrowInsRow, iter).Value <> 0 Then
        
        countyString = Worksheets(x).Cells(countyRow, iter).Value
       
        Worksheets("Hello").Cells(i, countyColumn).Value = _
        UCase(countyString)
        i = i + 1
        End If
        
        iter = iter + 1
    Else
        isempty = True
    End If
    
    Loop
    
    WriteCounty = i

End Function

Function WriteCollected(x As String, i As Integer) As Integer

Dim iter As Integer
Dim isempty As Boolean
Dim TAX As String
Dim INS As String

iter = 4
isempty = False


    Do While isempty = False
    If Worksheets(x).Cells(firstNameRow, iter).Value <> "" Then
    
        If Worksheets(x).Cells(escrowTaxRow, iter).Value <> 0 Or _
        Worksheets(x).Cells(escrowInsRow, iter).Value <> 0 Then
        
        TAX = Worksheets(x).Cells(escrowTaxRow, iter).Value
        INS = Worksheets(x).Cells(escrowInsRow, iter).Value
        
        Worksheets("Hello").Cells(i, taxColumn).Value = TAX
        Worksheets("Hello").Cells(i, insColumn).Value = INS
        i = i + 1
        End If
        
        iter = iter + 1
    Else
        isempty = True
    End If
    
    Loop
    
    WriteCollected = i

End Function
