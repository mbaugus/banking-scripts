Public Sub Begin()
'declare constants

DoIt
MsgBox ("0.00 amounts will NOT carry over to the payment file.  The row with zero amount due will be highlighted, the totals will match, but these lines with 0 amount will not count towards the total accounts count")

End Sub
Function DoIt()

Dim LastRowPayFile As Long
Dim LastRowOurListing As Long

Dim wksPayFile As Worksheet
Dim newWS As Worksheet
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wksPayFile = WB.Sheets(1)
Set newWS = Sheets.Add

newWS.NAME = "BatchFile"



Dim isempty As Boolean
Dim iter As Integer
Dim batchIter As Integer

'Keeps track of what line the created batch spreadsheet is on as we go down the list adding
batchIter = 1
'Keeps track of the payfile line we are on, I start it on row 2 because line one is a header
iter = 2
'Just a boolean value, when its triggered true the Loop ends
isempty = False

'These 6 stay constant for our batch payment file
Dim BatchDef As String
Dim AppCode As String
Dim CheckType As Integer
Dim Payee As String
Dim DueMonth As String
Dim GroupNumber As String

'These 3 change each loop
Dim accountNumber As String
Dim ToPay As Currency
Dim ParcelNumber As String

BatchDef = "ML252"
AppCode = "ML"
Payee = InputBox(Prompt:="Enter the payee number.", Title:="Enter the payee Number", Default:="104281")
CheckType = 2
DueMonth = InputBox(Prompt:="Enter the due month.", Title:="Enter due month", Default:="11")
GroupNumber = InputBox(Prompt:="Enter the group number for batch.", Title:="Group number", Default:="1")

    'Our loop
    Do While isempty = False
    'check first to make sure the value of cell isnt empty, if it is, there is a problem or the file ended
    If wksPayFile.Cells(iter, 1).Value <> "" Then
        
        ParcelNumber = wksPayFile.Cells(iter, 3).Value
        ToPay = wksPayFile.Cells(iter, 4).Value
        accountNumber = wksPayFile.Cells(iter, 5).Value
        
        If ToPay > 0 Then
        ZeroDue = False
        newWS.Cells(batchIter, 1).Value = BatchDef
        newWS.Cells(batchIter, 2).Value = AppCode
        newWS.Cells(batchIter, 3).Value = accountNumber
        newWS.Cells(batchIter, 4).Value = ToPay
        newWS.Cells(batchIter, 5).Value = CheckType
        newWS.Cells(batchIter, 6).Value = Payee
        newWS.Cells(batchIter, 7).Value = DueMonth
        newWS.Cells(batchIter, 8).Value = GroupNumber
        newWS.Cells(batchIter, 9).Value = ParcelNumber
        
        'Highlight the zero amount due line, so we can bump due dates up manually in system
        Else
        wksPayFile.Cells(iter, 1).EntireRow.Interior.ColorIndex = 6
        End If
        
    Else
        isempty = True
    End If
        iter = iter + 1
        ' only imcrement row if there wasn't 0 due.
        If ZeroDue = False Then
            batchIter = batchIter + 1
        End If
        ' Reset Zero Due Switch
        ZeroDue = True
    Loop
    
End Function
