Public Sub Begin()
'declare constants

MsgBox ("All corrections MUST BE DONE before this payment file is created.  Invalid accounts, duplicates, incorrect account numbers and amounts to pay.  These all carry over and are final.  Ye be warned.")
MsgBox ("0.00 amounts will NOT carry over to the payment file.  The row with zero amount due will be highlighted, the totals will match, but these lines with 0 amount will not count towards the total accounts count")

DoIt
    
MsgBox ("Finished.  Save BatchFile as a 'Tab delimited file' to use with Teller.")
End Sub

Function DoIt()

Dim LastRowPayFile As Long
Dim LastRowOurListing As Long

Dim wksPayFile As Worksheet
Dim newWS As Worksheet
Dim WB As Workbook

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

'Column inputs for account, amount to pay and parcel
Dim accountColumn As String
Dim toPayColumn As String
Dim parcelNumberColumn As String
Dim accountColumnInt As Integer
Dim toPayColumnInt As Integer
Dim parcelNumberColumnInt As Integer

'These 3 change each loop
Dim accountNumber As String
Dim ToPay As Currency
Dim ParcelNumber As String

CheckType = 2
BatchDef = "ML252"
AppCode = "ML"
Payee = InputBox(Prompt:="Enter the payee number.", Title:="Enter the payee Number", Default:="104281")
DueMonth = InputBox(Prompt:="Enter the due month for this payee.", Title:="Enter due month", Default:="11")
GroupNumber = InputBox(Prompt:="Group number for batch. Usually 1", Title:="Group number", Default:="1")
accountColumn = InputBox(Prompt:="The column your account numbers are located.  Use numbers, not letters.", Default:="1")
toPayColumn = InputBox(Prompt:="The column your amounts to pay are located.  Use numbers, not letters.", Default:="2")
parcelNumberColumn = InputBox(Prompt:="The column your parcels are located.  Use numbers, not letters.", Default:="3")

'Convert last 3 inputs into numbers, since we actually use these as number values in the script, not just outputting into text.
accountColumnInt = CInt(accountColumn)
toPayColumnInt = CInt(toPayColumn)
parcelNumberColumnInt = CInt(parcelNumberColumn)

'create the new worksheet.
Set WB = ActiveWorkbook
Set wksPayFile = WB.Sheets(1)
Set newWS = Sheets.Add
newWS.Name = "BatchFile"


    'Our loop
    Do While isempty = False
    'check first to make sure the value of cell isnt empty, if it is, there is a problem or the file ended
    If wksPayFile.Cells(iter, 1).Value <> "" Then
        'Pull the 3 things we care about from the payment file
        ParcelNumber = wksPayFile.Cells(iter, parcelNumberColumnInt).Value
        ToPay = wksPayFile.Cells(iter, toPayColumnInt).Value
        accountNumber = wksPayFile.Cells(iter, accountColumnInt).Value
        
        'Update the batch file, if not zero due.  The Else highlights the zero due line for fixin laters
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
    
    'change Column 4 to on batch file to "NUMBER" not "CURRENCY", will mess it all up kiddos.
    newWS.Columns(4).NumberFormat = "General"
                
End Function
