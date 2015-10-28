' We will need a recent Hendricks Listing.  Our first procedure
' will be similar to what we do with the request file.  Do manual file
' maintenance first to the accounts, make the adjustments to both file and the
' accounts.  Anything that is showing a due date of paid, needs investigated and
' deleted if it indeed is paid, else correct the issue.  All incorrect parcel
' formats must be fixed.  if a parcel has 0 due, you may delete it from the
' request file all together, however you must bump the due date up in the system
'
' Scan Hendricks Pay file versus our now cleaned and ready to go database listing
' All deletes need highlighted for investigation, and then subsequent deletion
' Do the opposite and make sure to append any additional parcels to the end of
' The spreadsheet.  We must manually look up amounts to pay on these.
' ** After deletions and adds and all maintenance is complete, we can run the 3rd
' ** module, which will create a tab delimited .txt file for.  This file can be
' ** uploaded to Teller for a large batch payment run **

' STEP 1. MANUAL MAINTENANCE
' STEP 2. FIND ADD/DELETE
' STEP 3. MAKE CORRECTIONS/DELETES AND ADD AMOUNTS/NAME/ADDRESS
' STEP 4. EXPORT OUR BATCH FILE
' STEP 5. RUN BATCH THROUGH TELLER, BALANCE.  GENERATE CHECK.
' STEP 6. ALL FAILURES TO PAY IN THE BATCH NEED TO BE MAINTAINED ON THE PAYMENT FILE
' STEP 7. EMAIL COUNTY OUR FINAL LISTING, NEXT DAY AIR THE CHECK.
' STEP 8. PROFIT      /O\
'                 /\_----_/\
'               ( o )    ( o )

Public Sub Begin()
'declare constants

doDeletes
doAdds

MsgBox ("Do your maintenance to the newly modified file. When this is completed run the batch file creation module.")
End Sub

Function doAdds()


Dim wksPay As Worksheet
Dim wksListing As Worksheet

Dim LastRowPayFile As Long
Dim LastRowOurListing As Long
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wksPay = WB.Sheets(1)
Set wksListing = WB.Sheets("listing")

LastRowPayFile = wksPay.Cells(wksPay.Rows.count, "A").End(xlUp).ROW
LastRowOurListing = wksListing.Cells(wksListing.Rows.count, "A").End(xlUp).ROW

Dim rowRangePayFile As Range
Dim rowRangeListing As Range
'colum B holds the parcel #
Set rowRangePayFile = wksPay.Range("C1:C" & LastRowPayFile)
'this is our listing, so far is column H that has the parcel #
Set rowRangeListing = wksListing.Range("H1:H" & LastRowOurListing)

Dim Match As Boolean
Match = False
Dim currentAccount As String
Dim currentParcel As String

Dim currentRow As Integer
Dim currentEndingRow As Integer
currentEndingRow = LastRowPayFile

Dim testValue As Integer
testValue = 0
currentRow = 1
For Each ListRow In rowRangeListing

    currentParcel = ListRow.Value
    currentAccount = wksListing.Cells(currentRow, 1)

    For Each reqRow In rowRangePayFile
        If currentParcel = reqRow.Value Then
            Match = True
            'add account number from listing since we are already on it
            'wksPay.Cells(currentRow, 5).Value = currentAccount
            
        End If
        
    Next reqRow
    
    If Match = False Then
        addLine currentEndingRow, currentParcel, currentAccount
        currentEndingRow = currentEndingRow + 1
        testValue = testValue + 1
    End If
    
    Match = False
    currentRow = currentRow + 1
Next ListRow

MsgBox ("Added " & testValue & " parcels to end of listing.")
End Function

Function addLine(ROW As Integer, PARCEL As String, ACCOUNT As String)
Dim wks As Worksheet
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wks = WB.Sheets(1)

wks.Cells(ROW + 1, 3).Value = PARCEL
wks.Cells(ROW + 1, 5).Value = ACCOUNT
End Function

Function doDeletes()


Dim wksPayFile As Worksheet
Dim wksListing As Worksheet

Dim LastRowRequestFile As Long
Dim LastRowOurListing As Long
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wksRequest = WB.Sheets(1)
Set wksListing = WB.Sheets("listing")

LastRowRequestFile = wksRequest.Cells(wksRequest.Rows.count, "A").End(xlUp).ROW
LastRowOurListing = wksListing.Cells(wksListing.Rows.count, "A").End(xlUp).ROW

Dim rowRangePayFilet As Range
Dim rowRangeListing As Range
Set rowRangePayFile = wksRequest.Range("C1:C" & LastRowRequestFile)
Set rowRangeListing = wksListing.Range("H1:H" & LastRowOurListing)


Dim Match As Boolean
Match = False

Dim nameLineOne As String
Dim nameLineTwo As String
Dim currentParcel As String
Dim currentAccountNumber As String

Dim currentRow As Integer
Dim currentEndingRow As Integer
Dim rowCounter As Integer

currentEndingRow = LastRowRequestFile

Dim testValue As Integer
rowCounter = 1
testValue = 0
currentRow = 1

For Each reqRow In rowRangePayFile
    currentParcel = reqRow.Value
    For Each ListRow In rowRangeListing
        If currentParcel = ListRow.Value Then
            Match = True
            currentAccountNumber = wksListing.Cells(ListRow.ROW, 1).Value
        End If
        
    'Inner loop end
    Next ListRow
    
    ' process the inner loops findings
    If Match = False Then
        highlightLine reqRow.ROW
        testValue = testValue + 1
    End If
    If Match = True Then
        addAccount reqRow.ROW, currentAccountNumber
    End If
    Match = False
    currentRow = currentRow + 1
    
'Outter loop end
Next reqRow

MsgBox ("Highlighted " & testValue & " parcels for deletion")
End Function



Function highlightLine(ROW As Integer)

Dim wks As Worksheet
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wks = WB.Sheets(1)
wks.Cells(ROW, 1).EntireRow.Interior.ColorIndex = 6
End Function

Function addAccount(ROW As Integer, ACCOUNT As String)

Dim wks As Worksheet
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wks = WB.Sheets(1)
wks.Cells(ROW, 5).Value = ACCOUNT

End Function
