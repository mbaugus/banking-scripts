' Module to read Hendricks listing and compare to my datawarehouse report (a current hendricks listing)
' This will read each line from their listing, and subsequently compare it to our listing
' If a parcel found on their request file is not found on our listing, we highlight the row
' Otherwise it simply continues to the next line.  Afterwards we will do the opposite and compare
' Our listing to the request file.  If we have a parcel that needs added (it was on our list but not theirs
' then it will be appended to the end of the request listing.  this will eliminate the need for a visual
' scan of each parcel that possibly could be deleted or not

'*************
' ** NOTE ****
' our listing needs to be named "listing", the only real requirement is that the parcels be in column 'H'
' the hendricks county request listing needs to be named "Union Savings Bank - Hendricks"
' these name definitions are required or the module will not be able to find the worksheets.
'
'*************
'*************


Public parcelColumnReport As Integer
Public parcelColumnRequestFile As Integer
Public nameColumnRequestFile As Integer



Public Sub Begin()
'declare constants
parcelColumnReport = 8
parcelColumnRequestFile = 1
nameColumnRequestFile = 3

doDeletes
doAdds

End Sub

Function doAdds()


Dim wksRequest As Worksheet
Dim wksListing As Worksheet

Dim LastRowRequestFile As Long
Dim LastRowOurListing As Long
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wksRequest = WB.Sheets(1)
Set wksListing = WB.Sheets("listing")

LastRowRequestFile = wksRequest.Cells(wksRequest.Rows.count, "A").End(xlUp).ROW
LastRowOurListing = wksListing.Cells(wksListing.Rows.count, "A").End(xlUp).ROW

Dim rowRangeRequest As Range
Dim rowRangeListing As Range
Set rowRangeRequest = wksRequest.Range("B1:B" & LastRowRequestFile)
Set rowRangeListing = wksListing.Range("H1:H" & LastRowOurListing)

Dim Match As Boolean
Match = False
Dim nameLineOne As String
Dim nameLineTwo As String
Dim currentParcel As String
Dim currentRow As Integer
Dim currentEndingRow As Integer
currentEndingRow = LastRowRequestFile

Dim testValue As Integer
testValue = 0
currentRow = 1
For Each ListRow In rowRangeListing
    currentParcel = ListRow.Value
    nameLineOne = wksListing.Cells(currentRow, 5)
    If wksListing.Cells(currentRow, 6) <> "" Then
        nameLineOne = nameLineOne & " & " & wksListing.Cells(currentRow, 6)
    End If

    For Each reqRow In rowRangeRequest
        If currentParcel = reqRow.Value Then
            Match = True
        End If
        
    Next reqRow
    
    If Match = False Then
        addLine currentEndingRow, currentParcel, nameLineOne
        currentEndingRow = currentEndingRow + 1
        testValue = testValue + 1
    End If
    
    Match = False
    currentRow = currentRow + 1
Next ListRow

MsgBox ("Added " & testValue & " parcels to end of listing.")
End Function


Function addLine(ROW As Integer, PARCEL As String, NAME As String)
Dim wks As Worksheet
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wks = WB.Sheets(1)

wks.Cells(ROW + 1, 2).Value = PARCEL
wks.Cells(ROW + 1, 3).Value = NAME
End Function



Function doDeletes()


Dim wksRequest As Worksheet
Dim wksListing As Worksheet

Dim LastRowRequestFile As Long
Dim LastRowOurListing As Long
Dim WB As Workbook

Set WB = ActiveWorkbook
Set wksRequest = WB.Sheets(1)
Set wksListing = WB.Sheets("listing")

LastRowRequestFile = wksRequest.Cells(wksRequest.Rows.count, "A").End(xlUp).ROW
LastRowOurListing = wksListing.Cells(wksListing.Rows.count, "A").End(xlUp).ROW

Dim rowRangeRequest As Range
Dim rowRangeListing As Range
Set rowRangeRequest = wksRequest.Range("B1:B" & LastRowRequestFile)
Set rowRangeListing = wksListing.Range("H1:H" & LastRowOurListing)

Dim Match As Boolean
Match = False
Dim nameLineOne As String
Dim nameLineTwo As String
Dim currentParcel As String
Dim currentRow As Integer
Dim currentEndingRow As Integer
Dim rowCounter As Integer
currentEndingRow = LastRowRequestFile

Dim testValue As Integer
rowCounter = 1
testValue = 0
currentRow = 1

For Each reqRow In rowRangeRequest
    currentParcel = reqRow.Value
    
    For Each ListRow In rowRangeListing
        If currentParcel = ListRow.Value Then
            Match = True
        End If
        
    'Inner loop end
    Next ListRow
    
    ' process the inner loops findings
    If Match = False Then
        highlightLine reqRow.ROW
        testValue = testValue + 1
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
