Attribute VB_Name = "Module1"
Option Explicit

Public Sub createWeekData()
    Dim x As Integer
    'Add data to the Week Number combobox.
    For x = 1 To 52
        StartUpForm.lstWeekNumber.AddItem ("Week " & x)
    Next x
End Sub

Public Sub getWorksOrderNumbers(dbConnection As adodb.Connection)
    Dim arr As Variant
    Dim x   As Integer
    Dim str As String
    Dim dbRecordSet As Recordset
        
    Set dbRecordSet = dbConnection.Execute("Select * from vw_WorksOrderNumber")
    
    ' Find out if the connection is valid.
    If dbConnection.State = adStateOpen Then
        str = dbRecordSet.GetString(adClipString, , " ", , "")
        
        'As there is only a single field we need to split on vbCr (end of row).
        arr = Split(str, vbCr, , vbTextCompare)
        
        For x = 0 To UBound(arr)
            'Add the Works Orders to the combo box.
            StartUpForm.lstWorksOrderNumber.AddItem (arr(x))
        Next x
    Else
        MsgBox "Can't connect to the database." & vbCr & vbCr & "Please contact your system administrator."
    End If
    
    dbRecordSet.Close
End Sub

Public Sub getProductCodes(dbConnection As adodb.Connection)
    Dim arr As Variant
    Dim x As Integer
    Dim str As String
    Dim dbRecordSet As Recordset
    
    Set dbRecordSet = dbConnection.Execute("Select * from vw_ProductCodes")
    
    ' Find out if the connection is valid.
    If dbConnection.State = adStateOpen Then
        str = dbRecordSet.GetString(adClipString)
        arr = Split(str, vbCr, , vbTextCompare)

        For x = 0 To UBound(arr)
            StartUpForm.lstProductCode.AddItem (arr(x))
        Next x
    Else
        MsgBox "Can't connect to the database." & vbCr & vbCr & "Please contact your system administrator."
    End If
    
    dbRecordSet.Close

End Sub

Public Sub createData()
    Dim woOffset, xlOffset, pumpsPerBox, pumpsOrdered, numberOfBoxes, remainder, weekNumber, worksOrderNumber, firstSerialNumber, lastSerialNumber, x As Integer
    Dim productCode, worksOrder, worksOrderPrefix, serialNumberSuffix, productCodeSuffix As String
    Dim currentYear As Date
    Dim sscorLabels As Boolean
    Dim Rng1 As Range
    
    ' Turn of updating to speed up the process.
    Application.ScreenUpdating = False
    
    'Clear all cells
    Worksheets("LabelData").Cells.Clear
    
    'The number of characters at the start of the Works Order Number that are not numeric.
    woOffset = 2
    'Start after the header line.
    xlOffset = 1
    
    productCode = StartUpForm.lstProductCode
    worksOrder = StartUpForm.lstWorksOrderNumber
    worksOrderPrefix = Left(worksOrder, woOffset)
    If Len(worksOrder) > 3 Then
        worksOrderNumber = Int(Mid(worksOrder, woOffset + 1))
    End If
    pumpsOrdered = Int(StartUpForm.numberOfPumps)
    pumpsPerBox = Int(StartUpForm.numberOfPumpsPerBox)
    weekNumber = Int(Right(StartUpForm.lstWeekNumber, 2))

    currentYear = Mid(Now(), 9, 2)
    productCodeSuffix = UCase(StartUpForm.txbProductCodeSuffix)
    serialNumberSuffix = UCase(StartUpForm.txbSerialNumberSuffix)
    sscorLabels = StartUpForm.chkSscor.Value
    
    If pumpsOrdered > 0 And pumpsPerBox > 0 Then
        numberOfBoxes = Int(pumpsOrdered / pumpsPerBox)
        remainder = pumpsOrdered Mod pumpsPerBox
    End If

    If remainder > 0 Then
        numberOfBoxes = numberOfBoxes + 1
    End If

    'Write the headers to the worksheet.
    With Worksheets("LabelData")
        Range("A1").Value = "Product Code"
        Range("B1").Value = "Works Order No."
        Range("C1").Value = "First Serial Number in the Box"
        Range("D1").Value = "Last Serial Number in the Box"
        Range("E1").Value = "Number of Pumps in the Box"
        Range("F1").Value = "Box X of Y"
    End With

    If (StartUpForm.txtSerialNumberStart.Text = "") Then
        firstSerialNumber = 1
    Else
        If (Int(StartUpForm.txtSerialNumberStart.Text) < 1) Then
            firstSerialNumber = 1
        Else
            firstSerialNumber = Int(StartUpForm.txtSerialNumberStart.Text)
        End If
    End If
     
    lastSerialNumber = firstSerialNumber + Int(pumpsPerBox) - 1

    For x = 1 To numberOfBoxes Step 1
        ' Information required on the labels: Product Code PRODUCT; Works Order Number WO12345;
        ' Serial Number From: 16130051 12345 to: 16130100 12345. This format is for RD1, APC and HiP.
        ' The format for SSCOR pumps is WO12345-0001 i.e. works order no. hyphen and pump number.
        ' Number of Pumps in the Box  50
        ' Box 2 of 30
        Cells(x + xlOffset, 1).Value = productCode & productCodeSuffix
        Cells(x + xlOffset, 2).Value = worksOrder
        Cells(x + xlOffset, 6).Value = "Box " & x & " of " & numberOfBoxes
        
        If ((x = numberOfBoxes) And (remainder > 0)) Then
            lastSerialNumber = Int(lastSerialNumber - pumpsPerBox + remainder)
            With Worksheets("LabelData")
                If sscorLabels Then
                    Cells(x + xlOffset, 3).Value = worksOrderPrefix & worksOrderNumber & " " & format(firstSerialNumber, "0000") & serialNumberSuffix
                    Cells(x + xlOffset, 4).Value = worksOrderPrefix & worksOrderNumber & " " & format(lastSerialNumber, "0000") & serialNumberSuffix
                    Cells(x + xlOffset, 5).Value = remainder
                    Cells(x + xlOffset, 6).Value = "Box " & x & " of " & numberOfBoxes
                Else
                    Cells(x + xlOffset, 3).Value = format(currentYear, "00") & format(weekNumber, "00") & format(firstSerialNumber, "0000") & " " & worksOrderNumber & serialNumberSuffix
                    Cells(x + xlOffset, 4).Value = format(currentYear, "00") & format(weekNumber, "00") & format(lastSerialNumber, "0000") & " " & worksOrderNumber & serialNumberSuffix
                    Cells(x + xlOffset, 5).Value = remainder
                    Cells(x + xlOffset, 6).Value = "Box " & x & " of " & numberOfBoxes
                End If
            End With
        Else
            With Worksheets("LabelData")
                If sscorLabels Then
                    Cells(x + xlOffset, 3).Value = worksOrderPrefix & worksOrderNumber & " " & format(firstSerialNumber, "0000") & serialNumberSuffix
                    Cells(x + xlOffset, 4).Value = worksOrderPrefix & worksOrderNumber & " " & format(lastSerialNumber, "0000") & serialNumberSuffix
                    Cells(x + xlOffset, 5).Value = pumpsPerBox
                Else
                    Cells(x + xlOffset, 3).Value = format(currentYear, "00") & format(weekNumber, "00") & format(firstSerialNumber, "0000") & " " & worksOrderNumber & serialNumberSuffix
                    Cells(x + xlOffset, 4).Value = format(currentYear, "00") & format(weekNumber, "00") & format(lastSerialNumber, "0000") & " " & worksOrderNumber & serialNumberSuffix
                    Cells(x + xlOffset, 5).Value = pumpsPerBox
                End If


            End With
            firstSerialNumber = lastSerialNumber + 1
            lastSerialNumber = Int(lastSerialNumber + pumpsPerBox)
        End If
    Next x
    
    Worksheets("LabelData").Cells(1, 1).Activate
    
    'Save the data in the existing document.
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    ' Turn updating back on to display the data.
    Application.ScreenUpdating = True
    
End Sub

Public Sub getWorksOrder(dbConnection As adodb.Connection, criterion As String)
    Dim arr As Variant
    Dim x   As Integer
    Dim str As String
    Dim dbRecordSet As Recordset
    criterion = " X82%"
        
    Set dbRecordSet = dbConnection.Execute("Select * from vw_WorksOrderNumber" & criterion)
    
    ' Find out if the connection is valid.
    If dbConnection.State = adStateOpen Then
        str = dbRecordSet.GetString(adClipString, , " ", , "")
        
        'As there is only a single field we need to split on vbCr (end of row).
        arr = Split(str, vbCr, , vbTextCompare)
        
        For x = 0 To UBound(arr)
            'Add the Works Orders to the combo box.
            StartUpForm.lstWorksOrderNumber.AddItem (arr(x))
        Next x
    Else
        MsgBox "Can't connect to the database." & vbCr & vbCr & "Please contact your system administrator."
    End If
    
    dbRecordSet.Close
End Sub

Public Function getDBHandle() As adodb.Connection
    
    Set getDBHandle = New adodb.Connection
    getDBHandle.ConnectionString = "driver={SQL Server};server=CAP-APPS64;uid=sa;pwd=CharlesA1;database=CAP-Live"
    getDBHandle.Open
    
    ' Find out if the connection is valid.
    'If dbConnection.State <> adStateOpen Then
    '    getDBHandle.Close
    '    MsgBox "Can't connect to the database." & vbCr & vbCr & "Please contact your system administrator."
    'End If
    
End Function


