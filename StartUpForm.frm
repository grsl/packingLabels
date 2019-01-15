VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartUpForm 
   Caption         =   "Create Packing Labels"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9615
   OleObjectBlob   =   "StartUpForm.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "StartUpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Added for the Search Functionality
' 2017/11/22
' Simon Long
'
'

Private Sub tbWOSearch_Change()
'
' Added for the Search Functionality
' 2018/01/04
' Simon Long
'
    Dim dbRecordSet As Recordset
    Dim str, searchItem As String
    Dim hDb As adodb.Connection
    Dim x As Integer
    
    Set hDb = getDBHandle
    Debug.Print hDb
     
    Let searchItem = UCase(StartUpForm.tbWOSearch.Value)
    Debug.Print searchItem
    
    If StartUpForm.tbWOSearch.Value = "" Then
    
        'Reset the listbox to all the available options.
        ' Get the list of Works Orders from the Sage database view vw_WorksOrderNumbers
        StartUpForm.lstWorksOrderNumber.Clear
        Call getWorksOrderNumbers(hDb)
  
    ElseIf (Len(searchItem) > 1) Then
        Debug.Print "Search Term " & searchItem
        
        Set hRs = getDBHandle.Execute("SELECT * FROM dbo.vw_WorksOrderNumber WHERE dbo.vw_WorksOrderNumber.WorksOrderNumber LIKE 'WO" & searchItem & "%'")
        
        If Not hRs.EOF Then
            str = hRs.GetString(adClipString, , " ", , "")
        
            'As there is only a single field we need to split on vbCr (end of row).
            arr = Split(str, vbCr, , vbTextCompare)
    
            StartUpForm.lstWorksOrderNumber.Clear
        
            For x = 0 To UBound(arr)
                'Add the Works Orders to the combo box.
                StartUpForm.lstWorksOrderNumber.AddItem (arr(x))
                'Debug.Print arr(x)
            Next x
        'Else
        '    StartUpForm.tbWOSearch.Value = ""
        '    MsgBox "Invalid value." & vbCrLf & "Please enter a Works Order Number.", vbOKOnly
        End If
        
    End If
    
    Let searchItem = ""
    
End Sub

Private Sub tbProductSearch_Change()
    
    Dim dbRecordSet As Recordset
    Dim str, searchItem As String
    Dim hDb As adodb.Connection
    Dim x As Integer
    
    Set hDb = getDBHandle()
     
    Let searchItem = UCase(StartUpForm.tbProductSearch.Value)
    
    If StartUpForm.tbProductSearch.Value = "" Then
        'Reset the listbox to all the available options.
        
        StartUpForm.lstProductCode.Clear
        Call getProductCodes(hDb)
    
    ElseIf (Len(searchItem) > 2) Then
    
        Set hRs = getDBHandle.Execute("SELECT * FROM dbo.vw_ProductCodes WHERE dbo.vw_ProductCodes.ProductCode LIKE '" & searchItem & "%'")
           
    '    If hRs.RecordCount > 0 Then
        If Not hRs.EOF Then
            str = hRs.GetString(adClipString, , " ", , "")
        
            'As there is only a single field we need to split on vbCr (end of row).
            arr = Split(str, vbCr, , vbTextCompare)
    
            StartUpForm.lstProductCode.Clear
        
            For x = 0 To UBound(arr)
                'Add the Works Orders to the combo box.
                StartUpForm.lstProductCode.AddItem (arr(x))
                'Debug.Print arr(x)
            Next x
        
        'Else
         '   StartUpForm.tbProductSearch.Value = ""
         '   MsgBox "Invalid value. vbCR Please enter a Works Order Number.", vbOKOnly
        End If
    
    End If
        
        searchItem = ""
        
End Sub

Private Sub btnCancel_Click()
    Application.DisplayAlerts = False
    StartUpForm.Hide
    Application.DisplayAlerts = True
End Sub

Private Sub btnSave_Click()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub

Private Sub btnClear_Click()
    'Clear all cells
    Worksheets("LabelData").Cells.Clear
    
    ' Set the text boxes to zero.
    Dim cntrl As MSForms.Control
        For Each cntrl In StartUpForm.Controls
            If TypeOf cntrl Is TextBox Then
                cntrl.tabkeybehaviour = False
            End If
        Next cntrl
    
    With StartUpForm
        ' Set the default in the three list boxes to the first item.
        .lstWorksOrderNumber.ListIndex = 0
        .lstWorksOrderNumber.Selected(0) = True
        .lstWeekNumber.ListIndex = 0
        .lstWeekNumber.Selected(0) = True
        .lstProductCode.ListIndex = 0
        .lstProductCode.Selected(0) = True
        .lstProductCode.SetFocus
    
        ' Initialise the text boxes to zero.
        .numberOfPumps.Value = 0
        .numberOfPumpsPerBox = 0
        .txtSerialNumberStart = 0
        
        ' Clear the suffixes
        .txbProductCodeSuffix = ""
        .txbSerialNumberSuffix = ""
        
        ' Clear the Search boxes.
        .tbProductSearch = ""
        .tbWOSearch = ""
    End With
End Sub

Private Sub btnPrintLabels_Click()
    ' //TODO
    ' Add functionality to print the labels straight from the user form rather than having
    ' the user open a Word document and processing the mailing merge to generate the labels.
    
    MsgBox "This button doesn't do anything."
End Sub

Private Sub CreateLabelData_Click()
    Call createData
End Sub

Private Sub lstProductCode_Change()
    Dim Selected As Boolean
    Selected = False
    
    For x = 0 To lstProductCode.ListCount - 1
        If lstProductCode.Selected(x) = True Then
            Selected = True
            bProductCode = True
        End If
    Next x
    
    If Not Selected Then
        StartUpForm.lstProductCode.SetFocus
        'MsgBox "Please select a product code."
    End If
End Sub

Private Sub lstWorksOrderNumber_Click()
    Dim Selected As Boolean
    Selected = False
    
    For x = 0 To lstWorksOrderNumber.ListCount - 1
        If lstWorksOrderNumber.Selected(x) = True Then
            Selected = True
            bWorksOrder = True
        End If
    Next x
    
    If Not Selected Then
        StartUpForm.lstWorksOrderNumber.SetFocus
        'MsgBox "Please select a works order number."
    End If
End Sub

Private Sub numberOfPumps_KeyPress(ByVal key As MSForms.ReturnInteger)
    If key < vbKey0 Or key > vbKey9 Then
        key = 0 ' this prevents the non-numeric data from showing up in the TextBox
        MsgBox "You can only enter numbers"
    End If
End Sub

Private Sub numberOfPumpsPerBox_KeyPress(ByVal key As MSForms.ReturnInteger)
    If key < vbKey0 Or key > vbKey9 Then
        key = 0 ' this prevents the non-numeric data from showing up in the TextBox
        MsgBox "You can only enter numbers"
    End If
End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Activate()
    Dim cntrl As MSForms.Control
    For Each cntrl In StartUpForm.Controls
        If TypeOf cntrl Is TextBox Then
            cntrl.tabkeybehaviour = False
        End If
    Next cntrl
    StartUpForm.lstProductCode.SetFocus
    
    ' Set the default item in the three list boxes to the first one.
    StartUpForm.lstProductCode.Selected(0) = True
    StartUpForm.lstWorksOrderNumber.Selected(0) = True
    StartUpForm.lstWeekNumber.Selected(0) = True
    
    ' Initialise the text boxes to zero.
    StartUpForm.numberOfPumps.Value = 0
    StartUpForm.numberOfPumpsPerBox = 0
    
    ' Initialise the search boxes.
    StartUpForm.tbProductSearch.Value = ""
    StartUpForm.tbWOSearch.Value = ""
    
End Sub

Private Sub UserForm_Terminate()
    Call btnCancel_Click
End Sub
