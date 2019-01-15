VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub Workbook_Open()
    
    Dim dbConnection As adodb.Connection
    Dim dbRecordSet As adodb.Recordset

    Worksheets("LabelData").Cells.Clear     ' Clear all cells on the worksheet of data.

    Call createWeekData                     ' Create the data for the Week drop down list
    
    ' Connect to the database.
    Set dbConnection = New adodb.Connection
    dbConnection.ConnectionString = "driver={SQL Server};server=CAP-APPS64;uid=sa;pwd=CharlesA1;database=CAP-Live"
    dbConnection.Open
      
    ' Get the list of Works Orders from the Sage database view vw_WorksOrderNumbers
    Call getWorksOrderNumbers(dbConnection)
  
    ' Create the data for the Week drop down list
    Call getProductCodes(dbConnection)

    dbConnection.Close
    
    Worksheets("LabelData").Cells(1, 1).Activate
    
    StartUpForm.Show
End Sub

