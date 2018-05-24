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
' ----------+-----------------------------------------------------------------------------------------------------
'    Author | G Bishop
'      Date | 08/07/2016
'      Name | ThisWorkbook
'   Purpose |
'     To Do |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' G Bishop  | 01-05-2018 | Various changes throughout 2018
' ----------+------------+----------------------------------------------------------------------------------------


' ----------+-------------------------------------------------------------------------------------------------------
' Procedure | Workbook_Open
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' G Bishop  | 08-07-2016 | now recreates toolbar
' ----------+------------+----------------------------------------------------------------------------------------
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    
'    Call populateHDCmboBox("InternalParameters", Worksheets("InternalParameters").cmboHighLightOrDelete)
'    Call populateTBCmboBox("InternalParameters", Worksheets("InternalParameters").cmboShowToolbarDescription)
'    Call populateCDCmboBox("InternalParameters", Worksheets("InternalParameters").cmboDifferences)
'    Call populateTTCmboBox("InternalParameters", Worksheets("InternalParameters").cmboDisplayTimeTaken)
'    Call populateModeCmboBox("InternalParameters", Worksheets("InternalParameters").cmboModeAorB)
    
    Call populateHDCmboBox(Worksheets("InternalParameters").cmboHighLightOrDelete)
    Call populateTBCmboBox(Worksheets("InternalParameters").cmboShowToolbarDescription)
    Call populateCDCmboBox(Worksheets("InternalParameters").cmboDifferences)
    Call populateTTCmboBox(Worksheets("InternalParameters").cmboDisplayTimeTaken)
    Call populateModeCmboBox(Worksheets("InternalParameters").cmboModeAorB)
    Call reCreateToolBar
    Application.ScreenUpdating = False
End Sub
    
