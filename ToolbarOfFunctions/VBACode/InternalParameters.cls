VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InternalParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmboModeAorB_Change()
    If Len(Trim(cmboModeAorB.Value)) > 0 Then
        Call handleRangesSub("Set", "rangeDelBlankLinesModeAorB", CStr(cmboModeAorB.Value))
        Call destroyToolBar
        Call createToolBar
    End If
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Private Sub cmboShowToolbarDescription_Change()
    If Len(Trim(cmboShowToolbarDescription.Value)) > 0 Then
        Call handleRangesSub("Set", "rangeShowDescriptionOption", CStr(cmboShowToolbarDescription.Value))
        Call destroyToolBar
        Call createToolBar
    End If
End Sub
'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Private Sub cmboHighLightOrDelete_Change()
    If Len(Trim(cmboHighLightOrDelete.Value)) > 0 Then
        Call handleRangesSub("Set", "rangeHighlightOrDeleteOption", CStr(cmboHighLightOrDelete.Value))
        Call destroyToolBar
        Call createToolBar
    End If
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Private Sub cmboDifferences_Change()
    If Len(Trim(cmboDifferences.Value)) > 0 Then
        Call handleRangesSub("Set", "rangeCompareOption", CStr(cmboDifferences.Value))
        Call destroyToolBar
        Call createToolBar
    End If
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Private Sub cmboDisplayTimeTaken_Change()
    If Len(Trim(cmboDisplayTimeTaken.Value)) > 0 Then
        Call handleRangesSub("Set", "rangeTimeTaken", CStr(cmboDisplayTimeTaken.Value))
    End If
End Sub

