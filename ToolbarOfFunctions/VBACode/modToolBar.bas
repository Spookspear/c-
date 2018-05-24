Attribute VB_Name = "modToolBar"
'-----------------------------------------------------------------------------
' Author:          Date
' Grant V Bishop   19th March 2010
'
' Description:
' This will contain all my reusable code
'-----------------------------------------------------------------------------
'Modified by   | Date       | Reason
'-----------------------------------------------------------------------------
' G Bishop       04/05/2013   Added a Read in group Membership from selection button
' G Bishop       01-05-2018   Added ping server option
'-----------------------------------------------------------------------------
' TODO:
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
Option Explicit
Const ToolBarName = "Excel Tools and Utilities Toolbar"

' ----------+-------------------------------------------------------------------------------------------------------
' Procedure | ReCreateToolBar
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub reCreateToolBar()
    Call destroyToolBar
    Call createToolBar
End Sub
'--------------------------------------------------------------------------
' Destroy the toolbar
'--------------------------------------------------------------------------
Sub destroyToolBar()
Dim strToolBarName  As String
    strToolBarName = ToolBarName
    On Error Resume Next
    Application.CommandBars(strToolBarName).Delete
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub createToolBar()
Dim strToolBarName As String
    'MsgBox "Is this being made twice?"
    strToolBarName = ToolBarName
    If Not IsToolbar(strToolBarName) Then
        Call newToolBar(strToolBarName)
    Else
        Application.CommandBars(ToolBarName).Visible = True
    End If
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Public Function IsToolbar(strToolBarName As String)
    On Error Resume Next
    IsToolbar = IsObject(CommandBars(strToolBarName))
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub newToolBar(strToolBarName As String)
'Dim cMenu1 As CommandBarControl
Dim cbMainMenuBar As CommandBar
'Dim iHelpMenu As Currency
'Dim cbcCustomMenu As CommandBarControl
Dim myNewButton As CommandBarButton
'Dim strWhichSheet As String
Dim intButtonCount As Currency
Dim lShowDescriptions As Boolean
Dim intStyle As Currency
Dim strHighlighOrDeleteDupliates As String
Dim strClearOrColour As String
Dim strModeAorB As String
'Dim barHeight As Integer            ' 2007 mods
    
    intButtonCount = 0
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Show Tool Bar Descriptions
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    lShowDescriptions = getRangeValue("rangeShowDescriptionOption")
    If lShowDescriptions Then
        intStyle = msoButtonIconAndCaption
    Else
        intStyle = msoButtonIcon
    End If
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' variables for toolbar text as chosen on: InternalParameters
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    strHighlighOrDeleteDupliates = getRangeValue("rangeHighlightOrDeleteOption")    ' Say Highlight or Delete (Duplicate Rows)
    strClearOrColour = getRangeValue("rangeCompareOption")                          ' Say Clear or Colour (on Differences)
    strModeAorB = getRangeValue("rangeDelBlankLinesModeAorB")                       ' Say what delete Mode A or B

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Destroy the toolbar
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Call destroyToolBar
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Add the new Toolbar (empty frame for the toolbar)
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Application.CommandBars.Add(Name:=strToolBarName).Visible = True
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Instantiate the toolbar
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Set cbMainMenuBar = Application.Application.CommandBars(ToolBarName)
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Populate File button
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        '.Height = barHeight * 2
        '.Width = 20
        .Caption = "Read Folders"
        .TooltipText = "Reads selected directory into worksheet"
        .Style = intStyle
        '.Style = msoButtonIconAndWrapCaption
        .FaceId = 270
        '.Height = 160
        
        '.Style = msoButtonIconAndCaptionBelow
        '.Style = msoButtonCaption
        '.Style = msoButtonIconAndWrapCaption
        '.Style = msoButtonIconAndWrapCaptionBelow
        
        .OnAction = "btnPopulateSheetFromFolder"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Compare Sheets
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        If strClearOrColour = "Colour" Then
            .Caption = "&Compare (Colour)"
            .TooltipText = "Compare Sheets (and colour the duplicate lines)"
        Else
            .Caption = "&Compare (Clear)"
            .TooltipText = "Compare Sheets (and clear the duplicate lines)"
        End If
        .Style = intStyle
        .FaceId = 694
        '.FaceId = 346
        .OnAction = "btnCompareSheets"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Zap Current Sheet
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        '.Caption = "&Zap Current Sheet"
        .Caption = ""       ' taken out to squeeze in new button
        .TooltipText = "Zap Current Sheet"
        .Style = intStyle
        .FaceId = 643
        '.FaceId = 7707                  '346
        .OnAction = "btnZapSheet"
    End With

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Del Blank Lines
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "&Del Blanks Mode:" & strModeAorB
        .TooltipText = "Delete Blank Lines using Mode: " & strModeAorB
        .Style = intStyle
        '.FaceId = 705
        .FaceId = 2055
        .OnAction = "btnDelBlankLines"
    End With

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Del Single Duplicates
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        If strHighlighOrDeleteDupliates = "Delete" Then
            .Caption = "Duplicates (Cols: Single): &Del"
            .TooltipText = "Delete Duplicates Rows"
        ElseIf strHighlighOrDeleteDupliates = "Highlight" Then
            .Caption = "Duplicates (Cols: Single): &Colour"
            .TooltipText = "Highlight the Duplicates Rows"
        ElseIf strHighlighOrDeleteDupliates = "ClearCell" Then
            .Caption = "Duplicates (Cols: Single): &Clear"
            .TooltipText = "Clear out the contents of the Duplicates Rows"
        End If
        .Style = intStyle
        '.FaceId = 6134
        .FaceId = 706
        .OnAction = "btnDealWithSingleDuplicates"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Del Many Duplicates
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        If strHighlighOrDeleteDupliates = "Delete" Then
            .Caption = "Duplicates (Cols: Many): &Del"
            .TooltipText = "Delete Duplicates Rows"
            
        ElseIf strHighlighOrDeleteDupliates = "Highlight" Then
            .Caption = "Duplicates (Cols: Many): &Colour"
            .TooltipText = "Highlight the Duplicates Rows"
            
        ElseIf strHighlighOrDeleteDupliates = "ClearCell" Then
            .Caption = "Duplicates (Cols: Many): &Clear"
            .TooltipText = "Clear out the contents of the Duplicates Rows"
            
        End If
        .Style = intStyle
        '.FaceId = 6134
        .FaceId = 706
        .OnAction = "btnDealWithManyDuplicates"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Populate Sheet with AD Group Members
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "&AD Group Members"
        .TooltipText = "Read in members of supplied Active Directory group"
        .Style = intStyle
        .FaceId = 2152
        .OnAction = "btnLoadADGroupIntoSpreadsheet"
    End With

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Populate Sheet with AD Group Members from active cell
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "AD &Group Members - Active Cell"
        .TooltipText = "Read in members of AD group from selected cell"
        .Style = intStyle
        .FaceId = 6134
        .OnAction = "btnLoadADGroupIntoSpreadsheetActiveCell"
    End With
   
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Read Users Group Membership
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "&Users Group Membership"
        .TooltipText = "The groups the user is a member of"
        .Style = intStyle
        .FaceId = 327
        .OnAction = "btnReadUsersGroupMembership"
    End With
    
    
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Read Users Group Membership
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "Get Details from AD Name "
        .TooltipText = "The user details"
        .Style = intStyle
        .FaceId = 329
        .OnAction = "btnReadUsers"
    End With
    
      
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Mypersonal timesheet option - remove if published
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Call addTimeSheet(cbMainMenuBar, intButtonCount, intStyle, True)
        
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Mypersonal timesheet option - remove if published
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    'MsgBox "Stop", , "1GVB1 01/05/2018"
    
    Call addPingSheetToToolbar(cbMainMenuBar, intButtonCount, intStyle, True)
      
      
    Application.CommandBars(ToolBarName).Position = msoBarTop
    
    ' Clean up the memory
    Set cbMainMenuBar = Nothing
    Set myNewButton = Nothing
        
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub NewToolBar_Old(strToolBarName As String)
'Dim cMenu1 As CommandBarControl
Dim cbMainMenuBar As CommandBar
'Dim iHelpMenu As Currency
'Dim cbcCustomMenu As CommandBarControl
Dim myNewButton As CommandBarButton
'Dim strWhichSheet As String
Dim intButtonCount As Currency
Dim lShowDescriptions As Boolean
Dim intStyle As Currency
Dim strHighlighOrDeleteDupliates As String
Dim strClearOrColour As String
Dim strModeAorB As String
'Dim barHeight As Integer            ' 2007 mods
    
    intButtonCount = 0
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Show Tool Bar Descriptions
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    lShowDescriptions = getRangeValue("rangeShowDescriptionOption")
    If lShowDescriptions Then
        intStyle = msoButtonIconAndCaption
    Else
        intStyle = msoButtonIcon
    End If
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' variables for toolbar text as chosen on: InternalParameters
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    strHighlighOrDeleteDupliates = getRangeValue("rangeHighlightOrDeleteOption")    ' Say Highlight or Delete (Duplicate Rows)
    strClearOrColour = getRangeValue("rangeCompareOption")                          ' Say Clear or Colour (on Differences)
    strModeAorB = getRangeValue("rangeDelBlankLinesModeAorB")                       ' Say what delete Mode A or B

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Destroy the toolbar
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Call destroyToolBar
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Add the new Toolbar (empty frame for the toolbar)
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Application.CommandBars.Add(Name:=strToolBarName).Visible = True
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Instantiate the toolbar
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    Set cbMainMenuBar = Application.Application.CommandBars(ToolBarName)
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Populate File button
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        '.Height = barHeight * 2
        '.Width = 20
        .Caption = "Read Folders"
        .TooltipText = "Reads selected directory into worksheet"
        .Style = intStyle
        '.Style = msoButtonIconAndWrapCaption
        .FaceId = 270
        '.Height = 160
        
        '.Style = msoButtonIconAndCaptionBelow
        '.Style = msoButtonCaption
        '.Style = msoButtonIconAndWrapCaption
        '.Style = msoButtonIconAndWrapCaptionBelow
        
        .OnAction = "btnPopulateSheetFromFolder"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Compare Sheets
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        If strClearOrColour = "Colour" Then
            .Caption = "&Compare (Colour)"
            .TooltipText = "Compare Sheets (and colour the duplicate lines)"
        Else
            .Caption = "&Compare (Clear)"
            .TooltipText = "Compare Sheets (and clear the duplicate lines)"
        End If
        .Style = intStyle
        .FaceId = 694
        '.FaceId = 346
        .OnAction = "btnCompareSheets"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Zap Current Sheet
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        '.Caption = "&Zap Current Sheet"
        .Caption = ""       ' taken out to squeeze in new button
        .TooltipText = "Zap Current Sheet"
        .Style = intStyle
        .FaceId = 643
        '.FaceId = 7707                  '346
        .OnAction = "btnZapSheet"
    End With

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Del Blank Lines
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "&Del Blanks Mode:" & strModeAorB
        .TooltipText = "Delete Blank Lines using Mode: " & strModeAorB
        .Style = intStyle
        '.FaceId = 705
        .FaceId = 2055
        .OnAction = "btnDelBlankLines"
    End With

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Del Single Duplicates
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        If strHighlighOrDeleteDupliates = "Delete" Then
            .Caption = "&Del Dups Rows (Single Col)"
            .TooltipText = "Delete Duplicates Rows"
        ElseIf strHighlighOrDeleteDupliates = "Highlight" Then
            .Caption = "&Colour Dups Rows (Single Col)"
            .TooltipText = "Highlight the Duplicates Rows"
        ElseIf strHighlighOrDeleteDupliates = "ClearCell" Then
            .Caption = "&Clear Dups Rows (Single Col)"
            .TooltipText = "Clear out the contents of the Duplicates Rows"
        End If
        .Style = intStyle
        '.FaceId = 6134
        .FaceId = 706
        .OnAction = "btnDealWithSingleDuplicates"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Del Many Duplicates
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        If strHighlighOrDeleteDupliates = "Delete" Then
            .Caption = "&Del Dups Rows (Many Cols)"
            .TooltipText = "Delete Duplicates Rows"
        ElseIf strHighlighOrDeleteDupliates = "Highlight" Then
            .Caption = "&Colour Rows (Many Cols)"
            .TooltipText = "Highlight the Duplicates Rows"
            
        ElseIf strHighlighOrDeleteDupliates = "ClearCell" Then
            .Caption = "&Clear Dups Rows (Many Cols)"
            .TooltipText = "Clear out the contents of the Duplicates Rows"
            
        End If
        .Style = intStyle
        '.FaceId = 6134
        .FaceId = 706
        .OnAction = "btnDealWithManyDuplicates"
    End With
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Populate Sheet with AD Group Members
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "&AD Group Members"
        .TooltipText = "Read in members of supplied Active Directory group"
        .Style = intStyle
        .FaceId = 2152
        .OnAction = "btnLoadADGroupIntoSpreadsheet"
    End With

    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Populate Sheet with AD Group Members from active cell
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "AD &Group Members - Active Cell"
        .TooltipText = "Read in members of AD group from selected cell"
        .Style = intStyle
        .FaceId = 6134
        .OnAction = "btnLoadADGroupIntoSpreadsheetActiveCell"
    End With
   
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Read Users Group Membership
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    intButtonCount = intButtonCount + 1
    Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
    With myNewButton
        .Caption = "&Users Group Membership"
        .TooltipText = "The groups the user is a member of"
        .Style = intStyle
        .FaceId = 327
        .OnAction = "btnReadUsersGroupMembership"
    End With
      
    Application.CommandBars(ToolBarName).Position = msoBarTop
    
    ' Clean up the memory
    Set cbMainMenuBar = Nothing
    Set myNewButton = Nothing
        
End Sub



