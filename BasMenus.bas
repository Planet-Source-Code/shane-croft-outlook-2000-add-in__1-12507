Attribute VB_Name = "BasMenus"
Option Explicit

Private Function IsMenuThere(sMenu As String, _
            sName As String) As Boolean
  'Returns True if menu sName exists in sMenu.
  Dim oCB As Office.CommandBar
  Dim oControl As Office.CommandBarControl
  
  IsMenuThere = False
  Set oCB = ActiveExplorer.CommandBars(sMenu)
  For Each oControl In oCB.Controls
    If oControl.Caption = sName Then
      IsMenuThere = True
      Exit For
    End If
  Next

  Set oCB = Nothing
  Set oControl = Nothing
End Function

Private Function AddMenu(sCBName As String, _
  sName As String, sTag As String) As CommandBarControl
  'Add the menu named in sName to the
  'Outlook command bar named in sCBName.
  
  Dim oBar As Office.CommandBar
  Dim oControl As Office.CommandBarControl
  Dim lCount As Long

  Set oBar = ActiveExplorer.CommandBars(sCBName)
  lCount = oBar.Controls.Count
  With oBar
    'A new menu is a popup control
    'Add it before the Help menu, which is last, and make the menu temp
    Set oControl = .Controls.Add(msoControlPopup, _
      , , lCount - 1, True)
    
    oControl.Caption = sName
    'Any Tag must be unique
    oControl.Tag = sTag
  End With
  Set AddMenu = oControl

  Set oBar = Nothing
  Set oControl = Nothing
End Function

Public Sub AddSMCMenu()
  Dim oButton As Office.CommandBarButton
  Dim oCB As Office.CommandBarPopup
  Dim oBar As Office.CommandBar
  Dim bResult As Boolean
  'Add a Temp menu
  bResult = IsMenuThere("Menu Bar", "&SMC")
  If bResult = False Then
    Set oCB = AddMenu("Menu Bar", "&SMC", "SMCTag")
    With oCB
Set oButton = .Controls.Add(msoControlButton)
      With oButton
      .Caption = "Settings..."
      .Tag = "Settings..."
      .FaceId = 548
      .HyperlinkType = msoCommandBarButtonHyperlinkOpen
      .ToolTipText = App.Path & "\" & "Settings.exe"
      .Style = msoButtonIconAndCaption
    End With
    End With

      With oCB
    Set oButton = .Controls.Add(msoControlButton)
    With oButton
      .BeginGroup = True
      .Caption = "SMC WebSite..."
      .Tag = "SMCwebSite"
      .FaceId = 1015
      .HyperlinkType = msoCommandBarButtonHyperlinkOpen
      .ToolTipText = "http://www.croftssoftware.com"
      .Style = msoButtonIconAndCaption
    End With
    End With
  Else
    Set oBar = ActiveExplorer.CommandBars("Menu Bar")
    Set oCB = oBar.FindControl(, , "SMCTag")
  End If
  

  Set oButton = Nothing
  Set oCB = Nothing
  Set oBar = Nothing
End Sub
