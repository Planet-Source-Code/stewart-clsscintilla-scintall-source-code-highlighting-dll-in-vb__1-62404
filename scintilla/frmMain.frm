VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6195
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11668
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1800
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportHTML 
         Caption         =   "Export to HTML"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "&Language"
      Begin VB.Menu mnuHighlighter 
         Caption         =   "&No Highlighters Installed"
         Index           =   0
      End
   End
   Begin VB.Menu mnuSamples 
      Caption         =   "Samples"
      Begin VB.Menu mnuDisplayText 
         Caption         =   "Display Text Value"
      End
      Begin VB.Menu mnuSetText 
         Caption         =   "Set Text"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetSel 
         Caption         =   "Get Sel Text"
      End
      Begin VB.Menu mnuSetSel 
         Caption         =   "Set Sel Text"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoIndent 
         Caption         =   "&AutoIndent"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuReadOnly 
         Caption         =   "&ReadOnly"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFixedFont 
         Caption         =   "Set Fixed Font"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextMenu 
         Caption         =   "&Context Menu"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEOLMode 
         Caption         =   "EOL Mode"
         Begin VB.Menu mnuCLRF 
            Caption         =   "CRLF"
         End
         Begin VB.Menu mnuCR 
            Caption         =   "CR"
         End
         Begin VB.Menu mnuLF 
            Caption         =   "LF"
         End
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWordWrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu mnuCallTip 
         Caption         =   "CallTip"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoComplete 
         Caption         =   "AutoComplete"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIgnoreAutoCompleteCase 
         Caption         =   "Ignore Auto Complete Case"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom In"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom Out"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewWhiteSpace 
         Caption         =   "View WhiteSpace"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowCallTips 
         Caption         =   "Show Call Tips"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindPrev 
         Caption         =   "Find &Previous"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSep18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto"
         Shortcut        =   ^G
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents sciMain As clsScintilla
Attribute sciMain.VB_VarHelpID = -1
Private strFile As String

Private Sub Form_Load()
Set sciMain = New clsScintilla
  sciMain.CreateScintilla Me
  sciMain.SetFixedFont "Courier New", 10

    
  ' Give the scrollbar a nice long width to handle a long line which may
  ' occur.
  sciMain.ScrollWidth = 10000
  'sciMain.Text = "This is a sample utilizing the Scintilla Wrapper Class. " & vbCrLf & "It shows all of the basic set forward abilities of the wrapper.  " & vbCrLf & "Though the wrapper doesn't incorperate everything Scintilla can handle " & vbCrLf & "it gives you easy access to a lot of it." & vbCrLf & vbCrLf & "http://www.ceditmx.com" & vbCrLf & vbCrLf & "http://www.vbaccelerator.com" & vbCrLf & vbCrLf & "http://www.scintilla.org"
  '+----------------------------------------+
  '| This is absolutly an imperative line   |
  '+----------------------------------------+
  sciMain.Attach Me
  '+----------------------------------------+
  '| This is absolutly an imperative line   |
  '+----------------------------------------+
  
  sciMain.Folding = True
  sciMain.ShowCallTips = True
  sciMain.LineNumbers = True
  sciMain.AutoIndent = True
  sciMain.SetMarginWidth MarginLineNumbers, 50
  
  sciMain.LoadAPIFile App.Path & "\api\cpp.api"
  
  LoadDirectory App.Path & "\highlighters"
  Call SetupMenu
  Call SetHighlighter(sciMain, "VB")
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  sciMain.SizeScintilla 0, 0, Me.ScaleWidth / Screen.TwipsPerPixelX, (Me.ScaleHeight / Screen.TwipsPerPixelY) - (stb.Height / Screen.TwipsPerPixelY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '+----------------------------------------+
  '| This is absolutly an imperative line   |
  '+----------------------------------------+
  sciMain.Detach Me
  '+----------------------------------------+
  '| This is absolutly an imperative line   |
  '+----------------------------------------+
End Sub

Private Sub mnuAutoComplete_Click()
  sciMain.ShowAutoComplete "else if then"
End Sub

Private Sub mnuAutoIndent_Click()
  mnuAutoIndent.Checked = Not mnuAutoIndent.Checked
  sciMain.AutoIndent = mnuAutoIndent.Checked
End Sub

Private Sub mnuCallTip_Click()
  sciMain.ShowCallTip "msgbox (Test1 as long, Test2 as long, Test3 as long)"
End Sub

Private Sub mnuCLRF_Click()
  sciMain.LineBreak = SC_EOL_CRLF
End Sub

Private Sub mnuContextMenu_Click()
  mnuContextMenu.Checked = Not mnuContextMenu.Checked
  sciMain.ContextMenu = mnuContextMenu.Checked
End Sub

Private Sub mnuCopy_Click()
  sciMain.Copy
End Sub


Private Sub mnuCR_Click()
  sciMain.LineBreak = SC_EOL_CR
End Sub

Private Sub mnuDisplayText_Click()
  MsgBox sciMain.Text
End Sub

Private Sub mnuExportHTML_Click()
  ExportToHTML App.Path & "\test.html", sciMain
End Sub

Private Sub mnuFind_Click()
  sciMain.DoFind
End Sub

Private Sub mnuFindNext_Click()
  sciMain.FindNext
End Sub

Private Sub mnuFindPrev_Click()
  sciMain.FindPrev
End Sub

Private Sub mnuFixedFont_Click()
  sciMain.SetFixedFont "Courier New", 10
End Sub

Private Sub mnuGetSel_Click()
  MsgBox sciMain.SelText
End Sub

Private Sub mnuGoto_Click()
  sciMain.DoGoto
End Sub

Private Sub mnuHighlighter_Click(index As Integer)
  SetHighlighter sciMain, mnuHighlighter(index).Tag
End Sub

Private Sub mnuIgnoreAutoCompleteCase_Click()
  mnuIgnoreAutoCompleteCase.Checked = Not mnuIgnoreAutoCompleteCase.Checked
  sciMain.IgnoreAutoCCase = mnuIgnoreAutoCompleteCase.Checked
End Sub

Private Sub mnuLF_Click()
  sciMain.LineBreak = SC_EOL_LF
End Sub

Private Sub mnuReadOnly_Click()
  mnuReadOnly.Checked = Not mnuReadOnly.Checked
  sciMain.ReadOnly = mnuReadOnly.Checked
  If sciMain.GetReadOnly = True Then
    stb.Panels(3).Text = "ReadOnly"
  Else
    stb.Panels(3).Text = ""
  End If
End Sub

Private Sub mnuReplace_Click()
  sciMain.DoReplace
End Sub

Private Sub mnuSave_Click()
  If strFile <> "" Then
    sciMain.SaveToFile strFile
  Else
    mnuSaveAs_Click
  End If
End Sub

Private Sub mnuSelectAll_Click()
  sciMain.SelectAll
End Sub

Private Sub mnuSetSel_Click()
  sciMain.SelText = "Hello world"
End Sub

Private Sub mnuSetText_Click()
  sciMain.Text = "Hello World"
End Sub


Private Sub mnuShowCallTips_Click()
  mnuShowCallTips.Checked = Not mnuShowCallTips.Checked
  sciMain.ShowCallTips = mnuShowCallTips.Checked
End Sub

Private Sub mnuViewWhiteSpace_Click()
  mnuViewWhiteSpace.Checked = Not mnuViewWhiteSpace.Checked
  sciMain.ViewWhiteSpace = mnuViewWhiteSpace.Checked
End Sub

Private Sub mnuWordWrap_Click()
  mnuWordWrap.Checked = Not mnuWordWrap.Checked
  sciMain.WordWrap = mnuWordWrap.Checked
End Sub

Private Sub mnuZoomIn_Click()
  sciMain.ZoomIn
  stb.Panels(4).Text = "Zoom: " & sciMain.GetZoom
End Sub

Private Sub mnuZoomOut_Click()
  sciMain.ZoomOut
  stb.Panels(4).Text = "Zoom: " & sciMain.GetZoom
End Sub

Private Sub sciMain_SavePointLeft()
  stb.Panels(2).Text = "Modified"
End Sub

Private Sub sciMain_SavePointReached()
  stb.Panels(2).Text = ""
End Sub

Private Sub sciMain_UpdateUI()
  stb.Panels(1).Text = "CurrentLine: " & sciMain.GetCurLine & " Column: " & sciMain.GetColumn & " Lines: " & sciMain.GetLineCount
End Sub
Private Sub mnuCut_Click()
  sciMain.Cut
End Sub

Private Sub mnuNew_Click()
  Dim msgRes As VbMsgBoxResult
  If strFile <> "" And sciMain.Modified = True Then
    msgRes = MsgBox("File: [" & strFile & "]" & vbCrLf & "has been modified.  Do you wish to save?", vbYesNoCancel, "Modified File")
    If msgRes = vbYes Then
      mnuSave_Click
    ElseIf msgRes = vbCancel Then
      Exit Sub
    End If
  End If
   
  sciMain.Text = ""
  sciMain.SetFocus
  
  ' Get rid of the previous filename
  strFile = ""
  
  ' This is a new document so we wan't scintilla be to be unmodified
  sciMain.SetSavePoint
  sciMain.ClearUndoBuffer
  ' For fun let's set the language to Visual Basic since this is after
  ' all a VB class to wrap Scintilla :)
  sciMain.SetHighlighter hlVB
End Sub

Private Sub mnuOpen_Click()
  Dim msgRes As VbMsgBoxResult
  Dim i As Integer
  If strFile <> "" And sciMain.Modified = True Then
    msgRes = MsgBox("File: [" & strFile & "]" & vbCrLf & "has been modified.  Do you wish to save?", vbYesNoCancel, "Modified File")
    If msgRes = vbYes Then
      mnuSave_Click
    ElseIf msgRes = vbCancel Then
      Exit Sub
    End If
  End If
  With cd
    For i = 0 To UBound(Highlighters) - 1
      If Highlighters(i).strFilter <> "" Then .Filter = .Filter & Highlighters(i).strFilter
    Next i
    .ShowOpen
    If cd.filename <> "" Then
      sciMain.LoadFile cd.filename
      SetHighlighter sciMain, SetHighlighterBasedOnExtension(.filename)
    End If
  End With
  strFile = cd.filename
  sciMain.SetFocus
End Sub

Private Sub mnuPaste_Click()
  sciMain.Paste
End Sub

Private Sub mnuRedo_Click()
  sciMain.Redo
End Sub

Private Sub mnuSaveAs_Click()
  With cd
    .Filter = "Visual Basic Files (*.frm, *.frx, *.bas)|*.frm;*.frx;*.bas|All Files (*.*)|*.*"
    .ShowOpen
    If cd.filename <> "" Then
      sciMain.SaveToFile cd.filename
    End If
  End With
  strFile = cd.filename
  sciMain.SetFocus
End Sub

Private Sub mnuUndo_Click()
  sciMain.Undo
End Sub


Public Function AddMenu(sCaption As String, sTag As String, iIndex As Integer) As Integer

  On Error Resume Next
  If iIndex > 0 Then Load mnuHighlighter(iIndex)
  mnuHighlighter(iIndex).Caption = sCaption ' sCaption we got from the "Identify" function on the plugin
  mnuHighlighter(iIndex).Visible = True
  mnuHighlighter(iIndex).Enabled = True
  mnuHighlighter(iIndex).Tag = sTag ' We store the interface to the plugin in here, to later use it on the event of a menu click

End Function

Public Function SetupMenu()
  Dim i As Integer
  For i = 0 To UBound(Highlighters) - 1
    AddMenu Highlighters(i).strName, Highlighters(i).strName, i
  Next i
End Function
