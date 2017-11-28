VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   3675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7680
   LinkTopic       =   "Form17"
   ScaleHeight     =   3675
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5530
      _Version        =   393217
      TextRTF         =   $"Text Editor.frx":0000
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu New 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save  As"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu Find 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu FindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Replace 
         Caption         =   "Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu TimeDate 
         Caption         =   "Time/Date"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim key As String, i As Integer, rpl As String, j As Integer
Dim str As String

Private Sub Copy_Click()
Clipboard.SetText (RTB1.SelText)
End Sub

Private Sub Cut_Click()
Clipboard.SetText (RTB1.SelText)
RTB1.SelText = ""
End Sub

Private Sub Delete_Click()
RTB1.SelText = ""
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Find_Click()
key = InputBox("Find What?")
i = RTB1.Find(key)
If i < 0 Then
    MsgBox ("Cannot Find " & key)
End If
End Sub

Private Sub FindNext_Click()
i = RTB1.Find(key, i + 1, Len(RTB1.Text))
If i < 0 Then
    MsgBox ("Cannot Find " & key)
End If
End Sub

Private Sub Form_Load()
RTB1.Width = Form17.Width
RTB1.Height = Form17.Height
End Sub

Private Sub Form_Resize()
RTB1.Width = Me.Width
RTB1.Height = Me.Height
End Sub

Private Sub New_Click()
If Len(RTB1.Text) > 0 Then
str = MsgBox("Do you want to save", vbYesNoCancel)
If str = vbYes Then
    CommonDialog1.ShowSave
    RTB1.SaveFile CommonDialog1.FileName
ElseIf str = vbNo Then
    RTB1.Text = ""
End If
Else
RTB1.Text = ""
End If
End Sub

Private Sub Open_Click()
If Len(RTB1.Text) > 0 Then
str = MsgBox("Do you want to save", vbYesNoCancel)
If str = vbYes Then
    CommonDialog1.ShowSave
    RTB1.SaveFile CommonDialog1.FileName
    CommonDialog1.ShowOpen
    RTB1.LoadFile CommonDialog1.FileName
ElseIf str = vbNo Then
    CommonDialog1.ShowOpen
    RTB1.LoadFile CommonDialog1.FileName
End If
Else
    CommonDialog1.ShowOpen
    RTB1.LoadFile CommonDialog1.FileName
End If
End Sub

Private Sub Paste_Click()
RTB1.SelText = Clipboard.GetText
End Sub

Private Sub Replace_Click()
key = InputBox("Find What?")
rpl = InputBox("Replace With ")
i = RTB1.Find(key, 0, Len(RTB1.Text))
If i >= 0 Then
    RTB1.SelText = rpl
Else
    MsgBox ("canot find " & key)
End If
End Sub

Private Sub Save_Click()
If CommonDialog1.FileName <> "" Then
    RTB1.SaveFile CommonDialog1.FileName
Else
    CommonDialog1.ShowSave
    RTB1.SaveFile CommonDialog1.FileName
End If
End Sub

Private Sub SaveAs_Click()
CommonDialog1.ShowSave
RTB1.SaveFile CommonDialog1.FileName
End Sub

Private Sub SelectAll_Click()
RTB1.SelStart = 0
RTB1.SelLength = Len(RTB1.Text)
End Sub

Private Sub TimeDate_Click()
RTB1.SelText = Now
End Sub
