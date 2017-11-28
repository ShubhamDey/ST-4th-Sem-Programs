VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form16"
   ScaleHeight     =   4935
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Clear List"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Remove Item"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add New Item"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear List"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove Item"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New Item"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   2595
      ItemData        =   "Listbox.frx":0000
      Left            =   4320
      List            =   "Listbox.frx":0002
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Listbox.frx":0004
      Left            =   240
      List            =   "Listbox.frx":0006
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim i As Integer
Private Sub Command1_Click()
str = List1.Text
For i = 0 To List2.ListCount - 1
    If str = List2.List(i) Then
        MsgBox (str & " already exists ")
        Exit For
    End If
Next i
If i = List2.ListCount Then
    List2.AddItem str
    List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub Command2_Click()
str = List2.Text
For i = 0 To List1.ListCount - 1
    If str = List1.List(i) Then
        MsgBox (str & " already exists ")
        Exit For
    End If
Next i
If i = List1.ListCount Then
    List1.AddItem str
    List2.RemoveItem List2.ListIndex
End If
End Sub

Private Sub Command3_Click()
str = InputBox("Enter a text")
For i = 0 To List1.ListCount - 1
    If str = List1.List(i) Then
        MsgBox (str & " already exists ")
        Exit For
    End If
Next i
If i = List1.ListCount Then
    List1.AddItem str
End If
End Sub

Private Sub Command4_Click()
If List1.ListIndex >= 0 Then
    List1.RemoveItem List1.ListIndex
Else
    MsgBox ("No items selected")
End If
End Sub

Private Sub Command5_Click()
List1.Clear
End Sub

Private Sub Command6_Click()
str = InputBox("Enter a text")
For i = 0 To List2.ListCount - 1
    If str = List2.List(i) Then
        MsgBox (str & " already exists ")
        Exit For
    End If
Next i
If i = List2.ListCount Then
    List2.AddItem str
End If
End Sub

Private Sub Command7_Click()
If List2.ListIndex >= 0 Then
    List2.RemoveItem List2.ListIndex
Else
    MsgBox ("No items selected")
End If
End Sub

Private Sub Command8_Click()
List2.Clear
End Sub
