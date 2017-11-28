VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2175
      Left            =   480
      Max             =   500
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Double
Dim f As Double

Private Sub VScroll1_Change()
f = VScroll1.Value
Label1.Caption = f & " Farenheit"
c = (f - 32) * (5 / 9)
Label2.Caption = c & " Celcius"
End Sub
