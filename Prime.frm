VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form4"
   ScaleHeight     =   5490
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer, i As Integer, ct As Integer
Private Sub Form_Load()
n = Val(InputBox("Enter the Number"))
ct = o
For i = 2 To n
If n Mod i = 0 Then
ct = ct + 1
End If
Next i
If ct = 1 Then
MsgBox (n & " is prime")
Else
MsgBox (n & " is not prime")
End If
End Sub
