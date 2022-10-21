VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RtnNum As Integer

Private Sub CommandButton1_Click()

'Select Case Frame1
'    Case 1: RtnNum = 1
'    Case 2: RtnNum = 2
'    Case 3: RtnNum = 3
    
'End Select
    


If OptionButton1 = True Then
'MsgBox ("1")
RtnNum = 1
End If

If OptionButton2 = True Then
'MsgBox ("2")
RtnNum = 2
End If

If OptionButton3 = True Then
'MsgBox ("3")
RtnNum = 3
End If

If OptionButton4 = True Then
'MsgBox ("4")
RtnNum = 4
End If

'Unload Me
Me.Hide
End Sub



