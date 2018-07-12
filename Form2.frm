VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotel Main Menu."
   ClientHeight    =   9540
   ClientLeft      =   450
   ClientTop       =   945
   ClientWidth     =   14295
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   9615
      Left            =   0
      Picture         =   "Form2.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14295
   End
   Begin VB.Menu Login 
      Caption         =   "&Login"
   End
   Begin VB.Menu reg 
      Caption         =   "&Register User"
   End
   Begin VB.Menu res 
      Caption         =   "&Reservation"
   End
   Begin VB.Menu reoprt 
      Caption         =   "&Report"
   End
   Begin VB.Menu Abt 
      Caption         =   "&About"
   End
   Begin VB.Menu Exit 
      Caption         =   "&Logout"
   End
   Begin VB.Menu Exitt 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Abt_Click()
MsgBox "Designed by PRATHAMESH JADHAV,9518907968,CSE WARNANAGAR", vbOKOnly, "HOTEL RESERVATION SYSTEM"
End Sub

Private Sub exit_Click()
If MsgBox("Are you sure to LogOut from this Application?", vbQuestion + vbYesNo, "System") = vbYes Then
Form2.Abt.Enabled = False
Form2.Login.Enabled = True
Form2.Exit.Enabled = False
Form2.res.Enabled = False
Form2.reoprt.Enabled = False
End If
End Sub

Private Sub Exitt_Click()
If MsgBox("Are you sure to End this Application?", vbQuestion + vbYesNo, "System") = vbYes Then
End
End If
End Sub

Private Sub Form_Load()
Form2.Abt.Enabled = True
Form2.Login.Enabled = True
Form2.Exit.Enabled = False
Form2.res.Enabled = False
Form2.reoprt.Enabled = False


End Sub

Private Sub Login_Click()
formLogin.Show
End Sub

Private Sub reg_Click()
Formchange.Show
End Sub

Private Sub reoprt_Click()
Form3.Show
End Sub

Private Sub res_Click()
Form1.Show
End Sub
