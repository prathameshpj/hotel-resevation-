VERSION 5.00
Begin VB.Form formLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HOTEL Login"
   ClientHeight    =   4725
   ClientLeft      =   5025
   ClientTop       =   3930
   ClientWidth     =   5610
   Icon            =   "Formlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Formlogin.frx":08CA
   ScaleHeight     =   4725
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Enter Username and Password"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   5175
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         Height          =   855
         Left            =   3480
         Picture         =   "Formlogin.frx":70B5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ok"
         Height          =   855
         Left            =   2160
         Picture         =   "Formlogin.frx":7CF7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtpassword 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Case sensitive !!"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox comboUser 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   2535
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   720
         Picture         =   "Formlogin.frx":8939
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     Password :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Username 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    Username :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "Formlogin.frx":8F30
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   3750
      Left            =   0
      Picture         =   "Formlogin.frx":D3A8
      Top             =   1200
      Width           =   7500
   End
End
Attribute VB_Name = "formLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub mydb()
Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\hotel.mdb;"
End Sub

Private Sub Command1_Click()
If txtPassword.Text = "" Then
MsgBox "Access Denied", vbCritical, "HOTEL RESERVATION SYSTEM"
Else

Set rs = New ADODB.Recordset
    rs.Open "Select*from users", db, 3, 3
            
    rs.Find ("Username = '" & comboUser.Text & "'")
If rs.Fields("Password") = txtPassword.Text Then
MsgBox "Welcome Manager", vbOKOnly, "HOTEL RESERVATION SYSTEM"
Form2.Show
Form2.Abt.Enabled = True
Form2.Login.Enabled = False
Form2.Exit.Enabled = True
Form2.res.Enabled = True
Form2.reoprt.Enabled = True
    Me.Hide
Else
    MsgBox "Invalid Password!", vbExclamation, "HOTEL RESERVATION SYSTEM"
    
End If
End If
comboUser.Text = ""
txtPassword.Text = ""
End Sub

Private Sub Users()
Set rs = New ADODB.Recordset
    rs.Open "Select*from users", db, 3, 3
Do Until rs.EOF

    comboUser.AddItem rs!Username
    
    rs.MoveNext

Loop
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call mydb
Call Users
End Sub
