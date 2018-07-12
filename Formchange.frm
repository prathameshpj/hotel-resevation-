VERSION 5.00
Begin VB.Form Formchange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create and Change Username & Password"
   ClientHeight    =   6060
   ClientLeft      =   5235
   ClientTop       =   2700
   ClientWidth     =   5730
   Icon            =   "Formchange.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin Project1.CandyButton CandyButton4 
      Height          =   1215
      Left            =   4200
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton CandyButton3 
      Height          =   1215
      Left            =   3000
      TabIndex        =   14
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Delete"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton CandyButton2 
      Height          =   1215
      Left            =   1680
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Edit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton CandyButton1 
      Height          =   1215
      Left            =   480
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Save"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.ComboBox ComboUtype 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   11
      Text            =   "--Select--"
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtConfirm 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "Case sensitive !!"
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      ToolTipText     =   "Case sensitive !!"
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtFirst 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      ToolTipText     =   "Case sensitive !!"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txtlast 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      ToolTipText     =   "Case sensitive !!"
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtuser 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      ToolTipText     =   "Case sensitive !!"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "Formchange.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "       Confirm :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     Password :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    User Type :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   First Name :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    Last Name :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   User Name :"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   4695
      Left            =   0
      Picture         =   "Formchange.frx":4D42
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7500
   End
End
Attribute VB_Name = "Formchange"
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
Private Sub MyRs()
Set rs = New ADODB.Recordset
    
    rs.Open "Select*from Users", db, adOpenStatic, adLockOptimistic
End Sub

Private Sub CandyButton1_Click()
If txtPassword.Text = txtConfirm.Text Then
With rs
    .AddNew
        !Username = txtuser.Text
        !FirstName = txtFirst.Text
        !LastName = txtlast.Text
        !Password = txtPassword.Text
        !Confirm = txtConfirm.Text
        !Usertype = ComboUtype.Text
    .Update
End With
    MsgBox "Successfully saving data", vbInformation, "AddUser"
    Call clear
    txtuser.SetFocus
Else
    MsgBox "Password do not match", vbInformation, "AddUser"
     txtPassword.Text = ""
     txtConfirm.Text = ""
     txtPassword.SetFocus
End If
End Sub

Private Sub CandyButton2_Click()
With rs
    .MoveFirst
            Do Until .EOF
                If txtuser.Text = !Username Then
    
    !FirstName = txtFirst.Text
    !LastName = txtlast.Text
    !Usertype = ComboUtype.Text
    !Username = txtuser.Text
    !Password = txtPassword.Text
    !Confirm = txtConfirm.Text
    .Update
            
            Exit Do
        Else
                    
        .MoveNext
                
        End If
        Loop
End With

MsgBox "Successfully Editing data", vbInformation, "HOTEL RESERVATION SYSTEM"
Call clear
txtuser.SetFocus
End Sub

Private Sub CandyButton3_Click()
If rs.RecordCount > 0 Then
If MsgBox("Are You Sure you want to Delete ?", vbExclamation + vbOKCancel, "User") = vbOK Then
rs.Delete
rs.MoveNext
MsgBox "Record Deleted", vbInformation, "HOTEL RESERVATION SYSTEM"
Call clear
txtuser.SetFocus
End If
End If
End Sub



Private Sub cmdSave_Click()
If txtPassword.Text = txtConfirm.Text Then
With rs
    .AddNew
        !Username = txtuser.Text
        !FirstName = txtFirst.Text
        !LastName = txtlast.Text
        !Password = txtPassword.Text
        !Confirm = txtConfirm.Text
        !Usertype = ComboUtype.Text
    .Update
End With
    MsgBox "Successfully saving data", vbInformation, "AddUser"
    Call clear
    txtuser.SetFocus
Else
    MsgBox "Password do not match", vbInformation, "AddUser"
     txtPassword.Text = ""
     txtConfirm.Text = ""
     txtPassword.SetFocus
End If
    
End Sub
Private Sub clear()
txtFirst.Text = ""
txtlast.Text = ""
txtuser.Text = ""
txtPassword.Text = ""
txtConfirm.Text = ""
ComboUtype.Text = ""
txtuser.SetFocus
End Sub

Private Sub CandyButton4_Click()
Me.Hide

End Sub

Private Sub Form_Load()
With ComboUtype
    .AddItem "Manager"
    .AddItem "Receptionist"
End With

Call mydb
Call MyRs
End Sub

Private Sub txtuser_Change()
With rs
    .MoveFirst
            Do Until .EOF
                If txtuser.Text = !Username Then
                    txtFirst.Text = !FirstName
                    txtlast.Text = !LastName
                    txtPassword.Text = !Password
                    txtConfirm.Text = !Confirm
                    ComboUtype.Text = !Usertype
                    Exit Do
                    Else
                    
                    .MoveNext
                
                End If
                Loop
                End With
End Sub
