VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HOTEL MANAGEMANT SYSTEM. "
   ClientHeight    =   8025
   ClientLeft      =   3120
   ClientTop       =   2115
   ClientWidth     =   10905
   Icon            =   "J.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000C000&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6960
      Width           =   1215
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   855
      Left            =   7920
      ScaleHeight     =   795
      ScaleWidth      =   2355
      TabIndex        =   65
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000C000&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000C000&
      Caption         =   "&Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox txtnumber 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   59
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000000&
      Caption         =   "Search using PHONE NUMBER"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   54
      Top             =   6120
      Width           =   3615
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000C000&
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtsearch 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Text            =   "+91"
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.ComboBox cmbtitle 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "J.frx":08CA
      Left            =   1560
      List            =   "J.frx":08E0
      TabIndex        =   53
      Text            =   "-select-"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C000&
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000000&
      Caption         =   "AVAILABLE SERVICES"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4320
      TabIndex        =   37
      Top             =   3600
      Width           =   3615
      Begin VB.ComboBox cmbtrans 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":0910
         Left            =   1680
         List            =   "J.frx":091A
         TabIndex        =   47
         Text            =   "-select-"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.ComboBox cmbinternet 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":0927
         Left            =   1680
         List            =   "J.frx":0931
         TabIndex        =   46
         Text            =   "-select-"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cmbbar 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":093E
         Left            =   1680
         List            =   "J.frx":0948
         TabIndex        =   45
         Text            =   "-select-"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cmbkitchen 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":0955
         Left            =   1680
         List            =   "J.frx":095F
         TabIndex        =   44
         Text            =   "-select-"
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmblaundry 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":096C
         Left            =   1680
         List            =   "J.frx":0976
         TabIndex        =   43
         Text            =   "-select-"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Transportation Service :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Internet Service :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Service :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Kitchen Service :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Laundry Service :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "ROOM INFORMATION"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4320
      TabIndex        =   23
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtreserved 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   61
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtspend 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   49
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox cmbpaid 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":0983
         Left            =   840
         List            =   "J.frx":098D
         TabIndex        =   36
         Text            =   "-select-"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.ComboBox cmboccupier 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":099A
         Left            =   1800
         List            =   "J.frx":09A1
         TabIndex        =   34
         Text            =   "-select-"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtamount 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cmbcapacity 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":09AA
         Left            =   1440
         List            =   "J.frx":09B4
         TabIndex        =   29
         Text            =   "-select-"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbrnum 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":09CC
         Left            =   1440
         List            =   "J.frx":09EE
         TabIndex        =   28
         Text            =   "-select-"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmbtype 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "J.frx":0A24
         Left            =   1440
         List            =   "J.frx":0A2E
         TabIndex        =   27
         Text            =   "-select-"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Reserved:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Days to Spend :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Occupier :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Number :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.ComboBox cmbid 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "J.frx":0A42
      Left            =   1560
      List            =   "J.frx":0A52
      TabIndex        =   19
      Text            =   "-select-"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ComboBox cmborigin 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "J.frx":0AAD
      Left            =   1560
      List            =   "J.frx":0B20
      TabIndex        =   17
      Text            =   "-select-"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.ComboBox cmbnational 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "J.frx":0C3C
      Left            =   1560
      List            =   "J.frx":0C46
      TabIndex        =   16
      Text            =   "-select-"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtphone 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   18
      Text            =   "+91"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtadd 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtothers 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtfirst 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtlast 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   12
      Top             =   690
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "VEHICLE INFORMATION"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   4095
      Begin VB.TextBox txtmodel 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   58
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtcolor 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtvnum 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtvname 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label26 
         Caption         =   "Vehicle Model :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Vehicle Color :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Vehicle Number:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Vehicle Name :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "UNICS 2017 SUBMITED BY PRATHAMESH JADHAV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   67
      Top             =   7800
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   7680
      Width           =   10935
   End
   Begin VB.Image Image3 
      Height          =   2295
      Left            =   8040
      Picture         =   "J.frx":0C5C
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   8040
      Picture         =   "J.frx":3944
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   8040
      Picture         =   "J.frx":25D91
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Card Number:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   60
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Title :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ID Card Issued :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "State of Origin :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Address :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Names :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
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
    
    rs.Open "Select*from hotel", db, adOpenStatic, adLockOptimistic
End Sub

Private Sub Command1_Click()
    rs.AddNew
            rs("title") = cmbtitle
            rs("firstname") = txtFirst
            rs("LastName") = txtlast
            rs("otherName") = txtothers
            rs("homeAddress") = txtadd
            rs("PhoneNumber") = txtphone
            rs("nationality") = cmbnational
            rs("origin") = cmborigin
            rs("idcard") = cmbid
            rs("idnumber") = txtnumber
            rs("vehiclename") = txtvname
            rs("vehiclenumber") = txtvnum
            rs("vehiclecolor") = txtcolor
            rs("vehiclemodel") = txtmodel
            rs("roomtype") = cmbtype
            rs("roomnumber") = cmbrnum
            rs("capacity") = cmbcapacity
            rs("amount") = txtamount
            rs("occupier") = cmboccupier
            rs("paid") = cmbpaid
            rs("spend") = txtspend
            rs("reserved") = txtreserved
            rs("laundry") = cmblaundry
            rs("kitchen") = cmbkitchen
            rs("bar") = cmbbar
            rs("internet") = cmbinternet
            rs("transportation") = cmbtrans
                    MsgBox "Information Saved"
            rs.Update
    
    MsgBox "Successfully saving data", vbInformation, "AddUser"
    Call clear
    txtlast.SetFocus


    
End Sub

Private Sub clear()
cmbtitle = ""
txtFirst = ""
            txtlast = ""
txtothers = ""
txtadd = ""
            txtphone = ""
            cmbnational = ""
       cmborigin = ""
            cmbid = ""
            txtnumber = ""
            txtvname = ""
txtvnum = ""
              txtcolor = ""
            txtmodel = ""
                cmbtype = ""
              cmbrnum = ""
                   cmbcapacity = ""
                   txtamount = ""
                    cmboccupier = ""
                    cmbpaid = ""
            txtspend = ""
                txtreserved = ""
                cmblaundry = ""
             cmbkitchen = ""
             cmbbar = ""
                   cmbinternet = ""
                     cmbtrans = ""
txtFirst.SetFocus
End Sub

Private Sub Command2_Click()
Call clear
MsgBox "Cleared !"
txtlast.SetFocus
End Sub

Private Sub Command4_Click()
  With rs
    .MoveFirst
            Do Until .EOF
                If txtsearch.Text = !phonenumber Then
                    txtphone.Text = !phonenumber
         cmbtitle = !Title
            txtFirst = !FirstName
           txtlast = !LastName
           txtothers = !otherName
             txtadd = !homeAddress
             txtphone = !phonenumber
            cmbnational = !nationality
            cmborigin = !origin
             cmbid = !idcard
            txtnumber = !idnumber
              txtvname = !vehiclename
               txtvnum = !vehiclenumber
               txtcolor = !vehiclecolor
                txtmodel = !vehiclemodel
                  cmbtype = !roomtype
                  cmbrnum = !roomnumber
                   cmbcapacity = !capacity
                     txtamount = !amount
                      cmboccupier = !occupier
                      cmbpaid = !paid
               txtspend = !spend
                txtreserved = !reserved
                  cmblaundry = !laundry
                   cmbkitchen = !kitchen
                    cmbbar = !bar
                      cmbinternet = !internet
                       cmbtrans = !transportation
                    Exit Do
                    Else
                    
                    .MoveNext
                
                End If
                Loop
                End With
                
End Sub

Private Sub Command5_Click()
Form2.Show
Me.Hide
End Sub

Private Sub Command6_Click()
With rs
    .MoveFirst
            Do Until .EOF
                If txtFirst.Text = !FirstName Then
    !Title = cmbtitle
     !FirstName = txtFirst
            !LastName = txtlast
            !otherName = txtothers
            !homeAddress = txtadd
            !phonenumber = txtphone
            !nationality = cmbnational
            !origin = cmborigin
            !idcard = cmbid
            !idnumber = txtnumber
            !vehiclename = txtvname
            !vehiclenumber = txtvnum
            !vehiclecolor = txtcolor
            !vehiclemodel = txtmodel
            !roomtype = cmbtype
            !roomnumber = cmbrnum
            !capacity = cmbcapacity
            !amount = txtamount
            !occupier = cmboccupier
            !paid = cmbpaid
            !spend = txtspend
            !reserved = txtreserved
            !laundry = cmblaundry
            !kitchen = cmbkitchen
            !bar = cmbbar
            !internet = cmbinternet
            !transportation = cmbtrans
    .Update
            
            Exit Do
        Else
                    
        .MoveNext
                
        End If
        Loop
End With

MsgBox "Successfully Updating data", vbInformation, "HOTEL RESERVATION SYSTEM"
Call clear
txtFirst.SetFocus
End Sub

Private Sub Command7_Click()
If rs.RecordCount > 0 Then
If MsgBox("Are You Sure you want to Delete ?", vbExclamation + vbOKCancel, "RESERVATION") = vbOK Then
rs.Delete
rs.MoveNext
MsgBox "Record Deleted", vbInformation, "HOTEL RESERVATION SYSTEM"
Call clear
txtlast.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
Call mydb
Call MyRs
End Sub
