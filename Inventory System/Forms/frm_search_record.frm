VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_search_record 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer Records"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "frm_search_record.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   7155
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1320
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customer Sales Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   840
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customer System Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Customer Records"
      TabPicture(0)   =   "frm_search_record.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(9)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(12)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LaVolpeButton2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "LaVolpeButton1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmd(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text1(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(5)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(11)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   2160
         TabIndex        =   5
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   2160
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   11
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   10
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   9
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   8
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   2160
         TabIndex        =   12
         Top             =   4440
         Width           =   2895
      End
      Begin LVbuttons.LaVolpeButton cmd 
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   13
         Top             =   4440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Modify"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   14737632
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_search_record.frx":0E5E
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   375
         Left            =   5520
         TabIndex        =   29
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Add New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_search_record.frx":0E7A
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   375
         Left            =   5520
         TabIndex        =   30
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frm_search_record.frx":0E96
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   22
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "STD Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   19
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Width           =   1815
      End
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   14
      Top             =   6240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "System Details (F3)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14737632
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_search_record.frx":0EB2
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   15
      Top             =   6240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Sales Book Entry (F2)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14737632
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_search_record.frx":0ECE
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5160
      Picture         =   "frm_search_record.frx":0EEA
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3000
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Customer Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frm_search_record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cust_rs As New ADODB.Recordset
Dim C_COUNT As New ADODB.Recordset
Dim cust_no As New ADODB.Recordset
Dim add_cust_rs As New ADODB.Recordset



Private Sub cmd_Click(Index As Integer)
If Index = 0 Then

        With CrystalReport1
            .DataFiles(0) = App.Path & "\Master_Database.mdb"
            .ReportFileName = App.Path & "\Report\SYSTEM_REPORT.rpt"
            .SelectionFormula = "{Customer_master.cutomer_name} = '" & Combo1.Text & "'"
            .username = "Admin"
            .Password = "1010101010" & Chr(10) & "1010101010"
            .Action = 1
            .PageZoom (100)
        End With


ElseIf Index = 2 Then
 
        Combo1.Enabled = False
        Dim cust_rs As New ADODB.Recordset
        If cmd(2).Caption = "&Modify" Then
                       SendKeys "{TAB}"
                       SendKeys "{TAB}"
                       SendKeys "{END}"
                cmd(0).Enabled = False
                cmd(1).Enabled = False
                

                LaVolpeButton1.Enabled = False
                
                For i = 1 To Text1.Count - 1
                    Text1(i).Enabled = True
                Next
                    cmd(2).Caption = "&Save"
        Else
                
               ' On Error GoTo a11:
                
                
                For i = 1 To Text1.Count - 1
                    Text1(i).Enabled = False
                Next
                cust_rs.Open "select * from Customer_master where cutomer_id='" & Text1(0).Text & "'", db, adOpenDynamic, adLockOptimistic
                For i = 1 To Text1.Count - 1
                    If Len(Text1(i).Text) > 0 Then
                        cust_rs.Fields(i).Value = Text1(i).Text
                    End If
                Next
                
                cust_rs.Update
                
                cmd(0).Enabled = True
                cmd(1).Enabled = True
                Combo1.Enabled = True
                db.Execute "UPDATE AMT_UNPAID_REMIND SET PARTY_NAME='" & Text1(1).Text & "' WHERE PARTY_NAME='" & Combo1.Text & "' AND TRAN_TYPE='SALES'"
                LaVolpeButton1.Enabled = True
                cust_rs.Close
                cust_rs.Open "SELECT * FROM Customer_master", db, adOpenDynamic, adLockOptimistic

                Combo1.Clear
                cust_rs.Requery
                
                While cust_rs.EOF <> True
                        Combo1.AddItem cust_rs.Fields(1).Value
                        cust_rs.MoveNext
                Wend
                
                Combo1.Text = Text1(1).Text
                cmd(2).Caption = "&Modify"
                
                SendKeys "{TAB}"
                SendKeys "{TAB}"
                SendKeys "{TAB}"
                SendKeys "{TAB}"
                SendKeys "{TAB}"

                Exit Sub
a11:
                cust_rs.CancelUpdate
                
                MsgBox "Duplicate Customer Name OR NULL Entry Not Allowed ...", vbInformation
                


                 

        End If
        
ElseIf Index = 1 Then
    With CrystalReport2
            .DataFiles(0) = App.Path & "\Master_Database.MDB"
            .ReportFileName = App.Path & "\Report\rpt_salesbook.rpt"
            .SelectionFormula = "{Sales_master.Party_name} = '" & Combo1.Text & "'"
            .username = "Admin"
            .Password = "1010101010" & Chr(10) & "1010101010"
            .Action = 1
            .PageZoom (100)
    End With
End If
End Sub

Private Sub Combo1_Click()
For i = 0 To Text1.Count - 1
    Text1(i).Text = Clear
Next


If Len(Combo1.Text) > 0 Then
    For i = 0 To cmd.Count - 1
        cmd(i).Enabled = True
    Next
    Dim R As New ADODB.Recordset
    R.Open "select * from Customer_master where cutomer_name='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
    For i = 0 To Text1.Count - 1
        If Len(R.Fields(i).Value) > 0 Then
            Text1(i).Text = R.Fields(i).Value
        End If
    Next
    R.Close
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 114 Then
    If cmd(0).Enabled = True Then
        Call cmd_Click(0)
    End If
ElseIf KeyCode = 27 Then
    If cmd(2).Enabled = False And Len(Combo1.Text) = 0 Then
        Unload Me
        Exit Sub
    End If
    
    If cmd(0).Enabled = True Then
        Unload Me
    End If
ElseIf KeyCode = 113 Then
        If cmd(0).Enabled = True Then
            Call cmd_Click(1)
        End If

End If
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
KeyPreview = True

For i = 0 To cmd.Count - 1
    cmd(i).Enabled = False
Next



cust_rs.Open "SELECT * FROM Customer_master", db, adOpenDynamic, adLockOptimistic
C_COUNT.Open "SELECT COUNT(*) FROM Customer_master", db, adOpenDynamic, adLockOptimistic
If C_COUNT.Fields(0).Value = 0 Then
    MsgBox "No Customer Record Found ...", vbInformation, "No Customer entry Found ..."
    Unload Me
    Exit Sub
End If

While cust_rs.EOF <> True
    Combo1.AddItem cust_rs.Fields(1).Value
    cust_rs.MoveNext
Wend

For i = 0 To Text1.Count - 1
    Text1(i).Enabled = False
Next


End Sub

Private Sub Form_Unload(Cancel As Integer)
cust_rs.Close
C_COUNT.Close
On Error Resume Next
add_cust_rs.CancelUpdate
cust_no.Close
        'Dim f As New FileSystemObject
        'f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True
       
End Sub

Private Sub LaVolpeButton1_Click()

If LaVolpeButton1.Caption = "&Add New" Then
                
                LaVolpeButton2.Visible = True
                
                cust_no.Open "select * from SORTED_CUST_NO", db, adOpenDynamic, adLockOptimistic
                
                If add_cust_rs.State = adStateOpen Then
                    add_cust_rs.Close
                End If
                
                add_cust_rs.Open "select * from Customer_master", db, adOpenKeyset, adLockOptimistic
                
                add_cust_rs.AddNew
                If cust_no.EOF <> True Then
                
                    cust_no.MoveLast
                    Dim LAST_ID As String
                    LAST_ID = Mid(cust_no.Fields(0).Value, 2, Len(cust_no.Fields(0).Value))
                    LAST_ID = VAL(LAST_ID) + 1
                    
                    If Len(LAST_ID) = 1 Then
                        Text1(0).Text = "C00000000" & LAST_ID
                    ElseIf Len(LAST_ID) = 2 Then
                        Text1(0).Text = "C0000000" & LAST_ID
                    ElseIf Len(LAST_ID) = 3 Then
                        Text1(0).Text = "C000000" & LAST_ID
                    ElseIf Len(LAST_ID) = 4 Then
                        Text1(0).Text = "C00000" & LAST_ID
                    ElseIf Len(LAST_ID) = 5 Then
                        Text1(0).Text = "C0000" & LAST_ID
                    ElseIf Len(LAST_ID) = 6 Then
                        Text1(0).Text = "C000" & LAST_ID
                    ElseIf Len(LAST_ID) = 7 Then
                        Text1(0).Text = "C00" & LAST_ID
                    ElseIf Len(LAST_ID) = 8 Then
                        Text1(0).Text = "C0" & LAST_ID
                    ElseIf Len(LAST_ID) = 9 Then
                        Text1(0).Text = "C" & LAST_ID
                    End If
                Else
                    Text1(0).Text = "C000000001"
                    
                End If
                
                cust_no.Close
                Combo1.Clear
                Combo1.Enabled = False
                
                For i = 1 To Text1.Count - 1
                    Text1(i).Text = Clear
                    Text1(i).Enabled = True
                Next
                
                LaVolpeButton1.Caption = "&Save"
                cmd(2).Enabled = False
                cmd(1).Enabled = False
                cmd(0).Enabled = False
                SendKeys "{TAB}"
                SendKeys "{TAB}"
                
Else
                
                
                
                
                
                
                
                For i = 0 To Text1.Count - 1
                            If Len(Text1(i).Text) > 0 Then
                                add_cust_rs.Fields(i).Value = Text1(i).Text
                            End If
                
                Next
                
                On Error GoTo A1:
                add_cust_rs.Update
                add_cust_rs.Close
                GoTo a2:
A1:
                MsgBox "Duplicate Customer Name OR NULL Entry Not Allowed ...", vbInformation
                SendKeys "{TAB}"
                SendKeys "{TAB}"
                Exit Sub
                
a2:
                
                
                cmd(0).Enabled = True
                cmd(1).Enabled = True
                cmd(2).Enabled = True
                LaVolpeButton1.Caption = "&Add New"
                

                
                
                
                
                LaVolpeButton2.Enabled = False
                
                
                cust_rs.Requery
                While cust_rs.EOF <> True
                    Combo1.AddItem cust_rs.Fields(1).Value
                    cust_rs.MoveNext
                Wend
                
                Combo1.Text = Text1(1).Text
                Combo1.Enabled = True
                
                For i = 0 To Text1.Count - 1
                        Text1(i).Enabled = False
                Next
                MsgBox "Customer Added Successfully ...", vbInformation
                SendKeys "{TAB}"
                
                
End If










End Sub

Private Sub LaVolpeButton2_Click()

add_cust_rs.CancelUpdate
add_cust_rs.Close
Combo1.Clear
For i = 0 To Text1.Count - 1
     Text1(i).Text = Clear
     Text1(i).Enabled = False
Next

    cmd(0).Enabled = True
    cmd(1).Enabled = True
    cmd(2).Enabled = True
    LaVolpeButton1.Caption = "&Add New"
    LaVolpeButton2.Visible = False
    Combo1.Enabled = True
    cust_rs.Requery
                While cust_rs.EOF <> True
                    Combo1.AddItem cust_rs.Fields(1).Value
                    cust_rs.MoveNext
                Wend
                
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    
    

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    SendKeys "{TAB}"
End If
End Sub

