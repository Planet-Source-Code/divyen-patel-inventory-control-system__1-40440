VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00808080&
   Caption         =   "Customer Support Division and Inventory System ( Version 1.0.0 )"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10605
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0E42
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr_cust_details 
      Left            =   120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customer Contact and Detail Report"
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
   Begin Crystal.CrystalReport monthly_profit_report 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Profit Report"
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5955
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
            Text            =   "Developed By : Divyen k Patel ( divyen@msn.com )"
            TextSave        =   "Developed By : Divyen k Patel ( divyen@msn.com )"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/3/2002"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:47 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   4  'Align Right
      Height          =   5955
      Left            =   9615
      TabIndex        =   6
      Top             =   0
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   10504
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   7200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":BB22
               Key             =   "cust"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":C974
               Key             =   "sec"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":D7C6
               Key             =   "report"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":E4A0
               Key             =   "party"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":F2F2
               Key             =   "profit"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":10574
               Key             =   "reminder"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":10E4E
               Key             =   "img7"
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton6 
         Height          =   615
         Left            =   0
         TabIndex        =   15
         Top             =   5040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "MDIForm1.frx":11728
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "7"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton5 
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "MDIForm1.frx":11744
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "5"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton4 
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "MDIForm1.frx":11760
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "3"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton3 
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "MDIForm1.frx":1177C
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "4"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   4080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "MDIForm1.frx":11798
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "2"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "MDIForm1.frx":117B4
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Reminder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   17
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Security"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   16
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Profit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   14
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   13
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   6000
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Logged In Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   6600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "12:00 AM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   7
         Top             =   7080
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   840
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":117D0
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":124AA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":132FC
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1414E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14A28
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15302
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15BDC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":165A6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":16E80
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1719A
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17A74
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1834E
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18C28
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18F42
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1981C
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A0F6
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A9D0
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B2AA
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BB84
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C45E
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CD38
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D612
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DEEC
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E7C6
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F0A0
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F97A
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20254
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20B2E
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21408
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21CE2
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22598
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22E72
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":232C4
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23716
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25EC8
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_sys_task 
      Caption         =   "System Task"
      Begin VB.Menu mnuSbar8 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Sys Task|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_log_out 
         Caption         =   "{IMG:12}Log Out ..."
      End
      Begin VB.Menu mnu_sep_logout 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "{IMG:14}Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu_int_admin 
      Caption         =   "Inventory"
      Begin VB.Menu mnuSbar1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Inventory Management|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_purchase 
         Caption         =   "{IMG:4}Purchase Data Entry"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_sales 
         Caption         =   "{IMG:4}Sales Data Entry"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_spe1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_internet_connection 
         Caption         =   "{IMG:4}Internet Connection Sale Entry"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_remove_cust_ex_check 
         Caption         =   "{IMG:17}Remove Customer From Expiry Check"
      End
      Begin VB.Menu mnu_ic_ex_report 
         Caption         =   "{IMG:1}Internet Connection Expiry Report"
      End
      Begin VB.Menu mnu_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_int_sales_report 
         Caption         =   "{IMG:1}Internet Connection Sales Report"
      End
      Begin VB.Menu mnu_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_purchase_book 
         Caption         =   "{IMG:1}Purchase Book Report"
      End
      Begin VB.Menu mnu_sales_book 
         Caption         =   "{IMG:1}Sales Book Report"
      End
      Begin VB.Menu mnu_int_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_stock 
         Caption         =   "{IMG:1}Current Stock Report"
      End
   End
   Begin VB.Menu mnu_cash_mgt 
      Caption         =   "Cash Management"
      Begin VB.Menu mnuSbar2 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Cash Management|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_income 
         Caption         =   "{IMG:4}Miscellaneous Income"
      End
      Begin VB.Menu mnu_expense 
         Caption         =   "{IMG:4}Miscellaneous Expense"
      End
      Begin VB.Menu mnu_sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_payment_received 
         Caption         =   "{IMG:4}Sales Payment Received"
      End
      Begin VB.Menu mnu_pur_payment 
         Caption         =   "{IMG:4}Purchase Payment Given"
      End
      Begin VB.Menu mnu_sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_payment_due_report 
         Caption         =   "{IMG:33}Sales Payment Due Report"
      End
      Begin VB.Menu mnu_pur_due_report 
         Caption         =   "{IMG:33}Purchase Payment Due Report"
      End
      Begin VB.Menu sep56 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_prof_loss_report 
         Caption         =   "{IMG:35}Profit  - Loss Report"
      End
   End
   Begin VB.Menu mnu_modify_entries 
      Caption         =   "Modify"
      Begin VB.Menu mnuSbar4 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Modify Entries|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_item_type_name 
         Caption         =   "{IMG:16}Item Type and Item Name"
      End
      Begin VB.Menu mod_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_pur_names 
         Caption         =   "{IMG:16}Purchase Party names"
      End
      Begin VB.Menu mod_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_update_purchase_entry 
         Caption         =   "{IMG:16}Purchase Book Entry"
      End
      Begin VB.Menu mnu_update_sales_entry 
         Caption         =   "{IMG:16}Sales Book Entry"
      End
      Begin VB.Menu mod_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_modify_internet_connection_sale 
         Caption         =   "{IMG:16}Internet Connection Sale Entry"
      End
   End
   Begin VB.Menu MNU_VIEW 
      Caption         =   "View"
      Begin VB.Menu mnuSbar5 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:View|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_purchase_bills 
         Caption         =   "{IMG:21}Purchase Entries"
      End
      Begin VB.Menu mnu_sales_bill 
         Caption         =   "{IMG:21}Sales Entries"
      End
      Begin VB.Menu MNU_VIEW_SEP1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnu_customer 
      Caption         =   "Tools"
      Begin VB.Menu mnuSbar3 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Tools|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_customer_record 
         Caption         =   "{IMG:2}Search Customer Records"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu_sep_tools 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_log_details 
         Caption         =   "{IMG:34}View Log Details"
      End
      Begin VB.Menu mnu_tt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sys_clean 
         Caption         =   "{IMG:11}Clear Log Details"
      End
      Begin VB.Menu mnu_toolsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_backup 
         Caption         =   "{IMG:9}Backup Database"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnu_restore 
         Caption         =   "{IMG:13}Restore Database"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_toolsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sec_pol 
         Caption         =   "{IMG:3}Security Policy"
      End
      Begin VB.Menu mnu_sep_tool 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_remind 
         Caption         =   "{IMG:31}Reminder Of the Day"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_tool7 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sys_info 
         Caption         =   "System Information"
      End
      Begin VB.Menu msto 
         Caption         =   "-Microsoft Tools"
      End
      Begin VB.Menu mnu_calc 
         Caption         =   "{IMG:10}Calculator"
      End
      Begin VB.Menu mnu_notepad 
         Caption         =   "{IMG:8}Notepad"
      End
      Begin VB.Menu mnu_we 
         Caption         =   "{IMG:6}Windows Explorer"
      End
      Begin VB.Menu mnu_osk 
         Caption         =   "{IMG:24}On-Screen Keyboard"
      End
      Begin VB.Menu sep_games 
         Caption         =   "-Games"
      End
      Begin VB.Menu mnu_sol 
         Caption         =   "{IMG:25}Solitaire"
      End
      Begin VB.Menu mnu_mine 
         Caption         =   "{IMG:26}Minesweeper"
      End
      Begin VB.Menu mnu_sep_tool3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_send_msg 
         Caption         =   "{IMG:30}Send Message to Network Computer"
      End
   End
   Begin VB.Menu mnu_windows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuSbar7 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Windows|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_cascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnu_tile_horizon 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnu_t_vertical 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnu_ar_icon 
         Caption         =   "Arrange Icon"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "Help"
      Begin VB.Menu mnuSbar10 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Help|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnu_keyboard_help 
         Caption         =   "{IMG:18}Keyboard Help"
      End
      Begin VB.Menu mnu_help_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_showst_tips 
         Caption         =   "{IMG:32}Show Start up Tips"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "{IMG:19}About"
      End
   End
   Begin VB.Menu un_report 
      Caption         =   "Reports"
      Begin VB.Menu st_top 
         Caption         =   "-Browse MIS Reports"
      End
      Begin VB.Menu un_cr_report 
         Caption         =   "{IMG:1}Current Stock Report"
      End
      Begin VB.Menu unsep 
         Caption         =   "-"
      End
      Begin VB.Menu un_pr_book 
         Caption         =   "{IMG:1}Purchase Book Report"
      End
      Begin VB.Menu un_sales_book 
         Caption         =   "{IMG:1}Sales Book Report"
      End
      Begin VB.Menu un_sep23 
         Caption         =   "-"
      End
      Begin VB.Menu in_ic_expiry_report 
         Caption         =   "{IMG:1}Internet Connection Expiry Report"
      End
      Begin VB.Menu un_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu rp_menu_int_sales 
         Caption         =   "{IMG:1}Internet Connection Sales Report"
      End
      Begin VB.Menu SEPSEPSEP 
         Caption         =   "-"
      End
      Begin VB.Menu un_sales_unpaid 
         Caption         =   "{IMG:33}Sales Payment Due Report"
      End
      Begin VB.Menu un_pur_un_report 
         Caption         =   "{IMG:33}Purchase Payment Due Report"
      End
      Begin VB.Menu un_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cust_details 
         Caption         =   "{IMG:1}Customer Detail Report"
      End
      Begin VB.Menu mnu_rep_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu un_log_report 
         Caption         =   "{IMG:34}View Log Details"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cur_user As String
Dim ST As Boolean
Public cur_company_name As String
Dim comp_saved As Boolean
Public LOGOUT_CLICKED As Boolean


Private Sub in_ic_expiry_report_Click()
mnu_ic_ex_report_Click
End Sub

Private Sub LaVolpeButton1_Click()
Call mnu_customer_record_Click
End Sub

Private Sub LaVolpeButton2_Click()
Call mnu_sec_pol_Click
End Sub

Private Sub LaVolpeButton3_Click()
Call mnu_pur_names_Click
End Sub

Private Sub LaVolpeButton4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    PopupMenu un_report
End If
End Sub

Private Sub LaVolpeButton5_Click()
Call mnu_prof_loss_report_Click
End Sub

Private Sub LaVolpeButton6_Click()
mnu_remind_Click
End Sub

Private Sub MDIForm_Activate()
comp_saved = False
LOGOUT_CLICKED = False
If cur_user = "OPERATOR" Then
    mnu_log_details.Enabled = False
    LaVolpeButton2.Enabled = False
    LaVolpeButton3.Enabled = False
    
    mnu_modify_entries.Enabled = False
    mnu_sec_pol.Enabled = False
    
    
    Dim R As New ADODB.Recordset
    R.Open "SELECT * FROM SECURITY_POLICY", db, adOpenKeyset, adLockOptimistic
    
    
    If R.Fields(0).Value = True Then
                mnu_purchase.Enabled = True
                mnu_sales.Enabled = True
                mnu_internet_connection.Enabled = True
                mnu_remove_cust_ex_check.Enabled = True
    Else
                mnu_purchase.Enabled = False
                mnu_sales.Enabled = False
                mnu_internet_connection.Enabled = False
                mnu_remove_cust_ex_check.Enabled = False
                
    End If
    
    If R.Fields(1).Value = True Then
            mnu_income.Enabled = True
            mnu_expense.Enabled = True
            mnu_payment_received.Enabled = True
            mnu_pur_payment.Enabled = True
    Else
            mnu_income.Enabled = False
            mnu_expense.Enabled = False
            mnu_payment_received.Enabled = False
            mnu_pur_payment.Enabled = False
    End If
    
    
    If R.Fields(2).Value = True Then
        mnu_backup.Enabled = True
    Else
        mnu_backup.Enabled = False
    End If
    
    If R.Fields(3).Value = True Then
        mnu_restore.Enabled = True
    Else
        mnu_restore.Enabled = False
    End If
    
    If R.Fields(4).Value = True Then
        mnu_sys_clean.Enabled = True
    Else
        mnu_sys_clean.Enabled = False
    End If
End If


End Sub

Private Sub MDIForm_Load()
SetMenus hwnd, SmallImages
MDIForm1.Toolbar2.Width = 990

If Len(cur_company_name) > 0 Then
            Dim clear_temp As New ADODB.Recordset
            clear_temp.Open "select * from SYS_CURRENT_INVOICE", db, adOpenKeyset, adLockOptimistic
            While clear_temp.EOF <> True
                clear_temp.Delete
                clear_temp.MoveNext
            Wend
            
            clear_temp.Close
            clear_temp.Open "select * from SYS_CURRENT_SALES_ITEMS", db, adOpenKeyset, adLockOptimistic
            While clear_temp.EOF <> True
                clear_temp.Delete
                clear_temp.MoveNext
            Wend
            clear_temp.Close
            
End If


ST = False
comp_saved = False
LOGOUT_CLICKED = False
                
                Dim R As New ADODB.Recordset
                R.Open "SELECT * FROM CHECK_SECURITY", db, adOpenKeyset, adLockOptimistic
                If R.Fields(0).Value = False Then
                    MDIForm1.mnu_log_out.Visible = False
                Else
                        MDIForm1.mnu_log_out.Visible = True
                End If
                
FRM_TASK.Show
                
Dim CHK_TIP_ST As New ADODB.Recordset
CHK_TIP_ST.Open "SELECT * FROM tip_status", db, adOpenKeyset, adLockOptimistic

If CHK_TIP_ST.Fields(0).Value = True Then
    FRM_TIPS.Show
End If
CHK_TIP_ST.Close

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ReleaseMenus hwnd
    
   If cur_user <> "PASS_DISABLE" Then
        If ST = False Then
                Dim d As New ADODB.Recordset
                On Error Resume Next
                d.Open "SELECT * FROM LOG_DETAIL WHERE LOG_OUT_TIME='CURRENT'", db, adOpenKeyset, adLockOptimistic
                d.Fields(2).Value = Now
                
                d.Update
                d.Close

        Else
            UPDATE_LOG
        End If
    End If
     
    On Error Resume Next
        db.Close
    Exit Sub
  
End Sub

Private Sub mnu_about_Click()
frmAbout.Show
End Sub

Private Sub mnu_ad_search_Click()
FRM_ADVANCE_SEARCH.Show
End Sub

Private Sub mnu_ar_icon_Click()
MDIForm1.Arrange vbArrangeIcons
End Sub

Private Sub mnu_backup_Click()
FRM_BACK_UP.Show
End Sub

Private Sub mnu_calc_Click()
On Error Resume Next
Shell ("calc"), vbMinimizedFocus
Exit Sub
End Sub

Private Sub mnu_cascade_Click()
    MDIForm1.Arrange vbCascade
End Sub

Private Sub mnu_cust_details_Click()

With cr_cust_details
        .DataFiles(0) = App.Path & "\Master_Database.MDB"
        .ReportFileName = App.Path & "\Report\rpt_cust_details.rpt"
        .username = "Admin"
        .Password = "1010101010" & Chr(10) & "1010101010"
        .Action = 1
        .PageZoom (100)
End With
End Sub

Private Sub mnu_customer_record_Click()
On Error Resume Next
frm_search_record.Show
Exit Sub
End Sub

Private Sub mnu_exit_Click()
comp_saved = True
comp_saved = True
UPDATE_LOG
ST = True
Unload Me
End
End Sub

Private Sub mnu_expense_Click()
frm_expense.Show
End Sub


Private Sub mnu_ic_ex_report_Click()
frm_internet_expiry_report.Show
End Sub


Private Sub mnu_income_Click()
frm_income.Show
End Sub

Private Sub mnu_int_sales_report_Click()
With Form1.cr_ct_int_sales
    .DataFiles(0) = App.Path & "\Master_Database.MDB"
    .ReportFileName = App.Path & "\Report\RPT_INTERNET_CD_SALES.rpt"
    .username = "Admin"
    .Password = "1010101010" & Chr(10) & "1010101010"
    .Action = 1
    .PageZoom (100)
End With
End Sub

Private Sub mnu_internet_connection_Click()
INTER_NET_CONNECTIONS.Show
End Sub

Private Sub mnu_item_type_name_Click()
On Error Resume Next
    FRM_CHANGE_ITEM_MASTER.Show
Exit Sub
End Sub

Private Sub mnu_keyboard_help_Click()
FRM_KEYBOARD_HELP.Show
End Sub

Private Sub mnu_log_details_Click()
frm_log_details.Show
End Sub

Private Sub mnu_log_out_Click()
    If cur_user <> "PASS_DISABLE" Then
            Dim R As New ADODB.Recordset
            R.Open "select * from CHECK_SECURITY", db, adOpenKeyset, adLockOptimistic
            If R.Fields(0).Value = "True" Then
                UPDATE_LOG
                comp_saved = True
                ST = True
                LOGOUT_CLICKED = True
                Unload Me
                frm_user_pass.Show
                Exit Sub
            Else
                MsgBox "Security is Disabled ..." & vbCrLf & "You can not logout ,to Close an Application Click on Exit ...", vbInformation, "Security Checked is Disabled..."
            End If
    End If
End Sub

Private Sub mnu_mine_Click()
On Error Resume Next
    Shell "winmine", vbNormalFocus
Exit Sub
End Sub

Private Sub mnu_modify_internet_connection_sale_Click()
On Error Resume Next
        FRM_MODIFY_INTERNET.Show
Exit Sub
End Sub

Private Sub mnu_notepad_Click()
On Error Resume Next
Shell ("notepad"), vbMaximizedFocus
Exit Sub

End Sub


Private Sub mnu_osk_Click()
On Error Resume Next
    Shell "osk", vbNormalFocus
Exit Sub
End Sub



Private Sub mnu_payment_due_report_Click()
 With Form1.unpaid_sales
        .DataFiles(0) = App.Path & "\Master_Database.mdb"
        .ReportFileName = App.Path & "\Report\rpt_unpaid_report.rpt"
        .SelectionFormula = "{QRY_UNPAID_REPORT.TRAN_TYPE} = 'SALES'"
        .username = "Admin"
        .Password = "1010101010" & Chr(10) & "1010101010"
        .Action = 1
        .PageZoom (100)
End With
End Sub

Private Sub mnu_payment_received_Click()
On Error Resume Next
sales_unpaid.Show
Exit Sub
End Sub

Private Sub mnu_prof_loss_report_Click()
Dim r_p As New ADODB.Recordset
Dim r_e As New ADODB.Recordset
Dim r_i As New ADODB.Recordset

r_p.Open "SELECT * FROM DATE_PROFIT WHERE YEAR(DAT) = '2002'", db, adOpenKeyset, adLockOptimistic
r_e.Open "SELECT * FROM EXPENSE WHERE YEAR(DAT) = '2002'", db, adOpenKeyset, adLockOptimistic
r_i.Open "SELECT * FROM INCOME WHERE YEAR(DAT) = '2002'", db, adOpenKeyset, adLockOptimistic

Dim P(1 To 12) As Double
Dim E(1 To 12) As Double
Dim i(1 To 12) As Double

While r_p.EOF <> True
        P(Month(Format(r_p.Fields(0).Value, "dd-mmm-yyyy"))) = P(Month(Format(r_p.Fields(0).Value, "dd-mmm-yyyy"))) + r_p.Fields(1).Value
        r_p.MoveNext
Wend



While r_e.EOF <> True
        E(Month(Format(r_e.Fields(0).Value, "dd-mmm-yyyy"))) = E(Month(Format(r_e.Fields(0).Value, "dd-mmm-yyyy"))) + r_e.Fields(2).Value
        r_e.MoveNext
Wend

While r_i.EOF <> True
    i(Month(Format(r_i.Fields(0).Value, "dd-mmm-yyyy"))) = i(Month(Format(r_i.Fields(0).Value, "dd-mmm-yyyy"))) + r_i.Fields(2).Value
    r_i.MoveNext
Wend



Dim UPDATE_PROFIT As New ADODB.Recordset

UPDATE_PROFIT.Open "SELECT * FROM MONTH_NAME_PROFIT", db, adOpenKeyset, adLockOptimistic

While UPDATE_PROFIT.EOF <> True
    UPDATE_PROFIT.Delete
    UPDATE_PROFIT.MoveNext
Wend

UPDATE_PROFIT.Requery

Dim j As Integer
For j = 1 To 12
    
    UPDATE_PROFIT.AddNew
    UPDATE_PROFIT.Fields(0).Value = MonthName(j)
    UPDATE_PROFIT.Fields(1).Value = P(j)
    UPDATE_PROFIT.Fields(2).Value = E(j)
    UPDATE_PROFIT.Fields(3).Value = i(j)
    UPDATE_PROFIT.Fields(4).Value = P(j) + (i(j) - E(j))
    UPDATE_PROFIT.Update
Next
    monthly_profit_report.DataFiles(0) = App.Path & "\Master_Database.mdb"
    monthly_profit_report.ReportFileName = App.Path & "\Report\MONTHLY_PROFIT.rpt"
    monthly_profit_report.username = "Admin"
    monthly_profit_report.Password = "1010101010" & Chr(10) & "1010101010"
    monthly_profit_report.Action = 1
End Sub

Private Sub mnu_pur_due_report_Click()
 With Form1.unpaid_purchase
.DataFiles(0) = App.Path & "\Master_Database.mdb"
.ReportFileName = App.Path & "\Report\rpt_unpaid_report.rpt"
.SelectionFormula = "{QRY_UNPAID_REPORT.TRAN_TYPE} = 'PURCHASE'"
.username = "Admin"
.Password = "1010101010" & Chr(10) & "1010101010"
.Action = 1
.PageZoom (100)
End With
End Sub

Private Sub mnu_pur_names_Click()
On Error Resume Next
    FRM_PURCHASE_P_NAMES.Show
Exit Sub
End Sub

Private Sub mnu_pur_payment_Click()
On Error Resume Next
unpaid_purchase.Show
Exit Sub
End Sub

Private Sub mnu_purchase_bills_Click()
On Error Resume Next
    FRM_VIEW_PURCHASE.Show
Exit Sub
End Sub

Private Sub mnu_purchase_book_Click()
frm_pr_book.Show
End Sub

Private Sub mnu_purchase_Click()
    Purchase_form.Show
End Sub



Private Sub mnu_remind_Click()
FRM_TASK.Show
End Sub

Private Sub mnu_remove_cust_ex_check_Click()
On Error Resume Next
    frm_remove_from_expiry.Show
Exit Sub
End Sub

Private Sub mnu_restore_Click()
FRM_RESTORE.Show
End Sub

Private Sub mnu_sales_bill_Click()
On Error Resume Next
frm_view_sales.Show
Exit Sub
End Sub

Private Sub mnu_sales_book_Click()
frm_sal_book.Show
End Sub

Private Sub mnu_sales_Click()
    sales_option.Show
End Sub

Private Sub mnu_sec_pol_Click()
frm_sec.Show
End Sub

Private Sub mnu_send_msg_Click()
frm_sent_msg.Show
End Sub

Private Sub mnu_showst_tips_Click()
        Dim chk_st As New ADODB.Recordset
        chk_st.Open "select * from tip_status", db, adOpenKeyset, adLockOptimistic
        
        If chk_st.Fields(0).Value = False Then
                chk_st.Fields(0).Value = True
                chk_st.Update
                chk_st.Close
            
        Else
                chk_st.Close
        End If
        FRM_TIPS.Show
End Sub

Private Sub mnu_sol_Click()
On Error Resume Next
    Shell "sol", vbNormalFocus
Exit Sub
End Sub

Private Sub mnu_stock_Click()
With Form1.CrystalReport1
    .DataFiles(0) = App.Path & "\Master_Database.MDB"
    .ReportFileName = App.Path & "\Report\RPT_AVA_STOCK.rpt"
    .username = "Admin"
    .Password = "1010101010" & Chr(10) & "1010101010"
    .Action = 1
    .PageZoom (100)
End With
End Sub

Private Sub mnu_sys_clean_Click()
Dim tr As Integer
tr = MsgBox("Are you sure you want to Clear Log Details ...", vbYesNo Or vbQuestion, "Want to Clear Log Details ...")
If tr = 6 Then
            Dim c As New ADODB.Recordset
            c.Open "select * from LOG_DETAIL where LOG_OUT_TIME <> 'CURRENT'", db, adOpenKeyset, adLockOptimistic
            While c.EOF <> True
                c.Delete
                c.MoveNext
            Wend
            MsgBox "Log Details Cleared ...", vbInformation, "Log File Cleared ..."
End If
End Sub

Private Sub mnu_sys_info_Click()
FRM_SYS_INFO.Show
End Sub

Private Sub mnu_t_vertical_Click()
MDIForm1.Arrange vbTileVertical
End Sub


Private Sub mnu_tile_horizon_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub

Public Sub UPDATE_LOG()
    On Error Resume Next
    
    
    Dim d As New ADODB.Recordset
    d.Open "SELECT * FROM LOG_DETAIL WHERE LOG_OUT_TIME='CURRENT'", db, adOpenDynamic, adLockOptimistic
    d.Fields(2).Value = Now
    d.Update
    d.Close
    
    DoEvents
End Sub

Private Sub mnu_update_purchase_entry_Click()
On Error Resume Next
    frm_invoice_no_pname_option.Show
Exit Sub
End Sub

Private Sub mnu_update_sales_entry_Click()
On Error Resume Next
    FRM_UPDATE_SALES.Show
Exit Sub
End Sub

Private Sub mnu_we_Click()
On Error Resume Next
Shell ("explorer"), vbMaximizedFocus
Exit Sub

End Sub

Private Sub rp_menu_int_sales_Click()
Call mnu_int_sales_report_Click
End Sub


Private Sub un_cr_report_Click()
mnu_stock_Click
End Sub

Private Sub un_log_report_Click()
mnu_log_details_Click
End Sub

Private Sub un_pr_book_Click()
mnu_purchase_book_Click
End Sub

Private Sub un_pur_un_report_Click()
mnu_pur_due_report_Click
End Sub

Private Sub un_sales_book_Click()
mnu_sales_book_Click
End Sub

Private Sub un_sales_unpaid_Click()
mnu_payment_due_report_Click
End Sub
