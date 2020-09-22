VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_sal_book 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Book"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frm_sal_book.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   22675459
      CurrentDate     =   37461
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Generate Report"
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
      MICON           =   "frm_sal_book.frx":0E42
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   22675459
      CurrentDate     =   37461
   End
   Begin VB.Label issues 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter From and to Date for viewing Sales Records between them"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "frm_sal_book.frx":0E5E
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frm_sal_book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
Me.Visible = False
With Form1.cr1
    .DataFiles(0) = App.Path & "\Master_Database.mdb"
    .ReportFileName = App.Path & "\Report\rpt_salesbook.rpt"
    .username = "Admin"
    .Password = "1010101010" & Chr(10) & "1010101010"
    .SelectionFormula = "{Sales_master.Sales_date} In Date (" & Format$(DTPicker1.Value, "yyyy,mm,dd") & ") To Date (" & Format$(DTPicker2.Value, "yyyy,mm,dd") & ")"
    .Action = 1
    .PageZoom (89)
End With
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Dim x As Integer
    x = MsgBox("Are you sure you want to exit Form Sales book report ...", vbYesNo Or vbQuestion, "Want to exit ?")
    If x = 6 Then
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
KeyPreview = True

Me.Left = MDIForm1.Width / 2 - Me.Width / 2
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - MDIForm1.StatusBar1.Height

DTPicker1.Format = dtpCustom
DTPicker1.CustomFormat = "dd-MMM-yyyy"
DTPicker1.Day = 1
DTPicker1.Month = Month(Date)
DTPicker1.Year = Year(Date)

DTPicker2.Value = Date
End Sub



