VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Management Information System ...."
   ClientHeight    =   7770
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   7455
   ClipControls    =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Continue >>"
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
      MICON           =   "frmSplash.frx":1CFA
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
      BackColor       =   &H00000000&
      Caption         =   "Jai Swaminarayan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   13
      Left            =   5160
      TabIndex        =   16
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   5400
      Picture         =   "frmSplash.frx":1D16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   7320
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Image Image1 
      Height          =   2430
      Left            =   240
      Picture         =   "frmSplash.frx":2994
      Top             =   4800
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   0
      X2              =   6960
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   7320
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   7320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Wish you a very very happy new year."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FD2F4&
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Wish you all a happy and Prosperous Diwali."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FD2F4&
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   14
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "User Name : Operator       / Password  :  op"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   13
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "User Name : Administrator / Password :  admin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   12
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "First Time Passwords : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   11
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Special Thanks to LaVolpe"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   10
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Please Vote me , I am Considering your vote as one of the Important element of my Resume"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   615
      Index           =   8
      Left            =   360
      TabIndex        =   9
      Top             =   4080
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Management Information System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FD2F4&
      Height          =   495
      Index           =   0
      Left            =   690
      TabIndex        =   8
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Including :  Inventory System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   1
      Left            =   990
      TabIndex        =   7
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Financial Management System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   2
      Left            =   2190
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Customer Support System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   3
      Left            =   2190
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Developed Exclusively For  Browse Infosys / Browse System / Browse Infocom"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   495
      Index           =   6
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Developed By Divyen k Patel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Contact : divyen_patel@rediffmail.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FD2F4&
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
If App.PrevInstance = True Then
    MsgBox "System is Already in Run Mode ...", vbCritical, "System is Already Running ..."
    Unload Me
    Exit Sub
End If
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Master_Database.mdb;Persist Security Info=False;Jet OLEDB:Database Password=1010101010"
End Sub

Private Sub LaVolpeButton1_Click()
    Me.Visible = False
    Dim c As New ADODB.Recordset
    c.Open "select * from CHECK_SECURITY", db, adOpenKeyset, adLockOptimistic
    If c.Fields(0).Value = True Then
        frm_user_pass.Show
    Else
        MDIForm1.cur_user = "PASS_DISABLE"
        MDIForm1.Show
    End If
End Sub



