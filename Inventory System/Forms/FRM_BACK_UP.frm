VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_BACK_UP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Back up Database ..."
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "FRM_BACK_UP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7590
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   6240
      TabIndex        =   4
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   4335
   End
   Begin LVbuttons.LaVolpeButton but_gen_rpt 
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Create Backup"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_BACK_UP.frx":0E42
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image2 
      Height          =   4830
      Left            =   0
      Picture         =   "FRM_BACK_UP.frx":0E5E
      Top             =   -120
      Width           =   2490
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2760
      TabIndex        =   6
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Path Where to Store Backup"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Path and Click on Create Backup."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "FRM_BACK_UP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1

Private Sub but_gen_rpt_Click()
If Len(Text2.Text) = 0 Then
    MsgBox "Select Backup File Path and then Click Create Backup Button...", vbInformation, "Select Backup Path ..."
    Exit Sub
End If

  Label3.Visible = True
  
  Dim dest As String
  dest = Text2.Text & "\Inventory_Backup.bk"
  Dim OldTimer As Single
  ProgressBar1.Visible = True
  OldTimer = Timer
  
  Call Huffman.EncodeFile(App.Path & "\Master_database.mdb", dest)
  ProgressBar1.Value = 0
  
  Unload Me
  Exit Sub
End Sub

Private Sub Command1_Click()
frm_sel_path.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If ProgressBar1.Value = 0 Or ProgressBar1.Value Then
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Left = 0
Me.Top = 0
Set Huffman = New clsHuffman
Label3.Visible = False
ProgressBar1.Visible = False

End Sub

Private Sub Huffman_Progress(Procent As Integer)
  Label3.Caption = "Compressing Database"
 ProgressBar1.Value = Procent
  If ProgressBar1.Value = 100 Then
    Label3.Caption = "Saving Compressed File ..."
    End If
  DoEvents

End Sub

