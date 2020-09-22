VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_RESTORE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Database ...."
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "FRM_RESTORE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8415
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   4800
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin LVbuttons.LaVolpeButton CMD_RESTORE 
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Restore It !"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "FRM_RESTORE.frx":0442
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
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   3465
      Left            =   120
      Picture         =   "FRM_RESTORE.frx":045E
      Top             =   840
      Width           =   1755
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
      Left            =   2040
      TabIndex        =   8
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Backed Up File and Click on  Restore Button"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected Back Up File Path"
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
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   3375
   End
End
Attribute VB_Name = "FRM_RESTORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1

Private Sub CMD_RESTORE_Click()
  If Len(Label1(1).Caption) = 0 Then
    MsgBox "Select Backup File Path and then Click Restore Button...", vbInformation, "Select Path ..."
    Exit Sub
  End If
  
  Dim OldTimer As Single
  ProgressBar1.Visible = True
  Label3.Visible = True
  
  OldTimer = Time
  Call Huffman.DecodeFile(Label1(1).Caption, App.Path & "\Master_Database_bk.mdb")
  
  ProgressBar1.Value = 0
  
  Unload Me
  Exit Sub
End Sub

Private Sub Dir1_Change()
Label1(1).Caption = ""
On Error GoTo A1:
    File1.Path = Dir1.Path
    Exit Sub
A1:
    MsgBox "Folder Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Drive1_Change()
Label1(1).Caption = ""
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"

End Sub

Private Sub File1_Click()
    Label1(1).Caption = File1.Path & "\" & File1.Filename
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

Me.Top = 0
Me.Left = 0
Label1(1).Caption = ""
Label3.Visible = False
ProgressBar1.Visible = False
Set Huffman = New clsHuffman
Drive1.Drive = "c:"
End Sub

Private Sub Huffman_Progress(Procent As Integer)
  Label3.Caption = "Uncompressing Database"
  ProgressBar1.Value = Procent
  If ProgressBar1.Value = 100 Then
    Label3.Caption = "Restoring Uncompressed File ..."
    End If
  DoEvents
End Sub

