VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_party_name 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Party Name to the List ..."
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frm_party_name.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
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
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add"
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
      MICON           =   "frm_party_name.frx":0E42
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Party Name"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm_party_name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_prt As New ADODB.Recordset
Dim ST As Boolean
Private Sub cmd_Click()
If Len(Text1.Text) > 0 Then
    rs_prt.AddNew
    rs_prt.Fields(0).Value = Text1.Text
    On Error GoTo A1:
    rs_prt.Update
    Purchase_form.Refresh_combobox (3)
    Purchase_form.Combo3.Text = Text1.Text
    ST = True
    Unload Me
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    
    Exit Sub
A1:
    MsgBox "Party Name Already Exist ...", vbInformation, "Enter Another Name ..."
Else
    MsgBox "Enter Party name ...", vbCritical, "Enter Name ..."
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Left = Purchase_form.Width / 2 - Me.Width / 2
Me.Top = Purchase_form.Height / 2 - Me.Height / 2

rs_prt.Open "SELECT * FROM purchase_partynames", db, adOpenDynamic, adLockOptimistic
ST = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs_prt.Close
    If ST = False Then
        Purchase_form.Combo3.Enabled = True
        Purchase_form.LaVolpeButton3.Enabled = True
        Purchase_form.PADD = False
    End If
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_Click
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then

ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then

ElseIf KeyAscii = 8 Then

ElseIf KeyAscii = 32 Then

Else
    KeyAscii = 0
End If
End Sub
