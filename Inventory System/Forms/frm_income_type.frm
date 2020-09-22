VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_income_types 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Income Types"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frm_income_type.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
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
      MICON           =   "frm_income_type.frx":0E42
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
      Caption         =   "Income Type"
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
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frm_income_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
If Len(Text1.Text) > 0 Then
Dim rs As New ADODB.Recordset
rs.Open "select * from Income_types", db, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields(0).Value = Text1.Text
On Error GoTo t1
    rs.Update
    rs.Close
    frm_income.refresh_extype
    frm_income.Combo1.Text = Text1.Text
    frm_income.Combo1.Enabled = False
    frm_income.LaVolpeButton1.Enabled = False
    frm_income.ST = True
    
    Unload Me
    
    
Exit Sub
t1:
    MsgBox "Income Type Already Exist ...", vbInformation, "Duplicate Entry Found ..."
    SendKeys "{TAB}"
Else
    MsgBox "Enter Income Type ...", vbCritical, "Null Entry Can not be Saved ..."
    SendKeys "{TAB}"
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Left = frm_income.Left + (frm_income.Width / 2 - Me.Width / 2)
Me.Top = frm_income.Top + (frm_income.Height / 2 - Me.Height / 2)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If frm_income.ST = False Then
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
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
