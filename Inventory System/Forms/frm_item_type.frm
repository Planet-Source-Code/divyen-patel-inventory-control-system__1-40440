VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_item_type 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter New Item Type"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "frm_item_type.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add Item type"
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
      MICON           =   "frm_item_type.frx":0E42
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
      Caption         =   "Item Type"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frm_item_type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim item_type_rs As New ADODB.Recordset
Public FN As String


Private Sub cmd_Click()
If Len(Text1.Text) <> 0 Then

item_type_rs.AddNew
item_type_rs.Fields(0).Value = Text1.Text
On Error GoTo A1:
item_type_rs.Update
If FN = "M" Then
        FRM_MODIFY_PURCHASE_BOOK.Refresh_combobox (2)
        FRM_MODIFY_PURCHASE_BOOK.Combo2.Text = Text1.Text
Else
        Purchase_form.Refresh_combobox (2)
        Purchase_form.Combo2.Text = Text1.Text
End If
Unload Me
Else
MsgBox "Enter Item type ...", vbInformation, "You can not save Zero length Item name ..."
End If
Exit Sub
A1:
MsgBox "Duplicate Item name Found ..." & vbCrLf & "Enter Another name of Close this form ...", vbCritical, "Duplicate Entry Found ..."
item_type_rs.CancelUpdate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    
    Unload Me
End If

End Sub

Private Sub Form_Load()
        KeyPreview = True
        item_type_rs.Open "select * from ItemType", db, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
item_type_rs.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Text1.Text) <> 0 Then
If KeyCode = 13 Then
    Call cmd_Click
End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then

ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then

ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then

ElseIf KeyAscii = 8 Then

ElseIf KeyAscii = 32 Then

Else
    KeyAscii = 0
End If
End Sub
