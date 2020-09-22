VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRM_PURCHASE_P_NAMES 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Purchase Party Names ..."
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   ControlBox      =   0   'False
   Icon            =   "FRM_PURCHASE_P_NAMES.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5925
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Modify"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
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
      MICON           =   "FRM_PURCHASE_P_NAMES.frx":0E42
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
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   5655
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9
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
      MICON           =   "FRM_PURCHASE_P_NAMES.frx":0E5E
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "FRM_PURCHASE_P_NAMES.frx":0E7A
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "FRM_PURCHASE_P_NAMES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    If LaVolpeButton1(1).Enabled = True Then
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    KeyPreview = True
    On Error Resume Next
    Dim ac_frm As String
    ac_frm = Clear
    ac_frm = MDIForm1.ActiveForm.Name

    If Len(ac_frm) <> 0 Then
        MsgBox "You can not Modify Supplier name When any of the Form is Opened ...", vbInformation, "Close all Forms ..."
        Unload Me
        Exit Sub
    End If

    
    
    
    
    Text1.Enabled = False
    rs.Open "SELECT * FROM purchase_partynames", db, adOpenKeyset, adLockOptimistic
    
    List1.Clear
    
    While rs.EOF <> True
        List1.AddItem rs.Fields(0).Value
        rs.MoveNext
    Wend
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    rs.Close
End Sub

Private Sub LaVolpeButton1_Click(Index As Integer)
If Index = 0 Then
                    If Len(Text1.Text) > 0 Then
                                If LaVolpeButton1(0).Caption = "&Modify" Then
                                        Text1.Enabled = True
                                        List1.Enabled = False
                                        LaVolpeButton1(0).Caption = "&SAVE"
                                        LaVolpeButton1(1).Enabled = False
                                        SendKeys "{TAB}"
                                        SendKeys "{END}"
                                        
                                Else
                                                    
                                    If Len(Text1.Text) > 0 Then
                                        Dim R As New ADODB.Recordset
                                        R.Open "select * from purchase_partynames where purchase_partyname='" & List1.List(List1.ListIndex) & "'", db, adOpenKeyset, adLockOptimistic
                                        R.Fields(0).Value = Text1.Text
                                        R.Update
                                        db.Execute "UPDATE AMT_UNPAID_REMIND SET PARTY_NAME='" & Text1.Text & "' WHERE PARTY_NAME='" & List1.List(List1.ListIndex) & "'"
                                        
                                        Text1.Enabled = False
                                        List1.Enabled = True
                                        LaVolpeButton1(1).Enabled = True
                                        LaVolpeButton1(0).Caption = "&Modify"
                                        Text1.Text = Clear
                                        MsgBox "Party name Updated Successfully", vbInformation, "Party name Changed ..."
                                        
                                        SendKeys "{TAB}"
                                        SendKeys "{TAB}"
                                        
                                        List1.Clear
                                        rs.Requery
                                        
                                        While rs.EOF <> True
                                            List1.AddItem rs.Fields(0).Value
                                            rs.MoveNext
                                        Wend
                                    Else
                                        MsgBox "Enter Party name ...", vbInformation, "Party name not found ..."
                                    End If
                                End If
                    End If
Else
        Unload Me
End If

End Sub

Private Sub List1_Click()
    Text1.Text = List1.List(List1.ListIndex)
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
