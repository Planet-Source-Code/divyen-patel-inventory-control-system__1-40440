VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_user_pass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter User name and Password to Login ..."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frm_user_pass.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unmask the Password"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   5415
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Login ..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "frm_user_pass.frx":0E42
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   120
      Picture         =   "frm_user_pass.frx":0E5E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      MouseIcon       =   "frm_user_pass.frx":13F0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Warning : Be Sure when you use this Option , It will show Password In Alphabetic Characters"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm_user_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text1.PasswordChar = ""
Else
    Text1.PasswordChar = "*"
End If

End Sub


Private Sub Combo1_Click()
Text1.Text = Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
            On Error Resume Next
            Unload FRM_COMPANY
            Unload MDIForm1
            Unload frm_comp_name
            Unload frmSplash
    
                Dim x As Integer
                x = MsgBox("Are you sure you want to Close the Application ...", vbQuestion Or vbYesNo, "Want to Close Application ..")
                If x = 6 Then
                        Unload Me
                End If
    End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Combo1.Clear
Combo1.AddItem "ADMINISTRATOR"
Combo1.AddItem "OPERATOR"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
            Unload FRM_COMPANY
            Unload MDIForm1
            Unload frm_comp_name
            Unload frmSplash
End Sub

Private Sub Label3_Click()
If Len(Combo1.Text) = 0 Then
    MsgBox "Please Select User Name and then Click on Change Password ...", vbInformation, "Select User Name ..."
Else
    Text1.Text = Clear
    frm_change_pass.usernam = Combo1.Text
    frm_change_pass.Show vbModal
    
End If
End Sub

Private Sub LaVolpeButton1_Click()

If db.State = adStateClosed Then
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Master_Database.mdb;Persist Security Info=False;Jet OLEDB:Database Password=1010101010"
End If

If Len(Combo1.Text) > 0 Then

Dim R As New ADODB.Recordset
R.Open "SELECT PASSWORD FROM UID_PASS WHERE USER_NAME='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic

If UCase(R.Fields(0).Value) = UCase(Text1.Text) Then
    MDIForm1.cur_user = Combo1.Text
    MDIForm1.Label1(1).Caption = LCase(Combo1.Text)
    MDIForm1.Label1(3).Caption = Time
    Dim UPDATE_LOG As New ADODB.Recordset
    UPDATE_LOG.Open "SELECT * FROM LOG_DETAIL", db, adOpenKeyset, adLockOptimistic
    
    UPDATE_LOG.AddNew
    UPDATE_LOG.Fields(0).Value = Combo1.Text
    UPDATE_LOG.Fields(1).Value = Now
    UPDATE_LOG.Fields(2).Value = "CURRENT"
    UPDATE_LOG.Update

    
    Me.Visible = False
    Text1.Text = Clear
    MDIForm1.Show
    
Else
    MsgBox "Invalid Password ...", vbCritical, "Enter Proper Password ..."
    Text1.Text = Clear
End If

Else
    MsgBox "Username Not Selected ...", vbCritical
End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call LaVolpeButton1_Click
End If
End Sub
