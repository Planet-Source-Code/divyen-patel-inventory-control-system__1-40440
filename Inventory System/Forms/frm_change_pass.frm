VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_change_pass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   Icon            =   "frm_change_pass.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
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
      Index           =   2
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2055
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
      Index           =   1
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2055
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
      Index           =   0
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Change"
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
      MICON           =   "frm_change_pass.frx":0E42
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
      Caption         =   "Confirm Password"
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
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Password"
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
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Old Password"
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
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frm_change_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public usernam As String
Private Sub Form_Load()
LaVolpeButton1.Enabled = False
End Sub

Private Sub LaVolpeButton1_Click()

    
    

    Dim check_len As New ADODB.Recordset
    check_len.Open "select * from PASS_LENGTH", db, adOpenKeyset, adLockOptimistic
If Len(Text1(2).Text) > check_len.Fields(0).Value Then
    MsgBox "Password Length is not matched with Administrator Policy ...", vbInformation, "Change the password length ..."
    Exit Sub
End If
check_len.Close

Dim R As New ADODB.Recordset
R.Open "select password from UID_PASS where user_name='" & usernam & "'", db, adOpenDynamic, adLockOptimistic
If UCase(R.Fields(0).Value) = UCase(Text1(0).Text) Then
    R.Fields(0).Value = Text1(2).Text
    R.Update
    MsgBox "Password Changed Successfully", vbInformation, "Password Changed .."
    Unload Me
Else
    MsgBox "Your Old Password is not valid ...", vbCritical, "Password Not Valid .."
End If


End Sub

Private Sub Text1_Change(Index As Integer)


If Index = 1 Or Index = 2 Then
    If Len(Text1(1).Text) Then
        LaVolpeButton1.Enabled = False
    End If
    
    If Len(Text1(2).Text) Then
        LaVolpeButton1.Enabled = False
    End If

    If Text1(1).Text = Text1(2).Text Then
        LaVolpeButton1.Enabled = True
    Else
        LaVolpeButton1.Enabled = False
    End If
    
End If
End Sub
