VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_sec 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Policy"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frm_sec.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   6945
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "General Sequrity Policy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   6735
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   960
         Width           =   255
      End
      Begin VB.ComboBox text1 
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
         Height          =   330
         ItemData        =   "frm_sec.frx":0E42
         Left            =   3360
         List            =   "frm_sec.frx":0E61
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5160
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_sec.frx":0E80
               Key             =   "lock"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_sec.frx":12D2
               Key             =   "unlock"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Maximum Password Length :"
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
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " between 1 to 9"
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
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Disable Sequrity :"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   1
         Left            =   4320
         Top             =   840
         Width           =   615
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   6120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Change It"
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
      MICON           =   "frm_sec.frx":1724
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Operator Sequrity Policy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6735
      Begin VB.CheckBox chk_operator 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Access to Inventory Masters"
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
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   4935
      End
      Begin VB.CheckBox chk_operator 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Access to Cash Management"
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
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   4935
      End
      Begin VB.CheckBox chk_operator 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Access to use Backup Feature"
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
         Index           =   2
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CheckBox chk_operator 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Access to use Restore Feature"
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
         Index           =   3
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   4575
      End
      Begin VB.CheckBox chk_operator 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow to use System Cleaner"
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
         Index           =   4
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear Operator Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3840
         MouseIcon       =   "frm_sec.frx":1740
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "It will Set ""OP""  as Default Password"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   1365
         Left            =   5160
         Picture         =   "frm_sec.frx":1892
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Label issues 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm_sec.frx":1E24
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
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   840
      Index           =   0
      Left            =   0
      Picture         =   "frm_sec.frx":1EAE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10080
   End
End
Attribute VB_Name = "frm_sec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim R1 As New ADODB.Recordset
Dim r2 As New ADODB.Recordset
Dim R3 As New ADODB.Recordset

Private Sub Check1_Click()
If Check1.Value = 0 Then
    Image1(1).Picture = ImageList1.ListImages(1).Picture
Else
    Image1(1).Picture = ImageList1.ListImages(2).Picture
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Left = 0
Me.Top = 0
Text1.Clear
For i = 1 To 9
    Text1.AddItem i
Next

R1.Open "SELECT * FROM SECURITY_POLICY", db, adOpenKeyset, adLockOptimistic
r2.Open "SELECT * FROM PASS_LENGTH", db, adOpenKeyset, adLockOptimistic
R3.Open "SELECT * FROM CHECK_SECURITY", db, adOpenKeyset, adLockOptimistic


For i = 0 To 4
    If R1.Fields(i).Value = True Then
        chk_operator(i).Value = 1
    Else
        chk_operator(i).Value = 0
    End If
Next

Text1.Text = r2.Fields(0).Value

If R3.Fields(0).Value = True Then
    Check1.Value = 0
Else
    Check1.Value = 1
End If


If Check1.Value = 0 Then
    Image1(1).Picture = ImageList1.ListImages(1).Picture
Else
    Image1(1).Picture = ImageList1.ListImages(2).Picture
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
R1.Close
r2.Close
R3.Close


End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = False
End Sub

Private Sub Label1_Click()
    Dim LOG_CO As New ADODB.Connection
    LOG_CO.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\COMPANY_MASTER.mdb;Persist Security Info=False;Jet OLEDB:Database Password=1010101010"

Dim clear_pass As New ADODB.Recordset
clear_pass.Open "select password from uid_pass where user_name='OPERATOR'", LOG_CO, adOpenKeyset, adLockOptimistic
clear_pass.Fields(0).Value = "OP"
clear_pass.Update
MsgBox "Password Changed ...", vbInformation, "Password Updated ..."
LOG_CO.Close

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = True
End Sub

Private Sub LaVolpeButton1_Click()
For i = 0 To 4
    If chk_operator(i).Value = 1 Then
            R1.Fields(i).Value = True
    Else
            R1.Fields(i).Value = False
    End If
Next
R1.Update

r2.Fields(0).Value = Text1.Text
r2.Update

If Check1.Value = 1 Then
    R3.Fields(0).Value = False
Else
    R3.Fields(0).Value = True
End If
R3.Update
If Check1.Value = 1 Then
    MDIForm1.mnu_log_out.Visible = False
End If

Unload Me
End Sub
