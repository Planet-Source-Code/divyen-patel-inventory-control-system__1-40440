VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_income 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Data Entry Form ..."
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frm_income.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   44498947
      CurrentDate     =   37457
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&New"
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
      MICON           =   "frm_income.frx":0E42
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   435
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "&Save"
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
      MICON           =   "frm_income.frx":0E5E
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frm_income"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ST As Boolean

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Combo1.Text) > 0 Then
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
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
    ST = False
    Me.Left = MDIForm1.Width / 2 - Me.Width / 2
    Me.Top = MDIForm1.Height / 2 - Me.Height / 2
    
    Call refresh_extype
    
    DTPicker1.Value = Date
End Sub

Public Sub refresh_extype()
Dim R As New ADODB.Recordset
R.Open "select * from Income_types", db, adOpenDynamic, adLockOptimistic
Combo1.Clear
While R.EOF <> True
    Combo1.AddItem R.Fields(0).Value
    R.MoveNext
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
If ST = True Then
    Dim R As New ADODB.Recordset
    R.Open "select * from Income_types where Income_types='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
    R.Delete
    R.Close
End If
End Sub

Private Sub LaVolpeButton1_Click()
frm_income_types.Show vbModal
End Sub

Private Sub LaVolpeButton3_Click()
    If Len(Text1.Text) = 0 Then
        MsgBox "Amount Not Found ...", vbCritical, "Enter Proper Data ..."
        Exit Sub
    End If
    If Len(Combo1.Text) = 0 Then
        MsgBox "Income Type Not Selected ...", vbCritical, "Enter Proper Data ..."
        Exit Sub
    End If
    
    Dim x As Integer
    x = MsgBox("Are you sure you want to save this entry ...", vbQuestion Or vbYesNo, "Want to save this entry ...")
    If x = 6 Then
    
    Dim add_ex As New ADODB.Recordset
    add_ex.Open "select * from Income", db, adOpenDynamic, adLockOptimistic
    
    add_ex.AddNew
    add_ex.Fields(1).Value = Combo1.Text
    add_ex.Fields(0).Value = Format(DTPicker1.Value, "dd-MMM-yyyy")
    add_ex.Fields(2).Value = Text1.Text
    add_ex.Update
    ST = False
    add_ex.Close
        'Dim f As New FileSystemObject
        'f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True
       
    Unload Me
    End If
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        If Len(Text1.Text) > 0 And VAL(Text1.Text) > 0 Then
                LaVolpeButton3_Click
        End If
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = 46 Then
        If InStr(1, Text1.Text, ".", vbTextCompare) > 0 Then
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

