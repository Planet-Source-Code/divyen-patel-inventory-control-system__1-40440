VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_internet_expiry_report 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Expiry Report"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frm_internet_expiry_report.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5145
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      ItemData        =   "frm_internet_expiry_report.frx":0E42
      Left            =   2400
      List            =   "frm_internet_expiry_report.frx":0E4F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Show"
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
      MICON           =   "frm_internet_expiry_report.frx":0E7A
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
      Caption         =   "Expire Option"
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frm_internet_expiry_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        If Len(Combo1.Text) > 0 Then
                Call LaVolpeButton3_Click
        End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Dim x As Integer
    x = MsgBox("Are you sure you want to exit Internet Expiry Report Form", vbYesNo Or vbQuestion, "Want to exit ?")
    If x = 6 Then
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
KeyPreview = True

Me.Left = MDIForm1.Width / 2 - Me.Width / 2
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - MDIForm1.StatusBar1.Height
End Sub

Private Sub LaVolpeButton3_Click()
Me.Visible = False
If Combo1.Text = "Today" Then
        
            With Form1.int_expire
                    .DataFiles(0) = App.Path & "\Master_Database.mdb"
                    .ReportFileName = App.Path & "\Report\rpt_internet_expire.rpt"
                    .username = "Admin"
                    .Password = "1010101010" & Chr(10) & "1010101010"
                    .SelectionFormula = "{INTERNET_CONNECTIONS.status} and {INTERNET_CONNECTIONS.EXPIRE_DATE} =CurrentDate"
                    .Action = 1
                    .PageZoom (100)
            End With
ElseIf Combo1.Text = "This Month" Then

        Dim S As New ADODB.Recordset
        S.Open "SELECT * FROM SYS_MONTH_YEAR_INTERNET_EXPIRE", db, adOpenDynamic, adLockOptimistic
        If S.EOF = True Then
            S.AddNew
        End If
        S.Fields(0).Value = Month(Date)
        S.Fields(1).Value = Year(Date)
        S.Update
        S.Close
        
        S.Open "SELECT * FROM SYS_MONTH_YEAR_INTERNET_EXPIRE", db, adOpenDynamic, adLockOptimistic
        If S.EOF = True Then
            S.AddNew
        End If
        S.Fields(0).Value = Month(Date)
        S.Fields(1).Value = Year(Date)
        S.Update
        S.Close
        
        With Form1.int_expire
            .DataFiles(0) = App.Path & "\Master_Database.mdb"
            .ReportFileName = App.Path & "\Report\rpt_internet_expire.rpt"
            .username = "Admin"
            .Password = "1010101010" & Chr(10) & "1010101010"
            .SelectionFormula = "{INTERNET_CONNECTIONS.sys_month}={SYS_MONTH_YEAR_INTERNET_EXPIRE.MONTH_NAME} and {INTERNET_CONNECTIONS.sys_year}={SYS_MONTH_YEAR_INTERNET_EXPIRE.YEAR} and {INTERNET_CONNECTIONS.status}"
            .Action = 1
            .PageZoom (100)
        End With
ElseIf Combo1.Text = "All Expiry Entries" Then


        
            With Form1.int_expire
                    .DataFiles(0) = App.Path & "\Master_Database.mdb"
                    .ReportFileName = App.Path & "\Report\rpt_internet_expire.rpt"
                    .username = "Admin"
                    .Password = "1010101010" & Chr(10) & "1010101010"
                    .SelectionFormula = "{INTERNET_CONNECTIONS.EXPIRE_DATE} <= CurrentDate and {INTERNET_CONNECTIONS.status}"
                    .Action = 1
                    .PageZoom (100)
            End With

        
End If
Unload Me
End Sub
