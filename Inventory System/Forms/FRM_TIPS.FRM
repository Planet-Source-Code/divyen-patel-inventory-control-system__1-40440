VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRM_TIPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip of the Day"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "FRM_TIPS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6375
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      Picture         =   "FRM_TIPS.frx":0442
      ScaleHeight     =   3615
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   4815
      End
      Begin VB.Line Line1 
         X1              =   960
         X2              =   6360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know ..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Do not Show tips on start up"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3135
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Close"
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_TIPS.frx":0FB3
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Next Tip"
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_TIPS.frx":0FCF
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
End
Attribute VB_Name = "FRM_TIPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim LAST_TIPNO_RS As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
KeyPreview = True
Me.Top = 300
Me.Left = (MDIForm1.Width / 2) - ((MDIForm1.Toolbar2.Width / 2) + (Me.Width / 2))
Dim LASTNO As Integer
Rs.Open "SELECT * FROM TIPS_OF_DAY", db, adOpenDynamic, adLockOptimistic
LAST_TIPNO_RS.Open "SELECT * FROM last_tip_no", db, adOpenDynamic, adLockOptimistic

LASTNO = LAST_TIPNO_RS.Fields(0).Value + 1

While Rs.EOF <> True
    If Rs.Fields(0).Value = LASTNO Then
        Label2.Caption = Rs.Fields(1).Value
        LAST_TIPNO_RS.Fields(0).Value = LASTNO
        LAST_TIPNO_RS.Update
        Exit Sub
    End If
    Rs.MoveNext
Wend
LASTNO = 1
Rs.MoveFirst
Label2.Caption = Rs.Fields(1).Value
LAST_TIPNO_RS.Fields(0).Value = LASTNO
LAST_TIPNO_RS.Update
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rs.Close

LAST_TIPNO_RS.Close
Dim TS As New ADODB.Recordset
TS.Open "SELECT * FROM tip_status", db, adOpenDynamic, adLockOptimistic

If Check1.Value = 1 Then
    TS.Fields(0).Value = 0
ElseIf Check1.Value = 0 Then
    TS.Fields(0).Value = 1
End If

TS.Update
TS.Close

End Sub

Private Sub LaVolpeButton1_Click()
Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
LASTNO = LAST_TIPNO_RS.Fields(0).Value + 1
Rs.MoveFirst
While Rs.EOF <> True
    If Rs.Fields(0).Value = LASTNO Then
        Label2.Caption = Rs.Fields(1).Value
        LAST_TIPNO_RS.Fields(0).Value = LASTNO
        LAST_TIPNO_RS.Update
        Exit Sub
    End If
    Rs.MoveNext
Wend
LASTNO = 1
Rs.MoveFirst
Label2.Caption = Rs.Fields(1).Value
LAST_TIPNO_RS.Fields(0).Value = LASTNO
LAST_TIPNO_RS.Update
End Sub
