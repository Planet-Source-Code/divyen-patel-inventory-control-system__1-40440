VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frm_remove_from_expiry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Customer From Expiry Check ..."
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frm_remove_from_expiry.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6120
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Save"
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
      MICON           =   "frm_remove_from_expiry.frx":0E42
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   16
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   15
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   14
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   13
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   12
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   11
      Top             =   480
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Do not Show this Record in the Expiry Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3240
      Width           =   5415
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
      MICON           =   "frm_remove_from_expiry.frx":0E5E
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
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
      MICON           =   "frm_remove_from_expiry.frx":0E7A
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
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ">"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
      MICON           =   "frm_remove_from_expiry.frx":0E96
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
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   9
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ">>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
      MICON           =   "frm_remove_from_expiry.frx":0EB2
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
      Caption         =   "Expire Date"
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
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registered Date"
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
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
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
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Package Details"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Name"
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
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer ID"
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
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   120
      Picture         =   "frm_remove_from_expiry.frx":0ECE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frm_remove_from_expiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset

Private Sub Check1_Click()
If Check1.Value = 1 Then
    LaVolpeButton1.Enabled = True
    For i = 0 To cmd.Count - 1
        cmd(i).Enabled = False
    Next
Else
    LaVolpeButton1.Enabled = False
    For i = 0 To cmd.Count - 1
        cmd(i).Enabled = True
    Next
    
End If

End Sub

Private Sub cmd_Click(Index As Integer)
If Index = 0 Then
    Rs.MoveFirst
ElseIf Index = 1 Then
    Rs.MovePrevious
    If Rs.BOF = True Then
        Rs.MoveFirst
        
    End If
ElseIf Index = 2 Then
       Rs.MoveNext
    If Rs.EOF = True Then
        Rs.MoveLast
  
    End If
    
ElseIf Index = 3 Then
    Rs.MoveLast
End If
Call FILLTEXT
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Dim x As Integer
    x = MsgBox("Are you sure you want to close Expiry Removal Form ...", vbQuestion Or vbYesNo, "Want to Close ...")
    If x = 6 Then
         'Dim f As New FileSystemObject
        'f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True
       
        Unload Me
    End If
End If

End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Left = 0
Me.Top = 0

Rs.Open "select * from sys_qry_expiry_check", db, adOpenKeyset, adLockOptimistic


If Rs.EOF <> True Then
    Call FILLTEXT
Else
    MsgBox "There is no Expired Internet Account Found ...", vbInformation, "No Expire Record Found ..."
    Unload Me
    Exit Sub

End If


LaVolpeButton1.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rs.Close
End Sub

Public Sub FILLTEXT()
For i = 0 To Text1.Count - 1
    Text1(i).Text = Clear
    If Len(Rs.Fields(i).Value) > 0 Then
        Text1(i).Text = Rs.Fields(i).Value
    End If
    
Next
End Sub

Private Sub LaVolpeButton1_Click()
LaVolpeButton1.Enabled = False
    For i = 0 To cmd.Count - 1
        cmd(i).Enabled = True
    Next
If Check1.Value = 1 Then
    Rs.Fields(8).Value = False
    Rs.Update
    Rs.Requery
    If Rs.RecordCount = 0 Then
        MsgBox "All Records Updated for Expiry Check ...", vbInformation, "Records Updated ..."
        'Dim f As New FileSystemObject
        'f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True
       
        Unload Me
        Exit Sub
    Else
        Call FILLTEXT
    End If
End If
Check1.Value = 0
End Sub
