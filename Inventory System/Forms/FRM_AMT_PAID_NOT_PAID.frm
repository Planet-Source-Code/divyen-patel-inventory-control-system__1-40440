VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRM_AMT_PAID_NOT_PAID 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount Paid Or Not ?"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   Icon            =   "FRM_AMT_PAID_NOT_PAID.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction Finished"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount Not Paid"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   3375
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Ok"
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
      MICON           =   "FRM_AMT_PAID_NOT_PAID.frx":0E42
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
      Left            =   120
      Picture         =   "FRM_AMT_PAID_NOT_PAID.frx":0E5E
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   14
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice No :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Party Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4320
      TabIndex        =   8
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Out Of"
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
      Left            =   4320
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "How much Paid ?"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Are the Amount of this Transaction Paid or not ?, if paid then how much amount paid ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   3615
   End
End
Attribute VB_Name = "FRM_AMT_PAID_NOT_PAID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dt As String

Private Sub cmd_Click()
If Option1(0).Value = True Then
    If Len(Text1.Text) = 0 Then
        MsgBox "Enter Amount of Rs Paid ...", vbInformation, "Paid Amount not Found ..."
        Exit Sub
    End If
    
    If VAL(Text1.Text) = 0 Then
        MsgBox "Paid Amount can not be Zero" & vbCrLf & "Select Amount Not paid Option , If Amount is not paid ...", vbCritical, "Zero Value not Allowed ..."
        Exit Sub
    End If
    
    
    Dim Rs As New ADODB.Recordset
    
    If Check1.Value = 1 Then
        Dim x As Integer
        x = MsgBox("Are you sure this Transaction is Complete", vbQuestion Or vbYesNo, "Trasaction Completed ?")
        If x <> 6 Then
            Check1.Value = 0
            Exit Sub
        End If
        
            
        If (VAL(Label2(2).Caption) - VAL(Text1.Text)) > 0 Then
            If Label3(5).Caption = "Sales" Then
                Rs.Open "SELECT * FROM EXPENSE", db, adOpenDynamic, adLockOptimistic
                Rs.AddNew
                Rs.Fields(0).Value = dt
                Rs.Fields(1).Value = "Discounted Amount(Sales)"
                Rs.Fields(2).Value = VAL(Label2(2).Caption) - VAL(Text1.Text)
                Rs.Fields(3).Value = Label3(2).Caption
                
                Rs.Update
                Rs.Close
            ElseIf Label3(5).Caption = "Purchase" Then
                Rs.Open "SELECT * FROM INCOME", db, adOpenDynamic, adLockOptimistic
                Rs.AddNew
                Rs.Fields(0).Value = dt
                Rs.Fields(1).Value = "Discounted Amount(Purchase)"
                Rs.Fields(2).Value = VAL(Label2(2).Caption) - VAL(Text1.Text)
                Rs.Fields(3).Value = Label3(2).Caption
                Rs.Update
                Rs.Close
            
            End If
            
        End If
        'Dim f As New FileSystemObject
        'f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True

        Unload Me
        Exit Sub
    End If
    
    Rs.Open "select * from AMT_UNPAID_REMIND", db, adOpenDynamic, adLockOptimistic
    Rs.AddNew
    
    If Label3(5).Caption = "Sales" Then
        Rs.Fields(0).Value = "SALES"
    ElseIf Label3(5).Caption = "Purchase" Then
        Rs.Fields(0).Value = "PURCHASE"
    End If
    
    'If Len(dt) = 0 Then
    '    dt = Format(Date, "dd-MMM-yyyy")
    'End If
    
    
    Rs.Fields(1).Value = dt
    Rs.Fields(2).Value = VAL(Label2(2).Caption) - VAL(Text1.Text)
    Rs.Fields(3).Value = Label3(2).Caption
    Rs.Fields(4).Value = Label3(0).Caption
    
    Rs.Update
        
       

    Unload Me

ElseIf Option1(1).Value = True Then
Rs.Open "select * from AMT_UNPAID_REMIND", db, adOpenDynamic, adLockOptimistic
    Rs.AddNew
    If Label3(5).Caption = "Sales" Then
        Rs.Fields(0).Value = "SALES"
    ElseIf Label3(5).Caption = "Purchase" Then
        Rs.Fields(0).Value = "PURCHASE"
    End If
    

    Rs.Fields(1).Value = dt
    Rs.Fields(2).Value = VAL(Label2(2).Caption) - VAL(Text1.Text)
    Rs.Fields(3).Value = Label3(2).Caption
    Rs.Fields(4).Value = Label3(0).Caption
    Rs.Update
        
       

    Unload Me
End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
    Me.Left = MDIForm1.Width / 2 - Me.Width / 2
    Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - MDIForm1.StatusBar1.Height
    
    Label2(0).Visible = False
    Text1.Visible = False
    Label2(1).Visible = False
    Label2(2).Visible = False
    Check1.Visible = False
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
    Label2(0).Visible = True
    Text1.Visible = True
    Text1.Text = Clear
    Label2(1).Visible = True
    Label2(2).Visible = True
    Check1.Visible = True
ElseIf Index = 1 Then
    Label2(0).Visible = False
    Text1.Text = Clear
    Text1.Visible = False
    Label2(1).Visible = False
    Label2(2).Visible = False
    Check1.Value = 0
    Check1.Visible = False
    
End If
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Then
    If KeyCode = 40 Then
        Option1_Click (1)
    End If
ElseIf Index = 1 Then
    If KeyCode = 38 Then
        Option1_Click (0)
    End If
End If

End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
If Index = 0 Then
    SendKeys "{TAB}"
Else
    Call cmd_Click
End If
End If
End Sub

Private Sub Text1_Change()

If VAL(Text1.Text) > VAL(Label2(2).Caption) Then
    MsgBox "You can not Enter Greater than the required value ...", vbCritical, "Enter Proper Value ..."
    Text1.Text = Clear
End If

If VAL(Text1.Text) = VAL(Label2(2).Caption) Then
    Check1.Value = 1
    Check1.Enabled = False
End If

If VAL(Text1.Text) <> VAL(Label2(2).Caption) Then
    Check1.Value = 0
    Check1.Enabled = True
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Text1.Text) > 0 Then
        Call cmd_Click
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
