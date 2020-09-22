VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FRM_CUST_SYSTEM_DETAILS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer System Details"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "FRM_CUST_SYSTEM_DETAILS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6705
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Next"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_CUST_SYSTEM_DETAILS.frx":0E42
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
      Index           =   2
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
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
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Text            =   "1"
      Top             =   2880
      Width           =   1335
   End
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
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
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
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Top             =   1440
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
      MICON           =   "FRM_CUST_SYSTEM_DETAILS.frx":0E5E
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
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Verify"
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
      MICON           =   "FRM_CUST_SYSTEM_DETAILS.frx":0E7A
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6000
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6120
      Picture         =   "FRM_CUST_SYSTEM_DETAILS.frx":0E96
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice Number"
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
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   13
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TO"
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
      Left            =   3360
      TabIndex        =   12
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System ID"
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
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Qty."
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
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer ID"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Name"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label issues 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Customer Name, Invoice Number and System Qty and Click on Next >> Button"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "FRM_CUST_SYSTEM_DETAILS.frx":1760
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6840
   End
End
Attribute VB_Name = "FRM_CUST_SYSTEM_DETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_custo_details As New ADODB.Recordset
Public Status As String
Public FNAME As String

Private Sub Combo1_Change()
If Len(Combo1.Text) = 0 Then
    LaVolpeButton2.Enabled = False
    Text1(0).Text = Clear
End If
End Sub

Private Sub Combo1_Click()
If Len(Combo1.Text) = 0 Then
    LaVolpeButton2.Enabled = False
    Text1(0).Text = Clear
Else
    LaVolpeButton2.Enabled = True
    Dim CUST_ID As New ADODB.Recordset
    CUST_ID.Open "SELECT cutomer_id FROM Customer_master WHERE cutomer_name='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
    Text1(0).Text = CUST_ID.Fields(0).Value
    CUST_ID.Close
End If


End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Combo1.Text) > 0 Then
If KeyCode = 13 Then
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
End If
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Dim x As Integer
    x = MsgBox("Are you sure you want to Cancel this Invoice ...", vbQuestion Or vbYesNo, "Want to cancel this invoice ...")
    If x = 6 Then
    Unload Me
    End If
    
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Left = 0
Me.Top = 0
Text1(0).Enabled = False

Label4.Visible = False
Text1(2).Text = SALES_INVOICE_NUMBER()
SendKeys "{TAB}"
Label3(0).Caption = SYSTEM_NO()
    
    LaVolpeButton2.Enabled = False
    FNAME = Clear
    Status = Clear
    rs_custo_details.Open "SELECT * FROM Customer_master", db, adOpenDynamic, adLockOptimistic
    REFRESH_COMBO (1)
End Sub

Public Sub REFRESH_COMBO(Index As Integer)
If Index = 1 Then
        Combo1.Clear
        rs_custo_details.Requery
        While rs_custo_details.EOF <> True
            Combo1.AddItem rs_custo_details.Fields(1).Value
            rs_custo_details.MoveNext
        Wend
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    rs_custo_details.Close
    If Status = "NOT SAVED" Then
        rs_custo_details.Open "SELECT * FROM Customer_master WHERE cutomer_name='" & Combo1.Text & "' AND cutomer_id='" & Text1(0).Text & "'", db, adOpenDynamic, adLockOptimistic
        rs_custo_details.Delete
        rs_custo_details.Close
    End If
    
End Sub

Private Sub issues_Click()

End Sub

Private Sub LaVolpeButton1_Click()
            FNAME = "CUST_SYS_FORM"
            Status = "NOT SAVED"
            Combo1.Enabled = False
            LaVolpeButton1.Enabled = False
            LaVolpeButton2.Enabled = False
            frm_cust_details.Show
End Sub

Private Sub LaVolpeButton2_Click()
        CrystalReport1.DataFiles(0) = App.Path & "\Master_Database.mdb"
        CrystalReport1.ReportFileName = App.Path & "\Report\rpt_Verify_cutomer_detail.rpt"
        CrystalReport1.SelectionFormula = "{Customer_master.cutomer_name} = '" & Combo1.Text & "'"
        CrystalReport1.username = "Admin"
        CrystalReport1.Password = "1010101010" & Chr(10) & "1010101010"
        CrystalReport1.Action = 1
End Sub

Private Sub LaVolpeButton3_Click()
If Len(Combo1.Text) > 0 Then
Status = "SAVED"
Dim CSNO As New ADODB.Recordset
CSNO.Open "SELECT * FROM CUSTOMER_SYSTEM_INVOICENO", db, adOpenDynamic, adLockOptimistic



    CSNO.AddNew
    CSNO.Fields(0).Value = Text1(0).Text
    CSNO.Fields(1).Value = Combo1.Text
    CSNO.Fields(2).Value = Text1(2).Text
    CSNO.Fields(3).Value = Text1(1).Text
    CSNO.Fields(4).Value = False
    CSNO.Update
    CSNO.Close
    
    
    Dim STR As String
    STR = Label3(0).Caption
    
    Static NO As Integer
    
    
    
    Dim C_S As String
    NO = Mid(STR, 2, Len(STR))
    
    For i = 0 To Text1(1).Text - 1
        
        Dim SSS As String
        SSS = NO
        If Len(SSS) = 1 Then
            C_S = "S0000" & NO
        ElseIf Len(SSS) = 2 Then
            C_S = "S000" & NO
        ElseIf Len(SSS) = 3 Then
            C_S = "S00" & NO
        ElseIf Len(SSS) = 4 Then
            C_S = "S0" & NO
        ElseIf Len(SSS) = 5 Then
            C_S = "S" & NO
        End If
        
        NO = NO + 1
        STR = C_S
        Dim R As New ADODB.Recordset
        R.Open "SELECT * FROM INVOICE_NUMBER_SYSTEM_ID", db, adOpenDynamic, adLockOptimistic
        R.AddNew
        R.Fields(0).Value = C_S
        R.Fields(1).Value = Text1(2).Text
        R.Update
        R.Close
    Next
    
    
    Load sales_form
    sales_form.Combo1 = Combo1.Text
    sales_form.Text1(4).Text = Text1(0).Text
    sales_form.SALE_TYPE = "SYSTEM"
    sales_form.SYSTEM_QTY = VAL(Text1(1).Text)
    sales_form.Text1(0).Enabled = False
    sales_form.Combo1.Enabled = False
    sales_form.Text1(4).Enabled = False
    sales_form.LaVolpeButton1.Enabled = False
    sales_form.LaVolpeButton2.Enabled = False
    sales_form.Visible = True
    
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    
    Unload Me
    Else
        MsgBox "Enter Customer name ...", vbCritical, "Customer name and ID not Found ..."
    End If
    
End Sub


Private Sub Text1_Change(Index As Integer)
If Index = 1 Then
    
    If Len(Text1(1).Text) > 0 Then
    Label4.Visible = True
    Dim ST As String
    ST = Label3(0).Caption
    
    Dim N As Integer
    N = Mid(ST, 2, Len(ST))
    N = N + VAL(Text1(1).Text)
    
    Dim NO As String
    NO = N
    
    If Len(NO) = 1 Then
        Label3(1) = "S0000" & NO
    ElseIf Len(NO) = 2 Then
        Label3(1) = "S000" & NO
    ElseIf Len(NO) = 3 Then
        Label3(1) = "S00" & NO
    ElseIf Len(NO) = 4 Then
        Label3(1) = "S0" & NO
    ElseIf Len(NO) = 5 Then
        Label3(1) = "S" & NO
    End If
    
    Else
        Label3(1).Caption = Clear
        Label4.Visible = False
    End If
    
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 1 Then
If KeyCode = 13 Then
    SendKeys "{TAB}"
End If
End If

End Sub
