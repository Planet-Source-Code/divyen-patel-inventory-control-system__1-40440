VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_cust_details 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Details"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frm_cust_details.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7215
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Name / Address"
      TabPicture(0)   =   "frm_cust_details.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(11)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text1(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1(11)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Contact Numbers"
      TabPicture(1)   =   "frm_cust_details.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(7)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(8)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(9)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text1(6)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text1(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1(8)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text1(9)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "E-mail Address"
      TabPicture(2)   =   "frm_cust_details.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(10)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(10)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
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
         Height          =   315
         Index           =   11
         Left            =   2280
         TabIndex        =   3
         Top             =   1440
         Width           =   3615
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
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   5
         Top             =   2160
         Width           =   3615
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
         Height          =   360
         Index           =   10
         Left            =   -72720
         TabIndex        =   14
         Top             =   1440
         Width           =   2895
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
         Height          =   360
         Index           =   9
         Left            =   -72720
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1800
         Width           =   3495
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
         Height          =   360
         Index           =   8
         Left            =   -72720
         MaxLength       =   13
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
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
         Height          =   360
         Index           =   7
         Left            =   -72720
         MaxLength       =   13
         TabIndex        =   10
         Top             =   1080
         Width           =   2655
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
         Height          =   360
         Index           =   6
         Left            =   -72720
         MaxLength       =   8
         TabIndex        =   9
         Top             =   720
         Width           =   1815
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
         Height          =   330
         Index           =   5
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
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
         Height          =   315
         Index           =   4
         Left            =   2280
         TabIndex        =   6
         Top             =   2520
         Width           =   1575
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
         Height          =   330
         Index           =   2
         Left            =   2280
         TabIndex        =   4
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   0
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   1815
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
         Height          =   330
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Index           =   11
         Left            =   360
         TabIndex        =   25
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Address"
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
         Index           =   10
         Left            =   -74640
         TabIndex        =   24
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
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
         Index           =   9
         Left            =   -74640
         TabIndex        =   23
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number 2"
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
         Index           =   8
         Left            =   -74640
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number 1"
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
         Index           =   7
         Left            =   -74640
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "STD Code"
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
         Index           =   6
         Left            =   -74640
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
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
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2"
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
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MICON           =   "frm_cust_details.frx":0E96
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
Attribute VB_Name = "frm_cust_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cust_rs As New ADODB.Recordset
Dim ST As String

Private Sub cmd_Click()
ST = "saved"

    cust_rs.AddNew
    Dim i As Integer
    
    For i = 0 To Text1.Count - 1
        If Len(Text1(i).Text) <> 0 Then
            cust_rs.Fields(i).Value = Text1(i).Text
        End If
        
    Next
    
On Error GoTo A1:
    cust_rs.Update
    
    
If sales_form.FNAME = "SALES_NAME" Then
    Call sales_form.REFRESH_COMBO(1)
    sales_form.Combo1.Text = Text1(1).Text
    sales_form.Text1(4).Text = Text1(0).Text
    Unload Me
    
    SendKeys "{TAB}"
    

    Exit Sub
End If

If FRM_CUST_SYSTEM_DETAILS.FNAME = "CUST_SYS_FORM" Then
    Call FRM_CUST_SYSTEM_DETAILS.REFRESH_COMBO(1)
    FRM_CUST_SYSTEM_DETAILS.Combo1.Text = Text1(1).Text
    FRM_CUST_SYSTEM_DETAILS.Text1(0).Text = Text1(0).Text
End If

If INTER_NET_CONNECTIONS.FNAME = "INTERNET" Then
    Call INTER_NET_CONNECTIONS.REFRESH_NAMES
    INTER_NET_CONNECTIONS.Text1(0).Text = Text1(0).Text
    INTER_NET_CONNECTIONS.Combo1.Text = Text1(1).Text
End If
    Unload Me
    Exit Sub
A1:
    MsgBox "Duplicate Customer Name Found ..." & vbCrLf & "Add Relative Keywords to the name ...", vbCritical, "Duplicate Entry Found ..."
    cust_rs.CancelUpdate
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True

SendKeys "{TAB}"
If sales_form.FNAME = "SALES_NAME" Then
    Me.Left = sales_form.Width / 2 - Me.Width / 2
    Me.Top = sales_form.Height / 2 - Me.Height / 2
ElseIf INTER_NET_CONNECTIONS.FNAME = "INTERNET" Then
    Me.Left = INTER_NET_CONNECTIONS.Width / 2 - Me.Width / 2
    Me.Top = INTER_NET_CONNECTIONS.Height / 2 - Me.Height / 2
End If


ST = "notsaved"

cust_rs.Open "select * from SORTED_CUST_NO", db, adOpenDynamic, adLockOptimistic

If cust_rs.EOF <> True Then

    cust_rs.MoveLast
    Dim LAST_ID As String
    LAST_ID = Mid(cust_rs.Fields(0).Value, 2, Len(cust_rs.Fields(0).Value))
    LAST_ID = VAL(LAST_ID) + 1
    
    If Len(LAST_ID) = 1 Then
        Text1(0).Text = "C00000000" & LAST_ID
    ElseIf Len(LAST_ID) = 2 Then
        Text1(0).Text = "C0000000" & LAST_ID
    ElseIf Len(LAST_ID) = 3 Then
        Text1(0).Text = "C000000" & LAST_ID
    ElseIf Len(LAST_ID) = 4 Then
        Text1(0).Text = "C00000" & LAST_ID
    ElseIf Len(LAST_ID) = 5 Then
        Text1(0).Text = "C0000" & LAST_ID
    ElseIf Len(LAST_ID) = 6 Then
        Text1(0).Text = "C000" & LAST_ID
    ElseIf Len(LAST_ID) = 7 Then
        Text1(0).Text = "C00" & LAST_ID
    ElseIf Len(LAST_ID) = 8 Then
        Text1(0).Text = "C0" & LAST_ID
    ElseIf Len(LAST_ID) = 9 Then
        Text1(0).Text = "C" & LAST_ID
    End If
Else
    Text1(0).Text = "C000000001"
    
End If
Text1(0).Enabled = False

cust_rs.Requery
cust_rs.Close
cust_rs.Open "SELECT * FROM Customer_master", db, adOpenDynamic, adLockOptimistic


End Sub

Private Sub Form_Unload(Cancel As Integer)
cust_rs.Close

If ST = "notsaved" And sales_form.FNAME = "SALES_NAME" Then
    sales_form.Combo1.Enabled = True
    sales_form.LaVolpeButton1.Enabled = True
    If Len(sales_form.Combo1.Text) = 0 Then
        sales_form.LaVolpeButton2.Enabled = False
    Else
    
        sales_form.LaVolpeButton2.Enabled = True
    End If
    
    sales_form.SNAME = Clear
    sales_form.Text1(4).Text = Clear
    sales_form.FNAME = Clear
    Exit Sub
End If

If ST = "notsaved" And FRM_CUST_SYSTEM_DETAILS.FNAME = "CUST_SYS_FORM" Then
    FRM_CUST_SYSTEM_DETAILS.Combo1.Enabled = True
    FRM_CUST_SYSTEM_DETAILS.Combo1.Text = Clear
    FRM_CUST_SYSTEM_DETAILS.Text1(0).Text = Clear
    FRM_CUST_SYSTEM_DETAILS.LaVolpeButton1.Enabled = True
    
    If Len(FRM_CUST_SYSTEM_DETAILS.Combo1.Text) = 0 Then
        FRM_CUST_SYSTEM_DETAILS.LaVolpeButton2.Enabled = False
    Else
        FRM_CUST_SYSTEM_DETAILS.LaVolpeButton2.Enabled = True
    End If
    
    
    FRM_CUST_SYSTEM_DETAILS.FNAME = Clear
    FRM_CUST_SYSTEM_DETAILS.Status = Clear
End If

If ST = "notsaved" And INTER_NET_CONNECTIONS.FNAME = "INTERNET" Then
    INTER_NET_CONNECTIONS.Status = False
    INTER_NET_CONNECTIONS.Text1(0).Text = Clear
    INTER_NET_CONNECTIONS.Combo1.Text = Clear
    INTER_NET_CONNECTIONS.LaVolpeButton1.Enabled = True
    INTER_NET_CONNECTIONS.LaVolpeButton2.Enabled = True
    INTER_NET_CONNECTIONS.Combo1.Enabled = True

End If



End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    If Index = 5 Or Index = 9 Then
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            
            SendKeys "{RIGHT}"
            SendKeys "{TAB}"
    Else
    SendKeys "{TAB}"
    End If
    
End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
  If KeyAscii >= 97 And KeyAscii <= 122 Then
        ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then
        ElseIf KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
  End If
ElseIf Index = 5 Or Index = 6 Or Index = 7 Or Index = 8 Or Index = 9 Then

        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        ElseIf KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

End If

End Sub
