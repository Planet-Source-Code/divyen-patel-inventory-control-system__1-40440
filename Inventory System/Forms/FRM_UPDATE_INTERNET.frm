VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_UPDATE_INTERNET 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Internet Connection Entry ..."
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8430
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   2415
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      ItemData        =   "FRM_UPDATE_INTERNET.frx":0000
      Left            =   4200
      List            =   "FRM_UPDATE_INTERNET.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   3855
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   4
      Left            =   4200
      TabIndex        =   7
      Top             =   4080
      Width           =   4095
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   11
      Top             =   5760
      Width           =   1335
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   4200
      TabIndex        =   8
      Top             =   4440
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   1920
      TabIndex        =   13
      Top             =   2280
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Select Package For Sale"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Invoice_no"
         Caption         =   "Invoice No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Item_type"
         Caption         =   "Item Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Item_name"
         Caption         =   "Item Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Party_name"
         Caption         =   "Party Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Purchase_Date"
         Caption         =   "Purchase Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Qty"
         Caption         =   "Qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "price_per_unit"
         Caption         =   "Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "total_amt"
         Caption         =   "Total Amt"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Item_Description"
         Caption         =   "Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   435
      Left            =   6600
      TabIndex        =   12
      Top             =   5760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "&Update"
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
      MICON           =   "FRM_UPDATE_INTERNET.frx":0004
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
      Left            =   7680
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   315
      Left            =   7560
      TabIndex        =   2
      Top             =   600
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
      MICON           =   "FRM_UPDATE_INTERNET.frx":0020
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
      Left            =   6120
      TabIndex        =   3
      Top             =   960
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
      MICON           =   "FRM_UPDATE_INTERNET.frx":003C
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   47710211
      CurrentDate     =   37463
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   47710211
      CurrentDate     =   37463
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
      Index           =   1
      Left            =   1920
      TabIndex        =   23
      Top             =   240
      Width           =   2175
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
      Index           =   0
      Left            =   1920
      TabIndex        =   22
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchased From"
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
      Left            =   1920
      TabIndex        =   21
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchased Price"
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
      Left            =   1920
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Package Detail"
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
      Left            =   1920
      TabIndex        =   19
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Serial Number"
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
      Left            =   1920
      TabIndex        =   18
      Top             =   4080
      Width           =   2175
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
      Index           =   8
      Left            =   1920
      TabIndex        =   17
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Billing Price"
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
      Left            =   1920
      TabIndex        =   16
      Top             =   5640
      Width           =   2175
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
      Index           =   6
      Left            =   1920
      TabIndex        =   15
      Top             =   4440
      Width           =   2175
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
      Left            =   1920
      TabIndex        =   14
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   0
      Picture         =   "FRM_UPDATE_INTERNET.frx":0058
      Top             =   0
      Width           =   1965
   End
End
Attribute VB_Name = "FRM_UPDATE_INTERNET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cust_names As New ADODB.Recordset
Public Status As Boolean
Public FNAME As String
Dim R As New ADODB.Recordset
Dim RS_PR As New ADODB.Recordset
Dim dtsource As New ADODB.Recordset
Dim INV As String
Public inno As String





Private Sub Combo1_Click()
If Len(Combo1.Text) = 0 Then
    LaVolpeButton2.Enabled = False
Else
    LaVolpeButton2.Enabled = True
    Dim R As New ADODB.Recordset
    R.Open "select cutomer_id from Customer_master where cutomer_name='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
    Text1(0).Text = R.Fields(0).Value
    R.Close
End If


End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
LaVolpeButton2.Enabled = False
Text1(0).Text = Clear
End If

If KeyCode = 13 Then
    If Len(Combo1.Text) > 0 Then
        SendKeys "{TAB}"
        SendKeys "{TAB}"
        SendKeys "{TAB}"
    End If
    
End If


End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
    
    Text1(1).Text = Clear
    Text1(2).Text = Clear
    Text1(3).Text = Clear
    
    DTPicker1(0).Value = Date
    DTPicker1(1).Value = Date
    

    


Dim pf As New ADODB.Recordset
pf.Open "SELECT DISTINCT Party_name FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='Internet Connection' AND Item_name='" & Combo2.Text & "'", db, adOpenDynamic, adLockOptimistic
Combo3.Clear
While pf.EOF <> True
    Combo3.AddItem pf.Fields(0).Value
    pf.MoveNext
Wend
pf.Close
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    Combo2.Text = Clear
    Combo3.Text = Clear
    Text1(1).Text = Clear
    Text1(2).Enabled = False
    DTPicker1(0).Value = Format(Date, "dd-MM-yyyy")
    DTPicker1(1).Value = Format(Date, "dd-MM-yyyy")
    
    DTPicker1(0).Enabled = False
    DTPicker1(1).Enabled = False
    Text1(3).Enabled = False
End If

If KeyCode = 13 Then
    If Len(Combo2.Text) > 0 Then
        SendKeys "{TAB}"
    End If
End If


End Sub

Private Sub Combo3_Click()

Dim price As New ADODB.Recordset
price.Open "SELECT price_per_unit,Invoice_no from SYS_QRY_INTERNET_SALE where Item_name='" & Combo2.Text & "' and Party_name='" & Combo3.Text & "'", db, adOpenDynamic, adLockOptimistic
Text1(1).Text = price.Fields(0).Value
INV = price.Fields(1).Value
price.Close


Set DataGrid1.DataSource = Nothing
dtsource.Close
dtsource.Open "SELECT * from SYS_QRY_INTERNET_SALE where Item_name='" & Combo2.Text & "' and Party_name='" & Combo3.Text & "'", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = dtsource
End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Combo3.Text) > 0 Then
        SendKeys "{TAB}"
        SendKeys "{TAB}"
    End If
End If

End Sub

Private Sub DataGrid1_Click()
Text1(1).Text = dtsource.Fields(6).Value
INV = dtsource.Fields(0).Value
End Sub


Private Sub Form_Load()
Me.Left = 0
Me.Top = 0


dtsource.CursorLocation = adUseClient
dtsource.Open "SELECT * from SYS_QRY_INTERNET_SALE where Item_name='" & Combo2.Text & "' and Party_name='" & Combo3.Text & "'", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = dtsource

    
Status = False
FNAME = Clear

Call REFRESH_NAMES
LaVolpeButton2.Enabled = False
RS_PR.Open "SELECT DISTINCT Item_name FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='Internet Connection' and Qty >0", db, adOpenDynamic, adLockOptimistic
While RS_PR.EOF <> True
    Combo2.AddItem RS_PR.Fields(0).Value
    RS_PR.MoveNext
Wend
RS_PR.Close

End Sub

Public Sub REFRESH_NAMES()
Combo1.Clear
cust_names.Open "SELECT cutomer_name FROM Customer_master", db, adOpenDynamic, adLockOptimistic
While cust_names.EOF <> True
    Combo1.AddItem cust_names.Fields(0).Value
    cust_names.MoveNext
Wend
cust_names.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
dtsource.Close
If Status = True Then
    R.Open "SELECT * FROM Customer_master WHERE cutomer_name='" & Combo1.Text & "' AND cutomer_id='" & Text1(0).Text & "'", db, adOpenDynamic, adLockOptimistic
    R.Delete
    R.Close
End If

End Sub

Private Sub LaVolpeButton1_Click()
Status = True
FNAME = "INTERNET"
LaVolpeButton1.Enabled = False
LaVolpeButton2.Enabled = False
Combo1.Enabled = False
frm_cust_details.Show
End Sub

Private Sub LaVolpeButton2_Click()
CrystalReport1.DataFiles(0) = App.Path & "\Master_Database.MDB"
CrystalReport1.ReportFileName = App.Path & "\Report\rpt_Verify_cutomer_detail.rpt"
CrystalReport1.SelectionFormula = "{Customer_master.cutomer_name} = '" & Combo1.Text & "'"
CrystalReport1.username = "Admin"
CrystalReport1.Password = "1010101010" & Chr(10) & "1010101010"
CrystalReport1.Action = 1
CrystalReport1.PageZoom (100)

End Sub

Private Sub LaVolpeButton3_Click()
If DTPicker1(1).Value = Format(Date, "dd-MM-yyyy") Then
    MsgBox "You can not set expiry date as todays date...", vbCritical, "Invalid Expiry Date..."
    Exit Sub
End If

If DTPicker1(0).Value < DTPicker1(1).Value Then
    If VAL(Text1(3).Text) > 0 Then
        If Len(Text1(0).Text) > 0 Then
        
            If Len(Text1(1).Text) > 0 Then
                
                Dim rs_p As New ADODB.Recordset
                
                rs_p.Open "SELECT * FROM DATE_PROFIT", db, adOpenDynamic, adLockOptimistic
                rs_p.AddNew
                rs_p.Fields(0).Value = Date
                
                rs_p.Fields(1).Value = VAL(Text1(3).Text) - VAL(Text1(1).Text)
                

                rs_p.Fields(2).Value = inno
                rs_p.Update
                rs_p.Close
                
                Dim rs_update As New ADODB.Recordset
                rs_update.Open "select * from INTERNET_CONNECTIONS", db, adOpenDynamic, adLockOptimistic
                rs_update.AddNew
                rs_update.Fields(0).Value = Text1(0).Text
                rs_update.Fields(1).Value = Combo2.Text
                rs_update.Fields(2).Value = Text1(2).Text
                rs_update.Fields(3).Value = Format(DTPicker1(0).Value, "mm-dd-yyyy")
                rs_update.Fields(4).Value = Format(DTPicker1(1).Value, "mm-dd-yyyy")
                rs_update.Fields(5).Value = Text1(3).Text
                rs_update.Fields(6).Value = Month(DTPicker1(1).Value)
                rs_update.Fields(7).Value = Year(DTPicker1(1).Value)
                
                     If Len(Text1(4).Text) > 0 Then
                    rs_update.Fields(8).Value = Text1(4).Text
                End If
                rs_update.Fields(9).Value = True
                ''''
                rs_update.Fields(10).Value = dtsource.Fields(0).Value
                rs_update.Fields(11).Value = dtsource.Fields(3).Value
                rs_update.Fields(12).Value = dtsource.Fields(6).Value
                rs_update.Fields(13).Value = dtsource.Fields(4).Value
                rs_update.Fields(14).Value = inno
                
               'On Error GoTo a1:
                rs_update.Update
                Status = False
                Dim up_stock As New ADODB.Recordset
                up_stock.Open "SELECT Qty FROM AVAILABLE_PURCHASED_STOCK WHERE Party_name='" & Combo3.Text & "' AND Item_type='Internet Connection' AND Item_name='" & Combo2.Text & "' AND Invoice_no='" & INV & "'", db, adOpenDynamic, adLockOptimistic
                up_stock.Fields(0).Value = up_stock.Fields(0).Value - 1
                
                up_stock.Update
                If up_stock.Fields(0).Value = 0 Then
                    up_stock.Delete
                End If
                
                up_stock.Close
                
                Dim SALES_UPDATE As New ADODB.Recordset
                SALES_UPDATE.Open "select * from Sales_master", db, adOpenDynamic, adLockOptimistic
                SALES_UPDATE.AddNew
                SALES_UPDATE.Fields(0).Value = inno
                SALES_UPDATE.Fields(1).Value = Text1(0).Text
                SALES_UPDATE.Fields(2).Value = Combo1.Text
                SALES_UPDATE.Fields(3).Value = Format(DTPicker1(0).Value, "mm-dd-yyyy")
                SALES_UPDATE.Fields(4).Value = "Internet Connection"
                SALES_UPDATE.Fields(5).Value = Combo2.Text
                SALES_UPDATE.Fields(6).Value = 1
                SALES_UPDATE.Fields(7).Value = Text1(3).Text
                SALES_UPDATE.Fields(8).Value = dtsource.Fields(0).Value
                
                SALES_UPDATE.Fields(9).Value = dtsource.Fields(3).Value
                SALES_UPDATE.Fields(10).Value = dtsource.Fields(4).Value
                SALES_UPDATE.Fields(11).Value = dtsource.Fields(6).Value
                
                
                SALES_UPDATE.Update
                SALES_UPDATE.Close
                
                Dim up_date_item_master As New ADODB.Recordset
                up_date_item_master.Open "select * from Item_master where Item_name='" & Combo2.Text & "' and Itemtype='Internet Connection'", db, adOpenKeyset, adLockOptimistic
                up_date_item_master.Fields(3).Value = up_date_item_master.Fields(3).Value - 1
                up_date_item_master.Update
                
                 FRM_AMT_PAID_NOT_PAID.Label3(5).Caption = "Sales"
                 FRM_AMT_PAID_NOT_PAID.Label3(2).Caption = inno
                 FRM_AMT_PAID_NOT_PAID.Label3(0).Caption = Combo1.Text
                 FRM_AMT_PAID_NOT_PAID.Label2(2).Caption = Text1(3).Text
                 Me.Visible = False
                 
                 FRM_AMT_PAID_NOT_PAID.dt = Format(DTPicker1(0).Value, "dd-MMM-yyyy")
                 FRM_AMT_PAID_NOT_PAID.Show vbModal
                '  Dim f As New FileSystemObject
                ' f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True
       
                 Unload Me
                 Exit Sub
A1:
                MsgBox "Duplicate Entry Found ...", vbCritical, "Check your data ..."
                Exit Sub
            Else
                MsgBox "Enter Proper Data ...", vbCritical, "Insufficient Data ..."
                
                Exit Sub
            End If
            
        Else
        MsgBox "Select Customer Name ...", vbCritical, "Name Not Found ..."
        Exit Sub
        End If
        
    Else
        MsgBox "Billing Price can not be Zero Value", vbCritical, "Enter Proper Data ..."
        Exit Sub
    End If
    
Else
    MsgBox "Registered Date Must be less than Expire Date ...", vbCritical, "Enter Proper Data ..."
    Exit Sub
End If



End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 1 Then
If Len(Text1(1).Text) > 0 Then
        Text1(2).Enabled = True
        DTPicker1(0).Enabled = True
        DTPicker1(1).Enabled = True
        Text1(3).Enabled = True
    End If
End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Len(Text1(Index).Text) > 0 Then
    SendKeys "{TAB}"
End If

End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
            
        ElseIf KeyAscii = 46 Then
            If InStr(1, Text1(3).Text, ".", vbTextCompare) > 0 Then
                KeyAscii = 0
            End If
        ElseIf KeyAscii = 8 Then
        
        Else
            KeyAscii = 0
        End If
    End If
End Sub
