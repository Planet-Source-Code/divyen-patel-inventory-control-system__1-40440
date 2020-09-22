VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_VIEW_PURCHASE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Purchase Entries ..."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
   Icon            =   "FRM_VIEW_PURCHASE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10755
   Begin VB.ComboBox Combo1 
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
      Index           =   0
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.ComboBox Combo1 
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
      Index           =   1
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   20
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
      Caption         =   "Purchase Book Entry"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Invoice_no"
         Caption         =   "Invoice Number"
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
      BeginProperty Column02 
         DataField       =   "Purchase_Date"
         Caption         =   "Date"
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
      BeginProperty Column04 
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
         Caption         =   "Price Per Unit"
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
         Caption         =   "Total Amount"
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
         Caption         =   "Item Description"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   -1  'True
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2700.284
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   9960
      Picture         =   "FRM_VIEW_PURCHASE.frx":0E42
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Party Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "FRM_VIEW_PURCHASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_data As New ADODB.Recordset

Dim p_name As New ADODB.Recordset




Private Sub Combo1_Click(Index As Integer)
If Index = 0 Then
    REFRESH_INVOICE_NO
    GETDATA
ElseIf Index = 1 Then
    GETDATA
    Text1.Text = rs_data.Fields(2).Value
End If

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Combo1(Index).Text) > 0 Then
        SendKeys "{TAB}"
    End If
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
    
    p_name.Open "select DISTINCT Party_name FROM Purchase_master", db, adOpenKeyset, adLockOptimistic
    
    If p_name.RecordCount = 0 Then
        Unload Me
        MsgBox "No Record Found ...", vbInformation, "No Record Found ..."
        Exit Sub
    End If
    
    While p_name.EOF <> True
        Combo1(0).AddItem p_name.Fields(0).Value
        p_name.MoveNext
    Wend
    
    
    
End Sub

Public Sub REFRESH_INVOICE_NO()
    Dim RS_INVOICE As New ADODB.Recordset
    RS_INVOICE.Open "select distinct Invoice_no from Purchase_master where Party_name='" & Combo1(0).Text & "'", db, adOpenKeyset, adLockOptimistic
    
    Combo1(1).Clear
    While RS_INVOICE.EOF <> True
        Combo1(1).AddItem RS_INVOICE.Fields(0).Value
        RS_INVOICE.MoveNext
    Wend
    
End Sub

Public Sub GETDATA()
If rs_data.State = adStateOpen Then
    rs_data.Close
End If
rs_data.CursorLocation = adUseClient
rs_data.Open "SELECT * FROM Purchase_master WHERE Party_name='" & Combo1(0).Text & "' AND Invoice_no='" & Combo1(1).Text & "' ORDER BY Invoice_no,Party_name,Purchase_Date,Item_type,Item_name,Qty,price_per_unit,total_amt,Item_Description", db, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = rs_data
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    rs_data.Close
    p_name.Close
End Sub

