VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_UPDATE_SALES 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Invoice to Modify ..."
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "FRM_UPDATE_SALES.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8040
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Modify"
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
      MICON           =   "FRM_UPDATE_SALES.frx":0E42
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Sales Bill Details"
      ColumnCount     =   11
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
         DataField       =   "Custo_id"
         Caption         =   "Customer Id"
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
         DataField       =   "Party_name"
         Caption         =   "Customer Name"
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
         DataField       =   "Sales_date"
         Caption         =   "Sales Date"
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
         DataField       =   "Item_type"
         Caption         =   "Item type"
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
         DataField       =   "Item_name"
         Caption         =   "Item name"
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
      BeginProperty Column07 
         DataField       =   "price_per_unit"
         Caption         =   "Rete Per Unit"
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
         DataField       =   "P_INVOICE_NO"
         Caption         =   "Purchase Invoice Number"
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
      BeginProperty Column09 
         DataField       =   "P_PARTY_NAME"
         Caption         =   "Purchase Party Name"
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
      BeginProperty Column10 
         DataField       =   "P_DATE"
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
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2415
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
      Index           =   0
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Delete"
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
      MICON           =   "FRM_UPDATE_SALES.frx":0E5E
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
      Caption         =   "Sales Date"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice Number"
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
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1575
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FRM_UPDATE_SALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cust_names As New ADODB.Recordset
Dim rs_data As New ADODB.Recordset


Private Sub Combo1_Click(Index As Integer)
If Index = 0 Then
    Dim invoice As New ADODB.Recordset
    invoice.Open "select distinct Invoice_no from Sales_master where Party_name='" & Combo1(0).Text & "'", db, adOpenKeyset, adLockOptimistic
    Combo1(1).Clear
    
    While invoice.EOF <> True
        Combo1(1).AddItem invoice.Fields(0).Value
        invoice.MoveNext
    Wend
    
ElseIf Index = 1 Then
    GETDATA
    Text1.Text = rs_data.Fields(3).Value
End If
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Combo1(Index)) > 0 Then
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
cust_names.Open "select distinct Party_name from Sales_master", db, adOpenKeyset, adLockOptimistic

If cust_names.RecordCount = 0 Then
    MsgBox "No Record Found ...", vbInformation, "No Record Found ..."
    Unload Me
    Exit Sub
End If

Combo1(0).Clear

While cust_names.EOF <> True
    Combo1(0).AddItem cust_names.Fields(0).Value
    cust_names.MoveNext
Wend


End Sub

Public Sub GETDATA()
If rs_data.State <> adStateClosed Then
    rs_data.Close
End If
rs_data.CursorLocation = adUseClient
rs_data.Open "select * from Sales_master where Party_name='" & Combo1(0).Text & "' and Invoice_no='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic
Set DataGrid1.DataSource = rs_data
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
cust_names.Close
rs_data.Close
Exit Sub
End Sub

Private Sub LaVolpeButton1_Click()

Dim CHECK_SALES_TYPE As New ADODB.Recordset
CHECK_SALES_TYPE.Open "SELECT * FROM Sales_master WHERE Party_name='" & Combo1(0).Text & "' AND Invoice_no='" & Combo1(1).Text & "' AND Item_type='Internet Connection'", db, adOpenKeyset, adLockOptimistic

If CHECK_SALES_TYPE.RecordCount > 0 Then
    MsgBox "This Sales Bill is of Internet Connection Sale ..." & vbCrLf & "If you want to change Intetnet Connection Bill then use Modify -> Internet Connection Sale Menu ...", vbInformation, "Internet Connection Sales can not be Modified here ..."
    CHECK_SALES_TYPE.Close
    Exit Sub
End If

Dim CR_UPDATE As New ADODB.Recordset
CR_UPDATE.Open "SELECT * FROM CRITICAL_SALES_DATA", db, adOpenKeyset, adLockOptimistic

While CR_UPDATE.EOF <> True
    CR_UPDATE.Delete
    CR_UPDATE.MoveNext
Wend

rs_data.Requery

Dim UPDATA_CUR_RECORD As New ADODB.Recordset
UPDATA_CUR_RECORD.Open "select * from SYS_CURRENT_SALES_ITEMS", db, adOpenKeyset, adLockOptimistic

While UPDATA_CUR_RECORD.EOF <> True
    UPDATA_CUR_RECORD.Delete
    UPDATA_CUR_RECORD.MoveNext
Wend



While rs_data.EOF <> True
    CR_UPDATE.AddNew
    
    For i = 0 To 11
        CR_UPDATE.Fields(i).Value = rs_data.Fields(i).Value
    Next
    CR_UPDATE.Update

    rs_data.MoveNext
Wend


CR_UPDATE.Requery

While CR_UPDATE.EOF <> True
    UPDATA_CUR_RECORD.AddNew
    UPDATA_CUR_RECORD.Fields(0).Value = CR_UPDATE.Fields(5).Value
    UPDATA_CUR_RECORD.Fields(1).Value = CR_UPDATE.Fields(6).Value
    UPDATA_CUR_RECORD.Fields(2).Value = CR_UPDATE.Fields(7).Value
    UPDATA_CUR_RECORD.Fields(3).Value = VAL(CR_UPDATE.Fields(6).Value) * VAL(CR_UPDATE.Fields(7).Value)
    UPDATA_CUR_RECORD.Fields(4).Value = CR_UPDATE.Fields(4).Value
    UPDATA_CUR_RECORD.Fields(5).Value = CR_UPDATE.Fields(9).Value
    UPDATA_CUR_RECORD.Fields(6).Value = CR_UPDATE.Fields(8).Value
    UPDATA_CUR_RECORD.Fields(7).Value = CR_UPDATE.Fields(10).Value
    UPDATA_CUR_RECORD.Fields(8).Value = CR_UPDATE.Fields(11).Value
    UPDATA_CUR_RECORD.Update
    CR_UPDATE.MoveNext
Wend

Dim ITEM_DESC As New ADODB.Recordset
ITEM_DESC.Open "SELECT * FROM sales_item_description WHERE invoice_number='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic
Dim CR_ITEM As New ADODB.Recordset
CR_ITEM.Open "SELECT * FROM CRITICAL_SALES_DESC", db, adOpenKeyset, adLockOptimistic

        While CR_ITEM.EOF <> True
            CR_ITEM.Delete
            CR_ITEM.MoveNext
        Wend

While ITEM_DESC.EOF <> True
    CR_ITEM.AddNew
    CR_ITEM.Fields(0).Value = ITEM_DESC.Fields(0).Value
    CR_ITEM.Fields(1).Value = ITEM_DESC.Fields(1).Value
    CR_ITEM.Update
    ITEM_DESC.MoveNext
Wend

ITEM_DESC.Requery

While ITEM_DESC.EOF <> True
    ITEM_DESC.Delete
    ITEM_DESC.MoveNext
Wend

CR_UPDATE.Requery

Dim check_sys As New ADODB.Recordset
check_sys.Open "select * from INVOICE_NUMBER_SYSTEM_ID WHERE INVOICE_NUMBER='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic

If check_sys.RecordCount > 0 Then
    FRM_SALES_UPDATE_FORM.SALE_TYPE = "SYSTEM"
    FRM_SALES_UPDATE_FORM.SYSTEM_QTY = check_sys.RecordCount
    While check_sys.EOF <> True
        db.Execute "DELETE FROM Customer_System_datail WHERE SYSTEM_ID='" & check_sys.Fields(0).Value & "'"
        check_sys.MoveNext
    Wend
End If


CR_UPDATE.Requery

Dim UPDATE_STOCK As New ADODB.Recordset
While CR_UPDATE.EOF <> True
    UPDATE_STOCK.Open "SELECT * FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & CR_UPDATE.Fields(4).Value & "' AND Item_name='" & CR_UPDATE.Fields(5).Value & "' AND Invoice_no='" & CR_UPDATE.Fields(8).Value & "' AND Party_name='" & CR_UPDATE.Fields(9).Value & "'", db, adOpenKeyset, adLockOptimistic
    If UPDATE_STOCK.RecordCount > 0 Then
        UPDATE_STOCK.Fields(5).Value = VAL(UPDATE_STOCK.Fields(5).Value) + VAL(CR_UPDATE.Fields(6).Value)
        UPDATE_STOCK.Fields(6).Value = VAL(CR_UPDATE.Fields(11).Value)
        UPDATE_STOCK.Fields(7).Value = VAL(UPDATE_STOCK.Fields(5).Value) * VAL(UPDATE_STOCK.Fields(6).Value)
        UPDATE_STOCK.Update
    Else
        UPDATE_STOCK.AddNew
        UPDATE_STOCK.Fields(0).Value = CR_UPDATE.Fields(8).Value
        UPDATE_STOCK.Fields(1).Value = CR_UPDATE.Fields(9).Value
        UPDATE_STOCK.Fields(2).Value = CR_UPDATE.Fields(10).Value
        UPDATE_STOCK.Fields(3).Value = CR_UPDATE.Fields(4).Value
        UPDATE_STOCK.Fields(4).Value = CR_UPDATE.Fields(5).Value
        UPDATE_STOCK.Fields(5).Value = CR_UPDATE.Fields(6).Value
        UPDATE_STOCK.Fields(6).Value = CR_UPDATE.Fields(11).Value
        UPDATE_STOCK.Fields(7).Value = VAL(CR_UPDATE.Fields(5).Value) * VAL(UPDATE_STOCK.Fields(6).Value)
        UPDATE_STOCK.Update
    End If
    CR_UPDATE.MoveNext
    UPDATE_STOCK.Close
Wend

db.Execute "DELETE FROM Sales_master WHERE Invoice_no='" & Combo1(1).Text & "'"
db.Execute "DELETE FROM DATE_PROFIT WHERE INVOICE_NUMBER='" & Combo1(1).Text & "'"


Dim del_unpaid As New ADODB.Recordset
del_unpaid.Open "select * from AMT_UNPAID_REMIND where PARTY_NAME='" & Combo1(0).Text & "' and INVOICE_NO='" & Combo1(1).Text & "' AND TRAN_TYPE='SALES'", db, adOpenKeyset, adLockOptimistic
    
If del_unpaid.RecordCount > 0 Then
        del_unpaid.Delete
End If
    
del_unpaid.Close

CR_UPDATE.Requery

While CR_UPDATE.EOF <> True
Dim R_ITEM_MASTER As New ADODB.Recordset
R_ITEM_MASTER.Open "SELECT Qty FROM Item_master WHERE Item_name='" & CR_UPDATE.Fields(5).Value & "' AND Itemtype='" & CR_UPDATE.Fields(4).Value & "'", db, adOpenKeyset, adLockOptimistic
R_ITEM_MASTER.Fields(0).Value = VAL(R_ITEM_MASTER.Fields(0).Value) + VAL(CR_UPDATE.Fields(6).Value)
R_ITEM_MASTER.Update
CR_UPDATE.MoveNext
R_ITEM_MASTER.Close
Wend

db.Execute "DELETE FROM EXPENSE WHERE INVOICE_NO='" & Combo1(1).Text & "' AND EXPENSE_TYPE='Discounted Amount(Sales)'"
db.Execute "DELETE FROM INCOME WHERE INVOICE_NUMBER='" & Combo1(1).Text & "' AND INCOME_TYPE='Discounted Amount(Sales)'"

CR_UPDATE.Requery
CR_ITEM.Requery
FRM_SALES_UPDATE_FORM.Text1(0).Text = Combo1(1).Text
FRM_SALES_UPDATE_FORM.Combo1.Text = Combo1(0).Text
FRM_SALES_UPDATE_FORM.DTPicker1.Value = Format(Text1.Text, "dd-MMM-yyyy")
FRM_SALES_UPDATE_FORM.Text1(4).Text = CR_UPDATE.Fields(1).Value
If CR_ITEM.RecordCount > 0 Then
    If Len(CR_ITEM.Fields(1).Value) > 0 Then
        FRM_SALES_UPDATE_FORM.Text2.Text = CR_ITEM.Fields(1).Value
    End If
End If

Unload Me
FRM_SALES_UPDATE_FORM.Show


End Sub

Private Sub LaVolpeButton2_Click()
Dim CHECK_SALES_TYPE As New ADODB.Recordset
CHECK_SALES_TYPE.Open "SELECT * FROM Sales_master WHERE Party_name='" & Combo1(0).Text & "' AND Invoice_no='" & Combo1(1).Text & "' AND Item_type='Internet Connection'", db, adOpenKeyset, adLockOptimistic

If CHECK_SALES_TYPE.RecordCount > 0 Then
    MsgBox "This Sales Bill is of Internet Connection Sale ..." & vbCrLf & "If you want to Delete Intetnet Connection Bill then use Modify -> Internet Connection Sale Menu ...", vbInformation, "Internet Connection Sales can not be Deleted here ..."
    CHECK_SALES_TYPE.Close
    Exit Sub
End If

Dim CR_UPDATE As New ADODB.Recordset
CR_UPDATE.Open "SELECT * FROM CRITICAL_SALES_DATA", db, adOpenKeyset, adLockOptimistic

While CR_UPDATE.EOF <> True
    CR_UPDATE.Delete
    CR_UPDATE.MoveNext
Wend

rs_data.Requery

Dim UPDATA_CUR_RECORD As New ADODB.Recordset
UPDATA_CUR_RECORD.Open "select * from SYS_CURRENT_SALES_ITEMS", db, adOpenKeyset, adLockOptimistic

While UPDATA_CUR_RECORD.EOF <> True
    UPDATA_CUR_RECORD.Delete
    UPDATA_CUR_RECORD.MoveNext
Wend



While rs_data.EOF <> True
    CR_UPDATE.AddNew
    
    For i = 0 To 11
        CR_UPDATE.Fields(i).Value = rs_data.Fields(i).Value
    Next
    CR_UPDATE.Update

    rs_data.MoveNext
Wend


CR_UPDATE.Requery

While CR_UPDATE.EOF <> True
    UPDATA_CUR_RECORD.AddNew
    UPDATA_CUR_RECORD.Fields(0).Value = CR_UPDATE.Fields(5).Value
    UPDATA_CUR_RECORD.Fields(1).Value = CR_UPDATE.Fields(6).Value
    UPDATA_CUR_RECORD.Fields(2).Value = CR_UPDATE.Fields(7).Value
    UPDATA_CUR_RECORD.Fields(3).Value = VAL(CR_UPDATE.Fields(6).Value) * VAL(CR_UPDATE.Fields(7).Value)
    UPDATA_CUR_RECORD.Fields(4).Value = CR_UPDATE.Fields(4).Value
    UPDATA_CUR_RECORD.Fields(5).Value = CR_UPDATE.Fields(9).Value
    UPDATA_CUR_RECORD.Fields(6).Value = CR_UPDATE.Fields(8).Value
    UPDATA_CUR_RECORD.Fields(7).Value = CR_UPDATE.Fields(10).Value
    UPDATA_CUR_RECORD.Fields(8).Value = CR_UPDATE.Fields(11).Value
    UPDATA_CUR_RECORD.Update
    CR_UPDATE.MoveNext
Wend

Dim ITEM_DESC As New ADODB.Recordset
ITEM_DESC.Open "SELECT * FROM sales_item_description WHERE invoice_number='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic
Dim CR_ITEM As New ADODB.Recordset
CR_ITEM.Open "SELECT * FROM CRITICAL_SALES_DESC", db, adOpenKeyset, adLockOptimistic

        While CR_ITEM.EOF <> True
            CR_ITEM.Delete
            CR_ITEM.MoveNext
        Wend

While ITEM_DESC.EOF <> True
    CR_ITEM.AddNew
    CR_ITEM.Fields(0).Value = ITEM_DESC.Fields(0).Value
    CR_ITEM.Fields(1).Value = ITEM_DESC.Fields(1).Value
    CR_ITEM.Update
    ITEM_DESC.MoveNext
Wend

ITEM_DESC.Requery

While ITEM_DESC.EOF <> True
    ITEM_DESC.Delete
    ITEM_DESC.MoveNext
Wend

CR_UPDATE.Requery

Dim check_sys As New ADODB.Recordset
check_sys.Open "select * from INVOICE_NUMBER_SYSTEM_ID WHERE INVOICE_NUMBER='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic

If check_sys.RecordCount > 0 Then
    FRM_SALES_UPDATE_FORM.SALE_TYPE = "SYSTEM"
    FRM_SALES_UPDATE_FORM.SYSTEM_QTY = check_sys.RecordCount
    While check_sys.EOF <> True
        db.Execute "DELETE FROM Customer_System_datail WHERE SYSTEM_ID='" & check_sys.Fields(0).Value & "'"
        check_sys.MoveNext
    Wend
End If


CR_UPDATE.Requery

Dim UPDATE_STOCK As New ADODB.Recordset
While CR_UPDATE.EOF <> True
    UPDATE_STOCK.Open "SELECT * FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & CR_UPDATE.Fields(4).Value & "' AND Item_name='" & CR_UPDATE.Fields(5).Value & "' AND Invoice_no='" & CR_UPDATE.Fields(8).Value & "' AND Party_name='" & CR_UPDATE.Fields(9).Value & "'", db, adOpenKeyset, adLockOptimistic
    If UPDATE_STOCK.RecordCount > 0 Then
        UPDATE_STOCK.Fields(5).Value = VAL(UPDATE_STOCK.Fields(5).Value) + VAL(CR_UPDATE.Fields(6).Value)
        UPDATE_STOCK.Fields(6).Value = VAL(CR_UPDATE.Fields(11).Value)
        UPDATE_STOCK.Fields(7).Value = VAL(UPDATE_STOCK.Fields(5).Value) * VAL(UPDATE_STOCK.Fields(6).Value)
        UPDATE_STOCK.Update
    Else
        UPDATE_STOCK.AddNew
        UPDATE_STOCK.Fields(0).Value = CR_UPDATE.Fields(8).Value
        UPDATE_STOCK.Fields(1).Value = CR_UPDATE.Fields(9).Value
        UPDATE_STOCK.Fields(2).Value = CR_UPDATE.Fields(10).Value
        UPDATE_STOCK.Fields(3).Value = CR_UPDATE.Fields(4).Value
        UPDATE_STOCK.Fields(4).Value = CR_UPDATE.Fields(5).Value
        UPDATE_STOCK.Fields(5).Value = CR_UPDATE.Fields(6).Value
        UPDATE_STOCK.Fields(6).Value = CR_UPDATE.Fields(11).Value
        UPDATE_STOCK.Fields(7).Value = VAL(CR_UPDATE.Fields(5).Value) * VAL(UPDATE_STOCK.Fields(6).Value)
        UPDATE_STOCK.Update
    End If
    CR_UPDATE.MoveNext
    UPDATE_STOCK.Close
Wend

db.Execute "DELETE FROM Sales_master WHERE Invoice_no='" & Combo1(1).Text & "'"
db.Execute "DELETE FROM DATE_PROFIT WHERE INVOICE_NUMBER='" & Combo1(1).Text & "'"


Dim del_unpaid As New ADODB.Recordset
del_unpaid.Open "select * from AMT_UNPAID_REMIND where PARTY_NAME='" & Combo1(0).Text & "' and INVOICE_NO='" & Combo1(1).Text & "' AND TRAN_TYPE='SALES'", db, adOpenKeyset, adLockOptimistic
    
If del_unpaid.RecordCount > 0 Then
        del_unpaid.Delete
End If
    
del_unpaid.Close

CR_UPDATE.Requery

While CR_UPDATE.EOF <> True
Dim R_ITEM_MASTER As New ADODB.Recordset
R_ITEM_MASTER.Open "SELECT Qty FROM Item_master WHERE Item_name='" & CR_UPDATE.Fields(5).Value & "' AND Itemtype='" & CR_UPDATE.Fields(4).Value & "'", db, adOpenKeyset, adLockOptimistic
R_ITEM_MASTER.Fields(0).Value = VAL(R_ITEM_MASTER.Fields(0).Value) + VAL(CR_UPDATE.Fields(6).Value)
R_ITEM_MASTER.Update
CR_UPDATE.MoveNext
R_ITEM_MASTER.Close
Wend

db.Execute "DELETE FROM EXPENSE WHERE INVOICE_NO='" & Combo1(1).Text & "' AND EXPENSE_TYPE='Discounted Amount(Sales)'"
db.Execute "DELETE FROM INCOME WHERE INVOICE_NUMBER='" & Combo1(1).Text & "' AND INCOME_TYPE='Discounted Amount(Sales)'"

MsgBox "Sales Bill Deleted Successfully...", vbInformation, "Invoice Deleted ..."
Unload Me

End Sub
