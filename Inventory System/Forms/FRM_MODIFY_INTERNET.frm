VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_MODIFY_INTERNET 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Internet Connection Sale ..."
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   Icon            =   "FRM_MODIFY_INTERNET.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7020
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Modify"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_MODIFY_INTERNET.frx":0E42
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
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
      Height          =   315
      Index           =   1
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
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
      Height          =   315
      Index           =   0
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   0   'False
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
      ColumnCount     =   12
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
         Caption         =   "Customer ID"
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
      BeginProperty Column05 
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
         Caption         =   "Rate Per Unit"
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
         Caption         =   "Purchase Invoice No"
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
      BeginProperty Column11 
         DataField       =   "P_RATE"
         Caption         =   "Purchase Rate"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_MODIFY_INTERNET.frx":0E5E
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   6000
      Picture         =   "FRM_MODIFY_INTERNET.frx":0E7A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   885
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
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1815
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
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "FRM_MODIFY_INTERNET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cust_names As New ADODB.Recordset
Dim rs_data As New ADODB.Recordset
Dim in_rs As New ADODB.Recordset

Private Sub Combo1_Click(Index As Integer)
If Index = 0 Then
    in_rs.Open "select Invoice_no from Sales_master where Party_name='" & Combo1(0).Text & "' and Item_type='Internet Connection'", db, adOpenKeyset, adLockOptimistic
    Combo1(1).Clear
    While in_rs.EOF <> True
        Combo1(1).AddItem in_rs.Fields(0).Value
        in_rs.MoveNext
    Wend
    in_rs.Close
    Set DataGrid1.DataSource = Nothing
ElseIf Index = 1 Then
    If rs_data.State = adStateOpen Then
        rs_data.Close
    End If
    rs_data.CursorLocation = adUseClient
    rs_data.Open "select * from Sales_master where Party_name='" & Combo1(0).Text & "' and Invoice_no='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = rs_data
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
    
    cust_names.Open "select Party_name from Sales_master where Item_type='Internet Connection'", db, adOpenKeyset, adLockOptimistic
    If cust_names.RecordCount = 0 Then
        Unload Me
        MsgBox "No Record Found ...", vbInformation, "No Record Found ..."
        Exit Sub
    End If
    
    Combo1(0).Clear
    While cust_names.EOF <> True
        Combo1(0).AddItem cust_names.Fields(0).Value
        cust_names.MoveNext
    Wend
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
cust_names.Close
rs_data.Close
in_rs.Close
Exit Sub
End Sub

Private Sub LaVolpeButton1_Click()
If Len(Combo1(1).Text) = 0 Then
    MsgBox "Select Invoice Number ...", vbInformation, "Invoice Number Not Selected ..."
    Exit Sub
End If

rs_data.Requery



    Dim update_awa_pur_stock As New ADODB.Recordset
    update_awa_pur_stock.Open "select * from AVAILABLE_PURCHASED_STOCK where Item_type='Internet Connection' and Item_name='" & rs_data.Fields(5).Value & "' and Invoice_no='" & rs_data.Fields(8).Value & "'", db, adOpenKeyset, adLockOptimistic

    

        If update_awa_pur_stock.RecordCount > 0 Then
            update_awa_pur_stock.Fields(5).Value = update_awa_pur_stock.Fields(5).Value + VAL(rs_data.Fields(6).Value)
            update_awa_pur_stock.Fields(7).Value = update_awa_pur_stock.Fields(5).Value * update_awa_pur_stock.Fields(6).Value
            update_awa_pur_stock.Update
        Else
            update_awa_pur_stock.AddNew
            update_awa_pur_stock.Fields(0).Value = rs_data.Fields(8).Value
            update_awa_pur_stock.Fields(1).Value = rs_data.Fields(9).Value
            update_awa_pur_stock.Fields(2).Value = rs_data.Fields(10).Value
            update_awa_pur_stock.Fields(3).Value = "Internet Connection"
            update_awa_pur_stock.Fields(4).Value = rs_data.Fields(5).Value
            update_awa_pur_stock.Fields(5).Value = rs_data.Fields(6).Value
            update_awa_pur_stock.Fields(6).Value = rs_data.Fields(11).Value
            update_awa_pur_stock.Fields(7).Value = update_awa_pur_stock.Fields(5).Value * update_awa_pur_stock.Fields(6).Value
            update_awa_pur_stock.Update
        End If
    update_awa_pur_stock.Close

rs_data.Requery

            Dim update_item_master As New ADODB.Recordset
            update_item_master.Open "select * from Item_master where Item_name='" & rs_data.Fields(5).Value & "' and Itemtype='Internet Connection'", db, adOpenKeyset, adLockOptimistic
            update_item_master.Fields(3).Value = update_item_master.Fields(3).Value + rs_data.Fields(6).Value
            update_item_master.Update

rs_data.Requery

Dim username As String
Dim rdate As String
Dim edate As String
Dim bpr As String
Dim seno As String

Dim rs_get_old_data As New ADODB.Recordset
rs_get_old_data.Open "select * from INTERNET_CONNECTIONS where sales_invoice_no='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic
username = rs_get_old_data.Fields(2).Value
rdate = rs_get_old_data.Fields(3).Value
edate = rs_get_old_data.Fields(4).Value
bpr = rs_get_old_data.Fields(5).Value
seno = rs_get_old_data.Fields(8).Value
rs_get_old_data.Close

DoEvents

db.Execute "delete from INTERNET_CONNECTIONS where sales_invoice_no='" & Combo1(1).Text & "'"
db.Execute "delete from AMT_UNPAID_REMIND where INVOICE_NO='" & Combo1(1).Text & "' and PARTY_NAME='" & Combo1(0).Text & "'"
db.Execute "delete from Sales_master where Invoice_no='" & Combo1(1).Text & "' and Party_name='" & Combo1(0).Text & "' and Item_type='Internet Connection'"
db.Execute "delete from DATE_PROFIT where INVOICE_NUMBER='" & Combo1(1).Text & "'"
db.Execute "delete from EXPENSE where INVOICE_NO='" & Combo1(1).Text & "'"

FRM_UPDATE_INTERNET.Text1(0).Text = rs_data.Fields(1).Value
FRM_UPDATE_INTERNET.Combo1.Text = Combo1(0).Text
FRM_UPDATE_INTERNET.Text1(0).Enabled = False
FRM_UPDATE_INTERNET.Combo1.Enabled = False
FRM_UPDATE_INTERNET.LaVolpeButton1.Enabled = False
FRM_UPDATE_INTERNET.LaVolpeButton2.Enabled = False
FRM_UPDATE_INTERNET.Combo2.Text = rs_data.Fields(5).Value
FRM_UPDATE_INTERNET.Combo3.Text = rs_data.Fields(9).Value
FRM_UPDATE_INTERNET.Text1(1).Text = rs_data.Fields(11).Value

FRM_UPDATE_INTERNET.inno = Combo1(1).Text
FRM_UPDATE_INTERNET.Text1(4).Text = seno
FRM_UPDATE_INTERNET.Text1(2).Text = username
FRM_UPDATE_INTERNET.Text1(3).Text = bpr
FRM_UPDATE_INTERNET.DTPicker1(0).Value = Format(rdate, "dd-MMM-yyyy")
FRM_UPDATE_INTERNET.DTPicker1(1).Value = Format(edate, "dd-MMM-yyyy")
FRM_UPDATE_INTERNET.Show
Unload Me
End Sub

Private Sub LaVolpeButton2_Click()

If Len(Combo1(1).Text) = 0 Then
    MsgBox "Select Invoice Number ...", vbInformation, "Invoice Number Not Selected ..."
    Exit Sub
End If
Dim x As Integer
x = MsgBox("Are you sure you want to delete this Sales Entry ...", vbQuestion Or vbYesNo, "Want to Delete this entry ...")
If x <> 6 Then
    Exit Sub
End If

rs_data.Requery



    Dim update_awa_pur_stock As New ADODB.Recordset
    update_awa_pur_stock.Open "select * from AVAILABLE_PURCHASED_STOCK where Item_type='Internet Connection' and Item_name='" & rs_data.Fields(5).Value & "' and Invoice_no='" & rs_data.Fields(8).Value & "'", db, adOpenKeyset, adLockOptimistic

    

        If update_awa_pur_stock.RecordCount > 0 Then
            update_awa_pur_stock.Fields(5).Value = update_awa_pur_stock.Fields(5).Value + VAL(rs_data.Fields(6).Value)
            update_awa_pur_stock.Fields(7).Value = update_awa_pur_stock.Fields(5).Value * update_awa_pur_stock.Fields(6).Value
            update_awa_pur_stock.Update
        Else
            update_awa_pur_stock.AddNew
            update_awa_pur_stock.Fields(0).Value = rs_data.Fields(8).Value
            update_awa_pur_stock.Fields(1).Value = rs_data.Fields(9).Value
            update_awa_pur_stock.Fields(2).Value = rs_data.Fields(10).Value
            update_awa_pur_stock.Fields(3).Value = "Internet Connection"
            update_awa_pur_stock.Fields(4).Value = rs_data.Fields(5).Value
            update_awa_pur_stock.Fields(5).Value = rs_data.Fields(6).Value
            update_awa_pur_stock.Fields(6).Value = rs_data.Fields(11).Value
            update_awa_pur_stock.Fields(7).Value = update_awa_pur_stock.Fields(5).Value * update_awa_pur_stock.Fields(6).Value
            update_awa_pur_stock.Update
        End If
    update_awa_pur_stock.Close

rs_data.Requery

            Dim update_item_master As New ADODB.Recordset
            update_item_master.Open "select * from Item_master where Item_name='" & rs_data.Fields(5).Value & "' and Itemtype='Internet Connection'", db, adOpenKeyset, adLockOptimistic
            update_item_master.Fields(3).Value = update_item_master.Fields(3).Value + rs_data.Fields(6).Value
            update_item_master.Update

rs_data.Requery



db.Execute "delete from INTERNET_CONNECTIONS where sales_invoice_no='" & Combo1(1).Text & "'"
db.Execute "delete from AMT_UNPAID_REMIND where INVOICE_NO='" & Combo1(1).Text & "' and PARTY_NAME='" & Combo1(0).Text & "'"
db.Execute "delete from Sales_master where Invoice_no='" & Combo1(1).Text & "' and Party_name='" & Combo1(0).Text & "' and Item_type='Internet Connection'"
db.Execute "delete from DATE_PROFIT where INVOICE_NUMBER='" & Combo1(1).Text & "'"
db.Execute "delete from EXPENSE where INVOICE_NO='" & Combo1(1).Text & "'"

End Sub
