VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_invoice_no_pname_option 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Invoice Number and Party name ..."
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "frm_invoice_no_pname_option.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9015
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Modify Entry"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_invoice_no_pname_option.frx":0E42
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
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
      Top             =   840
      Width           =   3735
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
      Top             =   360
      Width           =   5295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6165
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Delete Invoice"
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
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_invoice_no_pname_option.frx":0E5E
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
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
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1695
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
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frm_invoice_no_pname_option"
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
    Text1.Text = Clear
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
    LaVolpeButton1.Enabled = False
    LaVolpeButton2.Enabled = False
    
    p_name.Open "select DISTINCT Party_name FROM Purchase_master", db, adOpenKeyset, adLockOptimistic
    
    If p_name.RecordCount = 0 Then
        MsgBox "No Record Found ...", vbInformation, "No Record Found ..."
        Unload Me
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

Private Sub LaVolpeButton1_Click()
    Dim check_data As New ADODB.Recordset
    check_data.Open "select * from AVAILABLE_PURCHASED_STOCK where Invoice_no='" & Combo1(1).Text & "' and Party_name='" & Combo1(0).Text & "' ORDER BY Invoice_no,Party_name,Purchase_Date,Item_type,Item_name,Qty,price_per_unit,total_amt,Item_Description", db, adOpenKeyset, adLockOptimistic
    
    Dim check_st As Boolean
    check_st = True
    
    rs_data.Requery
    
    If rs_data.RecordCount = check_data.RecordCount Then
        While rs_data.EOF <> True
            For i = 0 To 7
                    If rs_data.Fields(i).Value <> check_data.Fields(i).Value Then
                        check_st = False
                        GoTo A1:
                    End If
            Next
            rs_data.MoveNext
            check_data.MoveNext
        Wend
    
    Else
        check_st = False
        GoTo A1:
    End If
    Dim CR_RS As New ADODB.Recordset
    CR_RS.Open "SELECT * FROM CRITICAL_PURCHASE_DATA", db, adOpenKeyset, adLockOptimistic
    
    While CR_RS.EOF <> True
        CR_RS.Delete
        CR_RS.MoveNext
    Wend
    
    rs_data.Requery
    While rs_data.EOF <> True
        CR_RS.AddNew
        For i = 0 To 8
            CR_RS.Fields(i).Value = rs_data.Fields(i).Value
        Next
        CR_RS.Update
        rs_data.MoveNext
    Wend
    
    CR_RS.Requery
    
    
    
    Dim sys_cur_ino As New ADODB.Recordset
    sys_cur_ino.Open "select * from SYS_CURRENT_INVOICE", db, adOpenKeyset, adLockOptimistic
    
    While CR_RS.EOF <> True
        sys_cur_ino.AddNew
            For i = 3 To 8
                sys_cur_ino.Fields(i - 3).Value = CR_RS.Fields(i).Value
            Next
        sys_cur_ino.Update
        CR_RS.MoveNext
    Wend
    CR_RS.Requery
    FRM_MODIFY_PURCHASE_BOOK.Show
    FRM_MODIFY_PURCHASE_BOOK.Text1(0).Text = CR_RS.Fields(0).Value
    FRM_MODIFY_PURCHASE_BOOK.DTPicker1.Value = Format(CR_RS.Fields(2).Value, "dd-MMM-yyyy")
    FRM_MODIFY_PURCHASE_BOOK.Combo3.Text = CR_RS.Fields(1).Value

    Unload Me
    
    Exit Sub
    
A1:
    MsgBox "You can not update this purchase entry , Item sold from this invoice ...", vbInformation, "Purchase Entry can not be saved ..."
End Sub

Private Sub LaVolpeButton2_Click()
Dim x As Integer
x = MsgBox("Are you sure you want to delete Purchase this Invoice ...", vbQuestion Or vbYesNo, "Are you sure ?")
If x = 6 Then
        Dim check_data As New ADODB.Recordset
        check_data.Open "select * from AVAILABLE_PURCHASED_STOCK where Invoice_no='" & Combo1(1).Text & "' and Party_name='" & Combo1(0).Text & "' ORDER BY Invoice_no,Party_name,Purchase_Date,Item_type,Item_name,Qty,price_per_unit,total_amt,Item_Description", db, adOpenKeyset, adLockOptimistic
    
        Dim check_st As Boolean
        check_st = True
    
        rs_data.Requery
    
        If rs_data.RecordCount = check_data.RecordCount Then
            While rs_data.EOF <> True
                For i = 0 To 7
                        If rs_data.Fields(i).Value <> check_data.Fields(i).Value Then
                            check_st = False
                            GoTo A1:
                        End If
                Next
                rs_data.MoveNext
                check_data.MoveNext
            Wend
    
        Else
                check_st = False
            GoTo A1:
        End If
                    Dim CR_RS As New ADODB.Recordset
                    CR_RS.Open "SELECT * FROM CRITICAL_PURCHASE_DATA", db, adOpenKeyset, adLockOptimistic
                
                        While CR_RS.EOF <> True
                            CR_RS.Delete
                            CR_RS.MoveNext
                        Wend
                        
                       rs_data.Requery
                        While rs_data.EOF <> True
                            CR_RS.AddNew
                            For i = 0 To 8
                                CR_RS.Fields(i).Value = rs_data.Fields(i).Value
                            Next
                            CR_RS.Update
                            rs_data.MoveNext
                        Wend
                        
                            Dim del_awa_st As New ADODB.Recordset
    Dim del_pur_mas As New ADODB.Recordset
    
    del_awa_st.Open "select * from AVAILABLE_PURCHASED_STOCK where Invoice_no='" & Combo1(1).Text & "' and Party_name='" & Combo1(0).Text & "'", db, adOpenKeyset, adLockOptimistic

    While del_awa_st.EOF <> True
        del_awa_st.Delete
        del_awa_st.MoveNext
    Wend
    
    del_pur_mas.Open "select * from Purchase_master where Invoice_no='" & Combo1(1).Text & "' and Party_name='" & Combo1(0).Text & "'", db, adOpenKeyset, adLockOptimistic
    
    While del_pur_mas.EOF <> True
        del_pur_mas.Delete
        del_pur_mas.MoveNext
    Wend
    
    Dim updata_item_master As New ADODB.Recordset
    Dim critical_pur_data As New ADODB.Recordset
    
    critical_pur_data.Open "select * from CRITICAL_PURCHASE_DATA", db, adOpenKeyset, adLockOptimistic
    
    While critical_pur_data.EOF <> True
            updata_item_master.Open "select * from Item_master where Itemtype='" & critical_pur_data.Fields(3).Value & "' and Item_name='" & critical_pur_data.Fields(4).Value & "'", db, adOpenKeyset, adLockOptimistic
            updata_item_master.Fields(3).Value = updata_item_master.Fields(3).Value - VAL(critical_pur_data.Fields(5).Value)
            updata_item_master.Update
            updata_item_master.Close
            critical_pur_data.MoveNext
    Wend
    
    
    Dim del_unpaid As New ADODB.Recordset
    del_unpaid.Open "select * from AMT_UNPAID_REMIND where PARTY_NAME='" & Combo1(0).Text & "' and INVOICE_NO='" & Combo1(1).Text & "'", db, adOpenKeyset, adLockOptimistic
    
    If del_unpaid.RecordCount > 0 Then
        del_unpaid.Delete
    End If
    
    del_unpaid.Close
    db.Execute "DELETE FROM EXPENSE WHERE INVOICE_NO='" & Combo1(1).Text & "' AND EXPENSE_TYPE='Discounted Amount(Purchase)'"
    db.Execute "DELETE FROM INCOME WHERE INVOICE_NUMBER='" & Combo1(1).Text & "' AND INCOME_TYPE='Discounted Amount(Purchase)'"

    MsgBox "Invoice Deleted Successfully ....", vbInformation, "Purchase Invoice Deleted ..."
    Unload Me
    Exit Sub
    End If
    Exit Sub
    
A1:
    MsgBox "You can not Delete this purchase entry , Item sold from this invoice ...", vbInformation, "Purchase Entry can not be saved ..."

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
    LaVolpeButton1.Enabled = True
    LaVolpeButton2.Enabled = True
Else
    LaVolpeButton1.Enabled = False
    LaVolpeButton2.Enabled = False
End If

End Sub
