VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form sales_unpaid 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unpaid Sales Amount"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "sales_unpaid.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5895
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "sales_unpaid.frx":0E42
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "DAT"
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
      BeginProperty Column01 
         DataField       =   "INVOICE_NO"
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
      BeginProperty Column02 
         DataField       =   "AMT_UNPAID"
         Caption         =   "Amount"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
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
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4320
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
      MICON           =   "sales_unpaid.frx":0E57
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
      Caption         =   "Received Amount"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Amount"
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
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1695
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "sales_unpaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cust_name As New ADODB.Recordset


Private Sub cmd_Click()
If Check1.Value = 1 Then
        Dim x As Integer
        x = MsgBox("Are you sure this Transaction is Complete", vbQuestion Or vbYesNo, "Trasaction Completed ?")
        If x <> 6 Then
            Check1.Value = 0
            Exit Sub
        End If
        
        If Check1.Enabled = True Then
            Dim ADD_EXPENSE As New ADODB.Recordset
            ADD_EXPENSE.Open "SELECT * FROM EXPENSE", db, adOpenDynamic, adLockOptimistic
            ADD_EXPENSE.AddNew
            
            Dim ld As New ADODB.Recordset
            ld.Open "select * from AMT_UNPAID_REMIND where TRAN_TYPE='SALES' AND PARTY_NAME='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
            ld.MoveLast
            ADD_EXPENSE.Fields(0).Value = Format(ld.Fields(1).Value, "dd-MMM-yyyy")
            ld.Close
            
            ADD_EXPENSE.Fields(1).Value = "Discounted Amount(Sales)"
            ADD_EXPENSE.Fields(2).Value = VAL(Text1(0).Text) - VAL(Text1(1).Text)
            ADD_EXPENSE.Update
            ADD_EXPENSE.Close
        End If
        
            Dim REM_CUST As New ADODB.Recordset
            REM_CUST.Open "SELECT * FROM AMT_UNPAID_REMIND WHERE TRAN_TYPE='SALES' AND PARTY_NAME='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
            While REM_CUST.EOF <> True
                REM_CUST.Delete
                REM_CUST.MoveNext
            Wend
            
            Unload Me
            Exit Sub
ElseIf Check1.Value = 0 Then
            Dim up_CUST As New ADODB.Recordset
            up_CUST.Open "SELECT * FROM AMT_UNPAID_REMIND WHERE TRAN_TYPE='SALES' AND PARTY_NAME='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
            Dim total As Double
            total = VAL(Text1(1).Text)
            
            
                If total < up_CUST.Fields(2).Value Then
                    up_CUST.Fields(2).Value = up_CUST.Fields(2).Value - total
                    up_CUST.Update
                    Unload Me
                    Exit Sub
                End If
                
                If total = up_CUST.Fields(2).Value Then
                    up_CUST.Delete
                    Unload Me
                    Exit Sub
                End If
                
                If total > up_CUST.Fields(2).Value Then
                    Dim ta As Double
                    ta = total
                    Dim t As Boolean
                    t = True
                    While t = True
                    If ta >= up_CUST.Fields(2).Value Then
                        ta = ta - up_CUST.Fields(2).Value
                        up_CUST.Delete
                        up_CUST.MoveNext
                    Else
                        up_CUST.Fields(2).Value = up_CUST.Fields(2).Value - ta
                        up_CUST.Update
                        t = False
                    End If
        
                    Wend
                    Unload Me
                    Exit Sub
                End If
End If

        
End Sub

Private Sub Combo1_Click()
If Len(Combo1.Text) > 0 Then
    refresh_grid
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
    Me.Left = MDIForm1.Width / 2 - Me.Width / 2
    Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - MDIForm1.StatusBar1.Height

cust_name.Open "select distinct PARTY_NAME from AMT_UNPAID_REMIND where TRAN_TYPE='SALES'", db, adOpenKeyset, adLockOptimistic
If cust_name.RecordCount = 0 Then
    MsgBox "No Unpaid Customer Found ...", vbInformation, "No Unpaid Amount Found ..."
    Unload Me
    Exit Sub
End If

Combo1.Clear
While cust_name.EOF <> True
    Combo1.AddItem cust_name.Fields(0).Value
    cust_name.MoveNext
Wend


End Sub

Public Sub refresh_grid()
Dim rs As New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select DAT,INVOICE_NO,AMT_UNPAID from grid_sales_unpaid_data where PARTY_NAME='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rs
Dim R As New ADODB.Recordset
R.Open "SELECT SumOfAMT_UNPAID FROM QRY_UNPAID_REPORT WHERE PARTY_NAME='" & Combo1.Text & "' AND TRAN_TYPE='SALES'", db, adOpenDynamic, adLockOptimistic
Text1(0).Text = R.Fields(0).Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
cust_name.Close
End Sub

Private Sub Text1_Change(Index As Integer)

If Index = 1 Then
    If VAL(Text1(1).Text) > VAL(Text1(0).Text) Then
        MsgBox "You can not enter More than the total unpaid Amount ...", vbCritical, "Enter Proper Value ..."
        Text1(1).Text = Clear
        Exit Sub
    End If
    
End If

If Index = 1 Then
    If VAL(Text1(0).Text) = VAL(Text1(1).Text) Then
        Check1.Value = 1
        Check1.Enabled = False
    Else
        Check1.Value = 0
        Check1.Enabled = True
    End If
    
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 1 Then
    If KeyCode = 13 Then
        cmd_Click
    End If
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    
    ElseIf KeyAscii = 46 Then
        If InStr(1, Text1(1).Text, ".", vbTextCompare) > 0 Then
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End If
End Sub
