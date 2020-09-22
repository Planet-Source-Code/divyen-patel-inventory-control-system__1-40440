VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form FRM_CHANGE_ITEM_MASTER 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Item Type and Item Name ..."
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   ControlBox      =   0   'False
   Icon            =   "FRM_CHANGE_ITEM_MASTER.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5265
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
   Begin VB.ListBox List2 
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
      Height          =   2565
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox List1 
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
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin LVbuttons.LaVolpeButton but_gen_rpt 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Modify"
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_CHANGE_ITEM_MASTER.frx":0E42
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
   Begin LVbuttons.LaVolpeButton but_gen_rpt 
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   3600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Modify"
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_CHANGE_ITEM_MASTER.frx":0E5E
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
   Begin LVbuttons.LaVolpeButton but_gen_rpt 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Close"
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "FRM_CHANGE_ITEM_MASTER.frx":0E7A
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
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Name"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Type"
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
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FRM_CHANGE_ITEM_MASTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As New ADODB.Recordset
Dim r2 As New ADODB.Recordset

Private Sub but_gen_rpt_Click(Index As Integer)
If Index = 0 Then
    If Len(Text1(0).Text) > 0 Then
        If but_gen_rpt(0).Caption = "Modify" Then
            List1.Enabled = False
            List2.Enabled = False
            but_gen_rpt(2).Enabled = False
            but_gen_rpt(1).Enabled = False
            but_gen_rpt(0).Caption = "&Save"
            Text1(0).Enabled = True
            SendKeys "{TAB}"
            SendKeys "{END}"
        Else
            Dim MODIFY_RS As New ADODB.Recordset
            MODIFY_RS.Open "SELECT * FROM ItemType WHERE Itemtype='" & List1.List(List1.ListIndex) & "'", db, adOpenKeyset, adLockOptimistic
            MODIFY_RS.Fields(0).Value = Text1(0).Text
            MODIFY_RS.Update
            Text1(0).Enabled = False
            List1.Enabled = True
            List2.Enabled = True
            but_gen_rpt(0).Caption = "Modify"
            but_gen_rpt(1).Enabled = True
            but_gen_rpt(2).Enabled = True
            r.Requery
            List1.Clear
            List2.Clear
    
            While r.EOF <> True
                    List1.AddItem r.Fields(0).Value
                    r.MoveNext
            Wend
            
            MsgBox "Item Type Updated Successfully ...", vbInformation, "Item type Updated ..."
        End If
        
    End If
ElseIf Index = 1 Then
    If Len(Text1(1).Text) > 0 Then
        If but_gen_rpt(1).Caption = "Modify" Then
            List1.Enabled = False
            List2.Enabled = False
            but_gen_rpt(0).Enabled = False
            but_gen_rpt(2).Enabled = False
            but_gen_rpt(1).Caption = "&Save"
            Text1(1).Enabled = True
            SendKeys "{TAB}"
            SendKeys "{END}"
        Else
            Dim update_item_nam As New ADODB.Recordset
            update_item_nam.Open "select Item_name from Item_master where Item_name='" & List2.List(List2.ListIndex) & "'", db, adOpenKeyset, adLockOptimistic
            update_item_nam.Fields(0).Value = Text1(1).Text
            update_item_nam.Update
            
                        
            Dim d As New ADODB.Connection
            d.Open db.ConnectionString
            d.Execute "update Customer_System_datail set Processor='" & Text1(1).Text & "' where Processor='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Motherboard='" & Text1(1).Text & "' where Motherboard='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Hardisk='" & Text1(1).Text & "' where Hardisk='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Floppydisk='" & Text1(1).Text & "' where Floppydisk='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set RAM='" & Text1(1).Text & "' where RAM='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set CD_ROM='" & Text1(1).Text & "' where CD_ROM='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Speaker='" & Text1(1).Text & "' where Speaker='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Mouse='" & Text1(1).Text & "' where Mouse='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Keyboard='" & Text1(1).Text & "' where Keyboard='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Printer='" & Text1(1).Text & "' where Printer='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Scanner='" & Text1(1).Text & "' where Scanner='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Sound_card='" & Text1(1).Text & "' where Sound_card='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set CD_writer='" & Text1(1).Text & "' where CD_writer='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Modem='" & Text1(1).Text & "' where Modem='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Web_cam='" & Text1(1).Text & "' where Web_cam='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Stabilizer='" & Text1(1).Text & "' where Stabilizer='" & List2.List(List2.ListIndex) & "'"
            d.Execute "update Customer_System_datail set Zip_Drive='" & Text1(1).Text & "' where Zip_Drive='" & List2.List(List2.ListIndex) & "'"
            
            
            
            
            
            
            
            
            
            
            
            
            
            Text1(1).Enabled = False
            but_gen_rpt(1).Caption = "Modify"
            but_gen_rpt(2).Enabled = True
            but_gen_rpt(0).Enabled = True
            List1.Enabled = True
            List2.Enabled = True
            Text1(0).Text = Clear
            Text1(1).Text = Clear
            List2.Clear
            List1.ListIndex = -1

            MsgBox "Item name Updated Successfully", vbInformation, "Item name updated ..."
            
            
            

        End If
        
    End If
    
ElseIf Index = 2 Then
    Unload Me
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 And but_gen_rpt(2).Enabled = True Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
    Me.Left = 0
    Me.Top = 0
    
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    
    r.Open "select * from ItemType", db, adOpenKeyset, adLockOptimistic
    List1.Clear
    List2.Clear
    
    While r.EOF <> True
        List1.AddItem r.Fields(0).Value
        r.MoveNext
    Wend
    
    r2.Open "select Item_name from Item_master where Itemtype='" & List1.List(0) & "'", db, adOpenKeyset, adLockOptimistic
    List2.Clear
    While r2.EOF <> True
        List2.AddItem r2.Fields(0).Value
        r2.MoveNext
    Wend
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
r.Close
r2.Close
Exit Sub
End Sub

Private Sub List1_Click()
    Text1(0).Text = List1.List(List1.ListIndex)
    r2.Close
    r2.Open "select Item_name from Item_master where Itemtype='" & Text1(0).Text & "'", db, adOpenKeyset, adLockOptimistic
    List2.Clear
    While r2.EOF <> True
        List2.AddItem r2.Fields(0).Value
        r2.MoveNext
    Wend
    
End Sub

Private Sub List2_Click()
Text1(1).Text = List2.List(List2.ListIndex)

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Then
    If KeyCode = 13 Then
        but_gen_rpt_Click (0)
    End If
ElseIf Index = 1 Then
    If KeyCode = 13 Then
        but_gen_rpt_Click (0)
    End If
End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then

ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then

ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then

ElseIf KeyAscii = 8 Then

ElseIf KeyAscii = 32 Then


Else
    KeyAscii = 0
End If

End Sub
