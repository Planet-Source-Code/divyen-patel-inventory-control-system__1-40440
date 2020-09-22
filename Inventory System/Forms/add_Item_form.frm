VERSION 5.00
Begin VB.Form add_Item_form 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter New Item Name ..."
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "add_Item_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   109
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   283
   Begin VB.CommandButton cmd 
      Caption         =   "&Add Item"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Number"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "add_Item_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim LAST_ID As String
Public FN As String

Private Sub cmd_Click()
If Len(Text1(1).Text) > 0 Then
            Rs.AddNew
            Rs.Fields(0).Value = Text1(0).Text
            Rs.Fields(1).Value = Text1(1).Text
            Rs.Fields(2).Value = Text1(2).Text
            Rs.Fields(3).Value = 0
            On Error GoTo A1
            Rs.Update
            If FN = "M" Then
                FRM_MODIFY_PURCHASE_BOOK.Refresh_combobox (1)
                FRM_MODIFY_PURCHASE_BOOK.Combo1.Text = Text1(1).Text
            Else
                Purchase_form.Refresh_combobox (1)
                Purchase_form.Combo1.Text = Text1(1).Text
            End If
            
                        Unload Me
            SendKeys ("{TAB}")
            Exit Sub
A1:
                MsgBox "Item Already Exist ...", vbCritical, "Item Exists..."
                Rs.CancelUpdate
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
    If FN = "M" Then
            Me.Left = (FRM_MODIFY_PURCHASE_BOOK.Width / 2) - (Me.Width / 2)
            Me.Top = (FRM_MODIFY_PURCHASE_BOOK.Height / 2) - (Me.Height / 2)
            Text1(2).Text = FRM_MODIFY_PURCHASE_BOOK.Combo2.Text
    Else
            Me.Left = (Purchase_form.Width / 2) - (Me.Width / 2)
            Me.Top = (Purchase_form.Height / 2) - (Me.Height / 2)
            Text1(2).Text = Purchase_form.Combo2.Text
    End If
    
    
       
    Rs.Open "select * from Item_master", db, adOpenDynamic, adLockOptimistic
    If Rs.EOF <> True Then
            Rs.MoveLast
            LAST_ID = Mid(Rs.Fields(0).Value, 2, Len(Rs.Fields(0).Value))
    
            Dim ID As Integer
            ID = VAL(LAST_ID)
            ID = ID + 1
            LAST_ID = ID
            
            If Len(LAST_ID) = 1 Then
                LAST_ID = "I000" & LAST_ID
            ElseIf Len(LAST_ID) = 2 Then
                LAST_ID = "I00" & LAST_ID
            ElseIf Len(LAST_ID) = 3 Then
                LAST_ID = "I0" & LAST_ID
            ElseIf Len(LAST_ID) = 4 Then
                LAST_ID = "I" & LAST_ID
            End If
            
             Text1(0).Text = LAST_ID
    
    Else
    
            Text1(0).Text = "I0001"
        
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
       Rs.Close
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Len(Text1(1).Text) > 0 Then
If KeyCode = 13 Then
    cmd_Click
End If
End If

End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then

ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then

ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then

ElseIf KeyAscii = 8 Then

ElseIf KeyAscii = 32 Then

ElseIf KeyAscii = 46 Then

Else
    KeyAscii = 0
End If
End If
End Sub
