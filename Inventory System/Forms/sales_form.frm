VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form sales_form 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Entry Form"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   Icon            =   "sales_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10035
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text2 
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
      Height          =   855
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   6120
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   9353162
      ForeColor       =   128
      HeadLines       =   1
      RowHeight       =   16
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
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Current Bill Items"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "item_name"
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
      BeginProperty Column01 
         DataField       =   "qty"
         Caption         =   "QTY"
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
         DataField       =   "price_p_unit"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "item_type"
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
         DataField       =   "purchased_from"
         Caption         =   "Purchased From"
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
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Width           =   9855
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1575
         Left            =   5760
         TabIndex        =   39
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2778
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   7714248
         ForeColor       =   128
         HeadLines       =   1
         RowHeight       =   16
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
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Purchaed Item Details"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   11
         Top             =   2160
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
         Index           =   3
         Left            =   6240
         TabIndex        =   12
         Top             =   2160
         Width           =   1455
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
         Height          =   345
         Index           =   6
         Left            =   1920
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
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
         Height          =   345
         Index           =   5
         Left            =   1920
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
      Begin LVbuttons.LaVolpeButton cmd_op 
         Height          =   375
         Index           =   0
         Left            =   1470
         TabIndex        =   13
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Add New"
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
         MICON           =   "sales_form.frx":0E42
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
      Begin LVbuttons.LaVolpeButton cmd_op 
         Height          =   375
         Index           =   1
         Left            =   2910
         TabIndex        =   14
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "sales_form.frx":0E5E
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
      Begin LVbuttons.LaVolpeButton cmd_op 
         Height          =   375
         Index           =   2
         Left            =   4335
         TabIndex        =   15
         Top             =   2640
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
         COLTYPE         =   2
         BCOL            =   14737632
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "sales_form.frx":0E7A
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
      Begin LVbuttons.LaVolpeButton cmd_op 
         Height          =   375
         Index           =   3
         Left            =   5790
         TabIndex        =   16
         Top             =   2640
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
         COLTYPE         =   2
         BCOL            =   14737632
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "sales_form.frx":0E96
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
      Begin LVbuttons.LaVolpeButton cmd_op 
         Height          =   375
         Index           =   4
         Left            =   7215
         TabIndex        =   17
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cancel"
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
         MICON           =   "sales_form.frx":0EB2
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
         Caption         =   "Qty"
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
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Price Per Unit"
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
         Left            =   2040
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Amount"
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
         Left            =   4800
         TabIndex        =   35
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purchased Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Availabel Qty"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
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
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1695
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
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1695
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
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   9855
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
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1455
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
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   315
         Left            =   5640
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
         MICON           =   "sales_form.frx":0ECE
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
         Left            =   4200
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
         MICON           =   "sales_form.frx":0EEA
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
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         CalendarBackColor=   16777215
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   47841283
         CurrentDate     =   37457
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
         Index           =   3
         Left            =   6840
         TabIndex        =   28
         Top             =   360
         Width           =   1095
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
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1695
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
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1695
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
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
   End
   Begin LVbuttons.LaVolpeButton cmd 
      Height          =   735
      Left            =   8040
      TabIndex        =   19
      Top             =   6240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Save and Generate Bill (F5)"
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
      MICON           =   "sales_form.frx":0F06
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
      Caption         =   "Items Description"
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
      Index           =   12
      Left            =   4920
      TabIndex        =   38
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label issues 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Be Sure while adding Sales entry in to Sales Master. You will not be able to delete it later or Modify any of the field."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   360
      Width           =   9135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Entry Form"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "sales_form.frx":0F22
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10080
   End
End
Attribute VB_Name = "sales_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_custo_datail As New ADODB.Recordset
Dim rs_sales_master As New ADODB.Recordset
Dim rs_cust_system_detail As New ADODB.Recordset
Dim RS_ITEM As New ADODB.Recordset
Dim rs_item_type As New ADODB.Recordset
Dim rs_cur_invoice_item As New ADODB.Recordset
Dim RS_PUR_ITEMS As New ADODB.Recordset
Dim INVOICE_NUMBER As String
Dim rs_cur_record_count As New ADODB.Recordset
Dim Status As Boolean
Dim RS_SALES_INVOICE_DESC As New ADODB.Recordset
Public SNAME As String
Public FNAME As String
Public SALE_TYPE As String
Public SYSTEM_QTY As Integer
Public TOTAL_TRAN_AMT As Double
Dim GRID2_CLICKED As Boolean








Private Sub cmd_Click()

If SALE_TYPE = "SYSTEM" Then
            Dim d As New ADODB.Recordset
            d.Open "SELECT * FROM CHECK_QTY", db, adOpenDynamic, adLockOptimistic
            
            While d.EOF <> True
                If d.Fields(1).Value <> SYSTEM_QTY Then
                    MsgBox "Qty of some Item is less or greater than the required Qty", vbCritical, "Enter Proper Qty from Each Items"
                    Exit Sub
                End If
                d.MoveNext
            Wend
End If

If Len(Combo1.Text) = 0 Then
    MsgBox "Customer Name not found ...", vbInformation, "Enter Party Name ..."
    Exit Sub
End If

SNAME = "SAVED"
FNAME = Clear







DoEvents
rs_cur_invoice_item.Requery
rs_cur_record_count.Requery




Dim rs_update_it_master As New ADODB.Recordset
Dim rs_aps As New ADODB.Recordset

If rs_cur_record_count.Fields(0).Value > 0 Then
    
    
    
    
    TOTAL_TRAN_AMT = TOTAL_AMT("SALES")
    
    While rs_cur_invoice_item.EOF <> True
        Dim rs_profit As New ADODB.Recordset
        rs_profit.Open "select price_per_unit from Purchase_master where Invoice_no='" & rs_cur_invoice_item.Fields(6).Value & "' AND Party_name='" & rs_cur_invoice_item.Fields(5).Value & "' AND Item_type='" & rs_cur_invoice_item.Fields(4).Value & "' AND Item_name='" & rs_cur_invoice_item.Fields(0).Value & "'", db, adOpenDynamic, adLockOptimistic
        Dim AC_VAL As Double
        AC_VAL = rs_profit.Fields(0).Value * rs_cur_invoice_item.Fields(1).Value
        Dim SAL_VAL As Double
        SAL_VAL = rs_cur_invoice_item.Fields(2).Value * rs_cur_invoice_item.Fields(1).Value
        Dim PROFIT As Double
        PROFIT = VAL(SAL_VAL) - VAL(AC_VAL)
        
        Dim R As New ADODB.Recordset
        R.Open "SELECT * FROM DATE_PROFIT", db, adOpenDynamic, adLockOptimistic
        R.AddNew
        R.Fields(0).Value = DTPicker1.Value
        R.Fields(1).Value = PROFIT
        R.Fields(2).Value = Text1(0).Text
        R.Update
        R.Close
        
        rs_cur_invoice_item.MoveNext
        rs_profit.Close
    Wend
    
    
    rs_cur_record_count.Requery
    rs_cur_invoice_item.Requery
    
    
    
    
    RS_SALES_INVOICE_DESC.AddNew
    RS_SALES_INVOICE_DESC.Fields(0).Value = Text1(0).Text
    RS_SALES_INVOICE_DESC.Fields(1).Value = Text2.Text
    RS_SALES_INVOICE_DESC.Update

    
    While rs_cur_invoice_item.EOF <> True
        rs_update_it_master.Open "select * from Item_master where Item_name='" & rs_cur_invoice_item.Fields(0).Value & "' AND Itemtype='" & rs_cur_invoice_item.Fields(4).Value & "'", db, adOpenDynamic, adLockOptimistic
        rs_aps.Open "select * from AVAILABLE_PURCHASED_STOCK where Party_name='" & rs_cur_invoice_item.Fields(5).Value & "' AND Item_type='" & rs_cur_invoice_item.Fields(4).Value & "' AND Item_name='" & rs_cur_invoice_item.Fields(0).Value & "' AND Invoice_no='" & rs_cur_invoice_item.Fields(6).Value & "'", db, adOpenDynamic, adLockOptimistic
        rs_sales_master.AddNew
        rs_sales_master.Fields(0).Value = Text1(0).Text
        rs_sales_master.Fields(1).Value = Text1(4).Text
        rs_sales_master.Fields(2).Value = Combo1.Text
        rs_sales_master.Fields(3).Value = DTPicker1.Value
        rs_sales_master.Fields(4).Value = rs_cur_invoice_item.Fields(4).Value
        rs_sales_master.Fields(5).Value = rs_cur_invoice_item.Fields(0).Value
        rs_sales_master.Fields(6).Value = rs_cur_invoice_item.Fields(1).Value
        rs_sales_master.Fields(7).Value = rs_cur_invoice_item.Fields(2).Value
        rs_sales_master.Fields(8).Value = rs_cur_invoice_item.Fields(6).Value
        rs_sales_master.Fields(9).Value = rs_cur_invoice_item.Fields(5).Value
        rs_sales_master.Fields(10).Value = rs_cur_invoice_item.Fields(7).Value
        rs_sales_master.Fields(11).Value = rs_cur_invoice_item.Fields(8).Value
        
        rs_update_it_master.Fields(3).Value = VAL(rs_update_it_master.Fields(3).Value) - VAL(rs_cur_invoice_item.Fields(1).Value)
        rs_update_it_master.Update
        rs_update_it_master.Close
        ''' CHANGE
        rs_aps.Fields(5).Value = VAL(rs_aps.Fields(5).Value) - VAL(rs_cur_invoice_item.Fields(1).Value)
        rs_aps.Fields(7).Value = VAL(rs_aps.Fields(5).Value) * VAL(rs_aps.Fields(6).Value)
        rs_aps.Update
        If rs_aps.Fields(5).Value = 0 Then
            rs_aps.Delete
        End If
        
        rs_aps.Close
        rs_cur_invoice_item.MoveNext
    Wend
    
    rs_sales_master.Update
    
    If SALE_TYPE = "SYSTEM" Then
    Dim UPDATE_SYS_REPORT As New ADODB.Recordset
    Dim sysnors As New ADODB.Recordset
    UPDATE_SYS_REPORT.Open "SELECT * FROM Customer_System_datail", db, adOpenDynamic, adLockOptimistic
    
    
    
    
    sysnors.Open "SELECT * FROM INVOICE_NUMBER_SYSTEM_ID WHERE INVOICE_NUMBER='" & Text1(0).Text & "'", db, adOpenDynamic, adLockOptimistic

    
    
    
    While sysnors.EOF <> True
        UPDATE_SYS_REPORT.AddNew
        rs_cur_invoice_item.Requery
        rs_cur_invoice_item.MoveFirst
        UPDATE_SYS_REPORT.Fields(0).Value = sysnors.Fields(0).Value
    
    While rs_cur_invoice_item.EOF <> True
    
    If rs_cur_invoice_item.Fields(4).Value = "Processor" Then
            UPDATE_SYS_REPORT.Fields(1).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "MotherBoard" Then
            UPDATE_SYS_REPORT.Fields(2).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Harddisk" Then
            UPDATE_SYS_REPORT.Fields(3).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Floppydisk" Then
            UPDATE_SYS_REPORT.Fields(4).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "RAM" Then
            UPDATE_SYS_REPORT.Fields(5).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "CD ROM" Then
            UPDATE_SYS_REPORT.Fields(6).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Speaker" Then
            UPDATE_SYS_REPORT.Fields(7).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Mouse" Then
            UPDATE_SYS_REPORT.Fields(8).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Keyboard" Then
            UPDATE_SYS_REPORT.Fields(9).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Printer" Then
            UPDATE_SYS_REPORT.Fields(10).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Scaner" Then
            UPDATE_SYS_REPORT.Fields(11).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "SoundCard" Then
            UPDATE_SYS_REPORT.Fields(12).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "CD WRITER" Then
            UPDATE_SYS_REPORT.Fields(13).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Modem" Then
            UPDATE_SYS_REPORT.Fields(14).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Monitor" Then
            UPDATE_SYS_REPORT.Fields(15).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Web-Cam" Then
            UPDATE_SYS_REPORT.Fields(16).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "Stebilizer" Then
            UPDATE_SYS_REPORT.Fields(17).Value = rs_cur_invoice_item.Fields(0).Value
    ElseIf rs_cur_invoice_item.Fields(4).Value = "ZIP Drive" Then
            UPDATE_SYS_REPORT.Fields(18).Value = rs_cur_invoice_item.Fields(0).Value
    End If
    
    
        
    rs_cur_invoice_item.MoveNext
    
    Wend
    
    UPDATE_SYS_REPORT.Update
    sysnors.MoveNext
    
    Wend
    rs_cur_invoice_item.Requery
    
    While rs_cur_invoice_item.EOF <> True
        rs_cur_invoice_item.Delete
        rs_cur_invoice_item.MoveNext
    Wend
    SALE_TYPE = Clear
    FRM_AMT_PAID_NOT_PAID.Label3(5).Caption = "Sales"
    FRM_AMT_PAID_NOT_PAID.Label3(2).Caption = Text1(0).Text
    FRM_AMT_PAID_NOT_PAID.Label3(0).Caption = Combo1.Text
    FRM_AMT_PAID_NOT_PAID.dt = Format(DTPicker1.Value, "dd-MMM-yyyy")
    FRM_AMT_PAID_NOT_PAID.Label2(2).Caption = TOTAL_TRAN_AMT
    Unload Me
       ' Dim f As New FileSystemObject
       ' f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True

    
    FRM_AMT_PAID_NOT_PAID.Show vbModal
    
    Exit Sub
    End If


    
    rs_cur_invoice_item.Requery
    
    While rs_cur_invoice_item.EOF <> True
        rs_cur_invoice_item.Delete
        rs_cur_invoice_item.MoveNext
    Wend
    
    
    With Form1.salesbill
            .DataFiles(0) = App.Path & "\Master_Database.mdb"
            .ReportFileName = App.Path & "\Report\sales_bill.rpt"
            .SelectionFormula = "{Sales_master.Invoice_no} = '" & Text1(0).Text & "'"
            .username = "Admin"
            .Password = "1010101010" & Chr(10) & "1010101010"
            .Action = 1
    End With
    
    FRM_AMT_PAID_NOT_PAID.Label3(5).Caption = "Sales"
    FRM_AMT_PAID_NOT_PAID.Label3(2).Caption = Text1(0).Text
    FRM_AMT_PAID_NOT_PAID.Label3(0).Caption = Combo1.Text
    FRM_AMT_PAID_NOT_PAID.dt = Format(DTPicker1.Value, "dd-MMM-yyyy")
    Unload Me
    
    FRM_AMT_PAID_NOT_PAID.Label2(2).Caption = TOTAL_TRAN_AMT
    FRM_AMT_PAID_NOT_PAID.Show vbModal
    
    
    
    
Else
    MsgBox "There is no item in the bill, you can not save it ...", vbCritical, "No item Found ..."
    
End If


End Sub



Private Sub cmd_op_Click(Index As Integer)

If Index = 0 Then
cmd.Enabled = False
DataGrid2.Enabled = False

If SALE_TYPE <> "SYSTEM" Then
If Combo1.Enabled <> False Then
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
Else
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"



End If

Else
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
End If


    Status = True
    cmd.Enabled = False
    
    Call OP_STATUS(False)
    ENABLE_DISABLE (True)
    rs_cur_invoice_item.AddNew
    clear_box
    Combo5.Enabled = False
    Text1(2).Enabled = False
    
ElseIf Index = 1 Then
    
        If SALE_TYPE = "SYSTEM" Then
            Dim CHECK_ITEM_NAME As New ADODB.Recordset
            CHECK_ITEM_NAME.Open "SELECT * from SYS_CURRENT_SALES_ITEMS", db, adOpenDynamic, adLockOptimistic
            While CHECK_ITEM_NAME.EOF <> True
                
                If CHECK_ITEM_NAME.Fields(4).Value = Combo3.Text Then
                        If CHECK_ITEM_NAME.Fields(0).Value <> Combo2.Text Then
                                MsgBox "All System must have same " & Combo3.Text & "...", vbCritical, "Item name must be same ..."
                                Exit Sub
                        End If
                End If
            
                CHECK_ITEM_NAME.MoveNext
            
            Wend
            
        End If
        
    
    
        If VAL(Combo5.Text) <> 0 And VAL(Text1(2).Text) <> 0 Then
                Text1(3).Text = VAL(Combo5.Text) * VAL(Text1(2).Text)
        End If

    Status = False
    If Len(Combo3.Text) > 0 And Len(Combo2.Text) > 0 And VAL(Text1(5).Text) > 0 And VAL(Combo5.Text) > 0 And VAL(Text1(2).Text) > 0 Then
    rs_cur_invoice_item.Fields(0).Value = Combo2.Text
    rs_cur_invoice_item.Fields(1).Value = Combo5.Text
    rs_cur_invoice_item.Fields(2).Value = Text1(2).Text
    rs_cur_invoice_item.Fields(3).Value = Text1(3).Text
    rs_cur_invoice_item.Fields(4).Value = Combo3.Text
    rs_cur_invoice_item.Fields(5).Value = Combo4.Text
    rs_cur_invoice_item.Fields(6).Value = INVOICE_NUMBER
    rs_cur_invoice_item.Fields(7).Value = Format(DTPicker1.Value, "DD-MMM-YYYY")
    rs_cur_invoice_item.Fields(8).Value = VAL(Text1(6).Text)
    
    On Error GoTo updateerr:
    rs_cur_invoice_item.Update
    rs_cur_invoice_item.UpdateBatch
    cmd.Enabled = True
    Set DataGrid2.DataSource = Nothing
    rs_cur_invoice_item.Requery
    Set DataGrid2.DataSource = rs_cur_invoice_item
    DataGrid2.Enabled = True
    ENABLE_DISABLE (False)
    OP_STATUS (True)
If SALE_TYPE <> "SYSTEM" Then
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
Else
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
End If

    cmd.Enabled = True
    Exit Sub
updateerr:
        rs_cur_invoice_item.CancelBatch
        rs_cur_invoice_item.CancelUpdate
        MsgBox "Item is already Exist in the bill" & vbCrLf & "Update Qty in the existing item entry ...", vbCritical, "Error: Duplicate item entry ..."
        ENABLE_DISABLE (False)
        OP_STATUS (True)
        rs_cur_record_count.Requery
        If rs_cur_record_count.Fields(0).Value > 0 Then
            rs_cur_invoice_item.MoveFirst
        End If
        cmd.Enabled = True
    Else
        MsgBox "Enter Proper and Sufficient Data", vbCritical, "Check your Data ..."
    End If
    
ElseIf Index = 2 Then
    If SALE_TYPE = "SYSTEM" Then
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
     Else
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            SendKeys "{TAB}"
            
    
    End If
    
    rs_cur_record_count.Requery
    
    If rs_cur_record_count.Fields(0).Value > 0 Then
        rs_cur_invoice_item.Delete
        rs_cur_invoice_item.MoveNext
        If rs_cur_invoice_item.EOF <> True Then
            
            Call FILLTEXT
        Else
            rs_cur_record_count.Requery
            If rs_cur_record_count.Fields(0).Value > 0 Then
                    rs_cur_invoice_item.MoveLast
                    Call FILLTEXT
            Else
                clear_box
                MsgBox "All Items Deleted ...", vbInformation, "Items Deleted.."
 
                
            End If
        End If
        
        OP_STATUS (True)
    Else
        clear_box
        MsgBox "All Items Deleted ...", vbInformation, "Items Deleted.."
        
        OP_STATUS (True)
    End If
    Call Combo4_Click
ElseIf Index = 3 Then
cmd.Enabled = False
ENABLE_DISABLE (True)
OP_STATUS (False)
If SALE_TYPE = "SYSTEM" Then
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    
Else
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"


End If



ElseIf Index = 4 Then
    Status = False
    rs_cur_record_count.Requery
    rs_cur_invoice_item.CancelBatch
    rs_cur_invoice_item.CancelUpdate
    cmd.Enabled = True
    If rs_cur_record_count.Fields(0).Value > 0 Then
        rs_cur_invoice_item.MoveFirst
    Else
        clear_box
    End If
    
    Call OP_STATUS(True)
    ENABLE_DISABLE (False)
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    cmd.Enabled = True

End If



End Sub

Private Sub Combo1_Change()
If Len(Combo1.Text) = 0 Then
LaVolpeButton2.Enabled = False
End If
End Sub

Private Sub Combo1_Click()
LaVolpeButton2.Enabled = True
Dim CUST_ID_RS As New ADODB.Recordset
CUST_ID_RS.Open "SELECT * FROM Customer_master WHERE cutomer_name='" & Combo1.Text & "'", db, adOpenDynamic, adLockOptimistic
Text1(4).Text = CUST_ID_RS.Fields(0).Value
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Combo1.Text) > 0 Then
If KeyCode = 13 Then
    SendKeys "{TAB}"
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
Text1(5).Text = Clear
Text1(6).Text = Clear
FILL_QTY (0)
Text1(2).Text = Clear
Text1(3).Text = Clear
'Combo4.Text = Clear
REFRESH_COMBO (4)

Combo5.Enabled = False
Text1(2).Enabled = False


REFRESH_COMBO (4)
Set DataGrid1.DataSource = Nothing
RS_PUR_ITEMS.Close
RS_PUR_ITEMS.CursorLocation = adUseClient
RS_PUR_ITEMS.Open "SELECT Invoice_no,Qty,price_per_unit,Purchase_Date FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & Combo3.Text & "' AND Item_name='" & Combo2.Text & "' AND Party_name='" & Combo4.Text & "'", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = RS_PUR_ITEMS
DataGrid1.Enabled = True

End Sub


Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)

If Len(Combo2.Text) > 0 Then
If KeyCode = 13 Then
SendKeys "{TAB}"
End If
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo3_Click()
REFRESH_COMBO (2)
REFRESH_COMBO (4)
FILL_QTY (0)
Set DataGrid1.DataSource = Nothing
RS_PUR_ITEMS.Close
RS_PUR_ITEMS.CursorLocation = adUseClient
RS_PUR_ITEMS.Open "SELECT Invoice_no,Qty,price_per_unit,Purchase_Date FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & Combo3.Text & "' AND Item_name='" & Combo2.Text & "' AND Party_name='" & Combo4.Text & "'", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = RS_PUR_ITEMS
DataGrid1.Enabled = True

End Sub





Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Combo3.Text) > 0 Then
If KeyCode = 13 Then
SendKeys "{TAB}"
End If
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo4_Click()
FILL_QTY (0)
Set DataGrid1.DataSource = Nothing
RS_PUR_ITEMS.Close
RS_PUR_ITEMS.CursorLocation = adUseClient
RS_PUR_ITEMS.Open "SELECT Invoice_no,Qty,price_per_unit,Purchase_Date,Item_Description FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & Combo3.Text & "' AND Item_name='" & Combo2.Text & "' AND Party_name='" & Combo4.Text & "'", db, adOpenDynamic, adLockOptimistic
If RS_PUR_ITEMS.EOF <> True Then
Text1(5).Text = RS_PUR_ITEMS.Fields(1).Value
Text1(6).Text = RS_PUR_ITEMS.Fields(2).Value
If Len(Text1(5).Text) > 0 Then
    Text1(2).Enabled = True
    Combo5.Enabled = True
    FILL_QTY (VAL(Text1(5).Text))
End If
End If

Set DataGrid1.DataSource = RS_PUR_ITEMS
DataGrid1.Enabled = True
Call DataGrid1_Click
End Sub



Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)

If Len(Combo4.Text) > 0 Then
If KeyCode = 13 Then
SendKeys "{TAB}"
End If
End If

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
End Sub

Private Sub Combo5_Change()
    If VAL(Combo5.Text) <> 0 And VAL(Text1(2).Text) <> 0 Then
        Text1(3).Text = VAL(Combo5.Text) * VAL(Text1(2).Text)
    Else
        Text1(3).Text = Clear
    End If

End Sub

Private Sub Combo5_Click()
    If VAL(Combo5.Text) <> 0 And VAL(Text1(2).Text) <> 0 Then
        Text1(3).Text = VAL(Combo5.Text) * VAL(Text1(2).Text)
    Else
        Text1(3).Text = Clear
    End If

End Sub

Private Sub Combo5_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Combo5.Text) > 0 Then
If KeyCode = 13 Then
SendKeys "{TAB}"
End If
End If

End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub DataGrid1_Click()
On Error GoTo ENDS:
If GRID2_CLICKED = True Then
    INVOICE_NUMBER = rs_cur_invoice_item.Fields(6).Value
    GRID2_CLICKED = False
Else
    INVOICE_NUMBER = RS_PUR_ITEMS.Fields(0).Value
End If

Dim RS_SET_QTY_PRICE As New ADODB.Recordset
RS_SET_QTY_PRICE.Open "SELECT Qty,price_per_unit FROM AVAILABLE_PURCHASED_STOCK WHERE Invoice_no ='" & INVOICE_NUMBER & "' AND Item_type='" & Combo3.Text & "' AND Item_name='" & Combo2.Text & "' AND Party_name='" & Combo4.Text & "'", db, adOpenDynamic, adLockOptimistic
Text1(5).Text = RS_SET_QTY_PRICE.Fields(0).Value
Text1(6).Text = RS_SET_QTY_PRICE.Fields(1).Value

RS_SET_QTY_PRICE.Close
FILL_QTY (VAL(Text1(5).Text))

If Len(Text1(5).Text) > 0 Then
    Text1(2).Enabled = True
    Combo5.Enabled = True
End If
ENDS:
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If cmd_op(0).Enabled = False Then
        Dim x As Integer
        x = MsgBox("Are you sure you want to Cancel updates ...", vbQuestion Or vbYesNo, "Want to cancel ?")
        If x = 6 Then
                Call cmd_op_Click(4)
        End If
        
    Exit Sub
    End If
    
    Dim y As Integer
    y = MsgBox("Are you Sure you want to Cancel Sales Bill?", vbQuestion Or vbYesNo, "Want to Cancel Purchase Invoice ?")
    If y = 6 Then
            Unload Me
    End If
    
ElseIf KeyCode = 116 Then
            Call cmd_Click
End If


End Sub

Private Sub DataGrid2_Click()
INVOICE_NUMBER = rs_cur_invoice_item.Fields(6).Value
GRID2_CLICKED = True
Call FILLTEXT
Combo5.Enabled = False
Text1(2).Enabled = False
DataGrid1.Enabled = False
End Sub

Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If cmd_op(0).Enabled = False Then
        Dim x As Integer
        x = MsgBox("Are you sure you want to Cancel updates ...", vbQuestion Or vbYesNo, "Want to cancel ?")
        If x = 6 Then
                Call cmd_op_Click(4)
        End If
        
    Exit Sub
    End If
    
    Dim y As Integer
    y = MsgBox("Are you Sure you want to Cancel Sales Bill?", vbQuestion Or vbYesNo, "Want to Cancel Purchase Invoice ?")
    If y = 6 Then
            Unload Me
    End If
    
ElseIf KeyCode = 116 Then
            Call cmd_Click
End If


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If cmd_op(0).Enabled = False Then
        Dim x As Integer
        x = MsgBox("Are you sure you want to Cancel updates ...", vbQuestion Or vbYesNo, "Want to cancel ?")
        If x = 6 Then
                Call cmd_op_Click(4)
        End If
        
    Exit Sub
    End If
    
    Dim y As Integer
    y = MsgBox("Are you Sure you want to Cancel Sales Bill?", vbQuestion Or vbYesNo, "Want to Cancel Purchase Invoice ?")
    If y = 6 Then
            Unload Me
    End If
    
ElseIf KeyCode = 116 Then
            Call cmd_Click
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 23 Then
        Text1(0).Text = "WithoutBill_" & Now
End If
End Sub

Private Sub Form_Load()

DTPicker1.Value = Date
SYSTEM_QTY = 0
SALE_TYPE = Clear
KeyPreview = True
Text1(0).Text = SALES_INVOICE_NUMBER()
SendKeys "{TAB}"
Dim DEL_RS As New ADODB.Recordset
DEL_RS.Open "SELECT * FROM SYS_CURRENT_SALES_ITEMS", db, adOpenDynamic, adLockOptimistic

While DEL_RS.EOF <> True
    DEL_RS.Delete
    DEL_RS.MoveNext
Wend

DEL_RS.Close


    SNAME = Clear
    LaVolpeButton2.Enabled = False
    Me.Left = 0
    Me.Top = 0
    
    rs_custo_datail.Open "SELECT * FROM Customer_master", db, adOpenDynamic, adLockOptimistic
    rs_cust_system_detail.Open "SELECT * FROM Customer_System_datail", db, adOpenDynamic, adLockOptimistic
    rs_sales_master.Open "SELECT * FROM Sales_master", db, adOpenDynamic, adLockOptimistic
    
    rs_cur_invoice_item.CursorLocation = adUseClient
    rs_cur_invoice_item.Open "SELECT * FROM SYS_CURRENT_SALES_ITEMS", db, adOpenDynamic, adLockOptimistic
    Set DataGrid2.DataSource = rs_cur_invoice_item
    
    
    rs_cur_record_count.Open "SELECT COUNT(*) FROM SYS_CURRENT_SALES_ITEMS", db, adOpenDynamic, adLockOptimistic
    
    RS_PUR_ITEMS.CursorLocation = adUseClient
    RS_PUR_ITEMS.Open "SELECT Invoice_no,Qty,price_per_unit,Purchase_Date,Item_Description FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & Combo3.Text & "' AND Item_name='" & Combo2.Text & "' AND Party_name='" & Combo4.Text & "'", db, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS_PUR_ITEMS
    
    

    RS_ITEM.Open "SELECT * FROM Item_master", db, adOpenDynamic, adLockOptimistic
    rs_item_type.Open "SELECT * FROM ItemType", db, adOpenDynamic, adLockOptimistic
    RS_SALES_INVOICE_DESC.Open "SELECT * FROM sales_item_description", db, adOpenDynamic, adLockOptimistic
    
    Call REFRESH_COMBO(1)
    Call REFRESH_COMBO(2)
    Call REFRESH_COMBO(3)
    Text1(3).Enabled = False

    ENABLE_DISABLE (False)
    OP_STATUS (True)
    
    
     DataGrid2.Enabled = False
     
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmd_op(0).Enabled = False Then
        Call cmd_op_Click(4)
End If

If SALE_TYPE = "SYSTEM" Then
    Dim DEL_UPDATE As New ADODB.Recordset
    DEL_UPDATE.Open "SELECT * FROM CUSTOMER_SYSTEM_INVOICENO WHERE CUSTOMER_NAME='" & Combo1.Text & "' AND INVOICE_NO='" & Text1(0).Text & "'", db, adOpenDynamic, adLockOptimistic
    While DEL_UPDATE.EOF <> True
        DEL_UPDATE.Delete
        DEL_UPDATE.MoveNext
    Wend
    
    DEL_UPDATE.Close
End If


If SNAME = "NOT SAVED" Then
        rs_custo_datail.Close
        rs_custo_datail.Open "SELECT * FROM Customer_master WHERE cutomer_name='" & Combo1.Text & "' AND cutomer_id='" & Text1(4).Text & "'", db, adOpenDynamic, adLockOptimistic
        rs_custo_datail.Delete
End If


    rs_cust_system_detail.Close
    rs_custo_datail.Close
    rs_sales_master.Close
    RS_ITEM.Close
    rs_item_type.Close
    
    
If Status = True Then
    rs_cur_invoice_item.CancelBatch
    rs_cur_invoice_item.CancelUpdate
End If

    rs_cur_invoice_item.Close
    rs_cur_record_count.Close
    RS_PUR_ITEMS.Close
    RS_SALES_INVOICE_DESC.Close
End Sub

Private Sub LaVolpeButton1_Click()
SNAME = "NOT SAVED"
FNAME = "SALES"

Combo1.Text = Clear
Combo1.Enabled = False
LaVolpeButton2.Enabled = False
LaVolpeButton1.Enabled = False

FNAME = "SALES_NAME"

frm_cust_details.Show vbModal

End Sub

Public Sub REFRESH_COMBO(Index As Integer)

If Index = 1 Then
    Combo1.Clear
    rs_custo_datail.Requery
    While rs_custo_datail.EOF <> True
        Combo1.AddItem rs_custo_datail.Fields(1).Value
        rs_custo_datail.MoveNext
    Wend
ElseIf Index = 2 Then
    Combo2.Clear
    
    RS_ITEM.Close
    RS_ITEM.Open "SELECT * FROM Item_master WHERE Itemtype='" & Combo3.Text & "'", db, adOpenDynamic, adLockOptimistic
    While RS_ITEM.EOF <> True
        Combo2.AddItem Trim(RS_ITEM.Fields(1).Value)
        RS_ITEM.MoveNext
    Wend
ElseIf Index = 3 Then
    Combo3.Clear
    rs_item_type.Requery
    While rs_item_type.EOF <> True
        Combo3.AddItem Trim(rs_item_type.Fields(0).Value)
        rs_item_type.MoveNext
    Wend
ElseIf Index = 4 Then
    Combo4.Clear
    Dim RS_P_NAME As New ADODB.Recordset
    RS_P_NAME.Open "SELECT DISTINCT Party_name FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & Combo3.Text & "' AND Item_name='" & Combo2.Text & "'", db, adOpenDynamic, adLockOptimistic
    
    While RS_P_NAME.EOF <> True
        Combo4.AddItem RS_P_NAME.Fields(0).Value
        RS_P_NAME.MoveNext
    Wend
    RS_P_NAME.Close
    
End If
End Sub

Private Sub LaVolpeButton2_Click()
CrystalReport1.DataFiles(0) = App.Path & "\Master_Database.mdb"
CrystalReport1.ReportFileName = App.Path & "\Report\rpt_Verify_cutomer_detail.rpt"
CrystalReport1.SelectionFormula = "{Customer_master.cutomer_name} = '" & Combo1.Text & "'"
CrystalReport1.username = "Admin"
CrystalReport1.Password = "1010101010" & Chr(10) & "1010101010"
CrystalReport1.Action = 1
CrystalReport1.PageZoom (100)
End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 1 Or Index = 2 Then
    If VAL(Combo5.Text) <> 0 And VAL(Text1(2).Text) <> 0 Then
        Text1(3).Text = VAL(Combo5.Text) * VAL(Text1(2).Text)
    End If
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Len(Text1(Index)) > 0 Then
If KeyCode = 13 Then
    SendKeys "{TAB}"
End If
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Or Index = 2 Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        ElseIf KeyAscii = 46 Then
           If InStr(1, Text1(2).Text, ".", vbTextCompare) > 0 Then
                KeyAscii = 0
           End If
        ElseIf KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
End If

End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
If Index = 0 Then
    If Len(Text1(Index).Text) = 0 Then
        MsgBox "Please Enter Invoice Number ...", vbInformation, "Enter Invoice Number ..."
        Cancel = True
    End If
    
End If

If Index = 1 Or Index = 2 Then
    If VAL(Text1(Index).Text) = 0 Then
        MsgBox "Zero value is nor Allowed ...", vbInformation, "Zero Value Error..."
        Cancel = True
    End If
End If
End Sub

Public Sub FILL_QTY(VAL As Integer)
    Combo5.Clear
    For i = 1 To VAL
        Combo5.AddItem i
    Next
End Sub

Public Sub ENABLE_DISABLE(t As Boolean)
If t = True Then
    Combo3.Enabled = True
    Combo2.Enabled = True
    Combo4.Enabled = True
    Combo5.Enabled = True
    Text1(2).Enabled = True
ElseIf t = False Then
    Combo3.Enabled = False
    Combo2.Enabled = False
    Combo4.Enabled = False
    Combo5.Enabled = False
    Text1(2).Enabled = False
End If

End Sub

Public Sub OP_STATUS(t As Boolean)
If t = True Then
    DataGrid2.Enabled = True
    cmd_op(0).Enabled = True
    cmd_op(1).Enabled = False
    Dim rs_check_data As New ADODB.Recordset
    rs_check_data.Open "select count(*) from SYS_CURRENT_SALES_ITEMS", db, adOpenDynamic, adLockOptimistic
    
    If rs_check_data.Fields(0).Value > 0 Then
        cmd_op(2).Enabled = True
        cmd_op(3).Enabled = True
    Else
        cmd_op(2).Enabled = False
        cmd_op(3).Enabled = False
    End If
    cmd_op(4).Enabled = False
    rs_check_data.Close
    DataGrid1.Enabled = False
    DataGrid2.Enabled = True
    
ElseIf t = False Then
    DataGrid1.Enabled = False
    cmd_op(0).Enabled = False
    cmd_op(1).Enabled = True
    cmd_op(2).Enabled = False
    cmd_op(3).Enabled = False
    cmd_op(4).Enabled = True
    DataGrid1.Enabled = True
    DataGrid2.Enabled = False
End If

End Sub

Public Sub clear_box()
'Combo3.Text = Clear
REFRESH_COMBO (3)
REFRESH_COMBO (2)
REFRESH_COMBO (4)
'Combo2.Text = Clear
'Combo4.Text = Clear
Text1(5).Text = Clear
Text1(6).Text = Clear
'Combo5.Text = Clear
FILL_QTY (0)
Text1(2).Text = Clear
Text1(3).Text = Clear

End Sub

Public Sub FILLTEXT()
    
'On Error GoTo Last:
    RS_PUR_ITEMS.Close
    Set DataGrid1.DataSource = Nothing
    RS_PUR_ITEMS.CursorLocation = adUseClient
    RS_PUR_ITEMS.Open "SELECT Invoice_no,Qty,price_per_unit,Purchase_Date FROM AVAILABLE_PURCHASED_STOCK WHERE Item_type='" & rs_cur_invoice_item.Fields(4).Value & "' AND Item_name='" & rs_cur_invoice_item.Fields(0).Value & "' AND Party_name='" & rs_cur_invoice_item.Fields(5).Value & "'", db, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS_PUR_ITEMS
    
    While RS_PUR_ITEMS.Fields(0).Value <> INVOICE_NUMBER
        RS_PUR_ITEMS.MoveNext
    Wend
    
    Text1(5).Text = RS_PUR_ITEMS.Fields(1).Value
    Text1(6).Text = RS_PUR_ITEMS.Fields(2).Value
    
    FILL_QTY (VAL(Text1(5).Text))
    
    Combo3.Text = rs_cur_invoice_item.Fields(4).Value
    Combo2.Text = rs_cur_invoice_item.Fields(0).Value
    
    Combo4.Text = rs_cur_invoice_item.Fields(5).Value
    Combo5.Text = rs_cur_invoice_item.Fields(1).Value
    Text1(2).Text = rs_cur_invoice_item.Fields(2).Value
    Text1(3).Text = rs_cur_invoice_item.Fields(3).Value
    
    
    
    
Last:
End Sub

