VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Purchase_form 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Entry Form"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "Purchase_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9225
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Purchase_form.frx":0E42
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      Caption         =   "Current Invoice Items"
      ColumnCount     =   5
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "Total_qty"
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
      BeginProperty Column02 
         DataField       =   "Rate_p_unit"
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
      BeginProperty Column03 
         DataField       =   "Total_amt"
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
      BeginProperty Column04 
         DataField       =   "item_description"
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
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3495.118
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   9015
      Begin VB.ComboBox Combo1 
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
         Height          =   360
         ItemData        =   "Purchase_form.frx":0E57
         Left            =   1320
         List            =   "Purchase_form.frx":0E59
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   3135
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
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   855
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
         Index           =   3
         Left            =   3960
         TabIndex        =   9
         Top             =   1080
         Width           =   855
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
         Index           =   4
         Left            =   6480
         TabIndex        =   10
         Top             =   1080
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
         Height          =   885
         Index           =   5
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   1440
         Width           =   5055
      End
      Begin VB.ComboBox Combo2 
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
         Height          =   360
         ItemData        =   "Purchase_form.frx":0E5B
         Left            =   1320
         List            =   "Purchase_form.frx":0E5D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   315
         Left            =   4560
         TabIndex        =   7
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "New"
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
         MICON           =   "Purchase_form.frx":0E5F
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
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "New"
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
         MICON           =   "Purchase_form.frx":0E7B
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
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   2400
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
         MICON           =   "Purchase_form.frx":0E97
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
         Left            =   2520
         TabIndex        =   13
         Top             =   2400
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
         MICON           =   "Purchase_form.frx":0EB3
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
         Left            =   3945
         TabIndex        =   14
         Top             =   2400
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
         MICON           =   "Purchase_form.frx":0ECF
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
         Left            =   5400
         TabIndex        =   15
         Top             =   2400
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
         MICON           =   "Purchase_form.frx":0EEB
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
         Left            =   6825
         TabIndex        =   16
         Top             =   2400
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
         MICON           =   "Purchase_form.frx":0F07
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
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Qty"
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
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
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
         Index           =   5
         Left            =   2400
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
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
         Index           =   6
         Left            =   5040
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item Description"
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
         TabIndex        =   27
         Top             =   1560
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
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   9015
      Begin VB.ComboBox Combo3 
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   3735
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   48168963
         CurrentDate     =   37457
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton3 
         Height          =   315
         Left            =   5760
         TabIndex        =   3
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "New"
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
         MICON           =   "Purchase_form.frx":0F23
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Party Name"
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
         TabIndex        =   23
         Top             =   600
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
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
   End
   Begin LVbuttons.LaVolpeButton cmd_update 
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   6480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save Bill (F5)"
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
      MICON           =   "Purchase_form.frx":0F3F
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Entry Form"
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
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label issues 
      BackStyle       =   0  'Transparent
      Caption         =   $"Purchase_form.frx":0F5B
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
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   8895
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   -120
      Picture         =   "Purchase_form.frx":0FE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9360
   End
End
Attribute VB_Name = "Purchase_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pur_rs As New ADODB.Recordset
Dim item_rs As New ADODB.Recordset
Dim item_type As New ADODB.Recordset
Dim rs_cur_invoice_item As New ADODB.Recordset
Dim rs_grid As New ADODB.Recordset
Dim rs_cur_record_count As New ADODB.Recordset
Dim pname As New ADODB.Recordset
Dim Status As Boolean
Public PADD As Boolean
Public TOTAL_TRAN_AMT As Double




Private Sub cmd_op_Click(Index As Integer)
If Index = 0 Then
If Combo3.Enabled = False Then
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    
Else
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
End If

    
    
    Call opbutton_status(False)
    ENABLE_DISABLE (True)
    rs_cur_invoice_item.AddNew
    clear_box
    Status = True
    cmd_update.Enabled = False
    
ElseIf Index = 1 Then
    
    If Len(Combo2.Text) > 0 And Len(Combo1.Text) > 0 And VAL(Text1(2).Text) > 0 And VAL(Text1(3).Text) > 0 And VAL(Text1(4).Text) > 0 Then
    
    rs_cur_invoice_item.Fields(0).Value = Combo2.Text
    rs_cur_invoice_item.Fields(1).Value = Combo1.Text
    rs_cur_invoice_item.Fields(2).Value = Text1(2).Text
    rs_cur_invoice_item.Fields(3).Value = Text1(3).Text
    rs_cur_invoice_item.Fields(4).Value = Text1(4).Text
    rs_cur_invoice_item.Fields(5).Value = Text1(5).Text
    
    On Error GoTo updateerr:
    rs_cur_invoice_item.Update
    rs_cur_invoice_item.UpdateBatch
    Set DataGrid1.DataSource = Nothing
    rs_cur_invoice_item.Requery
    Set DataGrid1.DataSource = rs_cur_invoice_item
    ENABLE_DISABLE (False)
    opbutton_status (True)
    Status = False
    cmd_update.Enabled = True
    If Combo3.Enabled = False Then
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
    End If
    
    Exit Sub
updateerr:
        rs_cur_invoice_item.CancelBatch
        rs_cur_invoice_item.CancelUpdate
        rs_cur_invoice_item.Requery
        MsgBox "Item is already Exist in the bill" & vbCrLf & "Update Qty in the existing item entry ...", vbCritical, "Error: Duplicate item entry ..."
        ENABLE_DISABLE (False)
        opbutton_status (True)
        Status = False
        cmd_update.Enabled = True
    Else
        MsgBox "Enter Proper and Sufficient Data", vbCritical, "Check your Data ..."
    End If
    
ElseIf Index = 2 Then
    
    
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
                    If Combo1.Enabled = False Then
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
                    End If
                    
            End If
        End If
    Else
        clear_box
        MsgBox "All Items Deleted ...", vbInformation, "Items Deleted.."
    End If
    
ElseIf Index = 3 Then
rs_cur_record_count.Requery
If rs_cur_record_count.Fields(0).Value > 0 Then
    ENABLE_DISABLE (True)
    opbutton_status (False)
    cmd_update.Enabled = False
End If

ElseIf Index = 4 Then
    
    
    Status = False
    rs_cur_record_count.Requery
    rs_cur_invoice_item.CancelBatch
    rs_cur_invoice_item.CancelUpdate
    rs_cur_invoice_item.Requery
    If rs_cur_record_count.Fields(0).Value > 0 Then
        rs_cur_invoice_item.MoveFirst
    Else
        clear_box
    End If
    
    Call opbutton_status(True)
    ENABLE_DISABLE (False)
    cmd_update.Enabled = True
    
End If

End Sub

Private Sub cmd_op_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Then
If KeyCode = 13 Then
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
    SendKeys "{TAB}"
End If
End If

End Sub

Private Sub cmd_update_Click()
Dim t As Integer
t = MsgBox("Are you sure you want to save purchase bill", vbQuestion Or vbYesNo, "Want to save Purchase bill")
If t = 7 Then
    Exit Sub
End If




Dim RS_AVA_PU_STOCK As New ADODB.Recordset
RS_AVA_PU_STOCK.Open "SELECT * FROM AVAILABLE_PURCHASED_STOCK", db, adOpenDynamic, adLockOptimistic

    rs_cur_invoice_item.Requery
    rs_cur_record_count.Requery
    If rs_cur_record_count.Fields(0).Value > 0 Then
    
    If Len(Text1(0).Text) > 0 And Len(Combo3.Text) > 0 Then
    TOTAL_TRAN_AMT = TOTAL_AMT("PURCHASE")
    
    While rs_cur_invoice_item.EOF <> True

                
        pur_rs.AddNew
        pur_rs.Fields(0).Value = Text1(0).Text
        pur_rs.Fields(1).Value = Combo3.Text
        pur_rs.Fields(2).Value = DTPicker1.Value
        pur_rs.Fields(3).Value = rs_cur_invoice_item.Fields(0).Value
        pur_rs.Fields(4).Value = rs_cur_invoice_item.Fields(1).Value
        pur_rs.Fields(5).Value = rs_cur_invoice_item.Fields(2).Value
        pur_rs.Fields(6).Value = rs_cur_invoice_item.Fields(3).Value
        pur_rs.Fields(7).Value = rs_cur_invoice_item.Fields(4).Value
        If Len(rs_cur_invoice_item.Fields(5).Value) > 0 Then
                pur_rs.Fields(8).Value = rs_cur_invoice_item.Fields(5).Value
        End If
        
        On Error GoTo OH_ER
        pur_rs.Update
        GoTo A1:
OH_ER:
        MsgBox "Duplicate Entry Found ...", vbCritical, "Duplicate Entry Found ..."
        pur_rs.CancelUpdate
        Exit Sub
A1:
        RS_AVA_PU_STOCK.AddNew
        RS_AVA_PU_STOCK.Fields(0).Value = Text1(0).Text
        RS_AVA_PU_STOCK.Fields(1).Value = Combo3.Text
        RS_AVA_PU_STOCK.Fields(2).Value = DTPicker1.Value
        RS_AVA_PU_STOCK.Fields(3).Value = rs_cur_invoice_item.Fields(0).Value
        RS_AVA_PU_STOCK.Fields(4).Value = rs_cur_invoice_item.Fields(1).Value
        RS_AVA_PU_STOCK.Fields(5).Value = rs_cur_invoice_item.Fields(2).Value
        RS_AVA_PU_STOCK.Fields(6).Value = rs_cur_invoice_item.Fields(3).Value
        RS_AVA_PU_STOCK.Fields(7).Value = rs_cur_invoice_item.Fields(4).Value
        RS_AVA_PU_STOCK.Fields(8).Value = rs_cur_invoice_item.Fields(5).Value
        
        RS_AVA_PU_STOCK.Update
        
    
        item_rs.Close
        item_rs.Open "select * from item_master where Itemtype='" & rs_cur_invoice_item.Fields(0).Value & "' and Item_name='" & rs_cur_invoice_item.Fields(1).Value & "'", db, adOpenDynamic, adLockOptimistic
        item_rs.Fields(3).Value = VAL(item_rs.Fields(3).Value) + VAL(rs_cur_invoice_item.Fields(2).Value)
        item_rs.Update
        
        rs_cur_invoice_item.MoveNext
    Wend
    
    
    rs_cur_invoice_item.MoveFirst
    
    While rs_cur_invoice_item.EOF <> True
        rs_cur_invoice_item.Delete
        rs_cur_invoice_item.MoveNext
    Wend
    PADD = False
    
    FRM_AMT_PAID_NOT_PAID.Label3(5).Caption = "Purchase"
    FRM_AMT_PAID_NOT_PAID.Label3(2).Caption = Text1(0).Text
    FRM_AMT_PAID_NOT_PAID.Label3(0).Caption = Combo3.Text
    FRM_AMT_PAID_NOT_PAID.dt = Format(DTPicker1.Value, "dd-MMM-yyyy")
    FRM_AMT_PAID_NOT_PAID.Label2(2).Caption = TOTAL_TRAN_AMT
    Unload Me
        'Dim f As New FileSystemObject
        'f.CopyFile App.Path & "\Master_Database.mdb", App.Path & "\data\" & cur_company_name & "\Master_Database.mdb", True
       
    FRM_AMT_PAID_NOT_PAID.Show vbModal
    Else
        MsgBox "Enter Proper data" & vbCrLf & "Some Important Data are missing", vbCritical, "Enter Proper Data ..."
    End If
    
    Else
        MsgBox "There is no item in this Purchase bill , You can not save it ...", vbInformation, "No item Found.."
    
    End If
    RS_AVA_PU_STOCK.Close
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Combo1.Text) > 0 Then
If KeyCode = 13 Then
    SendKeys "{TAB}"
    SendKeys "{TAB}"
End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo2_Click()
Refresh_combobox (1)
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Combo2.Text) > 0 Then
    If KeyCode = 13 Then
        SendKeys "{TAB}"
        SendKeys "{TAB}"
    End If
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'KeyAscii = 0

End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(Combo3.Text) > 0 Then
        If KeyCode = 13 Then
            SendKeys "{TAB}"
            SendKeys "{TAB}"
        End If
End If
End Sub

Private Sub DataGrid1_Click()
Call FILLTEXT
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
    y = MsgBox("Are you Sure you want to Cancel Purchase Bill?", vbQuestion Or vbYesNo, "Want to Cancel Purchase Invoice ?")
    If y = 6 Then
            Unload Me
    End If

ElseIf KeyCode = 116 Then
    Call cmd_update_Click
End If

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    If cmd_op(0).Enabled = False Then
        x = MsgBox("Are you sure you want to Cancel updates ...", vbQuestion Or vbYesNo, "Want to cancel ?")
        If x = 6 Then
                Call cmd_op_Click(4)
        End If
        Exit Sub
    End If
    
    Dim y As Integer
    y = MsgBox("Are you Sure you want to Cancel Purchase Bill?", vbQuestion Or vbYesNo, "Want to Cancel Purchase Invoice ?")
    If y = 6 Then
            Unload Me
    End If

ElseIf KeyCode = 116 Then
    Call cmd_update_Click
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 23 Then
        Text1(0).Text = "WithoutBill_" & Now
End If
End Sub

Private Sub Form_Load()

DTPicker1.Value = Date

Dim DEL_RS As New ADODB.Recordset
DEL_RS.Open "SELECT * FROM SYS_CURRENT_INVOICE", db, adOpenDynamic, adLockOptimistic

While DEL_RS.EOF <> True
    DEL_RS.Delete
    DEL_RS.MoveNext
Wend

DEL_RS.Close
KeyPreview = True
PADD = False

Me.Left = 0
Me.Top = 0

pur_rs.Open "select * from Purchase_master", db, adOpenDynamic, adLockOptimistic
item_rs.Open "select * from item_master", db, adOpenDynamic, adLockOptimistic
item_type.Open "select * from ItemType", db, adOpenDynamic, adLockOptimistic
rs_cur_invoice_item.CursorLocation = adUseClient
rs_cur_invoice_item.Open "select * from SYS_CURRENT_INVOICE", db, adOpenDynamic, adLockOptimistic
rs_cur_record_count.Open "select count(*) from SYS_CURRENT_INVOICE", db, adOpenDynamic, adLockOptimistic
pname.Open "select * from purchase_partynames", db, adOpenDynamic, adLockOptimistic

Combo3.Clear
While pname.EOF <> True
    
    Combo3.AddItem pname.Fields(0).Value
    pname.MoveNext
Wend


Set DataGrid1.DataSource = rs_cur_invoice_item
Call Refresh_combobox(2)

Call ENABLE_DISABLE(False)
opbutton_status (True)
A1:

End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmd_op(0).Enabled = False Then
        Call cmd_op_Click(4)
End If


If PADD = True Then
    Dim R As New ADODB.Recordset
        R.Open "SELECT * FROM purchase_partynames WHERE purchase_partyname='" & Combo3.Text & "'", db, adOpenDynamic, adLockOptimistic
    R.Delete
    R.Close
End If

pur_rs.Close
item_rs.Close
item_type.Close
pname.Close

rs_cur_record_count.Requery
If rs_cur_record_count.Fields(0).Value > 0 Then
    rs_cur_invoice_item.Requery
    While rs_cur_invoice_item.EOF <> True
        rs_cur_invoice_item.Delete
        rs_cur_invoice_item.MoveNext
    Wend
    
End If

If Status = True Then
rs_cur_invoice_item.CancelBatch
rs_cur_invoice_item.CancelUpdate
End If

rs_cur_invoice_item.Close
rs_cur_record_count.Close
End Sub

Private Sub LaVolpeButton1_Click()
'Combo1.Text = Clear
Refresh_combobox (1)


If Len(Combo2.Text) <> 0 Then
    add_Item_form.FN = "PR"
    add_Item_form.Show vbModal
Else
    MsgBox "Select Item Type to add new item ...", vbInformation, "Select Item Type ..."
    Combo2.SetFocus
End If
End Sub

Private Sub LaVolpeButton2_Click()

'Combo2.Text = Clear
Refresh_combobox (2)
frm_item_type.FN = "PR"

frm_item_type.Show vbModal
End Sub

Private Sub LaVolpeButton3_Click()
Combo3.Enabled = False
PADD = True

LaVolpeButton3.Enabled = False
frm_party_name.Show vbModal
End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 2 Or Index = 3 Then
    If VAL(Text1(2).Text) <> 0 And VAL(Text1(3).Text) <> 0 Then
        Text1(4).Text = VAL(Text1(2).Text) * VAL(Text1(3).Text)
    End If
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If KeyAscii = 34 Or KeyAscii = 39 Then
            KeyAscii = 0
        End If
    ElseIf Index = 1 Then
        If KeyAscii >= 97 And KeyAscii <= 122 Then
        ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then
        ElseIf KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    ElseIf Index = 2 Then
        
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        ElseIf KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    ElseIf Index = 3 Then
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



Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
If Index = 0 Then
    If Len(Text1(0).Text) = 0 Then
        Cancel = True
    End If
ElseIf Index = 1 Then
    If Len(Combo3.Text) = 0 Then
        Cancel = True
    End If
ElseIf Index = 2 Then
        If Len(Text1(Index).Text) = 0 Then
            Text1(Index).Text = 0
        End If
    
ElseIf Index = 3 Then
        If Len(Text1(Index).Text) = 0 Then
            Text1(Index).Text = 0
        End If
End If
End Sub

Public Sub Refresh_combobox(Index As Integer)
    If Index = 2 Then
    Combo2.Clear
    item_type.Requery
    While item_type.EOF <> True
            Combo2.AddItem item_type(0).Value
            item_type.MoveNext
    Wend
    
    ElseIf Index = 1 Then
    Combo1.Clear
    item_rs.Close
    item_rs.Open "select * from item_master where Itemtype ='" & Combo2.Text & "'", db, adOpenDynamic, adLockOptimistic
    While item_rs.EOF <> True
            Combo1.AddItem item_rs.Fields(1).Value
            item_rs.MoveNext
    Wend
    
    ElseIf Index = 3 Then
    Combo3.Clear
    pname.Requery
    While pname.EOF <> True
        Combo3.AddItem pname.Fields(0).Value
        pname.MoveNext
    Wend

    End If
End Sub

Public Function ENABLE_DISABLE(t As Boolean)
If t = True Then
    Combo2.Enabled = True
    Combo1.Enabled = True
    LaVolpeButton1.Enabled = True
    LaVolpeButton2.Enabled = True
    Text1(2).Enabled = True
    Text1(3).Enabled = True
    Text1(5).Enabled = True
ElseIf t = False Then
    Combo2.Enabled = False
    Combo1.Enabled = False
    LaVolpeButton1.Enabled = False
    LaVolpeButton2.Enabled = False
    Text1(2).Enabled = False
    Text1(3).Enabled = False
    Text1(5).Enabled = False
End If
End Function

Public Sub opbutton_status(t As Boolean)
If t = True Then
    DataGrid1.Enabled = True
    cmd_op(0).Enabled = True
    cmd_op(1).Enabled = False
    Dim rs_check_data As New ADODB.Recordset
    rs_check_data.Open "select count(*) from SYS_CURRENT_INVOICE", db, adOpenDynamic, adLockOptimistic
    
    If rs_check_data.Fields(0).Value > 0 Then
        cmd_op(2).Enabled = True
        cmd_op(3).Enabled = True
    Else
        cmd_op(2).Enabled = False
        cmd_op(3).Enabled = False
    End If
    cmd_op(4).Enabled = False
    rs_check_data.Close
    
ElseIf t = False Then
    DataGrid1.Enabled = False
    cmd_op(0).Enabled = False
    cmd_op(1).Enabled = True
    cmd_op(2).Enabled = False
    cmd_op(3).Enabled = False
    cmd_op(4).Enabled = True
End If

End Sub

Public Sub clear_box()
Refresh_combobox (2)
Refresh_combobox (1)

Text1(2).Text = Clear
Text1(3).Text = Clear
Text1(4).Text = Clear
Text1(5).Text = Clear
End Sub

Public Sub FILLTEXT()
Combo2.Text = rs_cur_invoice_item.Fields(0).Value
Combo1.Text = rs_cur_invoice_item.Fields(1).Value
Text1(2).Text = rs_cur_invoice_item.Fields(2).Value
Text1(3).Text = rs_cur_invoice_item.Fields(3).Value
Text1(4).Text = rs_cur_invoice_item.Fields(4).Value
Text1(5).Text = rs_cur_invoice_item.Fields(5).Value
End Sub
