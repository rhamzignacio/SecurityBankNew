VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SBTC"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   7890
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&PROCESS NOW!"
      DisabledPicture =   "frmMain.frx":7B65A
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      Picture         =   "frmMain.frx":7B99B
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Click this to process data from Head Folder"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse13 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&BROWSE 13 DIGITS"
      DisabledPicture =   "frmMain.frx":7BCDC
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      Picture         =   "frmMain.frx":7C192
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Click this to process data from Head Folder"
      Top             =   600
      Width           =   855
   End
   Begin VB.ListBox lstErrors 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   120
      TabIndex        =   38
      Top             =   2040
      Width           =   7335
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&CANCEL"
      DisabledPicture =   "frmMain.frx":7C648
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      Picture         =   "frmMain.frx":7CB05
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtItemType 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   32
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBatchno 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   4200
      MaxLength       =   8
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&DELETE"
      DisabledPicture =   "frmMain.frx":7CFC2
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      Picture         =   "frmMain.frx":7D497
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdLength 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&LENGTH"
      DisabledPicture =   "frmMain.frx":7D96C
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      Picture         =   "frmMain.frx":7DABD
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox txtRtno 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   3
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtAcctno 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1320
      MaxLength       =   12
      TabIndex        =   5
      Top             =   6600
      Width           =   4095
   End
   Begin VB.TextBox txtAcctnm1 
      BackColor       =   &H00C0FFFF&
      DataField       =   "Name1"
      DataSource      =   "DBFRecordset"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   1
      Top             =   5520
      Width           =   6015
   End
   Begin VB.TextBox txtAcctnm2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   2
      Top             =   5880
      Width           =   6015
   End
   Begin VB.TextBox txtChktype 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   6480
      MaxLength       =   1
      TabIndex        =   8
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox txtOrderqty 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   6480
      MaxLength       =   3
      TabIndex        =   6
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&EDIT"
      DisabledPicture =   "frmMain.frx":7DC0E
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "frmMain.frx":7E0B5
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&GENERATE"
      DisabledPicture =   "frmMain.frx":7E55C
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      Picture         =   "frmMain.frx":7EA05
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F1C4B6&
      Caption         =   "&SAVE"
      DisabledPicture =   "frmMain.frx":7EEAE
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Picture         =   "frmMain.frx":7F355
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowse10 
      BackColor       =   &H00F1C4B6&
      Caption         =   "Starter"
      DisabledPicture =   "frmMain.frx":7F7FC
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":7FD5D
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click this to process data from Head Folder"
      Top             =   600
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid grdDisplay 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   12648447
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "Chktype"
         Caption         =   "CT"
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
         DataField       =   "Rtno"
         Caption         =   "Routing no."
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
         DataField       =   "Acctno"
         Caption         =   "Account no."
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
         DataField       =   "Acctnm1"
         Caption         =   "Acct Name1"
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
         DataField       =   "Acctnm2"
         Caption         =   "Account Name2"
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
         DataField       =   "Formtype"
         Caption         =   "FT"
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
         DataField       =   "Orderqty"
         Caption         =   "Orderqty"
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
         DataField       =   "Branchname"
         Caption         =   "Branchname"
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
         DataField       =   "Address1"
         Caption         =   "Address1"
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
         DataField       =   "Address2"
         Caption         =   "Address2"
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
         DataField       =   "Address3"
         Caption         =   "Address3"
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
         DataField       =   "Batchno"
         Caption         =   "Batchno"
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
      BeginProperty Column12 
         DataField       =   "Status"
         Caption         =   "Status"
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
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4215.118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3044.977
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   299.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3960
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   4350.047
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   3674.835
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   3225.26
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column12 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPrimaryKey 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstGCCLick 
      Appearance      =   0  'Flat
      BackColor       =   &H00753E2B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -120
      TabIndex        =   42
      Top             =   7080
      Width           =   8655
   End
   Begin VB.Label lblBooksCA10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   50
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblBooksPA10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   49
      Top             =   1080
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   600
      Y2              =   1920
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "CA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   48
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "PA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   47
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "10 DIGITS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   46
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "13 DIGITS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   45
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image img13 
      Height          =   240
      Left            =   1320
      Picture         =   "frmMain.frx":802BE
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image img10 
      Height          =   240
      Left            =   360
      Picture         =   "frmMain.frx":803AB
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Bank Corporation"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FileName:"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   4200
      TabIndex        =   40
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      Caption         =   "FileName:"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   5280
      TabIndex        =   39
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblTotalAccounts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   37
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblGCBooks13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   36
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "GC:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   35
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "MC:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   34
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMCBooks13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   33
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label14 
      BackColor       =   &H007774F8&
      BackStyle       =   0  'Transparent
      Caption         =   "Batch No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lblMC 
      BackColor       =   &H0070F1EE&
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label lblGC 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   4920
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   1575
      Left            =   120
      Top             =   5400
      Width           =   7335
   End
   Begin VB.Label Label12 
      BackColor       =   &H007774F8&
      BackStyle       =   0  'Transparent
      Caption         =   "BRSTN:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H007774F8&
      BackStyle       =   0  'Transparent
      Caption         =   "Name 1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H007774F8&
      BackStyle       =   0  'Transparent
      Caption         =   "Name 2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H007774F8&
      BackStyle       =   0  'Transparent
      Caption         =   "Account No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H007774F8&
      BackStyle       =   0  'Transparent
      Caption         =   "Type A/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   23
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H007774F8&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Qty:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   6240
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   1335
      Left            =   3000
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL ACCOUNTS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL BOOKS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblTotalBooks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblBooksPA13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblBooksCA13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Books"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputFile, FilePath, FileLocation, PathOfFile As String
Dim Batchno, EditAccountNumber, EditChktype As String
Dim PageSeparator, PageNo, GrandTotal, TotalBooks As String
Dim TransPStartSerial, TransPEndSerial, EditPrimaryKey As String
Dim SeparatorCount, SeparatorValue, Counter, Incrementor As String
Dim BatchnoGC13, BatchnoMC13, BatchnoReg10, BatchnoReg13, ItemType, KeyAsciiFinal As String
Dim VerifyPass As Boolean
Dim EditYou As Boolean

Dim DoubleAccount_Total As String
Dim DoubleAccount_CheckType(0 To 999999) As String
Dim DoubleAccount_Accounts(0 To 999999) As String

Sub LimitProgramReg() '--> REGULAR!

BatchnoReg = UCase(BatchnoReg10 & "," & BatchnoReg13)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection
Set Rs = New ADODB.Recordset

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "UPDATE LastRefModified SET [Date] = '" & ExistingRefDateModified & "', [Time] = '" & ExistingRefTimeModified & "'"
Rs.Open strQuery, Conn, adOpenStatic

strQuery = "INSERT INTO LastRun ([Date], [Time], [Batch]) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & Format(Now, "HH:MM:SS") & "','" & BatchnoReg & "')"
Rs.Open strQuery, Conn, adOpenStatic

strQuery = "INSERT INTO Archive ([Date], [FileName]) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & InputFile & "')"
Rs.Open strQuery, Conn, adOpenStatic
End Sub

Sub LimitProgramMC() '--> MC!
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\MC\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection
Set Rs = New ADODB.Recordset

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\MC\MC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "UPDATE LastRefModified SET [Date] = '" & ExistingRefDateModified & "', [Time] = '" & ExistingRefTimeModified & "'"
Rs.Open strQuery, Conn, adOpenStatic

strQuery = "INSERT INTO LastRun ([Date], [Time], [Batch]) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & Format(Now, "HH:MM:SS") & "','" & BatchnoMC13 & "')"
Rs.Open strQuery, Conn, adOpenStatic

strQuery = "INSERT INTO Archive ([Date], [FileName]) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & InputFile & "')"
Rs.Open strQuery, Conn, adOpenStatic
End Sub

Sub LimitProgramGC() '--> GC!
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\GIFTCHK\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection
Set Rs = New ADODB.Recordset

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\GIFTCHK\GIFTCHK.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "UPDATE LastRefModified SET [Date] = '" & ExistingRefDateModified & "', [Time] = '" & ExistingRefTimeModified & "'"
Rs.Open strQuery, Conn, adOpenStatic

strQuery = "INSERT INTO LastRun ([Date], [Time], [Batch]) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & Format(Now, "HH:MM:SS") & "','" & BatchnoGC13 & "')"
Rs.Open strQuery, Conn, adOpenStatic

strQuery = "INSERT INTO Archive ([Date], [FileName]) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & InputFile & "')"
Rs.Open strQuery, Conn, adOpenStatic
End Sub



Sub VerifyRouting10()

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTC_10"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    ChkType = dbfRecordset.Fields(0)
    Rtno = dbfRecordset.Fields(1)
    Acctno = dbfRecordset.Fields(2)
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    
    Set Conn = New ADODB.Connection
    
    With Conn
      .CursorLocation = adUseClient
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & FileLocation & "; Jet OLEDB;"
      .Open
    End With
    
    strQuery = "SELECT * FROM BRANCH WHERE [Routing Number] = '" & Rtno & "'"
    
    Set Rs = New ADODB.Recordset
    
    With Rs
    
    Set .ActiveConnection = Conn
        .CursorType = adOpenStatic
        .Source = strQuery
        .Open
    End With
    
    If Rs.RecordCount < 1 Then
        MsgBox "Routing number " & Rtno & " is not valid, this input file cannot be proccess since it does not exists on Checkdat.mdb", vbInformation, "Error"
        VerifyPass = False
        Exit Sub
    Else
        VerifyPass = True
    End If

    'CHECK IF ROUTING EXISTS REF.DBF
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    If ChkType = "B" And Mid(Acctno, 5, 8) = "21100608" Then 'MC Personal
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
    
    ElseIf ChkType = "B" And Mid(Acctno, 5, 8) = "21200608" Then  'GC Commercial
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
    
    Else 'Regular
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    End If
    
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM REF WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    
    If DBFRecordset1.RecordCount < 1 Then
        MsgBox "Routing Number " & Rtno & " does not exists on REF.dbf, please check your data.", vbInformation, "Error"
        End
    End If
    'END CHECK IF ROUTING EXISTS REF.DBF

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
    Me.Caption = "Verifying Routing No " & FormatNumber((LoopCount / dbfRecordset.RecordCount) * 100, 5)
Loop
End Sub

Sub VerifyRouting13()

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTC_13"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    ChkType = dbfRecordset.Fields(0)
    Rtno = dbfRecordset.Fields(1)
    Acctno = dbfRecordset.Fields(2)
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    
    Set Conn = New ADODB.Connection
    
    With Conn
      .CursorLocation = adUseClient
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & FileLocation & "; Jet OLEDB;"
      .Open
    End With
    
    strQuery = "SELECT * FROM BRANCH WHERE [Routing Number] = '" & Rtno & "'"
    
    Set Rs = New ADODB.Recordset
    
    With Rs
    
    Set .ActiveConnection = Conn
        .CursorType = adOpenStatic
        .Source = strQuery
        .Open
    End With
    
    If Rs.RecordCount < 1 Then
        MsgBox "Routing number " & Rtno & " is not valid, or empty. This input file cannot be proccess since it doesn't exists on Checkdat.mdb", vbInformation, "Error"
        VerifyPass = False
        Exit Sub
    Else
        VerifyPass = True
    End If

    'CHECK IF ROUTING EXISTS REF.DBF
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    If ChkType = "B" And Mid(Acctno, 5, 8) = "21100608" Then 'MC Personal
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
    
    ElseIf ChkType = "B" And Mid(Acctno, 5, 8) = "21200608" Then  'GC Commercial
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
    
    Else 'Regular
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    End If
    
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM REF WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    
    If DBFRecordset1.RecordCount < 1 Then
        MsgBox "Routing Number " & Rtno & " does not exists on REF.dbf, please check your data.", vbInformation, "Error"
        End
    End If
    'END CHECK IF ROUTING EXISTS REF.DBF

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
    Me.Caption = "Verifying BRSTN: " & Rtno
Loop
End Sub

Sub CheckData()

Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")

If Mid(InputFile, 1, 2) = 10 Then
    SQL = "SELECT Accountno FROM SBTC_10"
Else
    SQL = "SELECT Accountno FROM SBTC_13"
End If

dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Acctno = dbfRecordset.Fields(0)
    
    If Mid(Acctno, 5, 8) = "21100608" Then 'MC
        lblMC.Caption = "MC"
    End If
    
    If Mid(Acctno, 5, 8) = "21200608" Then 'GC
        lblGC.Caption = "GC"
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
End Sub

Sub InsertNew_10()

'FOR REGULAR 10 DIGITS
LoopMyBatch:
    BatchnoReg10 = InputBox("Enter BATCH NUMBER FOR REGULAR 10 DIGITS:", "Batch Number Regular")
    
    If BatchnoReg10 = "" Then GoTo LoopMyBatch
    If Len(BatchnoReg10) <> 8 Then
        MsgBox "Please enter correct Batch Number", vbInformation, "Error"
        GoTo LoopMyBatch
    
    BatchnoReg10 = UCase(BatchnoReg10)
    End If

'GET SBTC DETAILS
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTC_10"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If dbfRecordset.RecordCount > 0 Then
    LoopCount = 0
    Do Until LoopCount = dbfRecordset.RecordCount
        
        'FIELDS
        CT = UCase(dbfRecordset.Fields(0))
        Rtno = dbfRecordset.Fields(1)
        Acctno = dbfRecordset.Fields(2)

        If Len(dbfRecordset.Fields(3)) >= 1 Then
            Acctnm1 = UCase(dbfRecordset.Fields(3))
        Else
            Acctnm1 = ""
        End If
        
        If Len(dbfRecordset.Fields(4)) >= 1 Then
            Acctnm2 = UCase(dbfRecordset.Fields(4))
        Else
            Acctnm2 = ""
        End If
        
        Orderqty = Val(dbfRecordset.Fields(5))
        PrimaryKey = getNewPrimaryKey
    
    'GET BRANCH NAME CHECKDAT.MDB
        Dim Conn As ADODB.Connection
        Dim Rs As ADODB.Recordset
        
        Set Conn = New ADODB.Connection
        
        With Conn
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & FileLocation & "; Jet OLEDB;"
          .Open
        End With
        
        strQuery = "SELECT * FROM BRANCH WHERE [Routing Number] = '" & Rtno & "'"
        
        Set Rs = New ADODB.Recordset
        
        With Rs
        
        Set .ActiveConnection = Conn
            .CursorType = adOpenStatic
            .Source = strQuery
            .Open
        End With
        
        If Len(Rs.Fields(2)) >= 1 Then
            BranchName = Replace(Rs.Fields(2), "'", "''")
        Else
            BranchName = ""
        End If
        
        If Len(Rs.Fields(3)) >= 1 Then
            Address1 = Replace(Rs.Fields(3), "'", "''")
        Else
            Address1 = ""
        End If
        
        If Len(Rs.Fields(4)) >= 1 Then
            Address2 = Replace(Rs.Fields(4), "'", "''")
        Else
            Address2 = ""
        End If
        
        If Len(Rs.Fields(5)) >= 1 Then
            Address3 = Replace(Rs.Fields(5), "'", "''")
        Else
            Address3 = ""
        End If
    'END GET BRANCH NAME
        
        If CT = "A" Then
            ItemType = "PA"
            Formtype = "05"
            ChkType = "A"
            Batchno = UCase(BatchnoReg10)
        End If
        
        If CT = "B" Then
            ItemType = "CA"
            Formtype = "16"
            ChkType = "B"
            Batchno = UCase(BatchnoReg10)
        End If
        
        'INSERT TO SBTC NEW
        Set DBFConnector1 = CreateObject("ADODB.Connection")
        
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
        
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        SQL1 = "INSERT INTO SBTCNEW (Chktype,Rtno,Acctno,Acctnm1,Acctnm2,Orderqty,Formtype,Batchno,Branchname,Address1,Address2,Address3,ItemType,PrimaryKey,Digits) VALUES ('" _
                & ChkType & "','" & Rtno & "','" & Acctno & "','" & Replace(Acctnm1, "'", "''") & "','" & Replace(Acctnm2, "'", "''") & "','" & Orderqty & "','" & Formtype & "','" _
                & Batchno & "','" & BranchName & "','" & Address1 & "','" & Address2 & "','" & Address3 & "','" & ItemType & "','" & PrimaryKey & "','10')"
        
            DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
        'END INSERT TO SBTC NEW
    
        dbfRecordset.MoveNext
        LoopCount = LoopCount + 1
        
        Me.Caption = "Creating new SBTC..." & Acctno
    Loop
    
    If Dir$(App.Path & "\SBTC_10.dbf") <> "" Then Kill (App.Path & "\SBTC_10.dbf")
End If
End Sub

Sub InsertNew_13()

'FOR REGULAR 10 DIGITS
LoopMyBatch:
    BatchnoReg13 = InputBox("Enter BATCH NUMBER FOR REGULAR 13 DIGITS:", "Batch Number Regular")
    
    If BatchnoReg13 = "" Then GoTo LoopMyBatch
    If Len(BatchnoReg13) <> 8 Then
        MsgBox "Please enter correct Batch Number", vbInformation, "Error"
        GoTo LoopMyBatch
    
    BatchnoReg13 = UCase(BatchnoReg13)
    End If

'FOR MC
If lblMC.Caption = "MC" Then
LoopMyBatchMC:
    BatchnoMC13 = InputBox("Enter BATCH NUMBER FOR MC 13 DIGITS:", "Batch Number MC")
    
    If BatchnoMC13 = "" Then GoTo LoopMyBatchMC
    If Len(BatchnoMC13) <> 8 Then
        MsgBox "Please enter correct Batch Number", vbInformation, "Error"
        GoTo LoopMyBatchMC
    End If

BatchnoMC13 = UCase(BatchnoMC13)
End If

'FOR MC
If lblGC.Caption = "GC" Then
LoopMyBatchGC:
    BatchnoGC13 = InputBox("Enter BATCH NUMBER FOR GC 13 DIGITS:", "Batch Number GC")
    
    If BatchnoGC13 = "" Then GoTo LoopMyBatchGC
    If Len(BatchnoGC13) <> 8 Then
        MsgBox "Please enter correct Batch Number", vbInformation, "Error"
        GoTo LoopMyBatchGC
    End If
    
BatchnoGC13 = UCase(BatchnoGC13)
End If

'GET SBTC DETAILS
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTC_13"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If dbfRecordset.RecordCount > 0 Then
    LoopCount = 0
    Do Until LoopCount = dbfRecordset.RecordCount
        
        'FIELDS
        CT = UCase(dbfRecordset.Fields(0))
        Rtno = dbfRecordset.Fields(1)
        Acctno = dbfRecordset.Fields(2)

        If Len(dbfRecordset.Fields(3)) >= 1 Then
            Acctnm1 = UCase(dbfRecordset.Fields(3))
        Else
            Acctnm1 = ""
        End If
        
        If Len(dbfRecordset.Fields(4)) >= 1 Then
            Acctnm2 = UCase(dbfRecordset.Fields(4))
        Else
            Acctnm2 = ""
        End If
        
        Orderqty = Val(dbfRecordset.Fields(5))
        PrimaryKey = getNewPrimaryKey
    
    'GET BRANCH NAME CHECKDAT.MDB
        Dim Conn As ADODB.Connection
        Dim Rs As ADODB.Recordset
        
        Set Conn = New ADODB.Connection
        
        With Conn
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & FileLocation & "; Jet OLEDB;"
          .Open
        End With
        
        strQuery = "SELECT * FROM BRANCH WHERE [Routing Number] = '" & Rtno & "'"
        
        Set Rs = New ADODB.Recordset
        
        With Rs
        
        Set .ActiveConnection = Conn
            .CursorType = adOpenStatic
            .Source = strQuery
            .Open
        End With
        
        If Len(Rs.Fields(2)) >= 1 Then
            BranchName = Replace(Rs.Fields(2), "'", "''")
        Else
            BranchName = ""
        End If
        
        If Len(Rs.Fields(3)) >= 1 Then
            Address1 = Replace(Rs.Fields(3), "'", "''")
        Else
            Address1 = ""
        End If
        
        If Len(Rs.Fields(4)) >= 1 Then
            Address2 = Replace(Rs.Fields(4), "'", "''")
        Else
            Address2 = ""
        End If
        
        If Len(Rs.Fields(5)) >= 1 Then
            Address3 = Replace(Rs.Fields(5), "'", "''")
        Else
            Address3 = ""
        End If
    'END GET BRANCH NAME
        
        If CT = "A" Then
            ItemType = "PA"
            Formtype = "05"
            ChkType = "A"
            Batchno = UCase(BatchnoReg13)
        End If
        
        If CT = "B" Then
            ItemType = "CA"
            Formtype = "16"
            ChkType = "B"
            Batchno = UCase(BatchnoReg13)
        End If
        
        If CT = "B" And Mid(Acctno, 5, 8) = "21100608" Then 'MC Personal
            ItemType = "MC"
            Formtype = "05"
            ChkType = "A"
            Batchno = UCase(BatchnoMC13)
            Acctnm1 = ""
            Acctnm2 = ""
        End If
        
        If CT = "B" And Mid(Acctno, 5, 8) = "21200608" Then  'GC Commercial
            ItemType = "GC"
            Formtype = "05"
            ChkType = "A"
            Batchno = UCase(BatchnoGC13)
            Acctnm1 = ""
            Acctnm2 = ""
        End If
        
        'INSERT TO SBTC NEW
        Set DBFConnector1 = CreateObject("ADODB.Connection")
        
        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
        
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        SQL1 = "INSERT INTO SBTCNEW (Chktype,Rtno,Acctno,Acctnm1,Acctnm2,Orderqty,Formtype,Batchno,Branchname,Address1,Address2,Address3,ItemType,PrimaryKey,Digits) VALUES ('" _
                & ChkType & "','" & Rtno & "','" & Acctno & "','" & Replace(Acctnm1, "'", "''") & "','" & Replace(Acctnm2, "'", "''") & "','" & Orderqty & "','" & Formtype & "','" _
                & Batchno & "','" & BranchName & "','" & Address1 & "','" & Address2 & "','" & Address3 & "','" & ItemType & "','" & PrimaryKey & "','13')"
        
            DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
        'END INSERT TO SBTC NEW
    
        dbfRecordset.MoveNext
        LoopCount = LoopCount + 1
        
        Me.Caption = "Creating new SBTC..." & Acctno
    Loop
    
    If Dir$(App.Path & "\SBTC_13.dbf") <> "" Then Kill (App.Path & "\SBTC_13.dbf")
End If
End Sub

Sub EditMeNow()

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "UPDATE SBTCNEW SET Acctnm1 = '" & Replace(txtAcctnm1.Text, "'", "''") & "', Acctnm2 = '" _
        & Replace(txtAcctnm2.Text, "'", "''") & "', Status = '>> MANUAL EDIT' WHERE PrimaryKey = " _
        & txtPrimaryKey.Text & ""
dbfRecordset.Open SQL, DBFConnector, 1, 1

DisplayData

MsgBox "Edited data has been saved.", vbInformation, "Confirm Save"
End Sub

Sub DisplayData()

'TOTAL ACCOUNTS
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTCNEW"
dbfRecordset.Open SQL, DBFConnector, 1, 1

lblTotalAccounts.Caption = Val(dbfRecordset.RecordCount)
DBFConnector.Close
'TOTAL ACCOUNTS

'TOTAL BOOKS
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(val(ORDERQTY)) FROM SBTCNEW"
dbfRecordset.Open SQL, DBFConnector, 1, 1

lblTotalBooks.Caption = Val(dbfRecordset.Fields(0))
DBFConnector.Close
'TOTAL Books

'======================================= FOR 10 DIGITS !!! ======================================

'TOTAL BOOKS PA REG 10
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(val(ORDERQTY)) FROM SBTCNEW WHERE CHKTYPE = 'A' AND ITEMTYPE ='PA' AND DIGITS = '10'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    lblBooksPA10.Caption = Val(dbfRecordset.Fields(0))
Else
    lblBooksPA10.Caption = "0"
End If

DBFConnector.Close
'TOTAL BOOKS PA REG

'TOTAL BOOKS CA REG 10
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(ORDERQTY) FROM SBTCNEW WHERE CHKTYPE = 'B' AND ITEMTYPE ='CA' AND DIGITS = '10'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    lblBooksCA10.Caption = dbfRecordset.Fields(0)
Else
    lblBooksCA10.Caption = "0"
End If

DBFConnector.Close
'TOTAL BOOKS CA REG

'======================================= FOR 13 DIGITS !!! ======================================

'TOTAL BOOKS PA REG 10
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(val(ORDERQTY)) FROM SBTCNEW WHERE CHKTYPE = 'A' AND ITEMTYPE ='PA' AND DIGITS = '13'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    lblBooksPA13.Caption = Val(dbfRecordset.Fields(0))
Else
    lblBooksPA13.Caption = "0"
End If

DBFConnector.Close
'TOTAL BOOKS PA REG

'TOTAL BOOKS CA REG 10
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(ORDERQTY) FROM SBTCNEW WHERE CHKTYPE = 'B' AND ITEMTYPE ='CA' AND DIGITS = '13'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    lblBooksCA13.Caption = dbfRecordset.Fields(0)
Else
    lblBooksCA13.Caption = "0"
End If

DBFConnector.Close
'TOTAL BOOKS CA REG

'TOTAL BOOKS MC 13
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(val(ORDERQTY)) FROM SBTCNEW WHERE CHKTYPE = 'A' AND ITEMTYPE = 'MC' AND DIGITS = '13'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    lblMCBooks13.Caption = Val(dbfRecordset.Fields(0))
Else
    lblMCBooks13.Caption = "0"
End If

DBFConnector.Close
'TOTAL BOOKS MC

'TOTAL BOOKS GC 13
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(val(ORDERQTY)) FROM SBTCNEW WHERE CHKTYPE = 'A' AND ITEMTYPE = 'GC' AND DIGITS = '13'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    lblGCBooks13.Caption = Val(dbfRecordset.Fields(0))
Else
    lblGCBooks13.Caption = "0"
End If

DBFConnector.Close
'TOTAL BOOKS MC
'==============================================================================================

'set for display grdDisplay
Set DBFConnectorDisplay = CreateObject("ADODB.Connection")
 
DBFConnectorDisplay.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnectorDisplay.CursorLocation = adUseClient
    
Set DBFRecordsetDisplay = CreateObject("ADODB.Recordset")
SQLDisplay = "SELECT * FROM SBTCNEW"
DBFRecordsetDisplay.Open SQLDisplay, DBFConnectorDisplay, 1, 1

Set grdDisplay.DataSource = DBFRecordsetDisplay
'end set for display grdDisplay

'text boxes for data
txtRtno.DataField = "Rtno"
Set txtRtno.DataSource = DBFRecordsetDisplay

txtAcctno.DataField = "Acctno"
Set txtAcctno.DataSource = DBFRecordsetDisplay

txtAcctnm1.DataField = "Acctnm1"
Set txtAcctnm1.DataSource = DBFRecordsetDisplay

txtAcctnm2.DataField = "Acctnm2"
Set txtAcctnm2.DataSource = DBFRecordsetDisplay

txtChktype.DataField = "Chktype"
Set txtChktype.DataSource = DBFRecordsetDisplay

txtOrderqty.DataField = "Orderqty"
Set txtOrderqty.DataSource = DBFRecordsetDisplay

txtPrimaryKey.DataField = "PrimaryKey"
Set txtPrimaryKey.DataSource = DBFRecordsetDisplay

txtBatchno.DataField = "Batchno"
Set txtBatchno.DataSource = DBFRecordsetDisplay

txtItemType.DataField = "Itemtype"
Set txtItemType.DataSource = DBFRecordsetDisplay
'end text boxes for data

txtAcctnm1.Enabled = False
txtAcctnm2.Enabled = False
txtAcctno.Enabled = False
txtRtno.Enabled = False
txtOrderqty.Enabled = False
txtChktype.Enabled = False

txtAcctnm1.BackColor = &HC0FFFF
txtAcctnm2.BackColor = &HC0FFFF
txtAcctno.BackColor = &HC0FFFF
txtRtno.BackColor = &HC0FFFF
txtOrderqty.BackColor = &HC0FFFF
txtChktype.BackColor = &HC0FFFF

cmdEdit.Enabled = True
cmdGenerate.Enabled = False
cmdGenerate.Default = True
cmdDelete.Enabled = True

cmdSave.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub cmdBrowse10_Click()

dlgBrowse.InitDir = App.Path
dlgBrowse.Filter = "ASC File |*.asc"
dlgBrowse.ShowOpen

If dlgBrowse.FileName = "" Then Exit Sub

FilePath = dlgBrowse.FileName

InputFile = Left(dlgBrowse.FileTitle, Val(Len(dlgBrowse.FileTitle) - 4))

If Mid(InputFile, 1, 3) <> "YSE" Then
    MsgBox "Invalid file! It should be 13 digits!", vbOKOnly + vbInformation, "Error"
    Exit Sub
End If

'Check if File has already been Processed
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With


Set Rs = New ADODB.Recordset
SQL = "SELECT [Date] FROM Archive_10 WHERE FileName = '" & InputFile & "'"
Rs.Open SQL, Conn, adOpenStatic

If Rs.RecordCount >= 1 Then
    MsgBox InputFile & " has already been processed last " & Rs.Fields(0).Value, vbCritical, "Error"
    Exit Sub
End If
'End Check if File has already been Processed

If MsgBox("Are you sure you want to process file " & dlgBrowse.FileTitle & "?", vbYesNo + vbInformation, "Confirm File") = vbNo Then Exit Sub

lblFileName.Caption = InputFile

Result = Open_SBTC_TextFile10(FilePath, App.Path, InputFile)

VerifyRouting10

If VerifyPass = False Then End

MsgBox "Starter finished!", vbOKOnly, "Thank you!"

img10.Visible = True
cmdBrowse10.Enabled = False
cmdProcess.Enabled = True
End Sub

Private Sub cmdBrowse13_Click()

dlgBrowse.InitDir = App.Path
dlgBrowse.Filter = "ASC File |*.asc"
dlgBrowse.ShowOpen

If dlgBrowse.FileName = "" Then Exit Sub

FilePath = dlgBrowse.FileName

InputFile = Left(dlgBrowse.FileTitle, Val(Len(dlgBrowse.FileTitle) - 4))

If Mid(InputFile, 1, 2) <> "13" Then
    MsgBox "Invalid file! It should be 13 digits!", vbOKOnly + vbInformation, "Error"
    Exit Sub
End If

'Check if File has already been Processed
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With


Set Rs = New ADODB.Recordset
SQL = "SELECT [Date] FROM Archive WHERE FileName = '" & InputFile & "'"
Rs.Open SQL, Conn, adOpenStatic

If Rs.RecordCount >= 1 Then
    DateLast = Rs.Fields(0)
    
    DateDifference = DateDiff("D", DateLast, Now)
    
    If DateDifference <= 200 Then
        MsgBox InputFile & " has already been processed last " & Rs.Fields(0).Value, vbCritical, "Error"
        Exit Sub
    End If
    
End If
'End Check if File has already been Processed

If MsgBox("Are you sure you want to process file " & dlgBrowse.FileTitle & "?", vbYesNo + vbInformation, "Confirm File") = vbNo Then Exit Sub

lblFileName.Caption = InputFile

Result = Open_SBTC_TextFile13(FilePath, App.Path, InputFile)

VerifyRouting13

If VerifyPass = False Then End

CheckData

MsgBox "13 Digits finished!", vbOKOnly, "Thank you!"

img13.Visible = True
cmdBrowse13.Enabled = False
cmdProcess.Enabled = True
End Sub

Private Sub cmdCancel_Click()
DisplayData
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure you want to delete the Account Number " & txtAcctno.Text & "?", vbInformation + vbYesNo, "Confirm Delete") = vbNo Then Exit Sub

'DELETE FROM FILE
Set conn1 = New ADODB.Connection

conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
conn1.CursorLocation = adUseClient

Set Rs1 = CreateObject("ADODB.Recordset")
SQL1 = "DELETE * From SBTCNEW where PrimaryKey = " & txtPrimaryKey.Text
Rs1.Open SQL1, conn1, 1, 1

On Error GoTo errDelete:
'PACK THE TABLE
Set conn2 = New ADODB.Connection
conn2.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd2 = New ADODB.Command
cmd2.CommandType = adCmdText
Set cmd2.ActiveConnection = conn1
cmd2.CommandText = "Set Exclusive On"
cmd2.Execute
cmd2.CommandText = "Pack SBTCNEW.dbf"
cmd2.Execute
conn2.Close

errDelete:
DisplayData

MsgBox "Account Number " & txtAcctno.Text & " deleted", vbInformation, ""
End Sub

Private Sub cmdEdit_Click()
If MsgBox("Are you sure you want to edit this data?", vbInformation + vbYesNo, "Confirm Edit") = vbNo Then Exit Sub

EditYou = True

txtAcctnm1.DataField = ""
txtAcctnm1.BackColor = &HFFFFFF
txtAcctnm1.Enabled = True
txtAcctnm1.SetFocus

txtAcctnm2.DataField = ""
txtAcctnm2.BackColor = &HFFFFFF
txtAcctnm2.Enabled = True

cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
cmdSave.Default = True
cmdGenerate.Enabled = False
End Sub

Private Sub cmdGenerate_Click()

'For Zipping Files
Zip_Reg_13 = False
Zip_Reg_10 = False
Zip_MC = False
Zip_GC = False

ProcessBy = InputBox("Enter Process By", "", "")
If ProcessBy = "" Then Exit Sub


Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

'For Batches
Batches = ""
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(BatchNo) FROM SBTCNEW"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Batchno = dbfRecordset.Fields(0)
    
    If Batches = "" Then
        Batches = Batchno
    Else
        Batches = Batches & "_" & Batchno
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
'End For Batches

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(ItemType), Digits FROM SBTCNEW"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    ItemType = dbfRecordset.Fields(0)
    Digits = dbfRecordset.Fields(1)
    
    If (ItemType = "PA" Or ItemType = "CA") And Digits = "10" Then Zip_Reg_10 = True
    If (ItemType = "PA" Or ItemType = "CA") And Digits = "13" Then Zip_Reg_13 = True
    If ItemType = "MC" And Digits = "13" Then Zip_MC = True
    If ItemType = "GC" And Digits = "13" Then Zip_GC = True
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
'For Zipping Files


'Check Whether Ref.dbf is Modified --> REGULAR!
If CheckIfRefIsModifiedReg = False Then
    cmdBrowse10.Enabled = False
    cmdBrowse13.Enabled = False
    cmdProcess.Enabled = False
    
    MsgBox "Unable to Process REGULAR! ! ! Ref.dbf has been Modified. Please Check Ref.dbf", vbCritical, "Error"
    Exit Sub
End If
'-----------------------------------------------------------------

'Check Whether Ref.dbf is Modified --> MC!
If CheckIfRefIsModifiedMC = False Then
    cmdBrowse10.Enabled = False
    cmdBrowse13.Enabled = False
    cmdProcess.Enabled = False
    
    MsgBox "Unable to Process MC! ! ! Ref.dbf has been Modified. Please Check Ref.dbf", vbCritical, "Error"
    Exit Sub
End If
'-----------------------------------------------------------------

'Check Whether Ref.dbf is Modified --> GC!
If CheckIfRefIsModifiedGIFTCHK = False Then
    cmdBrowse10.Enabled = False
    cmdBrowse13.Enabled = False
    cmdProcess.Enabled = False
    
    MsgBox "Unable to Process GC! ! ! Ref.dbf has been Modified. Please Check Ref.dbf", vbCritical, "Error"
    Exit Sub
End If

'============================================ BACK UP! =======================================

If MsgBox("Are you sure you want to generate files?", vbYesNo, "System Information") = vbNo Then Exit Sub

'***Back-Up Ref PERSONAL COMMERCIAL
If lblBooksPA13.Caption <> "0" Or lblBooksCA13.Caption <> "0" Or lblBooksPA10.Caption <> "0" Or lblBooksCA10.Caption <> "0" Then
    NewRefFileName = "Ref_Bef_REG_" & BatchnoReg & "_" & Format(Now, "MMDDYY") & "_" & Format(Now, "HHMMAMPM") & ".dbf"
    
    If Dir$(App.Path & "\Ref_Before\" & NewRefFileName) <> "" Then Kill App.Path & "\Ref_Before\" & NewRefFileName
    FileCopy App.Path & "\Ref.dbf", App.Path & "\Ref_Before\" & NewRefFileName
End If

'***Back-Up Ref MC
If lblMCBooks13.Caption <> "0" Then
    NewRefFileName = "Ref_Bef_MC_" & BatchnoMC & "_" & Format(Now, "MMDDYY") & "_" & Format(Now, "HHMMAMPM") & ".dbf"
    
    If Dir$(App.Path & "\Ref_Before\" & NewRefFileName) <> "" Then Kill App.Path & "\Ref_Before\" & NewRefFileName
    FileCopy App.Path & "\MC\Ref.dbf", App.Path & "\Ref_Before\" & NewRefFileName
End If

'***Back-Up Ref GC
If lblGCBooks13.Caption <> "0" Then
    NewRefFileName = "Ref_Bef_GC_" & BatchnoGC & "_" & Format(Now, "MMDDYY") & "_" & Format(Now, "HHMMAMPM") & ".dbf"
    
    If Dir$(App.Path & "\Ref_Before\" & NewRefFileName) <> "" Then Kill App.Path & "\Ref_Before\" & NewRefFileName
    FileCopy App.Path & "\GIFTCHK\Ref.dbf", App.Path & "\Ref_Before\" & NewRefFileName
End If

'====================================== SORT RT MUNA LAHAT! ====================================

If lblBooksPA13.Caption <> "0" Or lblBooksCA13.Caption <> "0" Then  '13 DIGITS!
    SortRT13
    
    If MsgBox("SortRT generated for 13 digits. Do you wish to continue?", vbYesNo, "Confirm Generate") = vbNo Then
        cmdGenerate.Enabled = True
        Me.Enabled = True
        Exit Sub
    End If
End If

If lblBooksPA10.Caption <> "0" Or lblBooksCA10.Caption <> "0" Then  '10 DIGITS
    SortRT10
    
    If MsgBox("SortRT generated for 10 digits. Do you wish to continue?", vbYesNo, "Confirm Generate") = vbNo Then
        cmdGenerate.Enabled = True
        Me.Enabled = True
        Exit Sub
    End If
End If

If lblMCBooks13.Caption <> "0" Then 'MANAGERS CHECK
    SortRT_MC
    
    If MsgBox("SortRT generated for MC. Do you wish to continue?", vbYesNo, "Confirm Generate") = vbNo Then
        cmdGenerate.Enabled = True
        Me.Enabled = True
        Exit Sub
    End If
End If

If lblGCBooks13.Caption <> "0" Then 'GIFT CHECKS
    SortRT_GC
    
    If MsgBox("SortRT generated for GC. Do you wish to continue?", vbYesNo, "Confirm Generate") = vbNo Then
        cmdGenerate.Enabled = True
        Me.Enabled = True
        Exit Sub
    End If
End If
'===================================== 13 DIGITS! OTHER FILES ==================================

If lblBooksPA13.Caption <> "0" Or lblBooksCA13.Caption <> "0" Then
    BlockP13
    BlockC13
    
    PackingDBF13
    PackingTxt13
    
    TransP13
    TransC13
    
    PrinterFilePA13
    PrinterFileCA13
    
    PrinterTXT_PA13
    PrinterTXT_CA13
    
    FixedPositionPersonal13
    FixedPositionCommercial13

    If lblBooksPA13.Caption <> "0" Then
        UpdateMastertoREF_PA13
    End If
    
    If lblBooksCA13.Caption <> "0" Then
        UpdateMastertoREF_CA13
    End If
End If

'===================================== 10 DIGITS! OTHER FILES ==================================

If lblBooksPA10.Caption <> "0" Or lblBooksCA10.Caption <> "0" Then
    BlockP10
    BlockC10
    
    PackingDBF10
    PackingTxt10
    
    TransP10
    TransC10
    
    PrinterFilePA10
    PrinterFileCA10
    
    PrinterTXT_PA10
    PrinterTXT_CA10
    
    FixedPositionPersonal10
    FixedPositionCommercial10

    If lblBooksPA10.Caption <> "0" Then
        UpdateMastertoREF_PA10
    End If
    
    If lblBooksCA10.Caption <> "0" Then
        UpdateMastertoREF_CA10
    End If
End If

'======================================= LIMIT PROGRAM!!!! ==================================

If lblBooksPA13.Caption <> "0" Or lblBooksCA13.Caption <> "0" Or lblBooksPA10.Caption <> "0" Or lblBooksCA10.Caption <> "0" Then
    LimitProgramReg
End If

'===================================== MC! OTHER FILES ==================================
If lblMCBooks13.Caption <> "0" Then
    BlockP_MC
    PackPDBF_MC
    PackingDBF_MC
    PackingATxt_MC
    TransP_MC
    PrinterFilePA_MC
    PrinterTXT_MC
    FixedPositionPersonal_MC
    UpdateMastertoREF_MC
    LimitProgramMC
End If

'===================================== GC DIGITS! OTHER FILES ==================================
If lblGCBooks13.Caption <> "0" Then
    BlockP_GC
    PackPDBF_GC
    PackingDBF_GC
    PackingATxt_GC
    TransP_GC
    PrinterFilePA_GC
    PrinterTXT_GC
    FixedPositionPersonal_GC
    UpdateMastertoREF_GC
    LimitProgramGC
End If

'================================== BACK UP OF REF AND MASTER! =================================

'UPDATE REF AND MASTER REGULAR
If Dir$("\\192.168.0.29\captive\Banks\sbtc\REF_Auto.dbf") <> "" Then Kill ("\\192.168.0.29\captive\Banks\sbtc\REF_Auto.dbf")
FileCopy App.Path & "\REF.dbf", "\\192.168.0.29\captive\Banks\sbtc\REF_Auto.dbf"

If Dir$("\\192.168.0.29\captive\Banks\sbtc\MSTRAuto.dbf") <> "" Then Kill ("\\192.168.0.29\captive\Banks\sbtc\MSTRAuto.dbf")
FileCopy App.Path & "\MASTER.dbf", "\\192.168.0.29\captive\Banks\sbtc\MSTRAuto.dbf"

'UPDATE REF AND MASTER MC
If Dir$("\\192.168.0.29\captive\Banks\sbtc\MC\REF_Auto.dbf") <> "" Then Kill ("\\192.168.0.29\captive\Banks\sbtc\MC\REF_Auto.dbf")
FileCopy App.Path & "\MC\REF.dbf", "\\192.168.0.29\captive\Banks\sbtc\MC\REF_Auto.dbf"

If Dir$("\\192.168.0.29\captive\Banks\sbtc\MC\MSTRAuto.dbf") <> "" Then Kill ("\\192.168.0.29\captive\Banks\sbtc\MC\MSTRAuto.dbf")
FileCopy App.Path & "\MC\MASTER.dbf", "\\192.168.0.29\captive\Banks\sbtc\MC\MSTRAuto.dbf"

'UPDATE REF AND MASTER GIFTCHECK
If Dir$("\\192.168.0.29\captive\Banks\sbtc\GIFTCHK\REF_Auto.dbf") <> "" Then Kill ("\\192.168.0.29\captive\Banks\sbtc\GIFTCHK\REF_Auto.dbf")
FileCopy App.Path & "\GIFTCHK\REF.dbf", "\\192.168.0.29\captive\Banks\sbtc\GIFTCHK\REF_Auto.dbf"

If Dir$("\\192.168.0.29\captive\Banks\sbtc\GIFTCHK\MSTRAuto.dbf") <> "" Then Kill ("\\192.168.0.29\captive\Banks\sbtc\GIFTCHK\MSTRAuto.dbf")
FileCopy App.Path & "\GIFTCHK\MASTER.dbf", "\\192.168.0.29\captive\Banks\sbtc\GIFTCHK\MSTRAuto.dbf"



'Zip everything
If Dir("C:\Windows\Temp\" & DateTimeToday, vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\"

FileCopy App.Path & "\SBTC_Regular.exe", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\SBTC_Regular.exe"

If Zip_Reg_10 = True Then CopyAll ("10_Digits")
If Zip_Reg_13 = True Then CopyAll ("12_Digits")
If Zip_GC = True Then CopyAll ("GIFTCHK")
If Zip_MC = True Then CopyAll ("MC")

ProgamExecute = """" & WinZipLocation & """" & " -u -r -p " & App.Path & "\AFT" & "_" & Batches & "_" & ProcessBy & "_" & ".zip" & " C:\Windows\Temp\" & DateTimeToday & "\*.*"
Result = Shell(ProgamExecute, vbHide)

ProgamExecute = """" & WinZipLocation & """" & " -u -r -p -sKathlynRose " & App.Path & "\Old_Files\AFT" & "_" & Batches & "_" & ProcessBy & "_" & DateTimeToday & ".zip" & " C:\Windows\Temp\" & DateTimeToday & "\*.*"
Result = Shell(ProgamExecute, vbHide)
'End Zip everything

MsgBox "Data has been successfully generated.", vbInformation, ""

cmdEdit.Enabled = False
cmdSave.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = False
cmdGenerate.Enabled = False
End
End Sub

Sub CopyAll(FolderLocation)

Dim fso As New FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")


If Dir("C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation, vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation

FileCopy App.Path & "\" & FolderLocation & "\BlockP.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\BlockP.txt"
FileCopy App.Path & "\" & FolderLocation & "\PackingA.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PackingA.txt"
FileCopy App.Path & "\" & FolderLocation & "\TransP.dbf", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\TransP.dbf"


If FolderLocation = "10_Digits" Or FolderLocation = "12_Digits" Then
    FileCopy App.Path & "\" & FolderLocation & "\FixedPositionP." & Format(Now, "YY") & "F", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\FixedPositionP." & Format(Now, "YY") & "F"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFileP." & Format(Now, "YY") & "P", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFileP." & Format(Now, "YY") & "P"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFileP.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFileP.txt"
End If

If FolderLocation = "MC" Then
    FileCopy App.Path & "\" & FolderLocation & "\FixedPositionMC." & Format(Now, "YY") & "F", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\FixedPositionMC." & Format(Now, "YY") & "F"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFileMC." & Format(Now, "YY") & "P", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFileMC." & Format(Now, "YY") & "P"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFilePA.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFilePA.txt"

    fso.CopyFile App.Path & "\MC\MC.captive", "C:\Windows\Temp\" & DateTimeToday & "\MC\MC.captive", True
End If

If FolderLocation = "GIFTCHK" Then
    FileCopy App.Path & "\" & FolderLocation & "\FixedPositionGC." & Format(Now, "YY") & "F", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\FixedPositionGC." & Format(Now, "YY") & "F"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFileGC." & Format(Now, "YY") & "P", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFileGC." & Format(Now, "YY") & "P"
    'FileCopy App.Path & "\" & FolderLocation & "\PrinterFileGC.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFileGC.txt"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFilePA.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFilePA.txt"
    
    fso.CopyFile App.Path & "\GIFTCHK\GIFTCHK.captive", "C:\Windows\Temp\" & DateTimeToday & "\GIFTCHK\GIFTCHK.captive", True
End If

FileCopy App.Path & "\" & FolderLocation & "\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\SortRT.txt"
FileCopy App.Path & "\" & FolderLocation & "\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\Packing.dbf"
FileCopy App.Path & "\" & FolderLocation & "\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\Packing.dbf"



If FolderLocation = "10_Digits" Or FolderLocation = "12_Digits" Then
    FileCopy App.Path & "\" & FolderLocation & "\BlockC.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\BlockC.txt"
    FileCopy App.Path & "\" & FolderLocation & "\FixedPositionC." & Format(Now, "YY") & "F", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\FixedPositionC." & Format(Now, "YY") & "F"
    FileCopy App.Path & "\" & FolderLocation & "\PackingB.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PackingB.txt"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFileC." & Format(Now, "YY") & "P", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFileC." & Format(Now, "YY") & "P"
    FileCopy App.Path & "\" & FolderLocation & "\PrinterFileC.txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\PrinterFileC.txt"

    FileCopy App.Path & "\" & FolderLocation & "\TransC.dbf", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\TransC.dbf"
End If


If FolderLocation = "10_Digits" Or FolderLocation = "12_Digits" Then
    If Dir("C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\Ref.dbf") <> "" Then Kill "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\Ref.dbf"
    FileCopy App.Path & "\Ref.dbf", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderLocation & "\Ref.dbf"
    
    fso.CopyFile App.Path & "\SBTC.captive", "C:\Windows\Temp\" & DateTimeToday & "\SBTC.captive", True
End If
End Sub

Private Sub cmdLength_Click()
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTCNEW"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    
    Acctnm1 = dbfRecordset.Fields(3)
    Acctnm2 = dbfRecordset.Fields(4)
    
    If Len(Acctnm1) > 40 Then
        MsgBox "Account Name1 " & Acctnm1 & " is too long. Please edit your data.", vbInformation, "Error"
        cmdGenerate.Enabled = False
        Exit Sub
    Else
        cmdGenerate.Enabled = True
    End If
    
    If Len(Acctnm2) > 40 Then
        MsgBox "Account Name2 " & Acctnm2 & " is too long. Please edit your data.", vbInformation, "Error"
        cmdGenerate.Enabled = False
        Exit Sub
    Else
        cmdGenerate.Enabled = True
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
dbfRecordset.Close
End Sub

Function CheckDoubleAccount(ChkType)
If ChkType = "A" Then ChequeName = "Personal"
If ChkType = "B" Then ChequeName = "Commercial"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Acctno) FROM SBTCNEW WHERE Chktype = '" & ChkType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    AccountNo = dbfRecordset.Fields(0)
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM SBTCNEW WHERE Acctno = '" & AccountNo & "' AND Chktype = '" & ChkType & "'"
    DBFRecordset1.Open SQL, DBFConnector, 1, 1
    
    If DBFRecordset1.RecordCount > 1 Then
        DoubleAccount_CheckType(DoubleAccount_Total) = ChequeName
        DoubleAccount_Accounts(DoubleAccount_Total) = AccountNo
        DoubleAccount_Total = DoubleAccount_Total + 1
    End If
    
    Me.Caption = "Checking for Double Account: " & AccountNo
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
End Function

Private Sub cmdProcess_Click()
    
    If (cmdBrowse10.Enabled = False) Then
        InsertNew_10
    End If
    
    If (cmdBrowse13.Enabled = False) Then
        InsertNew_13
    End If
    
    'For Double Account
    DoubleAccount_Total = 0
    CheckDoubleAccount ("A")
    CheckDoubleAccount ("B")
    
    If DoubleAccount_Total >= 1 Then
        Close #1
        Open App.Path & "\DoubleAccounts.txt" For Output As #1
        
        Print #1, "Double Accounts as of " & Format(Now, "Mmm. DD, YYYY")
        Print #1, ""
        
        LoopCount = 0
        Do Until LoopCount = DoubleAccount_Total
            Print #1, DoubleAccount_Accounts(LoopCount) & "  " & DoubleAccount_CheckType(LoopCount)
            
            LoopCount = LoopCount + 1
        Loop
        
        Close #1
        
        If MsgBox("There are " & DoubleAccount_Total & " double Accounts." & vbNewLine & "Do you want to Continue?" & vbNewLine & "See DoubleAccounts.txt", vbYesNo + vbInformation, "Confirm Continue") = vbNo Then Exit Sub
        
    End If
    'End For Double Account
    
    
    CheckUpdatedCheckDat

    'Generate Text Error
    Close #1
    Open App.Path & "\Errors.txt" For Output As #1
    Print #1, lstErrors.ListCount & " Error/s Found on " & Format(Now, "Mmm. DD, YYYY") & "  " & Format(Now, "HH:MM AMPM")
    Print #1, ""
    
    If lstErrors.ListCount <> 0 Then
        LoopCount = 0
        Do Until LoopCount = lstErrors.ListCount
            Print #1, lstErrors.List(LoopCount)
            
            LoopCount = LoopCount + 1
        Loop
    
        MsgBox "See Errors.txt for the details. " & vbNewLine & " The program will end and you must edit the data.", vbInformation + vbOKOnly, "Thank you"
        End
    End If
    
    Close #1
    'End Generate Text Error

    DisplayData

    Me.Caption = "SBTC"
    
    cmdProcess.Enabled = False
    cmdBrowse10.Enabled = False
    cmdBrowse13.Enabled = False
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdDelete.Enabled = True
    cmdCancel.Enabled = False
    cmdGenerate.Enabled = False
    cmdLength.Enabled = True

Me.Caption = "SBTC"
MsgBox "DBF Output has been generated", vbOKOnly, ""
End Sub

Private Sub cmdSave_Click()
If MsgBox("Are you sure you want to save this record?", vbInformation + vbYesNo, "Confirm Save") = vbNo Then Exit Sub

'ALL CAPS
txtAcctnm1.Text = UCase(txtAcctnm1.Text)
txtAcctnm2.Text = UCase(txtAcctnm2.Text)

'ROUTING NUMBER
If txtRtno.Text = "123456789" Or txtRtno.Text = "000000000" Then
 MsgBox "Routing number not valid.", vbInformation, "ERROR"
 txtRtno.Text = ""
 txtRtno.SetFocus
 Exit Sub
End If

If Len(txtRtno.Text) < 9 Then
 MsgBox "Routing number must be 9 characters.", vbInformation, "ERROR"
 txtRtno.Text = ""
 txtRtno.SetFocus
 Exit Sub
End If

If txtRtno.Text = "" Then
 MsgBox "Routing number must not be empty.", vbInformation, "ERROR"
 txtRtno.Text = ""
 txtRtno.SetFocus
 Exit Sub
End If

'ACCOUNT NUMBER
If txtAcctno.Text = "" Then
 MsgBox "Account number must not be empty.", vbInformation, "ERROR"
 txtAcctno.Text = ""
 txtAcctno.SetFocus
 Exit Sub
End If

'BATCHNO
If txtBatchno.Text = "" Then
    MsgBox "Batch No. should not be empty.", vbInformation, "Error"
    txtBatchno.SetFocus
    Exit Sub
End If

txtBatchno.Text = UCase(txtBatchno.Text)

'CHEQUE TYPE
txtChktype.Text = UCase(txtChktype.Text)

If (txtChktype.Text <> "A") And (txtChktype.Text <> "B") Then
 MsgBox "Type must only be to A or B.", vbInformation, "ERROR"
 txtChktype.Text = ""
 txtChktype.SetFocus
 Exit Sub
End If

If txtChktype.Text = "A" Then Formtype = "05"
If txtChktype.Text = "B" Then Formtype = "16"

'ORDER QTY
If txtOrderqty.Text = "" Then
    MsgBox "Order quantity must not be empty.", vbInformation, "ERROR"
    txtOrderqty.SetFocus
    Exit Sub
End If

If EditYou = True Then
    EditMeNow
Else

'GET BRANCH NAME CHECKDAT.MDB
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    
    Set Conn = New ADODB.Connection
    
    With Conn
      .CursorLocation = adUseClient
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & FileLocation & "; Jet OLEDB;"
      .Open
    End With
    
    strQuery = "SELECT * FROM BRANCH WHERE [Routing Number] = '" & txtRtno.Text & "'"
    
    Set Rs = New ADODB.Recordset
    
    With Rs
    
    Set .ActiveConnection = Conn
        .CursorType = adOpenStatic
        .Source = strQuery
        .Open
    End With
    
    If Rs.RecordCount < 1 Then
        MsgBox "Routing number " & txtRtno.Text & " does not exists on CHECKDAT.MDB, please check your data.", vbInformation, "Error"
        txtRtno.Text = ""
        txtRtno.SetFocus
        Exit Sub
    End If
    
    If Len(Rs.Fields(2)) >= 1 Then
        BranchName = Replace(Rs.Fields(2), "'", "''")
    Else
        BranchName = ""
    End If
    
    If Len(Rs.Fields(3)) >= 1 Then
        Address1 = Replace(Rs.Fields(3), "'", "''")
    Else
        Address1 = ""
    End If
    
    If Len(Rs.Fields(4)) >= 1 Then
        Address2 = Replace(Rs.Fields(4), "'", "''")
    Else
        Address2 = ""
    End If
    
    If Len(Rs.Fields(5)) >= 1 Then
        Address3 = Replace(Rs.Fields(5), "'", "''")
    Else
        Address3 = ""
    End If
'END GET BRANCH NAME

'CHECK IF ROUTING EXISTS REF.DBF
Set DBFConnector1 = CreateObject("ADODB.Connection")
DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"

DBFConnector1.CursorLocation = adUseClient
    
Set DBFRecordset1 = CreateObject("ADODB.Recordset")
SQL1 = "SELECT * FROM REF WHERE Rtno = '" & txtRtno.Text & "'"
DBFRecordset1.Open SQL1, DBFConnector1, 1, 1

If DBFRecordset1.RecordCount < 1 Then
    MsgBox "Routing Number " & txtRtno.Text & " does not exists on REF.dbf, please check your data.", vbInformation, "Error"
    txtRtno.Text = ""
    txtRtno.SetFocus
    Exit Sub
End If
'END CHECK IF ROUTING EXISTS REF.DBF

'SAVE DATA
Set DBFOutput = New ADODB.Connection
DBFOutput.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"

InsertData = "INSERT INTO SBTCNEW (Chktype,Rtno,Acctno,Acctnm1,Acctnm2,Orderqty,Formtype,Batchno,Branchname,Address1,Address2,Address3,Status,ItemType,PrimaryKey) VALUES ('" _
             & txtChktype.Text & "','" & txtRtno.Text & "','" & txtAcctno.Text & "','" & Replace(txtAcctnm1.Text, "'", "''") & "','" _
             & Replace(txtAcctnm2.Text, "'", "''") & "','" & txtOrderqty.Text & "','" & Formtype & "','" & txtBatchno.Text & "','" & BranchName & "','" _
             & Address1 & "','" & Address2 & "','" & Address3 & "','>> MANUAL ADD','" & ItemType & "', ')"

DBFOutput.Execute InsertData
DBFOutput.Close
'END SAVE DATA

DisplayData

MsgBox "New data has been saved.", vbInformation, ""
End If
End Sub
Sub ReadWinZip()
Close #1
Open "C:\WinZip.txt" For Input As #1

Do Until EOF(1)
    Line Input #1, WinZipLocation
Loop

Close #1
End Sub

Private Sub Form_Load()

If App.PrevInstance Then End

ReadWinZip

DateTimeToday = Format(Now, "YYYYMMDDHHMMSS")

img10.Visible = False
img13.Visible = False

If cmdBrowse10.Enabled = True Or cmdBrowse13.Enabled = True Then
    cmdProcess.Enabled = False
Else
    cmdProcess.Enabled = True
End If

CheckAll

'Check if the Program has already run today --> REGULAR!
If ProgramAlreadyRunReg <> "" Or ProgramAlreadyRunMC <> "" Or ProgramAlreadyRunGIFTCHK <> "" Then
    If MsgBox("Unable to Process ! ! ! This program can only process order only ONCE PER DAY ! " & vbNewLine & vbNewLine & "Are you sure you want to Process Today?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then End
    
    MyName = InputBox("Enter your Name", "", "")
    If MyName = "" Then End
    
    MyReason = InputBox("Enter Reason why you want to Re-process file", "", "")
    If MyReason = "" Then End
    
    Result = GenerateEmail(MyName, MyReason)
    
    ProcessOrdersToday
    
    
End If

'Check Whether Ref.dbf is Modified --> REGULAR!
If CheckIfRefIsModifiedReg = False Then
    cmdBrowse10.Enabled = False
    cmdBrowse13.Enabled = False
    cmdProcess.Enabled = False

    MsgBox "Unable to Process REGULAR! ! ! Ref.dbf has been Modified. Please Check Ref.dbf", vbCritical, "Error"
    Exit Sub
End If
'---------------------------------------------------------------

'Check if the Program has already run today --> MC!
If ProgramAlreadyRunMC <> "" Then
    cmdBrowse10.Enabled = False
    cmdBrowse13.Enabled = False
    cmdProcess.Enabled = False
    
    MsgBox "Unable to Process MC! ! ! This program can only process order only ONCE PER DAY ! " & vbNewLine & "Last Batch Processed: " & ProgramAlreadyRunMC, vbCritical, "Error"
    Exit Sub
End If


'---------------------------------------------------------------


'Check Whether Ref.dbf is Modified --> GC!
If CheckIfRefIsModifiedGIFTCHK = False Then
    cmdBrowse10.Enabled = False
    cmdBrowse13.Enabled = False
    cmdProcess.Enabled = False

    MsgBox "Unable to Process GC! ! ! Ref.dbf has been Modified. Please Check Ref.dbf", vbCritical, "Error"
    Exit Sub
End If
'---------------------------------------------------------------

lblFileName.Caption = ""

Open App.Path & "\Datasource\FileLocation.txt" For Input As #1
    Line Input #1, FileLocation
Close #1

Me.Caption = "SBTC"
cmdBrowse10.Default = True

DisableAll

'SBTC 10
If Dir$(App.Path & "\SBTCNEW.dbf") <> "" Then Kill (App.Path & "\SBTCNEW.dbf")
FileCopy App.Path & "\Datasource\SBTCNEW.dbf", App.Path & "\SBTCNEW.dbf"



'Check if Branches exists
CheckRef_if_exists ("MC")
CheckRef_if_exists ("GIFTCHK")
'End Check if Branches exists
End Sub

Sub CheckRef_if_exists(FolderName)
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set DBFConnector1 = CreateObject("ADODB.Connection")

DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\" & FolderName & ";Extended properties=dBase III"
DBFConnector1.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(RTNO), Branch_Tex FROM REF"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT Branch_Tex FROM REF WHERE RTNo = '" & BRSTN & "'"
    DBFRecordset1.Open SQL, DBFConnector1, 1, 1
    
    If DBFRecordset1.RecordCount = 0 Then
        MsgBox "BRSTN " & BRSTN & " with Branch Name " & BranchName & " does not exists on Ref.dbf on " & FolderName, vbCritical, "Error"
        End
    End If
        
    If DBFRecordset1.RecordCount > 1 Then
        MsgBox "BRSTN " & BRSTN & " with Branch Name " & BranchName & " contains more than 1 Branch on Ref.dbf on " & FolderName, vbCritical, "Error"
        End
    End If
        
    If DBFRecordset1.RecordCount = 1 Then
        BranchName_Folder = DBFRecordset1.Fields(0)
        
        If BranchName <> BranchName_Folder Then
            MsgBox "BRSTN " & BRSTN & " does not match on Ref.dbf" & vbNewLine & vbNewLine & "Regular: " & BranchName & vbNewLine & FolderName & ": " & BranchName_Folder, vbCritical, "Error"
            End
        End If
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
End Sub

Sub CheckAll()

'CHECKDAT
If Dir$("\\192.168.0.29\CAPTIVE\AUTO\SBTC\CHECKDAT.MDB", vbDirectory) = "" Then
    MsgBox "CHECKDAT.MDB does not exists.", vbCritical, "Error"
    End
End If

'SOURCE LOCATION
If Dir$(App.Path & "\SourceLocation.ini", vbDirectory) = "" Then
    MsgBox "SourceLocation.ini does not exists.", vbCritical, "Error"
    End
End If

'REGULAR
If Dir$(App.Path & "\SBTC.captive", vbDirectory) = "" Then
    MsgBox "SBTC.captive does not exists.", vbCritical, "Error"
    End
End If
    
'MC
If Dir$(App.Path & "\MC\MC.captive", vbDirectory) = "" Then
    MsgBox "MC.captive on MC Folder does not exists.", vbCritical, "Error"
    End
End If

'GIFTCHK
If Dir$(App.Path & "\GIFTCHK\GIFTCHK.captive", vbDirectory) = "" Then
    MsgBox "GIFTCHK.captive on GIFTCHK Folder does not exists.", vbCritical, "Error"
    End
End If

'FOLDERS:
    If Dir$(App.Path & "\Datasource", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource folder does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\Ref_Before", vbDirectory) = "" Then
        MsgBox App.Path & "\Ref_Before folder does not exists.", vbCritical, "Error"
        End
    End If

'DATABASE:
    If Dir$(App.Path & "\Datasource\FileLocation.txt", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource\FileLocation.txt does not exists.", vbCritical, "Error"
        End
    End If

'SBTC NEW
    If Dir$(App.Path & "\Datasource\SBTCNEW.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource\SBTCNEW.dbf does not exists.", vbCritical, "Error"
        End
    End If
    
'DBF:
    If Dir$(App.Path & "\Ref.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Ref.dbf does not exists.", vbCritical, "Error"
        End
    End If
    
    If Dir$(App.Path & "\Master.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Master.dbf does not exists.", vbCritical, "Error"
        End
    End If

'OUTSIDE DBF:

    If Dir$(App.Path & "\Datasource\PackP.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource\PackP.dbf does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\Datasource\PackC.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource\PackC.dbf does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\Datasource\Packing.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource\Packing.dbf does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\Datasource\TransP.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource\TransP.dbf does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\Datasource\TransC.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\Datasource\TransC.dbf does not exists.", vbCritical, "Error"
        End
    End If
    
'GG AND MC
    If Dir$(App.Path & "\MC", vbDirectory) = "" Then
        MsgBox App.Path & "\MC Folder does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\MC\Ref.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\MC\Ref.dbf does not exists.", vbCritical, "Error"
        End
    End If
    
    If Dir$(App.Path & "\MC\Master.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\MC\Master.dbf does not exists.", vbCritical, "Error"
        End
    End If
    
    If Dir$(App.Path & "\GIFTCHK", vbDirectory) = "" Then
        MsgBox App.Path & "\GIFTCHK Folder does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\GIFTCHK\Ref.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\GC\Ref.dbf does not exists.", vbCritical, "Error"
        End
    End If

    If Dir$(App.Path & "\GIFTCHK\Master.dbf", vbDirectory) = "" Then
        MsgBox App.Path & "\GIFTCHK\Master.dbf does not exists.", vbCritical, "Error"
        End
    End If
    
End Sub

Sub DisableAll()
'TOTAL
lblTotalAccounts.Caption = ""
lblTotalBooks.Caption = ""

'10 DIGITS
lblBooksPA10.Caption = ""
lblBooksCA10.Caption = ""

'13 DIGITS
lblBooksPA13.Caption = ""
lblBooksCA13.Caption = ""
lblMCBooks13.Caption = ""
lblGCBooks13.Caption = ""

txtAcctnm1.Enabled = False
txtAcctnm2.Enabled = False
txtRtno.Enabled = False
txtChktype.Enabled = False
txtOrderqty.Enabled = False
txtAcctno.Enabled = False
End Sub

Private Sub grdDisplay_KeyPress(KeyAscii As Integer) '--> REGULAR!
'For Initialization
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & GetSourceLocation & "; Jet OLEDB:Database Password=Elgae;"
  .Open
End With
'End For Initialization

'Get the Ascii Codes
Set Rs = New ADODB.Recordset
SQL = "SELECT Code FROM AccessCode"
Rs.Open SQL, Conn, adOpenStatic

Password = Rs.Fields(0)

KeyAsciiPassword = ""
LoopCount = 0
Do Until LoopCount = Len(Password)
    KeyAsciiPassword = KeyAsciiPassword & Asc(Mid(Password, LoopCount + 1, 1))
    
    LoopCount = LoopCount + 1
Loop
'End Get the Ascii Codes

'Check the Password
DataCount = 0
LoopCount = 0
Do Until LoopCount = Len(Password)
    If Chr(KeyAscii) = Mid(Password, LoopCount + 1, 1) Then DataCount = DataCount + 1
    
    LoopCount = LoopCount + 1
Loop
If DataCount = 0 Then
    KeyAsciiFinal = ""
Else
    KeyAsciiFinal = KeyAsciiFinal & KeyAscii
End If

If KeyAsciiFinal = KeyAsciiPassword Then
    frmModifiedReg.Show
    Me.Enabled = False
End If
'End Check the Password
End Sub

Private Sub lstErrors_KeyPress(KeyAscii As Integer) '--> MC!
'For Initialization
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & GetSourceLocation & "; Jet OLEDB:Database Password=Elgae;"
  .Open
End With
'End For Initialization

'Get the Ascii Codes
Set Rs = New ADODB.Recordset
SQL = "SELECT Code FROM AccessCode"
Rs.Open SQL, Conn, adOpenStatic

Password = Rs.Fields(0)

KeyAsciiPassword = ""
LoopCount = 0
Do Until LoopCount = Len(Password)
    KeyAsciiPassword = KeyAsciiPassword & Asc(Mid(Password, LoopCount + 1, 1))
    
    LoopCount = LoopCount + 1
Loop
'End Get the Ascii Codes

'Check the Password
DataCount = 0
LoopCount = 0
Do Until LoopCount = Len(Password)
    If Chr(KeyAscii) = Mid(Password, LoopCount + 1, 1) Then DataCount = DataCount + 1
    
    LoopCount = LoopCount + 1
Loop
If DataCount = 0 Then
    KeyAsciiFinal = ""
Else
    KeyAsciiFinal = KeyAsciiFinal & KeyAscii
End If

If KeyAsciiFinal = KeyAsciiPassword Then
    frmModifiedMC.Show
    Me.Enabled = False
End If
'End Check the Password
End Sub

Private Sub lstGCCLick_KeyPress(KeyAscii As Integer) '--> GC!
'For Initialization
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & GetSourceLocation & "; Jet OLEDB:Database Password=Elgae;"
  .Open
End With
'End For Initialization

'Get the Ascii Codes
Set Rs = New ADODB.Recordset
SQL = "SELECT Code FROM AccessCode"
Rs.Open SQL, Conn, adOpenStatic

Password = Rs.Fields(0)

KeyAsciiPassword = ""
LoopCount = 0
Do Until LoopCount = Len(Password)
    KeyAsciiPassword = KeyAsciiPassword & Asc(Mid(Password, LoopCount + 1, 1))
    
    LoopCount = LoopCount + 1
Loop
'End Get the Ascii Codes

'Check the Password
DataCount = 0
LoopCount = 0
Do Until LoopCount = Len(Password)
    If Chr(KeyAscii) = Mid(Password, LoopCount + 1, 1) Then DataCount = DataCount + 1
    
    LoopCount = LoopCount + 1
Loop
If DataCount = 0 Then
    KeyAsciiFinal = ""
Else
    KeyAsciiFinal = KeyAsciiFinal & KeyAscii
End If

If KeyAsciiFinal = KeyAsciiPassword Then
    frmModifiedGC.Show
    Me.Enabled = False
End If
'End Check the Password
End Sub

Private Sub txtAcctno_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case 8
    Case Else
    KeyAscii = 0
    End Select
End Sub

Private Sub txtOrderqty_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case 8
    Case Else
    KeyAscii = 0
    End Select
End Sub

Private Sub txtRtno_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case 8
    Case Else
    KeyAscii = 0
    End Select
End Sub

Sub PageDivider(ItemType)

If Val(PageSeparator) >= 35 Or Val(PageSeparator) = 0 Then
    If Val(PageSeparator) <> 0 Then Print #1, ""
    Print #1, ""
    Print #1, "    Page No.     " & Val(PageNo)
    Print #1, "    " & Format(Now, "MM/DD/YYYY")
    
    If ItemType = "MC" Then
        Print #1, "                        SBC M.C. - Summary of RT nos / # of Books"
    Else
        Print #1, "                          SBC - Summary of RT nos / # of Books"
    End If
    
    Print #1, ""
    Print #1, "    ACCTNO       QTY BRANCH                 ACCOUNT NAME"
    Print #1, ""
    Print #1, ""
    
    PageSeparator = 1
    PageNo = Val(PageNo + 1)
End If

PageSeparator = Val(PageSeparator) + 1

End Sub

Sub SortRT13()

PageSeparator = 0
PageNo = 1
GrandTotal = 0

Close #1
Open App.Path & "\12_Digits\SORTRT.txt" For Output As #1

ChkType = "A"
ItemType = "PA"
Digits = "13"

SortRTLoop:
'***Read the contents of DBF File
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Rtno), Branchname FROM SBTCNEW WHERE Chktype = '" _
        & ChkType & "' AND ItemType = '" & ItemType & "' AND Digits = '" & Digits & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'***End Read the contents of DBF File

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    
    'BATCHNO
    Set DBFConnectorB = CreateObject("ADODB.Connection")
    
    DBFConnectorB.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnectorB.CursorLocation = adUseClient
    
    Set DBFRecordsetB = CreateObject("ADODB.Recordset")
    SQLB = "SELECT Batchno FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = '" _
        & ChkType & "' AND ItemType = '" & ItemType & "' AND Digits = '" & Digits & "'"
    DBFRecordsetB.Open SQLB, DBFConnectorB, 1, 1
    
    Batchno = DBFRecordsetB.Fields(0)
    
    PageDivider (ItemType)
    Print #1, "   ** CHECK TYPE/BRSTN/BATCH # ---->  " & ChkType & "/" & Rtno & "/" & Batchno
    Print #1, "   ** BRANCH NAME ----> " & BranchName
    Print #1, ""
    
    '***Read the contents of DBF File
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient

    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = '" _
            & ChkType & "' AND ItemType = '" & ItemType & "' AND Digits = '" _
            & Digits & "' ORDER By Rtno, Acctno"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    '***End Read the contents of DBF File

    LoopCount2 = 0
    TotalOrderqty = 0
    
    Do Until LoopCount2 = dbfRecordset2.RecordCount
        Acctno = dbfRecordset2.Fields(2)
        Orderqty = Val(dbfRecordset2.Fields(5))
        TotalOrderqty = Val(Orderqty) + Val(TotalOrderqty)
        
        If Len(dbfRecordset2.Fields(12)) >= 1 Then
            ManualEdit = dbfRecordset2.Fields(12)
        Else
            ManualEdit = ""
        End If

        Do Until Len(Orderqty) >= 4
            Orderqty = " " & Orderqty
        Loop

        If Len(dbfRecordset2.Fields(3)) >= 1 Then
            Acctnm1 = dbfRecordset2.Fields(3)
        Else
            Acctnm1 = ""
        End If
        
        If Len(dbfRecordset2.Fields(4)) >= 1 Then
            Acctnm2 = dbfRecordset2.Fields(4)
        Else
            Acctnm2 = ""
        End If
        
        Print #1, "    " & Acctno & " " & Orderqty & "   " & Acctnm1 & "  " & ManualEdit
        
        PageDivider (ItemType)
        If Acctnm2 <> "" Then Print #1, "                        " & Acctnm2
            
        GrandTotal = Val(GrandTotal) + Val(Orderqty)
        
        dbfRecordset2.MoveNext
        LoopCount2 = LoopCount2 + 1
        
        Me.Caption = "SortRT 13 Digits.txt..." & Acctno
    Loop
    
    Do Until Len(TotalOrderqty) >= 3
        TotalOrderqty = " " & TotalOrderqty
    Loop
    
    PageDivider (ItemType)
    Print #1, "   ** Subtotal ** "
    Print #1, "                  " & TotalOrderqty
    
    PageDivider (ItemType)
    Print #1, ""
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
Loop

If ChkType = "A" Then
    ChkType = "B"
    ItemType = "CA"
    Digits = "13"
    GoTo SortRTLoop
    Exit Sub
Else
    Do Until Len(GrandTotal) >= 3
        GrandTotal = " " & GrandTotal
    Loop
    
    Print #1, "   *** Total ***  "
    Print #1, "                  " & GrandTotal
    Close #1
End If
End Sub

Sub SortRT10()

PageSeparator = 0
PageNo = 1
GrandTotal = 0

Close #1
Open App.Path & "\10_Digits\SORTRT.txt" For Output As #1

ChkType = "A"
ItemType = "PA"
Digits = "10"

SortRTLoop:
'***Read the contents of DBF File
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Rtno), Branchname FROM SBTCNEW WHERE Chktype = '" _
        & ChkType & "' AND ItemType = '" & ItemType & "' AND Digits = '" & Digits & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'***End Read the contents of DBF File

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    
    'BATCHNO
    Set DBFConnectorB = CreateObject("ADODB.Connection")
    
    DBFConnectorB.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnectorB.CursorLocation = adUseClient
    
    Set DBFRecordsetB = CreateObject("ADODB.Recordset")
    SQLB = "SELECT Batchno FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = '" _
        & ChkType & "' AND ItemType = '" & ItemType & "' AND Digits = '" & Digits & "'"
    DBFRecordsetB.Open SQLB, DBFConnectorB, 1, 1
    
    Batchno = DBFRecordsetB.Fields(0)
    
    PageDivider (ItemType)
    Print #1, "   ** CHECK TYPE/BRSTN/BATCH # ---->  " & ChkType & "/" & Rtno & "/" & Batchno
    Print #1, "   ** BRANCH NAME ----> " & BranchName
    Print #1, ""
    
    '***Read the contents of DBF File
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient

    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = '" _
            & ChkType & "' AND ItemType = '" & ItemType & "' AND Digits = '" _
            & Digits & "' ORDER By Rtno, Acctno"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    '***End Read the contents of DBF File

    LoopCount2 = 0
    TotalOrderqty = 0
    
    Do Until LoopCount2 = dbfRecordset2.RecordCount
        Acctno = dbfRecordset2.Fields(2)
        Orderqty = Val(dbfRecordset2.Fields(5))
        TotalOrderqty = Val(Orderqty) + Val(TotalOrderqty)
        
        If Len(dbfRecordset2.Fields(12)) >= 1 Then
            ManualEdit = dbfRecordset2.Fields(12)
        Else
            ManualEdit = ""
        End If

        Do Until Len(Orderqty) >= 4
            Orderqty = " " & Orderqty
        Loop

        If Len(dbfRecordset2.Fields(3)) >= 1 Then
            Acctnm1 = dbfRecordset2.Fields(3)
        Else
            Acctnm1 = ""
        End If
        
        If Len(dbfRecordset2.Fields(4)) >= 1 Then
            Acctnm2 = dbfRecordset2.Fields(4)
        Else
            Acctnm2 = ""
        End If
        
        Print #1, "    " & Acctno & " " & Orderqty & "   " & Acctnm1 & "  " & ManualEdit
        
        PageDivider (ItemType)
        If Acctnm2 <> "" Then Print #1, "                        " & Acctnm2
            
        GrandTotal = Val(GrandTotal) + Val(Orderqty)
        
        dbfRecordset2.MoveNext
        LoopCount2 = LoopCount2 + 1
        
        Me.Caption = "SortRT 10 Digits.txt..." & Acctno
    Loop
    
    Do Until Len(TotalOrderqty) >= 3
        TotalOrderqty = " " & TotalOrderqty
    Loop
    
    PageDivider (ItemType)
    Print #1, "   ** Subtotal ** "
    Print #1, "                  " & TotalOrderqty
    
    PageDivider (ItemType)
    Print #1, ""
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
Loop

If ChkType = "A" Then
    ChkType = "B"
    ItemType = "CA"
    Digits = "10"
    GoTo SortRTLoop
    Exit Sub
Else
    Do Until Len(GrandTotal) >= 3
        GrandTotal = " " & GrandTotal
    Loop
    
    Print #1, "   *** Total ***  "
    Print #1, "                  " & GrandTotal
    Close #1
End If
End Sub

Sub SortRT_GC()
PageSeparator = 0
PageNo = 1
GrandTotal = 0

Open App.Path & "\GIFTCHK\SORTRT.txt" For Output As #1

ChkType = "A"
ItemType = "GC"

SortRTLoop:
'***Read the contents of DBF File
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Rtno),Branchname FROM SBTCNEW WHERE Chktype = '" & ChkType & "' AND ITEMTYPE = '" & ItemType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'***End Read the contents of DBF File

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    
    'BATCHNO
    Set DBFConnectorB = CreateObject("ADODB.Connection")
    
    DBFConnectorB.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnectorB.CursorLocation = adUseClient
    
    Set DBFRecordsetB = CreateObject("ADODB.Recordset")
    SQLB = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Itemtype = 'GC'"
    DBFRecordsetB.Open SQLB, DBFConnectorB, 1, 1
    
    Batchno = DBFRecordsetB.Fields(7)
    
    PageDivider (ItemType)
    Print #1, "   ** CHECK TYPE/BRSTN/BATCH # ---->  " & ChkType & "/" & Rtno & "/" & Batchno
    Print #1, "   ** BRANCH NAME ----> " & BranchName
    Print #1, ""
    
    '***Read the contents of DBF File
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient

    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = '" & ChkType & "' AND Itemtype = '" & ItemType & "' ORDER By Rtno, Acctno"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    '***End Read the contents of DBF File

    LoopCount2 = 0
    TotalOrderqty = 0
    
    Do Until LoopCount2 = dbfRecordset2.RecordCount
        Acctno = dbfRecordset2.Fields(2)
        Orderqty = Val(dbfRecordset2.Fields(5))
        TotalOrderqty = Val(Orderqty) + Val(TotalOrderqty)
        
        If Len(dbfRecordset2.Fields(12)) >= 1 Then
            ManualEdit = dbfRecordset2.Fields(12)
        Else
            ManualEdit = ""
        End If

        Do Until Len(Orderqty) = 4
            Orderqty = " " & Orderqty
        Loop

        If Len(dbfRecordset2.Fields(3)) >= 1 Then
            Acctnm1 = dbfRecordset2.Fields(3)
        Else
            Acctnm1 = ""
        End If
        
        If Len(dbfRecordset2.Fields(4)) >= 1 Then
            Acctnm2 = dbfRecordset2.Fields(4)
        Else
            Acctnm2 = ""
        End If
        
        Print #1, "    " & Acctno & " " & Orderqty & "   " & Acctnm1 & "  " & ManualEdit
        
        PageDivider (ItemType)
        If Acctnm2 <> "" Then Print #1, "                        " & Acctnm2
            
        GrandTotal = Val(GrandTotal) + Val(Orderqty)
        
        dbfRecordset2.MoveNext
        LoopCount2 = LoopCount2 + 1
        
        Me.Caption = "SortRT_GC.txt..." & Acctno
    Loop
    
    Do Until Len(TotalOrderqty) = 3
        TotalOrderqty = " " & TotalOrderqty
    Loop
    
    PageDivider (ItemType)
    Print #1, "   ** Subtotal ** "
    Print #1, "                  " & TotalOrderqty
    
    PageDivider (ItemType)
    Print #1, ""
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
Loop

    Do Until Len(GrandTotal) = 3
        GrandTotal = " " & GrandTotal
    Loop
    
    Print #1, "   *** Total ***  "
    Print #1, "                  " & GrandTotal
    Close #1
End Sub

Sub SortRT_MC()
PageSeparator = 0
PageNo = 1
GrandTotal = 0

Open App.Path & "\MC\SORTRT.txt" For Output As #1

ChkType = "A"
ItemType = "MC"

SortRTLoop:
'***Read the contents of DBF File
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Rtno),Branchname FROM SBTCNEW WHERE Chktype = '" & ChkType & "' AND ITEMTYPE = '" & ItemType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

'***End Read the contents of DBF File

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    
    'BATCHNO
    Set DBFConnectorB = CreateObject("ADODB.Connection")
    
    DBFConnectorB.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnectorB.CursorLocation = adUseClient
    
    Set DBFRecordsetB = CreateObject("ADODB.Recordset")
    SQLB = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Itemtype = 'MC'"
    DBFRecordsetB.Open SQLB, DBFConnectorB, 1, 1
    
    Batchno = DBFRecordsetB.Fields(7)
    
    PageDivider (ItemType)
    Print #1, "   ** CHECK TYPE/BRSTN/BATCH # ---->  " & ChkType & "/" & Rtno & "/" & Batchno
    Print #1, "   ** BRANCH NAME ----> " & BranchName
    Print #1, ""
    
    '***Read the contents of DBF File
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient

    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = '" & ChkType & "' AND Itemtype = '" & ItemType & "' ORDER By Rtno, Acctno"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    '***End Read the contents of DBF File

    LoopCount2 = 0
    TotalOrderqty = 0
    
    Do Until LoopCount2 = dbfRecordset2.RecordCount
        Acctno = dbfRecordset2.Fields(2)
        Orderqty = Val(dbfRecordset2.Fields(5))
        TotalOrderqty = Val(Orderqty) + Val(TotalOrderqty)
        
        If Len(dbfRecordset2.Fields(12)) >= 1 Then
            ManualEdit = dbfRecordset2.Fields(12)
        Else
            ManualEdit = ""
        End If

        Do Until Len(Orderqty) = 4
            Orderqty = " " & Orderqty
        Loop

        If Len(dbfRecordset2.Fields(3)) >= 1 Then
            Acctnm1 = dbfRecordset2.Fields(3)
        Else
            Acctnm1 = ""
        End If
        
        If Len(dbfRecordset2.Fields(4)) >= 1 Then
            Acctnm2 = dbfRecordset2.Fields(4)
        Else
            Acctnm2 = ""
        End If
        
        Print #1, "    " & Acctno & " " & Orderqty & "   " & Acctnm1 & "  " & ManualEdit
        
        PageDivider (ItemType)
        If Acctnm2 <> "" Then Print #1, "                        " & Acctnm2
            
        GrandTotal = Val(GrandTotal) + Val(Orderqty)
        
        dbfRecordset2.MoveNext
        LoopCount2 = LoopCount2 + 1
        
        Me.Caption = "SortRT_MC.txt..." & Acctno
    Loop
    
    Do Until Len(TotalOrderqty) = 3
        TotalOrderqty = " " & TotalOrderqty
    Loop
    
    PageDivider (ItemType)
    Print #1, "   ** Subtotal ** "
    Print #1, "                  " & TotalOrderqty
    
    PageDivider (ItemType)
    Print #1, ""
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
Loop

    Do Until Len(GrandTotal) = 3
        GrandTotal = " " & GrandTotal
    Loop
    
    Print #1, "   *** Total ***  "
    Print #1, "                  " & GrandTotal
    Close #1
End Sub

Sub BlockP13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

If Dir$(App.Path & "\12_Digits\PACKP.dbf") <> "" Then Kill (App.Path & "\12_Digits\PACKP.dbf")
FileCopy App.Path & "\DataSource\PACKP.dbf", App.Path & "\12_Digits\PACKP.dbf"

Close #1
Open App.Path & "\12_Digits\BLOCKP.txt" For Output As #1

CheckType = "A"

'GET RTNO
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = '" _
        & CheckType & "' AND Itemtype = 'PA' AND Digits = '13'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'END GET RTNO

BlockCountDisplay = 1
BlockCount = 0
LoopCount = 0
Counting = 0
PageDisplay = 2

Print #1, ""
Print #1, "        Page No.     1"
Print #1, "        " & Format(Now, "MM/DD/YYYY")
Print #1, "                      SBTC - SUMMARY OF BLOCK - PERSONAL"
Print #1, ""
Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
Print #1,
Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

Do Until LoopCount = dbfRecordset.RecordCount
    RoutingNumber = dbfRecordset.Fields(0)
    
    'GET DETAILS
    Set DBFConnector5 = CreateObject("ADODB.Connection")
     
    DBFConnector5.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector5.CursorLocation = adUseClient
        
    Set DBFRecordset5 = CreateObject("ADODB.Recordset")
    SQL5 = "SELECT * FROM SBTCNEW WHERE Rtno = '" _
            & RoutingNumber & "' AND Chktype = 'A' AND Itemtype = 'PA' AND Digits = '13' ORDER BY Rtno,Acctno"
    DBFRecordset5.Open SQL5, DBFConnector5, 1, 1
    'END GET DETAILS
    
    LoopCountDr = 0
  
    Do Until LoopCountDr = DBFRecordset5.RecordCount
        Routingno = DBFRecordset5.Fields(1)
        AccountNo = DBFRecordset5.Fields(2)
        AcctNoWithHyphen = Mid(AccountNo, 1, 3) & "-" & Mid(AccountNo, 4, 6) & "-" & Mid(AccountNo, 10, 3)
        Books = DBFRecordset5.Fields(5)
        Batchno = DBFRecordset5.Fields(7)
        Status = DBFRecordset5.Fields(12)
        
        If Len(DBFRecordset5.Fields(3)) >= 1 Then
            Acctnm1 = DBFRecordset5.Fields(3)
        Else
            Acctnm1 = ""
        End If
        
        If Len(DBFRecordset5.Fields(4)) >= 1 Then
            Acctnm2 = DBFRecordset5.Fields(4)
        Else
            Acctnm2 = ""
        End If
            
        If LoopCountDr < 1 Then
    
        Set DBFConnector2 = CreateObject("ADODB.Connection")
 
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "SELECT * FROM REF WHERE Chktype = 'A' AND Rtno = '" & Routingno & "'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        LastSerial = dbfRecordset2.Fields(6)
        Serial = Val(LastSerial) + 1
        Serial = Format(Serial, "0000000")

        Do Until Val(Books) = 0
        
            EndingSerial = Val(Serial) + 49
            EndingSerial = Format(EndingSerial, "0000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & "    " & EndingSerial & " " & Status
        
        'INSERT TO PACKP
        Result = PackP13(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
       
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                      SBTC - SUMMARY OF BLOCK - PERSONAL"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."
            
            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
    
            Serial = Val(Serial) + 50
            Serial = Format(Serial, "0000000")
        
        Me.Caption = "BLOCKP 13 Digits... " & Serial
    Loop
 
 Else
        Do Until Val(Books) = 0
        Serial = (EndingSerial) + 1
        
        If Len(Serial) < 7 Then
            Serial = Format(Serial, "0000000")
        End If
        
            EndingSerial = Val(Serial) + 49
            EndingSerial = Format(EndingSerial, "0000000")

        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & "    " & EndingSerial & " " & Status
        
        'INSERT TO PACKP
        Result = PackP13(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
        
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                      SBTC - SUMMARY OF BLOCK - PERSONAL"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        
        Serial = Val(Serial) + 50
        Serial = Format(Serial, "0000000")
        
        Me.Caption = "BLOCKP 13 Digits... " & Serial
    Loop
End If
            DBFRecordset5.MoveNext
            LoopCountDr = LoopCountDr + 1
        Loop
        
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Close #1
End Sub

Sub BlockP10()

'13 DIGITS
PathOfFile = App.Path & "\10_Digits"

If Dir$(App.Path & "\10_Digits\PACKP.dbf") <> "" Then Kill (App.Path & "\10_Digits\PACKP.dbf")
FileCopy App.Path & "\DataSource\PACKP.dbf", App.Path & "\10_Digits\PACKP.dbf"

Close #1
Open App.Path & "\10_Digits\BLOCKP.txt" For Output As #1

CheckType = "A"

'GET RTNO
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = '" _
        & CheckType & "' AND Itemtype = 'PA' AND Digits = '10'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'END GET RTNO

BlockCountDisplay = 1
BlockCount = 0
LoopCount = 0
Counting = 0
PageDisplay = 2

Print #1, ""
Print #1, "        Page No.     1"
Print #1, "        " & Format(Now, "MM/DD/YYYY")
Print #1, "                       SBTC - SUMMARY OF BLOCK - PERSONAL"
Print #1, "                               STARTER 50 PIECES"
Print #1, ""
Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
Print #1,
Print #1, "        BLOCK RT_NO     M ACCT_NO       START_NO.  END_NO."

Do Until LoopCount = dbfRecordset.RecordCount
    RoutingNumber = dbfRecordset.Fields(0)
    
    
    'GET DETAILS
    Set DBFConnector5 = CreateObject("ADODB.Connection")
     
    DBFConnector5.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector5.CursorLocation = adUseClient
        
    Set DBFRecordset5 = CreateObject("ADODB.Recordset")
    SQL5 = "SELECT * FROM SBTCNEW WHERE Rtno = '" _
            & RoutingNumber & "' AND Chktype = 'A' AND Itemtype = 'PA' AND Digits = '10' ORDER BY Rtno,Acctno"
    DBFRecordset5.Open SQL5, DBFConnector5, 1, 1
    'END GET DETAILS
    
    
    LoopCountDr = 0
  
    Do Until LoopCountDr = DBFRecordset5.RecordCount
        Routingno = DBFRecordset5.Fields(1)
        AccountNo = DBFRecordset5.Fields(2)
        AcctNoWithHyphen = Mid(AccountNo, 1, 3) & "-" & Mid(AccountNo, 4, 6) & "-" & Mid(AccountNo, 10, 3)
        Books = DBFRecordset5.Fields(5)
        Batchno = DBFRecordset5.Fields(7)
        Status = DBFRecordset5.Fields(12)
        
        If Len(DBFRecordset5.Fields(3)) >= 1 Then
            Acctnm1 = DBFRecordset5.Fields(3)
        Else
            Acctnm1 = ""
        End If
        
        If Len(DBFRecordset5.Fields(4)) >= 1 Then
            Acctnm2 = DBFRecordset5.Fields(4)
        Else
            Acctnm2 = ""
        End If
        
        If LoopCountDr < 1 Then
    
        Set DBFConnector2 = CreateObject("ADODB.Connection")
 
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "SELECT * FROM REF WHERE Chktype = 'A' AND Rtno = '" & Routingno & "'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        LastSerial = dbfRecordset2.Fields(6)
        Serial = Val(LastSerial) + 1
        Serial = Format(Serial, "0000000")

        Do Until Val(Books) = 0
        
            EndingSerial = Val(Serial) + 49
            EndingSerial = Format(EndingSerial, "0000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & "    " & EndingSerial & " " & Status
        
        'INSERT TO PACKP
        Result = PackP10(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
       
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                       SBTC - SUMMARY OF BLOCK - PERSONAL"
            Print #1, "                               STARTER 50 PIECES"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO       START_NO.  END_NO."
            
            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
    
            Serial = Val(Serial) + 50
            Serial = Format(Serial, "0000000")
        
        Me.Caption = "BLOCKP 10 Digits... " & Serial
    Loop
 
 Else
        Do Until Val(Books) = 0
        Serial = (EndingSerial) + 1
        Serial = Format(Serial, "0000000")
        
            EndingSerial = Val(Serial) + 49
            EndingSerial = Format(EndingSerial, "0000000")

        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & "    " & EndingSerial & " " & Status
        
        'INSERT TO PACKP
        Result = PackP10(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
        
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                       SBTC - SUMMARY OF BLOCK - PERSONAL"
            Print #1, "                               STARTER 50 PIECES"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO       START_NO.  END_NO."

            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        
        Serial = Val(Serial) + 50
        Serial = Format(Serial, "0000000")
        
        Me.Caption = "BLOCKP 10 Digits... " & Serial
    Loop
End If
            DBFRecordset5.MoveNext
            LoopCountDr = LoopCountDr + 1
        Loop
        
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Close #1
End Sub

'13 DIGITS PA
Function PackP13(PathOfFile, Batchno, BlockCountDisplay2, Rtno, Acctno, AcctNoWithHyphen, ChkType, Acctnm1, Acctnm2, Orderqty, StartingA, StartingB, EndingA, EndingB)
Set DBFConnector4 = CreateObject("ADODB.Connection")

DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & PathOfFile & ";Extended properties=dBase III"
DBFConnector4.CursorLocation = adUseClient

Set DBFRecordset4 = CreateObject("ADODB.Recordset")

SQL4 = "INSERT INTO PACKP (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" _
        & Batchno & "','" & Val(BlockCountDisplay2) & "','" & Rtno & "','" & Acctno & "','" & AcctNoWithHyphen & "','A','" _
        & Replace(Acctnm1, "'", "''") & "','" & Replace(Acctnm2, "'", "''") & "','1','" & StartingA & "','" & StartingB & "','" _
        & EndingA & "','" & EndingB & "')"
DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
End Function

'13 DIGITS CA
Function PackC13(PathOfFile, Batchno, BlockCountDisplay2, Rtno, Acctno, AcctNoWithHyphen, ChkType, Acctnm1, Acctnm2, Orderqty, StartingA, StartingB, EndingA, EndingB)
Set DBFConnector4 = CreateObject("ADODB.Connection")

DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & PathOfFile & ";Extended properties=dBase III"
DBFConnector4.CursorLocation = adUseClient

Set DBFRecordset4 = CreateObject("ADODB.Recordset")

SQL4 = "INSERT INTO PACKC (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" _
        & Batchno & "','" & Val(BlockCountDisplay2) & "','" & Rtno & "','" & Acctno & "','" & AcctNoWithHyphen & "','B','" _
        & Replace(Acctnm1, "'", "''") & "','" & Replace(Acctnm2, "'", "''") & "','1','" & StartingA & "','" & StartingB & "','" _
        & EndingA & "','" & EndingB & "')"
DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
End Function

'10 DIGITS PA
Function PackP10(PathOfFile, Batchno, BlockCountDisplay2, Rtno, Acctno, AcctNoWithHyphen, ChkType, Acctnm1, Acctnm2, Orderqty, StartingA, StartingB, EndingA, EndingB)
Set DBFConnector4 = CreateObject("ADODB.Connection")

DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & PathOfFile & ";Extended properties=dBase III"
DBFConnector4.CursorLocation = adUseClient

Set DBFRecordset4 = CreateObject("ADODB.Recordset")

SQL4 = "INSERT INTO PACKP (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" _
        & Batchno & "','" & Val(BlockCountDisplay2) & "','" & Rtno & "','" & Acctno & "','" & AcctNoWithHyphen & "','A','" _
        & Replace(Acctnm1, "'", "''") & "','" & Replace(Acctnm2, "'", "''") & "','1','" & StartingA & "','" & StartingB & "','" _
        & EndingA & "','" & EndingB & "')"
DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
End Function

'10 DIGITS CA
Function PackC10(PathOfFile, Batchno, BlockCountDisplay2, Rtno, Acctno, AcctNoWithHyphen, ChkType, Acctnm1, Acctnm2, Orderqty, StartingA, StartingB, EndingA, EndingB)
Set DBFConnector4 = CreateObject("ADODB.Connection")

DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " & PathOfFile & ";Extended properties=dBase III"
DBFConnector4.CursorLocation = adUseClient

Set DBFRecordset4 = CreateObject("ADODB.Recordset")

SQL4 = "INSERT INTO PACKC (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" _
        & Batchno & "','" & Val(BlockCountDisplay2) & "','" & Rtno & "','" & Acctno & "','" & AcctNoWithHyphen & "','B','" _
        & Replace(Acctnm1, "'", "''") & "','" & Replace(Acctnm2, "'", "''") & "','1','" & StartingA & "','" & StartingB & "','" _
        & EndingA & "','" & EndingB & "')"
DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
End Function

Sub BlockP_MC()

CheckType = "A"

'GET RTNO
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'MC'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'END GET RTNO

BlockCountDisplay = 1
BlockCount = 0
LoopCount = 0
Counting = 0
PageDisplay = 2

Close #1
Open App.Path & "\MC\BLOCKP.txt" For Output As #1
            
Print #1, ""
Print #1, "        Page No.     1"
Print #1, "        " & Format(Now, "MM/DD/YYYY")
Print #1, "                         SBTC - SUMMARY OF BLOCK - M.C."
Print #1, ""
Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
Print #1,
Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

Do Until LoopCount = dbfRecordset.RecordCount
    RoutingNumber = dbfRecordset.Fields(0)
    
    'GET DETAILS
    Set DBFConnector5 = CreateObject("ADODB.Connection")
     
    DBFConnector5.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector5.CursorLocation = adUseClient
        
    Set DBFRecordset5 = CreateObject("ADODB.Recordset")
    SQL5 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & RoutingNumber & "' AND Chktype = 'A' AND Itemtype = 'MC' ORDER BY Rtno,Acctno"
    DBFRecordset5.Open SQL5, DBFConnector5, 1, 1
    'END GET DETAILS
    
    LoopCountDr = 0
  
    Do Until LoopCountDr = DBFRecordset5.RecordCount
        Routingno = DBFRecordset5.Fields(1)
        AccountNo = DBFRecordset5.Fields(2)
        Books = DBFRecordset5.Fields(5)
        Status = DBFRecordset5.Fields(12)
            
        If LoopCountDr < 1 Then
    
        Set DBFConnector2 = CreateObject("ADODB.Connection")
 
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC\;Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "SELECT * FROM REF WHERE ChkType = '" & CheckType & "' AND Rtno = '" & Routingno & "'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        LastSerial = dbfRecordset2.Fields(6)
        Serial = Val(LastSerial) + 1
        Serial = Format(Serial, "0000000000")

        Do Until Val(Books) = 0
            EndingSerial = Val(Serial) + 49
            EndingSerial = Format(EndingSerial, "0000000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & " " & EndingSerial & " " & Status
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                         SBTC - SUMMARY OF BLOCK - M.C."
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."
            
            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        Serial = Val(Serial) + 50
        Serial = Format(Serial, "0000000000")
        
        Me.Caption = "BLOCKP MC... " & Serial
    Loop
 
 Else
        Do Until Val(Books) = 0
        Serial = (EndingSerial) + 1
        Serial = Format(Serial, "0000000000")
             
        EndingSerial = Val(Serial) + 49
        EndingSerial = Format(EndingSerial, "0000000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & " " & EndingSerial & " " & Status
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                         SBTC - SUMMARY OF BLOCK - M.C."
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        Serial = Val(Serial) + 50
        Serial = Format(Serial, "0000000000")
        
        Me.Caption = "BLOCKP MC... " & Serial
    Loop
End If
            DBFRecordset5.MoveNext
            LoopCountDr = LoopCountDr + 1
        Loop
        
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Close #1
End Sub

Sub BlockP_GC()

CheckType = "A"

'GET RTNO
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'GC'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'END GET RTNO

BlockCountDisplay = 1
BlockCount = 0
LoopCount = 0
Counting = 0
PageDisplay = 2

Close #1
Open App.Path & "\GIFTCHK\BLOCKP.txt" For Output As #1
            
Print #1, ""
Print #1, "        Page No.     1"
Print #1, "        " & Format(Now, "MM/DD/YYYY")
Print #1, "                     SBTC - SUMMARY OF BLOCK - GIFTCHECK(GC)"
Print #1, ""
Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
Print #1,
Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

Do Until LoopCount = dbfRecordset.RecordCount
    RoutingNumber = dbfRecordset.Fields(0)
    
    'GET DETAILS
    Set DBFConnector5 = CreateObject("ADODB.Connection")
     
    DBFConnector5.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector5.CursorLocation = adUseClient
        
    Set DBFRecordset5 = CreateObject("ADODB.Recordset")
    SQL5 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & RoutingNumber & "' AND Chktype = 'A' AND Itemtype = 'GC' ORDER BY Rtno,Acctno"
    DBFRecordset5.Open SQL5, DBFConnector5, 1, 1
    'END GET DETAILS
    
    LoopCountDr = 0
  
    Do Until LoopCountDr = DBFRecordset5.RecordCount
        Routingno = DBFRecordset5.Fields(1)
        AccountNo = DBFRecordset5.Fields(2)
        Books = DBFRecordset5.Fields(5)
        Status = DBFRecordset5.Fields(12)
            
        If LoopCountDr < 1 Then
    
        Set DBFConnector2 = CreateObject("ADODB.Connection")
 
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK\;Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "SELECT * FROM REF WHERE ChkType = '" & CheckType & "' AND Rtno = '" & Routingno & "'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        LastSerial = dbfRecordset2.Fields(6)
        Serial = Val(LastSerial) + 1
        Serial = Format(Serial, "000000")

        Do Until Val(Books) = 0
            EndingSerial = Val(Serial) + 49
            EndingSerial = Format(EndingSerial, "000000")

        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & "    " & EndingSerial & " " & Status
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                     SBTC - SUMMARY OF BLOCK - GIFTCHECK(GC)"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."
            
            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        Serial = Val(Serial) + 50
        Serial = Format(Serial, "000000")
        
        Me.Caption = "BLOCKP GC... " & Serial
    Loop
 
 Else
        Do Until Val(Books) = 0
        Serial = (EndingSerial) + 1
        Serial = Format(Serial, "000000")
             
        EndingSerial = Val(Serial) + 49
        EndingSerial = Format(EndingSerial, "000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & "    " & EndingSerial & " " & Status
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                     SBTC - SUMMARY OF BLOCK - GIFTCHECK(GC)"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        Serial = Val(Serial) + 50
        Serial = Format(Serial, "000000")
        
        Me.Caption = "BLOCK GC... " & Serial
    Loop
End If
            DBFRecordset5.MoveNext
            LoopCountDr = LoopCountDr + 1
        Loop
        
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Close #1
End Sub

Sub BlockC13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

If Dir$(App.Path & "\12_Digits\PACKC.dbf") <> "" Then Kill (App.Path & "\12_Digits\PACKC.dbf")
FileCopy App.Path & "\DataSource\PACKC.dbf", App.Path & "\12_Digits\PACKC.dbf"

Close #1
Open App.Path & "\12_Digits\BLOCKC.txt" For Output As #1
    
CheckType = "B"

'GET RTNO
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = '" _
        & CheckType & "' AND Itemtype = 'CA' AND Digits = '13'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'END GET RTNO

BlockCountDisplay = 1
BlockCount = 0
LoopCount = 0
Counting = 0
PageDisplay = 2

Print #1, ""
Print #1, "        Page No.     1"
Print #1, "        " & Format(Now, "MM/DD/YYYY")
Print #1, "                     SBTC - SUMMARY OF BLOCK - COMMERCIAL"
Print #1, ""
Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
Print #1,
Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

Do Until LoopCount = dbfRecordset.RecordCount
    RoutingNumber = dbfRecordset.Fields(0)
    
    'GET DETAILS
    Set DBFConnector5 = CreateObject("ADODB.Connection")
     
    DBFConnector5.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector5.CursorLocation = adUseClient
        
    Set DBFRecordset5 = CreateObject("ADODB.Recordset")
    SQL5 = "SELECT * FROM SBTCNEW WHERE Rtno = '" _
            & RoutingNumber & "' AND Chktype = 'B' AND Itemtype = 'CA' AND Digits = '13' ORDER BY Rtno,Acctno"
    DBFRecordset5.Open SQL5, DBFConnector5, 1, 1
    'END GET DETAILS
    
    LoopCountDr = 0
  
    Do Until LoopCountDr = DBFRecordset5.RecordCount
        Routingno = DBFRecordset5.Fields(1)
        AccountNo = DBFRecordset5.Fields(2)
        AcctNoWithHyphen = Mid(AccountNo, 1, 3) & "-" & Mid(AccountNo, 4, 6) & "-" & Mid(AccountNo, 10, 3)
        Books = DBFRecordset5.Fields(5)
        Batchno = DBFRecordset5.Fields(7)
        
        If Len(DBFRecordset5.Fields(12)) >= 1 Then
            Status = DBFRecordset5.Fields(12)
        Else
            Status = ""
        End If

        If Len(DBFRecordset5.Fields(3)) >= 1 Then
            Acctnm1 = DBFRecordset5.Fields(3)
        Else
            Acctnm1 = ""
        End If

        If Len(DBFRecordset5.Fields(4)) >= 1 Then
            Acctnm2 = DBFRecordset5.Fields(4)
        Else
            Acctnm2 = ""
        End If
        
        If LoopCountDr < 1 Then
    
        Set DBFConnector2 = CreateObject("ADODB.Connection")
 
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "SELECT * FROM REF WHERE Chktype = '" & CheckType & "' AND Rtno = '" & Routingno & "'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        LastSerial = dbfRecordset2.Fields(6)
        Serial = Val(LastSerial) + 1
        Serial = Format(Serial, "0000000000")

        Do Until Val(Books) = 0
        
            EndingSerial = Val(Serial) + 99
            EndingSerial = Format(EndingSerial, "0000000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & " " & EndingSerial & " " & Status
        
        'INSERT TO PACKC
        Result = PackC13(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
        
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                     SBTC - SUMMARY OF BLOCK - COMMERCIAL"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."
            
            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        
        Serial = Val(Serial) + 100
        Serial = Format(Serial, "0000000000")
        
        Me.Caption = "BLOCKC 13 Digits... " & Serial
    Loop
 
 Else
        Do Until Val(Books) = 0
        Serial = (EndingSerial) + 1
        Serial = Format(Serial, "0000000000")
        
        EndingSerial = Val(Serial) + 99
        EndingSerial = Format(EndingSerial, "0000000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & " " & EndingSerial & " " & Status
        
        'INSERT TO PACKC
        Result = PackC13(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
       
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                     SBTC - SUMMARY OF BLOCK - COMMERCIAL"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."

            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        
        Serial = Val(Serial) + 100
        Serial = Format(Serial, "0000000000")
    
        Me.Caption = "BLOCKC 13 Digits... " & Serial
    Loop
End If
            DBFRecordset5.MoveNext
            LoopCountDr = LoopCountDr + 1
        Loop
        
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Close #1
End Sub

Sub BlockC10()

'13 DIGITS
PathOfFile = App.Path & "\10_Digits"

If Dir$(App.Path & "\10_Digits\PACKC.dbf") <> "" Then Kill (App.Path & "\10_Digits\PACKC.dbf")
FileCopy App.Path & "\DataSource\PACKC.dbf", App.Path & "\10_Digits\PACKC.dbf"

Close #1
Open App.Path & "\10_Digits\BLOCKC.txt" For Output As #1
    
CheckType = "B"

'GET RTNO
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = '" _
        & CheckType & "' AND Itemtype = 'CA' AND Digits = '10'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'END GET RTNO

BlockCountDisplay = 1
BlockCount = 0
LoopCount = 0
Counting = 0
PageDisplay = 2

Print #1, ""
Print #1, "        Page No.     1"
Print #1, "        " & Format(Now, "MM/DD/YYYY")
Print #1, "                     SBTC - SUMMARY OF BLOCK - COMMERCIAL"
Print #1, "                               STARTER 100 PIECES"
Print #1, ""
Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
Print #1,
Print #1, "        BLOCK RT_NO     M ACCT_NO       START_NO.  END_NO."

Do Until LoopCount = dbfRecordset.RecordCount
    RoutingNumber = dbfRecordset.Fields(0)
    
    
    'GET DETAILS
    Set DBFConnector5 = CreateObject("ADODB.Connection")
     
    DBFConnector5.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector5.CursorLocation = adUseClient
        
    Set DBFRecordset5 = CreateObject("ADODB.Recordset")
    SQL5 = "SELECT * FROM SBTCNEW WHERE Rtno = '" _
            & RoutingNumber & "' AND Chktype = 'B' AND Itemtype = 'CA' AND Digits = '10' ORDER BY Rtno,Acctno"
    DBFRecordset5.Open SQL5, DBFConnector5, 1, 1
    'END GET DETAILS
    
    
    LoopCountDr = 0
  
    Do Until LoopCountDr = DBFRecordset5.RecordCount
        Routingno = DBFRecordset5.Fields(1)
        AccountNo = DBFRecordset5.Fields(2)
        AcctNoWithHyphen = Mid(AccountNo, 1, 3) & "-" & Mid(AccountNo, 4, 6) & "-" & Mid(AccountNo, 10, 3)
        Books = DBFRecordset5.Fields(5)
        Batchno = DBFRecordset5.Fields(7)
        Status = DBFRecordset5.Fields(12)
        
        If Len(DBFRecordset5.Fields(3)) >= 1 Then
            Acctnm1 = DBFRecordset5.Fields(3)
        Else
            Acctnm1 = ""
        End If
        
        If Len(DBFRecordset5.Fields(4)) >= 1 Then
            Acctnm2 = DBFRecordset5.Fields(4)
        Else
            Acctnm2 = ""
        End If
        
        If LoopCountDr < 1 Then
    
        Set DBFConnector2 = CreateObject("ADODB.Connection")
 
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "SELECT * FROM REF WHERE Chktype = '" & CheckType & "' AND Rtno = '" & Routingno & "'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        LastSerial = dbfRecordset2.Fields(6)
        Serial = Val(LastSerial) + 1
        Serial = Format(Serial, "0000000000")

        Do Until Val(Books) = 0
        
            EndingSerial = Val(Serial) + 99
            EndingSerial = Format(EndingSerial, "0000000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & " " & EndingSerial & " " & Status
        
        'INSERT TO PACKC
        Result = PackC13(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
        
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                     SBTC - SUMMARY OF BLOCK - COMMERCIAL"
            Print #1, "                               STARTER 100 PIECES"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO       START_NO.  END_NO."
            
            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        
        Serial = Val(Serial) + 100
        Serial = Format(Serial, "0000000000")
        
        Me.Caption = "BLOCKC 10 Digits... " & Serial
    Loop
 
 Else
        Do Until Val(Books) = 0
        Serial = (EndingSerial) + 1
        Serial = Format(Serial, "0000000000")
        
        EndingSerial = Val(Serial) + 99
        EndingSerial = Format(EndingSerial, "0000000000")
        
        If BlockCount Mod 4 = 0 Then
            Print #1, " "
            Counting = Val(Counting) + 1
            Print #1, "       ** BLOCK   " & Val(BlockCountDisplay)
            Counting = Val(Counting) + 1
            BlockCountDisplay = Val(BlockCountDisplay) + 1
        End If
        
        BlockCountDisplay2 = Val(BlockCountDisplay) - 1
        
        Do Until Len(BlockCountDisplay2) = 3
            BlockCountDisplay2 = " " & BlockCountDisplay2
        Loop
        
        'PRINTING
        Print #1, "          " & BlockCountDisplay2 & " " & Routingno & "   " & AccountNo & "    " & Serial & " " & EndingSerial & " " & Status
        
        'INSERT TO PACKC
        Result = PackC10(PathOfFile, Batchno, BlockCountDisplay2, Routingno, AccountNo, AcctNoWithHyphen, CheckType, Acctnm1, Acctnm2, Books, Serial, Serial, EndingSerial, EndingSerial)
       
        Counting = Val(Counting) + 1
        
        If Val(Counting) >= 48 Then
            Print #1, ""
            Print #1, ""
            Print #1, "        Page No.     " & PageDisplay
            Print #1, "        " & Format(Now, "MM/DD/YYYY")
            Print #1, "                     SBTC - SUMMARY OF BLOCK - COMMERCIAL"
            Print #1, "                               STARTER 100 PIECES"
            Print #1, ""
            Print #1, "            RUSH, CANCELLED AND HOLD PRINT JOBS SHOULD ALWAYS HAVE"
            Print #1, "                   AN ATTACHED MEMO OR EMAIL FROM THE BANK!!!"
            Print #1,
            Print #1, "        BLOCK RT_NO     M ACCT_NO       START_NO.  END_NO."

            Counting = 0
            PageDisplay = Val(PageDisplay) + 1
        End If
        
        BlockCount = Val(BlockCount) + 1
        Books = Val(Books) - 1
        
        Serial = Val(Serial) + 100
        Serial = Format(Serial, "0000000000")
    
        Me.Caption = "BLOCKC 10 Digits... " & Serial
    Loop
End If
            DBFRecordset5.MoveNext
            LoopCountDr = LoopCountDr + 1
        Loop
        
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Close #1
End Sub


Sub PackPDBF_MC()
ItemType = "MC"

If Dir$(App.Path & "\MC\PACKP.dbf") <> "" Then Kill (App.Path & "\MC\PACKP.dbf")
FileCopy App.Path & "\Datasource\PACKP.dbf", App.Path & "\MC\PACKP.dbf"

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT (Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = '" & ItemType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

Counting = 0
BlockCount = 0

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount

RoutingNumber = dbfRecordset.Fields(0)

'Read DBF File
Set DBFConnector1 = CreateObject("ADODB.Connection")

DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector1.CursorLocation = adUseClient
    
Set DBFRecordset1 = CreateObject("ADODB.Recordset")
SQL1 = "SELECT * FROM SBTCNEW WHERE Chktype = 'A' And Rtno = '" & RoutingNumber & "' And Itemtype = '" & ItemType & "' ORDER BY Rtno, Acctno"
DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
'end Read DBF BOF

LoopCount1 = 0

Do Until LoopCount1 = DBFRecordset1.RecordCount
    Rtno = DBFRecordset1.Fields(1)
    Acctno = DBFRecordset1.Fields(2)
    Acctnm1 = DBFRecordset1.Fields(3)
    Acctnm2 = DBFRecordset1.Fields(4)
    Orderqty = DBFRecordset1.Fields(5)
    Batchno = DBFRecordset1.Fields(7)
    
    If LoopCount1 < 1 Then
    
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC\;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
    
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM REF WHERE Rtno = '" & Rtno & "' AND Chktype = 'A'"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
   
    LASTNO = dbfRecordset2.Fields(6)
   
    Serial = Val(LASTNO) + 1
   
    Do Until Len(Serial) = 10
        Serial = "0" & Serial
    Loop
    
    If Len(Acctnm1) >= 1 Then Acctnm1 = Replace(Acctnm1, "'", "''")
    If Len(Acctnm2) >= 1 Then Acctnm2 = Replace(Acctnm2, "'", "''")
    
    Do Until Val(Orderqty) = 0
        If Counting Mod 4 = 0 Then BlockCount = Val(BlockCount) + 1
        
        EndingSerial = Val(Serial) + 49
        
        Do Until Len(EndingSerial) = 10
            EndingSerial = "0" & EndingSerial
        Loop
 
        Set DBFConnector4 = CreateObject("ADODB.Connection")

        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
    
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
 
        SQL4 = "INSERT INTO PACKP (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" & Batchno & "','" & Val(BlockCount) & "','" & Rtno & "','" & Acctno & "','" & Mid(Acctno, 1, 3) & "-" & Mid(Acctno, 4, 6) & "-" & Mid(Acctno, 10, 3) & "','A','" & Acctnm1 & "','" & Acctnm2 & "','1','" & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
           
        Serial = Val(Serial) + 50
        
        Do Until Len(Serial) = 10
            Serial = "0" & Serial
        Loop
        
        Orderqty = Val(Orderqty) - 1
        Counting = Val(Counting) + 1
    
    Me.Caption = "PACKP.dbf..." & Acctno & " " & LASTNO & ""
    
    Loop
    
    Else
        
    Serial = Val(EndingSerial) + 1
   
    Do Until Len(Serial) = 10
        Serial = "0" & Serial
    Loop
    
    If Len(Acctnm1) >= 1 Then Acctnm1 = Replace(Acctnm1, "'", "''")
    If Len(Acctnm2) >= 1 Then Acctnm2 = Replace(Acctnm2, "'", "''")
    
    Do Until Val(Orderqty) = 0
        If Counting Mod 4 = 0 Then BlockCount = Val(BlockCount) + 1
        
        EndingSerial = Val(Serial) + 49
        
        Do Until Len(EndingSerial) = 10
            EndingSerial = "0" & EndingSerial
        Loop
 
        Set DBFConnector4 = CreateObject("ADODB.Connection")

        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
    
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
 
        SQL4 = "INSERT INTO PACKP (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" & Batchno & "','" & Val(BlockCount) & "','" & Rtno & "','" & Acctno & "','" & Mid(Acctno, 1, 3) & "-" & Mid(Acctno, 4, 6) & "-" & Mid(Acctno, 10, 3) & "','A','" & Acctnm1 & "','" & Acctnm2 & "','1','" & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
           
        Serial = Val(Serial) + 50
        
        Do Until Len(Serial) = 10
            Serial = "0" & Serial
        Loop
        
        Orderqty = Val(Orderqty) - 1
        Counting = Val(Counting) + 1
    
    Me.Caption = "PACKP.dbf..." & Acctno & " " & LASTNO
    Loop
    End If
    
    DBFRecordset1.MoveNext
    LoopCount1 = LoopCount1 + 1
Loop
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
dbfRecordset.Close
End Sub

Sub PackPDBF_GC()

ItemType = "GC"

If Dir$(App.Path & "\GIFTCHK\PACKP.dbf") <> "" Then Kill (App.Path & "\GIFTCHK\PACKP.dbf")
FileCopy App.Path & "\Datasource\PACKP.dbf", App.Path & "\GIFTCHK\PACKP.dbf"

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT (Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = '" & ItemType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

Counting = 0
BlockCount = 0

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount

RoutingNumber = dbfRecordset.Fields(0)

'Read DBF File
Set DBFConnector1 = CreateObject("ADODB.Connection")

DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector1.CursorLocation = adUseClient
    
Set DBFRecordset1 = CreateObject("ADODB.Recordset")
SQL1 = "SELECT * FROM SBTCNEW WHERE Chktype = 'A' And Rtno = '" & RoutingNumber & "' And Itemtype = '" & ItemType & "' ORDER BY Rtno, Acctno"
DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
'end Read DBF BOF

LoopCount1 = 0

Do Until LoopCount1 = DBFRecordset1.RecordCount
    Rtno = DBFRecordset1.Fields(1)
    Acctno = DBFRecordset1.Fields(2)
    Acctnm1 = DBFRecordset1.Fields(3)
    Acctnm2 = DBFRecordset1.Fields(4)
    Orderqty = DBFRecordset1.Fields(5)
    Batchno = DBFRecordset1.Fields(7)
    
    If LoopCount1 < 1 Then
    
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK\;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
    
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM REF WHERE Rtno = '" & Rtno & "' AND Chktype = 'A'"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
   
    LASTNO = dbfRecordset2.Fields(6)
   
    Serial = Val(LASTNO) + 1
   
    Do Until Len(Serial) = 6
        Serial = "0" & Serial
    Loop
    
    If Len(Acctnm1) >= 1 Then Acctnm1 = Replace(Acctnm1, "'", "''")
    If Len(Acctnm2) >= 1 Then Acctnm2 = Replace(Acctnm2, "'", "''")
    
    Do Until Val(Orderqty) = 0
        If Counting Mod 4 = 0 Then BlockCount = Val(BlockCount) + 1
        
        EndingSerial = Val(Serial) + 49
        
        Do Until Len(EndingSerial) = 6
            EndingSerial = "0" & EndingSerial
        Loop
 
        Set DBFConnector4 = CreateObject("ADODB.Connection")

        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
    
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
 
        SQL4 = "INSERT INTO PACKP (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" & Batchno & "','" & Val(BlockCount) & "','" & Rtno & "','" & Acctno & "','" & Mid(Acctno, 1, 3) & "-" & Mid(Acctno, 4, 6) & "-" & Mid(Acctno, 10, 3) & "','A','" & Acctnm1 & "','" & Acctnm2 & "','1','" & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
           
        Serial = Val(Serial) + 50
        
        Do Until Len(Serial) = 6
            Serial = "0" & Serial
        Loop
        
        Orderqty = Val(Orderqty) - 1
        Counting = Val(Counting) + 1
    
    Me.Caption = "PACKP.dbf..." & Acctno & " " & LASTNO & ""
    
    Loop
    
    Else
        
    Serial = Val(EndingSerial) + 1
   
    Do Until Len(Serial) = 6
        Serial = "0" & Serial
    Loop
    
    If Len(Acctnm1) >= 1 Then Acctnm1 = Replace(Acctnm1, "'", "''")
    If Len(Acctnm2) >= 1 Then Acctnm2 = Replace(Acctnm2, "'", "''")
    
    Do Until Val(Orderqty) = 0
        If Counting Mod 4 = 0 Then BlockCount = Val(BlockCount) + 1
        
        EndingSerial = Val(Serial) + 49
        
        Do Until Len(EndingSerial) = 6
            EndingSerial = "0" & EndingSerial
        Loop
 
        Set DBFConnector4 = CreateObject("ADODB.Connection")

        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
    
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
 
        SQL4 = "INSERT INTO PACKP (BatchNo,Block,[RT_NO],[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" & Batchno & "','" & Val(BlockCount) & "','" & Rtno & "','" & Acctno & "','" & Mid(Acctno, 1, 3) & "-" & Mid(Acctno, 4, 6) & "-" & Mid(Acctno, 10, 3) & "','A','" & Acctnm1 & "','" & Acctnm2 & "','1','" & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
           
        Serial = Val(Serial) + 50
        
        Do Until Len(Serial) = 6
            Serial = "0" & Serial
        Loop
        
        Orderqty = Val(Orderqty) - 1
        Counting = Val(Counting) + 1
    
    Me.Caption = "PACKP.dbf..." & Acctno & " " & LASTNO
    Loop
    End If
    
    DBFRecordset1.MoveNext
    LoopCount1 = LoopCount1 + 1
Loop
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
dbfRecordset.Close
End Sub

Sub PackingDBF13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

If Dir$(App.Path & "\12_Digits\PACKING.dbf") <> "" Then Kill (App.Path & "\12_Digits\PACKING.dbf")
FileCopy App.Path & "\DataSource\PACKING.dbf", App.Path & "\12_Digits\PACKING.dbf"

ChktypeFileName = "PACKP"

CheckLoop:
'Read PACK P/C
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM " & ChktypeFileName & " ORDER BY RT_NO, ACCT_NO, CK_NO_B"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read PACK P/C

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Batchno = dbfRecordset.Fields(0)
    Block = dbfRecordset.Fields(1)
    Rtno = dbfRecordset.Fields(2)
    
    'GET BRANCH FROM INPUT FILE
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT Branchname FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'GET BRANCH FROM INPUT FILE
    
    BranchName = UCase(Replace(DBFRecordset1.Fields(0), "'", "''"))
    
    ACCT_NO = dbfRecordset.Fields(4)
    ACCT_NO_P = dbfRecordset.Fields(5)
    ChkType = dbfRecordset.Fields(6)
    
    If Len(dbfRecordset.Fields(7)) >= 1 Then
        Acctnm1 = UCase(Replace(dbfRecordset.Fields(7), "'", "''"))
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(8)) >= 1 Then
        Acctnm2 = UCase(Replace(dbfRecordset.Fields(8), "'", "''"))
    Else
        Acctnm2 = ""
    End If
    
    Serial = dbfRecordset.Fields(11)
    EndingSerial = dbfRecordset.Fields(13)
 
    'Insert into Packing
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
    
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
         
    SQL2 = "INSERT INTO PACKING (BatchNo,Block,[RT_NO],Branch,[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" _
            & Batchno & "','" & Block & "','" & Rtno & "','" & BranchName & "','" & ACCT_NO & "','" _
            & ACCT_NO_P & "','" & ChkType & "','" & Acctnm1 & "','" & Acctnm2 & "','1','" _
            & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
    
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    'End Insert into Packing
    
    Me.Caption = "PACKING.dbf 13 Digits..." & Acctno & " " & Serial

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
dbfRecordset.Close

If ChktypeFileName = "PACKP" Then
    ChktypeFileName = "PACKC"
    DBFConnector.Close
    GoTo CheckLoop
    Exit Sub
End If
End Sub

Sub PackingDBF10()

'13 DIGITS
PathOfFile = App.Path & "\10_Digits"

If Dir$(App.Path & "\10_Digits\PACKING.dbf") <> "" Then Kill (App.Path & "\10_Digits\PACKING.dbf")
FileCopy App.Path & "\DataSource\PACKING.dbf", App.Path & "\10_Digits\PACKING.dbf"

ChktypeFileName = "PACKP"

CheckLoop:
'Read PACK P/C
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM " & ChktypeFileName & " ORDER BY RT_NO, ACCT_NO, CK_NO_B"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read PACK P/C

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Batchno = dbfRecordset.Fields(0)
    Block = dbfRecordset.Fields(1)
    Rtno = dbfRecordset.Fields(2)
    
    'GET BRANCH FROM INPUT FILE
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT Branchname FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'GET BRANCH FROM INPUT FILE
    
    BranchName = UCase(Replace(DBFRecordset1.Fields(0), "'", "''"))
    
    ACCT_NO = dbfRecordset.Fields(4)
    ACCT_NO_P = dbfRecordset.Fields(5)
    ChkType = dbfRecordset.Fields(6)
    
    If Len(dbfRecordset.Fields(7)) >= 1 Then
        Acctnm1 = UCase(Replace(dbfRecordset.Fields(7), "'", "''"))
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(8)) >= 1 Then
        Acctnm2 = UCase(Replace(dbfRecordset.Fields(8), "'", "''"))
    Else
        Acctnm2 = ""
    End If
    
    Serial = dbfRecordset.Fields(11)
    EndingSerial = dbfRecordset.Fields(13)
 
    'Insert into Packing
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
    
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
         
    SQL2 = "INSERT INTO PACKING (BatchNo,Block,[RT_NO],Branch,[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" _
            & Batchno & "','" & Block & "','" & Rtno & "','" & BranchName & "','" & ACCT_NO & "','" _
            & ACCT_NO_P & "','" & ChkType & "','" & Acctnm1 & "','" & Acctnm2 & "','1','" _
            & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
    
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    'End Insert into Packing
    
    Me.Caption = "PACKING.dbf 10 Digits..." & Acctno & " " & Serial

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
dbfRecordset.Close

If ChktypeFileName = "PACKP" Then
    ChktypeFileName = "PACKC"
    DBFConnector.Close
    GoTo CheckLoop
    Exit Sub
End If
End Sub

Sub PackingDBF_MC()
If Dir$(App.Path & "\MC\PACKING.dbf") <> "" Then Kill (App.Path & "\MC\PACKING.dbf")
FileCopy App.Path & "\Datasource\PACKING.dbf", App.Path & "\MC\PACKING.dbf"

ChktypeFileName = "PACKP"

CheckLoop:
'Read PACK P/C
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM " & ChktypeFileName & " ORDER BY RT_NO, ACCT_NO, CK_NO_B"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read PACK P/C

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Batchno = dbfRecordset.Fields(0)
    Block = dbfRecordset.Fields(1)
    Rtno = dbfRecordset.Fields(2)
    
    'GET BRANCH FROM INPUT FILE
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'GET BRANCH FROM INPUT FILE
    
    BranchName = Replace(DBFRecordset1.Fields(8), "'", "''")
    
    ACCT_NO = dbfRecordset.Fields(4)
    ACCT_NO_P = dbfRecordset.Fields(5)
    ChkType = dbfRecordset.Fields(6)
    
    If Len(dbfRecordset.Fields(7)) >= 1 Then
        Acctnm1 = Replace(dbfRecordset.Fields(7), "'", "''")
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(8)) >= 1 Then
        Acctnm2 = Replace(dbfRecordset.Fields(8), "'", "''")
    Else
        Acctnm2 = ""
    End If
    
    Serial = dbfRecordset.Fields(11)
    EndingSerial = dbfRecordset.Fields(13)
 
    'Insert into Packing
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
    
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
         
    SQL2 = "INSERT INTO PACKING (BatchNo,Block,[RT_NO],Branch,[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" & Batchno & "','" & Block & "','" & Rtno & "','" & BranchName & "','" & ACCT_NO & "','" & ACCT_NO_P & "','" & ChkType & "','" & Acctnm1 & "','" & Acctnm2 & "','1','" & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    'End Insert into Packing
    
    Me.Caption = "PACKING.dbf..." & Acctno & " " & Serial

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
dbfRecordset.Close
End Sub

Sub PackingDBF_GC()

If Dir$(App.Path & "\GIFTCHK\PACKING.dbf") <> "" Then Kill (App.Path & "\GIFTCHK\PACKING.dbf")
FileCopy App.Path & "\Datasource\PACKING.dbf", App.Path & "\GIFTCHK\PACKING.dbf"

ChktypeFileName = "PACKP"

CheckLoop:
'Read PACK P/C
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM " & ChktypeFileName & " ORDER BY RT_NO, ACCT_NO, CK_NO_B"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read PACK P/C

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Batchno = dbfRecordset.Fields(0)
    Block = dbfRecordset.Fields(1)
    Rtno = dbfRecordset.Fields(2)
    
    'GET BRANCH FROM INPUT FILE
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'GET BRANCH FROM INPUT FILE
    
    BranchName = Replace(DBFRecordset1.Fields(8), "'", "''")
    
    ACCT_NO = dbfRecordset.Fields(4)
    ACCT_NO_P = dbfRecordset.Fields(5)
    ChkType = dbfRecordset.Fields(6)
    
    If Len(dbfRecordset.Fields(7)) >= 1 Then
        Acctnm1 = Replace(dbfRecordset.Fields(7), "'", "''")
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(8)) >= 1 Then
        Acctnm2 = Replace(dbfRecordset.Fields(8), "'", "''")
    Else
        Acctnm2 = ""
    End If
    
    Serial = dbfRecordset.Fields(11)
    EndingSerial = dbfRecordset.Fields(13)
 
    'Insert into Packing
    Set DBFConnector2 = CreateObject("ADODB.Connection")

    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
    
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
         
    SQL2 = "INSERT INTO PACKING (BatchNo,Block,[RT_NO],Branch,[ACCT_NO],[ACCT_NO_P],CHKTYPE,[ACCT_NAME1],[ACCT_NAME2],[NO_BKS],[CK_NO_P],[CK_NO_B],[CK_NOE],[CK_NO_E]) VALUES ('" & Batchno & "','" & Block & "','" & Rtno & "','" & BranchName & "','" & ACCT_NO & "','" & ACCT_NO_P & "','" & ChkType & "','" & Acctnm1 & "','" & Acctnm2 & "','1','" & Serial & "','" & Serial & "','" & EndingSerial & "','" & EndingSerial & "')"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    'End Insert into Packing
    
    Me.Caption = "PACKING.dbf..." & Acctno & " " & Serial

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
dbfRecordset.Close
End Sub

Sub PackingATxt_MC()

Close #1
Open App.Path & "\MC\PACKINGA.txt" For Output As #1

PageNo = 1
Total = 0
LineCount = 0

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct ([Rt_no]) FROM PACKP"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do Until LoopCount = dbfRecordset.RecordCount
Routing = dbfRecordset.Fields(0)

Print #1,
Print #1, "  Page No.     " & PageNo
Print #1, "  " & Format(Now, "mm/dd/yyyy")
Print #1, "                                CAPTIVE PRINTING CORPORATION"
Print #1, "                              SBTC - Manager's Checks Summary"
Print #1, ""
Print #1, "  ACCT_NO          ACCOUNT NAME                     QTY CT START #    END #"
Print #1,
Print #1,

    'Read from PACKP
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM PACKP WHERE Rt_no = '" & Routing & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from PACKP
    
    Rtno = DBFRecordset1.Fields(2)
    
    'GET BRANCH FROM INPUT FILE
    Set DBFConnector2 = CreateObject("ADODB.Connection")
    
    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
        
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Itemtype = 'MC'"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    'GET BRANCH FROM INPUT FILE
    
    BranchName = dbfRecordset2.Fields(8)
    
    Batchno = DBFRecordset1.Fields(0)
    
    Print #1, " ** ORDERS OF BRSTN " & Rtno & "  " & BranchName
    Print #1,
    Print #1, " * BATCH #: " & Batchno
    
    Subtotal = 0
    LoopCount1 = 0
    
    Do Until LoopCount1 = DBFRecordset1.RecordCount
    AcctNoWithHyphen = DBFRecordset1.Fields(5)
    ChkType = DBFRecordset1.Fields(6)
    Orderqty = DBFRecordset1.Fields(9)
    
    StartSerial = DBFRecordset1.Fields(11)
    Do Until Len(StartSerial) >= 7
        StartSerial = "0" & StartSerial
     Loop
    
    EndSerial = DBFRecordset1.Fields(13)
    Do Until Len(EndSerial) >= 7
        EndSerial = "0" & EndSerial
    Loop
    
    If Len(DBFRecordset1.Fields(7)) >= 1 Then
        Acctnm1 = DBFRecordset1.Fields(7)
    
    Else
        Acctnm1 = ""
    End If
    
    Do Until Len(Acctnm1) >= 40
        Acctnm1 = Acctnm1 & " "
    Loop

     If Len(DBFRecordset1.Fields(8)) >= 1 Then
        Acctnm2 = DBFRecordset1.Fields(8)
        
        Do Until Len(Acctnm2) >= 40
            Acctnm2 = Acctnm2 & " "
        Loop
    Else
        Acctnm2 = ""
    End If

    'Printing
    Print #1, "  " & AcctNoWithHyphen & " " & Acctnm1 & " " & Orderqty & " " & ChkType & "  " & StartSerial & " " & EndSerial
    If Acctnm2 <> "" Then
        Print #1, "                 " & Acctnm2
        LineCount = LineCount + 1
    End If

    Subtotal = Val(Subtotal) + 1

    If LineCount >= 50 Then
        
        PageNo = Val(PageNo) + 1
        Print #1, ""
        Print #1, "  Page No.     " & PageNo
        Print #1, "  " & Format(Now, "mm/dd/yyyy")
        Print #1, "                                CAPTIVE PRINTING CORPORATION"
        Print #1, "                              SBTC - Manager's Checks Summary"
        Print #1, ""
        Print #1, "  ACCT_NO         ACCOUNT NAME                          QTY CT START # END #"
        Print #1,
        Print #1,

        LineCount = 0
    End If

        LineCount = LineCount + 1
        
        DBFRecordset1.MoveNext
        LoopCount1 = LoopCount1 + 1
    Loop
    
    Do Until Len(Subtotal) = 3
        Subtotal = " " & Subtotal
    Loop
    
    Print #1, " * Subsubtotal *                                        "
    Print #1, "                                                        " & Subtotal
    Print #1, " ** Subtotal **                                         "
    Print #1, "                                                        " & Subtotal
    
    Total = Val(Total) + Val(Subtotal)
    
    If LoopCount = Val(dbfRecordset.RecordCount) - 1 Then GoTo PrintTotal
    
    Print #1, ""
    
    Me.Caption = "PACKINGA.txt..." & Acctno & "  " & StartSerial
    
    PageNo = Val(PageNo) + 1
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

PrintTotal:
        Do Until Len(Total) = 3
            Total = " " & Total
        Loop
        
        Print #1, " *** Total ***                                          "
        Print #1, "                                                        " & Total
        Exit Sub
Close #1
End Sub

Sub PackingATxt_GC()

Close #1
Open App.Path & "\GIFTCHK\PACKINGA.txt" For Output As #1

PageNo = 1
Total = 0
LineCount = 0

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct ([Rt_no]) FROM PACKP"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do Until LoopCount = dbfRecordset.RecordCount
Routing = dbfRecordset.Fields(0)

Print #1,
Print #1, "  Page No.     " & PageNo
Print #1, "  " & Format(Now, "mm/dd/yyyy")
Print #1, "                                CAPTIVE PRINTING CORPORATION"
Print #1, "                                 SBTC - Gift Checks Summary"
Print #1, ""
Print #1, "  ACCT_NO          ACCOUNT NAME                     QTY CT START #    END #"
Print #1,
Print #1,

    'Read from PACKP
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM PACKP WHERE Rt_no = '" & Routing & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from PACKP
    
    Rtno = DBFRecordset1.Fields(2)
    
    'GET BRANCH FROM INPUT FILE
    Set DBFConnector2 = CreateObject("ADODB.Connection")
    
    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
        
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Itemtype = 'GC'"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    'GET BRANCH FROM INPUT FILE
    
    BranchName = dbfRecordset2.Fields(8)
    
    Batchno = DBFRecordset1.Fields(0)
    
    Print #1, " ** ORDERS OF BRSTN " & Rtno & "  " & BranchName
    Print #1,
    Print #1, " * BATCH #: " & Batchno
    
    Subtotal = 0
    LoopCount1 = 0
    
    Do Until LoopCount1 = DBFRecordset1.RecordCount
    AcctNoWithHyphen = DBFRecordset1.Fields(5)
    ChkType = DBFRecordset1.Fields(6)
    Orderqty = DBFRecordset1.Fields(9)
    
    StartSerial = DBFRecordset1.Fields(11)
    Do Until Len(StartSerial) >= 7
        StartSerial = "0" & StartSerial
     Loop
    
    EndSerial = DBFRecordset1.Fields(13)
    Do Until Len(EndSerial) >= 7
        EndSerial = "0" & EndSerial
    Loop
    
    If Len(DBFRecordset1.Fields(7)) >= 1 Then
        Acctnm1 = DBFRecordset1.Fields(7)
    
    Else
        Acctnm1 = ""
    End If
    
    Do Until Len(Acctnm1) >= 40
        Acctnm1 = Acctnm1 & " "
    Loop

     If Len(DBFRecordset1.Fields(8)) >= 1 Then
        Acctnm2 = DBFRecordset1.Fields(8)
        
        Do Until Len(Acctnm2) >= 40
            Acctnm2 = Acctnm2 & " "
        Loop
    Else
        Acctnm2 = ""
    End If

    'Printing
    Print #1, "  " & AcctNoWithHyphen & " " & Acctnm1 & " " & Orderqty & " " & ChkType & "  " & StartSerial & " " & EndSerial
    If Acctnm2 <> "" Then
        Print #1, "                 " & Acctnm2
        LineCount = LineCount + 1
    End If

    Subtotal = Val(Subtotal) + 1

    If LineCount >= 50 Then
        
        PageNo = Val(PageNo) + 1
        Print #1, ""
        Print #1, "  Page No.     " & PageNo
        Print #1, "  " & Format(Now, "mm/dd/yyyy")
        Print #1, "                                CAPTIVE PRINTING CORPORATION"
        Print #1, "                                 SBTC - Gift Checks Summary"
        Print #1, ""
        Print #1, "  ACCT_NO          ACCOUNT NAME                     QTY CT START #    END #"
        Print #1,
        Print #1,

        LineCount = 0
    End If

        LineCount = LineCount + 1
        
        DBFRecordset1.MoveNext
        LoopCount1 = LoopCount1 + 1
    Loop
    
    Do Until Len(Subtotal) = 3
        Subtotal = " " & Subtotal
    Loop
    
    Print #1, " * Subsubtotal *                                        "
    Print #1, "                                                        " & Subtotal
    Print #1, " ** Subtotal **                                         "
    Print #1, "                                                        " & Subtotal
    
    Total = Val(Total) + Val(Subtotal)
    
    If LoopCount = Val(dbfRecordset.RecordCount) - 1 Then GoTo PrintTotal
    
    Print #1, ""
    
    Me.Caption = "PACKINGA.txt..." & Acctno & "  " & StartSerial
    
    PageNo = Val(PageNo) + 1
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

PrintTotal:
        Do Until Len(Total) = 3
            Total = " " & Total
        Loop
        
        Print #1, " *** Total ***                                          "
        Print #1, "                                                        " & Total
        Exit Sub
Close #1
End Sub

Function PackingTxt13()

CheckType = "A"
Formtype = "05"

RepeatMe:

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"
Close #1
Open App.Path & "\12_Digits\PACKING" & CheckType & ".txt" For Output As #1


Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct (RT_NO), Branch, BatchNo FROM PACKING WHERE ChkType = '" & CheckType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Total = 0
LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    Batchno = dbfRecordset.Fields(2)
    
    If LoopCount >= 1 Then Print #1, ""
    
    Print #1,
    Print #1, "  Page No.     " & LoopCount + 1
    Print #1, "  " & Format(Now, "mm/dd/yyyy")
    Print #1, "                             CAPTIVE PRINTING CORPORATION"
    
    If CheckType = "A" Then Print #1, "                            SBTC - Personal Checks Summary"
    If CheckType = "B" Then Print #1, "                           SBTC - Commercial Checks Summary"
    
    Print #1, ""
    Print #1, "  ACCT_NO         ACCOUNT NAME                   QTY CT START #    END #"
    Print #1,
    Print #1,
    Print #1, " ** ORDERS OF BRSTN " & Rtno & "  " & BranchName
    Print #1,
    Print #1, " * BATCH #: " & Batchno
    
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient

    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT Acct_No_P, Acct_Name1, Acct_Name2, NO_BKS, CK_NO_B, CK_NO_E FROM PACKING WHERE ChkType = '" _
          & CheckType & "' AND RT_NO = '" & Rtno & "'"
    DBFRecordset1.Open SQL, DBFConnector1, 1, 1
    
    Subtotal = 0
    LoopCount1 = 0
    Do Until LoopCount1 = DBFRecordset1.RecordCount
        Acctno = DBFRecordset1.Fields(0)
        
        If Len(DBFRecordset1.Fields(1)) >= 1 Then
            Acctnm1 = DBFRecordset1.Fields(1)
        Else
            Acctnm1 = ""
        End If
        
        If Len(DBFRecordset1.Fields(2)) >= 1 Then
            Acctnm2 = DBFRecordset1.Fields(2)
        Else
            Acctnm2 = ""
        End If
        
        Orderqty = DBFRecordset1.Fields(3)
        StartSerial = DBFRecordset1.Fields(4)
        EndSerial = DBFRecordset1.Fields(5)
        
        Do Until Len(StartSerial) = 11
            StartSerial = StartSerial & " "
        Loop
        
        Do Until Len(Acctnm1) >= 32
            Acctnm1 = Acctnm1 & " "
        Loop
        
        If CheckType = "A" Then
            Print #1, "  " & Acctno & "  " & Acctnm1 & " " & Orderqty & " " & CheckType & "  " & StartSerial & EndSerial
            If Acctnm2 <> "" Then
                Print #1, "                  " & Acctnm2
                    LineCount = LineCount + 1
            End If
        
        Else
            Print #1, "  " & Acctno & "  " & Acctnm1 & " " & Orderqty & " " & CheckType & "  " & StartSerial & "" & EndSerial
            If Acctnm2 <> "" Then
                Print #1, "                  " & Acctnm2
                    LineCount = LineCount + 1
            End If
        End If
            
        Total = Total + Orderqty
        Subtotal = Subtotal + Orderqty
        
        Do Until Len(Subtotal) = 3
            Subtotal = " " & Subtotal
        Loop
        
        Do Until Len(Total) = 3
            Total = " " & Total
        Loop
        
        frmMain.Caption = "Packing.txt 13 Digits..." & StartSerial
        
        DBFRecordset1.MoveNext
        LoopCount1 = LoopCount1 + 1
    Loop

    Print #1, " * Subsubtotal *                                 " & Subtotal
    Print #1,
    Print #1, " ** Subtotal **                                  " & Subtotal
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Print #1,
Print #1,
Print #1, " *** Total ***                                   " & Total
Close #1

If CheckType = "A" And Formtype = "05" Then
    CheckType = "B"
    Formtype = "16"
    GoTo RepeatMe
    Exit Function
End If
End Function

Function PackingTxt10()

CheckType = "A"
Formtype = "05"

RepeatMe:

'13 DIGITS
PathOfFile = App.Path & "\10_Digits"
Close #1
Open App.Path & "\10_Digits\PACKING" & CheckType & ".txt" For Output As #1


Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct (RT_NO), Branch, BatchNo FROM PACKING WHERE ChkType = '" & CheckType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Total = 0
LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    Batchno = dbfRecordset.Fields(2)
    
    If LoopCount >= 1 Then Print #1, ""
    
    Print #1,
    Print #1, "  Page No.     " & LoopCount + 1
    Print #1, "  " & Format(Now, "mm/dd/yyyy")
    Print #1, "                             CAPTIVE PRINTING CORPORATION"
    
    If CheckType = "A" Then Print #1, "                            SBTC - Personal Checks Summary"
    If CheckType = "B" Then Print #1, "                           SBTC - Commercial Checks Summary"
    
    Print #1, ""
    Print #1, "  ACCT_NO         ACCOUNT NAME                   QTY CT START #    END #"
    Print #1,
    Print #1,
    Print #1, " ** ORDERS OF BRSTN " & Rtno & "  " & BranchName
    Print #1,
    Print #1, " * BATCH #: " & Batchno
    
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient

    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT Acct_No_P, Acct_Name1, Acct_Name2, NO_BKS, CK_NO_B, CK_NO_E FROM PACKING WHERE ChkType = '" _
          & CheckType & "' AND RT_NO = '" & Rtno & "'"
    DBFRecordset1.Open SQL, DBFConnector1, 1, 1
    
    Subtotal = 0
    LoopCount1 = 0
    Do Until LoopCount1 = DBFRecordset1.RecordCount
        Acctno = DBFRecordset1.Fields(0)
        
        If Len(DBFRecordset1.Fields(1)) >= 1 Then
            Acctnm1 = DBFRecordset1.Fields(1)
        Else
            Acctnm1 = ""
        End If
        
        If Len(DBFRecordset1.Fields(2)) >= 1 Then
            Acctnm2 = DBFRecordset1.Fields(2)
        Else
            Acctnm2 = ""
        End If
        
        Orderqty = DBFRecordset1.Fields(3)
        StartSerial = DBFRecordset1.Fields(4)
        EndSerial = DBFRecordset1.Fields(5)
        
        Do Until Len(StartSerial) = 11
            StartSerial = StartSerial & " "
        Loop
        
        Do Until Len(Acctnm1) >= 32
            Acctnm1 = Acctnm1 & " "
        Loop
        
        If CheckType = "A" Then
            Print #1, "  " & Acctno & "  " & Acctnm1 & " " & Orderqty & " " & CheckType & "  " & StartSerial & EndSerial
            If Acctnm2 <> "" Then
                Print #1, "                  " & Acctnm2
                    LineCount = LineCount + 1
            End If
        
        Else
            Print #1, "  " & Acctno & "  " & Acctnm1 & " " & Orderqty & " " & CheckType & "  " & StartSerial & "" & EndSerial
            If Acctnm2 <> "" Then
                Print #1, "                  " & Acctnm2
                    LineCount = LineCount + 1
            End If
        End If
            
        Total = Val(Total) + Orderqty
        Subtotal = Subtotal + Orderqty
        
        Do Until Len(Subtotal) = 3
            Subtotal = " " & Subtotal
        Loop
        
        Do Until Len(Total) >= 3
            Total = " " & Total
        Loop
        
        frmMain.Caption = "Packing.txt 10 Digits..." & StartSerial
        
        DBFRecordset1.MoveNext
        LoopCount1 = LoopCount1 + 1
    Loop

    Print #1, " * Subsubtotal *                                 " & Subtotal
    Print #1,
    Print #1, " ** Subtotal **                                  " & Subtotal
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Print #1,
Print #1,
Print #1, " *** Total ***                                   " & Total
Close #1

If CheckType = "A" And Formtype = "05" Then
    CheckType = "B"
    Formtype = "16"
    GoTo RepeatMe
    Exit Function
End If
End Function

Sub TransP13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

If Dir$(App.Path & "\12_Digits\TRANSP.dbf") <> "" Then Kill (App.Path & "\12_Digits\TRANSP.dbf")
FileCopy App.Path & "\DataSource\TRANSP.dbf", App.Path & "\12_Digits\TRANSP.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'PA' AND Digits = '13' ORDER BY Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
Rtno = dbfRecordset.Fields(0)

Set DBFConnector2 = CreateObject("ADODB.Connection")

DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector2.CursorLocation = adUseClient
    
Set dbfRecordset2 = CreateObject("ADODB.Recordset")
SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = 'A' AND Itemtype = 'PA' AND Digits = '13' ORDER BY Rtno,Acctno"
dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

LoopCount2 = 0

Do Until LoopCount2 = dbfRecordset2.RecordCount

    Routingno = dbfRecordset2.Fields(1)
    Routing2 = dbfRecordset2.Fields(1)
    AccountNo = dbfRecordset2.Fields(2)
    Accountno2 = dbfRecordset2.Fields(2)
    Books = Val(dbfRecordset2.Fields(5))
    Ctype = "A"
    Digits = dbfRecordset2.Fields(15)
    
    If Len(dbfRecordset2.Fields(3)) >= 1 Then
        Name1 = UCase(Replace(dbfRecordset2.Fields(3), "''", "'"))
    Else
        Name1 = ""
    End If
    
    If Len(dbfRecordset2.Fields(4)) >= 1 Then
        Name2 = UCase(Replace(dbfRecordset2.Fields(4), "''", "'"))
    Else
        Name2 = ""
    End If
    
    AccountnoWH = Mid(Accountno2, 1, 3) & "-" & Mid(Accountno2, 4, 6) & "-" & Mid(Accountno2, 10, 3)
    Routing1to5 = Mid(Routing2, 1, 5)
    Routing6to9 = Mid(Routing2, 6, 4)
    
        If LoopCount2 < 1 Then
            Temp = GetSerial(Routingno, Books, Ctype, Digits)
            
            TransPStartSerial = Format(TransPStartSerial, "0000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000")
    
        Else
           TransPStartSerial = TransPEndSerial
           
            TotalBooks = Books
            
            Do Until TotalBooks = 0
                TransPEndSerial = Val(TransPEndSerial) + 50
                TotalBooks = TotalBooks - 1
            Loop
    
            TransPStartSerial = Format(TransPStartSerial, "0000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000")
    
        End If
            
        Set DBFConnector4 = CreateObject("ADODB.Connection")
        
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO TRANSP ([Rt_no],[Acct_no],[No_bks],[Ck_no_p],[Rt_no_p1],[Rt_no_p2],[Acct_no_p],[Acct_name1],[Acct_name2],Sn,C,Blank) VALUES ('" _
                & Routingno & "','" & AccountNo & "','" & Books & "','" & TransPStartSerial & "','" _
                & Routing1to5 & "','" & Routing6to9 & "','" & AccountnoWH & "','" _
                & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','SN','C','XXXXXXXXXXXX')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
            
            dbfRecordset2.MoveNext
            LoopCount2 = LoopCount2 + 1
        Loop
    
    Me.Caption = "TRANSP 13 Digits..." & Routingno
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
End Sub

Sub TransP10()

'13 DIGITS
PathOfFile = App.Path & "\10_Digits"

If Dir$(App.Path & "\10_Digits\TRANSP.dbf") <> "" Then Kill (App.Path & "\10_Digits\TRANSP.dbf")
FileCopy App.Path & "\DataSource\TRANSP.dbf", App.Path & "\10_Digits\TRANSP.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'PA' AND Digits = '10' ORDER BY Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
Rtno = dbfRecordset.Fields(0)

Set DBFConnector2 = CreateObject("ADODB.Connection")

DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector2.CursorLocation = adUseClient
    
Set dbfRecordset2 = CreateObject("ADODB.Recordset")
SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = 'A' AND Itemtype = 'PA' AND Digits = '10' ORDER BY Rtno,Acctno"
dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

LoopCount2 = 0

Do Until LoopCount2 = dbfRecordset2.RecordCount

    Routingno = dbfRecordset2.Fields(1)
    Routing2 = dbfRecordset2.Fields(1)
    AccountNo = dbfRecordset2.Fields(2)
    Accountno2 = dbfRecordset2.Fields(2)
    Books = Val(dbfRecordset2.Fields(5))
    Ctype = "A"
    Digits = dbfRecordset2.Fields(15)
    
    If Len(dbfRecordset2.Fields(3)) >= 1 Then
        Name1 = UCase(Replace(dbfRecordset2.Fields(3), "''", "'"))
    Else
        Name1 = ""
    End If
    
    If Len(dbfRecordset2.Fields(4)) >= 1 Then
        Name2 = UCase(Replace(dbfRecordset2.Fields(4), "''", "'"))
    Else
        Name2 = ""
    End If
    
    AccountnoWH = Mid(Accountno2, 1, 3) & "-" & Mid(Accountno2, 4, 6) & "-" & Mid(Accountno2, 10, 3)
    Routing1to5 = Mid(Routing2, 1, 5)
    Routing6to9 = Mid(Routing2, 6, 4)
    
        If LoopCount2 < 1 Then
            Temp = GetSerial(Routingno, Books, Ctype, Digits)
            
            TransPStartSerial = Format(TransPStartSerial, "0000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000")
    
        Else
           TransPStartSerial = TransPEndSerial
           
            TotalBooks = Books
            
            Do Until TotalBooks = 0
                TransPEndSerial = Val(TransPEndSerial) + 50
                TotalBooks = TotalBooks - 1
            Loop
    
            TransPStartSerial = Format(TransPStartSerial, "0000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000")
    
        End If
            
        Set DBFConnector4 = CreateObject("ADODB.Connection")
        
            
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO TRANSP ([Rt_no],[Acct_no],[No_bks],[Ck_no_p],[Rt_no_p1],[Rt_no_p2],[Acct_no_p],[Acct_name1],[Acct_name2],Sn,C,Blank) VALUES ('" _
                & Routingno & "','" & AccountNo & "','" & Books & "','" & TransPStartSerial & "','" _
                & Routing1to5 & "','" & Routing6to9 & "','" & AccountnoWH & "','" _
                & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','SN','C','XXXXXXXXXXXX')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
            
                Me.Caption = "TRANSP 10 Digits..." & Routingno & "   " & FormatNumber((LoopCount / dbfRecordset.RecordCount) * 100, 5) & "       " & FormatNumber((LoopCount2 / dbfRecordset2.RecordCount) * 100, 5)
                
            dbfRecordset2.MoveNext
            LoopCount2 = LoopCount2 + 1
        Loop
    
    'Me.Caption = "TRANSP 10 Digits..." & Routingno & "   " & FormatNumber((LoopCount / dbfRecordset.RecordCount) * 100, 5)
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
End Sub

Sub TransP_MC()

If Dir$(App.Path & "\MC\TRANSP.DBF") <> "" Then Kill (App.Path & "\MC\TRANSP.DBF")
FileCopy App.Path & "\Datasource\TRANSP.dbf", App.Path & "\MC\TRANSP.DBF"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'MC' ORDER BY Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)
    
    Set DBFConnector2 = CreateObject("ADODB.Connection")
    
    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
        
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = 'A' AND Itemtype = 'MC' ORDER BY Rtno,Acctno"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
    
    LoopCount2 = 0

    Do Until LoopCount2 = dbfRecordset2.RecordCount
    
    Routingno = dbfRecordset2.Fields(1)
    Routing2 = dbfRecordset2.Fields(1)
    AccountNo = dbfRecordset2.Fields(2)
    Accountno2 = dbfRecordset2.Fields(2)
    Books = Val(dbfRecordset2.Fields(5))
    Ctype = "A"
    ItemType = "MC"
    
    If Len(dbfRecordset2.Fields(3)) >= 1 Then
        Name1 = Replace(dbfRecordset2.Fields(3), "''", "'")
    Else
        Name1 = ""
    End If
    
    If Len(dbfRecordset2.Fields(4)) >= 1 Then
        Name2 = Replace(dbfRecordset2.Fields(4), "''", "'")
    Else
        Name2 = ""
    End If
    
    AccountnoWH = Mid(Accountno2, 1, 3) & "-" & Mid(Accountno2, 4, 6) & "-" & Mid(Accountno2, 10, 3)
    Routing1to5 = Mid(Routing2, 1, 5)
    Routing6to9 = Mid(Routing2, 6, 4)

    If LoopCount2 < 1 Then
            Temp = GetSerialMC(Routingno, Books, Ctype)
    
            Do Until Len(TransPStartSerial) = 10
                TransPStartSerial = "0" & TransPStartSerial
            Loop
            
            Do Until Len(TransPEndSerial) = 10
                TransPEndSerial = "0" & TransPEndSerial
            Loop
    Else
           TransPStartSerial = TransPEndSerial
           
            TotalBooks = Books
            
            Do Until TotalBooks = 0
                TransPEndSerial = Val(TransPEndSerial) + 50
                TotalBooks = TotalBooks - 1
            Loop

            If Len(TransPStartSerial) < 10 Then
                Do Until Len(TransPStartSerial) = 10
                    TransPStartSerial = "0" & TransPStartSerial
                Loop
            End If
        
            If Len(TransPEndSerial) < 10 Then
                Do Until Len(TransPEndSerial) = 10
                    TransPEndSerial = "0" & TransPEndSerial
                Loop
            End If
    End If
        
    Set DBFConnector4 = CreateObject("ADODB.Connection")
    
    DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
    DBFConnector4.CursorLocation = adUseClient
        
    Set DBFRecordset4 = CreateObject("ADODB.Recordset")
    SQL4 = "INSERT INTO TRANSP ([Rt_no],[Acct_no],[No_bks],[Ck_no_p],[Rt_no_p1],[Rt_no_p2],[Acct_no_p],[Acct_name1],[Acct_name2],Sn,C,Blank) VALUES ('" & Routingno & "','" & AccountNo & "','" & Books & "','" & TransPStartSerial & "','" & Routing1to5 & "','" & Routing6to9 & "','" & AccountnoWH & "','" & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','SN','C','XXXXXXXXXXXX')"
    DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
        
        dbfRecordset2.MoveNext
        LoopCount2 = LoopCount2 + 1
    Loop
    
Me.Caption = "TRANSP.." & Routingno
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
End Sub

Sub TransP_GC()

If Dir$(App.Path & "\GIFTCHK\TRANSP.DBF") <> "" Then Kill (App.Path & "\GIFTCHK\TRANSP.DBF")
FileCopy App.Path & "\Datasource\TRANSP.dbf", App.Path & "\GIFTCHK\TRANSP.DBF"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'GC' ORDER BY Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
    Rtno = dbfRecordset.Fields(0)

    Set DBFConnector2 = CreateObject("ADODB.Connection")
    
    DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector2.CursorLocation = adUseClient
        
    Set dbfRecordset2 = CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = 'A' AND Itemtype = 'GC' ORDER BY Rtno,Acctno"
    dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

    LoopCount2 = 0
    
    Do Until LoopCount2 = dbfRecordset2.RecordCount
    
    Routingno = dbfRecordset2.Fields(1)
    Routing2 = dbfRecordset2.Fields(1)
    AccountNo = dbfRecordset2.Fields(2)
    Accountno2 = dbfRecordset2.Fields(2)
    Books = Val(dbfRecordset2.Fields(5))
    Ctype = "A"
    ItemType = "GC"
    
    If Len(dbfRecordset2.Fields(3)) >= 1 Then
        Name1 = Replace(dbfRecordset2.Fields(3), "''", "'")
    Else
        Name1 = ""
    End If

    If Len(dbfRecordset2.Fields(4)) >= 1 Then
        Name2 = Replace(dbfRecordset2.Fields(4), "''", "'")
    Else
        Name2 = ""
    End If
    
    AccountnoWH = Mid(Accountno2, 1, 3) & "-" & Mid(Accountno2, 4, 6) & "-" & Mid(Accountno2, 10, 3)
    Routing1to5 = Mid(Routing2, 1, 5)
    Routing6to9 = Mid(Routing2, 6, 4)

    If LoopCount2 < 1 Then
            Temp = GetSerialGC(Routingno, Books, Ctype)
    
            Do Until Len(TransPStartSerial) >= 6
                TransPStartSerial = "0" & TransPStartSerial
            Loop
            
            Do Until Len(TransPEndSerial) >= 6
                TransPEndSerial = "0" & TransPEndSerial
            Loop
    Else
           TransPStartSerial = TransPEndSerial
           
            TotalBooks = Books
            
            Do Until TotalBooks = 0
                TransPEndSerial = Val(TransPEndSerial) + 50
                TotalBooks = TotalBooks - 1
            Loop

            If Len(TransPStartSerial) < 6 Then
                Do Until Len(TransPStartSerial) >= 6
                    TransPStartSerial = "0" & TransPStartSerial
                Loop
            End If
        
            If Len(TransPEndSerial) < 6 Then
                Do Until Len(TransPEndSerial) >= 6
                    TransPEndSerial = "0" & TransPEndSerial
                Loop
            End If
    End If
        
    Set DBFConnector4 = CreateObject("ADODB.Connection")
    
    DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
    DBFConnector4.CursorLocation = adUseClient
        
    Set DBFRecordset4 = CreateObject("ADODB.Recordset")
    SQL4 = "INSERT INTO TRANSP ([Rt_no],[Acct_no],[No_bks],[Ck_no_p],[Rt_no_p1],[Rt_no_p2],[Acct_no_p],[Acct_name1],[Acct_name2],Sn,C,Blank) VALUES ('" & Routingno & "','" & AccountNo & "','" & Books & "','" & TransPStartSerial & "','" & Routing1to5 & "','" & Routing6to9 & "','" & AccountnoWH & "','" & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','SN','C','XXXXXXXXXXXX')"
    DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
        
        dbfRecordset2.MoveNext
        LoopCount2 = LoopCount2 + 1
    Loop
    
Me.Caption = "TRANSP.." & Routingno
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
End Sub

Function GetSerial(Routingno, Books, Ctype, Digits)

Set DBFConnector2 = CreateObject("ADODB.Connection")

DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector2.CursorLocation = adUseClient
    
Set dbfRecordset2 = CreateObject("ADODB.Recordset")
SQL2 = "SELECT * FROM REF WHERE Rtno = '" & Routingno & "' AND Chktype = '" & Ctype & "'"
dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
        
    LASTNO = dbfRecordset2.Fields(6)

    TransPStartSerial = Val(LASTNO) + 1
    TransPEndSerial = TransPStartSerial
    TotalBooks = Books
            
        Do Until TotalBooks = 0
            If Digits = "13" Then   '13 DIGITS
                If Ctype = "B" Then
                    TransPEndSerial = Val(TransPEndSerial) + 100
                Else
                    TransPEndSerial = Val(TransPEndSerial) + 50
                End If
                TotalBooks = TotalBooks - 1
            
            Else                    '10 DIGITS
                If Ctype = "B" Then
                    TransPEndSerial = Val(TransPEndSerial) + 50
                Else
                    TransPEndSerial = Val(TransPEndSerial) + 50
                End If
                TotalBooks = TotalBooks - 1
            End If
        Loop
End Function

Function GetSerialMC(Routingno, Books, Ctype)

Set DBFConnector2 = CreateObject("ADODB.Connection")

DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
DBFConnector2.CursorLocation = adUseClient
    
Set dbfRecordset2 = CreateObject("ADODB.Recordset")
SQL2 = "SELECT * FROM REF WHERE Rtno = '" & Routingno & "' AND Chktype = '" & Ctype & "'"
dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
        
    LASTNO = dbfRecordset2.Fields(6)

    TransPStartSerial = Val(LASTNO) + 1
    TransPEndSerial = TransPStartSerial
    TotalBooks = Books
            
        Do Until TotalBooks = 0
            TransPEndSerial = Val(TransPEndSerial) + 50
            TotalBooks = TotalBooks - 1
        Loop
End Function

Function GetSerialGC(Routingno, Books, Ctype)

Set DBFConnector2 = CreateObject("ADODB.Connection")

DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
DBFConnector2.CursorLocation = adUseClient
    
Set dbfRecordset2 = CreateObject("ADODB.Recordset")
SQL2 = "SELECT * FROM REF WHERE Rtno = '" & Routingno & "' AND Chktype = '" & Ctype & "'"
dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
        
    LASTNO = dbfRecordset2.Fields(6)

    TransPStartSerial = Val(LASTNO) + 1
    TransPEndSerial = TransPStartSerial
    TotalBooks = Books
            
        Do Until TotalBooks = 0
            TransPEndSerial = Val(TransPEndSerial) + 50
            TotalBooks = TotalBooks - 1
        Loop
End Function

Sub TransC13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

If Dir$(App.Path & "\12_Digits\TRANSC.dbf") <> "" Then Kill (App.Path & "\12_Digits\TRANSC.dbf")
FileCopy App.Path & "\DataSource\TRANSC.dbf", App.Path & "\12_Digits\TRANSC.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'B' And ItemType = 'CA' AND Digits = '13' ORDER BY Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
Rtno = dbfRecordset.Fields(0)

Set DBFConnector2 = CreateObject("ADODB.Connection")

DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector2.CursorLocation = adUseClient
    
Set dbfRecordset2 = CreateObject("ADODB.Recordset")
SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = 'B' AND Itemtype = 'CA' AND Digits = '13' ORDER BY Rtno,Acctno"
dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

LoopCount2 = 0

Do Until LoopCount2 = dbfRecordset2.RecordCount

    Routingno = dbfRecordset2.Fields(1)
    Routing2 = dbfRecordset2.Fields(1)
    AccountNo = dbfRecordset2.Fields(2)
    Accountno2 = dbfRecordset2.Fields(2)
    Books = Val(dbfRecordset2.Fields(5))
    Ctype = "B"
    Digits = dbfRecordset2.Fields(15)
    
    If Len(dbfRecordset2.Fields(3)) >= 1 Then
        Name1 = dbfRecordset2.Fields(3)
    Else
        Name1 = ""
    End If
    
    If Len(dbfRecordset2.Fields(4)) >= 1 Then
        Name2 = dbfRecordset2.Fields(4)
    Else
        Name2 = ""
    End If
    
    Name1 = Replace(Name1, "'", "''")
    Name2 = Replace(Name2, "'", "''")
    
    AccountnoWH = Mid(Accountno2, 1, 3) & "-" & Mid(Accountno2, 4, 6) & "-" & Mid(Accountno2, 10, 3)
    Routing1to5 = Mid(Routing2, 1, 5)
    Routing6to9 = Mid(Routing2, 6, 4)
    
        If LoopCount2 < 1 Then
            Temp = GetSerial(Routingno, Books, Ctype, Digits)
        
            TransPStartSerial = Format(TransPStartSerial, "0000000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000000")
            
        Else
            TransPStartSerial = TransPEndSerial
            TotalBooks = Books
                
            Do Until TotalBooks = 0
                TransPEndSerial = Val(TransPEndSerial) + 100
                TotalBooks = TotalBooks - 1
            Loop
                
            TransPStartSerial = Format(TransPStartSerial, "0000000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000000")
        End If
            
        Set DBFConnector4 = CreateObject("ADODB.Connection")
        
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO TRANSC ([Rt_no],[Acct_no],[No_bks],[Ck_no_p],[Rt_no_p1],[Rt_no_p2],[Acct_no_p],[Acct_name1],[Acct_name2],Sn,C,Blank) VALUES ('" _
                & Routingno & "','" & AccountNo & "','" & Books & "','" & TransPStartSerial & "','" _
                & Routing1to5 & "','" & Routing6to9 & "','" & AccountnoWH & "','" & Name1 & "','" _
                & Name2 & "','SN','C','XXXXXXXXXXXX')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
           
            dbfRecordset2.MoveNext
            LoopCount2 = LoopCount2 + 1
        Loop
    
    Me.Caption = "TRANSC 13 Digits..." & Routingno

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
End Sub

Sub TransC10()

'10 DIGITS
PathOfFile = App.Path & "\10_Digits"

If Dir$(App.Path & "\10_Digits\TRANSC.dbf") <> "" Then Kill (App.Path & "\10_Digits\TRANSC.dbf")
FileCopy App.Path & "\DataSource\TRANSC.dbf", App.Path & "\10_Digits\TRANSC.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'B' And ItemType = 'CA' AND Digits = '10' ORDER BY Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0

Do Until LoopCount = dbfRecordset.RecordCount
Rtno = dbfRecordset.Fields(0)

Set DBFConnector2 = CreateObject("ADODB.Connection")

DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
DBFConnector2.CursorLocation = adUseClient
    
Set dbfRecordset2 = CreateObject("ADODB.Recordset")
SQL2 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' AND Chktype = 'B' AND Itemtype = 'CA' AND Digits = '10' ORDER BY Rtno,Acctno"
dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

LoopCount2 = 0

Do Until LoopCount2 = dbfRecordset2.RecordCount

    Routingno = dbfRecordset2.Fields(1)
    Routing2 = dbfRecordset2.Fields(1)
    AccountNo = dbfRecordset2.Fields(2)
    Accountno2 = dbfRecordset2.Fields(2)
    Books = Val(dbfRecordset2.Fields(5))
    Ctype = "B"
    Digits = dbfRecordset2.Fields(15)
    
    If Len(dbfRecordset2.Fields(3)) >= 1 Then
        Name1 = dbfRecordset2.Fields(3)
    Else
        Name1 = ""
    End If
    
    If Len(dbfRecordset2.Fields(4)) >= 1 Then
        Name2 = dbfRecordset2.Fields(4)
    Else
        Name2 = ""
    End If
    
    Name1 = Replace(Name1, "'", "''")
    Name2 = Replace(Name2, "'", "''")
    
    AccountnoWH = Mid(Accountno2, 1, 3) & "-" & Mid(Accountno2, 4, 6) & "-" & Mid(Accountno2, 10, 3)
    Routing1to5 = Mid(Routing2, 1, 5)
    Routing6to9 = Mid(Routing2, 6, 4)
    
        If LoopCount2 < 1 Then
            Temp = GetSerial(Routingno, Books, Ctype, Digits)
        
            TransPStartSerial = Format(TransPStartSerial, "0000000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000000")
            
        Else
            TransPStartSerial = TransPEndSerial
            TotalBooks = Books
                
            Do Until TotalBooks = 0
                TransPEndSerial = Val(TransPEndSerial) + 50
                TotalBooks = TotalBooks - 1
            Loop
                
            TransPStartSerial = Format(TransPStartSerial, "0000000000")
            TransPEndSerial = Format(TransPEndSerial, "0000000000")
        End If
            
        Set DBFConnector4 = CreateObject("ADODB.Connection")
        
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO TRANSC ([Rt_no],[Acct_no],[No_bks],[Ck_no_p],[Rt_no_p1],[Rt_no_p2],[Acct_no_p],[Acct_name1],[Acct_name2],Sn,C,Blank) VALUES ('" _
                & Routingno & "','" & AccountNo & "','" & Books & "','" & TransPStartSerial & "','" _
                & Routing1to5 & "','" & Routing6to9 & "','" & AccountnoWH & "','" & Name1 & "','" _
                & Name2 & "','SN','C','XXXXXXXXXXXX')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
           
            dbfRecordset2.MoveNext
            LoopCount2 = LoopCount2 + 1
        Loop
    
    Me.Caption = "TRANSC 10 Digits..." & Routingno

    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
End Sub

Sub PrinterFilePA13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

Close #1
Open App.Path & "\12_Digits\PrinterFileP." & Format(Now, "YY") & "P" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
        Address1 = DBFRecordset1.Fields(1)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
       Address2 = DBFRecordset1.Fields(2)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
       Address3 = DBFRecordset1.Fields(3)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        EndingSerial = Format(EndingSerial, "0000000")

        NextSerial = Val(StartingSerial) + 50
        NextSerial = Format(NextSerial, "0000000")
        
        If Len(Incrementor) < 6 Then
            Incrementor = Format(Incrementor, "000000")
        End If

    'PRINTING
    If Counter = 0 Then
        Print #1, SeparatorValue & "  " & Incrementor & "3"
    Else
        Print #1, "3"
    End If
    
    Counter = Val(Counter) + 1
    
        Print #1, "3"
        Print #1, Rtno
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "A"
        Print #1, "     ONNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, "     ONNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, "SN"
        Print #1, ""
        Print #1, ""
        Print #1, Acctnm2
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "C"
        Print #1, "XXXX"
        Print #1, "XXXX"
        Print #1,
        Print #1,
        Print #1, BranchName
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address1
        Print #1, Address2
        Print #1, Address2
        Print #1, Address3
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
                
        If SeparatorCount Mod 4 = 0 Then
            Print #1, "\"
            Incrementor = Val(Incrementor) + 1
            Counter = 0
        Else
            SeparatorValue = ""
        End If
        
        StartingSerial = Val(StartingSerial) + 50
        StartingSerial = Format(StartingSerial, "0000000")
         
        Orderqty = Val(Orderqty) - 1
        SeparatorCount = Val(SeparatorCount) + 1
    
     Me.Caption = "Personal.P 13 Digits..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub PrinterFilePA10()

'10 DIGITS
PathOfFile = App.Path & "\10_Digits"

Close #1
Open App.Path & "\10_Digits\PrinterFileP." & Format(Now, "YY") & "P" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
        Address1 = DBFRecordset1.Fields(1)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
       Address2 = DBFRecordset1.Fields(2)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
       Address3 = DBFRecordset1.Fields(3)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        EndingSerial = Format(EndingSerial, "0000000")

        NextSerial = Val(StartingSerial) + 50
        NextSerial = Format(NextSerial, "0000000")
        
        If Len(Incrementor) < 6 Then
            Incrementor = Format(Incrementor, "000000")
        End If

    'PRINTING
    If Counter = 0 Then
        Print #1, SeparatorValue & "  " & Incrementor & "3"
    Else
        Print #1, "3"
    End If
    
    Counter = Val(Counter) + 1
    
        Print #1, "3"
        Print #1, Rtno
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "A"
        Print #1, "     ONNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, "     ONNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, "SN"
        Print #1, ""
        Print #1, ""
        Print #1, Acctnm2
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "C"
        Print #1, "XXXX"
        Print #1, "XXXX"
        Print #1,
        Print #1,
        Print #1, BranchName
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address1
        Print #1, Address2
        Print #1, Address2
        Print #1, Address3
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
                
        If SeparatorCount Mod 4 = 0 Then
            Print #1, "\"
            Incrementor = Val(Incrementor) + 1
            Counter = 0
        Else
            SeparatorValue = ""
        End If
        
        StartingSerial = Val(StartingSerial) + 50
        StartingSerial = Format(StartingSerial, "0000000")
         
        Orderqty = Val(Orderqty) - 1
        SeparatorCount = Val(SeparatorCount) + 1
    
     Me.Caption = "Personal.P 10 Digits..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub PrinterFilePA_MC()

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Close #1
Open App.Path & "\MC\PrinterFileMC." & Format(Now, "YY") & "P" For Output As #1

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' and Itemtype = 'MC'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(8)
    
    If Len(DBFRecordset1.Fields(9)) >= 1 Then
        Address1 = DBFRecordset1.Fields(9)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(10)) >= 1 Then
       Address2 = DBFRecordset1.Fields(10)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(11)) >= 1 Then
       Address3 = DBFRecordset1.Fields(11)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    
    If Len(StartingSerial) < 10 Then
        Do Until Len(StartingSerial) = 10
            StartingSerial = "0" & StartingSerial
        Loop
    End If
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        If Len(EndingSerial) < 10 Then
            Do Until Len(EndingSerial) = 10
                EndingSerial = "0" & EndingSerial
            Loop
        End If
        
        NextSerial = Val(StartingSerial) + 50
            
         If Len(NextSerial) < 10 Then
            Do Until Len(NextSerial) = 10
                NextSerial = "0" & NextSerial
            Loop
        End If
        
        If Len(Incrementor) < 6 Then
            Do Until Len(Incrementor) = 6
                Incrementor = "0" & Incrementor
            Loop
        End If

    'PRINTING
    If Counter = 0 Then
        Print #1, SeparatorValue & "  " & Incrementor & "3"
    Else
        Print #1, "3"
    End If
    
    Counter = Val(Counter) + 1
    
        Print #1, "3"
        Print #1, Rtno
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "A"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, "SN"
        Print #1, ""
        Print #1, ""
        Print #1, Acctnm2
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "C"
        Print #1, "XXXX"
        Print #1, "XXXX"
        Print #1,
        Print #1,
        Print #1, BranchName
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address1
        Print #1, Address2
        Print #1, Address2
        Print #1, Address3
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
                
        If SeparatorCount Mod 4 = 0 Then
            Print #1, "\"
            Incrementor = Val(Incrementor) + 1
            Counter = 0
        Else
            SeparatorValue = ""
        End If
        
        StartingSerial = Val(StartingSerial) + 50
            
          If Len(StartingSerial) < 10 Then
          
            Do Until Len(StartingSerial) = 10
                StartingSerial = "0" & StartingSerial
            Loop
         End If
         
        Orderqty = Val(Orderqty) - 1
        SeparatorCount = Val(SeparatorCount) + 1
    
     Me.Caption = "Personal.P..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub PrinterFilePA_GC()

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Close #1
Open App.Path & "\GIFTCHK\PrinterFileGC." & Format(Now, "YY") & "P" For Output As #1

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' and Itemtype = 'GC'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(8)
    
    If Len(DBFRecordset1.Fields(9)) >= 1 Then
        Address1 = DBFRecordset1.Fields(9)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(10)) >= 1 Then
       Address2 = DBFRecordset1.Fields(10)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(11)) >= 1 Then
       Address3 = DBFRecordset1.Fields(11)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    
    If Len(StartingSerial) < 6 Then
        Do Until Len(StartingSerial) = 6
            StartingSerial = "0" & StartingSerial
        Loop
    End If
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        If Len(EndingSerial) < 6 Then
            Do Until Len(EndingSerial) = 6
                EndingSerial = "0" & EndingSerial
            Loop
        End If
        
        NextSerial = Val(StartingSerial) + 50
            
         If Len(NextSerial) < 6 Then
            Do Until Len(NextSerial) = 6
                NextSerial = "0" & NextSerial
            Loop
        End If
        
        If Len(Incrementor) < 6 Then
            Do Until Len(Incrementor) = 6
                Incrementor = "0" & Incrementor
            Loop
        End If

    'PRINTING
    If Counter = 0 Then
        Print #1, SeparatorValue & "  " & Incrementor & "3"
    Else
        Print #1, "3"
    End If
    
    Counter = Val(Counter) + 1
    
        Print #1, "3"
        Print #1, Rtno
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "A"
        Print #1, "  O0000NNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, "  O0000NNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, "SN"
        Print #1, ""
        Print #1, ""
        Print #1, Acctnm2
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "C"
        Print #1, "XXXX"
        Print #1, "XXXX"
        Print #1,
        Print #1,
        Print #1, BranchName
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address1
        Print #1, Address2
        Print #1, Address2
        Print #1, Address3
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
                
        If SeparatorCount Mod 4 = 0 Then
            Print #1, "\"
            Incrementor = Val(Incrementor) + 1
            Counter = 0
        Else
            SeparatorValue = ""
        End If
        
        StartingSerial = Val(StartingSerial) + 50
            
          If Len(StartingSerial) < 6 Then
          
            Do Until Len(StartingSerial) = 6
                StartingSerial = "0" & StartingSerial
            Loop
         End If
         
        Orderqty = Val(Orderqty) - 1
        SeparatorCount = Val(SeparatorCount) + 1
    
     Me.Caption = "Personal.P..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub PrinterFileCA13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

Close #1
Open App.Path & "\12_Digits\PrinterFileC." & Format(Now, "YY") & "P" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSC Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno= '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
       Address1 = DBFRecordset1.Fields(1)
    Else
       Address1 = ""
    End If
   
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
      Address2 = DBFRecordset1.Fields(2)
    Else
      Address2 = ""
    End If
   
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
      Address3 = DBFRecordset1.Fields(3)
    Else
      Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 99
        EndingSerial = Format(EndingSerial, "0000000000")
    
        NextSerial = Val(StartingSerial) + 100
        NextSerial = Format(NextSerial, "0000000000")
       
       If Len(Incrementor) < 6 Then
            Do Until Len(Incrementor) = 6
                Incrementor = "0" & Incrementor
            Loop
        End If
        
    'PRINTING
    If Counter = 0 Then
        Print #1, SeparatorValue & "  " & Incrementor & "3"
    Else
        Print #1, "3"
    End If
    
    Counter = Val(Counter) + 1
    
        Print #1, "3"
        Print #1, Rtno
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Acctno
        Print #1, NextSerial
        Print #1, NextSerial
        Print #1, "A"
        Print #1, "A"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, "SN"
        Print #1, ""
        Print #1, ""
        Print #1, Acctnm2
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "C"
        Print #1, "XXXX"
        Print #1, "XXXX"
        Print #1,
        Print #1,
        Print #1, BranchName
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address1
        Print #1, Address2
        Print #1, Address2
        Print #1, Address3
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, StartingSerial
        Print #1, StartingSerial
        Print #1, EndingSerial
        Print #1, EndingSerial
        
        If SeparatorCount Mod 4 = 0 Then
            Print #1, "\"
            Incrementor = Val(Incrementor) + 1
            Counter = 0
        Else
            SeparatorValue = ""
        End If
        
            StartingSerial = Val(StartingSerial) + 100
            StartingSerial = Format(StartingSerial, "0000000000")
        
        Orderqty = Val(Orderqty) - 1
        SeparatorCount = Val(SeparatorCount) + 1
    
    Me.Caption = "Commercial.P 13 Digits..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub PrinterFileCA10()

'10 DIGITS
PathOfFile = App.Path & "\10_Digits"

Close #1
Open App.Path & "\10_Digits\PrinterFileC." & Format(Now, "YY") & "P" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSC Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno= '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
       Address1 = DBFRecordset1.Fields(1)
    Else
       Address1 = ""
    End If
   
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
      Address2 = DBFRecordset1.Fields(2)
    Else
      Address2 = ""
    End If
   
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
      Address3 = DBFRecordset1.Fields(3)
    Else
      Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        EndingSerial = Format(EndingSerial, "0000000000")
    
        NextSerial = Val(StartingSerial) + 50
        NextSerial = Format(NextSerial, "0000000000")
       
       If Len(Incrementor) < 6 Then
            Do Until Len(Incrementor) = 6
                Incrementor = "0" & Incrementor
            Loop
        End If
        
    'PRINTING
    If Counter = 0 Then
        Print #1, SeparatorValue & "  " & Incrementor & "3"
    Else
        Print #1, "3"
    End If
    
    Counter = Val(Counter) + 1
    
        Print #1, "3"
        Print #1, Rtno
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Acctno
        Print #1, NextSerial
        Print #1, NextSerial
        Print #1, "A"
        Print #1, "A"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, "SN"
        Print #1, ""
        Print #1, ""
        Print #1, Acctnm2
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "C"
        Print #1, "XXXX"
        Print #1, "XXXX"
        Print #1,
        Print #1,
        Print #1, BranchName
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address1
        Print #1, Address2
        Print #1, Address2
        Print #1, Address3
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, StartingSerial
        Print #1, StartingSerial
        Print #1, EndingSerial
        Print #1, EndingSerial
        
        If SeparatorCount Mod 4 = 0 Then
            Print #1, "\"
            Incrementor = Val(Incrementor) + 1
            Counter = 0
        Else
            SeparatorValue = ""
        End If
        
            StartingSerial = Val(StartingSerial) + 50
            StartingSerial = Format(StartingSerial, "0000000000")
        
        Orderqty = Val(Orderqty) - 1
        SeparatorCount = Val(SeparatorCount) + 1
    
    Me.Caption = "Commercial.P 10 Digits..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub FixedPositionPersonal13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

Close #1
Open App.Path & "\12_Digits\FixedPositionP." & Format(Now, "YY") & "F" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File
    
  Do While Not dbfRecordset.EOF
    
    Rtno = dbfRecordset.Fields(0)
    AccountNo = dbfRecordset.Fields(2)
    Books = dbfRecordset.Fields(3)
        Do Until Len(Books) >= 2
            Books = " " & Books
        Loop
    Serial = dbfRecordset.Fields(5)
    Routing1to5 = dbfRecordset.Fields(7)
    Routing6to9 = dbfRecordset.Fields(9)
    AccountnoWH = dbfRecordset.Fields(10)
    Name1 = dbfRecordset.Fields(12)
        Do Until Len(Name1) >= 33
            Name1 = Name1 & " "
        Loop
    Name2 = dbfRecordset.Fields(14)
        Do Until Len(Name2) >= 31
            Name2 = Name2 & " "
        Loop
    SN = dbfRecordset.Fields(15)
    C = dbfRecordset.Fields(19)
    Blank = dbfRecordset.Fields(20)
    
   Print #1, Rtno & " " & AccountNo & "   " & Books & " " & Serial & "    " & Routing1to5 & "  " & Routing6to9 & " " & AccountnoWH & "   " & Name1 & "" & Name2 & SN & "       " & C & Blank
    
    dbfRecordset.MoveNext
    Me.Caption = "Personal Fixed Position 13 Digits..." & Serial
  Loop

Close #1
Me.Caption = "SBTC"
End Sub

Sub FixedPositionPersonal10()

'10 DIGITS
PathOfFile = App.Path & "\10_Digits"

Close #1
Open App.Path & "\10_Digits\FixedPositionP." & Format(Now, "YY") & "F" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File
    
  Do While Not dbfRecordset.EOF
    
    Rtno = dbfRecordset.Fields(0)
    AccountNo = dbfRecordset.Fields(2)
    Books = dbfRecordset.Fields(3)
        Do Until Len(Books) >= 2
            Books = " " & Books
        Loop
    Serial = dbfRecordset.Fields(5)
    Routing1to5 = dbfRecordset.Fields(7)
    Routing6to9 = dbfRecordset.Fields(9)
    AccountnoWH = dbfRecordset.Fields(10)
    Name1 = dbfRecordset.Fields(12)
        Do Until Len(Name1) >= 33
            Name1 = Name1 & " "
        Loop
    Name2 = dbfRecordset.Fields(14)
        Do Until Len(Name2) >= 31
            Name2 = Name2 & " "
        Loop
    SN = dbfRecordset.Fields(15)
    C = dbfRecordset.Fields(19)
    Blank = dbfRecordset.Fields(20)
    
   Print #1, Rtno & " " & AccountNo & "   " & Books & " " & Serial & "    " & Routing1to5 & "  " & Routing6to9 & " " & AccountnoWH & "   " & Name1 & "" & Name2 & SN & "       " & C & Blank
    
    dbfRecordset.MoveNext
    Me.Caption = "Personal Fixed Position 10 Digits..." & Serial
  Loop

Close #1
Me.Caption = "SBTC"
End Sub

Sub FixedPositionPersonal_MC()

Close #1
Open App.Path & "\MC\FixedPositionMC." & Format(Now, "YY") & "F" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File
    
  Do While Not dbfRecordset.EOF
    
    Rtno = dbfRecordset.Fields(0)
    AccountNo = dbfRecordset.Fields(2)
    Books = dbfRecordset.Fields(3)
        Do Until Len(Books) >= 2
            Books = " " & Books
        Loop
    Serial = dbfRecordset.Fields(5)
    Routing1to5 = dbfRecordset.Fields(7)
    Routing6to9 = dbfRecordset.Fields(9)
    AccountnoWH = dbfRecordset.Fields(10)
    Name1 = dbfRecordset.Fields(12)
        Do Until Len(Name1) >= 33
            Name1 = Name1 & " "
        Loop
    Name2 = dbfRecordset.Fields(14)
        Do Until Len(Name2) >= 31
            Name2 = Name2 & " "
        Loop
    SN = dbfRecordset.Fields(15)
    C = dbfRecordset.Fields(19)
    Blank = dbfRecordset.Fields(20)
    
   Print #1, Rtno & " " & AccountNo & "   " & Books & " " & Serial & " " & Routing1to5 & "  " & Routing6to9 & " " & AccountnoWH & "   " & Name1 & "" & Name2 & SN & "       " & C & Blank
    
    dbfRecordset.MoveNext
    Me.Caption = "Personal Fixed Position... " & Serial
  Loop

Close #1
Me.Caption = "SBTC"
End Sub

Sub FixedPositionPersonal_GC()

Close #1
Open App.Path & "\GIFTCHK\FixedPositionGC." & Format(Now, "YY") & "F" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File
    
  Do While Not dbfRecordset.EOF
    
    Rtno = dbfRecordset.Fields(0)
    AccountNo = dbfRecordset.Fields(2)
    Books = dbfRecordset.Fields(3)
        Do Until Len(Books) >= 2
            Books = " " & Books
        Loop
    Serial = dbfRecordset.Fields(5)
    Routing1to5 = dbfRecordset.Fields(7)
    Routing6to9 = dbfRecordset.Fields(9)
    AccountnoWH = dbfRecordset.Fields(10)
    Name1 = dbfRecordset.Fields(12)
        Do Until Len(Name1) >= 33
            Name1 = Name1 & " "
        Loop
    Name2 = dbfRecordset.Fields(14)
        Do Until Len(Name2) >= 31
            Name2 = Name2 & " "
        Loop
    SN = dbfRecordset.Fields(15)
    C = dbfRecordset.Fields(19)
    Blank = dbfRecordset.Fields(20)
    
   Print #1, Rtno & " " & AccountNo & "   " & Books & " " & Serial & "     " & Routing1to5 & "  " & Routing6to9 & " " & AccountnoWH & "   " & Name1 & "" & Name2 & SN & "       " & C & Blank
    
    dbfRecordset.MoveNext
    Me.Caption = "Personal Fixed Position... " & Serial
  Loop

Close #1
Me.Caption = "SBTC"
End Sub

Sub FixedPositionCommercial13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

Close #1
Open App.Path & "\12_Digits\FixedPositionC." & Format(Now, "YY") & "F" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSC"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File
    
  Do While Not dbfRecordset.EOF
    
    Rtno = dbfRecordset.Fields(0)
    AccountNo = dbfRecordset.Fields(2)
    Books = dbfRecordset.Fields(3)
        Do Until Len(Books) >= 2
            Books = " " & Books
        Loop
    Serial = dbfRecordset.Fields(5)
    Routing1to5 = dbfRecordset.Fields(7)
    Routing6to9 = dbfRecordset.Fields(9)
    AccountnoWH = dbfRecordset.Fields(10)
    Name1 = dbfRecordset.Fields(12)
        Do Until Len(Name1) >= 33
            Name1 = Name1 & " "
        Loop
    Name2 = dbfRecordset.Fields(14)
        Do Until Len(Name2) >= 31
            Name2 = Name2 & " "
        Loop
    SN = dbfRecordset.Fields(15)
    C = dbfRecordset.Fields(19)
    Blank = dbfRecordset.Fields(20)
    
   Print #1, Rtno & " " & AccountNo & "   " & Books & " " & Serial & " " & Routing1to5 & "  " & Routing6to9 & " " & AccountnoWH & "   " & Name1 & "" & Name2 & SN & "       " & C & Blank
    
    dbfRecordset.MoveNext
    Me.Caption = "Commercial Fixed Position 13 Digits... " & Serial
  Loop

Close #1
Me.Caption = "SBTC"
End Sub

Sub FixedPositionCommercial10()

'10 DIGITS
PathOfFile = App.Path & "\10_Digits"

Close #1
Open App.Path & "\10_Digits\FixedPositionC." & Format(Now, "YY") & "F" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSC"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File
    
  Do While Not dbfRecordset.EOF
    
    Rtno = dbfRecordset.Fields(0)
    AccountNo = dbfRecordset.Fields(2)
    Books = dbfRecordset.Fields(3)
        Do Until Len(Books) >= 2
            Books = " " & Books
        Loop
    Serial = dbfRecordset.Fields(5)
    Routing1to5 = dbfRecordset.Fields(7)
    Routing6to9 = dbfRecordset.Fields(9)
    AccountnoWH = dbfRecordset.Fields(10)
    Name1 = dbfRecordset.Fields(12)
        Do Until Len(Name1) >= 33
            Name1 = Name1 & " "
        Loop
    Name2 = dbfRecordset.Fields(14)
        Do Until Len(Name2) >= 31
            Name2 = Name2 & " "
        Loop
    SN = dbfRecordset.Fields(15)
    C = dbfRecordset.Fields(19)
    Blank = dbfRecordset.Fields(20)
    
   Print #1, Rtno & " " & AccountNo & "   " & Books & " " & Serial & " " & Routing1to5 & "  " & Routing6to9 & " " & AccountnoWH & "   " & Name1 & "" & Name2 & SN & "       " & C & Blank
    
    dbfRecordset.MoveNext
    Me.Caption = "Commercial Fixed Position 10 Digits... " & Serial
  Loop

Close #1
Me.Caption = "SBTC"
End Sub

Sub PrinterTXT_PA13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

Close #1
Open App.Path & "\12_Digits\PrinterFileP.TXT" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
        Address1 = DBFRecordset1.Fields(1)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
       Address2 = DBFRecordset1.Fields(2)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
       Address3 = DBFRecordset1.Fields(3)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        EndingSerial = Format(EndingSerial, "0000000")

        NextSerial = Val(StartingSerial) + 50
        NextSerial = Format(NextSerial, "0000000")
        
        If Len(Incrementor) < 6 Then
            Incrementor = Format(Incrementor, "000000")
        End If
        
    'PRINTING
        Print #1, "3"
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "     ONNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, ""
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "XXXX"
        Print #1,
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address2
        Print #1, Address3
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        
        Counter = Val(Counter) + 1
        
        StartingSerial = Val(StartingSerial) + 50
        StartingSerial = Format(StartingSerial, "0000000")
         
        Orderqty = Val(Orderqty) - 1
    
     Me.Caption = "Personal.txt 13 Digits..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Print #1, "\"
Close #1
End Sub

Sub PrinterTXT_PA10()

'10 DIGITS
PathOfFile = App.Path & "\10_Digits"

Close #1
Open App.Path & "\10_Digits\PrinterFileP.TXT" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno = '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
        Address1 = DBFRecordset1.Fields(1)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
       Address2 = DBFRecordset1.Fields(2)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
       Address3 = DBFRecordset1.Fields(3)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        EndingSerial = Format(EndingSerial, "0000000")

        NextSerial = Val(StartingSerial) + 50
        NextSerial = Format(NextSerial, "0000000")
        
        If Len(Incrementor) < 6 Then
            Incrementor = Format(Incrementor, "000000")
        End If
        
    'PRINTING
        Print #1, "3"
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "     ONNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, ""
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "XXXX"
        Print #1,
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address2
        Print #1, Address3
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        
        Counter = Val(Counter) + 1
        
        StartingSerial = Val(StartingSerial) + 50
        StartingSerial = Format(StartingSerial, "0000000")
         
        Orderqty = Val(Orderqty) - 1
    
     
    Loop
    
    Me.Caption = "Personal.txt 10 Digits..." & StartingSerial & "    " & FormatNumber((LoopCount / dbfRecordset.RecordCount) * 100, 5)
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Print #1, "\"
Close #1
End Sub

Sub PrinterTXT_MC()

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Close #1
Open App.Path & "\MC\PrinterFilePA.TXT" For Output As #1

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' and Itemtype = 'MC'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(8)
    
    If Len(DBFRecordset1.Fields(9)) >= 1 Then
        Address1 = DBFRecordset1.Fields(9)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(10)) >= 1 Then
       Address2 = DBFRecordset1.Fields(10)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(11)) >= 1 Then
       Address3 = DBFRecordset1.Fields(11)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    
    If Len(StartingSerial) < 10 Then
        Do Until Len(StartingSerial) = 10
            StartingSerial = "0" & StartingSerial
        Loop
    End If
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        If Len(EndingSerial) < 10 Then
            Do Until Len(EndingSerial) = 10
                EndingSerial = "0" & EndingSerial
            Loop
        End If
        
        NextSerial = Val(StartingSerial) + 50
            
         If Len(NextSerial) < 10 Then
            Do Until Len(NextSerial) = 10
                NextSerial = "0" & NextSerial
            Loop
        End If
        
        If Len(Incrementor) < 6 Then
            Do Until Len(Incrementor) = 6
                Incrementor = "0" & Incrementor
            Loop
        End If
        
    'PRINTING
        Print #1, "3"
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, ""
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "XXXX"
        Print #1,
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address2
        Print #1, Address3
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        
        Counter = Counter + 1
        
        StartingSerial = Val(StartingSerial) + 50
            
          If Len(StartingSerial) < 10 Then
          
            Do Until Len(StartingSerial) = 10
                StartingSerial = "0" & StartingSerial
            Loop
         End If
         
        Orderqty = Val(Orderqty) - 1
    
     Me.Caption = "Personal.TXT..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub PrinterTXT_GC()

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSP Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Close #1
Open App.Path & "\GIFTCHK\PrinterFilePA.TXT" For Output As #1

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Rtno & "' and Itemtype = 'GC'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(8)
    
    If Len(DBFRecordset1.Fields(9)) >= 1 Then
        Address1 = DBFRecordset1.Fields(9)
    Else
        Address1 = ""
    End If
    
    If Len(DBFRecordset1.Fields(10)) >= 1 Then
       Address2 = DBFRecordset1.Fields(10)
    Else
       Address2 = ""
    End If
    
    If Len(DBFRecordset1.Fields(11)) >= 1 Then
       Address3 = DBFRecordset1.Fields(11)
    Else
       Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    
    If Len(StartingSerial) < 6 Then
        Do Until Len(StartingSerial) = 6
            StartingSerial = "0" & StartingSerial
        Loop
    End If
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        If Len(EndingSerial) < 6 Then
            Do Until Len(EndingSerial) = 6
                EndingSerial = "0" & EndingSerial
            Loop
        End If
        
        NextSerial = Val(StartingSerial) + 50
            
         If Len(NextSerial) < 6 Then
            Do Until Len(NextSerial) = 6
                NextSerial = "0" & NextSerial
            Loop
        End If
        
        If Len(Incrementor) < 6 Then
            Do Until Len(Incrementor) = 6
                Incrementor = "0" & Incrementor
            Loop
        End If
        
    'PRINTING
        Print #1, "3"
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "  O0000NNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, ""
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "XXXX"
        Print #1,
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address2
        Print #1, Address3
        Print #1, ""
        Print #1, ""
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        
        Counter = Counter + 1
        
        StartingSerial = Val(StartingSerial) + 50
            
          If Len(StartingSerial) < 6 Then
          
            Do Until Len(StartingSerial) = 6
                StartingSerial = "0" & StartingSerial
            Loop
         End If
         
        Orderqty = Val(Orderqty) - 1
    
     Me.Caption = "Personal.TXT..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"

Print #1, "\"
Close #1
End Sub

Sub PrinterTXT_CA13()

'13 DIGITS
PathOfFile = App.Path & "\12_Digits"

Close #1
Open App.Path & "\12_Digits\PrinterFileC.TXT" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSC Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno= '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
       Address1 = DBFRecordset1.Fields(1)
    Else
       Address1 = ""
    End If
   
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
      Address2 = DBFRecordset1.Fields(2)
    Else
      Address2 = ""
    End If
   
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
      Address3 = DBFRecordset1.Fields(3)
    Else
      Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 99
        EndingSerial = Format(EndingSerial, "0000000000")
    
        NextSerial = Val(StartingSerial) + 100
        NextSerial = Format(NextSerial, "0000000000")
       
        If Len(Incrementor) < 6 Then
             Incrementor = Format(Incrementor, "0000000")
        End If
         
        'PRINTING
        Print #1, "3"
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, ""
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "XXXX"
        Print #1,
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address2
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        
        Counter1 = Val(Counter1) + 1
        
        StartingSerial = Val(StartingSerial) + 100
        StartingSerial = Format(StartingSerial, "0000000000")

        Orderqty = Val(Orderqty) - 1
    
    Me.Caption = "Commercial.txt 13 Digits..." & StartingSerial
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Print #1, "\"
Close #1
End Sub

Sub PrinterTXT_CA10()

'10 DIGITS
PathOfFile = App.Path & "\10_Digits"

Close #1
Open App.Path & "\10_Digits\PrinterFileC.TXT" For Output As #1

'Read from DBF File
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & PathOfFile & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TRANSC Order By Rt_no, Acct_no, Ck_no_p"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Read from DBF File

'Create Text File .10P
LoopCount = 0
SeparatorCount = 1
Counter = 0
Incrementor = 0

Do Until LoopCount = dbfRecordset.RecordCount
    'Field Names
    Rtno = dbfRecordset.Fields(0)
    Acctno = dbfRecordset.Fields(2)
    AcctNoWithHyphen = dbfRecordset.Fields(10)
    
    If Len(dbfRecordset.Fields(12)) >= 1 Then
        Acctnm1 = dbfRecordset.Fields(12)
    Else
        Acctnm1 = ""
    End If
    
    If Len(dbfRecordset.Fields(14)) >= 1 Then
        Acctnm2 = dbfRecordset.Fields(14)
    Else
        Acctnm2 = ""
    End If
    
    Orderqty = dbfRecordset.Fields(3)
    
    'Read from BRANCHES to get addresses
    Set DBFConnector1 = CreateObject("ADODB.Connection")
 
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT BranchName, Address1, Address2, Address3 FROM SBTCNEW WHERE Rtno= '" & Rtno & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
    'End Read from BRANCHES to get addresses
    
    BranchName = DBFRecordset1.Fields(0)
    
    If Len(DBFRecordset1.Fields(1)) >= 1 Then
       Address1 = DBFRecordset1.Fields(1)
    Else
       Address1 = ""
    End If
   
    If Len(DBFRecordset1.Fields(2)) >= 1 Then
      Address2 = DBFRecordset1.Fields(2)
    Else
      Address2 = ""
    End If
   
    If Len(DBFRecordset1.Fields(3)) >= 1 Then
      Address3 = DBFRecordset1.Fields(3)
    Else
      Address3 = ""
    End If
       
    StartingSerial = dbfRecordset.Fields(5)
    StartingSerial = Format(StartingSerial, "0000000000")
            
    Do Until Val(Orderqty) = 0
        EndingSerial = Val(StartingSerial) + 49
        EndingSerial = Format(EndingSerial, "0000000000")
    
        NextSerial = Val(StartingSerial) + 50
        NextSerial = Format(NextSerial, "0000000000")
       
        If Len(Incrementor) < 6 Then
             Incrementor = Format(Incrementor, "0000000")
        End If
             
        'PRINTING
        Print #1, "3"
        Print #1, Rtno
        Print #1, Acctno
        Print #1, Replace(NextSerial, " ", "")
        Print #1, "A"
        Print #1, "  ONNNNNNNNNNO" & Mid(Rtno, 1, 5) & "D" & Mid(Rtno, 6, 4) & "T" & Acctno & "O"
        Print #1, Mid(Rtno, 1, 5)
        Print #1, " " & Mid(Rtno, 6, 4)
        Print #1, AcctNoWithHyphen
        Print #1, Acctnm1
        Print #1, "SN"
        Print #1, ""
        Print #1, Acctnm2
        Print #1, "C"
        Print #1, "XXXX"
        Print #1,
        Print #1, BranchName
        Print #1, Address1
        Print #1, Address2
        Print #1, Address3
        Print #1,
        Print #1,
        Print #1, "SECURITY BANK"
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Replace(StartingSerial, " ", "")
        Print #1, Replace(EndingSerial, " ", "")
        
        
        Counter1 = Val(Counter1) + 1
        
        
        StartingSerial = Val(StartingSerial) + 50
        StartingSerial = Format(StartingSerial, "0000000000")

        Orderqty = Val(Orderqty) - 1
    
        Me.Caption = "Commercial.txt 10 Digits..." & StartingSerial
    
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = "SBTC"
Print #1, "\"
Close #1
End Sub

Sub UpdateMastertoREF_PA13()

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' And Itemtype = 'PA' AND Digits = '13' Order by Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do While Not dbfRecordset.EOF

Routing = dbfRecordset.Fields(0)
BooksTotal = 0
TotalCount = 0

Set DBFConnector0 = CreateObject("ADODB.Connection")
    
    DBFConnector0.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector0.CursorLocation = adUseClient
       
    Set DBFRecordset0 = CreateObject("ADODB.Recordset")
    SQL0 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Routing & "' AND Chktype = 'A' And Itemtype ='PA' AND Digits = '13'"
    DBFRecordset0.Open SQL0, DBFConnector0, 1, 1
    
  Do While Not DBFRecordset0.EOF
     Books = Val(DBFRecordset0.Fields(5))

     BooksTotal = Val(BooksTotal) + Books
     DBFRecordset0.MoveNext
  Loop
     
    Set DBFConnector1 = CreateObject("ADODB.Connection")
     
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'A'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
            
    BatchDate = Format(Now, "mm/dd/yy")
    Rtno = DBFRecordset1.Fields(1)
    ChkType = "A"
    Pbefore = DBFRecordset1.Fields(4)
    Cbefore = DBFRecordset1.Fields(5)
    LASTNO = DBFRecordset1.Fields(6)
    BranchName = DBFRecordset1.Fields(9)
    
    TotalCount = Val(BooksTotal) * 50
    
    Lastnumber = Val(LASTNO) + Val(TotalCount)
        
        'DELETE OLD RECORD
        Set DBFConnector2 = CreateObject("ADODB.Connection")
         
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "DELETE * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'A'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        'INSERT NEW RECORD
        Set DBFConnector3 = CreateObject("ADODB.Connection")
         
        DBFConnector3.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector3.CursorLocation = adUseClient
            
        Set DBFRecordset3 = CreateObject("ADODB.Recordset")
        SQL3 = "INSERT INTO REF ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset3.Open SQL3, DBFConnector3, 1, 1
     
        'INSERT TO MASTER
        Set DBFConnector4 = CreateObject("ADODB.Connection")
         
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO MASTER ([Date1],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
     
     Me.Caption = "REF and MASTER..." & Routing
     dbfRecordset.MoveNext
Loop

Me.Caption = "SBTC"

Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack REF.dbf"
cmd1.Execute
conn1.Close

Set conn2 = New ADODB.Connection
conn2.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd2 = New ADODB.Command
cmd2.CommandType = adCmdText
Set cmd2.ActiveConnection = conn2
cmd2.CommandText = "Set Exclusive On"
cmd2.Execute
cmd2.CommandText = "Pack MASTER.dbf"
cmd2.Execute
conn2.Close
End Sub

Sub UpdateMastertoREF_PA10()

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' And Itemtype = 'PA' AND Digits = '10' Order by Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do While Not dbfRecordset.EOF

Routing = dbfRecordset.Fields(0)
BooksTotal = 0
TotalCount = 0

Set DBFConnector0 = CreateObject("ADODB.Connection")
    
    DBFConnector0.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector0.CursorLocation = adUseClient
       
    Set DBFRecordset0 = CreateObject("ADODB.Recordset")
    SQL0 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Routing & "' AND Chktype = 'A' And Itemtype ='PA' AND Digits = '10'"
    DBFRecordset0.Open SQL0, DBFConnector0, 1, 1
    
  Do While Not DBFRecordset0.EOF
     Books = Val(DBFRecordset0.Fields(5))

     BooksTotal = Val(BooksTotal) + Books
     DBFRecordset0.MoveNext
  Loop
     
    Set DBFConnector1 = CreateObject("ADODB.Connection")
     
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'A'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
            
    BatchDate = Format(Now, "mm/dd/yy")
    Rtno = DBFRecordset1.Fields(1)
    ChkType = "A"
    Pbefore = DBFRecordset1.Fields(4)
    Cbefore = DBFRecordset1.Fields(5)
    LASTNO = DBFRecordset1.Fields(6)
    BranchName = DBFRecordset1.Fields(9)
    
    TotalCount = Val(BooksTotal) * 50
    
    Lastnumber = Val(LASTNO) + Val(TotalCount)
        
        'DELETE OLD RECORD
        Set DBFConnector2 = CreateObject("ADODB.Connection")
         
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "DELETE * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'A'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        'INSERT NEW RECORD
        Set DBFConnector3 = CreateObject("ADODB.Connection")
         
        DBFConnector3.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector3.CursorLocation = adUseClient
            
        Set DBFRecordset3 = CreateObject("ADODB.Recordset")
        SQL3 = "INSERT INTO REF ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset3.Open SQL3, DBFConnector3, 1, 1
     
        'INSERT TO MASTER
        Set DBFConnector4 = CreateObject("ADODB.Connection")
         
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO MASTER ([Date1],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
     
     Me.Caption = "REF and MASTER 10 Digits..." & Routing
     dbfRecordset.MoveNext
Loop

Me.Caption = "SBTC"

Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack REF.dbf"
cmd1.Execute
conn1.Close

Set conn2 = New ADODB.Connection
conn2.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd2 = New ADODB.Command
cmd2.CommandType = adCmdText
Set cmd2.ActiveConnection = conn2
cmd2.CommandText = "Set Exclusive On"
cmd2.Execute
cmd2.CommandText = "Pack MASTER.dbf"
cmd2.Execute
conn2.Close
End Sub

Sub UpdateMastertoREF_CA13()

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'B' AND ItemType = 'CA' AND Digits = '13'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do While Not dbfRecordset.EOF

Routing = dbfRecordset.Fields(0)
BooksTotal = 0
TotalCount = 0

Set DBFConnector0 = CreateObject("ADODB.Connection")
    
    DBFConnector0.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector0.CursorLocation = adUseClient
       
    Set DBFRecordset0 = CreateObject("ADODB.Recordset")
    SQL0 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Routing & "' AND Chktype = 'B' AND Itemtype = 'CA' AND Digits = '13'"
    DBFRecordset0.Open SQL0, DBFConnector0, 1, 1
    
  Do While Not DBFRecordset0.EOF
     Books = Val(DBFRecordset0.Fields(5))
     
     BooksTotal = Val(BooksTotal) + Books
     DBFRecordset0.MoveNext
  Loop
     
    Set DBFConnector1 = CreateObject("ADODB.Connection")
     
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'B'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
            
    BatchDate = Format(Now, "mm/dd/yy")
    Rtno = DBFRecordset1.Fields(1)
    ChkType = "B"
    Pbefore = DBFRecordset1.Fields(4)
    Cbefore = DBFRecordset1.Fields(5)
    LASTNO = DBFRecordset1.Fields(6)
    BranchName = DBFRecordset1.Fields(9)
    
    TotalCount = Val(BooksTotal) * 100
    
    Lastnumber = Val(LASTNO) + Val(TotalCount)
        
        'DELETE OLD
        Set DBFConnector2 = CreateObject("ADODB.Connection")
         
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "DELETE * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'B'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
     
        'INSERT TO REF
        Set DBFConnector3 = CreateObject("ADODB.Connection")
         
        DBFConnector3.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector3.CursorLocation = adUseClient
            
        Set DBFRecordset3 = CreateObject("ADODB.Recordset")
        SQL3 = "INSERT INTO REF ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset3.Open SQL3, DBFConnector3, 1, 1
     
        'INSERT TO MASTER
        Set DBFConnector4 = CreateObject("ADODB.Connection")
         
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO MASTER ([Date1],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
     
     Me.Caption = "REF and MASTER 13 Digits..." & Routing
     
     dbfRecordset.MoveNext
Loop

Me.Caption = "SBTC"

Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack REF.dbf"
cmd1.Execute
conn1.Close

Set conn2 = New ADODB.Connection
conn2.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd2 = New ADODB.Command
cmd2.CommandType = adCmdText
Set cmd2.ActiveConnection = conn2
cmd2.CommandText = "Set Exclusive On"
cmd2.Execute
cmd2.CommandText = "Pack MASTER.dbf"
cmd2.Execute
conn2.Close
End Sub

Sub UpdateMastertoREF_CA10()

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'B' AND ItemType = 'CA' AND Digits = '10'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do While Not dbfRecordset.EOF

Routing = dbfRecordset.Fields(0)
BooksTotal = 0
TotalCount = 0

Set DBFConnector0 = CreateObject("ADODB.Connection")
    
    DBFConnector0.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector0.CursorLocation = adUseClient
       
    Set DBFRecordset0 = CreateObject("ADODB.Recordset")
    SQL0 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Routing & "' AND Chktype = 'B' AND Itemtype = 'CA' AND Digits = '10'"
    DBFRecordset0.Open SQL0, DBFConnector0, 1, 1
    
  Do While Not DBFRecordset0.EOF
     Books = Val(DBFRecordset0.Fields(5))
     
     BooksTotal = Val(BooksTotal) + Books
     DBFRecordset0.MoveNext
  Loop
     
    Set DBFConnector1 = CreateObject("ADODB.Connection")
     
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'B'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
            
    BatchDate = Format(Now, "mm/dd/yy")
    Rtno = DBFRecordset1.Fields(1)
    ChkType = "B"
    Pbefore = DBFRecordset1.Fields(4)
    Cbefore = DBFRecordset1.Fields(5)
    LASTNO = DBFRecordset1.Fields(6)
    BranchName = DBFRecordset1.Fields(9)
    
    TotalCount = Val(BooksTotal) * 50
    
    Lastnumber = Val(LASTNO) + Val(TotalCount)
        
        'DELETE OLD
        Set DBFConnector2 = CreateObject("ADODB.Connection")
         
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "DELETE * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'B'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1
     
        'INSERT TO REF
        Set DBFConnector3 = CreateObject("ADODB.Connection")
         
        DBFConnector3.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector3.CursorLocation = adUseClient
            
        Set DBFRecordset3 = CreateObject("ADODB.Recordset")
        SQL3 = "INSERT INTO REF ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset3.Open SQL3, DBFConnector3, 1, 1
     
        'INSERT TO MASTER
        Set DBFConnector4 = CreateObject("ADODB.Connection")
         
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & ";Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO MASTER ([Date1],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
     
     Me.Caption = "REF and MASTER 10 Digits..." & Routing
     
     dbfRecordset.MoveNext
Loop

Me.Caption = "SBTC"

Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack REF.dbf"
cmd1.Execute
conn1.Close

Set conn2 = New ADODB.Connection
conn2.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & ";"

Set cmd2 = New ADODB.Command
cmd2.CommandType = adCmdText
Set cmd2.ActiveConnection = conn2
cmd2.CommandText = "Set Exclusive On"
cmd2.Execute
cmd2.CommandText = "Pack MASTER.dbf"
cmd2.Execute
conn2.Close
End Sub




Sub UpdateMastertoREF_MC()
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'MC' Order by Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do While Not dbfRecordset.EOF

Routing = dbfRecordset.Fields(0)
BooksTotal = 0
TotalCount = 0

Set DBFConnector0 = CreateObject("ADODB.Connection")
    
    DBFConnector0.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector0.CursorLocation = adUseClient
       
    Set DBFRecordset0 = CreateObject("ADODB.Recordset")
    SQL0 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Routing & "' AND Chktype = 'A' AND Itemtype = 'MC'"
    DBFRecordset0.Open SQL0, DBFConnector0, 1, 1
    
  Do While Not DBFRecordset0.EOF
     Books = Val(DBFRecordset0.Fields(5))

     BooksTotal = Val(BooksTotal) + Books
     DBFRecordset0.MoveNext
  Loop
     
    Set DBFConnector1 = CreateObject("ADODB.Connection")
     
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * From REF WHERE Rtno = '" & Routing & "'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
            
    BatchDate = Format(Now, "mm/dd/yy")
    Rtno = DBFRecordset1.Fields(1)
    ChkType = "A"
    Pbefore = DBFRecordset1.Fields(4)
    Cbefore = DBFRecordset1.Fields(5)
    LASTNO = DBFRecordset1.Fields(6)
    BranchName = DBFRecordset1.Fields(9)
    
    TotalCount = Val(BooksTotal) * 50
    
    Lastnumber = Val(LASTNO) + Val(TotalCount)
        
        'DELETE OLD RECORD
        Set DBFConnector2 = CreateObject("ADODB.Connection")
         
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "DELETE * From REF WHERE Rtno = '" & Routing & "'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        'INSERT NEW RECORD
        Set DBFConnector3 = CreateObject("ADODB.Connection")
         
        DBFConnector3.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
        DBFConnector3.CursorLocation = adUseClient
            
        Set DBFRecordset3 = CreateObject("ADODB.Recordset")
        SQL3 = "INSERT INTO REF ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset3.Open SQL3, DBFConnector3, 1, 1
     
        'INSERT TO MASTER
        Set DBFConnector4 = CreateObject("ADODB.Connection")
         
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\MC;Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO MASTER ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
     
     Me.Caption = "REF and MASTER..." & Routing
     dbfRecordset.MoveNext
Loop

Me.Caption = "SBTC"

Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\MC;"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack REF"
cmd1.Execute
conn1.Close

Set conn2 = New ADODB.Connection
conn2.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\MC;"

Set cmd2 = New ADODB.Command
cmd2.CommandType = adCmdText
Set cmd2.ActiveConnection = conn2
cmd2.CommandText = "Set Exclusive On"
cmd2.Execute
cmd2.CommandText = "Pack MASTER.dbf"
cmd2.Execute
conn2.Close
End Sub

Sub UpdateMastertoREF_GC()

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Distinct(Rtno) FROM SBTCNEW WHERE Chktype = 'A' AND Itemtype = 'GC' Order by Rtno"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Do While Not dbfRecordset.EOF

Routing = dbfRecordset.Fields(0)
BooksTotal = 0
TotalCount = 0

Set DBFConnector0 = CreateObject("ADODB.Connection")
    
    DBFConnector0.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector0.CursorLocation = adUseClient
       
    Set DBFRecordset0 = CreateObject("ADODB.Recordset")
    SQL0 = "SELECT * FROM SBTCNEW WHERE Rtno = '" & Routing & "' AND Chktype = 'A' AND Itemtype = 'GC'"
    DBFRecordset0.Open SQL0, DBFConnector0, 1, 1
    
  Do While Not DBFRecordset0.EOF
     Books = Val(DBFRecordset0.Fields(5))

     BooksTotal = Val(BooksTotal) + Books
     DBFRecordset0.MoveNext
  Loop
     
    Set DBFConnector1 = CreateObject("ADODB.Connection")
     
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
        
    Set DBFRecordset1 = CreateObject("ADODB.Recordset")
    SQL1 = "SELECT * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'A'"
    DBFRecordset1.Open SQL1, DBFConnector1, 1, 1
            
    BatchDate = Format(Now, "mm/dd/yy")
    Rtno = DBFRecordset1.Fields(1)
    ChkType = "A"
    Pbefore = DBFRecordset1.Fields(4)
    Cbefore = DBFRecordset1.Fields(5)
    LASTNO = DBFRecordset1.Fields(6)
    BranchName = DBFRecordset1.Fields(9)
    
    TotalCount = Val(BooksTotal) * 50
    
    Lastnumber = Val(LASTNO) + Val(TotalCount)
        
        'DELETE OLD RECORD
        Set DBFConnector2 = CreateObject("ADODB.Connection")
         
        DBFConnector2.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
        DBFConnector2.CursorLocation = adUseClient
            
        Set dbfRecordset2 = CreateObject("ADODB.Recordset")
        SQL2 = "DELETE * From REF WHERE Rtno = '" & Routing & "' AND Chktype = 'A'"
        dbfRecordset2.Open SQL2, DBFConnector2, 1, 1

        'INSERT NEW RECORD
        Set DBFConnector3 = CreateObject("ADODB.Connection")
         
        DBFConnector3.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
        DBFConnector3.CursorLocation = adUseClient
            
        Set DBFRecordset3 = CreateObject("ADODB.Recordset")
        SQL3 = "INSERT INTO REF ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset3.Open SQL3, DBFConnector3, 1, 1
     
        'INSERT TO MASTER
        Set DBFConnector4 = CreateObject("ADODB.Connection")
         
        DBFConnector4.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\GIFTCHK;Extended properties=dBase III"
        DBFConnector4.CursorLocation = adUseClient
            
        Set DBFRecordset4 = CreateObject("ADODB.Recordset")
        SQL4 = "INSERT INTO MASTER ([Date],Rtno,Chktype,[P_before],[C_before],Lastno,[Branch_tex]) VALUES ('" & BatchDate & "','" & Rtno & "','" & ChkType & "','" & Cbefore & "','" & LASTNO & "','" & Lastnumber & "','" & BranchName & "')"
        DBFRecordset4.Open SQL4, DBFConnector4, 1, 1
     
     Me.Caption = "REF and MASTER..." & Routing
     dbfRecordset.MoveNext
Loop

Me.Caption = "SBTC"

Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\GIFTCHK\;"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack REF.dbf"
cmd1.Execute
conn1.Close

Set conn2 = New ADODB.Connection
conn2.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\GIFTCHK\;"

Set cmd2 = New ADODB.Command
cmd2.CommandType = adCmdText
Set cmd2.ActiveConnection = conn2
cmd2.CommandText = "Set Exclusive On"
cmd2.Execute
cmd2.CommandText = "Pack MASTER.dbf"
cmd2.Execute
conn2.Close
End Sub

Sub CheckUpdatedCheckDat()
Dim fso As New FileSystemObject
Dim Address1(0 To 10), Address2(0 To 10), Address3(0 To 10), Address4(0 To 10), Address5(0 To 10), Address6(0 To 10), AddressDescription(0 To 10) As String

'Copy First the MDB
If Dir$("C:\CheckDat_Y.mdb") <> "" Then Kill "C:\CheckDat_Y.mdb"
If fso.FileExists("\\Jenny\Y_CheckDat\Checkdat.mdb") = True Then
    fso.CopyFile "\\Jenny\Y_CheckDat\Checkdat.mdb", "C:\CheckDat_Y.mdb"
End If

If Dir$("C:\CheckDat_G.mdb") <> "" Then Kill "C:\CheckDat_G.mdb"
If fso.FileExists("\\Modem-Computer\G_CheckDat\Checkdat.mdb") = True Then
    fso.CopyFile "\\Modem-Computer\G_CheckDat\Checkdat.mdb", "C:\CheckDat_G.mdb"
End If

If Dir$("C:\CheckDat_K.mdb") <> "" Then Kill "C:\CheckDat_K.mdb"
If fso.FileExists("\\192.168.0.29\K_CheckDat\CheckDat.mdb") = True Then
    fso.CopyFile "\\192.168.0.29\K_CheckDat\CheckDat.mdb", "C:\CheckDat_K.mdb"
End If

If Dir$("C:\CheckDat_Q.mdb") <> "" Then Kill "C:\CheckDat_Q.mdb"
If fso.FileExists("\\Lenovo_xp\Q_CheckDat\Checkdat.mdb") = True Then
    fso.CopyFile "\\Lenovo_xp\Q_CheckDat\Checkdat.mdb", "C:\CheckDat_Q.mdb"
End If

If Dir$("C:\CheckDat_T.mdb") <> "" Then Kill "C:\CheckDat_T.mdb"
If fso.FileExists("\\Kapamilya\T_CheckDat\Checkdat.mdb") = True Then
    fso.CopyFile "\\Kapamilya\T_CheckDat\Checkdat.mdb", "C:\CheckDat_T.mdb"
End If

If Dir$("C:\CheckDat_Z.mdb") <> "" Then Kill "C:\CheckDat_Z.mdb"
If fso.FileExists("\\Emily\Z_CheckDat\CheckDat.mdb") = True Then
    fso.CopyFile "\\Emily\Z_CheckDat\CheckDat.mdb", "C:\CheckDat_Z.mdb"
End If
'End Copy First the MDB

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Rtno) FROM SBTCNEW"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    
    AddressCount = 0
    LoopCount1 = 0
    Do Until LoopCount1 = 7
        If LoopCount1 = 0 Then
            Location = "\\192.168.0.29\CAPTIVE\AUTO\SBTC\CheckDat.mdb"
            Description = "Checkdat.mdb"
        End If
        
        If LoopCount1 = 1 Then
            Location = "C:\Checkdat_Y.mdb"
            Description = "CheckDat.mdb on Drive Y"
        End If
        
        If LoopCount1 = 2 Then
            Location = "C:\Checkdat_G.mdb"
            Description = "CheckDat.mdb on Drive G"
        End If
        
        If LoopCount1 = 3 Then
            Location = "C:\Checkdat_K.mdb"
            Description = "CheckDat.mdb on Drive K"
        End If
        
        If LoopCount1 = 4 Then
            Location = "C:\Checkdat_Q.mdb"
            Description = "CheckDat.mdb on Drive Q"
        End If
        
        If LoopCount1 = 5 Then
            Location = "C:\Checkdat_T.mdb"
            Description = "CheckDat.mdb on Drive T"
        End If
    
        If LoopCount1 = 6 Then
            Location = "C:\Checkdat_Z.mdb"
            Description = "CheckDat.mdb on Drive Z"
        End If
        
        If fso.FileExists(Location) = True Then
            Dim Conn As ADODB.Connection
            Dim Rs As ADODB.Recordset
            
            Set Conn = New ADODB.Connection
            
            With Conn
              .CursorLocation = adUseClient
              .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & Location & "; Jet OLEDB:Database Password=CorpCaptive;"
              .Open
            End With
            
            strQuery = "SELECT [Branch Text 1], [Branch Text 2], [Branch Text 3], [Branch Text 4], [Branch Text 5], [Branch Text 6] FROM Branch WHERE [Routing Number] = '" & BRSTN & "'"
            
            Set Rs = New ADODB.Recordset
            
            With Rs
            
            Set .ActiveConnection = Conn
                .CursorType = adOpenStatic
                .Source = strQuery
                .Open
            End With
            
            If Rs.RecordCount <= 0 Then
                lstErrors.AddItem ("BRSTN " & BRSTN & " does not exists on " & Description)
            Else
            
                If Len(Rs.Fields(0)) >= 1 Then
                    Address1(AddressCount) = Rs.Fields(0)
                Else
                    Address1(AddressCount) = ""
                End If
                
                If Len(Rs.Fields(1)) >= 1 Then
                    Address2(AddressCount) = Rs.Fields(1)
                Else
                    Address2(AddressCount) = ""
                End If
                
                If Len(Rs.Fields(2)) >= 1 Then
                    Address3(AddressCount) = Rs.Fields(2)
                Else
                    Address3(AddressCount) = ""
                End If
                
                If Len(Rs.Fields(3)) >= 1 Then
                    Address4(AddressCount) = Rs.Fields(3)
                Else
                    Address4(AddressCount) = ""
                End If
                
                If Len(Rs.Fields(4)) >= 1 Then
                    Address5(AddressCount) = Rs.Fields(4)
                Else
                    Address5(AddressCount) = ""
                End If
                
                If Len(Rs.Fields(5)) >= 1 Then
                    Address6(AddressCount) = Rs.Fields(5)
                Else
                    Address6(AddressCount) = ""
                End If
                
                AddressDescription(AddressCount) = Description
                
                AddressCount = AddressCount + 1
            End If
        End If
        
        LoopCount1 = LoopCount1 + 1
    Loop

    LoopCount1 = 0
    Do Until LoopCount1 = AddressCount
        If LoopCount1 >= 1 Then
            If Address1(0) <> Address1(LoopCount1) Then
                lstErrors.AddItem ("Address 1 of BRSTN  " & BRSTN & " is Different on " & AddressDescription(LoopCount1) & " / " & Address1(0) & " -->" & Address1(LoopCount1))
            End If
            
            If Address2(0) <> Address2(LoopCount1) Then
                lstErrors.AddItem ("Address 2 of BRSTN  " & BRSTN & " is Different on " & AddressDescription(LoopCount1) & " / " & Address2(0) & " -->" & Address2(LoopCount1))
            End If
            
            If Address3(0) <> Address3(LoopCount1) Then
                lstErrors.AddItem ("Address 3 of BRSTN  " & BRSTN & " is Different on " & AddressDescription(LoopCount1) & " / " & Address3(0) & " -->" & Address3(LoopCount1))
            End If
            
            If Address4(0) <> Address4(LoopCount1) Then
                lstErrors.AddItem ("Address 4 of BRSTN  " & BRSTN & " is Different on " & AddressDescription(LoopCount1) & " / " & Address4(0) & " -->" & Address4(LoopCount1))
            End If
            
            If Address5(0) <> Address5(LoopCount1) Then
                lstErrors.AddItem ("Address 5 of BRSTN  " & BRSTN & " is Different on " & AddressDescription(LoopCount1) & " / " & Address5(0) & " -->" & Address5(LoopCount1))
            End If
            
            If Address6(0) <> Address6(LoopCount1) Then
                lstErrors.AddItem ("Address 6 of BRSTN  " & BRSTN & " is Different on " & AddressDescription(LoopCount1) & " / " & Address6(0) & " -->" & Address6(LoopCount1))
            End If
            
        End If
        
        LoopCount1 = LoopCount1 + 1
    Loop
    
    Me.Caption = "Checking Branch Addresses: " & BRSTN
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
End Sub

