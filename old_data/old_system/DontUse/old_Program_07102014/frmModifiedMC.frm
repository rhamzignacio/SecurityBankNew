VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmModifiedMC 
   BackColor       =   &H00932B2D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "For Program Modifications Only :) - MANAGER'S CHECK"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModifiedMC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00932B2D&
      Caption         =   "Recent Processed Files"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   6375
      Begin VB.TextBox txtFileName 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdReProcessFiles 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Re-Process File"
         DisabledPicture =   "frmModifiedMC.frx":030A
         Enabled         =   0   'False
         Height          =   855
         Left            =   3960
         Picture         =   "frmModifiedMC.frx":04CB
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid grdDisplay 
         Height          =   1935
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         HeadLines       =   1
         RowHeight       =   24
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "Mmm. DD, YYYY"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "FileName"
            Caption         =   "File Name"
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
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFileName 
         Caption         =   "File Name"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00932B2D&
         Caption         =   "Enter File Name:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00932B2D&
      Caption         =   "Process History"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6375
      Begin VB.CommandButton cmdProcessMeToday 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Process Me Today ! ! !"
         DisabledPicture =   "frmModifiedMC.frx":068C
         Height          =   855
         Left            =   3960
         Picture         =   "frmModifiedMC.frx":08FF
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblLastProcess 
         BackColor       =   &H00932B2D&
         Caption         =   "Last Process on "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00932B2D&
      Caption         =   "Ref.dbf"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdUseExisting 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Use Existing Ref.dbf ! ! !"
         DisabledPicture =   "frmModifiedMC.frx":0B72
         Height          =   855
         Left            =   3960
         Picture         =   "frmModifiedMC.frx":0F0B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblLastKnownRef 
         BackColor       =   &H00932B2D&
         Caption         =   "Last Known Ref.dbf : Sep. 23, 1988   06:00:45 PM"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label lblExistingRef 
         BackColor       =   &H00932B2D&
         Caption         =   "Existing Ref.dbf : Sep, 23, 1988  06:00:45 PM"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmModifiedMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcessMeToday_Click()

If MsgBox("Are you sure you want to make this program to Process today?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then Exit Sub

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

DataCount = 1
RepeatMe:

strQuery = "SELECT * FROM LastRun WHERE [Date] = '" & Format(Now, "MM/DD/YYYY") & "_" & DataCount & "'"
Rs.Open strQuery, Conn, adOpenStatic

If Rs.RecordCount >= 1 Then
    Rs.Close
    DataCount = DataCount + 1
    GoTo RepeatMe
    Exit Sub
End If

If Rs.RecordCount <= 0 Then
    Rs.Close
    
    strQuery = "UPDATE LastRun SET [Date] = '" & Format(Now, "MM/DD/YYYY") & "_" & DataCount & "' WHERE [Date] = '" & Format(Now, "MM/DD/YYYY") & "'"
    Rs.Open strQuery, Conn, adOpenStatic
End If

Unload Me
Me.Show

MsgBox "You can now Process Orders Today", vbInformation, "Success"

Unload Me
Me.Show
End Sub

Private Sub cmdReProcessFiles_Click()

If MsgBox("Are you sure you want to reprocess " & lblFileName.Caption & "?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then Exit Sub

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

strQuery = "DELETE * FROM Archive WHERE FileName = '" & lblFileName.Caption & "'"
Rs.Open strQuery, Conn, adOpenStatic

Description = "Reprocess File " & lblFileName.Caption
strQuery = "INSERT INTO AuditTrail (Description,[Date],[Time]) VALUES ('" _
    & Replace(Description, "'", "''") & "','" & Format(Now, "MM/DD/YYYY") & "','" & Format(Now, "HH:MM:SS") & "')"
Rs.Open strQuery, Conn, adOpenStatic

Unload Me
Me.Show

MsgBox "You can now reprocess the file again", vbInformation, ""

Unload Me
Me.Show
End Sub

Private Sub cmdUseExisting_Click()

If MsgBox("Are you sure you want to use this Ref for Serial Numbers?", vbYesNo + vbInformation, "Confirm Ref") = vbNo Then Exit Sub

Reasons = InputBox("Enter Reason why you want to change the ref.dbf", "", "")
If Reasons = "" Then Exit Sub

'Get Modified Date
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\MC\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")
'End Get Modified Date



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

strQuery = "SELECT [Date],[Time] FROM LastRefModified"
Rs.Open strQuery, Conn, adOpenStatic

LastKnownDate = Rs.Fields(0)
LastKnownTime = Rs.Fields(1)
Rs.Close

strQuery = "UPDATE LastRefModified SET [Date] = '" & ExistingRefDateModified & "', [Time] = '" & ExistingRefTimeModified & "'"
Rs.Open strQuery, Conn, adOpenStatic

Description = "Change Ref.dbf FROM " & LastKnownDate & " " & LastKnownTime & " to " & ExistingRefDateModified & " " & ExistingRefTimeModified
strQuery = "INSERT INTO AuditTrail ([Date],[Time], [Description],Reasons) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & Format(Now, "HH:MM:SS") & "','" & Description & "','" & Replace(Reasons, "'", "''") & "')"
'Rs.Open strQuery, Conn, adOpenStatic

Unload Me
Me.Show

MsgBox "Ref.dbf has been Changed", vbInformation, ""

Unload Me
Me.Show
End Sub

Private Sub Form_Load()

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

'For Strict Ref.dbf Modifications
strQuery = "SELECT [Date],[Time] FROM LastRefModified"
Rs.Open strQuery, Conn, adOpenStatic

lblLastKnownRef.Caption = "Last Known Ref.dbf : " & Format(Rs.Fields(0), "Mmm. DD, YYYY") & "   " & Format(Rs.Fields(1), "HH:MM:SS AMPM")
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\MC\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")

lblExistingRef.Caption = "Existing Ref.dbf : " & Format(ExistingRefDateModified, "Mmm. DD, YYYY") & "   " & Format(ExistingRefTimeModified, "HH:MM:SS AMPM")

If ExistingRefDateModified <> Rs.Fields(0) Or ExistingRefTimeModified <> Rs.Fields(1) Then
    cmdUseExisting.Enabled = True
Else
    cmdUseExisting.Enabled = False
End If
Rs.Close
'For Strict Ref.dbf Modifications

'For One-Time Run Only
strQuery = "SELECT [Date],[Time],Batch FROM LastRun ORDER BY PrimaryKey DESC"
Rs.Open strQuery, Conn, adOpenStatic

lblLastProcess.Caption = "Last Process on Batch " & Rs.Fields(2) & " on " & Format(Rs.Fields(0), "Mmm. DD, YYYY") & " " & Format(Rs.Fields(1), "HH:MM:SS AMPM")
Rs.Close

strQuery = "SELECT * FROM LastRun WHERE [Date] = '" & Format(Now, "MM/DD/YYYY") & "'"
Rs.Open strQuery, Conn, adOpenStatic

If Rs.RecordCount >= 1 Then
    cmdProcessMeToday.Enabled = True
Else
    cmdProcessMeToday.Enabled = False
End If
Rs.Close
'End For One-Time Run Only

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub txtFileName_Change()
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

strQuery = "SELECT * FROM Archive WHERE FileName Like '" & Replace(txtFileName.Text, "'", "''") & "%'"
Rs.Open strQuery, Conn, adOpenStatic

lblFileName.DataField = "FileName"
Set lblFileName.DataSource = Rs

Set grdDisplay.DataSource = Rs

If Rs.RecordCount >= 1 Then
    cmdReProcessFiles.Enabled = True
Else
    cmdReProcessFiles = False
End If
End Sub
