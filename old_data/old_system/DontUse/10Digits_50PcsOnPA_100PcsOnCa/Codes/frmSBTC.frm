VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSBTC 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SBTC 10 Digits (50 Pcs / Book on Personal, 100 Pcs / Book on Commercial)"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9804
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSBTC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4884
   ScaleWidth      =   9804
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton cmdGenerate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Generate ! ! !"
         DisabledPicture =   "frmSBTC.frx":030A
         Enabled         =   0   'False
         Height          =   975
         Left            =   7560
         Picture         =   "frmSBTC.frx":0630
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog dlgBrowse 
         Left            =   120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Browse ASC File"
         DisabledPicture =   "frmSBTC.frx":0956
         Height          =   975
         Left            =   120
         Picture         =   "frmSBTC.frx":0EB2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00404000&
         Caption         =   "Total Books Personal: 0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSDataGridLib.DataGrid grdDisplay 
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   9615
      _ExtentX        =   16955
      _ExtentY        =   4890
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648384
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "BRSTN"
         Caption         =   "BRSTN"
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
         DataField       =   "AccountNo"
         Caption         =   "Account No"
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
         DataField       =   "Name1"
         Caption         =   "Name"
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
         DataField       =   "OrderQty"
         Caption         =   "Order Qty"
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
            ColumnWidth     =   2124.283
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3780.284
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRefLocation 
      BackColor       =   &H00404000&
      Caption         =   "Ref Location:\"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   9495
   End
End
Attribute VB_Name = "frmSBTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GrandTotalSortRT, PageNo As String

Function PrintHeadingSortRT(PageNo)
If PageNo <> 1 Then Print #1, ""
Print #1, ""
Print #1, "    Page No. " & PageNo
Print #1, "    " & Format(Now, "Mmm. DD, YYYY")

Print #1, "                      Summary of RT nos / # of Books for SBTC"
Print #1, "                      SBTC Starter 50 Pcs Per Book on Personal"
Print #1, "                     SBTC Starter 100 Pcs Per Book on Commercial"

Print #1, ""
Print #1, "    ACCTNO       QTY ACCOUNT NAME"
Print #1, ""
End Function

Sub SortRT(Batch)
Close #1

Open App.Path & "\" & "SortRT.txt" For Output As #1

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(BRSTN),ChkType FROM SBTC ORDER BY ChkType"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LineNumber = 0
PageNo = 1
Result = PrintHeadingSortRT(PageNo)

GrandTotalSortRT = 0
LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    ChkType = dbfRecordset.Fields(1)
    
    If Val(LineNumber) >= 200 Then
        PageNo = Val(PageNo) + 1
        Result = PrintHeadingSortRT(PageNo)
        LineNumber = 0
    End If
    
    Print #1, ""
    Print #1, "   ** CHECK TYPE/BRSTN/BRANCH-->  " & ChkType & "/ " & BRSTN & " / " & getAddress(BRSTN, 1)
    Print #1, ""
    Print #1, "   * Batch # --> " & Batch
    LineNumber = LineNumber + 1
    
    Result = 0
    Result = PrintSortRT(BRSTN, Batch, LineNumber, ChkType)
    LineNumber = LineNumber + Val(Result)
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Print #1, ""
Print #1, "   ** Grand Total ** " & GrandTotalSortRT

Close #1
End Sub


Function PrintSortRT(BRSTN, Batch, LineNumber, ChkType)
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT AccountNo, OrderQty, Name1 FROM SBTC WHERE BRSTN = '" & BRSTN & "' AND ChkType = '" & ChkType & "' ORDER BY AccountNo"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
SubTotal = 0
Do Until LoopCount = dbfRecordset.RecordCount
    AccountNo = dbfRecordset.Fields(0)
    OrderQty = dbfRecordset.Fields(1)
    Do Until Len(OrderQty) = 4
        OrderQty = " " & OrderQty
    Loop
    
    If Len(dbfRecordset.Fields(2)) >= 1 Then
        Name1 = dbfRecordset.Fields(2)
    Else
        Name1 = ""
    End If
    
    Print #1, "    " & AccountNo & OrderQty & " " & Name1
    LineNumber = Val(LineNumber) + 1
    
    If LineNumber >= 50 Then
        PageNo = PageNo + 1
        Result = PrintHeadingSortRT(PageNo)
        LineNumber = 0
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    SubTotal = Val(SubTotal) + Val(OrderQty)
Loop

Print #1, ""
Print #1, "   ** Sub Total ** " & SubTotal
Print #1, ""
LineNumber = Val(LineNumber) + 1

GrandTotalSortRT = Val(GrandTotalSortRT) + Val(SubTotal)

PrintSortRT = LineNumber
End Function


Private Sub cmdBrowse_Click()
'For Configuration
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set DBFConnector1 = CreateObject("ADODB.Connection")
 
DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & getDataLocation & "\;Extended properties=dBase III"
DBFConnector1.CursorLocation = adUseClient

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection
Set Rs = New ADODB.Recordset

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & getDataLocation & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With
'End For Configuration

dlgBrowse.Filter = "ASC File|*.ASC"
dlgBrowse.InitDir = App.Path
dlgBrowse.ShowOpen

If dlgBrowse.FileName = "" Or Dir$(dlgBrowse.FileName) = "" Then Exit Sub

'Check if File has already been Processed
Set Rs = New ADODB.Recordset
strQuery = "SELECT [Date] FROM Archive_10 WHERE FileName = '" & dlgBrowse.FileTitle & "'"
Rs.Open strQuery, Conn, adOpenStatic

If Rs.RecordCount >= 1 Then
    MsgBox dlgBrowse.FileTitle & " has already been Processed Last " & Format(Rs.Fields(0), "Mmm. DD, YYYY"), vbCritical, "Error"
    Exit Sub
End If
'End Check if File has already been Processed



If MsgBox("Are you sure you want to Process " & dlgBrowse.FileTitle & "?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then Exit Sub




Close #1

Open dlgBrowse.FileName For Input As #1
Do Until EOF(1)
    Line Input #1, NotepadLine
    
    ChkType = Mid(NotepadLine, 1, 1)
    BRSTN = Mid(NotepadLine, 2, 9)
    AccountNo = Mid(NotepadLine, 13, 10)
    Name1 = Mid(NotepadLine, 23, 56)
    OrderQty = Val(Mid(NotepadLine, 82, 2))

    If ChkType = "A" Or ChkType = "B" Then
        Set dbfRecordset = CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO SBTC (BRSTN,AccountNo, Name1,OrderQty,ChkType) VALUES ('" _
            & BRSTN & "','" & AccountNo & "','" & Name1 & "','" & OrderQty & "','" & ChkType & "')"
        dbfRecordset.Open SQL, DBFConnector, 1, 1
    End If
Loop

Close #1

'Check if BRSTN Exists on Ref.dbf
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(BRSTN),ChkType FROM SBTC ORDER BY BRSTN"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    ChkType = dbfRecordset.Fields(1)
    
    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM REF WHERE RtNo = '" & BRSTN & "' AND ChkType = '" & ChkType & "'"
    dbfRecordset1.Open SQL, DBFConnector1, 1, 1
    
    If dbfRecordset1.RecordCount <= 0 Then
        MsgBox "BRSTN " & BRSTN & " with ChkType " & ChkType & " does not exists on Ref.dbf", vbInformation, "Error"
        Exit Sub
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
'End Check if BRSTN Exists on Ref.dbf

'Check Number of Books
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(OrderQty) FROM SBTC WHERE ChkType = 'A'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    TotalPA = FormatNumber(dbfRecordset.Fields(0), 0)
Else
    TotalPA = 0
End If

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT SUM(OrderQty) FROM SBTC WHERE ChkType = 'B'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    TotalCA = FormatNumber(dbfRecordset.Fields(0), 0)
Else
    TotalCA = 0
End If

lblTotal.Caption = "Total Books Personal: " & TotalPA & vbNewLine & "Total Books Commercial: " & TotalCA
'End Check Number of Books

If CheckUpdatedCheckDat = False Then Exit Sub

SortRT ("")

cmdBrowse.Enabled = False
cmdGenerate.Enabled = True
cmdGenerate.Default = True

DisplayDetails

MsgBox "Files has been Checked", vbInformation, ""
End Sub

Sub DisplayDetails()
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTC ORDER BY BRSTN,AccountNo,Name1"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Set grdDisplay.DataSource = dbfRecordset
End Sub

Function CheckUpdatedCheckDat()
Dim fso As New FileSystemObject

CheckUpdatedCheckDat = False

'Copy First the MDB
If fso.FileExists("\\Jenny\Y_CheckDat\Checkdat.mdb") = True Then
    If Dir$("C:\CheckDat_Y.mdb") <> "" Then Kill "C:\CheckDat_Y.mdb"
    fso.CopyFile "\\Jenny\Y_CheckDat\Checkdat.mdb", "C:\CheckDat_Y.mdb"
End If

If fso.FileExists("\\Modem-Computer\G_CheckDat\Checkdat.mdb") = True Then
    If Dir$("C:\CheckDat_G.mdb") <> "" Then Kill "C:\CheckDat_G.mdb"
    fso.CopyFile "\\Modem-Computer\G_CheckDat\Checkdat.mdb", "C:\CheckDat_G.mdb"
End If

If fso.FileExists("\\Karen\K_CheckDat\CheckDat.mdb") = True Then
    If Dir$("C:\CheckDat_K.mdb") <> "" Then Kill "C:\CheckDat_K.mdb"
    fso.CopyFile "\\Karen\K_CheckDat\CheckDat.mdb", "C:\CheckDat_K.mdb"
End If

If fso.FileExists("\\lenovo_xp\q_checkdat\CHECKDAT.MDB") = True Then
    If Dir$("C:\CheckDat_Q.mdb") <> "" Then Kill "C:\CheckDat_Q.mdb"
    fso.CopyFile "\\lenovo_xp\q_checkdat\CHECKDAT.MDB", "C:\CheckDat_Q.mdb"
End If

If fso.FileExists("\\Kapamilya\T_CheckDat\Checkdat.mdb") = True Then
    If Dir$("C:\CheckDat_T.mdb") <> "" Then Kill "C:\CheckDat_T.mdb"
    fso.CopyFile "\\Kapamilya\T_CheckDat\Checkdat.mdb", "C:\CheckDat_T.mdb"
End If

If fso.FileExists("\\Emily\Z_CheckDat\CheckDat.mdb") = True Then
    If Dir$("C:\CheckDat_Z.mdb") <> "" Then Kill "C:\CheckDat_Z.mdb"
    fso.CopyFile "\\Emily\Z_CheckDat\CheckDat.mdb", "C:\CheckDat_Z.mdb"
End If
'End Copy First the MDB

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(BRSTN) FROM SBTC"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    
    LoopCount1 = 0
    Do Until LoopCount1 = 6
        If LoopCount1 = 0 Then DatabaseLocation = getDataLocationCheckdat
        If LoopCount1 = 1 Then DatabaseLocation = "C:\CheckDat_Y.mdb"
        If LoopCount1 = 2 Then DatabaseLocation = "C:\CheckDat_G.mdb"
        If LoopCount1 = 3 Then DatabaseLocation = "C:\CheckDat_K.mdb"
        If LoopCount1 = 4 Then DatabaseLocation = "C:\CheckDat_Q.mdb"
        If LoopCount1 = 5 Then DatabaseLocation = "C:\CheckDat_T.mdb"
        If LoopCount1 = 6 Then DatabaseLocation = "C:\CheckDat_Z.mdb"
        
        If LoopCount1 = 0 Then Description = "Checkdat"
        If LoopCount1 = 1 Then Description = "Checkdat on Drive Y"
        If LoopCount1 = 2 Then Description = "Checkdat on Drive G"
        If LoopCount1 = 3 Then Description = "Checkdat on Drive K"
        If LoopCount1 = 4 Then Description = "Checkdat on Drive Q"
        If LoopCount1 = 5 Then Description = "Checkdat on Drive T"
        If LoopCount1 = 6 Then Description = "Checkdat on Drive Z"
        
        'For Configuration
        Dim Conn As ADODB.Connection
        Dim Rs As ADODB.Recordset
        
        Set Conn = New ADODB.Connection
        
        With Conn
          .CursorLocation = adUseClient
          .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & DatabaseLocation & "; Jet OLEDB:Database Password=CorpCaptive;"
          .Open
        End With
        'For Configuration
                
        Set Rs = New ADODB.Recordset
        strQuery = "SELECT [Branch Text 1], [Branch Text 2], [Branch Text 3], [Branch Text 4], [Branch Text 5], [Branch Text 6] FROM Branch WHERE [Routing Number] = '" & BRSTN & "'"
        Rs.Open strQuery, Conn, adOpenStatic
        
        If Rs.RecordCount <= 0 Then
            MsgBox "BRSTN " & BRSTN & " does not exists on " & Description, vbInformation, "Error"
            Exit Function
        Else
            If LoopCount1 = 0 Then
                If Len(Rs.Fields(0)) >= 1 Then
                    Orig_Address1 = Rs.Fields(0)
                Else
                    Orig_Address1 = ""
                End If
            
                If Len(Rs.Fields(1)) >= 1 Then
                    Orig_Address2 = Rs.Fields(1)
                Else
                    Orig_Address2 = ""
                End If
                
                If Len(Rs.Fields(2)) >= 1 Then
                    Orig_Address3 = Rs.Fields(2)
                Else
                    Orig_Address3 = ""
                End If
                
                If Len(Rs.Fields(3)) >= 1 Then
                    Orig_Address4 = Rs.Fields(3)
                Else
                    Orig_Address4 = ""
                End If
                
                If Len(Rs.Fields(4)) >= 1 Then
                    Orig_Address5 = Rs.Fields(4)
                Else
                    Orig_Address5 = ""
                End If
                
                If Len(Rs.Fields(5)) >= 1 Then
                    Orig_Address6 = Rs.Fields(5)
                Else
                    Orig_Address6 = ""
                End If
            Else
                If Len(Rs.Fields(0)) >= 1 Then
                    Address1 = Rs.Fields(0)
                Else
                    Address1 = ""
                End If
            
                If Len(Rs.Fields(1)) >= 1 Then
                    Address2 = Rs.Fields(1)
                Else
                    Address2 = ""
                End If
                
                If Len(Rs.Fields(2)) >= 1 Then
                    Address3 = Rs.Fields(2)
                Else
                    Address3 = ""
                End If
                
                If Len(Rs.Fields(3)) >= 1 Then
                    Address4 = Rs.Fields(3)
                Else
                    Address4 = ""
                End If
                
                If Len(Rs.Fields(4)) >= 1 Then
                    Address5 = Rs.Fields(4)
                Else
                    Address5 = ""
                End If
                
                If Len(Rs.Fields(5)) >= 1 Then
                    Address6 = Rs.Fields(5)
                Else
                    Address6 = ""
                End If
                
                If Address1 <> Orig_Address1 Then
                    MsgBox "Address 1 of BRSTN " & BRSTN & " does not Match on " & Description & vbNewLine & vbNewLine & Address1 & vbNewLine & Orig_Address1, vbInformation, "Error"
                    Exit Function
                End If
            
                If Address2 <> Orig_Address2 Then
                    MsgBox "Address 2 of BRSTN " & BRSTN & " does not Match on " & Description & vbNewLine & vbNewLine & Address2 & vbNewLine & Orig_Address2, vbInformation, "Error"
                    Exit Function
                End If
                
                If Address3 <> Orig_Address3 Then
                    MsgBox "Address 3 of BRSTN " & BRSTN & " does not Match on " & Description & vbNewLine & vbNewLine & Address3 & vbNewLine & Orig_Address3, vbInformation, "Error"
                    Exit Function
                End If
                
                If Address4 <> Orig_Address4 Then
                    MsgBox "Address 4 of BRSTN " & BRSTN & " does not Match on " & Description & vbNewLine & vbNewLine & Address4 & vbNewLine & Orig_Address4, vbInformation, "Error"
                    Exit Function
                End If
                
                If Address5 <> Orig_Address5 Then
                    MsgBox "Address 5 of BRSTN " & BRSTN & " does not Match on " & Description & vbNewLine & vbNewLine & Address5 & vbNewLine & Orig_Address5, vbInformation, "Error"
                    Exit Function
                End If
                
                If Address6 <> Orig_Address6 Then
                    MsgBox "Address 6 of BRSTN " & BRSTN & " does not Match on " & Desription & vbNewLine & vbNewLine & Address6 & vbNewLine & Orig_Address6, vbInformation, "Error"
                    Exit Function
                End If
            End If
        End If
        
        LoopCount1 = LoopCount1 + 1
    Loop
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

CheckUpdatedCheckDat = True
End Function

Private Sub cmdGenerate_Click()
Batch = InputBox("Enter Batch Number", "", "10CP" & Mid(dlgBrowse.FileTitle, 11, 4))
If Batch = "" Then Exit Sub

'For Configuration
Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
'End For Configuration

'BackUp Copy Ref.dbf
FileCopy getDataLocation & "\Ref.dbf", App.Path & "\Ref_Before\Ref_" & Batch & "_" & Format(Now, "MMDDYYHHMM") & ".dbf"
'End BackUp Copy Ref.dbf

DeleteDBF ("Packing")
DeleteDBF ("TransP")
DeleteDBF ("TransC")

Close #1, #2, #3, #4, #5, #6
Open App.Path & "\BlockP.txt" For Output As #1
Open App.Path & "\PrinterFileP.txt" For Output As #2

Open App.Path & "\BlockC.txt" For Output As #3
Open App.Path & "\PrinterFileC.txt" For Output As #4

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT BRSTN,AccountNo,Name1, OrderQty,ChkType FROM SBTC ORDER BY BRSTN,AccountNo"
dbfRecordset.Open SQL, DBFConnector, 1, 1


DataCount_PA = 0
DataCount_CA = 0
LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    AccountNo = dbfRecordset.Fields(1)
    
    If Len(dbfRecordset.Fields(2)) >= 1 Then
        Name1 = dbfRecordset.Fields(2)
    Else
        Name1 = ""
    End If
    
    OrderQty = dbfRecordset.Fields(3)
    ChkType = dbfRecordset.Fields(4)
    
    AccountNoWithHyphen = Mid(AccountNo, 1, 3) & "-" & Mid(AccountNo, 4, 6) & "-" & Mid(AccountNo, 10, 1)
    
    StartingSerial = getStartingSerialAndUpdate(BRSTN, OrderQty, ChkType) + 1
        
    Address1 = getAddress(BRSTN, 1)
    Address2 = getAddress(BRSTN, 2)
    Address3 = getAddress(BRSTN, 3)
    Address4 = getAddress(BRSTN, 4)
    Address5 = getAddress(BRSTN, 5)
    Address6 = getAddress(BRSTN, 6)
    
    If ChkType = "A" Then
        PcsPerBook = 50
        DBFFileName = "TransP"
        FormatSerial = "0000000"
        PrintNumber = 1
        DataCount = DataCount_PA
    End If

    If ChkType = "B" Then
        PcsPerBook = 100
        DBFFileName = "TransC"
        FormatSerial = "0000000000"
        PrintNumber = 3
        DataCount = DataCount_CA
    End If
    
    'TransP.dbf / TransC.dbf
    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "INSERT INTO " & DBFFileName & " (RT_NO,Acct_No, NO_BKS, CK_NO_P, RT_NO_P1, RT_NO_P2, Acct_No_P,Acct_Name1) VALUES ('" _
            & BRSTN & "','" & AccountNo & "','" & OrderQty & "','" & Format(StartingSerial, FormatSerial) & "','" & Mid(BRSTN, 1, 5) & "','" & Mid(BRSTN, 6, 4) & "','" & AccountNoWithHyphen & "','" & Name1 & "')"
    dbfRecordset1.Open SQL, DBFConnector, 1, 1
    'End TransP.dbf / TransC.dbf
    
    Do Until OrderQty = 0
        'For Do-Block
        If (DataCount_PA Mod 32 = 0 And ChkType = "A") Or (DataCount_CA Mod 32 = 0 And ChkType = "B") Then
            If DataCount <> 0 Then Print #PrintNumber, ""
            Print #PrintNumber, ""
            Print #PrintNumber, "        Page No. " & (DataCount / 32) + 1
            Print #PrintNumber, "        " & Format(Now, "Mmm. DD, YYYY")
            
            If ChkType = "A" Then
                Print #PrintNumber, "                       SBTC - SUMMARY OF BLOCK - PERSONAL"
                Print #PrintNumber, "                               STARTER 50 PIECES"
            End If
            
            If ChkType = "B" Then
                Print #PrintNumber, "                       SBTC - SUMMARY OF BLOCK - COMMERCIAL"
                Print #PrintNumber, "                                 STARTER 100 PIECES"
            End If
            
            Print #PrintNumber, ""
            Print #PrintNumber, "        BLOCK RT_NO     M ACCT_NO       START_NO.  END_NO."
        End If
        
        If (DataCount_PA Mod 4 = 0 And ChkType = "A") Or (DataCount_CA Mod 4 = 0 And ChkType = "B") Then
            If ChkType = "A" Then
                BlockCount_PA = BlockCount_PA + 1
                BlockCount = BlockCount_PA
            End If
            
            If ChkType = "B" Then
                BlockCount_CA = BlockCount_CA + 1
                BlockCount = BlockCount_CA
            End If
            
            Print #PrintNumber, ""
            Print #PrintNumber, "       ** BLOCK " & Val(BlockCount)
        End If
        
        Do Until Len(BlockCount) >= 13
            BlockCount = " " & BlockCount
        Loop
            
        If ChkType = "A" Then Print #PrintNumber, BlockCount & " " & BRSTN & "   " & AccountNo & "    " & Format(StartingSerial, FormatSerial) & "    " & Format(StartingSerial + PcsPerBook - 1, FormatSerial)
        If ChkType = "B" Then Print #PrintNumber, BlockCount & " " & BRSTN & "   " & AccountNo & "    " & Format(StartingSerial, FormatSerial) & " " & Format(StartingSerial + PcsPerBook - 1, FormatSerial)
        'End For Do-Block
        
        'Packing.dbf
        Set dbfRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO Packing (BatchNo, Block,RT_NO,Branch,Acct_No,ChkType,Acct_Name1,NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E,Acct_No_P) VALUES ('" _
            & Batch & "','" & Val(BlockCount) & "','" & BRSTN & "','" & Replace(Address1, "'", "''") & "','" & AccountNo & "','" & ChkType & "','" & Replace(Name1, "'", "''") & "','" & "1" & "','" & Format(StartingSerial, FormatSerial) & "','" & Format(StartingSerial, FormatSerial) & "','" & Format(StartingSerial + PcsPerBook - 1, FormatSerial) & "','" & Format(StartingSerial + PcsPerBook - 1, FormatSerial) & "','" & AccountNoWithHyphen & "')"
        dbfRecordset1.Open SQL, DBFConnector, 1, 1
        'End Packing.dbf
        
        'Printer File
        Print #PrintNumber + 1, "3"
        Print #PrintNumber + 1, BRSTN
        Print #PrintNumber + 1, AccountNo
        Print #PrintNumber + 1, Format(StartingSerial + PcsPerBook, FormatSerial)
        Print #PrintNumber + 1, "A"
        
        If ChkType = "A" Then Print #PrintNumber + 1, "     ONNNNNNNO" & Mid(BRSTN, 1, 5) & "D" & Mid(BRSTN, 6, 4) & "T" & Format(AccountNo, "000000000000") & "O"
        If ChkType = "B" Then Print #PrintNumber + 1, "  ONNNNNNNNNNO" & Mid(BRSTN, 1, 5) & "D" & Mid(BRSTN, 6, 4) & "T" & Format(AccountNo, "000000000000") & "O"
        
        Print #PrintNumber + 1, Mid(BRSTN, 1, 5)
        Print #PrintNumber + 1, " " & Mid(BRSTN, 6, 4)
        Print #PrintNumber + 1, AccountNoWithHyphen
        Print #PrintNumber + 1, Name1
        Print #PrintNumber + 1, "SN"
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, "C"
        Print #PrintNumber + 1, "XXXX"
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, Address1
        Print #PrintNumber + 1, Address2
        Print #PrintNumber + 1, Address3
        Print #PrintNumber + 1, Address4
        Print #PrintNumber + 1, Address5
        Print #PrintNumber + 1, Address6
        Print #PrintNumber + 1, "SECURITY BANK"
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, ""
        Print #PrintNumber + 1, Format(StartingSerial, FormatSerial)
        Print #PrintNumber + 1, Format(StartingSerial + PcsPerBook - 1, FormatSerial)
        'End Printer File
        
        If ChkType = "A" Then DataCount_PA = DataCount_PA + 1
        If ChkType = "B" Then DataCount_CA = DataCount_CA + 1
        
        OrderQty = OrderQty - 1
    Loop
    
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
    oldBRSTN_A = BRSTN
    oldBRSTN_B = BRSTN
Loop

Print #PrintNumber, "\"
Print #PrintNumber + 1, "\"


SortRT (Batch)
Result = PackingText(Batch, "A")
Result = PackingText(Batch, "B")

LimitData

'Copy Ref.dbf
If Dir$(App.Path & "\Ref.dbf") <> "" Then Kill App.Path & "\Ref.dbf"
FileCopy getDataLocation & "\Ref.dbf", App.Path & "\Ref.dbf"
'End Copy Ref.dbf

MsgBox "Files has been Generated", vbInformation, ""
End
End Sub

Function PackingText(Batch, ChkType)
PageNo = 1
Close #1
Open App.Path & "\Packing" & ChkType & ".txt" For Output As #1

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT BRSTN,COUNT(BRSTN) FROM SBTC WHERE ChkType = '" & ChkType & "' GROUP BY BRSTN ORDER BY BRSTN"
dbfRecordset.Open SQL, DBFConnector, 1, 1

DataCount = 0
LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    SubTotal = dbfRecordset.Fields(1)
    
    Result = PrintHeadingPacking(PageNo, ChkType)
    Print #1, " ** ORDERS OF BRSTN " & BRSTN & " " & getAddress(BRSTN, 1)
    Print #1, ""
    Print #1, " * Batch Number " & Batch
    Print #1, ""
    
    'Select the Details
    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT Acct_NO_P, Acct_Name1, CK_NO_B, CK_NO_E FROM Packing WHERE RT_NO = '" & BRSTN & "' AND ChkType = '" & ChkType & "' ORDER BY Acct_NO_P"
    dbfRecordset1.Open SQL, DBFConnector, 1, 1
    'End Select the Details
    
    LoopCount1 = 0
    Do Until LoopCount1 = dbfRecordset1.RecordCount
        AccountNo = dbfRecordset1.Fields(0)
        Acct_Name1 = dbfRecordset1.Fields(1)
        StartingSerial = dbfRecordset1.Fields(2)
        EndingSerial = dbfRecordset1.Fields(3)
        
        Do Until Len(Acct_Name1) >= 33
            Acct_Name1 = Acct_Name1 & " "
        Loop
        
        Do Until Len(StartingSerial) >= 11
            StartingSerial = StartingSerial & " "
        Loop

        If DataCount >= 50 Then
            DataCount = 0
            PageNo = PageNo + 1
            Result = PrintHeadingPacking(PageNo, ChkType)
        End If
        
        Print #1, "     " & AccountNo & "  " & Acct_Name1 & "1 " & ChkType & "  " & StartingSerial & EndingSerial
        
        DataCount = DataCount + 1
        dbfRecordset1.MoveNext
        LoopCount1 = LoopCount1 + 1
    Loop
    
    Print #1, ""
    Print #1, " ** Sub Total ** " & SubTotal
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Close #1
End Function

Function PrintHeadingPacking(PageNo, ChkType)
If PageNo <> "1" Then Print #1, ""

Print #1, "  Page No. " & PageNo
Print #1, "  " & Format(Now, "Mmm. DD, YYYY")
Print #1, "                             CAPTIVE PRINTING CORPORATION"

If ChkType = "A" Then Print #1, "                          SBTC - Personal Checks Summary"
If ChkType = "B" Then Print #1, "                         SBTC - Commercial Checks Summary"

Print #1, ""
Print #1, "  ACCT_NO          ACCOUNT NAME                   QTY CT START #    END #"
Print #1, ""
End Function

Sub LimitData()
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(getDataLocation & "\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & getDataLocation & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With
'For Configuration
        
Set Rs = New ADODB.Recordset
strQuery = "UPDATE LastRefModified SET [Date] = '" & ExistingRefDateModified & "', [Time] = '" & ExistingRefTimeModified & "'"
Rs.Open strQuery, Conn, adOpenStatic
        
Set Rs = New ADODB.Recordset
strQuery = "INSERT INTO Archive_10 ([Date], FileName) VALUES ('" _
    & Format(Now, "MM/DD/YYYY") & "','" & dlgBrowse.FileTitle & "')"
Rs.Open strQuery, Conn, adOpenStatic

End Sub

Private Sub Form_Load()
lblRefLocation.Caption = "Ref.dbf Location: " & Replace(getDataLocation, "\\Karen\Captive\", "K:\")
If App.PrevInstance Then End

lblTotal.Caption = "Total Books Personal: 0" & vbNewLine & "Total Books Commercial: 0"
CheckIfRefHasBeenModified
CreateDBF

End Sub

Sub CheckIfRefHasBeenModified()
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(getDataLocation & "\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection
Set Rs = New ADODB.Recordset

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & getDataLocation & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "SELECT [Date],[Time] FROM LastRefModified"
Rs.Open strQuery, Conn, adOpenStatic

If ExistingRefDateModified <> Rs.Fields(0) Or ExistingRefTimeModified <> Rs.Fields(1) Then
    MsgBox "Unable to Run. Ref.dbf on Regular Checks has been Modified", vbCritical, "Error"
    End
End If
End Sub

Sub CreateDBF()
If Dir$(App.Path & "\SBTC.dbf") <> "" Then Kill App.Path & "\SBTC.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE SBTC (BRSTN Varchar(50), AccountNo Varchar(50), Name1 Varchar(50), OrderQty Varchar(50), ChkType Varchar(50))"
dbfRecordset.Open SQL, DBFConnector, 1, 1
End Sub
