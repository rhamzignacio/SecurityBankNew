VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Bank"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDeliveryDate 
      BackColor       =   &H00404000&
      Caption         =   "Delivery Date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5535
      Begin MSComCtl2.DTPicker dteDeliveryDate 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   873
         _Version        =   393216
         Format          =   96665600
         CurrentDate     =   42741
      End
   End
   Begin VB.CommandButton cmdEncode 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Encode"
      Height          =   975
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.FileListBox fleTemp 
      Height          =   345
      Left            =   1320
      Pattern         =   "*.zip"
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheckFilesHead 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Check files on Head"
      Default         =   -1  'True
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.FileListBox fleHead 
      Height          =   4920
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "This will auto 'Rename' Files:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1935
      Left            =   5760
      TabIndex        =   5
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label lblHashTotal 
      BackColor       =   &H000000FF&
      Caption         =   "Warning: 1 Hash Total hasn't been Sent Yet"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00404000&
      Caption         =   "Total: 0"
      ForeColor       =   &H000080FF&
      Height          =   4335
      Left            =   5760
      TabIndex        =   2
      Top             =   1200
      Width           =   5535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub CreateTable()
If Dir$(App.Path & "\SBTC.dbf") <> "" Then Kill App.Path & "\SBTC.dbf"
If Dir$(App.Path & "\Temp.dbf") <> "" Then Kill App.Path & "\Temp.dbf"
If Dir$(App.Path & "\Errors.dbf") <> "" Then Kill App.Path & "\Errors.dbf"
If Dir$(App.Path & "\Batch.dbf") <> "" Then Kill App.Path & "\Batch.dbf"
If Dir$(App.Path & "\Temp1.dbf") <> "" Then Kill App.Path & "\Temp1.dbf"




Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE SBTC (ChkType Varchar(6),BRSTN Varchar(9) , AccountNo Varchar(12), Name1 Varchar(60), Name2 Varchar(60), FormType Varchar(2),OrderQty Varchar(3), Batch Varchar(30), Address1 Varchar(60), Address2 Varchar(60), Address3 Varchar(60), Address4 Varchar(60), Address5 Varchar(60), Address6 Varchar(60), PKey Numeric, BStock Varchar(50), FileName Varchar(50), StartSN Varchar(50), PcsPerBook Varchar(3) , StartSN1 numeric)"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE Temp (AccountNo Varchar(12), AcctName Varchar(60), Batch Varchar(30), Pkey Numeric, FileName Varchar(50))"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE Errors (Errors Varchar(244))"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE Batch (Batch Varchar(244))"
dbfRecordset.Open SQL, DBFConnector, 1, 1

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE Temp1 (Orig Varchar(244), New Varchar(244))"
dbfRecordset.Open SQL, DBFConnector, 1, 1
End Sub








Sub CheckModifiedFileName(FileName)


If Dir$("\\karen\captive\auto\SBTC\Regular\Attachments\" & FileName) = "" Then
    MsgBox "Unable to Process since Filename " & FileName & " has been renamed", vbInformation, "Error"
    End
End If




Close #1
Open fleHead.Path & "\" & FileName For Input As #1

LineNumber_1 = 0

Do Until EOF(1)
    Line Input #1, NotepadLine1
    
    
    LineNumber_2 = 0
    Close #2
    Open "\\karen\captive\auto\SBTC\Regular\Attachments\" & FileName For Input As #2
    Do Until EOF(2)
        Line Input #2, NotepadLine2
        If LineNumber_1 = LineNumber_2 Then
            If NotepadLine1 <> NotepadLine2 Then
                MsgBox "Unable to Process since filename " & FileName & " has been modified" & vbNewLine & vbNewLine & vbNewLine & vbnewnline & "Original: " & vbNewLine & NotepadLine2 & vbNewLine & vbNewLine & FileName & ":" & vbNewLine & NotepadLine1, vbInformation, "Error"
                End
            End If
        End If
        
        LineNumber_2 = LineNumber_2 + 1
    Loop
    Close #2
    
    LineNumber_1 = LineNumber_1 + 1
Loop
Close #1


If LineNumber_1 <> LineNumber_2 Then
    MsgBox "Unable to Process since filename " & FileName & "has been modified", vbInformation, "Error"
    End
End If
End Sub

Private Sub cmdCheckFilesHead_Click()
cmdEncode.Enabled = False



Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient





If cmdCheckFilesHead.Caption = "&Check files on Head" Then

    
    
    'Check the Delivery Date
    If DateValue(dteDeliveryDate.Value) = DateValue(Now) Then
        MsgBox "Please set the Delivery Date", vbCritical, "Error"
        Exit Sub
    End If
    
    fraDeliveryDate.Enabled = False
    'End Check the Delivery Date
    
    
    
    
    
    If MsgBox("Are you sure you want to check files?", vbYesNo + vbInformation, "Confirm Check") = vbNo Then Exit Sub
        
        
        
        
    
    CreateTable
    
    
    
    
    RenameFiles
    
    
    
    LoopCount = 0
    Do Until LoopCount = fleHead.ListCount
        FileName = fleHead.List(LoopCount)



        ReadASCFiles (FileName)
        
        
        
        LoopCount = LoopCount + 1
    Loop
    
    
    
    CheckUpdatedCheckDat
    
    
    
    CheckRef
    
    
    
    If CheckErrors >= 1 Then
    
        MsgBox "Errors has been found. See errors.txt", vbCritical, "Error"
        
        End
        
    End If
        
    
    
    CheckDoubleAccount
    
    
    
    DisplayDetails
    
    
    
    cmdCheckFilesHead.Caption = "Process ! ! !"
    
    
    
    'For Sort RT
    SortRT ("Regular")
    SortRT ("Regular\PreEncoded")
    SortRT ("MC")
    SortRT ("CheckOne")
    SortRT ("CheckPower")
    SortRT ("GiftCheck")
    'End For Sort RT
    
    
    
    MsgBox "Data has been Checked. No Errors Found", vbInformation, ""
    Exit Sub
Else



    Batch = InputBox("Enter Batch Name", "", Format(Now, "MMDDYYYY"))
    If Batch = "" Then Exit Sub
    
    
    
    If Dir$(App.Path & "\Archive\" & Batch, vbDirectory) <> "" Then
        MsgBox "Batch " & Batch & " has already been processed", vbInformation, "Error"
        Exit Sub
    End If
        
        
        
        
    
    
    
    
    'Process By / Checked By
    ProcessBy = InputBox("Enter Process By", "", "")
    If ProcessBy = "" Then Exit Sub
    
    CheckedBy = InputBox("Enter Checked By", "", "")
    If CheckedBy = "" Then Exit Sub
    
    If UCase(ProcessBy) = UCase(CheckedBy) Then
        MsgBox "Prepared by and Checked By should not be the same", vbCritical, "Error"
        Exit Sub
    End If
    'End Process By / Checked By
    
        
    'For Zip
    DateTimeToday = Format(Now, "MMDDYYHHMMSS")
    
    MkDir ("C:\Windows\Temp\" & DateTimeToday)
    MkDir ("C:\Windows\Temp\" & DateTimeToday & "\" & Batch)
    'End For Zip



    'Clear Folders
    fleTemp.Pattern = "*.txt;*.mdb;*.**P"
    
    LoopCount = 0
    Do Until LoopCount = 10
        If LoopCount = 0 Then fleTemp.Path = App.Path & "\Charge_Slip\"
        If LoopCount = 1 Then fleTemp.Path = App.Path & "\CheckOne\"
        If LoopCount = 2 Then fleTemp.Path = App.Path & "\CheckPower\"
        If LoopCount = 3 Then fleTemp.Path = App.Path & "\Customized\"
        If LoopCount = 4 Then fleTemp.Path = App.Path & "\GiftCheck\"
        If LoopCount = 5 Then fleTemp.Path = App.Path & "\MC\"
        If LoopCount = 6 Then fleTemp.Path = App.Path & "\MC\Continues\"
        If LoopCount = 7 Then fleTemp.Path = App.Path & "\Regular\"
        If LoopCount = 8 Then fleTemp.Path = App.Path & "\Regular\PreEncoded\"
        fleTemp.Pattern = "*.txt;*.mdb;*.**P"
        
        If LoopCount = 9 Then
            fleTemp.Path = App.Path
            fleTemp.Pattern = "*.zip"
        End If
        
        
        fleTemp.Refresh
        LoopCount1 = 0
        Do Until LoopCount1 = fleTemp.ListCount
            Filee = UCase(fleTemp.List(LoopCount1))
            
            If Filee <> "SORTRT.TXT" Then Kill fleTemp.Path & "\" & Filee
            
            LoopCount1 = LoopCount1 + 1
        Loop
        
        LoopCount = LoopCount + 1
    Loop
    'End Clear Folders
    
    
    TransferAll (Batch)
    
    Result = ProcessAll(Batch, ProcessBy, CheckedBy, dteDeliveryDate.Value)
    
    If MsgBox("Data has been processed", vbInformation, "") = vbNo Then End
    End
End If
End Sub




Sub CheckDoubleAccount()
Close #1
Open App.Path & "\Double_Accounts.txt" For Output As #1


Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient


Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT AccountNo, COUNT(AccountNo) FROM SBTC GROUP BY AccountNo"
dbfRecordset.Open SQL, DBFConnector, 1, 1


All_Double = ""

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    AccountNo = dbfRecordset.Fields(0)
    OrderCount = dbfRecordset.Fields(1)
    
    If OrderCount > 1 Then
        If All_Double = "" Then
            All_Double = AccountNo
        Else
            All_Double = All_Double & vbNewLine & AccountNo
        End If
        
        'Check the Double Account
        Set dbfRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "SELECT OrderQty, FileName FROM SBTC WHERE AccountNo = '" & AccountNo & "'"
        dbfRecordset1.Open SQL, DBFConnector, 1, 1
        
        LoopCount1 = 0
        Do Until LoopCount1 = dbfRecordset1.RecordCount
            OrderQty = dbfRecordset1.Fields(0)
            FileName = dbfRecordset1.Fields(1)
            
            Do Until Len(OrderQty) >= 5
                OrderQty = OrderQty & " "
            Loop
            
            Print #1, AccountNo & "  " & OrderQty & FileName
            
            dbfRecordset1.MoveNext
            LoopCount1 = LoopCount1 + 1
        Loop
        
        Print #1, ""
        'End Check the Double Account
    End If
    
    
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Close #1

If All_Double <> "" Then
    If MsgBox("There are double Accounts on Order. Continue?" & vbNewLine & vbNewLine & "Please see 'Double Accounts.txt' for more details", vbYesNo + vbInformation, "Confirm Continue") = vbNo Then End
End If


End Sub


Sub RenameFiles()
Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient




LoopCount = 0
Do Until LoopCount = fleHead.ListCount
    FileName = UCase(fleHead.List(LoopCount))
    NewFileName = ""
    
    
    If Mid(FileName, 1, 6) = "YSECPT" Then NewFileName = "YSE13Digit_cp" & Mid(FileName, 7, Len(FileName))
    If Mid(FileName, 1, 6) = "NOVABU" Then NewFileName = "13Digit_nb" & Mid(FileName, 7, Len(FileName))
    If Mid(FileName, 1, 6) = "CONSOL" Then NewFileName = "13Digit_cs" & Mid(FileName, 7, Len(FileName))
    If Mid(FileName, 1, 6) = "CPTIVE" Then NewFileName = "13Digit_cp" & Mid(FileName, 7, Len(FileName))
    If Mid(FileName, 1, 6) = "CAPTIVE" Then NewFileName = "13Digit_cp" & Mid(FileName, 8, Len(FileName))
    
    If Mid(FileName, 1, 5) = "YSETG" Then
        MsgBox "Unable to process file with YSETG", vbInformation, "Tone Guide File"
        End
    End If

    If NewFileName = "" Then
        MsgBox "Unable to read text file. Please save ALL the original file", vbCritical, "Error"
        End
    End If

    Set dbfRecordset = CreateObject("ADODB.Recordset")
    SQL = "INSERT INTO Temp1 (Orig , New) VALUES ('" & Mid(FileName, 1, Len(FileName) - 4) & "','" & Mid(NewFileName, 1, Len(NewFileName) - 4) & "')"
    dbfRecordset.Open SQL, DBFConnector, 1, 1



    FileCopy fleHead.Path & "\" & FileName, fleHead.Path & "\" & NewFileName
    Kill fleHead.Path & "\" & FileName
            
            
    LoopCount = LoopCount + 1
Loop

fleHead.Refresh
End Sub




Function CheckErrors()
Close #1
Open App.Path & "\Errors.txt" For Output As #1

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT ERRORS FROM Errors"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    
    Print #1, dbfRecordset.Fields(0)
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Close #1

CheckErrors = dbfRecordset.RecordCount
End Function



Sub CheckRef()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT ChkType,BRSTN, FormType, Address1, Batch FROM SBTC GROUP BY ChkType,BRSTN, FormType, Address1,Batch"
dbfRecordset.Open SQL, DBFConnector, 1, 1

MeCaption = Me.Caption



LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Me.Caption = "Checking Ref --> " & FormatNumber((LoopCount / dbfRecordset.RecordCount) * 100, 5)
    
    ChkType = dbfRecordset.Fields(0)
    BRSTN = dbfRecordset.Fields(1)
    FormType = dbfRecordset.Fields(2)
    
    If Len(dbfRecordset.Fields(3)) >= 1 Then
        BranchName = dbfRecordset.Fields(3)
    Else
        BranchName = ""
    End If
    
    Batch = dbfRecordset.Fields(4)
    

    
    Ref_Location = ""
    If ChkType = "A" And FormType = "05" Then Ref_Location = App.Path & "\Regular\"
    If ChkType = "B" And FormType = "16" Then Ref_Location = App.Path & "\Regular\"
    If ChkType = "AA" And FormType = "05" Then Ref_Location = App.Path & "\Regular\"
    If ChkType = "BB" And FormType = "16" Then Ref_Location = App.Path & "\Regular\"
    
    If ChkType = "MC" And FormType = "20" Then Ref_Location = App.Path & "\MC\"
        
    If ChkType = "F" And FormType = "25" Then Ref_Location = App.Path & "\CheckOne\"
    If ChkType = "F" And FormType = "26" Then Ref_Location = App.Path & "\CheckOne\"
    
    If ChkType = "E" And FormType = "23" Then Ref_Location = App.Path & "\CheckPower\"
    If ChkType = "E" And FormType = "22" Then Ref_Location = App.Path & "\CheckPower\"
    
    If ChkType = "GC" And FormType = "20" Then Ref_Location = App.Path & "\GiftCheck\"
        
    
    If ChkType = "A" And FormType = "05" Then RefChkType = "A"
    If ChkType = "B" And FormType = "16" Then RefChkType = "B"
    
    If ChkType = "AA" And FormType = "05" Then RefChkType = "A"
    If ChkType = "BB" And FormType = "16" Then RefChkType = "B"
    
    If ChkType = "MC" And FormType = "20" Then RefChkType = "A"
    
    If ChkType = "F" And FormType = "25" Then RefChkType = "A"
    If ChkType = "F" And FormType = "26" Then RefChkType = "B"
    
    If ChkType = "GC" And FormType = "20" Then RefChkType = "A"
    
    If ChkType = "E" And FormType = "23" Then RefChkType = "A"
    If ChkType = "E" And FormType = "22" Then RefChkType = "B"
    
    
    
    
    'For Ref Location
    Set DBFConnector1 = CreateObject("ADODB.Connection")

    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Ref_Location & "\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    'End For Ref Location

    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT Branch_Tex FROM REF WHERE RTNO = '" & BRSTN & "' AND ChkType = '" & RefChkType & "'"
    dbfRecordset1.Open SQL, DBFConnector1, 1, 1
    
    
    If dbfRecordset1.RecordCount <> 1 Then
        Description = "BRSTN " & BRSTN & " with Chktype " & RefChkType & " contains error on Ref.dbf on File " & Batch
        
        SaveError (Description)
    Else
        Ref_BranchName = dbfRecordset1.Fields(0)
        
        If Ref_BranchName <> BranchName Then
            Description = vbNewLine & "BRSTN " & BRSTN & " does not match on Branches.dbf" & vbNewLine & "FTS: " & BranchName & vbNewLine & "Ref: " & Ref_BranchName & vbNewLine & Ref_Location & vbNewLine
            
            SaveError (Description)
        End If
    End If
        

    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

Me.Caption = MeCaption

End Sub


Sub TransferAll(Batch)


LoopCount = 0
Do Until LoopCount = fleHead.ListCount
    If LoopCount = 0 Then
        If Dir$(App.Path & "\Archive\" & Batch, vbDirectory) = "" Then MkDir (App.Path & "\Archive\" & Batch)
    End If
    
    FileName = fleHead.List(LoopCount)
    
    FileCopy App.Path & "\Head\" & FileName, "C:\Windows\Temp\" & DateTimeToday & "\" & FileName
    FileCopy App.Path & "\Head\" & FileName, App.Path & "\Archive\" & Batch & "\" & FileName
    Kill App.Path & "\Head\" & FileName
    
    
    LoopCount = LoopCount + 1
Loop
End Sub




Sub CheckUpdatedCheckDat()
Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(BRSTN) FROM SBTC"
dbfRecordset.Open SQL, DBFConnector, 1, 1


MeCaption = Me.Caption



LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Me.Caption = FormatNumber((LoopCount / dbfRecordset.RecordCount) * 100, 5) & " % D o n e . . . Checking Address"
    
    BRSTN = dbfRecordset.Fields(0)
    
    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT Address1 , Address2 , Address3 , Address4 , Address5 , Address6 FROM Branches WHERE BRSTN = '" & BRSTN & "'"
    dbfRecordset1.Open SQL, DBFConnector, 1, 1
    
    If dbfRecordset1.RecordCount <= 0 Then
        
        SaveError ("BRSTN " & BRSTN & " does not exists on " & Description)
        
    Else
        If Len(dbfRecordset1.Fields(0)) >= 1 Then
            Address1 = dbfRecordset1.Fields(0)
        Else
            Address1 = ""
        End If
        
        If Len(dbfRecordset1.Fields(1)) >= 1 Then
            Address2 = dbfRecordset1.Fields(1)
        Else
            Address2 = ""
        End If
        
        If Len(dbfRecordset1.Fields(2)) >= 1 Then
            Address3 = dbfRecordset1.Fields(2)
        Else
            Address3 = ""
        End If
        
        If Len(dbfRecordset1.Fields(3)) >= 1 Then
            Address4 = dbfRecordset1.Fields(3)
        Else
            Address4 = ""
        End If
        
        If Len(dbfRecordset1.Fields(4)) >= 1 Then
            Address5 = dbfRecordset1.Fields(4)
        Else
            Address5 = ""
        End If
        
        If Len(dbfRecordset1.Fields(5)) >= 1 Then
            Address6 = dbfRecordset1.Fields(5)
        Else
            Address6 = ""
        End If
                
        Set dbfRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "UPDATE SBTC SET Address1 = '" & Replace(Address1, "'", "''") & "', Address2 = '" & Replace(Address2, "'", "''") & "', Address3 = '" & Replace(Address3, "'", "''") & "', Address4 = '" & Replace(Address4, "'", "''") & "', Address5 = '" & Replace(Address5, "'", "''") & "', Address6 = '" & Replace(Address6, "'", "''") & "' WHERE BRSTN = '" & BRSTN & "'"
        dbfRecordset1.Open SQL, DBFConnector, 1, 1
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop



Me.Caption = MeCaption

End Sub



Sub DisplayDetails()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT ChkType, FormType, SUM(OrderQty) FROM SBTC GROUP BY ChkType, FormType ORDER BY ChkType, FormType"
dbfRecordset.Open SQL, DBFConnector, 1, 1

lblTotal.Caption = ""

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    ChkType = dbfRecordset.Fields(0)
    FormType = dbfRecordset.Fields(1)
    OrderQty = dbfRecordset.Fields(2)
    
    
    
    ChequeName = getChequeName(ChkType, FormType)
    
    
    Total_Qty = ChequeName & ": " & OrderQty
    
    
    If lblTotal.Caption = "" Then
        lblTotal.Caption = Total_Qty
    Else
        lblTotal.Caption = lblTotal.Caption & vbNewLine & Total_Qty
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

lblTotal.Caption = lblTotal.Caption & vbNewLine & vbNewLine


Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Batch, SUM(OrderQty) FROM SBTC GROUP BY Batch ORDER BY Batch"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Batch = dbfRecordset.Fields(0)
    OrderQty = dbfRecordset.Fields(1)
    
    lblTotal.Caption = lblTotal.Caption & vbNewLine & Batch & ": " & OrderQty
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
End Sub



Sub ReadExcel(FileName)
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient



'Close All Excel Files
TaskID = Shell("taskkill.exe /f /t /im Excel.exe", vbHide)

hProcess = OpenProcess(SYNCHRONIZE, True, TaskID)
Call WaitForSingleObject(hProcess, WAIT_INFINITE)
CloseHandle hProcess
'End Close All Excel Files




Dim wbkmvr As Excel.Workbook
Dim xlapp As Excel.Application
Dim wksline As Excel.Worksheet

Set xlapp = Excel.Application

xlapp.Workbooks.Open (App.Path & "\Head\" & FileName)
Set wbkmvr = xlapp.ActiveWorkbook
Set wksline = wbkmvr.ActiveSheet

TotalRecords = Val(wksline.UsedRange.Rows.Count)

LoopCount = 0
Do Until LoopCount = TotalRecords
    BRSTN = Replace(Trim(wksline.Cells(LoopCount + 1, 1)), "-", "")
    MC_AccountNo = Right(Replace(Trim(wksline.Cells(LoopCount + 1, 4)), "-", ""), 12)
    GC_AccountNo = Right(Replace(Trim(wksline.Cells(LoopCount + 1, 5)), "-", ""), 12)
    
    
    
    If Len(BRSTN) = 9 And IsNumeric(BRSTN) = True And Len(MC_AccountNo) = 12 And IsNumeric(MC_AccountNo) And Len(GC_AccountNo) And IsNumeric(GC_AccountNo) = True Then
        Set dbfRecordset = CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO SBTC (BRSTN, ChkType, AccountNo , FormType,OrderQty,Batch, pkey) VALUES ('" _
            & BRSTN & "','MC','" & MC_AccountNo & "','20','1','" & Mid(FileName, 1, Len(FileName) - 4) & "'," & getPKey & ")"
        dbfRecordset.Open SQL, DBFConnector, 1, 1

        Set dbfRecordset = CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO SBTC (BRSTN, ChkType, AccountNo , FormType,OrderQty,Batch , pkey) VALUES ('" _
            & BRSTN & "','GC','" & GC_AccountNo & "','20','1','" & Mid(FileName, 1, Len(FileName) - 4) & "'," & getPKey & ")"
        dbfRecordset.Open SQL, DBFConnector, 1, 1
    End If
    
    LoopCount = LoopCount + 1
Loop



'Close All Excel Files
TaskID = Shell("taskkill.exe /f /t /im Excel.exe", vbHide)

hProcess = OpenProcess(SYNCHRONIZE, True, TaskID)
Call WaitForSingleObject(hProcess, WAIT_INFINITE)
CloseHandle hProcess
'End Close All Excel Files
End Sub


Sub ReadASCFiles(FileName)
If UCase(Right(FileName, 3)) = "XLS" Then
    ReadExcel (FileName)
    Exit Sub
End If


Batch = UCase(Mid(FileName, 1, Len(FileName) - 4))



'For Configuration
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
'End For Configuration




MeCaption = Me.Caption


LoopCount = 0


Close #1
Open App.Path & "\Head\" & FileName For Input As #1

Do Until EOF(1)
    Line Input #1, NotepadLine
    
    Pkey = Pkey + 1
    
    ChkType = Mid(NotepadLine, 1, 1)
    BRSTN = Mid(NotepadLine, 2, 9)
    AccountNo = Mid(NotepadLine, 12, 12)
    
    AccountName = Trim(Replace(Replace(Replace(UCase(Mid(NotepadLine, 24, 56)), "Ñ", "N"), "NO NAME", ""), "¥", "N"))
    ContCode = Trim(Mid(NotepadLine, 80, 1))
    FormType = Mid(NotepadLine, 81, 2)
    OrderQty = Mid(NotepadLine, 83, 2)
    Extension = Trim(Mid(NotepadLine, 85, 3))
    
    Me.Caption = "Saving file " & AccountNo & "    Line # " & LoopCount
    
    
    
    
    
    
    'A     05      PA Reg
    'B     16      CA Reg
    
    'B     20      MC      --> mid(AccountNo,5,3) = "211" ---> MC (not 212)
    'B     20      GC      --> mid(AccountNo,5,3) = "212" ---> GC
    
    'Y S E
    'A     05      PA PreEncoded --> AA
    'B     16      CA PreEncoded --> BB
    
    '13 CS
    'F     25      PA CheckOne
    'F     26      CA CheckOne
    
    
    '13 NB
    'E     23      PA CheckPower
    'E     22      CA CheckPower
    
    
    'Temp only
    'B      21  MC
    'B      22  GC
    
    
    
    
    

    
    If ((ChkType = "A" And FormType = "05") Or (ChkType = "B" And FormType = "16") Or (ChkType = "F" And FormType = "25") Or (ChkType = "F" And FormType = "26") Or (ChkType = "B" And FormType = "20" And Mid(AccountNo, 5, 3) <> "212") Or (ChkType = "B" And FormType = "20" And Mid(AccountNo, 5, 3) = "212") Or (ChkType = "E" And FormType = "22") Or (ChkType = "E" And FormType = "23") Or (ChkType = "B" And FormType = "21") Or (ChkType = "B" And FormType = "22")) _
        And Extension <> "CKR" And Extension <> "CKC" And Extension <> "CK1" Then
        
        
        If Mid(UCase(FileName), 1, 3) = "YSE" Then
            If ChkType = "A" And FormType = "05" Then ChkType = "AA"
            If ChkType = "B" And FormType = "16" Then ChkType = "BB"
        End If
        
        If (ChkType = "B" And FormType = "20" And Mid(AccountNo, 5, 3) = "212") Or _
           (ChkType = "B" And FormType = "20" And Mid(AccountNo, 1, 1) = "9" And Mid(AccountNo, 6, 7) = "2000022") Then
            ChkType = "GC"
            
            AccountName = ""
        End If
        
        If ChkType = "B" And FormType = "20" And Mid(AccountNo, 5, 3) <> "212" Then ChkType = "MC"

        
        
        
        
        'Temp only
        If ChkType = "B" And FormType = "21" Then
            FormType = "20"
            ChkType = "MC"
        End If
        
        If ChkType = "B" And FormType = "22" Then
            FormType = "20"
            ChkType = "GC"
        End If
        'End Temp only
        
        
        
        If ContCode = "" Or ContCode = "1" Then
            Set dbfRecordset = CreateObject("ADODB.Recordset")
            SQL = "INSERT INTO SBTC (BRSTN, ChkType, AccountNo , Name1 , FormType,OrderQty,Batch,Pkey,FileName) VALUES ('" _
                & BRSTN & "','" & ChkType & "','" & AccountNo & "','" & Replace(AccountName, "'", "''") & "','" & FormType & "','" & OrderQty & "','" & Batch & "'," & Pkey & ",'" & FileName & "')"
            dbfRecordset.Open SQL, DBFConnector, 1, 1
            
            If Val(OrderQty) >= 50 Then
                If MsgBox("The Order Qty of Account Number " & AccountNo & " on " & FileName & " is " & Val(OrderQty) & vbNewLine & "Are you sure you want to continue?", vbYesNo + vbInformation, "Confirm Continue") = vbNo Then End
            End If
        End If
        
        
        If ContCode = "2" Then
            Set dbfRecordset = CreateObject("ADODB.Recordset")
            SQL = "INSERT INTO Temp (AccountNo, AcctName,Batch,PKey, FileName) VALUES ('" & AccountNo & "','" & Replace(AccountName, "'", "''") & "','" & Batch & "'," & Pkey & ",'" & FileName & "')"
            dbfRecordset.Open SQL, DBFConnector, 1, 1
        End If
        
        
    Else
    
    
        If Extension <> "CKR" And Extension <> "CKC" And Extension <> "CK1" Then
            MsgBox "Unable to find the chequename of ChkType " & ChkType & " with FormType " & FormType & " on " & FileName & ". Account No: " & AccountNo, vbCritical, "Error"
            End
        End If
    End If
    
    LoopCount = LoopCount + 1
Loop

Close #1













'For Cont Code
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT AccountNo, AcctName, Batch FROM Temp WHERE FileName = '" & FileName & "' ORDER BY PKey"
dbfRecordset.Open SQL, DBFConnector, 1, 1




LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    AccountNo = dbfRecordset.Fields(0)
    
    If Len(dbfRecordset.Fields(1)) >= 1 Then
        AcctName = dbfRecordset.Fields(1)
    Else
        AcctName = ""
    End If

    Batch = dbfRecordset.Fields(2)
    
    

    
    
    
    
    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT PKey FROM SBTC WHERE FileName = '" & FileName & "' AND AccountNo = '" & AccountNo & "' AND Name2 IS NULL ORDER BY PKey"
    dbfRecordset1.Open SQL, DBFConnector, 1, 1
    
    
    
    
    If dbfRecordset1.RecordCount = 0 Then
        MsgBox "Error on Cont Code on Account No " & AccountNo, vbInformation, "Error"
        End
    Else
        Pkey = dbfRecordset1.Fields(0)
        
        Set dbfRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "UPDATE SBTC SET Name2 = '" & Replace(AcctName, "'", "''") & "' WHERE FileName = '" & FileName & "' AND AccountNo = '" & AccountNo & "' AND PKey = " & Pkey
        dbfRecordset1.Open SQL, DBFConnector, 1, 1
    End If
    
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
'End For Cont Code



Me.Caption = MeCaption
End Sub







Sub ReadWinZip()
Close #1
Open "C:\WinZip.txt" For Input As #1

Do Until EOF(1)
    Line Input #1, WinZipLocation
Loop

Close #1

If Dir$(WinZipLocation) = "" Then
    MsgBox "Winzip Location is Invalid", vbInformation, "Error"
    End
End If
End Sub




Private Sub cmdEncode_Click()
If DateValue(dteDeliveryDate.Value) = DateValue(Now) Then
    MsgBox "Please select Delivery Date", vbInformation, ""
    Exit Sub
End If

If MsgBox("Are you sure you want to encode?", vbYesNo + vbInformation, "Confirm Encode") = vbNo Then Exit Sub


CreateTable


frmEncode.Show
frmEncode.lblDeliveryDate.Caption = Format(dteDeliveryDate.Value, "Mmm. DD, YYYY")

Unload Me

End Sub



Private Sub dteDeliveryDate_Change()
'For MySQL
Dim Conn_SQL As ADODB.Connection

Set Conn_SQL = New ADODB.Connection
Conn_SQL.ConnectionString = "uid=cpc;pwd=CorpCaptive;server=" & Target_ip & ";driver={MySQL ODBC 5.1 Driver};database=captive_database;dsn=;"
Conn_SQL.Open
'End For MySQL



SQL = "SELECT COUNT(PrimaryKey) FROM Holidays WHERE Date = '" & Format(dteDeliveryDate.Value, "YYYY-MM-DD") & "'"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
Rs.Requery


If Val(Rs.Fields(0).Value) >= 1 Then
    MsgBox "Can't set delivery date on Holidays", vbCritical, "Error"
    dteDeliveryDate.Value = Now
End If


If Weekday(dteDeliveryDate.Value) = 1 Or Weekday(dteDeliveryDate.Value) = 7 Then
    MsgBox "Delivery Date can't be set on weekends", vbCritical, "Error"
    dteDeliveryDate.Value = Now
End If



End Sub




Private Sub Form_Load()
'Close All Excel Files
TaskID = Shell("taskkill.exe /f /t /im Excel.exe", vbHide)

hProcess = OpenProcess(SYNCHRONIZE, True, TaskID)
Call WaitForSingleObject(hProcess, WAIT_INFINITE)
CloseHandle hProcess
'End Close All Excel Files




GetSettings


ReadWinZip



fleHead.Path = App.Path & "\Head\"
fleHead.Pattern = "*.asc;*.xls;*.txt"



If fleHead.ListCount >= 1 Then
    cmdCheckFilesHead.Enabled = True
Else
    cmdCheckFilesHead.Enabled = False
End If




CheckHashTotal



TimeStart = Format(Now, "HH:MM")



Label1.Caption = "This will auto 'Rename' Files in Head Folder" & vbneline & vbNewLine & vbNewLine & "YSECPT0408.txt -> YSE13Digit_cp0408.txt" & vbNewLine & "NOVABU0408.txt --> 13Digit_nb0408.txt" & vbNewLine & "CONSOL0408.txt --> 13Digit_cs0408.txt" & vbNewLine & "CAPTIV0408.txt  --> 13Digit_cp0408.txt" & vbNewLine & "CPTIV0408.txt  --> 13Digit_cp0408.txt"



dteDeliveryDate.Value = Now


DateToday_Final = Format(Now, "YYYY-MM-DD")
TimeToday_Final = Format(Now, "HH:MM:SS")
End Sub





Sub CheckHashTotal()


'For MySQL
Dim Conn_SQL As ADODB.Connection

Set Conn_SQL = New ADODB.Connection
Conn_SQL.ConnectionString = "uid=cpc;pwd=CorpCaptive;server=" & Target_ip & ";driver={MySQL ODBC 5.1 Driver};database=captive_database;dsn=;"
Conn_SQL.Open
'End For MySQL




If CodesOnly = True Then DBase = "Master_Database_SBTC_Temp"
If CodesOnly = False Then DBase = "Master_Database_SBTC"




SQL = "SELECT DISTINCT(DeliveryDate) FROM " & DBase & " WHERE DeliveryDate <= '" & Format(Now, "YYYY-MM-DD") & "' AND HashSentDate IS NULL"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
Rs.Requery




ToolTip = ""



LoopCount = 0
Do Until LoopCount = Rs.RecordCount
    DeliveryDate = Rs.Fields(0)
    
    If ToolTip = "" Then
        ToolTip = DeliveryDate
    Else
        ToolTip = ToolTip & " " & DeliveryDate
    End If
    
    Rs.MoveNext
    LoopCount = LoopCount + 1
Loop


lblHashTotal.ToolTipText = ToolTip


If Rs.RecordCount >= 1 Then
    lblHashTotal.Caption = Rs.RecordCount & " Delivery Date/s hasn't been Sent Yet"
    
    lblHashTotal.Visible = True
Else
    lblHashTotal.Visible = False
End If

End Sub





Private Sub lblHashTotal_DblClick()
frmSendHashTotal.Show
End Sub
