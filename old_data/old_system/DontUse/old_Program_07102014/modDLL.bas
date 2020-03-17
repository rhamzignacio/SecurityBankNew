Attribute VB_Name = "modDLL"
Dim StoredText(0 To 999999) As String
Dim StoreTextTotal As String

Public Zip_Reg_13 As Boolean
Public Zip_Reg_10 As Boolean
Public Zip_MC As Boolean
Public Zip_GC As Boolean

Public ProcessBy As String
Public WinZipLocation As String
Public DateTimeToday As String

Function GenerateEmail(MyName, MyReason)
    Body = "Hello and Good Day," & vbNewLine _
            & vbNewLine _
            & "Note: Security Bank CheckOne has been run for more than ONCE Today " & Format(Now, "Mmm. DD, YYYY") & " " & Format(Now, "HH:MM:SS AMPM") & " due to the following Reason/s:" & vbNewLine _
            & vbNewLine _
            & "> > > Re-Run of System Requested by: " & MyName & vbNewLine _
            & "> > > Reason/s: " & MyReason & vbNewLine & vbNewLine _
            & vbNewLine _
            & "This is a System Generated Message . . ." & vbNewLine & vbNewLine _
            & vbNewLine _
            & "Best Regards," & vbNewLine _
            & vbNewLine _
            & "Captive Printing Corporation"
     
     Result = SendMail("Security Bank Report", Body)
     
End Function

Function SendMail(Subject, Body)
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & GetEmailClient & "; Jet OLEDB:Database Password=CorpCaptive;"
  .Open
End With

'Get Email Recipient
Set Rs = New ADODB.Recordset
SQL = "SELECT EmailAddress FROM Reports_Receipient"
Rs.Open SQL, Conn, adOpenStatic

EmailAddress = ""
LoopCount = 0
Do Until LoopCount = Rs.RecordCount
    Email = Rs.Fields(0)
    If EmailAddress = "" Then
        EmailAddress = Email
    Else
        Email = EmailAddress & "; " & Email
    End If
    
    Rs.MoveNext
    LoopCount = LoopCount + 1
Loop
'End Get Email Recipient

'Send the Email
Set Rs = New ADODB.Recordset
SQL = "INSERT INTO PendingEmail (EmailAddress,Subject,Body,Recieved_Date,Recieved_Time,Sent_Date,Sent_Time,Batch) VALUES ('" _
    & EmailAddress & "','" & Replace(Subject, "'", "''") & "','" & Replace(Body, "'", "''") & "','" & Format(Now, "MM/DD/YYYY") & "','" & Format(Now, "HH:MM:SS") & "','<Not Yet Sent>','<Not Yet Sent>','" & Format(Now, "MMDDYYHHMMSS") & "')"
Rs.Open SQL, Conn, adOpenStatic
'End Send the Email

End Function

Function GetEmailClient()
Close #1
Open App.Path & "\EmailClient.ini" For Input As #1
Do Until EOF(1)
    Line Input #1, GetEmailClient
Loop
Close #1
End Function

Sub ProcessOrdersToday()

LoopCount = 0
Do Until LoopCount = 3
    If LoopCount = 0 Then DataLocation = App.Path & "\SBTC.captive"
    If LoopCount = 1 Then DataLocation = App.Path & "\GIFTCHK\GIFTCHK.captive"
    If LoopCount = 2 Then DataLocation = App.Path & "\MC\MC.captive"
    
    'For Initialisation
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    
    Set Conn = New ADODB.Connection
    
    With Conn
      .CursorLocation = adUseClient
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & DataLocation & "; Jet OLEDB:Database Password=Elgae;"
      .Open
    End With
    'End For Initialisation
    
    DataCount = 1
    
RepeatMe:
    
    Set Rs = New ADODB.Recordset
    SQL = "SELECT [Date] FROM LastRUN WHERE [Date] = '" & Format(Now, "MM/DD/YYYY") & "_" & Val(DataCount) & "'"
    Rs.Open SQL, Conn, adOpenStatic
    
    If Rs.RecordCount >= 1 Then
        DataCount = DataCount + 1
        
        GoTo RepeatMe
        Exit Sub
    End If
    
    Set Rs = New ADODB.Recordset
    SQL = "UPDATE LastRUN SET [Date] = '" & Format(Now, "MM/DD/YYYY") & "_" & Val(DataCount) & "' WHERE [Date] = '" & Format(Now, "MM/DD/YYYY") & "'"
    Rs.Open SQL, Conn, adOpenStatic
    
    LoopCount = LoopCount + 1
Loop

End Sub


Function Open_SBTC_TextFile10(FileLocation, OutputDirectory, InputFile)
'Create Temp
If Dir$(App.Path & "\Temp10.dbf") <> "" Then Kill App.Path & "\Temp10.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE TEMP10 (CheckType Varchar(1), BRSTN Varchar(9), AccountNo Varchar(12), AccountNM Varchar(60), OrderQty Varchar(4), ContCode varchar(2))"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Create Temp

Close #1
Open FileLocation For Input As #1

Do Until EOF(1)
    Line Input #1, NotepadLine

    '10 DIGITS
    If Mid(InputFile, 1, 3) = "YSE" Then
        CheckType = Mid(NotepadLine, 1, 1)
        BRSTN = Mid(NotepadLine, 2, 9)
        AccountNumber = Mid(NotepadLine, 12, 12)
        AccountName = Mid(NotepadLine, 24, 56)
        ContCode = Val(Mid(NotepadLine, 80, 1))
        Orderqty = Mid(NotepadLine, 83, 2)
    End If

    Set DBFConnector = CreateObject("ADODB.Connection")

    DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector.CursorLocation = adUseClient
        
    Set dbfRecordset = CreateObject("ADODB.Recordset")
    SQL = "INSERT INTO TEMP10 (CheckType, BRSTN, AccountNo, AccountNM, OrderQty,ContCode) VALUES ('" _
        & CheckType & "','" & BRSTN & "','" & AccountNumber & "','" & Replace(AccountName, "'", "''") & "','" & Orderqty & "','" & ContCode & "')"
    dbfRecordset.Open SQL, DBFConnector, 1, 1
    
    frmMain.Caption = "Creating DBF..." & BRSTN

Loop
Close #1

'Create Temp
If Dir$(App.Path & "\SBTC_10.dbf") <> "" Then Kill App.Path & "\SBTC_10.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & OutputDirectory & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE SBTC_10 (CheckType Varchar(1), BRSTN Varchar(9), AccountNo Varchar(12), Name1 Varchar(57), Name2 Varchar(57), OrderQty Varchar(4),PrimaryKey Numeric)"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Create Temp

'Open Temp File
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT CheckType, BRSTN, AccountNo, AccountNM, OrderQty,ContCode FROM TEMP10"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'Open Temp File

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    CheckType = dbfRecordset.Fields(0)
    BRSTN = dbfRecordset.Fields(1)
    AccountNo = dbfRecordset.Fields(2)
    
    If Len(dbfRecordset.Fields(3)) >= 1 Then
        AccountNM = dbfRecordset.Fields(3)
    Else
        AccountNM = ""
    End If
    
    Orderqty = dbfRecordset.Fields(4)
    ContCode = Val(dbfRecordset.Fields(5))
    
    If Val(ContCode) = 0 Then
        Set DBFConnector1 = CreateObject("ADODB.Connection")

        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & OutputDirectory & "\;Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
            
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO SBTC_10 (CheckType, BRSTN, AccountNo, Name1, Name2, OrderQty, PrimaryKey) VALUES ('" _
            & CheckType & "','" & BRSTN & "','" & AccountNo & "','" & Replace(AccountNM, "'", "''") & "','" & "','" _
            & Orderqty & "','" & LoopCount + 1 & "')"
        DBFRecordset1.Open SQL, DBFConnector1, 1, 1
    End If
    
    If Val(ContCode) = 1 Then
        'Find The 2nd Account Name
        Set DBFConnector1 = CreateObject("ADODB.Connection")

        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
            
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "SELECT AccountNM FROM Temp10 WHERE ContCode = '2' AND AccountNo = '" & AccountNo & "'"
        DBFRecordset1.Open SQL, DBFConnector1, 1, 1
        
        If Len(DBFRecordset1.Fields(0)) >= 1 Then
            Name2 = DBFRecordset1.Fields(0)
        Else
            Name2 = ""
        End If
        DBFRecordset1.Close
        
        Set DBFConnector1 = CreateObject("ADODB.Connection")

        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & OutputDirectory & "\;Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
            
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        
        SQL = "INSERT INTO SBTC_10 (CheckType, BRSTN, AccountNo, Name1, Name2, OrderQty, PrimaryKey) VALUES ('" _
              & CheckType & "','" & BRSTN & "','" & AccountNo & "','" & Replace(AccountNM, "'", "''") & "','" _
              & Replace(Name2, "'", "''") & "','" & Orderqty & "','" & LoopCount + 1 & "')"
        DBFRecordset1.Open SQL, DBFConnector1, 1, 1
        'End Find The 2nd Account Name
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
    frmMain.Caption = FormatNumber((LoopCount / dbfRecordset.RecordCount) * 100, 5) & " % D o n e . . ."
Loop

frmMain.Caption = "SBTC"

If Dir$(App.Path & "\Temp10.dbf") <> "" Then Kill App.Path & "\Temp10.dbf"
End Function

Function Open_SBTC_TextFile13(FileLocation, OutputDirectory, InputFile)
'Create Temp
If Dir$(App.Path & "\Temp13.dbf") <> "" Then Kill App.Path & "\Temp13.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE TEMP13 (CheckType Varchar(1), BRSTN Varchar(9), AccountNo Varchar(12), AccountNM Varchar(57), OrderQty Varchar(4), ContCode varchar(1))"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Create Temp

Close #1
Open FileLocation For Input As #1

Do Until EOF(1)
    Line Input #1, NotepadLine
    
    '13 DIGITS
    If Mid(InputFile, 1, 2) = "13" Then
        CheckType = Mid(NotepadLine, 1, 1)
        BRSTN = Mid(NotepadLine, 2, 9)
        AccountNumber = Trim(Mid(NotepadLine, 12, 12))
        AccountName = Mid(NotepadLine, 24, 56)
        ContCode = Val(Mid(NotepadLine, 80, 1))
        Orderqty = Mid(NotepadLine, 83, 4)
    
        If Len(AccountNumber) <> 12 Then
            frmMain.lstErrors.AddItem ("Account Number " & AccountNumber & " is Invalid")
'            MsgBox AccountNumber
        End If
        
    End If
    

    
    Set DBFConnector = CreateObject("ADODB.Connection")

    DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector.CursorLocation = adUseClient
        
    Set dbfRecordset = CreateObject("ADODB.Recordset")
    SQL = "INSERT INTO TEMP13 (CheckType, BRSTN, AccountNo, AccountNM, OrderQty,ContCode) VALUES ('" _
        & CheckType & "','" & BRSTN & "','" & AccountNumber & "','" & Replace(AccountName, "'", "''") & "','" & Orderqty & "','" & ContCode & "')"
    dbfRecordset.Open SQL, DBFConnector, 1, 1
    
    frmMain.Caption = "Creating DBF..." & BRSTN

Loop
Close #1

'Create Temp
If Dir$(App.Path & "\SBTC_13.dbf") <> "" Then Kill App.Path & "\SBTC_13.dbf"

Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & OutputDirectory & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "CREATE TABLE SBTC_13 (CheckType Varchar(1), BRSTN Varchar(9), AccountNo Varchar(12), Name1 Varchar(57), Name2 Varchar(57), OrderQty Varchar(4),PrimaryKey Numeric)"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'End Create Temp

'Open Temp File
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT CheckType, BRSTN, AccountNo, AccountNM, OrderQty,ContCode FROM TEMP13"
dbfRecordset.Open SQL, DBFConnector, 1, 1
'Open Temp File

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    CheckType = dbfRecordset.Fields(0)
    BRSTN = dbfRecordset.Fields(1)
    AccountNo = dbfRecordset.Fields(2)
    
    If Len(dbfRecordset.Fields(3)) >= 1 Then
        AccountNM = dbfRecordset.Fields(3)
    Else
        AccountNM = ""
    End If
    
    Orderqty = dbfRecordset.Fields(4)
    ContCode = Val(dbfRecordset.Fields(5))
    
    If Val(ContCode) = 0 Then
        Set DBFConnector1 = CreateObject("ADODB.Connection")

        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & OutputDirectory & "\;Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
            
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO SBTC_13 (CheckType, BRSTN, AccountNo, Name1, Name2, OrderQty, PrimaryKey) VALUES ('" _
            & CheckType & "','" & BRSTN & "','" & AccountNo & "','" & Replace(AccountNM, "'", "''") & "','" & "','" _
            & Orderqty & "','" & LoopCount + 1 & "')"
        DBFRecordset1.Open SQL, DBFConnector1, 1, 1
    End If
    
    If Val(ContCode) = 1 Then
        'Find The 2nd Account Name
        Set DBFConnector1 = CreateObject("ADODB.Connection")

        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
            
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "SELECT AccountNM FROM Temp13 WHERE ContCode = '2' AND AccountNo = '" & AccountNo & "'"
        DBFRecordset1.Open SQL, DBFConnector1, 1, 1
        
        If Len(DBFRecordset1.Fields(0)) >= 1 Then
            Name2 = DBFRecordset1.Fields(0)
        Else
            Name2 = ""
        End If
        DBFRecordset1.Close
        
        Set DBFConnector1 = CreateObject("ADODB.Connection")

        DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & OutputDirectory & "\;Extended properties=dBase III"
        DBFConnector1.CursorLocation = adUseClient
            
        Set DBFRecordset1 = CreateObject("ADODB.Recordset")
        
        SQL = "INSERT INTO SBTC_13 (CheckType, BRSTN, AccountNo, Name1, Name2, OrderQty, PrimaryKey) VALUES ('" _
              & CheckType & "','" & BRSTN & "','" & AccountNo & "','" & Replace(AccountNM, "'", "''") & "','" _
              & Replace(Name2, "'", "''") & "','" & Orderqty & "','" & LoopCount + 1 & "')"
        DBFRecordset1.Open SQL, DBFConnector1, 1, 1
        'End Find The 2nd Account Name
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
    
    frmMain.Caption = "Creating DBF..." & BRSTN
Loop

If Dir$(App.Path & "\Temp13.dbf") <> "" Then Kill App.Path & "\Temp13.dbf"
End Function

Function getNewPrimaryKey()
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient
    
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT Max(PrimaryKey) FROM SBTCNEW"
dbfRecordset.Open SQL, DBFConnector, 1, 1

If Len(dbfRecordset.Fields(0)) >= 1 Then
    getNewPrimaryKey = dbfRecordset.Fields(0) + 1
Else
    getNewPrimaryKey = "1"
End If
End Function

Function ProgramAlreadyRunReg() '--> REGULAR!
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "SELECT Batch FROM LastRun WHERE Date = '" & Format(Now, "MM/DD/YYYY") & "' ORDER BY PrimaryKey DESC"

Set Rs = New ADODB.Recordset

With Rs

Set .ActiveConnection = Conn
    .CursorType = adOpenStatic
    .Source = strQuery
    .Open
End With

If Rs.RecordCount >= 1 Then
    ProgramAlreadyRunReg = Rs.Fields(0)
Else
    ProgramAlreadyRunReg = ""
End If
End Function

Function CheckIfRefIsModifiedReg() '--> REGULAR!

Dim fso As New FileSystemObject

'Get the Last Modified on Ref.dbf
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")
'End Get the Last Modified on Ref.dbf

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "SELECT [Date],[Time] FROM LastRefModified"

Set Rs = New ADODB.Recordset

With Rs

Set .ActiveConnection = Conn
    .CursorType = adOpenStatic
    .Source = strQuery
    .Open
End With

SaveDate = Rs.Fields(0)
SaveTime = Rs.Fields(1)

If ExistingRefDateModified = SaveDate And ExistingRefTimeModified = SaveTime Then
    CheckIfRefIsModifiedReg = True
Else
    CheckIfRefIsModifiedReg = False
End If
End Function

Function ProgramAlreadyRunMC() '--> MC!
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\MC\MC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "SELECT Batch FROM LastRun WHERE Date = '" & Format(Now, "MM/DD/YYYY") & "' ORDER BY PrimaryKey DESC"

Set Rs = New ADODB.Recordset

With Rs

Set .ActiveConnection = Conn
    .CursorType = adOpenStatic
    .Source = strQuery
    .Open
End With

If Rs.RecordCount >= 1 Then
    ProgramAlreadyRunMC = Rs.Fields(0)
Else
    ProgramAlreadyRunMC = ""
End If
End Function

Function CheckIfRefIsModifiedMC() '--> MC!

Dim fso As New FileSystemObject

'Get the Last Modified on Ref.dbf
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\MC\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")
'End Get the Last Modified on Ref.dbf

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\MC\MC.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "SELECT [Date],[Time] FROM LastRefModified"

Set Rs = New ADODB.Recordset

With Rs

Set .ActiveConnection = Conn
    .CursorType = adOpenStatic
    .Source = strQuery
    .Open
End With

SaveDate = Rs.Fields(0)
SaveTime = Rs.Fields(1)

If ExistingRefDateModified = SaveDate And ExistingRefTimeModified = SaveTime Then
    CheckIfRefIsModifiedMC = True
Else
    CheckIfRefIsModifiedMC = False
End If
End Function

Function ProgramAlreadyRunGIFTCHK() '--> GIFTCHK!
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\GIFTCHK\GIFTCHK.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "SELECT Batch FROM LastRun WHERE Date = '" & Format(Now, "MM/DD/YYYY") & "' ORDER BY PrimaryKey DESC"

Set Rs = New ADODB.Recordset

With Rs

Set .ActiveConnection = Conn
    .CursorType = adOpenStatic
    .Source = strQuery
    .Open
End With

If Rs.RecordCount >= 1 Then
    ProgramAlreadyRunGIFTCHK = Rs.Fields(0)
Else
    ProgramAlreadyRunGIFTCHK = ""
End If
End Function

Function CheckIfRefIsModifiedGIFTCHK() '--> GIFTCHK!

Dim fso As New FileSystemObject

'Get the Last Modified on Ref.dbf
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(App.Path & "\GIFTCHK\Ref.dbf")

ExistingRefDateModified = Format(f.DateLastModified, "MM/DD/YYYY")
ExistingRefTimeModified = Format(f.DateLastModified, "HH:MM:SS")
'End Get the Last Modified on Ref.dbf

Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & App.Path & "\GIFTCHK\GIFTCHK.captive; Jet OLEDB:Database Password=Elgae;"
  .Open
End With

strQuery = "SELECT [Date],[Time] FROM LastRefModified"

Set Rs = New ADODB.Recordset

With Rs

Set .ActiveConnection = Conn
    .CursorType = adOpenStatic
    .Source = strQuery
    .Open
End With

SaveDate = Rs.Fields(0)
SaveTime = Rs.Fields(1)

If ExistingRefDateModified = SaveDate And ExistingRefTimeModified = SaveTime Then
    CheckIfRefIsModifiedGIFTCHK = True
Else
    CheckIfRefIsModifiedGIFTCHK = False
End If
End Function

Function GetSourceLocation()
Close #1
Open App.Path & "\SourceLocation.ini" For Input As #1
Do Until EOF(1)
    Line Input #1, GetSourceLocation
Loop
Close #1
End Function
