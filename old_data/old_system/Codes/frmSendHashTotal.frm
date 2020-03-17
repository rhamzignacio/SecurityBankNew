VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSendHashTotal 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Hash Total"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   Icon            =   "frmSendHashTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSendHashTotal 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Send Hash Total"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Generate Hash Total"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dteDeliveryDate 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97976320
      CurrentDate     =   42741
   End
   Begin VB.Label lblSummaryChequeName 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2415
      Left            =   3480
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label lblSummaryBatch 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "Delivery Date:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmSendHashTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerate_Click()


'For MySQL
Dim Conn_SQL As ADODB.Connection

Set Conn_SQL = New ADODB.Connection
Conn_SQL.ConnectionString = "uid=cpc;pwd=CorpCaptive;server=" & Target_ip & ";driver={MySQL ODBC 5.1 Driver};database=captive_database;dsn=;"
Conn_SQL.Open
'End For MySQL


If CodesOnly = True Then DBase = "Master_Database_SBTC_Temp"
If CodesOnly = False Then DBase = "Master_Database_SBTC"

SQL = "SELECT Batch, COUNT(PrimaryKey) FROM " & DBase & " WHERE DeliveryDate = '" & Format(dteDeliveryDate.Value, "YYYY-MM-DD") & "' GROUP BY Batch"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
Rs.Requery

lblSummaryBatch.Caption = ""
TotalBatch = 0

LoopCount = 0
Do Until LoopCount = Rs.RecordCount
    Batch = Rs.Fields(0)
    Qty = Rs.Fields(1)
    
    
    Do Until Len(Batch) >= 15
        Batch = Batch & " "
    Loop
    
    
    If lblSummaryBatch.Caption = "" Then
        lblSummaryBatch.Caption = Batch & Qty
    Else
        lblSummaryBatch.Caption = lblSummaryBatch.Caption & vbNewLine & Batch & Qty
    End If
    TotalBatch = TotalBatch + Qty
    
    Rs.MoveNext
    LoopCount = LoopCount + 1
Loop

lblSummaryBatch.Caption = lblSummaryBatch.Caption & vbNewLine & vbNewLine & "Total: " & TotalBatch





If CodesOnly = True Then DBase = "Master_Database_SBTC_Temp"
If CodesOnly = False Then DBase = "Master_Database_SBTC"

SQL = "SELECT ChequeName, COUNT(PrimaryKey) FROM " & DBase & " WHERE DeliveryDate = '" & Format(dteDeliveryDate.Value, "YYYY-MM-DD") & "' GROUP BY ChequeName"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
Rs.Requery


lblSummaryChequeName.Caption = ""
TotalBatch = 0

LoopCount = 0
Do Until LoopCount = Rs.RecordCount
    ChequeName = Rs.Fields(0)
    Qty = Rs.Fields(1)
    
    
    Do Until Len(ChequeName) >= 25
        ChequeName = ChequeName & " "
    Loop
    
    
    If lblSummaryChequeName.Caption = "" Then
        lblSummaryChequeName.Caption = ChequeName & Qty
    Else
        lblSummaryChequeName.Caption = lblSummaryChequeName.Caption & vbNewLine & ChequeName & Qty
    End If
    
    TotalBatch = TotalBatch + Qty
    
    Rs.MoveNext
    LoopCount = LoopCount + 1
Loop

lblSummaryChequeName.Caption = lblSummaryChequeName.Caption & vbNewLine & vbNewLine & "Total: " & TotalBatch



If Val(TotalBatch) >= 1 Then
    cmdSendHashTotal.Enabled = True
Else
    cmdSendHashTotal.Enabled = False
End If

MsgBox "Hash Total has been Generated", vbInformation, ""

End Sub

Private Sub cmdSendHashTotal_Click()
'On Error GoTo Err

'If Dir(App.Path & "\HashTotal\" & Format(dteDeliveryDate.Value, "MMDDYYYY") & ".xls") <> "" Then Kill App.Path & "\HashTotal\" & Format(dteDeliveryDate.Value, "MMDDYYYY") & ".xls"
'FileCopy App.Path & "\HashTotal_Source.xls", App.Path & "\HashTotal\" & Format(dteDeliveryDate.Value, "MMDDYYYY") & ".xls"






'For MySQL
Dim Conn_SQL As ADODB.Connection

Set Conn_SQL = New ADODB.Connection
Conn_SQL.ConnectionString = "uid=cpc;pwd=CorpCaptive;server=" & Target_ip & ";driver={MySQL ODBC 5.1 Driver};database=captive_database;dsn=;"
Conn_SQL.Open
'End For MySQL












'For Output in Excel Format
'Dim wbkmvr As Excel.Workbook
'Dim xlapp As Excel.Application
'Dim wksline As Excel.Worksheet

'Set xlapp = Excel.Application

'xlapp.Workbooks.Open (App.Path & "\HashTotal\" & Format(dteDeliveryDate.Value, "MMDDYYYY") & ".xls")
'Set wbkmvr = xlapp.ActiveWorkbook
'Set wksline = wbkmvr.Sheets(1)
'End For Output in Excel Format
    

If CodesOnly = True Then DBase = "Master_Database_SBTC_Temp"
If CodesOnly = False Then DBase = "Master_Database_SBTC"


'Get the Max
SQL = "SELECT MAX(LENGTH(ChequeName)) , MAX(LENGTH(Batch)) , MAX(LENGTH(BRSTN)) , MAX(LENGTH(Address1)) , MAX(LENGTH(AccountNo)) , MAX(LENGTH(Name1)) , MAX(LENGTH(Name2)) , MAX(LENGTH(StartingSerial)) , MAX(LENGTH(EndingSerial)) FROM " & DBase & " WHERE DeliveryDate = '" & Format(dteDeliveryDate.Value, "YYYY-MM-DD") & "'"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
Rs.Requery

Max_ChequeName = Rs.Fields(0)
Max_Batch = Rs.Fields(1)
Max_BRSTN = Rs.Fields(2)
Max_Address1 = Rs.Fields(3)
Max_AccountNo = Rs.Fields(4)
Max_Name1 = Rs.Fields(5)
Max_Name2 = Rs.Fields(6)
Max_StartingSerial = Rs.Fields(7)
Max_EndingSerial = Rs.Fields(8)
'End Get the Max



SQL = "SELECT ChequeName, Batch, BRSTN, AccountNo, Name1, Name2 , StartingSerial , EndingSerial , Address1 FROM " & DBase & " WHERE DeliveryDate = '" & Format(dteDeliveryDate.Value, "YYYY-MM-DD") & "' ORDER BY ChequeName, BRSTN, AccountNo, StartingSerial"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
Rs.Requery



Close #1
Open App.Path & "\HashTotal\" & Format(dteDeliveryDate.Value, "MMDDYYYY") & ".txt" For Output As #1


'wksline.Cells(2, 1) = "For Delivery Date: " & Format(dteDeliveryDate.Value, "Mmm. DD, YYYY")
Print #1, "For Delivery Date: " & Format(dteDeliveryDate.Value, "Mmm. DD, YYYY")
Print #1, ""


LoopCount = 0
Do Until LoopCount = Rs.RecordCount
    ChequeName = Rs.Fields(0)
    Batch = Rs.Fields(1)
    BRSTN = Rs.Fields(2)
    AccountNo = Rs.Fields(3)
    Name1 = Rs.Fields(4)
    Name2 = Rs.Fields(5)
    StartingSerial = Rs.Fields(6)
    EndingSerial = Rs.Fields(7)
    Address1 = Rs.Fields(8)
    
    
    
    
    Do Until Len(ChequeName) >= Max_ChequeName + 5
        ChequeName = ChequeName & " "
    Loop
    
    Do Until Len(Batch) >= Max_Batch + 5
        Batch = Batch & " "
    Loop
    
    Do Until Len(BRSTN) >= Max_BRSTN + 5
        BRSTN = BRSTN & " "
    Loop
        
    Do Until Len(AccountNo) >= Max_AccountNo + 5
        AccountNo = AccountNo & " "
    Loop
    
    Name1_2 = Name1 & " " & Name2
    Do Until Len(Name1_2) >= Max_Name1 + Max_Name2 + 5
        Name1_2 = Name1_2 & " "
    Loop
    
    Do Until Len(StartingSerial) >= Max_StartingSerial + 5
        StartingSerial = StartingSerial & " "
    Loop
    
    Do Until Len(EndingSerial) >= Max_EndingSerial + 5
        EndingSerial = EndingSerial & " "
    Loop
    
    Print #1, ChequeName & Batch & BRSTN & AccountNo & Name1_2 & StartingSerial & EndingSerial & Address1
    
    Rs.MoveNext
    LoopCount = LoopCount + 1
Loop
Close #1


'wbkmvr.Save
'wbkmvr.Close




Subject_Email = "SBTC Hash Total for Delivery Date " & Format(Now, "Mmm. DD, YYYY")

Heading = "  Hello and Good Day," _
        & vbNewLine _
        & vbNewLine _
        & vbNewLine _
        & "     Kindly see the attached file for the Hash Total." _
        & vbNewLine _
        & vbNewLine _
        & vbNewLine _
        & "     Orders as of this Batch:" _
        & vbNewLine _
        & vbNewLine _
        & vbNewLine _
        & lblSummaryBatch.Caption _
        & vbNewLine _
        & vbNewLine _
        & vbNewLine _
        & lblSummaryChequeName.Caption

Footer = vbNewLine _
       & vbNewLine _
       & vbNewLine _
       & vbNewLine _
       & "     This is a System Generated Message." & vbNewLine _
       & vbNewLine _
       & vbNewLine _
       & vbNewLine _
       & vbNewLine _
       & "     Thanks and Best Regards," _
       & vbNewLine _
       & vbNewLine _
       & vbNewLine _
       & "     Captive Printing Corporation    "
        



Body = Heading & vbNewLine & Footer








'Add Save Attachment
If CodesOnly = True Then
    Recipient = "orders@captiveprinting.com.ph"
Else
    Recipient = "gsdpurchasing7@securitybank.com.ph,gsdpurchasing5@securitybank.com.ph,gsdpurchasing2@securitybank.com.ph,ctimusan@securitybank.com.ph,rmenguito@securitybank.com.ph,virtualsupport@securitybank.com.ph,GSDPurchasing4@securitybank.com.ph,orders@captiveprinting.com.ph,cpc_services@captiveprinting.com.ph"
End If



SQL = "SELECT MAX(PrimaryKey) FROM Emails"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic

PrimaryKey_Email = Rs.Fields(0).Value + 1

SQL = "INSERT INTO Emails (Bank , Recipient_Email , Subject , Body , DateRequest , TimeRequest , Status , PrimaryKey , source_email) VALUES ('" _
    & "SBTC','" & Recipient & "','" & Subject_Email & "','" & Replace(Body, "'", "''") & "','" & Format(Now, "YYYY-MM-DD") & "','" & Format(Now, "HH:MM:SS") & "','Received'," & PrimaryKey_Email & ",'orders@captiveprinting.com.ph')"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic





Dim mystream As ADODB.Stream
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
mystream.LoadFromFile App.Path & "\HashTotal\" & Format(dteDeliveryDate.Value, "MMDDYYYY") & ".txt"



SQL = "INSERT INTO Emails_Blob (PrimaryKey_Source, FileName) VALUES (" _
    & PrimaryKey_Email & ",'" & Format(dteDeliveryDate.Value, "MMDDYYYY") & ".txt')"
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic



SQL = "SELECT * FROM Emails_Blob WHERE PrimaryKey_Source = " & PrimaryKey_Email
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic

Rs!Attachment = mystream.Read
Rs.Update
'End Save Attachment




'Check until send
RepeatMe:

SQL = "SELECT Status,ErrorMessage FROM Emails WHERE PrimaryKey = " & PrimaryKey_Email
Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic


Status = Rs.Fields(0)


If Status = "Received" Then
    GoTo RepeatMe
    Exit Sub
End If

If Status = "Sent" Then
    If CodesOnly = True Then DBase = "Master_Database_SBTC_Temp"
    If CodesOnly = False Then DBase = "Master_Database_SBTC"

    SQL = "UPDATE " & DBase & " SET HashSentDate = '" & Format(Now, "YYYY-MM-DD") & "', HashSentTime = '" & Format(Now, "HH:MM:SS") & "' WHERE DeliveryDate = '" & Format(dteDeliveryDate.Value, "YYYY-MM-DD") & "'"
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
    
    MsgBox "Hash Total has been Sent", vbInformation, ""
    End
End If

If Status = "Failed" Then
    MsgBox Rs.Fields(1), vbCritical, "Error"
    End
End If
'End Check until send





'Exit Sub
'err:
'
'    MsgBox err.Number & vbNewLine & err.Description
'
End Sub

Private Sub Form_Load()
dteDeliveryDate.Value = Now
End Sub
