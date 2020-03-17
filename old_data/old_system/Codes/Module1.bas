Attribute VB_Name = "Module1"
Public DateTimeToday As String
Public WinZipLocation As String

Public TimeStart As String
Public ProcessBy As String

Public Basetock_Verify As String


Public Const WAIT_INFINITE = -1&
Public Const SYNCHRONIZE = &H100000

Public Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long
   
Public Declare Function WaitForSingleObject Lib "kernel32" _
  (ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long
   
Public Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long
  
Public Pkey
Public DeliveryDate As String
Public Target_ip As String
Public DateToday_Final As String
Public TimeToday_Final As String
Public CodesOnly As Boolean

Public Resting_Folder  As String
Public PrinterFiles_Folder  As String





Function getPKey()
Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM SBTC"
dbfRecordset.Open SQL, DBFConnector, 1, 1

getPKey = dbfRecordset.RecordCount + 1
End Function






Function getChequeName(ChkType, FormType)
If ChkType = "A" And FormType = "05" Then getChequeName = "Regular Personal"
If ChkType = "B" And FormType = "16" Then getChequeName = "Regular Commercial"

If ChkType = "MC" And FormType = "20" Then getChequeName = "Manager's Checks"
If ChkType = "GC" And FormType = "20" Then getChequeName = "Gift Check"

If ChkType = "AA" And FormType = "05" Then getChequeName = "Personal Pre-Encoded"
If ChkType = "BB" And FormType = "16" Then getChequeName = "Commercial Pre-Encoded"

If ChkType = "F" And FormType = "25" Then getChequeName = "CheckOne Personal"
If ChkType = "F" And FormType = "26" Then getChequeName = "CheckOne Commercial"

If ChkType = "E" And FormType = "23" Then getChequeName = "CheckPower Personal"
If ChkType = "E" And FormType = "22" Then getChequeName = "CheckPower Commercial"

If ChkType = "GC" And FormType = "20" Then getChequeName = "Gift Check"

If ChkType = "CUSTOM" And FormType = "00" Then getChequeName = "Customized Checks"
If ChkType = "MC_1" And FormType = "00" Then getChequeName = "Manager's Check Continues"
End Function




Sub GetSettings()


Temp = UCase(App.Path)
Temp2 = ""


LoopCount = 0
Do Until LoopCount = Len(Temp)
    If Mid(Temp, LoopCount + 1, 4) = "AUTO" Then
        Temp2 = Mid(Temp, 1, LoopCount)
    End If

    If Mid(Temp, LoopCount + 1, 5) = "CODES" Then
        CodesOnly = True
    End If
    
    LoopCount = LoopCount + 1
Loop


Close #1
Open Temp2 & "\Auto\Settings.ini" For Input As #1



TotalDataLine = 0



LoopCount = 0
Do Until EOF(1)
    Line Input #1, LineInputData
    
    LoopCount = LoopCount + 1

    If LoopCount = 1 Then Target_ip = LineInputData
    If LoopCount = 2 Then Resting_Folder = LineInputData
    If LoopCount = 3 Then PrinterFiles_Folder = LineInputData
Loop
Close #1



DateToday_Final = Format(Now, "YYYY-MM-DD")
TimeToday_Final = Format(Now, "HH:MM:SS")
End Sub





Function ProcessMe(ChkType, FormType, FolderName, DeleteDBF_Value, FinalBatch, DeliveryDate)
    If ChkType = "A" And FormType = "05" Then
        PcsPerBook = 50
        ChkType2 = "P"
        Description = "PERSONAL"
        FormatSerial = "0000000"
        MICRLine = "     ONNNNNNNO"
        Ref_Location = App.Path & "\Regular\"
        RefChkType = "A"
        FileName = Mid(FinalBatch, 1, 4) & "_P12" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "A"
        ChkType_2 = "B"
        FormType_1 = "05"
        FormType_2 = "16"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    
    If ChkType = "B" And FormType = "16" Then
        PcsPerBook = 100
        ChkType2 = "C"
        Description = "COMMERCIAL"
        FormatSerial = "0000000000"
        MICRLine = "  ONNNNNNNNNNO"
        Ref_Location = App.Path & "\Regular\"
        RefChkType = "B"
        FileName = Mid(FinalBatch, 1, 4) & "_C12" & Mid(FinalBatch, 9, Len(FinalBatch))
    
        ChkType_1 = "A"
        ChkType_2 = "B"
        FormType_1 = "05"
        FormType_2 = "16"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\" & Format(Now, "YYYY") & "\"
        End If
    End If

    
    If ChkType = "AA" And FormType = "05" Then
        PcsPerBook = 50
        ChkType2 = "P"
        Description = "PERSONAL"
        FormatSerial = "0000000"
        MICRLine = "     ONNNNNNNO"
        Ref_Location = App.Path & "\Regular\"
        RefChkType = "A"
        
        FileName = "SB" & Mid(FinalBatch, 1, 4) & "P" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "AA"
        ChkType_2 = "BB"
        FormType_1 = "05"
        FormType_2 = "16"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\" & Format(Now, "YYYY") & "\"
        End If
    End If

    
    If ChkType = "BB" And FormType = "16" Then
        PcsPerBook = 100
        ChkType2 = "C"
        Description = "COMMERCIAL"
        FormatSerial = "0000000000"
        MICRLine = "  ONNNNNNNNNNO"
        Ref_Location = App.Path & "\Regular\"
        RefChkType = "B"
        FileName = "SB" & Mid(FinalBatch, 1, 4) & "C" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "AA"
        ChkType_2 = "BB"
        FormType_1 = "05"
        FormType_2 = "16"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    
    If ChkType = "MC" And FormType = "20" Then
        PcsPerBook = 50
        ChkType2 = "P"
        Description = "MANAGER'S CHECK"
        FormatSerial = "0000000000"
        MICRLine = "  ONNNNNNNNNNO"
        Ref_Location = App.Path & "\MC\"
        RefChkType = "A"
        FileName = "MC" & Mid(FinalBatch, 1, 4) & "P" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "MC"
        ChkType_2 = "MC"
        FormType_1 = "20"
        FormType_2 = "20"
        ChkType_31 = "MC"
        ChkType_32 = "MC"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\MC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\MC\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\MC\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\MC\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    
    If ChkType = "MC_1" And FormType = "00" Then
        PcsPerBook = 100
        ChkType2 = "C"
        Description = "MANAGER'S CHECK CONTINUES"
        FormatSerial = "0000000000"
        MICRLine = "  ONNNNNNNNNNO"
        Ref_Location = App.Path & "\MC\Continues"
        RefChkType = "B"
        
        If Mid(FinalBatch, 1, 2) = "MC" Then
            FileName = "MCC" & Mid(FinalBatch, 3, Len(FinalBatch))
        Else
            FileName = "MCC" & Mid(FinalBatch, 1, 4) & Mid(FinalBatch, 9, Len(FinalBatch))
        End If
        
        ChkType_1 = "MC_1"
        ChkType_2 = "MC_1"
        FormType_1 = "00"
        FormType_2 = "00"
        ChkType_31 = "MC_1"
        ChkType_32 = "MC_1"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\MC\CONTINUOUS\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\CONTINUOUS_MC\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\MC\CONTINUOUS\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\CONTINUOUS_MC\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    If ChkType = "CUSTOM" And FormType = "00" Then
        PcsPerBook = 100
        ChkType2 = "C"
        Description = "CUSTOMIZED CHECKS"
        FormatSerial = "0000000000"
        MICRLine = "  ONNNNNNNNNNO"
        Ref_Location = App.Path & "\Customized"
        RefChkType = "B"
        
        FileName = "CUS" & Mid(FinalBatch, 1, 4) & "C" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        
        ChkType_1 = "CUSTOM"
        ChkType_2 = "CUSTOM"
        FormType_1 = "00"
        FormType_2 = "00"
        ChkType_31 = "CUSTOM"
        ChkType_32 = "CUSTOM"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\CUSTOM\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\Customized\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\CUSTOM\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\Customized\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    If ChkType = "F" And FormType = "25" Then
        PcsPerBook = 25
        ChkType2 = "P"
        Description = "PERSONAL CHECKONE"
        FormatSerial = "0000000"
        MICRLine = "     ONNNNNNNO"
        Ref_Location = App.Path & "\CheckOne\"
        RefChkType = "A"
        FileName = "13D" & Mid(FinalBatch, 1, 4) & "P" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "F"
        ChkType_2 = "F"
        FormType_1 = "25"
        FormType_2 = "26"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    
    If ChkType = "F" And FormType = "26" Then
        PcsPerBook = 50
        ChkType2 = "C"
        Description = "COMMERCIAL CHECKONE"
        FormatSerial = "0000000000"
        MICRLine = "  ONNNNNNNNNNO"
        Ref_Location = App.Path & "\CheckOne\"
        RefChkType = "B"
        FileName = "13D" & Mid(FinalBatch, 1, 4) & "C" & Mid(FinalBatch, 9, Len(FinalBatch))
    
        ChkType_1 = "F"
        ChkType_2 = "F"
        FormType_1 = "25"
        FormType_2 = "26"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
    
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\CheckOne\" & Format(Now, "YYYY") & "\"
        End If
    End If

    
    If ChkType = "E" And FormType = "23" Then
        PcsPerBook = 50
        ChkType2 = "P"
        Description = "PERSONAL CHECKPOWER"
        FormatSerial = "0000000"
        MICRLine = "     ONNNNNNNO"
        Ref_Location = App.Path & "\CheckPower\"
        RefChkType = "A"
        FileName = "CKP" & Mid(FinalBatch, 1, 4) & "P" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "E"
        ChkType_2 = "E"
        FormType_1 = "23"
        FormType_2 = "22"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\CheckPower\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\CKPOWER\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\CheckPower\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\CKPOWER\" & Format(Now, "YYYY") & "\"
        End If
    End If

    
    If ChkType = "E" And FormType = "22" Then
        PcsPerBook = 100
        ChkType2 = "C"
        Description = "COMMERCIAL CHECKPOWER"
        FormatSerial = "0000000000"
        MICRLine = "  ONNNNNNNNNNO"
        Ref_Location = App.Path & "\CheckPower\"
        RefChkType = "B"
        FileName = "CKP" & Mid(FinalBatch, 1, 4) & "C" & Mid(FinalBatch, 9, Len(FinalBatch))
    
        ChkType_1 = "E"
        ChkType_2 = "E"
        FormType_1 = "23"
        FormType_2 = "22"
        ChkType_31 = "PA"
        ChkType_32 = "CA"
        
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\CheckPower\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\CKPOWER\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\CheckPower\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\CKPOWER\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    
    If ChkType = "GC" And FormType = "20" Then
        PcsPerBook = 50
        ChkType2 = "P"
        Description = "GIFT CHECK"
        FormatSerial = "000000"
        MICRLine = "  O0000NNNNNNO"
        Ref_Location = App.Path & "\GiftCheck\"
        RefChkType = "A"
        FileName = "GC" & Mid(FinalBatch, 1, 4) & "P" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "GC"
        ChkType_2 = "GC"
        FormType_1 = "20"
        FormType_2 = "20"
        ChkType_31 = "GC"
        ChkType_32 = "GC"
    
        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\GiftCheck\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\GC\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\GiftCheck\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\GC\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    
    If ChkType = "CS" Then
        PcsPerBook = 50
        ChkType2 = "P"
        Description = "CHARGE SLIP"
        FormatSerial = "0000000000"
        MICRLine = "  O0000NNNNNNO"
        Ref_Location = App.Path & "\Charge_Slip\"
        RefChkType = "A"
        FileName = "CS" & Mid(FinalBatch, 1, 4) & "P" & Mid(FinalBatch, 9, Len(FinalBatch))
        
        ChkType_1 = "CS"
        ChkType_2 = "CS"
        FormType_1 = "00"
        FormType_2 = "00"
        ChkType_31 = "CS"
        ChkType_32 = "CS"

        If CodesOnly = True Then
            Temp_DriveR = PrinterFiles_Folder & "\Codes\SBTC\Charge_Slip\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\Codes\SBTC\Charge_Slip\" & Format(Now, "YYYY") & "\"
        End If
        
        If CodesOnly = False Then
            Temp_DriveR = PrinterFiles_Folder & "\SBTC\Charge_Slip\" & Format(Now, "YYYY") & "\"
            Temp_CTC = Resting_Folder & "\CTC\SBTC\Charge_Slip\" & Format(Now, "YYYY") & "\"
        End If
    End If
    
    
    
    
    
    
    BlockCount = 0
    TotalData = 0
    
    
    
    
    
    
    
    'For Output File
    Close #1, #2, #3
    Open App.Path & "\" & FolderName & "\Block" & ChkType2 & ".txt" For Output As #1
    'End For Output File
    
    
    
    
    
    
    Set DBFConnector = CreateObject("ADODB.Connection")
    DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
    DBFConnector.CursorLocation = adUseClient

    Set DBFConnector_Ref = CreateObject("ADODB.Connection")
    DBFConnector_Ref.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Ref_Location & "\;Extended properties=dBase III"
    DBFConnector_Ref.CursorLocation = adUseClient
    
    Set DBFConnector1 = CreateObject("ADODB.Connection")
    DBFConnector1.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\" & FolderName & "\;Extended properties=dBase III"
    DBFConnector1.CursorLocation = adUseClient
    
    
    
    
    
    
    'For MySQL
    Dim Conn_SQL As ADODB.Connection

    Set Conn_SQL = New ADODB.Connection
    Conn_SQL.ConnectionString = "uid=cpc;pwd=CorpCaptive;server=" & Target_ip & ";driver={MySQL ODBC 5.1 Driver};database=captive_database;dsn=;"
    Conn_SQL.Open
    'End For MySQL

    
    
    
    
    If DeleteDBF_Value = "TRUE" Then
        Result = DeleteDBF("Packing", FolderName)
    End If
    


    
    
    Set dbfRecordset = CreateObject("ADODB.Recordset")
    SQL = "SELECT BRSTN, AccountNo, OrderQty, Name1, Name2, Address1, Address2, Address3, Address4, Address5, Address6, Batch, BStock, StartSN, PcsPerBook , PKey FROM SBTC WHERE ChkType = '" & ChkType & "' AND FormType = '" & FormType & "' ORDER BY BRSTN, AccountNo, Name1"
    dbfRecordset.Open SQL, DBFConnector, 1, 1
    
    
    
    DataNumber = 0
    
    LoopCount = 0
    Do Until LoopCount = dbfRecordset.RecordCount
        
        
        If LoopCount = 0 Then
            Close #2, #3
            Open App.Path & "\" & FolderName & "\" & FileName & ".txt" For Output As #2
            Open App.Path & "\" & FolderName & "\" & FileName & "." & Format(Now, "YY") & "P" For Output As #3
            
            
            
            'Copy Printer File MDB
            If (ChkType = "F" And FormType = "25") Or (ChkType = "F" And FormType = "26") Or (ChkType = "GC" And FormType = "20") Or (ChkType = "MC" And FormType = "20") Or (ChkType = "MC_1" And FormType = "00") Or (ChkType = "CUSTOM" And FormType = "00") Or ChkType = "CS" Then
                FileCopy App.Path & "\DataSource.mdb", App.Path & "\" & FolderName & "\" & FileName & ".mdb"
                
                
                Dim Conn1 As ADODB.Connection
                Dim Rs1 As ADODB.Recordset
                
                Set Conn1 = New ADODB.Connection
                
                With Conn1
                .CursorLocation = adUseClient
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source=" & App.Path & "\" & FolderName & "\" & FileName & ".mdb; Jet OLEDB:Database Password=CorpCaptive;"
                .Open
                End With
                
            End If
            'End Copy Printer File MDB
        End If
        
        
        
        BRSTN = dbfRecordset.Fields(0)
        AccountNo = dbfRecordset.Fields(1)
        OrderQty = dbfRecordset.Fields(2)
        
        If Len(dbfRecordset.Fields(3)) >= 1 Then
            Name1 = dbfRecordset.Fields(3)
        Else
            Name1 = ""
        End If
        
        If Len(dbfRecordset.Fields(4)) >= 1 Then
            Name2 = dbfRecordset.Fields(4)
        Else
            Name2 = ""
        End If
        
        If Len(dbfRecordset.Fields(5)) >= 1 Then
            Address1 = dbfRecordset.Fields(5)
        Else
            Address1 = ""
        End If

        If Len(dbfRecordset.Fields(6)) >= 1 Then
            Address2 = dbfRecordset.Fields(6)
        Else
            Address2 = ""
        End If
        
        If Len(dbfRecordset.Fields(7)) >= 1 Then
            Address3 = dbfRecordset.Fields(7)
        Else
            Address3 = ""
        End If
        
        If Len(dbfRecordset.Fields(8)) >= 1 Then
            Address4 = dbfRecordset.Fields(8)
        Else
            Address4 = ""
        End If
        
        If Len(dbfRecordset.Fields(9)) >= 1 Then
            Address5 = dbfRecordset.Fields(9)
        Else
            Address5 = ""
        End If
        
        If Len(dbfRecordset.Fields(10)) >= 1 Then
            Address6 = dbfRecordset.Fields(10)
        Else
            Address6 = ""
        End If
        
        Batch = UCase(dbfRecordset.Fields(11))
        Batch_Orig = UCase(dbfRecordset.Fields(11))
        
        Basestock = dbfRecordset.Fields(12)
        
        If Len(dbfRecordset.Fields(13)) >= 1 Then
            StartingSerial = dbfRecordset.Fields(13)
        Else
            StartingSerial = ""
        End If
        
        If ChkType = "CUSTOM" And FormType = "00" Then
            PcsPerBook = dbfRecordset.Fields(14)
        End If
        
        Pkey = dbfRecordset.Fields(15)
        
        
        
        
        
        'Rename Batch
        If Mid(Batch, 1, 10) = "13DIGIT_CP" Or Mid(Batch, 1, 10) = "13DIGIT_CS" Or Mid(Batch, 1, 10) = "13DIGIT_NB" Then Batch = Mid(Batch_Orig, 1, 2) & Mid(Batch_Orig, 9, Len(Batch_Orig))
        If Mid(Batch, 4, 10) = "13DIGIT_CP" Then Batch = Mid(Batch_Orig, 4, 2) & Mid(Batch_Orig, 12, Len(Batch_Orig))
        If FolderName = "MC" Then
            Batch = "MC" & Mid(Batch_Orig, 11, 4) & Mid(FinalBatch, 7, 8)
        End If
        
        If FolderName = "GiftCheck" Then
            Batch = "GC" & Mid(Batch_Orig, 11, 4) & Mid(FinalBatch, 7, 8)
        End If
        
        
        Set dbfRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO Batch (Batch) VALUES ('" & Batch & "')"
        dbfRecordset1.Open SQL, DBFConnector, 1, 1
        'End Rename Batch
        
        
        
        
        
        
        
        
        'For Total
        If Val(TotalData) = 0 Then
            
            If Weekday(DeliveryDate, vbSunday) = 1 Then DayOfWeekResult = "SUN"
            If Weekday(DeliveryDate, vbSunday) = 2 Then DayOfWeekResult = "MON"
            If Weekday(DeliveryDate, vbSunday) = 3 Then DayOfWeekResult = "TUE"
            If Weekday(DeliveryDate, vbSunday) = 4 Then DayOfWeekResult = "WED"
            If Weekday(DeliveryDate, vbSunday) = 5 Then DayOfWeekResult = "THU"
            If Weekday(DeliveryDate, vbSunday) = 6 Then DayOfWeekResult = "FRI"
            If Weekday(DeliveryDate, vbSunday) = 7 Then DayOfWeekResult = "SAT"
                    
            Temp = Batch
            Do Until Len(Temp) >= 45
                Temp = Temp & " "
            Loop
            
            Summary_DoBlock = "    " & Temp & "DLVR: " & Format(DeliveryDate, "MM-DD") & "(" & DayOfWeekResult & ")" & vbNewLine & vbNewLine
            
            
            Set dbfRecordset1 = CreateObject("ADODB.Recordset")
            SQL = "SELECT SUM(OrderQty) FROM SBTC WHERE ChkType = '" & ChkType_1 & "' AND FormType = '" & FormType_1 & "'"
            dbfRecordset1.Open SQL, DBFConnector, 1, 1
            
            Sum_DoBlock = "    " & ChkType_31 & " = " & dbfRecordset1.Fields(0) & "                 " & FileName & ".txt"
            
            If ChkType_31 <> ChkType_32 Then
                Set dbfRecordset1 = CreateObject("ADODB.Recordset")
                SQL = "SELECT SUM(OrderQty) FROM SBTC WHERE ChkType = '" & ChkType_2 & "' AND FormType = '" & FormType_2 & "'"
                dbfRecordset1.Open SQL, DBFConnector, 1, 1
                
                Sum_DoBlock = Sum_DoBlock & vbNewLine & "    " & ChkType_32 & " = " & dbfRecordset1.Fields(0)
            End If
            
            Summary_DoBlock = Summary_DoBlock & vbNewLine & Sum_DoBlock
            
            Summary_DoBlock = Summary_DoBlock & vbNewLine & vbNewLine & "    Prepared By : " & ProcessBy & vbNewLine & "    Updated By  : " & ProcessBy & vbNewLine & "    Time Start  : " & TimeStart & vbNewLine & "    Time Fnished:                                              RECHECKED BY:  " & vbNewLine & "    File Rcvd   :"
        End If
        'End For Total
        
        
        
        
        
        
        
        
        
        
        If FolderName = "MC" Or FolderName = "GiftCheck" Then Name1 = ""
        
        
        
        
        
        
        'For Ref.dbf Starting SN
        If ChkType = "CS" Then BranchName = "CHARGE SLIP"
        
        If ChkType <> "CS" Then
            If (ChkType = "MC_1" And FormType = "00") Or (ChkType = "CUSTOM" And FormType = "00") Then
                Set dbfRecordset1 = CreateObject("ADODB.Recordset")
                SQL = "INSERT INTO Master ([Date], Batch, BRSTN, AccountNo, StartSN, EndSN, OrderQty, Address1, Address2, Address3, Address4, Address5, Address6, Name1, Name2) VALUES ('" _
                    & Now & "','" & Batch & "','" & BRSTN & "','" & AccountNo & "','" & StartingSerial & "','" & Val(StartingSerial) + (Val(OrderQty) * Val(PcsPerBook)) - 1 & "','" & OrderQty & "','" & Replace(Address1, "'", "''") & "','" & Replace(Address2, "'", "''") & "','" & Replace(Address3, "'", "''") & "','" & Replace(Address4, "'", "''") & "','" & Replace(Address5, "'", "''") & "','" & Replace(Address6, "'", "''") & "','" & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "')"
                dbfRecordset1.Open SQL, DBFConnector1, 1, 1
                
                BranchName = Address1
            Else
                Set dbfRecordset1 = CreateObject("ADODB.Recordset")
                SQL = "SELECT LastNo,C_Before, Branch_Tex FROM REF WHERE RTNO = '" & BRSTN & "' AND ChkType = '" & RefChkType & "'"
                dbfRecordset1.Open SQL, DBFConnector_Ref, 1, 1
                
                
                EndingSerial = dbfRecordset1.Fields(0)
                C_Before = dbfRecordset1.Fields(1)
                StartingSerial = Val(dbfRecordset1.Fields(0)) + 1
                NewLastNo = (Val(PcsPerBook) * Val(OrderQty)) + EndingSerial
                BranchName = dbfRecordset1.Fields(2)
                
                Set dbfRecordset1 = CreateObject("ADODB.Recordset")
                SQL = "UPDATE REF SET [Date] = '" & Format(Now, "MM/DD/YYYY") & "', LastNo = '" & NewLastNo & "', P_Before = '" & C_Before & "', C_Before = '" & EndingSerial & "' WHERE RTNO = '" & BRSTN & "' AND ChkType = '" & RefChkType & "'"
                dbfRecordset1.Open SQL, DBFConnector_Ref, 1, 1
            End If
        End If
        'End For Ref.dbf Starting SN
            
            
            
        
        
        
        
        
        
        'Update SBTC StartSN1
        Set dbfRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "UPDATE SBTC SET StartSN1 = " & StartingSerial & ",PcsPerBook = '" & PcsPerBook & "' WHERE PKey = " & Pkey
        dbfRecordset1.Open SQL, DBFConnector, 1, 1
        'End Update SBTC StartSN1
            
            
            
            

        Do Until Val(OrderQty) = 0
            
            'For Do-Block
            If TotalData Mod 32 = 0 Then
                
                
                If TotalData <> 0 Then
                    If Val(TotalData) = 32 Then
                        Print #1, ""
                        Print #1, Summary_DoBlock
                    End If
                    
                    
                    Print #1, ""
                End If
                
                Print #1, ""
                Print #1, "        Page No. " & (TotalData / 32) + 1
                Print #1, "        " & Format(Now, "Mmm. DD, YYYY")
                Print #1, "                   SBTC - SUMMARY OF BLOCK - " & Description
                If ChkType = "AA" Or ChkType = "BB" Then Print #1, "                                Pre-Encoded"
                Print #1, "                Basestock = " & Basestock & " --> " & Basetock_Verify
                Print #1, ""
                
                
                
                'For Heading Basestock
                If (ChkType = "F" And FormType = "25") Then
                    Print #1, "    Starting 04072016, please use program 'SBTC Personal NCDSS (25 Pcs per Book)'"
                    Print #1, "and use the Basestock of 'SBTC Regular PA' if Basestock IS NOT AVAILABLE for Check One PA"
                    Print #1, ""
                End If
                
                If (ChkType = "F" And FormType = "26") Then
                    Print #1, "    Starting 04072016, please use program 'SBTC Commercial NCDSS (50 Pcs per Book)'"
                    Print #1, "and use the Basestock of 'SBTC Regular CA' if Basestock IS NOT AVAILABLE for Check One CA"
                    Print #1, ""
                End If
                
                If (ChkType = "E" And FormType = "23") Then
                    Print #1, "    Starting 04072016, please use program 'SBTC Personal NCDSS'"
                    Print #1, "and use the Basestock of 'SBTC Regular PA' if Basestock IS NOT AVAILABLE for Check Power PA"
                    Print #1, ""
                End If
                
                If (ChkType = "E" And FormType = "22") Then
                    Print #1, "    Starting 04072016, please use program 'SBTC Commercial NCDSS'"
                    Print #1, "and use the Basestock of 'SBTC Regular CA' if Basestock IS NOT AVAILABLE for Check Power CA"
                    Print #1, ""
                End If
                
                If (ChkType = "MC_1" And FormType = "00") Or (ChkType = "CUSTOM" And FormType = "00") Then
                    Print #1, "                  A L L  M A N U A L  E N C O D E D ! ! !"
                    Print #1, ""
                End If
                'End For Heading Basestock
                
                
                
                
                'For Heading Carbon
                If (ChkType = "F" And FormType = "25") Or (ChkType = "F" And FormType = "26") Or (ChkType = "GC" And FormType = "20") Or (ChkType = "MC" And FormType = "20") Then
                    Print #1, "            *** With Duplicate Copy ---> " & FileName & ".mdb" & " ***"
                    Print #1, ""
                End If
                
                If (ChkType = "MC_1" And FormType = "00") Then
                    Print #1, "            *** With Triplicate Copy ---> " & FileName & ".mdb" & " ***"
                    Print #1, ""
                End If
                'End For Heading Carbon
                
                
                Print #1, "    Starting Batch 02042016, New MICR Alignment of NCDSS is 15-54 ! ! !"
                Print #1, ""
                Print #1, "        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO."
                Print #1, ""
            End If
            
            
            If TotalData Mod 4 = 0 Then
                BlockCount = Val(BlockCount) + 1
                
                Print #1, ""
                Print #1, "       ** BLOCK " & Val(BlockCount)
            End If
            
            Do Until Len(BlockCount) >= 13
                BlockCount = " " & BlockCount
            Loop
            
            StartingSerial = Format(Val(StartingSerial), FormatSerial)
            Do Until Len(StartingSerial) >= 11
                StartingSerial = StartingSerial & " "
            Loop
            
            Print #1, BlockCount & " " & BRSTN & "   " & AccountNo & "    " & StartingSerial & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) & " " & Status
            'End For Do-Block
            
            
            
            
            
            
            
            
            'For CTC Printer File
            If TotalData Mod 4 = 0 Then
                If TotalData <> 0 Then Print #3, "\"
                Print #3, "  " & Format(Val(TotalData / 4), "000000") & "3"
                Print #3, "3"
            Else
                Print #3, "3"
                Print #3, "3"
            End If
            
            
            Print #3, BRSTN
            Print #3, BRSTN
            Print #3, AccountNo
            Print #3, AccountNo
            Print #3, Format(Val(StartingSerial) + Val(PcsPerBook), FormatSerial)
            Print #3, Format(Val(StartingSerial) + Val(PcsPerBook), FormatSerial)
            Print #3, "A"
            Print #3, "A"
            Print #3, MICRLine & Mid(BRSTN, 1, 5) & "D" & Mid(BRSTN, 6, 4) & "T" & AccountNo & "O"
            Print #3, MICRLine & Mid(BRSTN, 1, 5) & "D" & Mid(BRSTN, 6, 4) & "T" & AccountNo & "O"
            Print #3, Mid(BRSTN, 1, 5)
            Print #3, Mid(BRSTN, 1, 5)
            Print #3, " " & Mid(BRSTN, 6, 4)
            Print #3, " " & Mid(BRSTN, 6, 4)
            Print #3, Format(Val(AccountNo), "000-000000-000")
            Print #3, Format(Val(AccountNo), "000-000000-000")
            Print #3, Name1
            Print #3, Name1
            Print #3, "SN"
            Print #3, "SN"
            Print #3, ""
            Print #3, ""
            Print #3, Name2
            Print #3, Name2
            Print #3, "C"
            Print #3, "C"
            Print #3, "XXXX"
            Print #3, "XXXX"
            Print #3, Name3
            Print #3, Name3
            
            If ChkType = "GC" And FormType = "20" Then
                Print #3, Trim(Replace(Address1, "BRANCH", ""))
                Print #3, Trim(Replace(Address1, "BRANCH", ""))
            Else
                Print #3, Address1
                Print #3, Address1
            End If
            
            Print #3, Address2
            Print #3, Address2
            Print #3, Address3
            Print #3, Address3
            Print #3, Address4
            Print #3, Address4
            Print #3, Address5
            Print #3, Address5
            Print #3, Address6
            Print #3, Address6
            Print #3, "SECURITY BANK"
            Print #3, "SECURITY BANK"
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, ""
            Print #3, Format(Val(StartingSerial), FormatSerial)
            Print #3, Format(Val(StartingSerial), FormatSerial)
            Print #3, Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial)
            Print #3, Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial)
            'End For CTC Printer File
            
            
            
            
            'For Printer File
            If TotalData Mod 4 = 0 Then
                If TotalData = 0 Then
                    Print #2, "3" '1
                Else
                    Print #2, "3" '1
                End If
            Else
                Print #2, "3" '1
            End If
            
            Print #2, BRSTN '2
            Print #2, AccountNo '3
            Print #2, Format(Val(StartingSerial) + Val(PcsPerBook), FormatSerial) '4
            Print #2, "A" '5
            Print #2, MICRLine & Mid(BRSTN, 1, 5) & "D" & Mid(BRSTN, 6, 4) & "T" & AccountNo & "O" '6
            
            If ChkType = "MC_1" And FormType = "00" Then
                Print #2, Format(Val(BRSTN), "00000-000-0") '7
                Print #2, "" '8"
            Else
                Print #2, Mid(BRSTN, 1, 5) '7
                Print #2, " " & Mid(BRSTN, 6, 4) '8
            End If
            
            Print #2, Format(Val(AccountNo), "000-000000-000") '9
            Print #2, Name1 '10
            Print #2, "SN" '11
            Print #2, "" '12
            Print #2, Name2 '13
            Print #2, "C" '14
            Print #2, "XXXX" '15
            Print #2, Name3 '16
            
            
            
            If ChkType = "GC" And FormType = "20" Then
                Print #2, Trim(Replace(Address1, "BRANCH", "")) '17
                
                Temp_Address1 = Trim(Replace(Address1, "BRANCH", ""))
            Else
                Print #2, Address1 '17
                
                Temp_Address1 = Address1
            End If
            
            
            
            Print #2, Address2 '18
            Print #2, Address3 '19
            Print #2, Address4 '20
            Print #2, Address5 '21
            Print #2, Address6 '22
            Print #2, "SECURITY BANK" '23
            Print #2, "" '24
            Print #2, "" '25
            Print #2, "" '26
            Print #2, "" '27
            Print #2, "" '28
            Print #2, "" '29
            Print #2, "" '30
            Print #2, Format(Val(StartingSerial), FormatSerial) '31
            Print #2, Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) '32
            'End For Printer File
            
            
            
            
            
            
            
            'Save to Master_Database
            If CodesOnly = True Then DBase = "Master_Database_SBTC_Temp"
            If CodesOnly = False Then DBase = "Master_Database_SBTC"
            
            
            SQL = "INSERT INTO " & DBase & " (Date , Time , DeliveryDate , ChkType , ChequeName , BRSTN , AccountNo , Name1 , Name2 , StartingSerial , EndingSerial , Batch , Address1 , Address2 , Address3 , Address4 , Address5 , Address6) VALUES ('" _
                & DateToday_Final & "','" & TimeToday_Final & "','" & Format(DeliveryDate, "YYYY-MM-DD") & "','" & ChkType & "','" & Replace(getChequeName(ChkType, FormType), "'", "''") & "','" & BRSTN & "','" & AccountNo & "','" & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','" & Format(Val(StartingSerial), FormatSerial) & "','" & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) & "','" & Batch & "','" & Replace(Address1, "'", "''") & "','" & Replace(Address2, "'", "''") & "','" & Replace(Address3, "'", "''") & "','" & Replace(Address4, "'", "''") & "','" & Replace(Address5, "'", "''") & "','" & Replace(Address6, "'", "''") & "')"
            Set Rs = New ADODB.Recordset
            Rs.CursorLocation = adUseClient
            Rs.Open SQL, Conn_SQL, adOpenKeyset, adLockOptimistic
            'End Save to Master_Database
            
            
            
            'For MDB File
            If (ChkType = "F" And FormType = "25") Or (ChkType = "F" And FormType = "26") Or (ChkType = "GC" And FormType = "20") Or (ChkType = "MC" And FormType = "20") Or (ChkType = "MC_1" And FormType = "00") Or (ChkType = "CUSTOM" And FormType = "00") Or ChkType = "CS" Then
                TempStartingSerial = Val(StartingSerial)
                
                LoopCount1 = 0
                Do Until LoopCount1 = Val(PcsPerBook)
                    Set Rs = New ADODB.Recordset
                    SQL = "INSERT INTO InputFile_Temp (BRSTN, AccountNumber, RT1to5, RT6to9,AccountNumberWithHyphen,Serial,Name1,Name2,Name3,Address1,Address2,Address3,Address4,Address5,Address6,BankName, StartingSerial, EndingSerial,PcsPerBook,DataNumber) VALUES ('" _
                        & BRSTN & "','" & AccountNo & "','" & Mid(BRSTN, 1, 5) & "','" & Mid(BRSTN, 6, 4) & "','" & Format(Val(AccountNo), "000-000000-000") & "','" & Format(Val(TempStartingSerial), FormatSerial) & "','" & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','" & Replace(Name3, "'", "''") & "','" & Replace(Temp_Address1, "'", "''") & "','" & Replace(Address2, "'", "''") & "','" & Replace(Address3, "'", "''") & "','" & Replace(Address4, "'", "''") & "','" & Replace(Address5, "'", "''") & "','" & Replace(Address6, "'", "''") & "','SECURITY BANK','" & Format(Val(StartingSerial), FormatSerial) & "','" & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) & "','" & PcsPerBook & "','" & DataNumber + 1 & "')"
                    Rs.Open SQL, Conn1, adOpenStatic
                    
                    TempStartingSerial = TempStartingSerial + 1
                    DataNumber = DataNumber + 1
                    LoopCount1 = LoopCount1 + 1
                Loop
            End If
            'End For MDB File
            
            
            
            
            'For Packing
            Set dbfRecordset2 = CreateObject("ADODB.Recordset")
            SQL = "INSERT INTO Packing (BatchNo, RT_NO, Branch, Acct_No, Acct_No_P, Acct_Name1, Acct_Name2, No_Bks, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, Block, ChkType) VALUES ('" _
                & Batch & "','" & BRSTN & "','" & Replace(BranchName, "'", "''") & "','" & AccountNo & "','" & Format(Val(AccountNo), "000000-00000-0") & "','" & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','" & "1" & "','" & Format(Val(StartingSerial), FormatSerial) & "','" & Format(Val(StartingSerial), FormatSerial) & "','" & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) & "','" & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) & "','" & Val(BlockCount) & "','" & RefChkType & "')"
            dbfRecordset2.Open SQL, DBFConnector1, 1, 1
            'End for Packing
            
            
            TotalData = Val(TotalData) + 1
            StartingSerial = Val(StartingSerial) + Val(PcsPerBook)
            OrderQty = OrderQty - 1
        Loop
        
        
        dbfRecordset.MoveNext
        LoopCount = LoopCount + 1
    Loop
    
    If LoopCount >= 1 Then
        Print #2, "\"
        Print #3, "\"
    End If
    
    If Val(TotalData) <= 32 Then
        Print #1, ""
        Print #1, Summary_DoBlock
    End If
    Close #1, #2, #3
    
    
    
    
    
    
    Result = PackingList(ChkType, FolderName, FormType, RefChkType)
    
    
    
    
    
    
    
    
    ProcessMe = TotalData
    
    
    
    
    
    
    If TotalData >= 1 Then
        
        
        If Dir("C:\Windows\Temp\" & DateTimeToday, vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday
        
        
        If UCase(FolderName) = "REGULAR\PREENCODED" Then
            If Dir("C:\Windows\Temp\" & DateTimeToday & "\Regular\", vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\Regular\"
        End If
        
        
        If ChkType = "MC_1" And FormType = "00" Then
            If Dir("C:\Windows\Temp\" & DateTimeToday & "\MC", vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\MC\"
        End If
        
        If Dir("C:\Windows\Temp\" & DateTimeToday & "\" & FolderName, vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\" & FolderName & "\"
        
        
        
        
        FileCopy App.Path & "\" & FolderName & "\Block" & ChkType2 & ".txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderName & "\Block" & ChkType2 & ".txt"
        FileCopy App.Path & "\" & FolderName & "\" & FileName & ".txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderName & "\" & FileName & ".txt"
        FileCopy App.Path & "\" & FolderName & "\" & FileName & "." & Format(Now, "YY") & "P", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderName & "\" & FileName & "." & Format(Now, "YY") & "P"
        FileCopy App.Path & "\" & FolderName & "\Packing" & RefChkType & ".txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderName & "\Packing" & RefChkType & ".txt"
        
        If FolderName = "Regular\PreEncoded" Then
            If Dir("C:\Windows\Temp\" & DateTimeToday & "\" & FinalBatch & "\Regular\", vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\" & FinalBatch & "\Regular"
        End If
        
        If ChkType = "MC_1" And FormType = "00" Then
            If Dir("C:\Windows\Temp\" & DateTimeToday & "\" & FinalBatch & "\MC", vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\" & FinalBatch & "\MC"
        End If
        
        If Dir("C:\Windows\Temp\" & DateTimeToday & "\" & FinalBatch & "\" & FolderName, vbDirectory) = "" Then MkDir "C:\Windows\Temp\" & DateTimeToday & "\" & FinalBatch & "\" & FolderName
        
        
        FileCopy App.Path & "\" & FolderName & "\Packing" & RefChkType & ".txt", "C:\Windows\Temp\" & DateTimeToday & "\" & FinalBatch & "\" & FolderName & "\Packing" & RefChkType & ".txt"
        
        
        
        
        If (ChkType = "F" And FormType = "25") Or (ChkType = "F" And FormType = "26") Or (ChkType = "GC" And FormType = "20") Or (ChkType = "MC" And FormType = "20") Or (ChkType = "MC_1" And FormType = "00") Or (ChkType = "CUSTOM" And FormType = "00") Or ChkType = "CS" Then
            'Make it Padded
            Do Until DataNumber Mod (PcsPerBook * 4) = 0
                Set Rs = New ADODB.Recordset
                SQL = "INSERT INTO InputFile_Temp (DataNumber) VALUES ('" & DataNumber + 1 & "')"
                Rs.Open SQL, Conn1, adOpenStatic
                
                DataNumber = DataNumber + 1
            Loop
                   
                   
            LineNumber1 = PcsPerBook * 0
            LineNumber2 = PcsPerBook * 1
            LineNumber3 = PcsPerBook * 2
            LineNumber4 = PcsPerBook * 3
            
RepeatMe:
            
            LoopCount = 0
            Do Until LoopCount = Val(PcsPerBook)
                'Line Number 1
                LineNumber1 = LineNumber1 + 1

                Set Rs = New ADODB.Recordset
                SQL = "INSERT INTO InputFile SELECT * FROM InputFile_Temp WHERE DataNumber = '" & LineNumber1 & "'"
                Rs.Open SQL, Conn1, adOpenStatic
                'End Line Number 1
                
                
                'Line Number 2
                LineNumber2 = LineNumber2 + 1

                Set Rs = New ADODB.Recordset
                SQL = "INSERT INTO InputFile SELECT * FROM InputFile_Temp WHERE DataNumber = '" & LineNumber2 & "'"
                Rs.Open SQL, Conn1, adOpenStatic
                'End Line Number 2
                
                
                'Line Number 3
                LineNumber3 = LineNumber3 + 1

                Set Rs = New ADODB.Recordset
                SQL = "INSERT INTO InputFile SELECT * FROM InputFile_Temp WHERE DataNumber = '" & LineNumber3 & "'"
                Rs.Open SQL, Conn1, adOpenStatic
                'End Line Number 3
                
                
                'Line Number 4
                LineNumber4 = LineNumber4 + 1

                Set Rs = New ADODB.Recordset
                SQL = "INSERT INTO InputFile SELECT * FROM InputFile_Temp WHERE DataNumber = '" & LineNumber4 & "'"
                Rs.Open SQL, Conn1, adOpenStatic
                'End Line Number 4
                
                
                LoopCount = LoopCount + 1
            Loop
            
            If LineNumber4 <> DataNumber Then
                LineNumber1 = LineNumber1 + (PcsPerBook * 3)
                LineNumber2 = LineNumber2 + (PcsPerBook * 3)
                LineNumber3 = LineNumber3 + (PcsPerBook * 3)
                LineNumber4 = LineNumber4 + (PcsPerBook * 3)
                
                GoTo RepeatMe
                Exit Function
            End If
            'End Make it Padded
            
            Conn1.Close
            
            
            FileCopy App.Path & "\" & FolderName & "\" & FileName & ".mdb", "C:\Windows\Temp\" & DateTimeToday & "\" & FolderName & "\" & FileName & ".mdb"
            
            
            'Copy to Drive R
            CreateDirectory (Temp_DriveR)
            
            FileCopy App.Path & "\" & FolderName & "\" & FileName & ".mdb", Temp_DriveR & "\" & FileName & ".mdb"
            'End Copy to Drive R
            
            'Copy to Drive R
            CreateDirectory (Temp_CTC)
            
            FileCopy App.Path & "\" & FolderName & "\" & FileName & ".mdb", Temp_CTC & "\" & FileName & ".mdb"
            'End Copy to Drive R
        End If
        
        
        
        'Copy to Drive R
        CreateDirectory (Temp_DriveR)
        
        FileCopy App.Path & "\" & FolderName & "\" & FileName & ".txt", Temp_DriveR & "\" & FileName & ".txt"
        'Copy to Drive R
        
        
        
        
        'Copy to CTC
        CreateDirectory (Temp_CTC)
        
        FileCopy App.Path & "\" & FolderName & "\" & FileName & ".txt", Temp_CTC & "\" & FileName & ".txt"
        'End Copy to CTC
        
        
        
        
    End If
End Function



Sub CreateDirectory(Directory_Location)
Dim FSO As New FileSystemObject

Temp_Count = 0

LoopCount = 2
Do Until LoopCount = Len(Directory_Location)


    If Mid(Directory_Location, LoopCount + 1, 1) = "\" Then
        CurrentDirectory = Mid(Directory_Location, 1, LoopCount)

        'Count the \
        LoopCount1 = 0
        Do Until LoopCount1 = Len(CurrentDirectory)
            If Mid(CurrentDirectory, LoopCount1 + 1, 1) = "\" Then
                Temp_Count = Temp_Count + 1
            End If

            LoopCount1 = LoopCount1 + 1
        Loop
        'End Count the \

        If Temp_Count >= 4 Then
            If FSO.FolderExists(CurrentDirectory) = False Then FSO.CreateFolder (CurrentDirectory)
        End If

    End If

    LoopCount = LoopCount + 1
Loop
End Sub






Function PackingList(ChkType, FolderName, FormType, ChkTypeRef)
PageNo = 0

Close #1
Open App.Path & "\" & FolderName & "\Packing" & ChkTypeRef & ".txt" For Output As #1


Print_Front_Cover = False


RepeatMe:



Dim DBFConnector, dbfRecordset, dbfRecordset1, LoopCount, LoopCount1, SQL


Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\" & FolderName & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient


Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT RT_NO, Branch, Sum(No_Bks), BatchNo FROM Packing WHERE ChkType = '" & ChkTypeRef & "' GROUP BY RT_NO, Branch, BatchNo ORDER BY BatchNo, RT_NO"
dbfRecordset.Open SQL, DBFConnector, 1, 1
    


LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    BRSTN = dbfRecordset.Fields(0)
    BranchName = dbfRecordset.Fields(1)
    OrderQty = dbfRecordset.Fields(2)
    Batch = dbfRecordset.Fields(3)
    
    
    
    'For Heading
    PageNo = Val(PageNo) + 1
    
    If LoopCount <> 0 Then Print #1, ""
    Print #1, ""
    Print #1, "  Page No. " & Val(PageNo)
    Print #1, "  " & Format(Now, "Mmm. DD, YYYY")
    Print #1, "                                CAPTIVE PRINTING CORPORATION"
    
    If ChkType = "A" And FormType = "05" Then Print #1, "                               SBTC - Personal Checks Summary"
    If ChkType = "B" And FormType = "16" Then Print #1, "                               SBTC - Commercial Checks Summary"
    
    If ChkType = "AA" And FormType = "05" Then Print #1, "                               SBTC - Personal PreEncoded Checks Summary"
    If ChkType = "BB" And FormType = "16" Then Print #1, "                               SBTC - Commercial PreEncoded Checks Summary"
    
    If ChkType = "MC" And FormType = "20" Then Print #1, "                               SBTC - Manager's Checks Summary"
    
    If ChkType = "F" And FormType = "25" Then Print #1, "                               SBTC - Personal CheckOne Summary"
    If ChkType = "F" And FormType = "26" Then Print #1, "                               SBTC - Commercial CheckOne Summary"
    
    If ChkType = "E" And FormType = "23" Then Print #1, "                               SBTC - Personal CheckPower Summary"
    If ChkType = "E" And FormType = "22" Then Print #1, "                               SBTC - Commercial CheckPower Summary"
    
    If ChkType = "GC" And FormType = "20" Then Print #1, "                               SBTC - Gift Checks Summary"
    If ChkType = "MC_1" And FormType = "00" Then Print #1, "                               SBTC - Manager's Checks Continues Summary"
    If ChkType = "CUSTOM" And FormType = "00" Then Print #1, "                               SBTC - Customized Checks Summary"
    
    If ChkType = "CS" Then Print #1, "                               SBTC - Charge Slip Checks Summary"
    
    
    
    'For WithOut Name
    If Print_Front_Cover = True Then
        Print #1, "                                 ( F R O N T  C O V E R )"
    End If
    'End For WithOut Name
    
    Print #1, ""
    Print #1, "  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #"
    Print #1, ""
    Print #1, ""
    Print #1, " ** ORDERS OF BRSTN " & BRSTN & " " & BranchName
    Print #1, ""
    Print #1, " * Batch #: " & Batch
    'End For Heading
    
    
    
    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT Acct_No_P, Acct_Name1, Acct_Name2, Ck_No_B, CK_NO_E FROM Packing WHERE CHkTYpe = '" & ChkTypeRef & "' AND RT_NO = '" & BRSTN & "' AND Branch = '" & Replace(BranchName, "'", "''") & "' AND BatchNo = '" & Batch & "' ORDER BY Acct_No_P, CK_NO_B"
    dbfRecordset1.Open SQL, DBFConnector, 1, 1
    
    
    
    LoopCount1 = 0
    Do Until LoopCount1 = dbfRecordset1.RecordCount
        
        AccountNo = dbfRecordset1.Fields(0)
        
        If Len(dbfRecordset1.Fields(1)) >= 1 Then
            Name1 = dbfRecordset1.Fields(1)
        Else
            Name1 = ""
        End If
        
        If Len(dbfRecordset1.Fields(2)) >= 1 Then
            Name2 = dbfRecordset1.Fields(2)
        Else
            Name2 = ""
        End If
    
        'For WithOut Name
        If Print_Front_Cover = True Then
            Name1 = ""
            Name2 = ""
        End If
        'End For WithOut Name
    
        StartingSerial = dbfRecordset1.Fields(3)
        EndingSerial = dbfRecordset1.Fields(4)
        
        Do Until Len(Name1) >= 35
            Name1 = Name1 & " "
        Loop
        
        Do Until Len(StartingSerial) >= 11
            StartingSerial = StartingSerial & " "
        Loop
        
        Print #1, "  " & AccountNo & "  " & Name1 & "1 " & Replace(Replace(ChkType, "MC_1", "B"), "CUSTOM", "B") & "  " & StartingSerial & EndingSerial
        If Name2 <> "" Then Print #1, "                  " & Name2
        
        dbfRecordset1.MoveNext
        LoopCount1 = LoopCount1 + 1
    Loop
    
    Print #1, ""
    Print #1, ""
    Print #1, " * * * Sub Total * * *                              " & OrderQty
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop




If Print_Front_Cover = False Then
    Print #1, ""
    
    Print_Front_Cover = True
    
    GoTo RepeatMe
    Exit Function
End If
Close #1

End Function
  
  
  
  
  
  
Function ProcessAll(Batch, ProcessBy, CheckedBy, DeliveryDate)
Dim FSO As New FileSystemObject

Set DBFConnector = CreateObject("ADODB.Connection")
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient










Reg_PA = ProcessMe("A", "05", "Regular", "TRUE", Batch, DeliveryDate)
Reg_CA = ProcessMe("B", "16", "Regular", "", Batch, DeliveryDate)

PreEncoded_PA = ProcessMe("AA", "05", "Regular\PreEncoded", "TRUE", Batch, DeliveryDate)
PreEncoded_CA = ProcessMe("BB", "16", "Regular\PreEncoded", "", Batch, DeliveryDate)

MC = ProcessMe("MC", "20", "MC", "TRUE", Batch, DeliveryDate)
    
CheckOne_PA = ProcessMe("F", "25", "CheckOne", "TRUE", Batch, DeliveryDate)
CheckOne_CA = ProcessMe("F", "26", "CheckOne", "", Batch, DeliveryDate)

CheckPower_PA = ProcessMe("E", "23", "CheckPower", "TRUE", Batch, DeliveryDate)
CheckPower_CA = ProcessMe("E", "22", "CheckPower", "", Batch, DeliveryDate)

GiftCheck = ProcessMe("GC", "20", "GiftCheck", "TRUE", Batch, DeliveryDate)

MC_Continues = ProcessMe("MC_1", "00", "MC\Continues", "TRUE", Batch, DeliveryDate)

Customized = ProcessMe("CUSTOM", "00", "Customized", "TRUE", Batch, DeliveryDate)

Charge_Slip = ProcessMe("CS", "00", "Charge_Slip", "TRUE", Batch, DeliveryDate)






    
    
'For Zip
If (Val(Reg_PA) + Val(Reg_CA) + Val(PreEncoded_PA) + Val(PreEncoded_CA)) >= 1 Then
    FileCopy App.Path & "\Regular\Ref.dbf", "C:\Windows\Temp\" & DateTimeToday & "\Regular\Ref.dbf"
    FileCopy App.Path & "\Regular\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\Regular\Packing.dbf"
    FileCopy App.Path & "\Regular\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\Regular\SortRT.txt"
    
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\" & Format(Now, "YYYY") & "\"
End If

If (Val(PreEncoded_PA) + Val(PreEncoded_CA)) >= 1 Then
    FileCopy App.Path & "\Regular\PreEncoded\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\Regular\PreEncoded\Packing.dbf"
    FileCopy App.Path & "\Regular\PreEncoded\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\Regular\PreEncoded\SortRT.txt"

    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\" & Format(Now, "YYYY") & "\"
End If

If Val(MC) >= 1 Then
    FileCopy App.Path & "\MC\Ref.dbf", "C:\Windows\Temp\" & DateTimeToday & "\MC\Ref.dbf"
    FileCopy App.Path & "\MC\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\MC\Packing.dbf"
    FileCopy App.Path & "\MC\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\MC\SortRT.txt"
    
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\" & Format(Now, "YYYY") & "\"
End If

If (Val(CheckOne_PA) + Val(CheckOne_CA)) >= 1 Then
    FileCopy App.Path & "\CheckOne\Ref.dbf", "C:\Windows\Temp\" & DateTimeToday & "\CheckOne\Ref.dbf"
    FileCopy App.Path & "\CheckOne\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\CheckOne\Packing.dbf"
    FileCopy App.Path & "\CheckOne\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\CheckOne\SortRT.txt"
    
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\" & Format(Now, "YYYY") & "\"
End If

If (Val(CheckPower_PA) + Val(CheckPower_CA)) >= 1 Then
    FileCopy App.Path & "\CheckPower\Ref.dbf", "C:\Windows\Temp\" & DateTimeToday & "\CheckPower\Ref.dbf"
    FileCopy App.Path & "\CheckPower\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\CheckPower\Packing.dbf"
    FileCopy App.Path & "\CheckPower\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\CheckPower\SortRT.txt"
    
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\" & Format(Now, "YYYY") & "\"
End If

If Val(GiftCheck) >= 1 Then
    FileCopy App.Path & "\GiftCheck\Ref.dbf", "C:\Windows\Temp\" & DateTimeToday & "\GiftCheck\Ref.dbf"
    FileCopy App.Path & "\GiftCheck\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\GiftCheck\Packing.dbf"
    FileCopy App.Path & "\GiftCheck\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\GiftCheck\SortRT.txt"
    
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\" & Format(Now, "YYYY") & "\"
End If

If Val(MC_Continues) >= 1 Then
    FileCopy App.Path & "\MC\Continues\Branches.dbf", "C:\Windows\Temp\" & DateTimeToday & "\MC\Continues\Branches.dbf"
    FileCopy App.Path & "\MC\Continues\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\MC\Continues\Packing.dbf"
    FileCopy App.Path & "\MC\Continues\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\MC\Continues\SortRT.txt"
    
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\CONTINUOUS_MC\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\CONTINUOUS_MC\" & Format(Now, "YYYY") & "\"
End If

If Val(Customized) >= 1 Then
    FileCopy App.Path & "\Customized\Packing.dbf", "C:\Windows\Temp\" & DateTimeToday & "\Customized\Packing.dbf"
    FileCopy App.Path & "\Customized\SortRT.txt", "C:\Windows\Temp\" & DateTimeToday & "\Customized\SortRT.txt"
    
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\CUSTOMIZED\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\CUSTOMIZED\" & Format(Now, "YYYY") & "\"
End If


If Val(Charge_Slip) >= 1 Then
    If CodesOnly = True Then Temp_Zip = Resting_Folder & "\Zips\Codes\SBTC\Charge_Slip\" & Format(Now, "YYYY") & "\"
    If CodesOnly = False Then Temp_Zip = Resting_Folder & "\Zips\SBTC\Charge_Slip\" & Format(Now, "YYYY") & "\"
End If


FSO.CopyFile App.Path & "\SecurityBank.exe", "C:\Windows\Temp\" & DateTimeToday & "\SecurityBank.exe", True
'End For Zip
    
    
    
    
    
    
    
    
'For Batch
Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Batch) FROM Batch"
dbfRecordset.Open SQL, DBFConnector, 1, 1

File_Batch = ""

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Temp = dbfRecordset.Fields(0)
    
    If File_Batch = "" Then
        File_Batch = Temp
    Else
        File_Batch = File_Batch & "_" & Temp
    End If
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop

If File_Batch = "" Then File_Batch = Batch
'End For Batch










'For FTP Status

If Dir(App.Path & "\HashTotal\FTP\", vbDirectory) = "" Then MkDir App.Path & "\HashTotal\FTP\"


Close #1
Open App.Path & "\temp.bat" For Output As #1
Print #1, "rmdir /S /Q " & App.Path & "\HashTotal\FTP\" & Batch
Close #1


Shell App.Path & "\Temp.bat"



DateTimeToday_Temp = Format(Now, "MM/DD/YYYY HH:MM:SS")


Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT DISTINCT(Orig),New FROM Temp1"
dbfRecordset.Open SQL, DBFConnector, 1, 1

LoopCount = 0
Do Until LoopCount = dbfRecordset.RecordCount
    Orig_FileName = dbfRecordset.Fields(0)
    New_FileName = dbfRecordset.Fields(1)
    
    
    
    Close #1
    If Dir(App.Path & "\HashTotal\FTP\" & Batch, vbDirectory) = "" Then MkDir App.Path & "\HashTotal\FTP\" & Batch
    Open App.Path & "\HashTotal\FTP\" & Batch & "\" & Orig_FileName & "_Status.txt" For Output As #1
    
    
    
    Set dbfRecordset1 = CreateObject("ADODB.Recordset")
    SQL = "SELECT AccountNo, Name1, Name2, ChkType , BRSTN, Address1, StartSN1, PcsPerBook, OrderQty  FROM SBTC WHERE FileName = '" & New_FileName & ".txt' ORDER BY PKey"
    dbfRecordset1.Open SQL, DBFConnector, 1, 1

    LoopCount1 = 0
    Do Until LoopCount1 = dbfRecordset1.RecordCount
        AccountNo = dbfRecordset1.Fields(0)
        
        If Len(dbfRecordset1.Fields(1)) >= 1 Then
            Name1 = dbfRecordset1.Fields(1)
        Else
            Name1 = ""
        End If
        
        If Len(dbfRecordset1.Fields(2)) >= 1 Then
            Name2 = dbfRecordset1.Fields(2)
        Else
            Name2 = ""
        End If
        
        ChkType = dbfRecordset1.Fields(3)
        BRSTN = dbfRecordset1.Fields(4)
        Address1 = dbfRecordset1.Fields(5)
        StartingSerial = dbfRecordset1.Fields(6)
        PcsPerBook = dbfRecordset1.Fields(7)
        OrderQty = dbfRecordset1.Fields(8)
        
        EndingSerial = Val(StartingSerial) + (Val(PcsPerBook) * Val(OrderQty)) - 1
        
        
        Print #1, AccountNo & "|" _
                & Trim(Name1 & " " & Name2) & "|" _
                & ChkType & "|" _
                & OrderQty & "|" _
                & BRSTN & "|" _
                & Address1 & "|" _
                & "For Processing" & "|" _
                & DateTimeToday_Temp & "|" _
                & "|" _
                & DateTimeToday_Temp & "|" _
                & "|" _
                & Format(Val(StartingSerial), "0000000000") & " - " & Format(Val(EndingSerial), "0000000000") & "|" _
                & "|"
                
        
        
        dbfRecordset1.MoveNext
        LoopCount1 = LoopCount1 + 1
    Loop
    Close #1
    
    
    
    
    dbfRecordset.MoveNext
    LoopCount = LoopCount + 1
Loop
'End For FTP Status












'For Zip File
Dim ProgramExecute As String

ProgramExecute = """" & WinZipLocation & """" & " -u -r -p " & """" & App.Path & "\AFT" & "_" & File_Batch & "_Process.by_" & ProcessBy & "__Checked.By_" & CheckedBy & ".zip" & """" & " C:\Windows\Temp\" & DateTimeToday & "\*.*"
TaskID = Shell(ProgramExecute, vbNormalFocus)


hProcess = OpenProcess(SYNCHRONIZE, True, TaskID)
Call WaitForSingleObject(hProcess, WAIT_INFINITE)
CloseHandle hProcess
'End For Zip File










'For Zip File Packing List
ProgamExecute = """" & WinZipLocation & """" & " -u -r -p " & """" & App.Path & "\HashTotal\" & Batch & ".zip" & """" & " C:\Windows\Temp\" & DateTimeToday & "\" & Batch & "\" & "*.*"

TaskID = Shell(ProgamExecute, vbNormalFocus)

hProcess = OpenProcess(SYNCHRONIZE, True, TaskID)
Call WaitForSingleObject(hProcess, WAIT_INFINITE)
CloseHandle hProcess
'End For Zip File Packing List





'Copy the Zip File
CreateDirectory (Temp_Zip)

FileCopy App.Path & "\AFT" & "_" & File_Batch & "_Process.by_" & ProcessBy & "__Checked.By_" & CheckedBy & ".zip", Temp_Zip & "\AFT" & "_" & File_Batch & "_Process.by_" & ProcessBy & "__Checked.By_" & CheckedBy & ".zip"
'End Copy the Zip File
End Function
  
  
  
  

  
  
  
Sub SaveError(Description)
Set DBFConnector = CreateObject("ADODB.Connection")

DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "INSERT INTO Errors (Errors) VALUES ('" & Description & "')"
dbfRecordset.Open SQL, DBFConnector, 1, 1
End Sub


Function DeleteDBF(FileName, FolderName)
' First delete all the records
Dim Conn1, cmd1
Set Conn1 = New ADODB.Connection
Conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\" & FolderName & "\;"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
cmd1.ActiveConnection = Conn1
cmd1.CommandText = "Delete From " & FileName
cmd1.Execute

Conn1.Close
Set Conn1 = Nothing

' Now Pack the table to shrink its size
Set Conn1 = New ADODB.Connection
Conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\" & FolderName & "\;"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = Conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack " & FileName & ".dbf"
cmd1.Execute
Conn1.Close
'End Delete the Packing.dbf
End Function


Function SortRT(FolderName)

    Dim GrandTotal
    Dim PageNo As String
    Dim LineNumber

    GrandTotal = 0
    PageNo = 1
    LineNumber = 0

    Close #1
    Open App.Path & "\" & FolderName & "\SortRT.txt" For Output As #1

    Dim DBFConnector
    Set DBFConnector = CreateObject("ADODB.Connection")
    DBFConnector.Open ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\;Extended properties=dBase III")
    DBFConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient


    Dim dbfRecordset
    Set dbfRecordset = CreateObject("ADODB.Recordset")
    
    If FolderName = "Regular" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE (ChkType = 'A' AND FormType = '05') or (ChkType = 'B' AND FormType = '16') GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "Regular\PreEncoded" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE (ChkType = 'AA' AND FormType = '05') or (ChkType = 'BB' AND FormType = '16') GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "MC" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE ChkType = 'MC' AND FormType = '20' GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "CheckOne" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE ChkType = 'F' AND (FormType = '25' or FormType = '26') GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "CheckPower" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE ChkType = 'E' AND (FormType = '23' or FormType = '22') GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "GiftCheck" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE ChkType = 'GC' AND FormType = '20' GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "MC\Continues" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE ChkType = 'MC_1' AND FormType = '00' GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "Customized" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE ChkType = 'CUSTOM' AND FormType = '00' GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    If FolderName = "Charge_Slip" Then SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty), FormType FROM SBTC WHERE ChkType = 'CS' GROUP BY ChkType, BRSTN, Address1,FormType  ORDER BY ChkType, BRSTN"
    
    dbfRecordset.Open SQL, DBFConnector, 1, 1
    
    Dim LoopCount
    LoopCount = 0
    Do Until LoopCount = dbfRecordset.RecordCount
        If LoopCount = 0 Or LineNumber >= 50 Then
            Print #1, ""
            If PageNo <> 1 Then Print #1, ""

            Print #1, "    Page No. " & PageNo
            Print #1, "    " & Format(Now, "Mmm. DD, YYYY")
            Print #1, "                             Summary of RT nos / # of Books"
            
            
            If FolderName = "Regular" Then Print #1, "                               SBTC - Regular Checks"
            If FolderName = "Regular\PreEncoded" Then Print #1, "                               SBTC - PreEncoded Checks"
            If FolderName = "MC" Then Print #1, "                               SBTC - Manager's Checks"
            If FolderName = "CheckOne" Then Print #1, "                               SBTC - Check One"
            If FolderName = "CheckPower" Then Print #1, "                               SBTC - Check Power"
            If FolderName = "GiftCheck" Then Print #1, "                               SBTC - Gift Check"
            If FolderName = "MC\Continues" Then Print #1, "                               SBTC - Manager's Check Continues"
            If FolderName = "Customized" Then Print #1, "                               SBTC - Customized Checks"
            If FolderName = "Charge_Slip" Then Print #1, "                               SBTC - Charge Slip"
            
            
            Print #1, ""
            Print #1, "    ACCTNO       QTY BRANCH                 ACCOUNT NAME"
            Print #1, ""



            PageNo = PageNo + 1
            LineNumber = 0
        End If



        
        ChkType = dbfRecordset.Fields(0).Value
        BRSTN = dbfRecordset.Fields(1).Value
        Address1 = dbfRecordset.Fields(2).Value
        Subtotal = Val(dbfRecordset.Fields(3).Value)
        FormType = dbfRecordset.Fields(4).Value
        
        

        Print #1, ""
        Print #1, "   ** CHECK TYPE/BRSTN/BATCH # ---->  " & Replace(Replace(Replace(ChkType, "AA", "A"), "BB", "B"), "MC_1", "MC Continues") & "/" & BRSTN & "/" & FinalBatch
        Print #1, "   ** Branch: " & Address1
        LineNumber = LineNumber + 3



        'For Details
        Set dbfRecordset1 = CreateObject("ADODB.Recordset")
        SQL = "SELECT AccountNo, Name1, Name2, OrderQty FROM SBTC WHERE ChkType = '" & ChkType & "' AND BRSTN = '" & BRSTN & "' AND Address1 = '" & Replace(Address1, "'", "''") & "' AND FormType = '" & FormType & "' ORDER BY AccountNo, Name1, Name2"
        dbfRecordset1.Open SQL, DBFConnector, 1, 1

        Dim LoopCount1
        LoopCount1 = 0
        Do Until LoopCount1 = dbfRecordset1.RecordCount
            AccountNo = dbfRecordset1.Fields(0).Value

            If Len(dbfRecordset1.Fields(1)) >= 1 Then
                Name1 = dbfRecordset1.Fields(1).Value
            Else
                Name1 = ""
            End If

            If Len(dbfRecordset1.Fields(2)) >= 1 Then
                Name2 = dbfRecordset1.Fields(2).Value
            Else
                Name2 = ""
            End If
            
            If FolderName = "MC" Then Name1 = ""
            
            
            OrderQty = Val(dbfRecordset1.Fields(3).Value)

            Do Until Len(OrderQty) = 4
                OrderQty = " " & OrderQty
            Loop

            Print #1, "    " & AccountNo & OrderQty & " " & Name1
            LineNumber = LineNumber + 1

            If Name2 <> "" Then
                Print #1, "                     " & Name2
                LineNumber = LineNumber + 1
            End If

            If LineNumber >= 50 Then
                Print #1, ""
                If PageNo <> 1 Then Print #1, ""

                Print #1, "    Page No. " & PageNo
                Print #1, "    " & Format(Now, "Mmm. DD, YYYY")
                Print #1, "                             Summary of RT nos / # of Books"
                
                If FolderName = "Regular" Then Print #1, "                               SBTC - Regular Checks"
                If FolderName = "Regular\PreEncoded" Then Print #1, "                               SBTC - PreEncoded Checks"
                If FolderName = "MC" Then Print #1, "                               SBTC - Manager's Checks"
                If FolderName = "CheckOne" Then Print #1, "                               SBTC - Check One"
                If FolderName = "CheckPower" Then Print #1, "                               SBTC - Check Power"
                If FolderName = "GiftCheck" Then Print #1, "                               SBTC - Gift Check"
                If FolderName = "MC\Continues" Then Print #1, "                               SBTC - Manager's Check Continues"
                If FolderName = "Customized" Then Print #1, "                               SBTC - Customized Checks"
                If FolderName = "Charge_Slip" Then Print #1, "                               SBTC - Charge Slip"
                
                
                Print #1, ""
                Print #1, "    ACCTNO       QTY BRANCH                 ACCOUNT NAME"
                Print #1, ""

                PageNo = PageNo + 1
                LineNumber = 0
            End If

            dbfRecordset1.MoveNext
            LoopCount1 = LoopCount1 + 1
        Loop
        'End For Details

        Print #1, ""
        Print #1, "    Sub Total: " & Subtotal
        LineNumber = LineNumber + 2

        GrandTotal = GrandTotal + Subtotal

        dbfRecordset.MoveNext
        LoopCount = LoopCount + 1
    Loop

    Print #1, ""
    Print #1, "    Grand Total: " & GrandTotal
    
    Close #1

End Function

