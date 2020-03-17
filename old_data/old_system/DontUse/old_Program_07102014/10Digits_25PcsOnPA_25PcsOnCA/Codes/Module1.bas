Attribute VB_Name = "Module1"
Function getDataLocation()
Close #99
Open App.Path & "\DataLocation.ini" For Input As #99
Do Until EOF(99)
    Line Input #99, getDataLocation
Loop
Close #99
End Function

Function getDataLocationCheckdat()
Close #99
Open App.Path & "\DataLocationCheckDat.ini" For Input As #99
Do Until EOF(99)
    Line Input #99, getDataLocationCheckdat
Loop
Close #99
End Function

Function getStartingSerialAndUpdate(BRSTN, OrderQty, ChkType)
If ChkType = "A" Then PcsPerBook = 25
If ChkType = "B" Then PcsPerBook = 25

Set DBFConnector = CreateObject("ADODB.Connection")
 
DBFConnector.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & getDataLocation & "\;Extended properties=dBase III"
DBFConnector.CursorLocation = adUseClient

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "SELECT LastNo FROM REF WHERE RTNO = '" & BRSTN & "' AND ChkType = '" & ChkType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1

getStartingSerialAndUpdate = dbfRecordset.Fields(0)

Set dbfRecordset = CreateObject("ADODB.Recordset")
SQL = "UPDATE REF SET LastNo = '" & (getStartingSerialAndUpdate + (PcsPerBook * OrderQty)) & "' WHERE RTNO = '" & BRSTN & "' AND ChkType = '" & ChkType & "'"
dbfRecordset.Open SQL, DBFConnector, 1, 1
End Function

Function DeleteDBF(FileName)
' First delete all the records
Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\;"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
cmd1.ActiveConnection = conn1
cmd1.CommandText = "Delete From " & FileName
cmd1.Execute

conn1.Close
Set conn1 = Nothing

' Now Pack the table to shrink its size
Set conn1 = New ADODB.Connection
conn1.Open _
"Provider=VfpOleDB.1;" & _
"Data Source=" & App.Path & "\;"

Set cmd1 = New ADODB.Command
cmd1.CommandType = adCmdText
Set cmd1.ActiveConnection = conn1
cmd1.CommandText = "Set Exclusive On"
cmd1.Execute
cmd1.CommandText = "Pack " & FileName & ".dbf"
cmd1.Execute
conn1.Close
End Function

Function getAddress(BRSTN, AddressNumber)
Dim Conn As ADODB.Connection
Dim Rs As ADODB.Recordset

Set Conn = New ADODB.Connection

With Conn
  .CursorLocation = adUseClient
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & getDataLocationCheckdat & "; Jet OLEDB:Database Password=CorpCaptive;"
  .Open
End With
'For Configuration
        
Set Rs = New ADODB.Recordset
strQuery = "SELECT [Branch Text 1], [Branch Text 2], [Branch Text 3], [Branch Text 4], [Branch Text 5], [Branch Text 6] FROM Branch WHERE [Routing Number] = '" & BRSTN & "'"
Rs.Open strQuery, Conn, adOpenStatic

If Len(Rs.Fields(AddressNumber - 1)) >= 1 Then
    getAddress = Rs.Fields(AddressNumber - 1)
Else
    getAddress = ""
End If
End Function
