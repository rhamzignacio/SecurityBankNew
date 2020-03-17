Public Class frmMain

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dlgBrowse.FileName = ""
        dlgBrowse.Filter = "ASC File | *.ASC"
        dlgBrowse.ShowDialog()

        If dlgBrowse.CheckFileExists = False Or dlgBrowse.FileName = "" Then Exit Sub


        If MsgBox("Are you sure you want to process " & dlgBrowse.SafeFileName & "?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then Exit Sub

        'Back Up Ref
        FileCopy(LocationRefDBF() & "\Ref.dbf", application.startuppath & "\Ref_Before\" & Replace(Replace(Replace(FormatDateTime(Now, DateFormat.GeneralDate), "/", ""), ":", ""), " ", "_") & ".dbf")

        If Dir(application.startuppath & "\Ref.dbf") <> "" Then Kill(application.startuppath & "\Ref.dbf")
        FileCopy(LocationRefDBF() & "\Ref.dbf", application.startuppath & "\Ref.dbf")
        'End Back Up Ref

        'Create DBase Temp
        If Dir(application.startuppath & "\Temp.dbf") <> "" Then Kill(application.startuppath & "\Temp.dbf")

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & application.startuppath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "CREATE TABLE Temp (ChkType Varchar(200), BRSTN Varchar(200), AccountNo Varchar(200), OrderQty Varchar (200), Name1 Varchar(200), Name2 Varchar(200), Address1 Varchar(200), Address2 Varchar(200), Address3 Varchar(200), Address4 Varchar(200), Address5 Varchar(200), Address6 Varchar(200))"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)
        'End Create DBase Temp

        'Read the File Name
        FileClose(1)
        FileOpen(1, dlgBrowse.FileName, OpenMode.Input)

        Do Until EOF(1)
            Dim LineInputData As String = LineInput(1)

            Dim ChkType As String = Mid(LineInputData, 1, 1)
            Dim BRSTN As String = Mid(LineInputData, 2, 9)
            Dim AccountNo As String = Mid(LineInputData, 12, 12)
            Dim Name1 As String = Trim(Mid(LineInputData, 24, 57))
            Dim OrderQty As String = Val(Trim(Mid(LineInputData, 83, 2)))

            If ChkType = "A" Or ChkType = "B" Then
                dbfRecordset = CreateObject("ADODB.Recordset")
                Sql = "INSERT INTO Temp (ChkType, BRSTN, AccountNo, OrderQty, Name1) VALUES ('" _
                        & ChkType & "','" & BRSTN & "','" & AccountNo & "','" & OrderQty & "','" & Replace(Name1, "'", "''") & "')"
                dbfRecordset.Open(Sql, dbfConnector, 1, 1)
            End If
        Loop
        FileClose(1)
        'End Read the File Name

        'Check the Ref.dbf and Checkdat.mdb
        Dim Conn As New ADODB.Connection
        Dim Rs As New ADODB.Recordset

        Conn = New ADODB.Connection
        Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=C:\Checktho\Checkdat.mdb; Jet OLEDB:Database Password=CorpCaptive;"
        Conn.Open()

        dbfRecordset = CreateObject("ADODB.Recordset")
        Sql = "SELECT DISTINCT(ChkType), BRSTN FROM Temp"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim LoopCount As String = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim ChkType As String = dbfRecordset.fields(0).value
            Dim BRSTN As String = dbfRecordset.fields(1).value

            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
            Sql = "SELECT * FROM Ref WHERE ChkType = '" & ChkType & "' AND RTNO = '" & BRSTN & "'"
            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

            If dbfRecordset.recordcount <= 0 Then
                MsgBox("BRSTN " & BRSTN & " with ChkType " & ChkType & " does not exists on Ref.dbf", vbInformation, "Error")
                End
            End If

            Sql = "SELECT [Branch Text 1], [Branch Text 2], [Branch Text 3], [Branch Text 4], [Branch Text 5], [Branch Text 6] FROM Branch WHERE [Routing Number] = '" & BRSTN & "'"
            Rs = New ADODB.Recordset
            Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            Rs.Open(Sql, Conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

            Dim Address1 As String
            Dim Address2 As String
            Dim Address3 As String
            Dim Address4 As String
            Dim Address5 As String
            Dim Address6 As String

            If Rs.RecordCount <= 0 Then
                MsgBox("BRSTN " & BRSTN & " with ChkType " & ChkType & " does not exists on Checkdat.mdb", vbInformation, "Error")
                End
            Else

                If (IsDBNull(Rs.Fields(0).Value)) = False Then
                    Address1 = Rs.Fields(0).Value
                Else
                    Address1 = ""
                End If

                If (IsDBNull(Rs.Fields(1).Value)) = False Then
                    Address2 = Rs.Fields(1).Value
                Else
                    Address2 = ""
                End If

                If (IsDBNull(Rs.Fields(2).Value)) = False Then
                    Address3 = Rs.Fields(2).Value
                Else
                    Address3 = ""
                End If

                If (IsDBNull(Rs.Fields(3).Value)) = False Then
                    Address4 = Rs.Fields(3).Value
                Else
                    Address4 = ""
                End If

                If (IsDBNull(Rs.Fields(4).Value)) = False Then
                    Address5 = Rs.Fields(4).Value
                Else
                    Address5 = ""
                End If

                If (IsDBNull(Rs.Fields(5).Value)) = False Then
                    Address6 = Rs.Fields(5).Value
                Else
                    Address6 = ""
                End If
            End If

            dbfRecordset1 = CreateObject("ADODB.Recordset")
            Sql = "UPDATE Temp SET Address1 = '" & Replace(Address1, "'", "''") & "',  Address2 = '" & Replace(Address2, "'", "''") & "', Address3 = '" & Replace(Address3, "'", "''") & "', Address4 = '" & Replace(Address4, "'", "''") & "', Address5 = '" & Replace(Address5, "'", "''") & "', Address6 = '" & Replace(Address6, "'", "''") & "' WHERE BRSTN = '" & BRSTN & "'"
            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop
        'End Check the Ref.dbf and Checkdat.mdb

        SortRT()

        If MsgBox("SortRt.txt has been generated. Are you sure you want to process?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then Exit Sub

        Dim Batch As String = Mid(InputBox("Enter Batch Number", "", ""), 1, 8)
        If Batch = "" Then Exit Sub

        ProcessAll("A", Batch)
        ProcessAll("B", Batch)

        FileCopy(Application.StartupPath & "\Ref.dbf", LocationRefDBF() & "\Ref.dbf")

        UpdateRefModified()

        MsgBox("Data has been Processed", vbInformation, "")
        End
    End Sub



    Sub ProcessAll(ChkType, Batch)
        Dim ChkType2 As String = ""
        Dim DataCount As String = 0
        Dim BlockCount As String = 0
        Dim PcsPerBook As String = 0
        Dim FormatSerial As String = ""
        Dim Temp As String = ""

        If ChkType = "A" Then
            ChkType2 = "P"
            PcsPerBook = 20
            FormatSerial = "0000000"
        End If

        If ChkType = "B" Then
            ChkType2 = "C"
            PcsPerBook = 20
            FormatSerial = "0000000000"
        End If

        Dim DoBlock As System.IO.StreamWriter
        DoBlock = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Block" & ChkType2 & ".txt", False)

        Dim PrinterFile As System.IO.StreamWriter
        PrinterFile = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\PrinterFile" & ChkType2 & ".txt", False)

        'Read the Contents
        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "DELETE FROM Packing WHERE ChkType = '" & ChkType & "'"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        CompactDBF("Packing")

        dbfRecordset = CreateObject("ADODB.Recordset")
        Sql = "SELECT BRSTN, AccountNo, OrderQty, Name1, Name2, Address1, Address2, Address3, Address4, Address5, Address6 FROM Temp WHERE ChkType = '" & ChkType & "'"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim LoopCount As String = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim BRSTN As String = dbfRecordset.fields(0).value
            Dim AccountNo As String = dbfRecordset.fields(1).value
            Dim OrderQty As String = dbfRecordset.fields(2).value

            Dim Name1 As String
            If IsDBNull(dbfRecordset.fields(3).value) = False Then
                Name1 = dbfRecordset.fields(3).value
            Else
                Name1 = ""
            End If

            Dim Name2 As String
            If IsDBNull(dbfRecordset.fields(4).value) = False Then
                Name2 = dbfRecordset.fields(4).value
            Else
                Name2 = ""
            End If

            Dim Address1 As String
            If IsDBNull(dbfRecordset.fields(5).value) = False Then
                Address1 = dbfRecordset.fields(5).value
            Else
                Address1 = ""
            End If

            Dim Address2 As String
            If IsDBNull(dbfRecordset.fields(6).value) = False Then
                Address2 = dbfRecordset.fields(6).value
            Else
                Address2 = ""
            End If

            Dim Address3 As String
            If IsDBNull(dbfRecordset.fields(7).value) = False Then
                Address3 = dbfRecordset.fields(7).value
            Else
                Address3 = ""
            End If

            Dim Address4 As String
            If IsDBNull(dbfRecordset.fields(8).value) = False Then
                Address4 = dbfRecordset.fields(8).value
            Else
                Address4 = ""
            End If

            Dim Address5 As String
            If IsDBNull(dbfRecordset.fields(9).value) = False Then
                Address5 = dbfRecordset.fields(9).value
            Else
                Address5 = ""
            End If

            Dim Address6 As String
            If IsDBNull(dbfRecordset.fields(10).value) = False Then
                Address6 = dbfRecordset.fields(10).value
            Else
                Address6 = ""
            End If

            Dim StartingSerial As String = GetStartingSerialAndUpdate(BRSTN, ChkType, OrderQty, PcsPerBook)

            Do Until Val(OrderQty) = 0

                'For Do-Block
                If Val(DataCount) Mod 32 = 0 Then
                    If DataCount <> 0 Then DoBlock.WriteLine("")
                    DoBlock.WriteLine("        Page No. " & (DataCount / 32) + 1)
                    DoBlock.WriteLine("        " & FormatDateTime(Now, DateFormat.ShortDate))

                    If ChkType = "A" Then DoBlock.WriteLine("                     SBTC  - SUMMARY OF BLOCK - PERSONAL")
                    If ChkType = "B" Then DoBlock.WriteLine("                     SBTC  - SUMMARY OF BLOCK - COMMERCIAL")

                    DoBlock.WriteLine("                                   (20 PCS PER BOOKLET)")
                    DoBlock.WriteLine("                                   ALL MANUAL ENCODED")
                    DoBlock.WriteLine("")
                    DoBlock.WriteLine("        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.")

                End If

                If Val(DataCount Mod 4) = 0 Then
                    BlockCount = BlockCount + 1
                    Do Until Len(BlockCount) >= 13
                        BlockCount = " " & BlockCount
                    Loop

                    DoBlock.WriteLine("")
                    DoBlock.WriteLine("       ** BLOCK " & Val(BlockCount))
                End If

                Temp = Format(Val(StartingSerial), FormatSerial)

                Do Until Len(Temp) >= 11
                    Temp = Temp & " "
                Loop
                DoBlock.WriteLine(BlockCount & " " & BRSTN & "   " & AccountNo & "     " & Temp & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial))
                'End For Do-Block

                'For Printer File
                PrinterFile.WriteLine("3")
                PrinterFile.WriteLine(BRSTN)
                PrinterFile.WriteLine(AccountNo)
                PrinterFile.WriteLine(Format(Val(StartingSerial) + Val(PcsPerBook), FormatSerial))
                PrinterFile.WriteLine("A")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine(Mid(BRSTN, 1, 5))
                PrinterFile.WriteLine(" " & Mid(BRSTN, 6, 4))
                PrinterFile.WriteLine(Format(Val(AccountNo), "000-000000-000"))
                PrinterFile.WriteLine(Name1)
                PrinterFile.WriteLine("SN")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine(Name2)
                PrinterFile.WriteLine("C")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine(Address1)
                PrinterFile.WriteLine(Address2)
                PrinterFile.WriteLine(Address3)
                PrinterFile.WriteLine(Address4)
                PrinterFile.WriteLine(Address5)
                PrinterFile.WriteLine(Address6)
                PrinterFile.WriteLine("SECURITY BANK")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine("")
                PrinterFile.WriteLine(Temp)
                PrinterFile.WriteLine(Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial))
                'End For Printer File

                'Packing.dbf
                Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                Sql = "INSERT INTO Packing (BatchNo, Block, RT_NO, Branch, Acct_No, Acct_No_P, ChkType, Acct_Name1, Acct_Name2, NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E) VALUES ('" _
                        & Mid(Batch, 1, 8) & "','" & Val(BlockCount) & "','" & BRSTN & "','" & Address1 & "','" & AccountNo & "','" & Format(Val(AccountNo), "000-000000-000") & "','" & ChkType & "','" & Mid(Name1, 1, 30) & "','" & Mid(Name2, 1, 30) & "','1','" & Temp & "','" & Temp & "','" & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) & "','" & Format(Val(StartingSerial) + Val(PcsPerBook) - 1, FormatSerial) & "')"
                dbfRecordset1.Open(Sql, dbfConnector, 1, 1)
                'End Packing.dbf

                DataCount = Val(DataCount) + 1
                StartingSerial = Val(StartingSerial) + Val(PcsPerBook)
                OrderQty = OrderQty - 1
            Loop

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop
        'End Read the Contents


        DoBlock.Close()
        PrinterFile.Close()

        PackingList(ChkType, ChkType2, Batch)
    End Sub

    Sub PackingList(ChkType, ChkType2, Batch)

        Dim PackingListText As System.IO.StreamWriter
        PackingListText = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Packing" & ChkType & ".txt", False)

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "SELECT Branch, RT_NO, Count(Branch) FROM Packing WHERE ChkType = '" & ChkType & "' GROUP BY Branch, RT_NO ORDER BY RT_NO"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim LoopCount = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim BranchName As String = dbfRecordset.fields(0).value
            Dim BRSTN As String = dbfRecordset.fields(1).value
            Dim OrderQty As String = dbfRecordset.fields(2).value

            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
            Sql = "SELECT Acct_No_P, Acct_Name1, Acct_Name2, CK_NO_B, CK_NO_E FROM Packing WHERE RT_No = '" & BRSTN & "' AND ChkType= '" & ChkType & "' ORDER BY Acct_No_P, Acct_Name1"
            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

            If LoopCount <> 0 Then PackingListText.WriteLine("")

            PackingListText.WriteLine("")
            PackingListText.WriteLine("  Page No. " & Val(LoopCount) + 1)
            PackingListText.WriteLine("  " & FormatDateTime(Now, DateFormat.ShortDate))
            PackingListText.WriteLine("                                CAPTIVE PRINTING CORPORATION")

            If ChkType = "A" Then PackingListText.WriteLine("                               SBTC  - Personal Checks Summary")
            If ChkType = "B" Then PackingListText.WriteLine("                               SBTC - Commercial Checks Summary")

            PackingListText.WriteLine("")
            PackingListText.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #")
            PackingListText.WriteLine("")
            PackingListText.WriteLine("")
            PackingListText.WriteLine(" ** ORDERS OF BRSTN " & BRSTN & " " & BranchName)
            PackingListText.WriteLine("")
            PackingListText.WriteLine(" * BATCH #: " & Mid(Batch, 1, 8))

            Dim LoopCount1 = 0
            Do Until LoopCount1 = dbfRecordset1.recordcount
                Dim AccountNo As String = dbfRecordset1.fields(0).value

                Dim Name1 As String
                If IsDBNull(dbfRecordset1.fields(1).value) = False Then
                    Name1 = dbfRecordset1.fields(1).value
                Else
                    Name1 = ""
                End If

                Dim Name2 As String
                If IsDBNull(dbfRecordset1.fields(2).value) = False Then
                    Name2 = dbfRecordset1.fields(2).value
                Else
                    Name2 = ""
                End If

                Dim StartingSerial As String = dbfRecordset1.fields(3).value
                Dim EndingSerial As String = dbfRecordset1.fields(4).value

                Do Until Len(Name1) >= 35
                    Name1 = Name1 & " "
                Loop

                Do Until Len(StartingSerial) >= 11
                    StartingSerial = StartingSerial & " "
                Loop

                PackingListText.WriteLine("  " & AccountNo & "  " & Name1 & "1 " & ChkType & "  " & StartingSerial & EndingSerial)
                If Name2 <> "" Then PackingListText.WriteLine("                  " & Name2)

                dbfRecordset1.movenext()
                LoopCount1 = LoopCount1 + 1
            Loop

            PackingListText.WriteLine("")
            PackingListText.WriteLine("  * * * Sub Total * * * " & OrderQty)

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop

        PackingListText.Close()

    End Sub


    Function GetStartingSerialAndUpdate(BRSTN, ChkType, OrderQty, PcsPerBook)
        GetStartingSerialAndUpdate = ""

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "SELECT LastNo FROM Ref WHERE RTNO = '" & BRSTN & "' AND ChkType = '" & ChkType & "'"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim EndingSerial As String = dbfRecordset.fields(0).value
        GetStartingSerialAndUpdate = Val(EndingSerial) + 1

        Dim NewEndingSerial As String = Val(EndingSerial) + (Val(PcsPerBook) * Val(OrderQty))

        dbfRecordset = CreateObject("ADODB.Recordset")
        Sql = "UPDATE REF SET LastNo = '" & NewEndingSerial & "', [Date] = '" & FormatDateTime(Now, DateFormat.ShortDate) & "' WHERE RTNO = '" & BRSTN & "' AND ChkType = '" & ChkType & "'"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

    End Function

    Sub SortRT()
        Dim PageNo As String = 1
        Dim NumberOfLines = 0

        Dim PrintSortRT As System.IO.StreamWriter
        PrintSortRT = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Sort.txt", False)

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & application.startuppath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "SELECT Address1, BRSTN, ChkType, COUNT(BRSTN) FROM TEMP GROUP BY ADDRESS1, BRSTN, ChkType ORDER BY ChkType, BRSTN"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim LoopCount As String = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim Address1 As String = dbfRecordset.fields(0).value
            Dim BRSTN As String = dbfRecordset.fields(1).value
            Dim ChkType As String = dbfRecordset.fields(2).value
            Dim SubTotal As String = dbfRecordset.fields(3).value

            'PrintSortRT.WriteLine("")
            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
            Sql = "SELECT AccountNo, Name1, Name2, OrderQty FROM Temp WHERE BRSTN = '" & BRSTN & "' AND Address1 = '" & Address1 & "' AND ChkType = '" & ChkType & "' ORDER BY AccountNo, Name1, Name2"
            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

            Dim Name1 As String
            Dim Name2 As String

            Dim LoopCount1 As String = 0
            Do Until LoopCount1 = dbfRecordset1.Recordcount

                If (NumberOfLines Mod 50 = 0) Then
                    If NumberOfLines <> 0 Then PrintSortRT.WriteLine("")

                    PrintSortRT.WriteLine("")
                    PrintSortRT.WriteLine("    Page No. " & PageNo)
                    PrintSortRT.WriteLine("    " & FormatDateTime(Now, DateFormat.GeneralDate))
                    PrintSortRT.WriteLine("                          SBC - Summary of RT nos / # of Books")
                    PrintSortRT.WriteLine("")
                    PrintSortRT.WriteLine("    ACCTNO       QTY BRANCH                 ACCOUNT NAME")

                End If

                Dim AccountNo As String = Format(Val(dbfRecordset1.fields(0).value), "000000000000")

                If (IsDBNull(dbfRecordset1.fields(1).value)) = False Then
                    Name1 = dbfRecordset1.fields(1).value
                Else
                    Name1 = ""
                End If

                If (IsDBNull(dbfRecordset1.fields(2).value)) = False Then
                    Name2 = dbfRecordset1.fields(2).value

                    NumberOfLines = NumberOfLines + 2
                Else
                    Name2 = ""

                    NumberOfLines = NumberOfLines + 1
                End If

                Dim OrderQty As String = dbfRecordset1.fields(3).value
                Do Until Len(OrderQty) = 4
                    OrderQty = " " & OrderQty
                Loop



                If LoopCount1 = 0 Then
                    PrintSortRT.WriteLine("")
                    PrintSortRT.WriteLine("   ** CHECK TYPE/BRSTN/BATCH # ---->  " & ChkType & "/" & BRSTN)
                    PrintSortRT.WriteLine("   ** BRANCH NAME ----> " & Address1)
                    PrintSortRT.WriteLine("")
                End If

                PrintSortRT.WriteLine("    " & AccountNo & "  " & OrderQty & " " & Name1)
                If Name2 <> "" Then PrintSortRT.WriteLine("                   " & Name2)


                dbfRecordset1.movenext()
                LoopCount1 = LoopCount1 + 1
            Loop

            PrintSortRT.WriteLine("")
            PrintSortRT.WriteLine(" ** Sub Total * * " & SubTotal)
            PrintSortRT.WriteLine("")

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop

        PrintSortRT.Close()


    End Sub
    Function LocationRefDBF()
        LocationRefDBF = ""

        FileClose(1)

        FileOpen(1, application.startuppath & "\Ref_Location.ini", OpenMode.Input)

        Do Until EOF(1)
            LocationRefDBF = LineInput(1)
        Loop
    End Function

    Sub UpdateRefModified()
        Dim DateTime = FileDateTime(LocationRefDBF() & "\Ref.dbf")

        Dim DateModified = Format(Val(DateTime.Month.ToString), "00") & "/" & DateTime.Day.ToString & "/" & Format(Val(DateTime.Year.ToString), "0000")
        Dim TimeModified = Format(Val(DateTime.Hour.ToString), "00") & ":" & Format(Val(DateTime.Minute.ToString), "00") & ":" & Format(Val(DateTime.Second.ToString), "00")

        Dim Conn As New ADODB.Connection
        Dim Rs As New ADODB.Recordset

        Conn = New ADODB.Connection
        Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & LocationRefDBF() & "\SBTC.captive; Jet OLEDB:Database Password=Elgae;"
        Conn.Open()

        Dim Sql = "UPDATE LastRefModified SET [Date] = '" & DateModified & "', [Time] = '" & TimeModified & "'"
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(Sql, Conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
    End Sub

    Sub CompactDBF(FileName)
        ' First delete all the records
        'Dim Conn1 = New ADODB.Connection
        'Conn1.Open("Provider=VfpOleDB.1; Data Source=" & Application.StartupPath & "\;")

        'Dim Cmd1 = New ADODB.Command
        'Cmd1.CommandType = ADODB.CommandTypeEnum.adCmdText
        'cmd1.ActiveConnection = conn1
        'cmd1.CommandText = "Delete From " & FileName
        'cmd1.Execute()

        'conn1.Close()
        'conn1 = Nothing

        ' Now Pack the table to shrink its size
        Dim conn1 = New ADODB.Connection
        Conn1.Open("Provider=VfpOleDB.1; Data Source=" & Application.StartupPath & "\;")

        Dim cmd1 = New ADODB.Command
        Cmd1.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmd1.ActiveConnection = conn1
        cmd1.CommandText = "Set Exclusive On"
        cmd1.Execute()
        cmd1.CommandText = "Pack " & FileName & ".dbf"
        cmd1.Execute()
        conn1.Close()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Set the Time same to host
        Dim P As System.Diagnostics.Process
        Dim PInfo As New System.Diagnostics.ProcessStartInfo
        Dim POut As String

        PInfo.FileName = "C:\Windows\System32\Net.exe "
        PInfo.Arguments = "TIME \\192.168.0.29 /SET /Y"
        PInfo.RedirectStandardOutput = True
        PInfo.UseShellExecute = False
        PInfo.CreateNoWindow = True
        P = Diagnostics.Process.Start(PInfo)
        P.WaitForExit()
        POut = P.StandardOutput.ReadLine
    End Sub
End Class
