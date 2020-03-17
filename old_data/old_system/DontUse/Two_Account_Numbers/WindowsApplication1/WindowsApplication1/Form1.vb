

Public Class Form1
    Dim FinalBatch As String
    Dim TotalErrors As String

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        If btnBrowse.Text = "Browse .ASC File" Then
            'Browse the File
            dlgBrowse.InitialDirectory = Application.StartupPath
            dlgBrowse.Filter = "ASC File | *.ASC"
            dlgBrowse.ShowDialog()

            If Dir$(dlgBrowse.FileName) = "" Then Exit Sub

            If MsgBox("Are you sure you want to Process " & dlgBrowse.SafeFileName & "?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then Exit Sub

            'End Browse the File
            TotalErrors = 0

            ReadTextFile()

            CheckAddressCheckdat()

            If TotalErrors = 0 Then
                SortRT()

                btnBrowse.Text = "Process " & FinalBatch

                DisplayDetails()

                MsgBox("Data has been Checked. No Errors Found ! ! !", vbInformation, " ")

            Else
                GenerateErrors()

                MsgBox("Unable to Process. " & TotalErrors & " Error/s Found ! ! !", vbCritical, "Error!")
                End
            End If

        Else
            If MsgBox("Are you sure you want to process Batch " & FinalBatch & "?", vbYesNo + vbInformation, "Confirm Process") = vbNo Then Exit Sub

            ProcessData()

            MsgBox("Data has been Processed", vbInformation, " ")
            End
        End If
    End Sub

    Sub ProcessData()
        ProcessMe("A")
        ProcessMe("B")
    End Sub

    Sub CompactDBF(FileName As String)
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
        conn1.Open("Provider=VfpOleDB.1; Data Source=" & Application.StartupPath & "\;")

        Dim cmd1 = New ADODB.Command
        cmd1.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmd1.ActiveConnection = conn1
        cmd1.CommandText = "Set Exclusive On"
        cmd1.Execute()
        cmd1.CommandText = "Pack " & FileName & ".dbf"
        cmd1.Execute()
        conn1.Close()

    End Sub

    Sub ProcessMe(ChkType)
        Dim Temp As String = ""
        Dim TotalData = 0
        Dim BlockCount As String = 0
        Dim PcsPerBook = 0
        Dim ChkType2 = ""
        Dim FormatSerial = ""
        Dim PrinterFileName = ""

        If ChkType = "A" Then
            PcsPerBook = 50
            ChkType2 = "P"
            FormatSerial = "0000000"
            PrinterFileName = "SB" & Format(Val(Now.Month), "00") & Format(Val(Now.Day), "00") & "P.txt"
        End If

        If ChkType = "B" Then
            PcsPerBook = 100
            ChkType2 = "C"
            FormatSerial = "0000000000"
            PrinterFileName = "SB" & Format(Val(Now.Month), "00") & Format(Val(Now.Day), "00") & "C.txt"
        End If

        'For Configuration DBase
        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        'End For Configuration DBase
        


        Dim PrinterFile As System.IO.StreamWriter

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "DELETE FROM Packing"
        If ChkType = "A" Then dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        CompactDBF("Packing")

        Dim DoBlock As System.IO.StreamWriter
        DoBlock = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Block" & ChkType2 & ".txt", False)

        dbfRecordset = CreateObject("ADODB.Recordset")
        Sql = "SELECT BRSTN, AccountNo, Name1, Name2, Address1, Address2, Address3,Address4,Address5,Address6, OrderQty, AccountNo2 FROM SBTC WHERE ChkType = '" & ChkType & "' ORDER BY BRSTN, AccountNo, Name1, Name2"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim LoopCount = 0
        Do Until LoopCount = dbfRecordset.recordcount
            If LoopCount = 0 Then
                PrinterFile = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\" & PrinterFileName, False)
            End If

            Dim BRSTN = dbfRecordset.fields(0).value
            Dim AccountNo = dbfRecordset.fields(1).value

            Dim Name1 = ""
            If IsDBNull(dbfRecordset.fields(2).value) = False Then
                Name1 = dbfRecordset.fields(2).value
            Else
                Name1 = ""
            End If

            Dim Name2 = ""
            If IsDBNull(dbfRecordset.fields(3).value) = False Then
                Name2 = dbfRecordset.fields(3).value
            Else
                Name2 = ""
            End If

            Dim Address1 = ""
            If IsDBNull(dbfRecordset.fields(4).value) = False Then
                Address1 = dbfRecordset.fields(4).value
            Else
                Address1 = ""
            End If

            Dim Address2 = ""
            If IsDBNull(dbfRecordset.fields(5).value) = False Then
                Address2 = dbfRecordset.fields(5).value
            Else
                Address2 = ""
            End If

            Dim Address3 = ""
            If IsDBNull(dbfRecordset.fields(6).value) = False Then
                Address3 = dbfRecordset.fields(6).value
            Else
                Address3 = ""
            End If

            Dim Address4 = ""
            If IsDBNull(dbfRecordset.fields(7).value) = False Then
                Address4 = dbfRecordset.fields(7).value
            Else
                Address4 = ""
            End If

            Dim Address5 = ""
            If IsDBNull(dbfRecordset.fields(8).value) = False Then
                Address5 = dbfRecordset.fields(8).value
            Else
                Address5 = ""
            End If

            Dim Address6 = ""
            If IsDBNull(dbfRecordset.fields(9).value) = False Then
                Address6 = dbfRecordset.fields(9).value
            Else
                Address6 = ""
            End If

            Dim OrderQty = Val(dbfRecordset.fields(10).value)
            Dim AccountNo2 As String = dbfRecordset.fields(11).value
            Dim AccountNo2WithHyphen As String = Mid(AccountNo2, 1, 4) & "-" & Mid(AccountNo2, 5, 6) & "-" & Mid(AccountNo2, 11, 3)

            'For StartingSerial
            Dim StartSN = 0

            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
            Sql = "SELECT LastNo FROM Ref WHERE RTNo = '" & BRSTN & "' AND ChkTYpe = '" & ChkType & "'"
            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

            StartSN = Val(dbfRecordset1.fields(0).value) + 1
            Dim LastNo = Val(dbfRecordset1.fields(0).value) + (Val(PcsPerBook) * OrderQty)

            dbfRecordset1 = CreateObject("ADODB.Recordset")
            Sql = "UPDATE REF SET LastNo = '" & LastNo & "'WHERE RTNo = '" & BRSTN & "' AND ChkTYpe = '" & ChkType & "'"
            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)
            'End For StartingSerial

            Do Until Val(OrderQty) = 0
                If Val(TotalData) Mod 32 = 0 Then
                    DoBlock.WriteLine("")

                    Temp = (TotalData / 32) + 1

                    If Temp <> 1 Then DoBlock.WriteLine("")
                    DoBlock.WriteLine("        Page No. " & Temp)
                    DoBlock.WriteLine("        " & Now.ToShortDateString)

                    If ChkType = "A" Then DoBlock.WriteLine("                         SUMMARY OF BLOCK - PERSONAL")
                    If ChkType = "B" Then DoBlock.WriteLine("                         SUMMARY OF BLOCK - COMMERCIAL")

                    DoBlock.WriteLine("                             SBTC - Regular_Checks")

                    DoBlock.WriteLine("")
                    DoBlock.WriteLine("        BLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.")
                    DoBlock.WriteLine("")
                End If

                If Val(TotalData) Mod 4 = 0 Then
                    BlockCount = BlockCount + 1

                    DoBlock.WriteLine("")
                    DoBlock.WriteLine("       ** BLOCK " & Val(BlockCount))
                End If

                Do Until Len(BlockCount) >= 13
                    BlockCount = " " & BlockCount
                Loop

                Temp = Format(Val(StartSN), FormatSerial)
                Do Until Len(Temp) = 11
                    Temp = Temp & " "
                Loop
                DoBlock.WriteLine(BlockCount & " " & BRSTN & "   " & AccountNo & "    " & Temp & Format(Val(Temp) + Val(PcsPerBook) - 1, FormatSerial))

                'For Printer File
                PrinterFile.WriteLine("32")
                PrinterFile.WriteLine(BRSTN)
                PrinterFile.WriteLine(AccountNo)
                PrinterFile.WriteLine(Format(Val(StartSN) + Val(PcsPerBook), FormatSerial))
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
                PrinterFile.WriteLine(AccountNo2WithHyphen)
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
                PrinterFile.WriteLine(Trim(Temp))
                PrinterFile.WriteLine(Format(Val(StartSN) + Val(PcsPerBook) - 1, FormatSerial))
                'End For Printer File

                'FOr Packing.dbf
                Dim ChkTypeDisplay = ""
                If ChkType = "A" Then ChkTypeDisplay = "A"
                If ChkType = "B" Then ChkTypeDisplay = "B"

                dbfRecordset1 = CreateObject("ADODB.Recordset")
                Sql = "INSERT INTO Packing (BatchNo, Block, RT_NO, Branch, Acct_no, Acct_No_P, ChkType, Acct_Name1, Acct_Name2, NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E) VALUES ('" _
                    & FinalBatch & "','" & Trim(BlockCount) & "','" & BRSTN & "','" & Replace(Address1, "'", "''") & "','" & AccountNo & "','" & AccountNo2WithHyphen & "','" & ChkTypeDisplay & "','" & Replace(Name1, "'", "''") & "','" & Replace(Name2, "'", "''") & "','" & "1" & "','" & Trim(Temp) & "','" & Trim(Temp) & "','" & Format(Val(StartSN) + Val(PcsPerBook) - 1, FormatSerial) & "','" & Format(Val(StartSN) + Val(PcsPerBook) - 1, FormatSerial) & "')"
                dbfRecordset1.Open(Sql, dbfConnector, 1, 1)
                'End FOr Packing.dbf

                StartSN = Val(StartSN) + Val(PcsPerBook)

                TotalData = TotalData + 1
                OrderQty = OrderQty - 1
            Loop

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop

        DoBlock.Close()
        If LoopCount > 0 Then PrinterFile.Close()

        PackingList(ChkType, ChkType2)
    End Sub

    Sub PackingList(ChkType As String, ChkType2 As String)

        Dim PackingListText As System.IO.StreamWriter
        PackingListText = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Packing" & ChkType & ".txt", False)

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim SQL = "SELECT Branch, RT_NO, Count(Branch) FROM Packing WHERE ChkType = '" & ChkType & "' GROUP BY Branch, RT_NO ORDER BY RT_NO"
        dbfRecordset.Open(SQL, dbfConnector, 1, 1)

        Dim LoopCount = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim BranchName As String = dbfRecordset.fields(0).value
            Dim BRSTN As String = dbfRecordset.fields(1).value
            Dim OrderQty As String = dbfRecordset.fields(2).value

            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
            SQL = "SELECT Acct_No_P, Acct_Name1, Acct_Name2, CK_NO_B, CK_NO_E FROM Packing WHERE RT_No = '" & BRSTN & "' AND ChkType= '" & ChkType & "' ORDER BY Acct_No_P, Acct_Name1, CK_NO_B"
            dbfRecordset1.Open(SQL, dbfConnector, 1, 1)

            If LoopCount <> 0 Then PackingListText.WriteLine("")

            PackingListText.WriteLine("")
            PackingListText.WriteLine("  Page No. " & Val(LoopCount) + 1)
            PackingListText.WriteLine("  " & FormatDateTime(Now, DateFormat.ShortDate))
            PackingListText.WriteLine("                                CAPTIVE PRINTING CORPORATION")

            If ChkType = "A" Then PackingListText.WriteLine("                     EWB  - Personal Checks Summary")
            If ChkType = "B" Then PackingListText.WriteLine("                     EWB - Commercial Checks Summary")

            PackingListText.WriteLine("")
            PackingListText.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #")
            PackingListText.WriteLine("")
            PackingListText.WriteLine("")
            PackingListText.WriteLine(" ** ORDERS OF BRSTN " & BRSTN & " " & BranchName)
            PackingListText.WriteLine("")
            PackingListText.WriteLine(" * BATCH #: " & FinalBatch)

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

                Dim ChkTypeDisplay = ""
                If ChkType = "A" Then ChkTypeDisplay = "A"
                If ChkType = "B" Then ChkTypeDisplay = "B"

                PackingListText.WriteLine("  " & AccountNo & "  " & Name1 & "1 " & ChkTypeDisplay & "  " & StartingSerial & EndingSerial)
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



    Sub DisplayDetails()
        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "SELECT ChkType, SUM(OrderQty) FROM SBTC GROUP BY ChkTYpe ORDER BY ChkType"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim Display As String = ""

        Dim LoopCount = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim ChkType As String = dbfRecordset.fields(0).value
            Dim OrderQty As String = dbfRecordset.fields(1).value

            Dim Description As String = ""
            If ChkType = "A" Then Description = "Personal"
            If ChkType = "B" Then Description = "Commercial"

            If Display = "" Then
                Display = Description & ": " & OrderQty
            Else
                Display = Display & vbNewLine & Description & ": " & OrderQty
            End If

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop

        lblTotal.Text = Display

    End Sub

    Sub GenerateErrors()

        Dim Errors_Text As System.IO.StreamWriter
        Errors_Text = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\Errors.txt", False)

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "SELECT Errors FROM Errors"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Errors_Text.WriteLine(Now.ToShortDateString)
        Errors_Text.WriteLine(TotalErrors & " Errors Found ! ! !")
        Errors_Text.WriteLine("")

        Dim LoopCount = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim Errors As String = dbfRecordset.fields(0)

            Errors_Text.WriteLine(Errors)

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop

        Errors_Text.Close()

    End Sub

    Sub CheckAddressCheckdat()
        'Check for Errors
        Dim TotalCheckdat = 0
        Dim CheckDatMDB(0 To 999999) As String
        Dim CheckDatMDBDrive(0 To 999999) As String

        FileClose(1)
        FileOpen(1, "C:\CheckDat_Location.ini", OpenMode.Input)
        Do Until EOF(1)
            Dim LineInputData As String = LineInput(1)
            Dim Drive = Mid(LineInputData, 1, 1)
            Dim Checkdat_Location = Mid(LineInputData, 3, Len(LineInputData))

            If Drive <> "" And Dir(Checkdat_Location, FileAttribute.Directory) <> "" Then
                If Dir("C:\CheckDat_" & Drive & ".mdb") <> "" Then Kill("C:\CheckDat_" & Drive & ".mdb")
                FileCopy(Checkdat_Location, "C:\CheckDat_" & Drive & ".mdb")

                CheckDatMDB(TotalCheckdat) = "C:\CheckDat_" & Drive & ".mdb"
                CheckDatMDBDrive(TotalCheckdat) = Drive

                TotalCheckdat = TotalCheckdat + 1
            End If
        Loop

        FileClose(1)


        'Checkdat.mdb
        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "SELECT DISTINCT(BRSTN), ChkType FROM SBTC"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        Dim LoopCount = 0
        Do Until LoopCount = dbfRecordset.recordcount
            Dim BRSTN = dbfRecordset.fields(0).value
            Dim ChkType = dbfRecordset.fields(1).value

            Dim dbfConnector1 = CreateObject("ADODB.Connection")
            dbfConnector1 = CreateObject("ADODB.Connection")
            dbfConnector1.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
            dbfConnector1.CursorLocation = ADODB.CursorLocationEnum.adUseClient

            Dim Orig_Address1 = ""
            Dim Orig_Address2 = ""
            Dim Orig_Address3 = ""
            Dim Orig_Address4 = ""
            Dim Orig_Address5 = ""
            Dim Orig_Address6 = ""

            Dim LoopCount1 = 0
            Do Until LoopCount1 = Val(TotalCheckdat) + 1
                Dim CheckDatLocation As String = ""
                Dim Drive_Checkdat As String = ""

                If LoopCount1 = 0 Then
                    CheckDatLocation = "C:\Checktho\Checkdat.mdb"
                    Drive_Checkdat = "C:\CheckTho\"
                Else
                    CheckDatLocation = CheckDatMDB(LoopCount1 - 1)
                    Drive_Checkdat = CheckDatMDBDrive(LoopCount1 - 1)
                End If

                'Checkdat Location
                Dim Conn As New ADODB.Connection
                Dim Rs As New ADODB.Recordset

                Conn = New ADODB.Connection
                Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                   "Data Source=" & CheckDatLocation & "; Jet OLEDB:Database Password=CorpCaptive;"
                Conn.Open()
                'End Checkdat Location

                'Search for BRSTN
                Sql = "SELECT [Branch Text 1], [Branch Text 2], [Branch Text 3], [Branch Text 4], [Branch Text 5], [Branch Text 6] FROM Branch WHERE [Routing Number] = '" & BRSTN & "'"
                Rs = New ADODB.Recordset
                Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                Rs.Open(Sql, Conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                If Rs.RecordCount <= 0 Then
                    Dim Description = "BRSTN " & BRSTN & " does not exists on Drive " & Drive_Checkdat

                    Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                    Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description, "'", "''") & "')"
                    dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                    TotalErrors = TotalErrors + 1
                Else
                    If LoopCount1 = 0 Then
                        If (IsDBNull(Rs.Fields(0).Value)) = False Then
                            Orig_Address1 = Rs.Fields(0).Value
                        Else
                            Orig_Address1 = ""
                        End If

                        If (IsDBNull(Rs.Fields(1).Value)) = False Then
                            Orig_Address2 = Rs.Fields(1).Value
                        Else
                            Orig_Address2 = ""
                        End If

                        If (IsDBNull(Rs.Fields(2).Value)) = False Then
                            Orig_Address3 = Rs.Fields(2).Value
                        Else
                            Orig_Address3 = ""
                        End If

                        If (IsDBNull(Rs.Fields(3).Value)) = False Then
                            Orig_Address4 = Rs.Fields(3).Value
                        Else
                            Orig_Address4 = ""
                        End If

                        If (IsDBNull(Rs.Fields(4).Value)) = False Then
                            Orig_Address5 = Rs.Fields(4).Value
                        Else
                            Orig_Address5 = ""
                        End If

                        If (IsDBNull(Rs.Fields(5).Value)) = False Then
                            Orig_Address6 = Rs.Fields(5).Value
                        Else
                            Orig_Address6 = ""
                        End If

                        Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                        Sql = "UPDATE SBTC SET Address1 = '" & Replace(Orig_Address1, "'", "''") & "', Address2 = '" & Replace(Orig_Address2, "'", "''") & "', Address3 = '" & Replace(Orig_Address3, "'", "''") & "', Address4 = '" & Replace(Orig_Address4, "'", "''") & "', Address5 = '" & Replace(Orig_Address5, "'", "''") & "', Address6 = '" & Replace(Orig_Address6, "'", "''") & "' WHERE BRSTN = '" & BRSTN & "'"
                        dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                        dbfRecordset1 = CreateObject("ADODB.Recordset")
                        Sql = "SELECT Branch_Tex FROM REF WHERE ChkType = '" & ChkType & "' and RTNO = '" & BRSTN & "'"
                        dbfRecordset1.Open(Sql, dbfConnector1, 1, 1)

                        If dbfRecordset1.Recordcount = 1 Then
                            Dim BranchName = Trim(dbfRecordset1.fields(0).value)

                            If BranchName <> Trim(Orig_Address1) Then
                                Dim Description = "BRSTN " & BRSTN & " with ChkType " & ChkType & " has a different Address on Ref.dbf" & vbNewLine & "Ref: " & BranchName & vbNewLine & "Checkdat: " & Orig_Address1

                                dbfRecordset1 = CreateObject("ADODB.Recordset")
                                Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description, "'", "''") & "')"
                                dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                                TotalErrors = TotalErrors + 1
                            End If
                        Else
                            Dim Description = "BRSTN " & BRSTN & " with ChkType " & ChkType & " does not exists on Ref.dbf"

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            TotalErrors = TotalErrors + 1
                        End If

                    Else
                        Dim Address1 = ""
                        If (IsDBNull(Rs.Fields(0).Value)) = False Then
                            Address1 = Rs.Fields(0).Value
                        Else
                            Address1 = ""
                        End If

                        Dim Address2 = ""
                        If (IsDBNull(Rs.Fields(1).Value)) = False Then
                            Address2 = Rs.Fields(1).Value
                        Else
                            Address2 = ""
                        End If

                        Dim Address3 = ""
                        If (IsDBNull(Rs.Fields(2).Value)) = False Then
                            Address3 = Rs.Fields(2).Value
                        Else
                            Address3 = ""
                        End If

                        Dim Address4 = ""
                        If (IsDBNull(Rs.Fields(3).Value)) = False Then
                            Address4 = Rs.Fields(3).Value
                        Else
                            Address4 = ""
                        End If

                        Dim Address5 = ""
                        If (IsDBNull(Rs.Fields(4).Value)) = False Then
                            Address5 = Rs.Fields(4).Value
                        Else
                            Address5 = ""
                        End If

                        Dim Address6 = ""
                        If (IsDBNull(Rs.Fields(5).Value)) = False Then
                            Address6 = Rs.Fields(5).Value
                        Else
                            Address6 = ""
                        End If






                        If Address1 <> Orig_Address1 Then
                            Dim Description1 = "C:\ --> " & Orig_Address1
                            Dim Description2 = Drive_Checkdat & ":\ -->" & Address1

                            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description1, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description2, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            TotalErrors = TotalErrors + 1
                        End If

                        If Address2 <> Orig_Address2 Then
                            Dim Description1 = "C:\ --> " & Orig_Address2
                            Dim Description2 = Drive_Checkdat & ":\ -->" & Address2


                            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description1, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description2, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            TotalErrors = TotalErrors + 1
                        End If

                        If Address3 <> Orig_Address3 Then
                            Dim Description1 = "C:\ --> " & Orig_Address3
                            Dim Description2 = Drive_Checkdat & ":\ -->" & Address3

                            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description1, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description2, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            TotalErrors = TotalErrors + 1
                        End If

                        If Address4 <> Orig_Address4 Then
                            Dim Description1 = "C:\ --> " & Orig_Address4
                            Dim Description2 = Drive_Checkdat & ":\ -->" & Address4


                            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description1, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description2, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            TotalErrors = TotalErrors + 1
                        End If

                        If Address5 <> Orig_Address5 Then
                            Dim Description1 = "C:\ --> " & Orig_Address5
                            Dim Description2 = Drive_Checkdat & ":\ -->" & Address5


                            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description1, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description2, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            TotalErrors = TotalErrors + 1
                        End If

                        If Address6 <> Orig_Address6 Then
                            Dim Description1 = "C:\ --> " & Orig_Address6
                            Dim Description2 = Drive_Checkdat & ":\ -->" & Address6


                            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description1, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('" & Replace(Description2, "'", "''") & "')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            dbfRecordset1 = CreateObject("ADODB.Recordset")
                            Sql = "INSERT INTO Errors (Errors) VALUES ('')"
                            dbfRecordset1.Open(Sql, dbfConnector, 1, 1)

                            TotalErrors = TotalErrors + 1
                        End If


                    End If
                End If
                'End Search for BRSTN

                Rs.Close()
                Conn.Close()

                LoopCount1 = LoopCount1 + 1
            Loop

            dbfRecordset.movenext()
            LoopCount = LoopCount + 1
        Loop
        'End Checkdat.mdb
    End Sub

    Sub SortRT()
        Dim GrandTotal = 0
        Dim PageNo As String = 1
        Dim LineNumber = 0

        Dim SortRT_Output As System.IO.StreamWriter

        SortRT_Output = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\SortRT.txt", False)

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim SQL = "SELECT ChkType, BRSTN, Address1, SUM(OrderQty) FROM SBTC GROUP BY ChkType, BRSTN, Address1  ORDER BY ChkType, BRSTN"
        dbfRecordset.Open(SQL, dbfConnector, 1, 1)

        Dim LoopCount = 0
        Do Until LoopCount = dbfRecordset.recordcount
            If LoopCount = 0 Or LineNumber >= 50 Then
                SortRT_Output.WriteLine("")
                If PageNo <> 1 Then SortRT_Output.WriteLine("")

                SortRT_Output.WriteLine("    Page No. " & PageNo)
                SortRT_Output.WriteLine("    " & Now.ToShortDateString)
                SortRT_Output.WriteLine("                             Summary of RT nos / # of Books")
                SortRT_Output.WriteLine("                               SBTC - Regular Checks")
                SortRT_Output.WriteLine("")
                SortRT_Output.WriteLine("    ACCTNO       QTY BRANCH                 ACCOUNT NAME")
                SortRT_Output.WriteLine("")

                PageNo = PageNo + 1
                LineNumber = 0
            End If


            Dim ChkType = dbfRecordset.fields(0).value
            Dim BRSTN = dbfRecordset.fields(1).value
            Dim Address1 = dbfRecordset.fields(2).value
            Dim SubTotal = dbfRecordset.fields(3).value

            SortRT_Output.WriteLine("")
            SortRT_Output.WriteLine("   ** CHECK TYPE/BRSTN/BATCH # ---->  " & ChkType & "/" & BRSTN & "/" & FinalBatch)
            SortRT_Output.WriteLine("   ** Branch: " & Address1)
            LineNumber = LineNumber + 3

            'For Details
            Dim dbfRecordset1 = CreateObject("ADODB.Recordset")
            SQL = "SELECT AccountNo, Name1, Name2, OrderQty FROM SBTC WHERE ChkType = '" & ChkType & "' AND BRSTN = '" & BRSTN & "' AND Address1 = '" & Replace(Address1, "'", "''") & "'"
            dbfRecordset1.Open(SQL, dbfConnector, 1, 1)

            Dim LoopCount1 = 0
            Do Until LoopCount1 = dbfRecordset1.recordcount
                Dim AccountNo = dbfRecordset1.fields(0).value

                Dim Name1 = "'"
                If IsDBNull(dbfRecordset1.fields(1).value) = False Then
                    Name1 = dbfRecordset1.fields(1).value
                Else
                    Name1 = ""
                End If

                Dim Name2 = "'"
                If IsDBNull(dbfRecordset1.fields(2).value) = False Then
                    Name2 = dbfRecordset1.fields(2).value
                Else
                    Name2 = ""
                End If

                Dim OrderQty = dbfRecordset1.fields(3).value

                Do Until Len(OrderQty) = 4
                    OrderQty = " " & OrderQty
                Loop

                SortRT_Output.WriteLine("    " & AccountNo & OrderQty & " " & Name1)
                LineNumber = LineNumber + 1

                If Name2 <> "" Then
                    SortRT_Output.WriteLine("                    " & Name2)
                    LineNumber = LineNumber + 1
                End If

                If LineNumber >= 50 Then
                    SortRT_Output.WriteLine("")
                    If PageNo <> 1 Then SortRT_Output.WriteLine("")

                    SortRT_Output.WriteLine("    Page No. " & PageNo)
                    SortRT_Output.WriteLine("    " & Now.ToShortDateString)
                    SortRT_Output.WriteLine("                             Summary of RT nos / # of Books")
                    SortRT_Output.WriteLine("                               SBTC - Regular Checks")
                    SortRT_Output.WriteLine("")
                    SortRT_Output.WriteLine("    ACCTNO       QTY BRANCH                 ACCOUNT NAME")
                    SortRT_Output.WriteLine("")

                    PageNo = PageNo + 1
                    LineNumber = 0
                End If

                dbfRecordset1.movenext()
                LoopCount1 = LoopCount1 + 1
            Loop
            'End For Details

            SortRT_Output.WriteLine("")
            SortRT_Output.WriteLine("    Sub Total: " & SubTotal)
            LineNumber = LineNumber + 2

            GrandTotal = GrandTotal + SubTotal

            dbfRecordset.Movenext()
            LoopCount = LoopCount + 1
        Loop

        SortRT_Output.WriteLine("")
        SortRT_Output.WriteLine("    Grand Total: " & GrandTotal)

        SortRT_Output.Close()

    End Sub

    Sub CreateDBaseFile()
        If Dir(Application.StartupPath & "\SBTC.dbf") <> "" Then Kill(Application.StartupPath & "\SBTC.dbf")
        If Dir(Application.StartupPath & "\Errors.dbf") <> "" Then Kill(Application.StartupPath & "\Errors.dbf")

        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim dbfRecordset = CreateObject("ADODB.Recordset")
        Dim Sql = "CREATE TABLE SBTC (ChkType Varchar(254), BRSTN Varchar(254), AccountNo Varchar(254), AccountNo2 Varchar(254), Name1 Varchar(254), Name2 Varchar(254), OrderQty Varchar(254), Address1 Varchar(254), Address2 Varchar(254), Address3 Varchar(254), Address4 Varchar(254), Address5 Varchar(254), Address6 Varchar(254))"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)

        dbfRecordset = CreateObject("ADODB.Recordset")
        Sql = "CREATE TABLE Errors (Errors Varchar(200))"
        dbfRecordset.Open(Sql, dbfConnector, 1, 1)
    End Sub

    Sub ReadTextFile()

        FinalBatch = InputBox("Enter Batch Number", " ", "")
        If FinalBatch = "" Then End
        FinalBatch = UCase(FinalBatch)

        'For Dbase Configuration
        Dim dbfConnector = CreateObject("ADODB.Connection")
        dbfConnector = CreateObject("ADODB.Connection")

        dbfConnector.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Application.StartupPath & "\;Extended properties=dBase III")
        dbfConnector.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        'End For Dbase Configuration

        FileClose(1)
        FileOpen(1, dlgBrowse.FileName, OpenMode.Input)

        CreateDBaseFile()

        Dim TotalDataLine = 0

        Do Until EOF(1)
            Dim LineInputData As String = LineInput(1)

            Dim ChkType As String = Mid(LineInputData, 1, 1)
            Dim BRSTN As String = Mid(LineInputData, 2, 9)
            Dim AccountNo2 As String = Mid(LineInputData, 11, 13)
            Dim AccountName As String = Trim(Mid(LineInputData, 24, 57))
            Dim OrderQty As String = Val(Mid(LineInputData, 83, 2))
            Dim AccountNo As String = Mid(LineInputData, 86, 12)


            'Save It
            Dim dbfRecordset = CreateObject("ADODB.Recordset")
            Dim Sql = "INSERT INTO SBTC (ChkType, BRSTN, AccountNo, AccountNo2, Name1, OrderQty) VALUES ('" _
                      & ChkType & "','" & BRSTN & "','" & AccountNo & "','" & AccountNo2 & "','" & Replace(AccountName, "'", "''") & "','" & OrderQty & "')"
            dbfRecordset.Open(Sql, dbfConnector, 1, 1)
            'End Save It

        Loop

        FileClose(1)
    End Sub
End Class
