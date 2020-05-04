using ADOX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;


namespace sbtc
{
    public static class GenerateService
    {
        static string regPath = Application.StartupPath + "\\Output\\Regular";

        static string regPrePath = Application.StartupPath + "\\Output\\Regular\\PreEncoded";

        static string chargeSlipPath = Application.StartupPath + "\\Output\\Charge_Slip";

        static string checkOnePath = Application.StartupPath + "\\Output\\CheckOne";

        static string checkPowerPath = Application.StartupPath + "\\Output\\CheckPower";

        static string customPath = Application.StartupPath + "\\Output\\Customized";

        static string gcPath = Application.StartupPath + "\\Output\\GiftCheck";

        static string mcPath = Application.StartupPath + "\\Output\\MC";

        static string mcContPath = Application.StartupPath + "\\Output\\MC\\Continues";

        private static void SortHeader(StreamWriter sw, string _folderName, int _page)
        {
            sw.WriteLine("");

            sw.WriteLine("\t\tPage No. " + _page.ToString());

            sw.WriteLine(""); sw.WriteLine("");

            sw.WriteLine("\t\t\t\tSummary of RT nos / # of Books");

            if (_folderName == "Regular")
                sw.WriteLine("\t\t\t\t\tSBTC - Regular Checks");
            else if (_folderName == "Regular\\PreEncoded")
                sw.WriteLine("\t\t\t\t\tSBTC - PreEncoded Checks");
            else if (_folderName == "MC")
                sw.WriteLine("\t\t\t\t\tSBTC - Manager's Check");
            else if (_folderName == "CheckOne")
                sw.WriteLine("\t\t\t\t\tSBTC - Check One");
            else if (_folderName == "GiftCheck")
                sw.WriteLine("\t\t\t\t\tSBTC - Gift Check");
            else if (_folderName == "MC\\Continuous")
                sw.WriteLine("\t\t\t\t\tSBTC - Manager's Check Continuous");
            else if (_folderName == "Customized")
                sw.WriteLine("\t\t\t\t\tSBTC - Customized Checks");
            else if (_folderName == "Charge_Slip")
                sw.WriteLine("\t\t\t\t\tSBTC - Charge Slip");

            sw.WriteLine("\t\tACCTNO      QTY ACCOUNT NAME");

            sw.WriteLine("");
        }

        public static string GenerateSortRT(string _folderName, List<OrderModel> _orders)
        {
            int page = 1, lineCounter = 0;

            string directory = Application.StartupPath + "\\" + _folderName;

            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            if (File.Exists(directory + "\\SortRT.txt"))
                File.Delete(directory + "\\SortRT.txt");

            StreamWriter sw;

            sw = File.CreateText(directory + "\\SortRT.txt");
            sw.Close();

            using (sw = new StreamWriter(File.Open(directory + "\\SortRT.txt", FileMode.Append)))
            {
                SortHeader(sw, _folderName, page);

                lineCounter += 10;

                List<OrderModel> checkListPersonal = new List<OrderModel>();

                List<OrderModel> checkListCommercial = new List<OrderModel>();

                string checkTypePersonal = "", checkTypeCommercial = "", formTypePersonal = "", formTypeCommercial = "";

                var brstnList = _orders.Select(r => r.BRSTN).Distinct().OrderBy(r => r).ToList();

                if (_folderName == "MC" || _folderName == "GiftCheck" || _folderName == "MC\\Continues")
                {
                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        var orderList = _orders.Where(r => r.BRSTN == brstnList[x]).ToList();

                        sw.WriteLine("\t\t** CHECK TYPE/BRSTN/BATCH # ---->\t\t A/" + orderList[0].BRSTN);

                        sw.WriteLine("\t\t** Branch: " + orderList[0].Address1);

                        lineCounter += 2;

                        orderList.ForEach(order =>
                        {
                            if (lineCounter == 0 || lineCounter >= 50)
                            {
                                lineCounter = 0;//RESET LINE COUNTER

                                sw.WriteLine("");

                                SortHeader(sw, _folderName, page);

                                lineCounter += 10;

                                page++;
                            }

                            string qty;

                            if (order.OrderQuantity.ToString().Length == 1)
                                qty = " " + order.OrderQuantity.ToString();
                            else
                                qty = order.OrderQuantity.ToString();

                            sw.WriteLine("\t\t" + order.AccountNo + " " + qty + " " + order.Name);
                            lineCounter++;

                        });//END OF FOREACH OrderList

                        sw.WriteLine(""); sw.WriteLine("\t\tSub Total: " + orderList.Count.ToString());
                    }

                    sw.WriteLine(""); sw.WriteLine("");

                    sw.WriteLine("\t\tGrand Total: " + _orders.Count.ToString());
                }
                else
                {
                    if (_folderName == "Regular")
                    {
                        checkTypePersonal = "A";
                        formTypePersonal = "05";

                        checkTypeCommercial = "B";
                        formTypeCommercial = "16";

                    }
                    else if (_folderName == "Regular\\PreEncoded")
                    {
                        checkTypePersonal = "AA";

                        checkTypePersonal = "BB";
                    }
                    else if (_folderName == "CheckOne")
                    {
                        checkTypePersonal = checkTypeCommercial = "F";

                        formTypePersonal = "25";

                        formTypeCommercial = "26";
                    }
                    else if (_folderName == "CheckPower")
                    {
                        checkTypePersonal = checkTypeCommercial = "E";

                        formTypePersonal = "23";

                        formTypeCommercial = "22";
                    }

                    checkListPersonal = _orders.Where(r => r.CheckType == checkTypePersonal && r.FormType == formTypePersonal).ToList();

                    checkListCommercial = _orders.Where(r => r.CheckType == checkTypeCommercial && r.FormType == formTypeCommercial).ToList();

                    if (checkListPersonal.Count > 0)
                    {
                        var brstn = checkListPersonal.Select(r => r.BRSTN).Distinct().ToList();

                        for (int x = 0; x < brstn.Count; x++)
                        {
                            var orderList = checkListPersonal.Where(r => r.BRSTN == brstn[x]).ToList();

                            sw.WriteLine("\t\t** CHECK TYPE/BRSTN/BATCH # ---->\t\t A/" + orderList[0].BRSTN);

                            sw.WriteLine("\t\t** Branch: " + orderList[0].Address1);

                            lineCounter += 2;

                            orderList.ForEach(order =>
                            {
                                if(lineCounter == 0 || lineCounter >= 50)
                                {
                                    lineCounter = 0;//RESET LINECOUNTER

                                    sw.WriteLine("");

                                    SortHeader(sw, _folderName, page);

                                    lineCounter += 10;

                                    page++;
                                }

                                string qty;

                                if (order.OrderQuantity.ToString().Length == 1)
                                    qty = " " + order.OrderQuantity.ToString();
                                else
                                    qty = order.OrderQuantity.ToString();

                                sw.WriteLine("\t\t" + order.AccountNo + " " + qty + " " + order.Name);

                                lineCounter++;
                            });//END FOREACH

                            sw.WriteLine(""); sw.WriteLine("\t\tSub Total: " + orderList.Count.ToString());
                        }//END FOR

                        sw.WriteLine(""); sw.WriteLine("");

                        sw.WriteLine("\t\tGrand Total: " + _orders.Count.ToString());
                    }//END IF PERSONAL IS NOT NULL

                    if (checkListCommercial.Count > 0)
                    {
                        var brstn = checkListCommercial.Select(r => r.BRSTN).Distinct().ToList();

                        for (int x = 0; x < brstn.Count; x++)
                        {
                            var orderList = checkListCommercial.Where(r => r.BRSTN == brstn[x]).ToList();

                            sw.WriteLine("\t\t** CHECK TYPE/BRSTN/BATCH # ---->\t\t A/" + orderList[0].BRSTN);

                            sw.WriteLine("\t\t** Branch: " + orderList[0].Address1);

                            lineCounter += 2;

                            orderList.ForEach(order =>
                            {
                                if (lineCounter == 0 || lineCounter >= 50)
                                {
                                    lineCounter = 0;//RESET LINECOUNTER

                                    sw.WriteLine("");

                                    SortHeader(sw, _folderName, page);

                                    lineCounter += 10;

                                    page++;
                                }

                                string qty;

                                if (order.OrderQuantity.ToString().Length == 1)
                                    qty = " " + order.OrderQuantity.ToString();
                                else
                                    qty = order.OrderQuantity.ToString();

                                sw.WriteLine("\t\t" + order.AccountNo + " " + qty + " " + order.Name);

                                lineCounter++;
                            });//END FOREACH

                            sw.WriteLine(""); sw.WriteLine("\t\tSub Total: " + orderList.Count.ToString());
                        }//END FOR

                        sw.WriteLine(""); sw.WriteLine("");

                        sw.WriteLine("\t\tGrand Total: " + _orders.Count.ToString());
                    }

                }//END OF ELSE
            }//END OF USING STREAMWRITER
            return "";
        }//END OF FUNCTION

        public static void GeneratePackingList(OrderSorted _orders, string _batch, DateTime _deliveryDate, 
            List<BranchesModel> _branches)
        {
            #region Regular Personal
            if (_orders.RegularPersonal.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.RegularPersonal.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = regPath + "\\PackingA.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Personal Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.RegularPersonal[0].Batch);

                        lineCount = 11;

                        var checks = _orders.RegularPersonal.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string tempStart = check.StartingSerial.ToString();

                            while (tempStart.Length < 7)
                                tempStart = "0" + tempStart;

                            while (tempStart.Length < 10)
                                tempStart += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + tempStart + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                            
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {                           
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FOR FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Personal Checks Summary");

                        sw.WriteLine("                                  (F R O N T  C O V E R)");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.RegularPersonal[0].Batch);

                        lineCount = 11;

                        var checks = _orders.RegularPersonal.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string tempStart = check.StartingSerial.ToString();

                            while (tempStart.Length < 7)
                                tempStart = "0" + tempStart;

                            while (tempStart.Length < 11)
                                tempStart += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + tempStart + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region Regular Commercial
            if (_orders.RegularCommercial.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.RegularCommercial.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = regPath + "\\PackingB.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.RegularCommercial[0].Batch);

                        lineCount = 11;

                        var checks = _orders.RegularCommercial.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " "  + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.RegularCommercial[0].Batch);

                        lineCount = 11;

                        var checks = _orders.RegularCommercial.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                }//END USING
            }//END IF
            #endregion

            #region Personal Pre-Encoded
            if (_orders.PersonalPreEncoded.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.PersonalPreEncoded.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = regPrePath + "\\PackingA.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Personal Pre-Encoded Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.PersonalPreEncoded[0].Batch);

                        lineCount = 11;

                        var checks = _orders.PersonalPreEncoded.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 7)
                                start = "0" + start;

                            while (start.Length < 11)
                                start += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Personal Pre-Encoded Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.PersonalPreEncoded[0].Batch);

                        lineCount = 11;

                        var checks = _orders.PersonalPreEncoded.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 7)
                                start = "0" + start;

                            while (start.Length < 11)
                                start += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region  Commercial Pre-Encoded
            if (_orders.CommercialPreEncoded.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.CommercialPreEncoded.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = regPrePath + "\\PackingB.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CommercialPreEncoded[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CommercialPreEncoded.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CommercialPreEncoded[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CommercialPreEncoded.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region CheckOne Personal
            if (_orders.CheckOnePersonal.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.CheckOnePersonal.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = checkOnePath + "\\PackingA.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckOne Personal Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #:" + _orders.CheckOnePersonal[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckOnePersonal.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 7)
                                start = "0" + start;

                            while (start.Length < 10)
                                start += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckOne Personal Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #:" + _orders.CheckOnePersonal[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckOnePersonal.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 7)
                                start = "0" + start;

                            while (start.Length < 10)
                                start += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region CheckOne Commercial
            if (_orders.CheckOneCommerical.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.CheckOneCommerical.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = checkOnePath + "\\PackingB.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckOne Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CheckOneCommerical[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckOneCommerical.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckOne Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CheckOneCommerical[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckOneCommerical.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region CheckPower Personal
            if (_orders.CheckPowerPersonal.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.CheckPowerPersonal.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = checkPowerPath + "\\PackingA.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckPower Personal Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CheckPowerPersonal[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckPowerPersonal.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string tempStart = check.StartingSerial.ToString();

                            while (tempStart.Length < 7)
                                tempStart = "0" + tempStart;

                            while (tempStart.Length < 11)
                                tempStart += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + tempStart + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckPower Personal Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CheckPowerPersonal[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckPowerPersonal.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string tempStart = check.StartingSerial.ToString();

                            while (tempStart.Length < 7)
                                tempStart = "0" + tempStart;

                            while (tempStart.Length < 11)
                                tempStart += " ";

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 7)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + tempStart + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region CheckPower Commercial
            if (_orders.CheckPowerCommercial.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.CheckPowerCommercial.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = checkPowerPath + "\\PackingB.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckPower Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CheckPowerCommercial[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckPowerCommercial.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - CheckPower Commercial Checks Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.CheckPowerCommercial[0].Batch);

                        lineCount = 11;

                        var checks = _orders.CheckPowerCommercial.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region Manager's Check
            if (_orders.ManagersCheck.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.ManagersCheck.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = mcPath + "\\PackingMC.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Manager's Check Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.ManagersCheck[0].Batch);

                        lineCount = 11;

                        var checks = _orders.ManagersCheck.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Manager's Check Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.ManagersCheck[0].Batch);

                        lineCount = 11;

                        var checks = _orders.ManagersCheck.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR
                }//END USING
            }//END IF
            #endregion

            #region Manager's Contl Check
            if (_orders.ManagersCheckCont.Count > 0)
            {
                int pageNo = 1, lineCount = 0;

                var brstnList = _orders.ManagersCheckCont.Select(r => r.BRSTN).Distinct().ToList();

                StreamWriter sw;

                string fileName = mcContPath + "\\PackingMC.txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    sw.WriteLine("");

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Manager's Contl Check Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.ManagersCheckCont[0].Batch);

                        lineCount = 11;

                        var checks = _orders.ManagersCheckCont.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "  " + tempName + " 1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                    //FRONT COVER
                    sw.WriteLine("");

                    lineCount = 0;

                    for (int x = 0; x < brstnList.Count; x++)
                    {
                        sw.WriteLine("  Page No. " + pageNo.ToString());

                        sw.WriteLine("  " + DateTime.Now.ToString("MMMM dd yyyy"));

                        sw.WriteLine("                                CAPTIVE PRINTING CORPORATION");

                        sw.WriteLine("                               SBTC - Manager's Contl Check Summary");

                        sw.WriteLine("");

                        sw.WriteLine("  ACCT_NO         ACCOUNT NAME                     QTY CT START #    END #");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        var branch = _branches.FirstOrDefault(r => r.BRSTN == brstnList[x]);

                        sw.WriteLine(" ** ORDERS OF BRSTN " + brstnList[x] + " " + branch.Address1);

                        sw.WriteLine("");

                        sw.WriteLine(" * Batch #: " + _orders.ManagersCheckCont[0].Batch);

                        lineCount = 11;

                        var checks = _orders.ManagersCheckCont.Where(r => r.BRSTN == brstnList[x]).ToList();

                        checks.ForEach(check =>
                        {
                            string temp = check.AccountNo.Substring(0, 3) + "-" + check.AccountNo.Substring(3, 6) + "-" + check.AccountNo.Substring(9, 3);

                            string tempName = check.Name;

                            while (tempName.Length < 35)
                                tempName += " ";

                            string start = check.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = check.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("  " + temp + "                                     1 A  " + start + " " + end);

                            lineCount++;

                            if (lineCount >= 60)
                            {
                                sw.WriteLine("");

                                lineCount = 0;
                            }
                        });

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine(" * * * Sub Total * * *                               " + checks.Count.ToString());

                        if (x + 1 < brstnList.Count)
                        {
                            sw.WriteLine("");

                            sw.WriteLine("");

                            pageNo++;
                        }
                    }//END FOR

                }//END USING
            }//END IF
            #endregion
        }//END FUNCTION

        public static void GenerateDoBlock(OrderSorted _orders, string _batch, string _ext, DateTime _deliveryDate,
            string _preparedBy)
        {
            if (_batch == "0000")
                _preparedBy = "TEST ONLY";

            if(_orders.RegularPersonal.Count> 0)
            {
                StreamWriter sw;

                string fileName = regPath + "\\BlockP.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.RegularPersonal.Select(r => r.BRSTN).Distinct().ToList();

                    for(int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.RegularPersonal.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {                               
                                if(page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.RegularPersonal.Count.ToString() + "                 " + _batch.Substring(0,4) + "_P12" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.RegularCommercial.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");                                

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - PERSONAL");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }
                            
                            if(blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;                           
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.RegularPersonal[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.RegularPersonal.Count.ToString() + "                 " + _batch.Substring(0, 4) + "_P12" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.RegularCommercial.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.RegularCommercial.Count > 0)
            {
                StreamWriter sw;

                string fileName = regPath + "\\BlockC.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.RegularCommercial.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.RegularCommercial.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.RegularPersonal.Count.ToString() + "                 " + _batch.Substring(0, 4) + "_C12" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.RegularCommercial.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - COMMERCIAL");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.RegularCommercial[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.RegularPersonal.Count.ToString() + "                 " + _batch.Substring(0, 4) + "_C12" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.RegularCommercial.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.PersonalPreEncoded.Count > 0)
            {
                StreamWriter sw;

                string fileName = regPrePath + "\\BlockP.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.PersonalPreEncoded.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.PersonalPreEncoded.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.PersonalPreEncoded.Count.ToString() + "                 " + _batch.Substring(0, 4) + "P" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.CommercialPreEncoded.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - PERSONAL [Pre-Encoded]");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.PersonalPreEncoded[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.PersonalPreEncoded.Count.ToString() + "                 " + _batch.Substring(0, 4) + "P" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.CommercialPreEncoded.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.CommercialPreEncoded.Count > 0)
            {
                StreamWriter sw;

                string fileName = regPrePath + "\\BlockC.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.CommercialPreEncoded.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.CommercialPreEncoded.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.PersonalPreEncoded.Count.ToString() + "                 " + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.CommercialPreEncoded.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - COMMERCIAL [Pre-Encoded]");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.CommercialPreEncoded[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.PersonalPreEncoded.Count.ToString() + "                 " + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.CommercialPreEncoded.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.CheckOnePersonal.Count > 0)
            {
                StreamWriter sw;

                string fileName = checkOnePath + "\\BlockP.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.CheckOnePersonal.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.CheckOnePersonal.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.CheckOnePersonal.Count.ToString() + "                 13D" + _batch.Substring(0, 4) + "P" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.CheckOneCommerical.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - CHECKONE PERSONAL");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.CheckOnePersonal[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.CheckOnePersonal.Count.ToString() + "                 13D" + _batch.Substring(0, 4) + "P" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.CheckOneCommerical.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.CheckOneCommerical.Count > 0)
            {
                StreamWriter sw;

                string fileName = checkOnePath + "\\BlockC.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.CheckOneCommerical.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.CheckOneCommerical.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.CheckOnePersonal.Count.ToString() + "                 13D" + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.CheckOneCommerical.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - CHECKONE COMMERCIAL");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.CheckOneCommerical[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.CheckOnePersonal.Count.ToString() + "                 13D" + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.CheckOneCommerical.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.CheckPowerPersonal.Count > 0)
            {
                StreamWriter sw;

                string fileName = checkPowerPath + "\\BlockP.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.CheckPowerPersonal.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.CheckPowerPersonal.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.CheckPowerPersonal.Count.ToString() + "                 CKP" + _batch.Substring(0, 4) + "P" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.CheckPowerCommercial.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - CHECKPOWER PERSONAL");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.CheckPowerPersonal[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.CheckPowerPersonal.Count.ToString() + "                 CKP" + _batch.Substring(0, 4) + "P" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.CheckPowerCommercial.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.CheckPowerCommercial.Count > 0)
            {
                StreamWriter sw;

                string fileName = checkPowerPath + "\\BlockC.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.CheckPowerCommercial.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.CheckPowerCommercial.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPA = " + _orders.CheckPowerPersonal.Count.ToString() + "                 CKP" + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                                    sw.WriteLine("\t\tCA = " + _orders.CheckPowerCommercial.Count.ToString());

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - CHECKPOWER COMMERCIAL");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.CheckPowerCommercial[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.CheckPowerPersonal.Count.ToString() + "                 CKP" + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                        sw.WriteLine("\t\tCA = " + _orders.CheckPowerCommercial.Count.ToString());

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF

            if(_orders.ManagersCheck.Count > 0)
            {
                StreamWriter sw;

                string fileName = mcPath + "\\BlockMC.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.ManagersCheck.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.ManagersCheck.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tMC = " + _orders.ManagersCheck.Count.ToString() + "                 MC" + _batch.Substring(0, 4) + "P" + _ext + ".txt");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - Manager's Check");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.ManagersCheck[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tMC = " + _orders.ManagersCheck.Count.ToString() + "                 CKP" + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF
            
            if(_orders.ManagersCheckCont.Count > 0)
            {
                StreamWriter sw;

                string fileName = mcContPath + "\\BlockMC.txt";

                bool footer = true;

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    int page = 1, block = 1, blockCounter = 0;

                    var brstn = _orders.ManagersCheckCont.Select(r => r.BRSTN).Distinct().ToList();

                    for (int x = 0; x < brstn.Count; x++)
                    {
                        var checks = _orders.ManagersCheckCont.Where(r => r.BRSTN == brstn[x]).ToList();

                        checks.ForEach(c =>
                        {
                            if ((block % 8 == 0 && blockCounter == 4) || (block == 1 && blockCounter == 0))
                            {
                                if (page == 2 && footer)
                                {
                                    sw.WriteLine("");

                                    sw.WriteLine("\t\t" + c.Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                                    sw.WriteLine("");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tMC = " + _orders.ManagersCheckCont.Count.ToString() + "                 MCC" + _batch.Substring(0, 4) + "B" + _ext + ".txt");

                                    sw.WriteLine("");

                                    sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                                    sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                                    footer = false;
                                }

                                if (block % 8 == 0 && blockCounter == 4)
                                    sw.WriteLine("");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tPage No." + page.ToString());

                                sw.WriteLine("\t\t" + DateTime.Today.ToShortDateString());

                                sw.WriteLine("\t\t\t SBTC - SUMMARY OF BLOCK - Manager's Check Continuous");

                                sw.WriteLine("");

                                sw.WriteLine("\t\tBLOCK RT_NO     M ACCT_NO         START_NO.  END_NO.");

                                sw.WriteLine("");

                                sw.WriteLine("");

                                page++;
                            }

                            if (blockCounter == 4)
                            {
                                block++;

                                blockCounter = 0;

                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }
                            else if (block == 1 && blockCounter == 0)
                            {
                                sw.WriteLine("");

                                sw.WriteLine("\t\t** BLOCK " + block.ToString());
                            }

                            string start = c.StartingSerial.ToString();

                            while (start.Length < 10)
                                start = "0" + start;

                            string end = c.EndingSerial.ToString();

                            while (end.Length < 10)
                                end = "0" + end;

                            sw.WriteLine("\t\t\t" + block.ToString() + " " + c.BRSTN + "   " + c.AccountNo + "    " + start + "    " + end);

                            blockCounter++;
                        });//END FOREACH
                    }//END FOR

                    if (footer)
                    {
                        sw.WriteLine("");

                        sw.WriteLine("\t\t" + _orders.ManagersCheckCont[0].Batch + "                                 DLVR: " + String.Format("{0:MM-dd(ddd)}", _deliveryDate));

                        sw.WriteLine("");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPA = " + _orders.ManagersCheckCont.Count.ToString() + "                 CKP" + _batch.Substring(0, 4) + "C" + _ext + ".txt");

                        sw.WriteLine("");

                        sw.WriteLine("\t\tPrepared By   : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tUpdated By    : " + _preparedBy.ToUpper());

                        sw.WriteLine("\t\tTime Finished : " + DateTime.Now.ToShortTimeString());

                        footer = false;
                    }
                }//END USING
            }//END IF
        }//END FUNCTION

        public static void GeneratePackingDBF(OrderSorted _orders, string _batch, string _ext)
        {
            string path, connectionString, query;

            int blockNo, blockCounter;

            #region Regular Checks
            //REGULAR CHECKS
            path = regPath + "\\Packing.dbf";

            connectionString = "Provider=VFPOLEDB.1; Data Source=" + path + "; Mode=ReadWrite;";

            query = "";

            if(_orders.RegularPersonal.Count > 0 || _orders.RegularCommercial.Count > 0)
            {
                //CHECK IF FILE EXIST
                if (!File.Exists(path))
                    File.WriteAllBytes(path, Properties.Resources.Packing);

                OleDbConnection conn = new OleDbConnection(connectionString);

                OleDbCommand cmd;

                conn.Open();

                cmd = new OleDbCommand("DELETE FROM PACKING", conn);

                cmd.ExecuteNonQuery();

                blockNo = 1; 
                
                blockCounter = 0;

                foreach(var check in _orders.RegularPersonal)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();       
                    
                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," + 
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" + 
                        check.AccountNo + "','" + check.CheckType + "','" + 
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }//END FOR EACH

                blockNo = 1;
                blockCounter = 0;

                foreach(var check in _orders.RegularCommercial)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }
                conn.Close();
            }
            #endregion

            #region PreEncoded Checks
            //REGULAR CHECKS
            path = regPrePath + "\\Packing.dbf";

            connectionString = "Provider=VFPOLEDB.1; Data Source=" + path + "; Mode=ReadWrite;";

            query = "";

            if (_orders.PersonalPreEncoded.Count > 0 || _orders.CommercialPreEncoded.Count > 0)
            {
                //CHECK IF FILE EXIST
                if (!File.Exists(path))
                    File.WriteAllBytes(path, Properties.Resources.Packing);

                OleDbConnection conn = new OleDbConnection(connectionString);

                OleDbCommand cmd;

                conn.Open();

                cmd = new OleDbCommand("DELETE FROM PACKING", conn);

                cmd.ExecuteNonQuery();

                blockNo = 1;

                blockCounter = 0;

                foreach (var check in _orders.PersonalPreEncoded)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }//END FOR EACH

                blockNo = 1;
                blockCounter = 0;

                foreach (var check in _orders.CommercialPreEncoded)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }
                conn.Close();
            }
            #endregion

            #region CheckOne Checks
            //REGULAR CHECKS
            path = checkOnePath + "\\Packing.dbf";

            connectionString = "Provider=VFPOLEDB.1; Data Source=" + path + "; Mode=ReadWrite;";

            query = "";

            if (_orders.CheckOnePersonal.Count > 0 || _orders.CheckOneCommerical.Count > 0)
            {
                //CHECK IF FILE EXIST
                if (!File.Exists(path))
                    File.WriteAllBytes(path, Properties.Resources.Packing);

                OleDbConnection conn = new OleDbConnection(connectionString);

                OleDbCommand cmd;

                conn.Open();

                cmd = new OleDbCommand("DELETE FROM PACKING", conn);

                cmd.ExecuteNonQuery();

                blockNo = 1;

                blockCounter = 0;

                foreach (var check in _orders.CheckOnePersonal)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }//END FOR EACH

                blockNo = 1;
                blockCounter = 0;

                foreach (var check in _orders.CheckOneCommerical)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }
                conn.Close();
            }
            #endregion

            #region CheckPower Checks
            //REGULAR CHECKS
            path = checkPowerPath + "\\Packing.dbf";

            connectionString = "Provider=VFPOLEDB.1; Data Source=" + path + "; Mode=ReadWrite;";

            query = "";

            if (_orders.CheckPowerPersonal.Count > 0 || _orders.CheckPowerCommercial.Count > 0)
            {
                //CHECK IF FILE EXIST
                if (!File.Exists(path))
                    File.WriteAllBytes(path, Properties.Resources.Packing);

                OleDbConnection conn = new OleDbConnection(connectionString);

                OleDbCommand cmd;

                conn.Open();

                cmd = new OleDbCommand("DELETE FROM PACKING", conn);

                cmd.ExecuteNonQuery();

                blockNo = 1;

                blockCounter = 0;

                foreach (var check in _orders.CheckPowerPersonal)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 7)
                        start = "0" + start;

                    while (end.Length < 7)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }//END FOR EACH

                blockNo = 1;
                blockCounter = 0;

                foreach (var check in _orders.CheckPowerCommercial)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }
                conn.Close();
            }
            #endregion

            #region Manager's Checks
            //REGULAR CHECKS
            path = mcPath + "\\Packing.dbf";

            connectionString = "Provider=VFPOLEDB.1; Data Source=" + path + "; Mode=ReadWrite;";

            query = "";

            if (_orders.ManagersCheck.Count > 0)
            {
                //CHECK IF FILE EXIST
                if (!File.Exists(path))
                    File.WriteAllBytes(path, Properties.Resources.Packing);

                OleDbConnection conn = new OleDbConnection(connectionString);

                OleDbCommand cmd;

                conn.Open();

                cmd = new OleDbCommand("DELETE FROM PACKING", conn);

                cmd.ExecuteNonQuery();

                blockNo = 1;

                blockCounter = 0;

                foreach (var check in _orders.ManagersCheck)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }
                conn.Close();
            }
            #endregion

            #region Manager's Contl Checks
            //REGULAR CHECKS
            path = mcContPath + "\\Packing.dbf";

            connectionString = "Provider=VFPOLEDB.1; Data Source=" + path + "; Mode=ReadWrite;";

            query = "";

            if (_orders.ManagersCheckCont.Count > 0)
            {
                //CHECK IF FILE EXIST
                if (!File.Exists(path))
                    File.WriteAllBytes(path, Properties.Resources.Packing);

                OleDbConnection conn = new OleDbConnection(connectionString);

                OleDbCommand cmd;

                conn.Open();

                cmd = new OleDbCommand("DELETE FROM PACKING", conn);

                cmd.ExecuteNonQuery();

                blockNo = 1;

                blockCounter = 0;

                foreach (var check in _orders.ManagersCheckCont)
                {
                    string start = check.StartingSerial.ToString(), end = check.EndingSerial.ToString();

                    while (start.Length < 10)
                        start = "0" + start;

                    while (end.Length < 10)
                        end = "0" + end;

                    //initialize SQL command
                    query = "INSERT INTO PACKING (BATCHNO, BLOCK, RT_NO, BRANCH, ACCT_NO, ACCT_NO_P, CHKTYPE, ACCT_NAME1, ACCT_NAME2," +
                        "NO_BKS, CK_NO_P, CK_NO_B, CK_NOE, CK_NO_E, DELIVERTO, M) VALUES('" + check.Batch + "'," +
                        blockNo.ToString() + ",'" + check.BRSTN + "','" + check.Address1 + "','" + check.AccountNo + "','" +
                        check.AccountNo + "','" + check.CheckType + "','" +
                        check.Name.Replace("'", "''") + "','" + check.Name2.Replace("'", "''") + "',1," +
                        start + ",'" + start + "'," + end + ",'" + end + "','', '')";

                    cmd = new OleDbCommand(query, conn) ;

                    cmd.ExecuteNonQuery();

                    if (blockCounter < 4)
                        blockCounter++;
                    else
                    {
                        blockNo++;

                        blockCounter = 0;
                    }
                }
                conn.Close();
            }
            #endregion
        }//END FUNCTION

        public static void GeneratePrinterFiles(OrderSorted _orders, string _batch, string _ext)
        {
            #region Regular Personal
            if (_orders.RegularPersonal.Count > 0)
            {

                if (!Directory.Exists(regPath))
                    Directory.CreateDirectory(regPath);

                StreamWriter sw;

                string fileName = regPath + "\\" + _batch.Substring(0, 4) + "_P12" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.RegularPersonal.ForEach(regPersonal =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(regPersonal.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(regPersonal.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQty = (regPersonal.StartingSerial + QuantityPerBooklet.RegularPersonal).ToString();

                        while (startQty.Length < 7)
                            startQty = "0" + startQty;

                        sw.WriteLine(startQty.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(regPersonal.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + regPersonal.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((regPersonal.AccountNo.Substring(0, 3) + "-" + regPersonal.AccountNo.Substring(3, 6) + "-" + regPersonal.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(regPersonal.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(regPersonal.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(regPersonal.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(regPersonal.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(regPersonal.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(regPersonal.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(regPersonal.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(regPersonal.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = regPersonal.StartingSerial.ToString();

                        while (start.Length < 7)
                            start = "0" + start;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries

                        string end = regPersonal.EndingSerial.ToString();

                        while (end.Length < 7)
                            end = "0" + end;

                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });//END OF USING STREAMWRITER
                }//END OF USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(regPath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region Regular Commerical
            if (_orders.RegularCommercial.Count > 0)
            {
                if (!Directory.Exists(regPath))
                    Directory.CreateDirectory(regPath);

                StreamWriter sw;

                string fileName = regPath + "\\" + _batch.Substring(0, 4) + "_C12" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.RegularCommercial.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQty = (c.StartingSerial + QuantityPerBooklet.RegularCommercial).ToString();

                        while (startQty.Length < 10)
                            startQty = "0" + startQty;

                        sw.WriteLine(startQty.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(c.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 10)
                            start = "0" + start;

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 10)
                            end = "0" + end;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries
                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });
                }//END OF USING
                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(regPath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region Personal PreEncoded
            if (_orders.PersonalPreEncoded.Count > 0)
            {
                if (!Directory.Exists(regPrePath))
                    Directory.CreateDirectory(regPrePath);

                StreamWriter sw;

                string fileName = regPrePath + "\\" + _batch.Substring(0, 4) + "P" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using(sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.PersonalPreEncoded.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQty = (c.StartingSerial + QuantityPerBooklet.RegularPersonalPre).ToString();

                        while (startQty.Length < 7)
                            startQty = "0" + startQty;

                        sw.WriteLine(startQty.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine("**********"); // 10 - NAME 1 (FIXED)
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 7)
                            start = "0" + start;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 7)
                            end = "0" + end;

                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });//END OF FOREACH
                }//END OF USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(regPrePath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region Commercial PreEncoded
            if (_orders.CommercialPreEncoded.Count > 0)
            {
                if (!Directory.Exists(regPrePath))
                    Directory.CreateDirectory(regPrePath);

                StreamWriter sw;

                string fileName = regPrePath + "\\" + _batch.Substring(0, 4) + "C" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using(sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.CommercialPreEncoded.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQty = (c.StartingSerial + QuantityPerBooklet.RegularCommercialPre).ToString();

                        while (startQty.Length < 10)
                            startQty = "0" + startQty;

                        sw.WriteLine(startQty.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine("**********"); // 10 - NAME 1 (FIXED)
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 10)
                            start = "0" + start;

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 10)
                            end = "0" + end;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries
                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });//END OF FOREACH
                }//END OF USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(regPrePath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region CheckOne Personal
            if (_orders.CheckOnePersonal.Count > 0)
            {
                if (!Directory.Exists(checkOnePath))
                    Directory.CreateDirectory(checkOnePath);

                StreamWriter sw;

                string fileName = checkOnePath + "\\13D" + _batch.Substring(0, 4) + "P" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.CheckOnePersonal.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQTY = (c.StartingSerial + QuantityPerBooklet.CheckOnePersonal).ToString();

                        while (startQTY.Length < 7)
                            startQTY = "0" + startQTY;

                        sw.WriteLine(startQTY.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(c.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        if (start.Length < 7)
                            start = "0" + start;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries

                        string end = c.EndingSerial.ToString();

                        if (end.Length < 7)
                            end = "0" + end;

                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });//END OF FOREACH
                }//END OF USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(checkOnePath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\CHECKPOWER\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END OF IF
            #endregion

            #region CheckOne Commercial
            if (_orders.CheckOneCommerical.Count > 0)
            {
                if (!Directory.Exists(checkOnePath))
                    Directory.CreateDirectory(checkOnePath);

                StreamWriter sw;

                string fileName = checkOnePath + "\\13D" + _batch.Substring(0, 4) + "C" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using(sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.CheckOneCommerical.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQTY = (c.StartingSerial + QuantityPerBooklet.CheckOneCommercial).ToString();

                        while (startQTY.Length < 10)
                            startQTY = "0" + startQTY;

                        sw.WriteLine(startQTY.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(c.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 10)
                            start = "0" + start;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 10)
                            end = "0" + end;

                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });//END FOREACH
                }//END USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(checkOnePath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\CHECKONE\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region CheckPower Personal
            if (_orders.CheckPowerPersonal.Count > 0)
            {
                if (!Directory.Exists(checkPowerPath))
                    Directory.CreateDirectory(checkPowerPath);

                StreamWriter sw;

                string fileName = checkPowerPath + "\\CKP" + _batch.Substring(0, 4) + "P" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.CheckPowerPersonal.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQTY = (c.StartingSerial + QuantityPerBooklet.CheckPowerPersonal).ToString();

                        while (startQTY.Length < 7)
                            startQTY = "0" + startQTY;

                        sw.WriteLine(startQTY.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(c.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 7)
                            start = "0" + start;

                        sw.WriteLine(c.StartingSerial.ToString().Trim(' ')); // 31 - StartingSeries

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 7)
                            end = "0" + end;

                        sw.WriteLine(c.EndingSerial.ToString().Trim(' ')); // 32 - EndingSeries
                    });//END FOREACH
                }//END USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(checkPowerPath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\CHECKPOWER\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region CheckPower Commercial
            if (_orders.CheckPowerCommercial.Count > 0)
            {
                if (!Directory.Exists(checkOnePath))
                    Directory.CreateDirectory(checkOnePath);

                StreamWriter sw;

                string fileName = checkOnePath + "\\CKP" + _batch.Substring(0, 4) + "C" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using(sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.CheckPowerCommercial.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQTY = (c.StartingSerial + QuantityPerBooklet.CheckPowerCommercial).ToString();

                        while (startQTY.Length < 10)
                            startQTY = "0" + startQTY;

                        sw.WriteLine((c.StartingSerial + 50).ToString().Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(c.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 10)
                            start = "0" + start;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 10)
                            end = "0" + end;

                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });
                }//END USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(checkPowerPath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\CHECKPOWER\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region Manager's Check
            if (_orders.ManagersCheck.Count > 0)
            {
                if (!Directory.Exists(mcPath))
                    Directory.CreateDirectory(mcPath);

                StreamWriter sw;

                string fileName = mcPath + "\\MC" + _batch.Substring(0, 4) + "P" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using (sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.ManagersCheck.ForEach(c =>
                    {
                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQTY = (c.StartingSerial + QuantityPerBooklet.ManagersCheck).ToString();

                        while (startQTY.Length < 10)
                            startQTY = "0" + startQTY;

                        sw.WriteLine(startQTY.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(c.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 10)
                            start = "0" + start;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 10)
                            end = "0" + end;

                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });//END FOREACH
                }//END USING

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(mcPath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\MC\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion

            #region Manager's Check Cont
            if (_orders.ManagersCheckCont.Count > 0)
            {
                if (!Directory.Exists(mcContPath))
                    Directory.CreateDirectory(mcContPath);

                StreamWriter sw;

                string fileName = mcContPath + "\\MCC" + _batch.Substring(0, 4) + "B" + _ext + ".txt";

                sw = File.CreateText(fileName);
                sw.Close();

                using(sw = new StreamWriter(File.Open(fileName, FileMode.Append)))
                {
                    _orders.ManagersCheckCont.ForEach(c =>
                    {

                        sw.WriteLine("3"); //1 - FIXED;
                        sw.WriteLine(c.BRSTN.Trim(' ')); //2 - BRSTN or RT No
                        sw.WriteLine(c.AccountNo.Trim(' ')); //3 - AccountNo

                        string startQTY = (c.StartingSerial + QuantityPerBooklet.ManagersCheckCont).ToString();

                        while (startQTY.Length < 10)
                            startQTY = "0" + startQTY;

                        sw.WriteLine(startQTY.Trim(' ')); //4 - Starting Serial + 50pcs per booklet
                        sw.WriteLine("A"); // 5 - FIXED
                        sw.WriteLine(""); // 6 - BLANK
                        sw.WriteLine(c.BRSTN.Substring(0, 5).Trim(' ')); // 7 - First 5 Digits BRSTN
                        sw.WriteLine(" " + c.BRSTN.Substring(5, 4).Trim(' ')); // 8 - Last 4 Digits of BRSTN
                        sw.WriteLine((c.AccountNo.Substring(0, 3) + "-" + c.AccountNo.Substring(3, 6) + "-" + c.AccountNo.Substring(9, 3)).Trim(' ')); // 9 - ACCTNO
                        sw.WriteLine(c.Name.Trim(' ')); // 10 - NAME 1
                        sw.WriteLine("SN"); // 11 - FIXED
                        sw.WriteLine(""); // 12 - BLANK
                        sw.WriteLine(c.Name2.Trim(' ')); // 13 - NAME 2
                        sw.WriteLine("C"); // 14 - FIXED
                        sw.WriteLine("XXXX"); // 15 - FIXED
                        sw.WriteLine(""); // 16 - BLANK
                        sw.WriteLine(c.Address1.Trim(' ')); // 17 - Address 1 or BranchName
                        sw.WriteLine(c.Address2.Trim(' ')); // 18 - Address 2
                        sw.WriteLine(c.Address3.Trim(' ')); // 19 - Address 3
                        sw.WriteLine(c.Address4.Trim(' ')); // 20 - Address 4
                        sw.WriteLine(c.Address5.Trim(' ')); // 21 - Address 5
                        sw.WriteLine(c.Address6.Trim(' ')); // 22 - Address 6
                        sw.WriteLine("SECURITY BANK"); // 23 - FIXED
                        sw.WriteLine(""); // 24 - BLANK
                        sw.WriteLine(""); // 25 - BLANK
                        sw.WriteLine(""); // 26 - BLANK
                        sw.WriteLine(""); // 27 - BLANK
                        sw.WriteLine(""); // 28 - BLANK
                        sw.WriteLine(""); // 29 - BLANK
                        sw.WriteLine(""); // 30 - BLANK

                        string start = c.StartingSerial.ToString();

                        while (start.Length < 10)
                            start = "0" + start;

                        sw.WriteLine(start.Trim(' ')); // 31 - StartingSeries

                        string end = c.EndingSerial.ToString();

                        while (end.Length < 10)
                            end = "0" + end;

                        sw.WriteLine(end.Trim(' ')); // 32 - EndingSeries
                    });//END FOREACH
                }//END USING    

                if (_batch != "0000")
                    File.Copy(fileName, fileName.Replace(regPath, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\MC\\CONTINUOUS\\" + DateTime.Now.Year.ToString() + "\\"), true);
            }//END IF
            #endregion
        }//END OF FUNCTION

        public static void GenerateMDBFile(OrderSorted _orders, string _batch, string _ext)
        {
            if(_orders.ManagersCheck.Count > 0)
            {
                string fileName = mcPath + "\\MC" + _batch.Substring(0, 4) + _ext + ".mdb";

                FileInfo fileInfo = new FileInfo(fileName);

                if (fileInfo.Exists)
                    fileInfo.Delete();

                string mdbConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + fileName + "; Jet OLEDB:Engine Type=5";

                Catalog tClass = new Catalog();

                tClass.Create(mdbConn);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(tClass);

                GC.Collect();

                tClass = new Catalog();

                tClass.let_ActiveConnection(mdbConn);

                for (int x = 1; x <= 4; x++)
                {

                    Table tTable = new Table();

                    string tableTemp = "";

                    if (x == 1)
                        tableTemp = "Out";
                    else
                        tableTemp = "Outs";

                    tTable.Name = "InputFile_" + x.ToString() + tableTemp;

                    //COLUMN NAMES
                    tTable.Columns.Append("BRSTN", DataTypeEnum.adVarWChar, 10);

                    tTable.Columns.Append("AccountNumber", DataTypeEnum.adVarWChar, 30);

                    tTable.Columns.Append("RT1to5", DataTypeEnum.adVarWChar, 5);

                    tTable.Columns.Append("RT6to9", DataTypeEnum.adVarWChar, 5);

                    tTable.Columns.Append("AccountNumberWithHypen", DataTypeEnum.adVarWChar, 30);

                    tTable.Columns.Append("Serial", DataTypeEnum.adVarWChar, 10);

                    tTable.Columns.Append("Name1", DataTypeEnum.adVarWChar, 50);

                    tTable.Columns.Append("Name2", DataTypeEnum.adVarWChar, 50);

                    tTable.Columns.Append("Name3", DataTypeEnum.adVarWChar, 50);

                    tTable.Columns.Append("Address1", DataTypeEnum.adVarWChar, 100);

                    tTable.Columns.Append("Address2", DataTypeEnum.adVarWChar, 100);

                    tTable.Columns.Append("Address3", DataTypeEnum.adVarWChar, 100);

                    tTable.Columns.Append("Address4", DataTypeEnum.adVarWChar, 100);

                    tTable.Columns.Append("Address5", DataTypeEnum.adVarWChar, 100);

                    tTable.Columns.Append("Address6", DataTypeEnum.adVarWChar, 100);

                    tTable.Columns.Append("BankName", DataTypeEnum.adVarWChar, 100);

                    tTable.Columns.Append("StartingSerial", DataTypeEnum.adVarWChar, 20);

                    tTable.Columns.Append("EndingSerial", DataTypeEnum.adVarWChar, 20);

                    tTable.Columns.Append("PcsPerBook", DataTypeEnum.adVarWChar, 3);

                    tTable.Columns.Append("FileName", DataTypeEnum.adVarWChar, 30);

                    tTable.Columns.Append("PrimaryKey", DataTypeEnum.adVarWChar);

                    tTable.Columns["PrimaryKey"].Attributes = ColumnAttributesEnum.adColNullable;

                    tTable.Columns.Append("PageNumber", DataTypeEnum.adVarWChar);

                    tTable.Columns.Append("DataNumber", DataTypeEnum.adVarWChar, 20);

                    tClass.Tables.Append((object)tTable);
                }//END FOR

                GC.Collect(); //IF NOT INCLUDED FILE REMAIN IN OPENED STATUS

                OleDbConnection connection = new OleDbConnection(mdbConn);

                OleDbCommand cmd = new OleDbCommand();

                cmd.Connection = connection;

                cmd.Connection.Open();

                int primaryKey = 1;

                string txtName = "MC" + _batch.Substring(0, 4) + _ext + ".txt";

                int dataNumber1 = 0, dataNumber2 = 0, dataNumber3 = 0, dataNumber4 = 0;//SERVE AS DATANUMBER

                #region 1Out Format
                //1OUTS FORMAT
                foreach (var check in _orders.ManagersCheck)
                {
                    string RTFirst = check.BRSTN.Substring(0,5);

                    string RTLast = check.BRSTN.Substring(check.BRSTN.Length - 4, 4);

                    string startSeries = check.StartingSerial.ToString();

                    string endSeries = check.EndingSerial.ToString();

                    string acctNoHypen = check.AccountNo.Substring(0,3) + "-" + check.AccountNo.Substring(3,6) + "-" + check.AccountNo.Substring(9,3);

                    while (startSeries.Length < 10)
                        startSeries = "0" + startSeries;

                    while (endSeries.Length < 10)
                        endSeries = "0" + endSeries;

                    Int64 start = check.StartingSerial;

                    while(start <= check.EndingSerial)
                    {
                        string temp = start.ToString();

                        while (temp.Length < 10)
                            temp = "0" + temp;

                        cmd.CommandText = "INSERT INTO InputFile_1Out (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + check.BRSTN + "','" + check.AccountNo + "','" + RTFirst + "','" + RTLast + "','" + acctNoHypen + "','" +
                            temp + "','" + check.Name + "','" + check.Name2 + "','','" + check.Address1.Replace("'","''") + "','" + check.Address2.Replace("'", "''") + "','" + check.Address3.Replace("'", "''") +
                            "','" + check.Address4.Replace("'", "''") + "','" + check.Address5.Replace("'", "''") + "','" + check.Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries + "','" + endSeries +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + primaryKey + "');";

                        cmd.ExecuteNonQuery();

                        start++;

                        primaryKey++;
                    }//END WHILE
                }//END FOR
                #endregion

                #region 2Outs Format
                //2Outs Format
                dataNumber1 = 0; 
                dataNumber2 = QuantityPerBooklet.ManagersCheck; //SERVE AS DATANUMBER

                primaryKey = 0;

                for (int x1 = 0; x1 < _orders.ManagersCheck.Count; x1++)
                {                  
                    Int64 start1 = _orders.ManagersCheck[x1].StartingSerial;

                    string RTFirst1 = _orders.ManagersCheck[x1].BRSTN.Substring(0, 5);

                    string RTLast1 = _orders.ManagersCheck[x1].BRSTN.Substring(_orders.ManagersCheck[x1].BRSTN.Length - 4, 4);

                    string startSeries1 = _orders.ManagersCheck[x1].StartingSerial.ToString();

                    string endSeries1 = _orders.ManagersCheck[x1].EndingSerial.ToString();

                    string acctNoHypen1 = _orders.ManagersCheck[x1].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x1].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x1].AccountNo.Substring(9, 3);

                    while (startSeries1.Length < 10)
                        startSeries1 = "0" + startSeries1;

                    while (endSeries1.Length < 10)
                        endSeries1 = "0" + endSeries1;

                    int x2 = 0;

                    Int64 start2 = 0;

                    string RTFirst2 = "", RTLast2 = "", startSeries2 = "", endSeries2 = "", acctNoHypen2 = "";

                    if (x1 + 1 < _orders.ManagersCheck.Count)
                    {
                        x2 = x1 + 1;

                        RTFirst2 = _orders.ManagersCheck[x2].BRSTN.Substring(0, 5);

                        RTLast2 = _orders.ManagersCheck[x2].BRSTN.Substring(_orders.ManagersCheck[x2].BRSTN.Length - 4, 4);

                        startSeries2 = _orders.ManagersCheck[x2].StartingSerial.ToString();

                        endSeries2 = _orders.ManagersCheck[x2].EndingSerial.ToString();

                        acctNoHypen2 = _orders.ManagersCheck[x2].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x2].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x2].AccountNo.Substring(9, 3);

                        while (startSeries2.Length < 10)
                            startSeries2 = "0" + startSeries2;

                        while (endSeries2.Length < 10)
                            endSeries2 = "0" + endSeries2;

                        start2 = _orders.ManagersCheck[x2].StartingSerial;

                        x1++;
                    }

                    for (int x = 0; x < QuantityPerBooklet.ManagersCheck; x++)
                    {
                        dataNumber1++;

                        string temp1 = start1.ToString(); //Serial with 0 format

                        while (temp1.Length < 10)
                            temp1 = "0" + temp1;

                        cmd.CommandText = "INSERT INTO InputFile_2Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x1].BRSTN + "','" + _orders.ManagersCheck[x1].AccountNo + "','" + RTFirst1 + "','" + RTFirst2 + "','" + acctNoHypen1 + "','" +
                            temp1 + "','" + _orders.ManagersCheck[x1].Name + "','" + _orders.ManagersCheck[x1].Name2 + "','','" + _orders.ManagersCheck[x1].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x1].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries1 + "','" + endSeries1 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber1 + "');";

                        cmd.ExecuteNonQuery();

                        start1++;

                        primaryKey++;

                        if (x1 + 1 < _orders.ManagersCheck.Count)
                        {
                            dataNumber2++;

                            string temp2 = start2.ToString();

                            while (temp2.Length < 10)
                                temp2 = "0" + temp2;

                            cmd.CommandText = "INSERT INTO InputFile_2Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x2].BRSTN + "','" + _orders.ManagersCheck[x2].AccountNo + "','" + RTFirst2 + "','" + RTFirst2 + "','" + acctNoHypen2 + "','" +
                            temp2 + "','" + _orders.ManagersCheck[x2].Name + "','" + _orders.ManagersCheck[x2].Name2 + "','','" + _orders.ManagersCheck[x2].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x2].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries2 + "','" + endSeries2 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            start2++;
                            primaryKey++;
                        }
                        else // INSERT BLANK FIELD
                        {
                            dataNumber2++;

                            cmd.CommandText = "INSERT INTO InputFile_2Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('','','','','','','','','','','','','','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            primaryKey++;
                        }

                    }//END WHILE
                }//END FOR
                #endregion

                #region 3Outs Format
                dataNumber1 = 0;
                dataNumber2 = QuantityPerBooklet.ManagersCheck; //SERVE AS DATANUMBER
                dataNumber3 = QuantityPerBooklet.ManagersCheck * 2;

                primaryKey = 0;

                for (int x1 = 0; x1 < _orders.ManagersCheck.Count; x1++)
                {
                    Int64 start1 = _orders.ManagersCheck[x1].StartingSerial;

                    string RTFirst1 = _orders.ManagersCheck[x1].BRSTN.Substring(0, 5);

                    string RTLast1 = _orders.ManagersCheck[x1].BRSTN.Substring(_orders.ManagersCheck[x1].BRSTN.Length - 4, 4);

                    string startSeries1 = _orders.ManagersCheck[x1].StartingSerial.ToString();

                    string endSeries1 = _orders.ManagersCheck[x1].EndingSerial.ToString();

                    string acctNoHypen1 = _orders.ManagersCheck[x1].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x1].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x1].AccountNo.Substring(9, 3);

                    while (startSeries1.Length < 10)
                        startSeries1 = "0" + startSeries1;

                    while (endSeries1.Length < 10)
                        endSeries1 = "0" + endSeries1;

                    bool Out2 = false, Out3 = false;

                    int x2 = 0;

                    Int64 start2 = 0;

                    string RTFirst2 = "", RTLast2 = "", startSeries2 = "", endSeries2 = "", acctNoHypen2 = "";

                    if (x1 + 1 < _orders.ManagersCheck.Count)
                    {
                        x2 = x1 + 1;

                        RTFirst2 = _orders.ManagersCheck[x2].BRSTN.Substring(0, 5);

                        RTLast2 = _orders.ManagersCheck[x2].BRSTN.Substring(_orders.ManagersCheck[x2].BRSTN.Length - 4, 4);

                        startSeries2 = _orders.ManagersCheck[x2].StartingSerial.ToString();

                        endSeries2 = _orders.ManagersCheck[x2].EndingSerial.ToString();

                        acctNoHypen2 = _orders.ManagersCheck[x2].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x2].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x2].AccountNo.Substring(9, 3);

                        while (startSeries2.Length < 10)
                            startSeries2 = "0" + startSeries2;

                        while (endSeries2.Length < 10)
                            endSeries2 = "0" + endSeries2;

                        start2 = _orders.ManagersCheck[x2].StartingSerial;

                        x1++;

                        Out2 = true;
                    }
                    else
                        Out2 = false;

                    int x3 = 0;

                    Int64 start3 = 0;

                    string RTFirst3 = "", RTLast3 = "", startSeries3 = "", endSeries3 = "", acctNoHypen3 = "";

                    if (x1 + 1 < _orders.ManagersCheck.Count)
                    {
                        x3 = x1 + 1;

                        RTFirst3 = _orders.ManagersCheck[x3].BRSTN.Substring(0, 5);

                        RTLast3 = _orders.ManagersCheck[x3].BRSTN.Substring(_orders.ManagersCheck[x3].BRSTN.Length - 4, 4);

                        startSeries3 = _orders.ManagersCheck[x3].StartingSerial.ToString();

                        endSeries3 = _orders.ManagersCheck[x3].EndingSerial.ToString();

                        acctNoHypen3 = _orders.ManagersCheck[x3].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x3].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x3].AccountNo.Substring(9, 3);

                        while (startSeries3.Length < 10)
                            startSeries3 = "0" + startSeries3;

                        while (endSeries3.Length < 10)
                            endSeries3 = "0" + endSeries3;

                        start3 = _orders.ManagersCheck[x3].StartingSerial;

                        x1++;

                        Out3 = true;
                    }
                    else
                        Out3 = false;

                    for (int x = 0; x < QuantityPerBooklet.ManagersCheck; x++)
                    {
                        dataNumber1++;

                        string temp1 = start1.ToString(); //Serial with 0 format

                        while (temp1.Length < 10)
                            temp1 = "0" + temp1;

                        cmd.CommandText = "INSERT INTO InputFile_3Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x1].BRSTN + "','" + _orders.ManagersCheck[x1].AccountNo + "','" + RTFirst1 + "','" + RTFirst2 + "','" + acctNoHypen1 + "','" +
                            temp1 + "','" + _orders.ManagersCheck[x1].Name + "','" + _orders.ManagersCheck[x1].Name2 + "','','" + _orders.ManagersCheck[x1].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x1].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries1 + "','" + endSeries1 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber1 + "');";

                        cmd.ExecuteNonQuery();

                        start1++;

                        primaryKey++;

                        if (Out2)
                        {
                            dataNumber2++;

                            string temp2 = start2.ToString();

                            while (temp2.Length < 10)
                                temp2 = "0" + temp2;

                            cmd.CommandText = "INSERT INTO InputFile_3Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x2].BRSTN + "','" + _orders.ManagersCheck[x2].AccountNo + "','" + RTFirst2 + "','" + RTFirst2 + "','" + acctNoHypen2 + "','" +
                            temp2 + "','" + _orders.ManagersCheck[x2].Name + "','" + _orders.ManagersCheck[x2].Name2 + "','','" + _orders.ManagersCheck[x2].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x2].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries2 + "','" + endSeries2 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            start2++;
                            primaryKey++;
                        }
                        else // INSERT BLANK FIELD
                        {
                            dataNumber2++;

                            cmd.CommandText = "INSERT INTO InputFile_3Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('','','','','','','','','','','','','','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            primaryKey++;
                        }

                        if (Out3)
                        {
                            dataNumber3++;

                            string temp3 = start3.ToString();

                            while (temp3.Length < 10)
                                temp3 = "0" + temp3;

                            cmd.CommandText = "INSERT INTO InputFile_3Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x3].BRSTN + "','" + _orders.ManagersCheck[x3].AccountNo + "','" + RTFirst3 + "','" + RTFirst3 + "','" + acctNoHypen3 + "','" +
                            temp3 + "','" + _orders.ManagersCheck[x3].Name + "','" + _orders.ManagersCheck[x3].Name2 + "','','" + _orders.ManagersCheck[x3].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x3].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries3 + "','" + endSeries3 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber3 + "');";

                            cmd.ExecuteNonQuery();

                            start3++;
                            primaryKey++;
                        }
                        else // INSERT BLANK FIELD
                        {
                            dataNumber2++;

                            cmd.CommandText = "INSERT INTO InputFile_3Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('','','','','','','','','','','','','','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            primaryKey++;
                        }


                    }//END WHILE
                }//END FOR

                #endregion

                #region 4Outs Format
                dataNumber1 = 0;
                dataNumber2 = QuantityPerBooklet.ManagersCheck; //SERVE AS DATANUMBER
                dataNumber3 = QuantityPerBooklet.ManagersCheck * 2;
                dataNumber4 = QuantityPerBooklet.ManagersCheck * 3;

                primaryKey = 0;

                for (int x1 = 0; x1 < _orders.ManagersCheck.Count; x1++)
                {
                    Int64 start1 = _orders.ManagersCheck[x1].StartingSerial;

                    string RTFirst1 = _orders.ManagersCheck[x1].BRSTN.Substring(0, 5);

                    string RTLast1 = _orders.ManagersCheck[x1].BRSTN.Substring(_orders.ManagersCheck[x1].BRSTN.Length - 4, 4);

                    string startSeries1 = _orders.ManagersCheck[x1].StartingSerial.ToString();

                    string endSeries1 = _orders.ManagersCheck[x1].EndingSerial.ToString();

                    string acctNoHypen1 = _orders.ManagersCheck[x1].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x1].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x1].AccountNo.Substring(9, 3);

                    while (startSeries1.Length < 10)
                        startSeries1 = "0" + startSeries1;

                    while (endSeries1.Length < 10)
                        endSeries1 = "0" + endSeries1;

                    int x2 = 0;

                    Int64 start2 = 0;

                    string RTFirst2 = "", RTLast2 = "", startSeries2 = "", endSeries2 = "", acctNoHypen2 = "";

                    bool Out2 = false, Out3 = false, Out4 = false;

                    //For 2ndLine
                    if (x1 + 1 < _orders.ManagersCheck.Count)
                    {
                        x2 = x1 + 1;

                        RTFirst2 = _orders.ManagersCheck[x2].BRSTN.Substring(0, 5);

                        RTLast2 = _orders.ManagersCheck[x2].BRSTN.Substring(_orders.ManagersCheck[x2].BRSTN.Length - 4, 4);

                        startSeries2 = _orders.ManagersCheck[x2].StartingSerial.ToString();

                        endSeries2 = _orders.ManagersCheck[x2].EndingSerial.ToString();

                        acctNoHypen2 = _orders.ManagersCheck[x2].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x2].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x2].AccountNo.Substring(9, 3);

                        while (startSeries2.Length < 10)
                            startSeries2 = "0" + startSeries2;

                        while (endSeries2.Length < 10)
                            endSeries2 = "0" + endSeries2;

                        start2 = _orders.ManagersCheck[x2].StartingSerial;

                        x1++;

                        Out2 = true;
                    }
                    else
                        Out2 = false;

                    int x3 = 0;

                    Int64 start3 = 0;

                    string RTFirst3 = "", RTLast3 = "", startSeries3 = "", endSeries3 = "", acctNoHypen3 = "";

                    //FOR 3rdLine
                    if (x1 + 1 < _orders.ManagersCheck.Count)
                    {
                        x3 = x1 + 1;

                        RTFirst3 = _orders.ManagersCheck[x3].BRSTN.Substring(0, 5);

                        RTLast3 = _orders.ManagersCheck[x3].BRSTN.Substring(_orders.ManagersCheck[x3].BRSTN.Length - 4, 4);

                        startSeries3 = _orders.ManagersCheck[x3].StartingSerial.ToString();

                        endSeries3 = _orders.ManagersCheck[x3].EndingSerial.ToString();

                        acctNoHypen3 = _orders.ManagersCheck[x3].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x3].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x3].AccountNo.Substring(9, 3);

                        while (startSeries3.Length < 10)
                            startSeries3 = "0" + startSeries3;

                        while (endSeries3.Length < 10)
                            endSeries3 = "0" + endSeries3;

                        start3 = _orders.ManagersCheck[x3].StartingSerial;

                        x1++;

                        Out3 = true;
                    }
                    else
                        Out3 = false;

                    int x4 = 0;

                    Int64 start4 = 0;

                    string RTFirst4 = "", RTLast4 = "", startSeries4 = "", endSeries4 = "", acctNoHypen4 = "";

                    //For 4thLine
                    if (x1 + 1 < _orders.ManagersCheck.Count)
                    {
                        x4 = x1 + 1;

                        RTFirst4 = _orders.ManagersCheck[x4].BRSTN.Substring(0, 5);

                        RTLast4 = _orders.ManagersCheck[x4].BRSTN.Substring(_orders.ManagersCheck[x4].BRSTN.Length - 4, 4);

                        startSeries4 = _orders.ManagersCheck[x4].StartingSerial.ToString();

                        endSeries4 = _orders.ManagersCheck[x4].EndingSerial.ToString();

                        acctNoHypen4 = _orders.ManagersCheck[x4].AccountNo.Substring(0, 3) + "-" + _orders.ManagersCheck[x4].AccountNo.Substring(3, 6) + "-" + _orders.ManagersCheck[x4].AccountNo.Substring(9, 3);

                        while (startSeries4.Length < 10)
                            startSeries4 = "0" + startSeries4;

                        while (endSeries4.Length < 10)
                            endSeries4 = "0" + endSeries4;

                        start4 = _orders.ManagersCheck[x4].StartingSerial;

                        x1++;

                        Out4 = true;
                    }
                    else
                        Out4 = false;

                    for (int x = 0; x < QuantityPerBooklet.ManagersCheck; x++)
                    {
                        dataNumber1++;

                        string temp1 = start1.ToString(); //Serial with 0 format

                        while (temp1.Length < 10)
                            temp1 = "0" + temp1;

                        cmd.CommandText = "INSERT INTO InputFile_4Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x1].BRSTN + "','" + _orders.ManagersCheck[x1].AccountNo + "','" + RTFirst1 + "','" + RTFirst2 + "','" + acctNoHypen1 + "','" +
                            temp1 + "','" + _orders.ManagersCheck[x1].Name + "','" + _orders.ManagersCheck[x1].Name2 + "','','" + _orders.ManagersCheck[x1].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x1].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x1].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries1 + "','" + endSeries1 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber1 + "');";

                        cmd.ExecuteNonQuery();

                        start1++;

                        primaryKey++;

                        if (Out2)
                        {
                            dataNumber2++;

                            string temp2 = start2.ToString();

                            while (temp2.Length < 10)
                                temp2 = "0" + temp2;

                            cmd.CommandText = "INSERT INTO InputFile_4Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x2].BRSTN + "','" + _orders.ManagersCheck[x2].AccountNo + "','" + RTFirst2 + "','" + RTFirst2 + "','" + acctNoHypen2 + "','" +
                            temp2 + "','" + _orders.ManagersCheck[x2].Name + "','" + _orders.ManagersCheck[x2].Name2 + "','','" + _orders.ManagersCheck[x2].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x2].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x2].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries2 + "','" + endSeries2 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            start2++;
                            primaryKey++;
                        }
                        else // INSERT BLANK FIELD
                        {
                            dataNumber2++;

                            cmd.CommandText = "INSERT INTO InputFile_4Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('','','','','','','','','','','','','','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            primaryKey++;
                        }

                        if (Out3)
                        {
                            dataNumber3++;

                            string temp3 = start3.ToString();

                            while (temp3.Length < 10)
                                temp3 = "0" + temp3;

                            cmd.CommandText = "INSERT INTO InputFile_4Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x3].BRSTN + "','" + _orders.ManagersCheck[x3].AccountNo + "','" + RTFirst3 + "','" + RTFirst3 + "','" + acctNoHypen3 + "','" +
                            temp3 + "','" + _orders.ManagersCheck[x3].Name + "','" + _orders.ManagersCheck[x3].Name2 + "','','" + _orders.ManagersCheck[x3].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x3].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x3].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries3 + "','" + endSeries3 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber3 + "');";

                            cmd.ExecuteNonQuery();

                            start3++;
                            primaryKey++;
                        }
                        else // INSERT BLANK FIELD
                        {
                            dataNumber2++;

                            cmd.CommandText = "INSERT INTO InputFile_4Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('','','','','','','','','','','','','','','','','','','','','" + primaryKey + "','0','" + dataNumber2 + "');";

                            cmd.ExecuteNonQuery();

                            primaryKey++;
                        }

                        if (Out4)
                        {
                            dataNumber4++;

                            string temp4 = start4.ToString();

                            while (temp4.Length < 10)
                                temp4 = "0" + temp4;

                            cmd.CommandText = "INSERT INTO InputFile_4Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('" + _orders.ManagersCheck[x4].BRSTN + "','" + _orders.ManagersCheck[x4].AccountNo + "','" + RTFirst4 + "','" + RTFirst4 + "','" + acctNoHypen4 + "','" +
                            temp4 + "','" + _orders.ManagersCheck[x4].Name + "','" + _orders.ManagersCheck[x4].Name2 + "','','" + _orders.ManagersCheck[x4].Address1.Replace("'", "''") + "','" + _orders.ManagersCheck[x4].Address2.Replace("'", "''") + "','" + _orders.ManagersCheck[x4].Address3.Replace("'", "''") +
                            "','" + _orders.ManagersCheck[x4].Address4.Replace("'", "''") + "','" + _orders.ManagersCheck[x4].Address5.Replace("'", "''") + "','" + _orders.ManagersCheck[x4].Address6.Replace("'", "''") + "','SECURITY BANK','" + startSeries4 + "','" + endSeries4 +
                            "','50','" + txtName + "','" + primaryKey + "','0','" + dataNumber4 + "');";

                            cmd.ExecuteNonQuery();

                            start4++;

                            primaryKey++;
                        }
                        else // INSERT BLANK FIELD
                        {
                            dataNumber4++;

                            cmd.CommandText = "INSERT INTO InputFile_4Outs (BRSTN, AccountNumber, RT1to5, RT6to9, AccountNumberWithHypen, Serial, Name1, " +
                            "Name2, Name3, Address1, Address2, Address3, Address4, Address5, Address6, BankName, StartingSerial, EndingSerial, " +
                            "PcsPerBook, FileName, PrimaryKey, PageNumber, DataNumber) " +
                            "VALUES ('','','','','','','','','','','','','','','','','','','','','" + primaryKey + "','0','" + dataNumber4 + "');";

                            cmd.ExecuteNonQuery();

                            primaryKey++;
                        }

                    }//END WHILE
                }//END FOR

                #endregion

                connection.Close();

                connection.Dispose();

                GC.Collect();

                if (!Directory.Exists("\\\\192.168.0.254\\PrinterFiles\\SBTC\\MC\\" + DateTime.Now.Year))
                    Directory.CreateDirectory("\\\\192.168.0.254\\PrinterFiles\\SBTC\\MC\\" + DateTime.Now.Year);

                if(_batch != "0000")
                    File.Copy(fileName, "\\\\192.168.0.254\\PrinterFiles\\SBTC\\MC\\" + DateTime.Now.Year + "\\" + fileName.Replace(mcPath, ""), true);
            }//END IF
        }

        #region private class
        public static void CheckPaths()
        {
            if (Directory.Exists(regPrePath))
            {
                DeleteFilesInDirectory(regPrePath);

                Directory.Delete(regPrePath);
            }

            if (Directory.Exists(regPath))
            {
                DeleteFilesInDirectory(regPath);

                Directory.Delete(regPath);
            }

            if (Directory.Exists(chargeSlipPath))
            {
                DeleteFilesInDirectory(chargeSlipPath);

                Directory.Delete(chargeSlipPath);
            }

            if (Directory.Exists(checkOnePath))
            {
                DeleteFilesInDirectory(checkOnePath);

                Directory.Delete(checkOnePath);
            }

            if (Directory.Exists(checkPowerPath))
            {
                DeleteFilesInDirectory(checkPowerPath);

                Directory.Delete(checkPowerPath);
            }

            if (Directory.Exists(customPath))
            {
                DeleteFilesInDirectory(customPath);

                Directory.Delete(customPath);
            }

            if (Directory.Exists(gcPath))
            {
                DeleteFilesInDirectory(gcPath);

                Directory.Delete(gcPath);
            }


            if (Directory.Exists(mcContPath))
            {
                DeleteFilesInDirectory(mcContPath);

                Directory.Delete(mcContPath);
            }

            if (Directory.Exists(mcPath))
            {
                DeleteFilesInDirectory(mcPath);

                Directory.Delete(mcPath);
            }
        }//END OF FUNCTION
        private static void DeleteFilesInDirectory(string _directory)
        {
            DirectoryInfo info = new DirectoryInfo(_directory);

            foreach(FileInfo file in info.GetFiles())
            {
                file.Delete();
            }    
        }//END OF FUNCTION
        #endregion
    }//END OF CLASS
}//END OF NAMESPACE
