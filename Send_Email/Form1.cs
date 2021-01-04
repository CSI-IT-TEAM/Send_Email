using DevExpress.Utils;
using JPlatform.Client.Controls6;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Send_Email
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            grdBase1.Size = new Size(1950, 1100);
            grdBase2.Size = new Size(1950, 1100);
            grdBase3.Size = new Size(1950, 1200);
            grdBase4.Size = new Size(1950, 1000);

            grdBaseNpi.Size = new Size(4000, 700);

            panel1.Size = new Size(1950, 1100);
            chart2.Size = new Size(1950, 1035);

            tmrLoad.Enabled = true;
            this.Text = "20210102080000";
        }

        DataTable dtEmail;
        bool _isRun = false, _isRun2 = false;
        int _start_column = 0;
        //"jungbo.shim@dskorea.com", "nguyen.it@changshininc.com", "dien.it@changshininc.com", "do.it@changshininc.com"
        //, "nguyen.it@changshininc.com", "dien.it@changshininc.com", "ngoc.it@changshininc.com", "yen.it@changshininc.com"
        //readonly string[] _emailTest = {   "do.it@changshininc.com", "nguyen.it@changshininc.com", "dien.it@changshininc.com", "ngoc.it@changshininc.com", "yen.it@changshininc.com" };
        readonly string[] _emailTest = { "jungbo.shim@dskorea.com" };

        #region Event
        private void tmrLoad_Tick(object sender, EventArgs e)
        {
            //12h
            RunToPo("Q1");
            RunToPoIe("Q1");
            RunProduction("Q1");

            //7h
            RunEScan("Q1");
            RunAndon("Q1");
            Run("Q1");

            //8h
            RunTMSDash("Q1");

            RunNPI("Q1");

            //16h
            RunCutting("Q1");
        }

        private void cmdRunProd_Click(object sender, EventArgs e)
        {
            RunProduction("Q");
        }

        private void cmdRunAndon_Click(object sender, EventArgs e)
        {
            RunAndon("Q");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Run("Q");
        }

        private void cmdRunEscan_Click(object sender, EventArgs e)
        {
            RunEScan("Q");
        }

        private void cmdPoTo_Click(object sender, EventArgs e)
        {
            RunToPo("Q");
        }

        private void cmdPoToIe_Click(object sender, EventArgs e)
        {
            RunToPoIe("Q");
        }

        private void cmdCutting_Click(object sender, EventArgs e)
        {
            RunCutting("Q");
        }

        private void cmdNpi_Click(object sender, EventArgs e)
        {
            RunNPI("Q");
            //RunNPI2();
        }

        #endregion Event

        private void CreateMail(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email 
                if (app.Session.CurrentUser.AddressEntry.Address.Contains("IT.GMES"))
                {
                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }
                oRecips = null;
                mailItem.BCC = "phuoc.it@changshininc.com";

                mailItem.HTMLBody = htmlBody;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        #region Email TO/PO

        private void RunToPo(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_TO_PO(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;
                WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunToPo({argType}): Run");
                DataTable dtData = dsData.Tables[0];
                DataTable dtDate = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog($"  dtData:{dtData.Rows.Count};  dtDate: {dtEmail.Rows.Count}; dtEmail: {dtEmail.Rows.Count}");

                CreateMailToPo(dtDate, dtData, dtEmail);

                WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunToPo({argType}): End");
            }
            catch (Exception ex)
            {
                WriteLog($"  RunToPo({argType}): " + ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void CreateMailToPo(DataTable dtDate, DataTable dtData, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "TO&PO List";

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email 
                if (!app.Session.CurrentUser.AddressEntry.Address.Contains("IT.NGOC"))
                {
                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }

                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                mailItem.Body = "This is the message.";

                var query = from row in dtData.AsEnumerable()
                            group row by row.Field<string>("DEPT_NM") into dept
                            orderby dept.Key
                            select new
                            {
                                Name = dept.Key,
                                cntLine = dept.Count()
                            };
                System.Collections.Hashtable ht = new System.Collections.Hashtable();
                foreach (var row in query)
                {
                    ht.Add(row.Name, row.cntLine);
                }
                string[] strValue = new string[14];

                string strDate = "";
                foreach (DataRow row in dtDate.Rows)
                {
                    strDate += "<th bgcolor = '#cc66ff' style = 'color:#ffffff' align = 'center' width = '80' >" + row["YMD"].ToString() + " </th >";
                }

                string HtmlTableBody = "";

                string HtmlTableHeader =
                        "<tr>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' rowspan='2' width = '150'>Plant</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' rowspan='2' width = '150'>Line</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' rowspan='2' width = '80'>TO</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' rowspan='2' width = '80'>PO Actual</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' rowspan='2' width = '80'>Relief Actual</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' rowspan='2' width = '80'>Balance</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' colspan='7' >Staffing Ratio(%)</th>" +
                            "<th bgcolor='#ff9900' style='color:#ffffff' align='center' rowspan='2' width = '80'>Next TO</th>" +
                            "<th bgcolor='#ff9900' style='color:#ffffff' align='center' rowspan='2' width = '80'>Staffing Ratio(%)</th>" +
                        "</tr>" +
                        "<tr>" +
                            strDate +
                            "<th bgcolor='#0066ff' style='color:#ffffff' align='center' width = '80'>AVG</th>" +

                        "</tr>";

                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    string Plant = dtData.Rows[iRow]["DEPT_NM"].ToString();
                    string Line = dtData.Rows[iRow]["LINE_NAME"].ToString();
                    string To = dtData.Rows[iRow]["TO"].ToString();
                    string Po = dtData.Rows[iRow]["PO_ACTUAL"].ToString();
                    string Relief = dtData.Rows[iRow]["RELIEF"].ToString();
                    string Balance = dtData.Rows[iRow]["BALANCE"].ToString();

                    string Day1 = dtData.Rows[iRow]["DAY1"].ToString();
                    string Day1BgColor = dtData.Rows[iRow]["BG_COLOR_DAY1"].ToString();
                    string Day1ForeColor = dtData.Rows[iRow]["FORE_COLOR_DAY1"].ToString();

                    string Day2 = dtData.Rows[iRow]["DAY2"].ToString();
                    string Day2BgColor = dtData.Rows[iRow]["BG_COLOR_DAY2"].ToString();
                    string Day2ForeColor = dtData.Rows[iRow]["FORE_COLOR_DAY2"].ToString();

                    string Day3 = dtData.Rows[iRow]["DAY3"].ToString();
                    string Day3BgColor = dtData.Rows[iRow]["BG_COLOR_DAY3"].ToString();
                    string Day3ForeColor = dtData.Rows[iRow]["FORE_COLOR_DAY3"].ToString();

                    string Day4 = dtData.Rows[iRow]["DAY4"].ToString();
                    string Day4BgColor = dtData.Rows[iRow]["BG_COLOR_DAY4"].ToString();
                    string Day4ForeColor = dtData.Rows[iRow]["FORE_COLOR_DAY4"].ToString();

                    string Day5 = dtData.Rows[iRow]["DAY5"].ToString();
                    string Day5BgColor = dtData.Rows[iRow]["BG_COLOR_DAY5"].ToString();
                    string Day5ForeColor = dtData.Rows[iRow]["FORE_COLOR_DAY5"].ToString();

                    string Day6 = dtData.Rows[iRow]["DAY6"].ToString();
                    string Day6BgColor = dtData.Rows[iRow]["BG_COLOR_DAY6"].ToString();
                    string Day6ForeColor = dtData.Rows[iRow]["FORE_COLOR_DAY6"].ToString();

                    string DAvg = dtData.Rows[iRow]["DAVG"].ToString();
                    string DAvgBgColor = dtData.Rows[iRow]["BG_COLOR_DAVG"].ToString();
                    string DAvgForeColor = dtData.Rows[iRow]["FORE_COLOR_DAVG"].ToString();

                    string NextTo = dtData.Rows[iRow]["TO_TOTAL_WEEK"].ToString();
                    string Next = dtData.Rows[iRow]["NEXT"].ToString();
                    string NextBgColor = dtData.Rows[iRow]["BG_COLOR_NEXT"].ToString();
                    string NextForeColor = dtData.Rows[iRow]["FORE_COLOR_NEXT"].ToString();

                    string TotalBgColor = dtData.Rows[iRow]["TOTAL_BG_COLOR"].ToString();
                    string TotalForeColor = dtData.Rows[iRow]["TOTAL_FORE_COLOR"].ToString();

                    string rowspan = "";
                    if (iRow == 0)
                    {
                        rowspan = $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='left' rowspan='{ht[Plant]}' >{Plant}</td>";
                    }
                    else
                    {
                        rowspan = Plant == dtData.Rows[iRow - 1]["DEPT_NM"].ToString()
                            ? ""
                            : $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='left' rowspan='{ht[Plant]}' >{Plant}</td>";
                    }

                    HtmlTableBody +=
                            "<tr>" +
                                rowspan +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='left' >{Line}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{To}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{Po}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{Relief}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{Balance}</td>" +
                                $"<td bgcolor='{Day1BgColor}'  style='color:{Day1ForeColor}'  align='right'>{Day1}</td>" +
                                $"<td bgcolor='{Day2BgColor}'  style='color:{Day2ForeColor}'  align='right'>{Day2}</td>" +
                                $"<td bgcolor='{Day3BgColor}'  style='color:{Day3ForeColor}'  align='right'>{Day3}</td>" +
                                $"<td bgcolor='{Day4BgColor}'  style='color:{Day4ForeColor}'  align='right'>{Day4}</td>" +
                                $"<td bgcolor='{Day5BgColor}'  style='color:{Day5ForeColor}'  align='right'>{Day5}</td>" +
                                $"<td bgcolor='{Day6BgColor}'  style='color:{Day6ForeColor}'  align='right'>{Day6}</td>" +
                                $"<td bgcolor='{DAvgBgColor}'  style='color:{DAvgForeColor}'  align='right'>{DAvg}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{NextTo}</td>" +
                                $"<td bgcolor='{NextBgColor}'  style='color:{NextForeColor}' align='right'>{Next}</td>" +
                            "</tr>";
                }

                string html = "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;' >" +
                                      "<b style='background-color:black; color:yellow' >Formular and staffing ratio color explanation</b><br>" +
                                      "&nbsp;&nbsp;&nbsp;Balance = PO Actual + Relief – TO<br>" +
                                      "&nbsp;&nbsp;&nbsp;Staffing Ratio = (PO + Relief) / TO<br>" +
                                      "&nbsp;&nbsp;&nbsp;More than 105: red<br>" +
                                      "&nbsp;&nbsp;&nbsp;103 ~ 105: yellow<br>" +
                                      "&nbsp;&nbsp;&nbsp;97 ~ 103: green<br>" +
                                      "&nbsp;&nbsp;&nbsp;95 ~ 97: yellow<br>" +
                                      "&nbsp;&nbsp;&nbsp;Less than 95: grey<br>" +
                                      "* System will inform everyday and W/S have to check and update<br>" +
                                      "* IE will inform to workshop if continue red color for 3 days<br>" +
                                      "* If within 1 weeks, there are plant to change model, it will be allowed to keep manpower<br><br>" +
                                      "Remark:<br>" +
                                      "● Support include indirect / OH (Material handler, Office, W/H worker …)<br>" +
                                      "● TL / GL for line include in line<br>" +
                               "</p>" +
                              "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' >" +
                                    HtmlTableHeader +
                                    HtmlTableBody +
                              "</table>";

                mailItem.HTMLBody = html;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("  CreateMailToPo: " + ex.ToString());
            }
        }

        public DataSet SEL_TO_PO(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_TO_PO_V2";
                MyOraDB.ReDim_Parameter(5);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";
                MyOraDB.Parameter_Name[3] = "CV_2";
                MyOraDB.Parameter_Name[4] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("P_SEND_EMAIL_PROD: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_PROD_DATA: " + ex.ToString());
                return null;
            }
        }

        #endregion

        #region Email TO/PO IE

        private void RunToPoIe(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_TO_PO_IE(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;

                DataTable dtData = dsData.Tables[0];
                DataTable dtEmail = dsData.Tables[1];

                WriteLog(dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                CreateMailToPoIe(dtData, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog("RunToPo(): " + ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void CreateMailToPoIe(DataTable dtData, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "TO&PO List";

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email 
                if (!app.Session.CurrentUser.AddressEntry.Address.Contains("IT.NGOC"))
                {
                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }

                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                mailItem.Body = "This is the message.";

                string rowValue = "";

                var query = from row in dtData.AsEnumerable()
                            group row by row.Field<string>("DEPT_NM") into dept
                            orderby dept.Key
                            select new
                            {
                                Name = dept.Key,
                                cntLine = dept.Count()
                            };


                var query2 = from row in dtData.AsEnumerable()
                             group row by row.Field<string>("MLINE") into dept
                             orderby dept.Key
                             select new
                             {
                                 Name = dept.Key,
                                 cntLine = dept.Count()
                             };

                System.Collections.Hashtable ht = new System.Collections.Hashtable();
                System.Collections.Hashtable ht2 = new System.Collections.Hashtable();
                foreach (var row in query)
                {
                    ht.Add(row.Name, row.cntLine);
                }
                foreach (var row in query2)
                {
                    ht2.Add(row.Name, row.cntLine);
                }

                string[] strValue = new string[16];
                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    string deptName = dtData.Rows[iRow]["DEPT_NM"].ToString();
                    string mline = dtData.Rows[iRow]["MLINE"].ToString();
                    strValue[0] = dtData.Rows[iRow]["TOTAL_BG_COLOR"].ToString();
                    strValue[1] = dtData.Rows[iRow]["TOTAL_FORE_COLOR"].ToString();
                    strValue[2] = dtData.Rows[iRow]["BG_COLOR"].ToString();
                    strValue[3] = dtData.Rows[iRow]["FORE_COLOR"].ToString();
                    strValue[4] = deptName;
                    strValue[5] = dtData.Rows[iRow]["LINE_NAME"].ToString();
                    strValue[6] = dtData.Rows[iRow]["TO"].ToString();
                    strValue[7] = dtData.Rows[iRow]["PO"].ToString();
                    strValue[8] = dtData.Rows[iRow]["Rate"].ToString();
                    strValue[9] = ht[deptName].ToString();
                    strValue[10] = dtData.Rows[iRow]["RELIEF"].ToString();
                    strValue[11] = dtData.Rows[iRow]["BALANCE"].ToString();
                    strValue[12] = dtData.Rows[iRow]["PROC"].ToString();
                    strValue[13] = ht2[mline].ToString();
                    strValue[14] = dtData.Rows[iRow]["TO3"].ToString();
                    strValue[15] = dtData.Rows[iRow]["PO_ACTUAL"].ToString();

                    //string[] strValue2 =
                    //{
                    //    dtData.Rows[iRow]["TOTAL_BG_COLOR"].ToString(),
                    //    dtData.Rows[iRow]["TOTAL_FORE_COLOR"].ToString(),
                    //    dtData.Rows[iRow]["BG_COLOR"].ToString(),
                    //    dtData.Rows[iRow]["FORE_COLOR"].ToString(),
                    //    deptName,
                    //    dtData.Rows[iRow]["LINE_NAME"].ToString(),
                    //    dtData.Rows[iRow]["TO"].ToString(),
                    //    dtData.Rows[iRow]["PO"].ToString(),
                    //    dtData.Rows[iRow]["Rate"].ToString(),
                    //    ht[deptName].ToString()
                    //};

                    string rowspan = "", rowspan2 = "";
                    if (iRow == 0)
                    {
                        rowspan = "<td bgcolor='{0}' style='color:{1}' align='left' rowspan='{9}' >{4}</td>";
                        rowspan2 = "<td bgcolor='{0}' style='color:{1}' align='left' rowspan='{13}'>{5}</td>";
                    }
                    else
                    {
                        rowspan = deptName == dtData.Rows[iRow - 1]["DEPT_NM"].ToString()
                            ? ""
                            : "<td bgcolor='{0}' style='color:{1}' align='left' rowspan='{9}' >{4}</td>";

                        rowspan2 = mline == dtData.Rows[iRow - 1]["MLINE"].ToString()
                            ? ""
                            : "<td bgcolor='{0}' style='color:{1}' align='left' rowspan='{13}'>{5}</td>";
                    }

                    rowValue += string.Format(
                            "<tr>" +
                                rowspan +
                                rowspan2 +
                                // "<td bgcolor='{0}' style='color:{1}' align='left'>{5}</td>" +
                                "<td bgcolor='{0}' style='color:{1}' align='left'>{12}</td>" +
                                "<td bgcolor='{0}' style='color:{1}' align='right'>{6}</td>" +
                                "<td bgcolor='{0}' style='color:{1}' align='right'>{15}</td>" +
                                "<td bgcolor='{0}' style='color:{1}' align='right'>{10}</td>" +
                                "<td bgcolor='{0}' style='color:{1}' align='right'>{11}</td>" +
                                "<td bgcolor='{2}' style='color:{3}' align='right'>{8}</td>" +
                                "<td bgcolor='{0}' style='color:{1}' align='right'>{14}</td>" +
                                "<td bgcolor='{0}' style='color:{1}' align='right'>{7}</td>" +
                            "</tr>",
                            strValue);
                }

                string html = "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;' >" +
                                      "<b style='background-color:black; color:yellow' >Formular and staffing ratio color explanation</b><br>" +
                                      "&nbsp;&nbsp;&nbsp;Balance = PO + Relief – TO<br>" +
                                      "&nbsp;&nbsp;&nbsp;Staffing Ratio = (PO + Relief) / TO<br>" +
                                      "&nbsp;&nbsp;&nbsp;More than 105: orange<br>" +
                                      "&nbsp;&nbsp;&nbsp;102 ~ 105&nbsp;&nbsp;&nbsp;: yellow<br>" +
                                      "&nbsp;&nbsp;&nbsp;100 ~ 102&nbsp;&nbsp;&nbsp;: green<br>" +
                                      "&nbsp;&nbsp;&nbsp;98 ~ 100&nbsp;&nbsp;&nbsp;&nbsp;: yellow<br>" +
                                      "&nbsp;&nbsp;&nbsp;Less than 98&nbsp;: red" +
                                    "</p>" +
                                "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' >" +
                                  "<tr>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '150'>Plant</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>Line</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>Process</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>TO</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>PO Actual</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>Relief Actual</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>Balance</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>Staffing Ratio(%)</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>TO 3%</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width = '80'>PO Total</th>" +
                                  "</tr>" +
                                    rowValue +
                              "</table>";

                mailItem.HTMLBody = html;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailToPo: " + ex.ToString());
            }
        }

        public DataSet SEL_TO_PO_IE(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_TO_PO_IE";
                MyOraDB.ReDim_Parameter(4);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";
                MyOraDB.Parameter_Name[3] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("P_SEND_EMAIL_PROD: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_PROD_DATA: " + ex.ToString());
                return null;
            }
        }

        #endregion

        #region Email E-SCAN

        private void RunEScan(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_ESCAN_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null || dsData.Tables.Count == 0) return;

                DataTable dtLTF1 = dsData.Tables[0];
                DataTable dtLTNosN = dsData.Tables[1];
                DataTable dtTP = dsData.Tables[2];
                DataTable dtEmail = dsData.Tables[3];

                WriteLog("RunEScan Data: " + dtLTF1.Rows.Count.ToString() + " " + dtLTNosN.Rows.Count.ToString() + " " + dtTP.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                CreateMailEScan(dtLTF1, dtLTNosN, dtTP, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog("RunEScan: " + ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void CreateMailEScan(DataTable dtLTF1, DataTable dtLTNosN, DataTable dtTP, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Current Status of E-SCAN";

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email 
                if (!app.Session.CurrentUser.AddressEntry.Address.Contains("IT.NGOC"))
                {
                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }

                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                mailItem.Body = "This is the message.";

                string HtmlLongThanhF1 = CreateTableHtml(dtLTF1);
                string HtmlLongThanhNosN = CreateTableHtml(dtLTNosN);
                string HtmlTanPhu = CreateTableHtml(dtTP);

                mailItem.HTMLBody = "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;' >" +
                                      "Control limit = result &plusmn;10%<br>" +
                                      "Within control limit is green<br>" +
                                      "Lower or higher control limit is red" +
                                    "</p>" +
                                    "<b style='font-family:Times New Roman; font-size:22px; font-style:Italic'>Long Thanh - F1</b>" +
                                    "<br>" +
                                    HtmlLongThanhF1 +
                                    "<br>" +
                                     "<b style='font-family:Times New Roman; font-size:22px; font-style:Italic'>Long Thanh - N</b>" +
                                    HtmlLongThanhNosN +
                                    "<br>" +
                                    "<b style='font-family:Times New Roman; font-size:22px; font-style:Italic;'>Tan Phu</b>" +
                                    "<br>" +
                                    HtmlTanPhu;


                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailEScan " + ex.ToString());
            }
        }


        private string CreateTableHtml(DataTable argDt)
        {
            string rowValue = "";
            int iRowCount = argDt.Rows.Count;
            string[] value = new string[5];

            for (int iRow = 0; iRow < iRowCount; iRow++)
            {
                string RowCurrdate = argDt.Rows[iRow]["DATE"].ToString();

                value[0] = RowCurrdate;
                value[1] = argDt.Rows[iRow]["BG_COLOR"].ToString();
                value[2] = argDt.Rows[iRow]["FORE_COLOR"].ToString();
                value[3] = argDt.Rows[iRow]["SCAN_QTY"].ToString();
                value[4] = argDt.Rows[iRow]["E_SCAN_QTY"].ToString();

                if (iRow == 0 || RowCurrdate != argDt.Rows[iRow - 1]["DATE"].ToString())
                {
                    rowValue += string.Format("<tr><td align='center' width='80'>{0}</td>", value);
                }

                rowValue += string.Format(
                            "<td bgcolor='{1}' style='color:{2}' align='right' width='50'>{3}</td>" +
                            "<td bgcolor='{1}' style='color:{2}' align='right' width='50'>{4}</td>",
                            value);
                if (iRow + 1 >= iRowCount || RowCurrdate != argDt.Rows[iRow + 1]["DATE"].ToString())
                {
                    rowValue += "</tr>";
                }
            }

            string strHeader = GetHeader(argDt);

            return "<table style='font-family:Calibri; font-size:14px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0'>" +
                        strHeader +
                        rowValue +
                    "</table>";
        }



        private string GetHeader(DataTable argTable)
        {
            string[] selectedColumns = new[] { "LINE", "ORD" };
            DataTable dtHeader = new DataView(argTable).ToTable(true, selectedColumns);

            dtHeader = dtHeader.Select("", "ORD").Distinct().CopyToDataTable();

            string strHeaderRow1 = "", strHeaderRow2 = "";
            foreach (DataRow row in dtHeader.Rows)
            {
                strHeaderRow1 += " <th colspan = '2' align='center'>" + row["LINE"].ToString() + "</th>";
                strHeaderRow2 += " <th bgcolor='#ff9900' style='color:#ffffff' align='center'  >Scan</th>" +
                                 " <th bgcolor='#366cc9' style='color:#ffffff' align='center'  >E-Scan</th>";
            }

            string strHtml = "<tr bgcolor='#ffe5cc'>" +
                                " <th rowspan = '2' align='center' >Date</th>" +
                                 strHeaderRow1 +
                             "</tr>" +
                             "<tr>" +
                                 strHeaderRow2 +
                             "</tr>";
            return strHtml;
        }

        public DataSet SEL_ESCAN_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_ESCAN_TEST";

                MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
                // MyOraDB.ShowErr = true;
                MyOraDB.ReDim_Parameter(7);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_LOC";
                MyOraDB.Parameter_Name[2] = "V_P_DATE";
                MyOraDB.Parameter_Name[3] = "CV_1";
                MyOraDB.Parameter_Name[4] = "CV_2";
                MyOraDB.Parameter_Name[5] = "CV_3";
                MyOraDB.Parameter_Name[6] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = V_P_DATE;
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";
                MyOraDB.Parameter_Values[6] = "";


                MyOraDB.Add_Select_Parameter(true);
                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("P_SEND_EMAIL_ESCAN: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region Email Production

        private void RunProduction(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_PROD_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;
                WriteLog($"RunProduction({argType}): BEGIN ");
                DataTable dtDate = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog("  " + dtDate.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                CreateMailProduction(dtDate, dtData, dtEmail);
                WriteLog($"RunProduction({argType}): END ");
            }
            catch (Exception ex)
            {
                WriteLog($"  RunProduction({argType}) " + ex.ToString());
            }
            finally
            {

                _isRun2 = false;
            }

        }

        private void CreateMailProduction(DataTable dtDate, DataTable dtData, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Productivity achievement ratio at this time";

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email 
                if (!app.Session.CurrentUser.AddressEntry.Address.Contains("IT.NGOC"))
                {

                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }

                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                mailItem.Body = "This is the message.";

                string rowValue = "";

                string strRowSpan = "";

                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    strRowSpan = dtData.Rows[iRow]["CNT"].ToString();
                    if (iRow == 0)
                    {
                        rowValue += "<tr>" +
                                       "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + " </td>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["MLINE"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D6_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D6_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D6"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D5_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D5_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D5"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D4_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D4_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D4"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D3_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D3_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D3"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D2_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D2_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D2"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D1_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D1_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D1"].ToString() + "</td>" +
                                       "<td align='right' >" + dtData.Rows[iRow]["TARGET"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["RPLAN"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["ACT"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["TODAY_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["TODAY_FORE_COLOR"].ToString() + "' align='right' >" + dtData.Rows[iRow]["RATIO"].ToString() + " </td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["REASON_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["REASON_FORE_COLOR"].ToString() + "' align='left' >" + dtData.Rows[iRow]["REASON"].ToString() + " </td>" +
                                  "</tr>";
                    }
                    else
                    {
                        if (dtData.Rows[iRow]["PLANT"].ToString() == dtData.Rows[iRow - 1]["PLANT"].ToString())
                        {
                            rowValue += "<tr>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["MLINE"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D6_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D6_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D6"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D5_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D5_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D5"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D4_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D4_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D4"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D3_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D3_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D3"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D2_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D2_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D2"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D1_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D1_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D1"].ToString() + "</td>" +
                                       "<td align='right' >" + dtData.Rows[iRow]["TARGET"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["RPLAN"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["ACT"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["TODAY_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["TODAY_FORE_COLOR"].ToString() + "' align='right' >" + dtData.Rows[iRow]["RATIO"].ToString() + " </td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["REASON_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["REASON_FORE_COLOR"].ToString() + "' align='left' >" + dtData.Rows[iRow]["REASON"].ToString() + " </td>" +
                                  "</tr>";
                        }
                        else
                        {
                            rowValue += "<tr>" +
                                         "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + " </td>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["MLINE"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D6_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D6_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D6"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D5_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D5_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D5"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D4_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D4_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D4"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D3_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D3_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D3"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D2_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D2_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D2"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D1_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D1_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D1"].ToString() + "</td>" +
                                       "<td align='right' >" + dtData.Rows[iRow]["TARGET"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["RPLAN"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["ACT"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["TODAY_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["TODAY_FORE_COLOR"].ToString() + "' align='right' >" + dtData.Rows[iRow]["RATIO"].ToString() + " </td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["REASON_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["REASON_FORE_COLOR"].ToString() + "' align='left' >" + dtData.Rows[iRow]["REASON"].ToString() + " </td>" +
                                  "</tr>";
                        }

                    }
                }

                string strDate = "";
                foreach (DataRow row in dtDate.Rows)
                {
                    strDate += "<th bgcolor = '#ff9900' style = 'color:#ffffff' align = 'center' width = '70' >" + row["YMD"].ToString() + " </th >";
                }

                string html = "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1400'>" +
                               "<tr bgcolor='#ffe5cc'>" +
                                  " <th rowspan = '2' align='center' width='70'>Plant</th>" +
                                  " <th rowspan = '2' align='center' width='70'>Mini Line</th>" +
                                  " <th bgcolor = '#ff9900' style = 'color:#ffffff' colspan = '6' align='center'>Full time on previous day performance</th>" +
                                  " <th bgcolor = '#366cc9' style = 'color:#ffffff' colspan = '4' align='center'> Before lunch on today performance</th>" + " <th rowspan = '2' align='center' bgcolor = '#000000' style = 'color:#ffffff' width='200'>Reason of underproduction</th>" +

                               "</tr>" +
                               "<tr>" +
                                  strDate +
                                  "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='100'>Daily Plan</th>" +
                                  "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='100'>Real Plan</th>" +
                                  "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='100'>Actual</th>" +
                                  "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='100'>Ratio(%)</th>" +
                               "</tr>" +

                                 rowValue +
                           "</table>";

                //string text = "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic; color:#0000ff' >" +
                //                    "SPR(Sequence Production Ratio) = How many follow passcard scan sequence of ratio" +
                //               "</p>";




                mailItem.HTMLBody = html;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("  CreateMailProduction: " + ex.ToString());
            }
        }

        public DataSet SEL_PROD_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_PROD";
                MyOraDB.ReDim_Parameter(5);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";
                MyOraDB.Parameter_Name[3] = "CV_2";
                MyOraDB.Parameter_Name[4] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";


                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("P_SEND_EMAIL_PROD: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_PROD_DATA: " + ex.ToString());
                return null;
            }
        }

        #endregion

        #region Email ANDON

        private void RunAndon(string argType)
        {
            try
            {
                DataSet dsData = SEL_ANDON_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;
                WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunAndon({argType}): BEGIN");
                DataTable dtData = dsData.Tables[0];
                DataTable dtEmail = dsData.Tables[1];
                WriteLog($"  dtData:{dtData.Rows.Count} dtEmail: {dtEmail.Rows.Count}");
                CreateMailAndon(dtData, dtEmail);
                WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunAndon({argType}): END");
            }
            catch (Exception ex)
            {
                WriteLog($"  RunAndon({argType}): {ex}");
                throw;
            }
            finally
            {

            }

        }

        private void CreateMailAndon(DataTable dtData, DataTable dtEmail)
        {
            try
            {


                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Andon information of yesterday";
                string str = app.Name;
                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (!app.Session.CurrentUser.AddressEntry.Address.Contains("IT.NGOC"))
                {
                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }

                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                mailItem.Body = "This is the message.";

                string rowValue = "";

                string strRowSpan = "";


                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    strRowSpan = dtData.Rows[iRow]["MLINE_CNT"].ToString();
                    if (iRow == 0)
                    {
                        rowValue += "<tr>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["RANKING"].ToString() + " </td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" +
                                            // "bgcolor='" + dtData.Rows[iRow]["BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_LINE"].ToString() +
                                        "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR2"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR2"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_LINE_AVG"].ToString() +
                                        "</td>" +

                                        //  "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["DOWNTIME_LINE_AVG"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["CALLING_TIMES_LINE"].ToString() + "</td>" +

                                        "<td rowspan='" + strRowSpan + "' align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR3"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR3"].ToString() + "'>" +
                                            dtData.Rows[iRow]["AVERAGE_ELAPSE_LINE"].ToString() +
                                        "</td>" +

                                        // "<td rowspan='" + strRowSpan + "' align ='center' >" + dtData.Rows[iRow]["AVERAGE_ELAPSE_LINE"].ToString() + "</td>" +
                                        "<td align='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +

                                        "<td  align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR4"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR4"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() +
                                        "</td>" +

                                        // "<td align='center'>" + dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() + "</td>" +
                                        "<td align='center'>" + dtData.Rows[iRow]["CALLING_TIMES_MLINE"].ToString() + "</td>" +

                                        "<td align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR5"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR5"].ToString() + "'>" +
                                            dtData.Rows[iRow]["AVERAGE_ELAPSE_MLINE"].ToString() +
                                        "</td>" +

                                        "<td rowspan='" + strRowSpan + "' align ='center'>" +

                                            dtData.Rows[iRow]["MACHINE_CNT_LINE"].ToString() +
                                        "</td>" +

                                   //"<td rowspan='" + strRowSpan + "' align ='center'" +
                                   //     "bgcolor='" + dtData.Rows[iRow]["BG_COLOR6"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR6"].ToString() + "'>" +
                                   //    dtData.Rows[iRow]["DOWN_TIME"].ToString() +
                                   //"</td>" +

                                   //   "<td align='center'>" + dtData.Rows[iRow]["AVERAGE_ELAPSE_MLINE"].ToString() + "</td>" +

                                   "</tr>";
                    }
                    // bgcolor='" + dtData.Rows[iRow]["BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR"].ToString() + "'
                    if (iRow > 0)
                    {
                        if (dtData.Rows[iRow]["PLANT"].ToString() == dtData.Rows[iRow - 1]["PLANT"].ToString())
                        {
                            rowValue += "<tr>" +
                                        //"<td align='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + "</td>" +
                                        //"<td align='center'>" + dtData.Rows[iRow]["CALLING_TIMES_LINE"].ToString() + "</td>" +
                                        //"<td align='right' >" + dtData.Rows[iRow]["DOWNTIME_LINE"].ToString() + "</td>" +
                                        //"<td align='right'>" + dtData.Rows[iRow]["AVERAGE_ELAPSE_LINE"].ToString() + "</td>" +
                                        "<td align='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +
                                         "<td  align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR4"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR4"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() +
                                        "</td>" +

                                        // "<td align='center'>" + dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() + "</td>" +
                                        "<td align='center'>" + dtData.Rows[iRow]["CALLING_TIMES_MLINE"].ToString() + "</td>" +

                                        "<td align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR5"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR5"].ToString() + "'>" +
                                            dtData.Rows[iRow]["AVERAGE_ELAPSE_MLINE"].ToString() +
                                        "</td>" +
                                   //     "<td bgcolor='" + row["BG_COLOR"].ToString() + "' style='color:" + row["FORE_COLOR"].ToString() + "' align='right' >" + row["RATIO"].ToString() + " </td>" +
                                   "</tr>";
                        }
                        else
                        {
                            rowValue += "<tr>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["RANKING"].ToString() + " </td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" +
                                            //  "bgcolor='" + dtData.Rows[iRow]["BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_LINE"].ToString() +
                                        "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR2"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR2"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_LINE_AVG"].ToString() +
                                        "</td>" +

                                        //  "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["DOWNTIME_LINE_AVG"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["CALLING_TIMES_LINE"].ToString() + "</td>" +

                                        "<td rowspan='" + strRowSpan + "' align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR3"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR3"].ToString() + "'>" +
                                            dtData.Rows[iRow]["AVERAGE_ELAPSE_LINE"].ToString() +
                                        "</td>" +

                                        // "<td rowspan='" + strRowSpan + "' align ='center' >" + dtData.Rows[iRow]["AVERAGE_ELAPSE_LINE"].ToString() + "</td>" +
                                        "<td align='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +

                                        "<td  align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR4"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR4"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() +
                                        "</td>" +

                                        // "<td align='center'>" + dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() + "</td>" +
                                        "<td align='center'>" + dtData.Rows[iRow]["CALLING_TIMES_MLINE"].ToString() + "</td>" +

                                        "<td align ='center'" +
                                             "bgcolor='" + dtData.Rows[iRow]["BG_COLOR5"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR5"].ToString() + "'>" +
                                            dtData.Rows[iRow]["AVERAGE_ELAPSE_MLINE"].ToString() +
                                        "</td>" +

                                        "<td rowspan='" + strRowSpan + "' align ='center'>" +

                                            dtData.Rows[iRow]["MACHINE_CNT_LINE"].ToString() +
                                        "</td>" +

                                   //"<td rowspan='" + strRowSpan + "' align ='center'" +
                                   //     "bgcolor='" + dtData.Rows[iRow]["BG_COLOR6"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR6"].ToString() + "'>" +
                                   //    dtData.Rows[iRow]["DOWN_TIME"].ToString() +
                                   //"</td>" +

                                   "</tr>"; ;

                        }
                    }
                }

                string text = "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;' >" +
                                    "Total Downtime per Line & Downtime by line = under 10 minutes is <b style='color:green'>green </b> " +
                                    "and from 10 min to 29:59 is <b style='background-color:black; color:yellow'>yellow</b> " +
                                    "and from 30 min to 59:59 min is <b style='color:orange'>orange</b> " +
                                    "and then more than 1 hour is <b style='color:red'>red</b>" +
                               "</p>" +
                              "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;'>" +
                                    "Total average measure & Downtime average by line = under 2 minutes is <b style='color:green'>green </b> " +
                                    "and from 2 min to 4:59 is <b style='background-color:black; color:yellow'>yellow</b> " +
                                    "and from 5 min to 09:59 is <b style='color:orange'>orange</b> " +
                                    "and then more than 10 min is <b style='color:red'>red</b>" +
                              "</p>"
                              ;

                string html = text +
                            "<br>" +
                            "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1000px'>" +
                                  "<tr bgcolor='#366cc9' style='color:#ffffff'>" +
                                     "<th style='color:#ffffff' align='center' width='100'>Ranking</th>" +
                                     "<th style='color:#ffffff' align='center' width='100'>Plant</th>" +
                                     "<th style='color:#ffffff' align='center' width='200'>Total Downtime</th>" +
                                     "<th style='color:#ffffff' align='center' width='200'>Total Downtime per Line</th>" +
                                     "<th style='color:#ffffff' align='center' width='200'>Total Calling Times</th>" +
                                     "<th style='color:#ffffff' align='center' width='200'>Total Average Measure</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width='100'>Line</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width='200'>Downtime by Line</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width='200'>Calling Times by Line</th>" +
                                     "<th bgcolor='#00ced1' style='color:#ffffff' align='center' width='200'>Downtime Average by Line</th>" +
                                     "<th bgcolor='#8b2cb0' style='color:#ffffff' align='center' width='200'>Machine Total</th>" +
                                  //"<th bgcolor='#8b2cb0' style='color:#ffffff' align='center' width='200'>Machine D/T(Min)</th>" +
                                  "</tr>" +
                                    rowValue +
                              "</table>";

                mailItem.HTMLBody = html;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailAndon: " + ex.ToString());
            }
        }


        public DataSet SEL_ANDON_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_ANDON_V2";

                MyOraDB.ReDim_Parameter(4);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";
                MyOraDB.Parameter_Name[3] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";


                MyOraDB.Add_Select_Parameter(true);
                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null) return null;
                return ds_ret;
            }
            catch
            {
                return null;
            }
        }

        #endregion        

        #region Email Bottom Inventory

        private void CreateMailBottomInventory()
        {
            try
            {
                //Outlook.MailItem mailItem = (Outlook.MailItem)
                // this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttach = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Chart.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPicGrid1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid1.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPicGrid2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid2.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPicGrid3 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid3.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPicGrid4 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid4.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = "Bottom Inventory Set Analysis";


                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (!app.Session.CurrentUser.AddressEntry.Address.Contains("IT.NGOC"))
                {
                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                //Get List Send email Test
                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }

                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                mailItem.Body = "This is the message.";
                string imgChart = "imgChart", imgGrid1 = "imgGrid1", imgGrid2 = "imgGrid2", imgGrid3 = "imgGrid3", imgGrid4 = "imgGrid4";
                oAttach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgChart);
                oAttachPicGrid1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgGrid1);
                oAttachPicGrid2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgGrid2);
                oAttachPicGrid3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgGrid3);
                oAttachPicGrid4.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgGrid4);
                mailItem.HTMLBody = String.Format(
                    "<body>" +
                          "<img src=\"cid:{0}\">" +
                        "<br><img src=\"cid:{1}\">" +
                        "<br><img src=\"cid:{2}\">" +
                        "<br><img src=\"cid:{3}\">" +
                        "<br><img src=\"cid:{4}\">" +
                    "</body>",
                    imgChart, imgGrid1, imgGrid2, imgGrid3, imgGrid4);
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }


        }

        private void Run(string argType)
        {
            if (_isRun) return;
            if (!LoadData(argType))
            {
                // MessageBox.Show("Do not Send!");

                _isRun = false;
                return;
            }

            try
            {
                _isRun = true;
                CaptureControl(panel1, "Chart");
                CaptureControl(grdBase1, "Grid1");
                CaptureControl(grdBase2, "Grid2");
                CaptureControl(grdBase3, "Grid3");
                CaptureControl(grdBase4, "Grid4");
                CreateMailBottomInventory();
            }
            catch { lblStatus.Text = DateTime.Now.ToString() + "Do not Send!"; }
            finally { _isRun = false; }
        }

        private bool LoadData(string typeSearch)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                DataSet dsReturn = SEL_LOAD_DATA(typeSearch, DateTime.Now.ToString("yyyyMMdd"), "", "N");
                if (dsReturn == null) return false;
                DataTable dtChart = dsReturn.Tables[0];
                DataTable dtGrid = dsReturn.Tables[1];
                dtEmail = dsReturn.Tables[2];
                LoadDataChart(dtChart);

                // Grid 
                DataTable dtSource = new DataTable();
                if (buildHeader_detail(dtSource, dtGrid))
                {
                    if (bindingDataSource_detail(dtSource, dtGrid))
                    {
                        grdBase1.DataSource = dtSource.Select("Factory in ('F1','F2')").CopyToDataTable();
                        grdBase2.DataSource = dtSource.Select("Factory in ('F3')").CopyToDataTable();
                        grdBase3.DataSource = dtSource.Select("Factory in ('F4')").CopyToDataTable();
                        grdBase4.DataSource = dtSource.Select("Factory in ('F5','VJ2')").CopyToDataTable();

                        formatGrid2(gvwBase1);
                        formatGrid2(gvwBase2);
                        formatGrid2(gvwBase3);
                        formatGrid2(gvwBase4);
                    }
                }

                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private bool buildHeader_detail(DataTable dtSource, DataTable dt)
        {
            try
            {
                int.TryParse(dt.Rows[0]["START_COLUMN"].ToString(), out _start_column);
                for (int i = 0; i < _start_column - 1; i++)
                {
                    if (i == _start_column - 2 || i == _start_column - 3)
                        dtSource.Columns.Add(dt.Columns[i].Caption, typeof(decimal));
                    else
                        dtSource.Columns.Add(dt.Columns[i].Caption);
                }
                dtSource.Columns.Add("Total", typeof(decimal));
                DataRow[] row = dt.Select("", "SIZE_NUM ASC");
                double min = 1, max = 22;
                double.TryParse(row[0]["SIZE_NUM"].ToString(), out min);
                double.TryParse(row[row.Length - 1]["SIZE_NUM"].ToString(), out max);
                for (double i = min; i <= max; i = i + 0.5)
                {
                    dtSource.Columns.Add(i.ToString().Replace(".5", "T"), typeof(decimal));
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }
        }

        private bool bindingDataSource_detail(DataTable dtSource, DataTable dtTemp)
        {
            int cnt = 0;
            try
            {
                int[] rowtotal = new int[dtSource.Columns.Count];
                string distinct_row = "";
                int temp1, temp2;

                for (int i = 0; i < dtTemp.Rows.Count; i++)
                {
                    cnt = i;
                    if (distinct_row != dtTemp.Rows[i]["DISTINCTROW"].ToString())
                    {
                        dtSource.Rows.Add();
                    }
                    distinct_row = dtTemp.Rows[i]["DISTINCTROW"].ToString();
                    for (int j = 0; j < _start_column - 1; j++)
                    {
                        dtSource.Rows[dtSource.Rows.Count - 1][j] = dtTemp.Rows[i][j];
                    }

                    int.TryParse(dtTemp.Rows[i]["QTY"].ToString(), out temp1);
                    int.TryParse(dtSource.Rows[dtSource.Rows.Count - 1][dtTemp.Rows[i]["CS_SIZE"].ToString()].ToString(), out temp2);
                    dtSource.Rows[dtSource.Rows.Count - 1][dtTemp.Rows[i]["CS_SIZE"].ToString()] = Convert.ToDecimal(temp1 + temp2);
                    int.TryParse(dtSource.Rows[dtSource.Rows.Count - 1][_start_column - 1].ToString(), out temp2);
                    dtSource.Rows[dtSource.Rows.Count - 1][_start_column - 1] = (temp1 + temp2).ToString();
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }
        }

        #region Set Data For Chart & Grid
        private void LoadDataChart(DataTable argDtChart)
        {
            chart2.DataSource = argDtChart;
            chart2.Series[2].ArgumentDataMember = "LINE_NM";
            chart2.Series[2].ValueDataMembers.AddRange(new string[] { "BT_QTY" });
            chart2.Series[1].ArgumentDataMember = "LINE_NM";
            chart2.Series[1].ValueDataMembers.AddRange(new string[] { "FS_QTY" });
            chart2.Series[0].ArgumentDataMember = "LINE_NM";
            chart2.Series[0].ValueDataMembers.AddRange(new string[] { "FN_QTY" });
            //chart2.Series[3].ArgumentDataMember = "LINE_NM";
            //chart2.Series[3].ValueDataMembers.AddRange(new string[] { "TAR_QTY" });
            chart2.Series[3].ArgumentDataMember = "LINE_NM";
            chart2.Series[3].ValueDataMembers.AddRange(new string[] { "TAR_QTY3" });
            chart2.Series[4].ArgumentDataMember = "LINE_NM";
            chart2.Series[4].ValueDataMembers.AddRange(new string[] { "TAR_QTY2" });
            chart2.Series[5].ArgumentDataMember = "LINE_NM";
            chart2.Series[5].ValueDataMembers.AddRange(new string[] { "TAR_QTY" });

            ((DevExpress.XtraCharts.XYDiagram)chart2.Diagram).AxisX.Label.Staggered = false;
        }

        private void formatGrid2(JPlatform.Client.Controls6.GridViewEx gvwBase)
        {
            try
            {
                gvwBase.BeginUpdate();
                //gvwBase2.RowHeight = 45;
                //gvwBase2.Appearance.Row.Font = new System.Drawing.Font("Calibri", 12F);
                //gvwBase2.Appearance.HeaderPanel.Font = new System.Drawing.Font("Calibri", 12F);

                gvwBase.RowHeight = -1;
                gvwBase.Appearance.Row.Font = new System.Drawing.Font("Calibri", 11F);
                gvwBase.Appearance.HeaderPanel.Font = new System.Drawing.Font("Calibri", 11F);

                for (int i = 0; i < gvwBase1.Columns.Count; i++)
                {
                    gvwBase.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwBase.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwBase.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwBase.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                    gvwBase.Columns[i].OptionsColumn.ReadOnly = true;
                    gvwBase.Columns[i].OptionsColumn.AllowEdit = false;
                    if (i > 5)
                    {
                        gvwBase.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                        gvwBase.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                        gvwBase.Columns[i].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                        gvwBase.Columns[i].DisplayFormat.FormatString = i <= 8 ? "#,0" : "#,#";
                        gvwBase.Columns[i].Width = 40;
                        if (gvwBase.Columns[i].FieldName.Equals("Total"))
                            gvwBase.Columns[i].Visible = false;
                        if (i <= 7)
                        {
                            gvwBase.Columns[i].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                            gvwBase.Columns[i].Width = 65;
                        }
                    }
                    else
                    {
                        gvwBase.Columns[i].OptionsColumn.AllowMerge = i <= 4 ? DevExpress.Utils.DefaultBoolean.True : DevExpress.Utils.DefaultBoolean.False;
                        gvwBase.Columns[i].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                        gvwBase.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    }
                }
                gvwBase.Columns[0].Visible = false;
                gvwBase.Columns[1].Visible = false;
                //gvwBase2.Columns["Target"].Visible = false;

                gvwBase.Columns[2].Width = 90;
                gvwBase.Columns[3].Width = 65;
                gvwBase.Columns[4].Width = 65;
                gvwBase.Columns[5].Width = 180;

                //gvwBase2.Columns[2].Width = 100;
                //gvwBase2.Columns[3].Width = 100;
                //gvwBase2.Columns[4].Width = 150;
                //gvwBase2.Columns[5].Width = 300;
                gvwBase.Columns[5].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                //gvwBase2.Columns[6].OptionsColumn.AllowMerge = DefaultBoolean.True;
                //gvwBase.Columns[5].Width = 95;
                //gvwBase.Columns[7].Width = 170;
                gvwBase.EndUpdate();
            }
            catch { }
        }

        #endregion

        private void CaptureControl(Control control, string nameImg)
        {
            //  MemoryStream ms = new MemoryStream();
            string Path = Application.StartupPath + @"\Capture\";
            Bitmap bmp = new Bitmap(control.Width, control.Height);
            if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);
            control.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, control.Width, control.Height));
            bmp.Save(Path + nameImg + @".png", System.Drawing.Imaging.ImageFormat.Png); //you could ave in BPM, PNG  etc format.
            //byte[] Pic_arr = new byte[ms.Length];
            //ms.Position = 0;
            //ms.Read(Pic_arr, 0, Pic_arr.Length);
            //ms.Close();
        }

        public DataSet SEL_LOAD_DATA(string V_P_WORK_TYPE, string V_P_DATE, string V_P_COMP, string V_P_SET_YN)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
                string process_name = "P_SEND_EMAIL_TO_GM";

                MyOraDB.ReDim_Parameter(7);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_WORK_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "V_P_COMP";
                MyOraDB.Parameter_Name[3] = "V_P_SET_YN";
                MyOraDB.Parameter_Name[4] = "CV_1";
                MyOraDB.Parameter_Name[5] = "CV_2";
                MyOraDB.Parameter_Name[6] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_WORK_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = V_P_COMP;
                MyOraDB.Parameter_Values[3] = V_P_SET_YN;
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";
                MyOraDB.Parameter_Values[6] = "";


                MyOraDB.Add_Select_Parameter(true);
                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null) return null;
                return ds_ret;
            }
            catch
            {
                return null;
            }
        }



        private void gvwBase2_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridViewEx ex = sender as GridViewEx;

            if (e.Column.ColumnHandle > 4)
            {
                if (ex.GetRowCellValue(e.RowHandle, "Div").ToString().ToUpper().Equals("BOTTOM SET"))
                {
                    e.Appearance.BackColor = Color.FromArgb(242, 242, 242);
                }
                else if (ex.GetRowCellValue(e.RowHandle, "Div").ToString().ToUpper().Equals("STOCKFIT INCOMING SET"))
                {
                    e.Appearance.BackColor = Color.White;
                }
            }

            if (e.Column.ColumnHandle > 5 && ex.GetRowCellValue(e.RowHandle, "Plant").ToString() != "Total")
            {

                if (e.CellValue == null || e.CellValue.ToString() == "") return;
                //if (Convert.ToDouble(e.CellValue.ToString().Replace(",", "")) < 0)
                //{
                //    e.Appearance.BackColor = Color.Red;
                //    e.Appearance.ForeColor = Color.White;
                //}
                if (e.Column.FieldName.Contains("Total Inv"))
                {
                    if (Convert.ToDouble(e.CellValue.ToString().Replace(",", "")) < 0)
                    {
                        e.Appearance.BackColor = Color.Red;
                        e.Appearance.ForeColor = Color.White;
                    }
                    else if (Convert.ToDouble(e.CellValue.ToString().Replace(",", "")) >= Convert.ToDouble(ex.GetRowCellValue(e.RowHandle, "Target").ToString().Replace(",", "")))
                    {
                        e.Appearance.BackColor = Color.Green;
                        e.Appearance.ForeColor = Color.White;
                    }
                    else if (Convert.ToDouble(e.CellValue.ToString().Replace(",", "")) < Convert.ToDouble(ex.GetRowCellValue(e.RowHandle, "Target").ToString().Replace(",", "")))
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
            }
            if (e.Column.ColumnHandle > 1)
            {
                if (ex.GetRowCellValue(e.RowHandle, "Plant").ToString() == "Total")
                {
                    e.Appearance.BackColor = Color.LightCyan;
                    e.Appearance.ForeColor = Color.Coral;
                }
            }
        }

        #endregion

        #region Email Cutting

        private void RunCutting(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_CUTTING_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;

                DataTable dtHeader = dsData.Tables[1];
                DataTable dtData = dsData.Tables[0];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                string html = GetHtmlBodyCutting(dtHeader, dtData);

                CreateMail("Cutting current situation in front of input stitching line", html, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }

        }

        private string GetHtmlBodyCutting(DataTable dtHeader, DataTable dtData)
        {
            try
            {
                var query = from row in dtData.AsEnumerable()
                            group row by row.Field<string>("LINE_NAME") into dept
                            orderby dept.Key
                            select new
                            {
                                Name = dept.Key,
                                cntLine = dept.Count()
                            };

                System.Collections.Hashtable htDept = new System.Collections.Hashtable();
                System.Collections.Hashtable htLine = new System.Collections.Hashtable();

                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    string HtKey = dtData.Rows[iRow]["LINE_NAME"].ToString() + dtData.Rows[iRow]["MLINE_CD"].ToString();
                    if (htLine.ContainsKey(HtKey))
                    {
                        int.TryParse(htLine[HtKey].ToString(), out int CurrentValue);
                        htLine[HtKey] = CurrentValue + 1;
                    }
                    else
                    {
                        htLine.Add(HtKey, 1);
                    }

                }

                foreach (var row in query)
                {
                    htDept.Add(row.Name, row.cntLine);
                }

                string strColorExplain = dtHeader.Rows[0]["TEXT1"].ToString();
                string TableHeader = "";
                //Header
                //string[] ColumHead = new string[dtHeader.Rows.Count];

                TableHeader = "<tr> " +
                                 // "<td bgcolor = '#00ced1' style = 'color:#ffffff' rowspan = '2' align = 'center' width = '50'>Rank</td>" +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' rowspan = '2' align = 'center' width = '80'>Plant</td>" +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' rowspan = '2' align = 'center' width = '50'>Line</td>" +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' rowspan = '2' align = 'center' width = '80'>Stitching Line</td>" +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' rowspan = '2' align = 'center' width = '120'>Stitching Input</td > " +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' colspan = '2' align = 'center' >UPC<br>(D-D +2H)</td > " +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' colspan = '2' align = 'center' >UPA<br>(D-D +6H)</td > " +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' colspan = '2' align = 'center' >UPA2<br>(D-D +10H)</td > " +
                                 "<td bgcolor = '#00ced1' style = 'color:#ffffff' colspan = '2' align = 'center' >UPO<br>(D-D +16H)</td > " +





                              "</tr> " +
                              "<tr> " +
                                 "<td bgcolor = '#9966ff' style = 'color:#ffffff' align = 'center' width = '120' >Fast</td > " +
                                 "<td bgcolor = '#366cc9' style = 'color:#ffffff' align = 'center' width = '120' >Slow</td > " +
                                 "<td bgcolor = '#9966ff' style = 'color:#ffffff' align = 'center' width = '120' >Fast</td > " +
                                 "<td bgcolor = '#366cc9' style = 'color:#ffffff' align = 'center' width = '120' >Slow</td > " +
                                 "<td bgcolor = '#9966ff' style = 'color:#ffffff' align = 'center' width = '120' >Fast</td > " +
                                 "<td bgcolor = '#366cc9' style = 'color:#ffffff' align = 'center' width = '120' >Slow</td > " +
                                 "<td bgcolor = '#9966ff' style = 'color:#ffffff' align = 'center' width = '120' >Fast</td > " +
                                 "<td bgcolor = '#366cc9' style = 'color:#ffffff' align = 'center' width = '120' >Slow</td > " +
                              //"<td bgcolor = '#ff3399' style = 'color:#ffffff' align = 'center' width = '100' >Assembly Date</td > " +
                              //"<td bgcolor = '#9966ff' style = 'color:#ffffff' align = 'center' width = '100' >Input Date</td > " +
                              //"<td bgcolor = '#366cc9' style = 'color:#ffffff' align = 'center' width = '100' >Input Time</td > " +
                              //"<td bgcolor = '#9966ff' style = 'color:#ffffff' align = 'center' width = '100' >Date</td > " +
                              //"<td bgcolor = '#366cc9' style = 'color:#ffffff' align = 'center' width = '100' >Time</td > " +
                              "</tr> ";


                //Row
                string TableRow = "", rowspan = "", rowspan2 = "", rowspan3 = "";
                int iRanking = 0;
                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    string deptName = dtData.Rows[iRow]["LINE_NAME"].ToString();
                    string mline = dtData.Rows[iRow]["MLINE_CD"].ToString();
                    string cLine = dtData.Rows[iRow]["UPS_LINE"].ToString();
                    //   string Component = dtData.Rows[iRow]["PART_NM"].ToString();

                    //  string AssYmdQty = dtData.Rows[iRow]["ASY_YMD"].ToString();
                    string UpsYmdQty = dtData.Rows[iRow]["UPS_YMD"].ToString();
                    // string UpsHmsQty = dtData.Rows[iRow]["UPS_HMS"].ToString();
                    string UpsBColor = ColorNull(dtData.Rows[iRow]["UPS_BCOLOR"].ToString());
                    string UpsFColor = ColorNull(dtData.Rows[iRow]["UPS_FCOLOR"].ToString());

                    //string UpsSlowQty = dtData.Rows[iRow]["UPS_SLOW"].ToString();
                    //string UpsFastQty = dtData.Rows[iRow]["UPS_FAST"].ToString();
                    //string UpsSlowBColor = ColorNull(dtData.Rows[iRow]["UPS_BCOLOR_S"].ToString());
                    //string UpsSlowFColor = ColorNull(dtData.Rows[iRow]["UPS_FCOLOR_S"].ToString());
                    //string UpsFastBColor = ColorNull(dtData.Rows[iRow]["UPS_BCOLOR_F"].ToString());
                    //string UpsFastFColor = ColorNull(dtData.Rows[iRow]["UPS_FCOLOR_F"].ToString());

                    string UpcSlowQty = dtData.Rows[iRow]["UPC_SLOW"].ToString();
                    string UpcFastQty = dtData.Rows[iRow]["UPC_FAST"].ToString();
                    string UpcSlowBColor = ColorNull(dtData.Rows[iRow]["UPC_BCOLOR_S"].ToString());
                    string UpcSlowFColor = ColorNull(dtData.Rows[iRow]["UPC_FCOLOR_S"].ToString());
                    string UpcFastBColor = ColorNull(dtData.Rows[iRow]["UPC_BCOLOR_F"].ToString());
                    string UpcFastFColor = ColorNull(dtData.Rows[iRow]["UPC_FCOLOR_F"].ToString());

                    string UpaSlowQty = dtData.Rows[iRow]["UPA_SLOW"].ToString();
                    string UpaFastQty = dtData.Rows[iRow]["UPA_FAST"].ToString();
                    string UpaSlowBColor = ColorNull(dtData.Rows[iRow]["UPA_BCOLOR_S"].ToString());
                    string UpaSlowFColor = ColorNull(dtData.Rows[iRow]["UPA_FCOLOR_S"].ToString());
                    string UpaFastBColor = ColorNull(dtData.Rows[iRow]["UPA_BCOLOR_F"].ToString());
                    string UpaFastFColor = ColorNull(dtData.Rows[iRow]["UPA_FCOLOR_F"].ToString());

                    string Upa2SlowQty = dtData.Rows[iRow]["UPA2_SLOW"].ToString();
                    string Upa2FastQty = dtData.Rows[iRow]["UPA2_FAST"].ToString();
                    string Upa2SlowBColor = ColorNull(dtData.Rows[iRow]["UPA2_BCOLOR_S"].ToString());
                    string Upa2SlowFColor = ColorNull(dtData.Rows[iRow]["UPA2_FCOLOR_S"].ToString());
                    string Upa2FastBColor = ColorNull(dtData.Rows[iRow]["UPA2_BCOLOR_F"].ToString());
                    string Upa2FastFColor = ColorNull(dtData.Rows[iRow]["UPA2_FCOLOR_F"].ToString());

                    string UpoSlowQty = dtData.Rows[iRow]["UPO_SLOW"].ToString();
                    string UpoFastQty = dtData.Rows[iRow]["UPO_FAST"].ToString();
                    string UpoSlowBColor = ColorNull(dtData.Rows[iRow]["UPO_BCOLOR_S"].ToString());
                    string UpoSlowFColor = ColorNull(dtData.Rows[iRow]["UPO_FCOLOR_S"].ToString());
                    string UpoFastBColor = ColorNull(dtData.Rows[iRow]["UPO_BCOLOR_F"].ToString());
                    string UpoFastFColor = ColorNull(dtData.Rows[iRow]["UPO_FCOLOR_F"].ToString());

                    rowspan3 = $"<td  bgcolor='WHITE' style='color:BLACK' align='center' rowspan='{htDept[deptName]}' >{++iRanking}</td>";
                    rowspan = $"<td  bgcolor='WHITE' style='color:BLACK' align='center' rowspan='{htDept[deptName]}' >{deptName}</td>";
                    ;
                    rowspan2 = $"<td bgcolor='WHITE' style='color:BLACK' align='center' rowspan='{htLine[deptName + mline]}'>{mline.TrimStart('0')}</td>";


                    if (deptName == "LEAN L" && mline == "005")
                    { }


                    if (iRow > 0 && deptName == dtData.Rows[iRow - 1]["LINE_NAME"].ToString())
                    {
                        rowspan = "";
                        rowspan3 = "";
                        iRanking--;
                    }

                    if (iRow > 0 && deptName + mline == dtData.Rows[iRow - 1]["LINE_NAME"].ToString() + dtData.Rows[iRow - 1]["MLINE_CD"].ToString())
                    {
                        rowspan2 = "";
                    }


                    TableRow += "<tr>" +
                                   //  rowspan3 +
                                   rowspan +
                                   rowspan2 +
                                   //$"<td bgcolor='WHITE' style='color:BLACK' align='center'>{(iRow+1).ToString()}</td>" +
                                   //$"<td bgcolor='WHITE' style='color:BLACK' align='left'>&nbsp;{deptName}</td>" +
                                   //$"<td bgcolor='WHITE' style='color:BLACK' align='center'>{mline}</td>" +
                                   $"<td bgcolor='WHITE' style='color:BLACK' align='center'>{cLine}</td>" +
                                   //  $"<td bgcolor='WHITE' style='color:BLACK' align='left'>{Component}</td>" +
                                   $"<td bgcolor='{UpsBColor}' style='color:{UpsFColor}' align='center'>{UpsYmdQty}</td>" +
                                   $"<td bgcolor='{UpcFastBColor}' style='color:{UpcFastFColor }' align='center'>{UpcFastQty}</td>" +
                                   $"<td bgcolor='{UpcSlowBColor}' style='color:{UpcSlowFColor}' align='center'>{UpcSlowQty}</td>" +
                                   $"<td bgcolor='{UpaFastBColor}' style='color:{UpaFastFColor}' align='center'>{UpaFastQty}</td>" +
                                   $"<td bgcolor='{UpaSlowBColor}' style='color:{UpaSlowFColor}' align='center'>{UpaSlowQty}</td>" +
                                   $"<td bgcolor='{Upa2FastBColor}' style='color:{Upa2FastFColor}' align='center'>{Upa2FastQty}</td>" +
                                   $"<td bgcolor='{Upa2SlowBColor}' style='color:{Upa2SlowFColor}' align='center'>{Upa2SlowQty}</td>" +
                                   $"<td bgcolor='{UpoFastBColor}' style='color:{UpoFastFColor}' align='center'>{UpoFastQty}</td>" +
                                   $"<td bgcolor='{UpoSlowBColor}' style='color:{UpoSlowFColor}' align='center'>{UpoSlowQty}</td>" +






                               //  $"<td bgcolor='{UpsBColor}' style='color:{UpsFColor}' align='center'>{AssYmdQty}</td>" +

                               //  $"<td bgcolor='{UpsBColor}' style='color:{UpsFColor}' align='center'>{UpsHmsQty}</td>" +
                               // $"<td bgcolor='{FgaBColor}' style='color:{FgaFColor}' align='center'>{FgaYmdQty}</td>" +
                               //    $"<td bgcolor='{FgaBColor}' style='color:{FgaFColor}' align='center'>{FgaHmsQty}</td>" +

                               "</tr>";
                }

                //"<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;' >" +
                //          "<b style='background-color:yellow; color:black' >Color explanation</b><br>" +
                //            "In comparison with stitching input<br>" +
                //            "When faster or slower<br>" +
                //            "Green&nbsp;&nbsp;: 1 hour<br>" +
                //            "Yellow&nbsp;: from 2 hours to 3 hours<br>" +
                //            "Red&nbsp;&nbsp;&nbsp;&nbsp;: more than 4 hours" +
                //       "</p>"

                return strColorExplain +
                       "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' >" +
                            TableHeader + TableRow +
                       "</table>";
            }
            catch (Exception ex)
            {
                WriteLog("GetHtmlBodyCutting: " + ex.ToString());
                return "";
            }
        }

        public DataSet SEL_CUTTING_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();


            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_CUTTING";
                MyOraDB.ReDim_Parameter(5);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";
                MyOraDB.Parameter_Name[3] = "CV_2";
                MyOraDB.Parameter_Name[4] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";


                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("SEL_CUTTING_DATA: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }

        #endregion



        #region Email NPI Capture

        private void RunNPI2()
        {

            fn_BindingHeader();
            fn_BindingData();
            CaptureControl(grdBaseNpi, "GridNpi");
            CreateMailNpi();
        }

        private void CreateMailNpi()
        {
            try
            {
                //Outlook.MailItem mailItem = (Outlook.MailItem)
                // this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPicGrid1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\GridNpi.png", Outlook.OlAttachmentType.olByValue, null, "tr");

                mailItem.Subject = "NPI";


                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (!app.Session.CurrentUser.AddressEntry.Address.Contains("IT.NGOC"))
                {
                    foreach (DataRow row in dtEmail.Rows)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(row["EMAIL"].ToString());
                        oRecip.Resolve();
                    }
                }

                //Get List Send email Test
                if (chkTest.Checked)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }

                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                mailItem.Body = "This is the message.";
                string imgGrid1 = "imgGrid1";
                oAttachPicGrid1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgGrid1);
                //oAttachPicGrid1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x00390040", imgGrid1);



                mailItem.HTMLBody = String.Format(
                    "<body>" +
                          "<img src=\"cid:{0}\">" +
                    "</body>",
                     imgGrid1);
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }


        }

        public DataSet SEL_NPI_DATA2(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();

            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_NPI_TEST";
                MyOraDB.ReDim_Parameter(6);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_FACTORY";
                MyOraDB.Parameter_Name[2] = "V_P_PLANT";
                MyOraDB.Parameter_Name[3] = "V_P_FROM";
                MyOraDB.Parameter_Name[4] = "V_P_TO";
                MyOraDB.Parameter_Name[5] = "CV_1";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[4] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "2110";
                MyOraDB.Parameter_Values[2] = "ALL";
                MyOraDB.Parameter_Values[3] = DateTime.Now.ToString("yyyyMMdd");
                MyOraDB.Parameter_Values[4] = DateTime.Now.AddDays(-60).ToString("yyyyMMdd"); ;
                MyOraDB.Parameter_Values[5] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("SEL_NPI_DATA2: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_NPI_DATA2: " + ex.ToString());
                return null;
            }
        }

        private void fn_BindingHeader()
        {
            try
            {
                DataSet ds = SEL_NPI_DATA2("H");
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    grdViewNpi.Bands.Clear();
                    grdViewNpi.Columns.Clear();

                    GridBandEx gridBand1 = new GridBandEx();
                    GridBandEx gridBand2 = new GridBandEx();
                    GridBandEx gridBand3 = new GridBandEx();
                    GridBandEx gridBand4 = new GridBandEx();
                    GridBandEx gridBand5 = new GridBandEx();
                    GridBandEx gridBand6 = new GridBandEx();
                    GridBandEx gridBand7 = new GridBandEx();

                    BandedGridColumnEx column_Band1 = new BandedGridColumnEx();
                    BandedGridColumnEx column_Band2 = new BandedGridColumnEx();
                    BandedGridColumnEx column_Band3 = new BandedGridColumnEx();
                    BandedGridColumnEx column_Band4 = new BandedGridColumnEx();
                    BandedGridColumnEx column_Band5 = new BandedGridColumnEx();
                    BandedGridColumnEx column_Band6 = new BandedGridColumnEx();
                    BandedGridColumnEx column_Band7 = new BandedGridColumnEx();
                    column_Band1.Caption = "PLANT";
                    column_Band1.FieldName = "PLANT_NM";
                    column_Band1.Name = "PLANT_NM";
                    column_Band1.Visible = true;
                    column_Band1.Width = 60;

                    column_Band2.Caption = "LINE";
                    column_Band2.FieldName = "LINE_CD";
                    column_Band2.Name = "LINE_CD";
                    column_Band2.Visible = true;
                    column_Band2.Width = 60;

                    column_Band3.Caption = "CATEGORY";
                    column_Band3.FieldName = "CATEGORY_NAME";
                    column_Band3.Name = "CATEGORY_NAME";
                    column_Band3.Visible = true;
                    column_Band3.Width = 150;

                    column_Band4.Caption = "TD_CODE";
                    column_Band4.FieldName = "TD_CODE";
                    column_Band4.Name = "TD_CODE";
                    column_Band4.Visible = true;
                    column_Band4.Width = 90;

                    column_Band5.Caption = "STYLE_CODE";
                    column_Band5.FieldName = "STYLE_CODE";
                    column_Band5.Name = "STYLE_CODE";
                    column_Band5.Visible = true;
                    column_Band5.Width = 120;

                    column_Band6.Caption = "MODEL_NAME";
                    column_Band6.FieldName = "MODEL_NAME";
                    column_Band6.Name = "MODEL_NAME";
                    column_Band6.Visible = true;
                    column_Band6.Width = 320;

                    column_Band7.Caption = "PROD_DATE";
                    column_Band7.FieldName = "PROD_DATE";
                    column_Band7.Name = "PROD_DATE";
                    column_Band7.Visible = true;
                    column_Band7.Width = 120;

                    //6 Fixed band
                    gridBand1.Caption = "Plant";
                    gridBand1.Name = "gridBand1";
                    gridBand1.VisibleIndex = 0;
                    gridBand1.Columns.Add(column_Band1);
                    gridBand1.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand2.Caption = "Line";
                    gridBand2.Name = "gridBand2";
                    gridBand2.VisibleIndex = 1;
                    gridBand2.Columns.Add(column_Band2);
                    gridBand2.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand3.Caption = "Category";
                    gridBand3.Name = "gridBand3";
                    gridBand3.VisibleIndex = 2;
                    gridBand3.Columns.Add(column_Band3);
                    gridBand3.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand4.Caption = "TD Code";
                    gridBand4.Name = "gridBand4";
                    gridBand4.VisibleIndex = 3;
                    gridBand4.Columns.Add(column_Band4);
                    gridBand4.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand5.Caption = "Style Code";
                    gridBand5.Name = "gridBand5";
                    gridBand5.VisibleIndex = 4;
                    gridBand5.Columns.Add(column_Band5);
                    gridBand5.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand6.Caption = "Model Name";
                    gridBand6.Name = "gridBand6";
                    gridBand6.VisibleIndex = 5;
                    gridBand6.Columns.Add(column_Band6);
                    gridBand6.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand7.Caption = "Prod Date";
                    gridBand7.Name = "gridBand7";
                    gridBand7.VisibleIndex = 6;
                    gridBand7.Columns.Add(column_Band7);
                    gridBand7.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    grdViewNpi.Columns.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn[] { column_Band1, column_Band2, column_Band3, column_Band4, column_Band5, column_Band6, column_Band7 });
                    grdViewNpi.Bands.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { gridBand1, gridBand2, gridBand3, gridBand4, gridBand5, gridBand6, gridBand7 });
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        //2 band chung cặp
                        //  GridBandEx first_Band = new GridBandEx();
                        GridBandEx second_Band = new GridBandEx();
                        GridBandEx third_Band = new GridBandEx();
                        BandedGridColumnEx column_Band = new BandedGridColumnEx();

                        //first_Band.Caption = dt.Rows[i]["NPI_CODE"].ToString();
                        //first_Band.Name = string.Concat("first_Band", i);
                        //first_Band.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                        //first_Band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        //first_Band.VisibleIndex = i;

                        second_Band.Caption = dt.Rows[i]["NPI_DATE"].ToString();
                        second_Band.Name = string.Concat("second_Band", i);
                        second_Band.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                        second_Band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        second_Band.VisibleIndex = i;
                        second_Band.RowCount = 2;
                        third_Band.Caption = dt.Rows[i]["NPI_NAME"].ToString();
                        third_Band.Name = string.Concat("third_Band", i);
                        third_Band.VisibleIndex = i;
                        third_Band.AppearanceHeader.Options.UseTextOptions = true;
                        third_Band.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                        third_Band.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                        third_Band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        third_Band.RowCount = 20;
                        third_Band.VisibleIndex = i;

                        column_Band.Caption = dt.Rows[i]["NPI_CODE"].ToString();
                        column_Band.FieldName = dt.Rows[i]["NPI_CODE"].ToString();
                        column_Band.Name = dt.Rows[i]["NPI_CODE"].ToString();
                        column_Band.Visible = true;
                        column_Band.Width = 85;
                        third_Band.Columns.Add(column_Band);
                        grdViewNpi.Columns.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn[] { column_Band });
                        //first_Band.Children.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { second_Band });
                        second_Band.Children.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { third_Band });
                        grdViewNpi.Bands.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { second_Band });
                        grdViewNpi.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.True;
                    }

                    for (int i = 0; i < grdViewNpi.Columns.Count; i++)
                    {
                        if (i >= 7)
                        {
                            grdViewNpi.Columns[i].OwnerBand.ParentBand.AppearanceHeader.Font = new Font("Calibri", 14, FontStyle.Bold);
                        }
                        grdViewNpi.Columns[i].AppearanceHeader.Font = new Font("Calibri", 14, FontStyle.Bold);
                        grdViewNpi.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 14, FontStyle.Bold);
                        grdViewNpi.Columns[i].AppearanceCell.Font = new Font("Calibri", 14);
                        grdViewNpi.Columns[i].AppearanceCell.Options.UseTextOptions = true;
                        grdViewNpi.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        private void fn_BindingData()
        {
            try
            {
                DataSet ds = SEL_NPI_DATA2("Q");
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    DataTable dtPivot = Pivot(dt, dt.Columns["NPI_CODE"], dt.Columns["VALUE1"]);
                    DataTable dtSource = dtPivot.Copy();
                    grdBaseNpi.DataSource = dtSource;

                    for (int i = 0; i < grdViewNpi.Columns.Count; i++)
                    {
                        grdViewNpi.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                        if (i >= 7)
                        {
                            grdViewNpi.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Near;
                        }

                    }

                    grdViewNpi.Columns["CATEGORY_NAME"].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Near;
                    grdViewNpi.Columns["MODEL_NAME"].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Near;

                    grdViewNpi.OptionsBehavior.Editable = false;
                    grdViewNpi.OptionsBehavior.ReadOnly = true;
                }
                else
                {
                    grdBaseNpi.DataSource = null;
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        DataTable Pivot(DataTable dt, DataColumn pivotColumn, DataColumn pivotValue)
        {
            // find primary key columns 
            //(i.e. everything but pivot column and pivot value)
            DataTable temp = dt.Copy();
            temp.Columns.Remove(pivotColumn.ColumnName);
            temp.Columns.Remove(pivotValue.ColumnName);
            string[] pkColumnNames = temp.Columns.Cast<DataColumn>()
            .Select(c => c.ColumnName)
            .ToArray();

            // prep results table
            DataTable result = temp.DefaultView.ToTable(true, pkColumnNames).Copy();
            result.PrimaryKey = result.Columns.Cast<DataColumn>().ToArray();
            dt.AsEnumerable()
            .Select(r => r[pivotColumn.ColumnName].ToString())
            .Distinct().ToList()
            .ForEach(c => result.Columns.Add(c, pivotValue.DataType));
            //.ForEach(c => result.Columns.Add(c, pivotColumn.DataType));

            // load it
            foreach (DataRow row in dt.Rows)
            {
                // find row to update
                DataRow aggRow = result.Rows.Find(
                pkColumnNames
                .Select(c => row[c])
                .ToArray());
                // the aggregate used here is LATEST 
                // adjust the next line if you want (SUM, MAX, etc...)
                aggRow[row[pivotColumn.ColumnName].ToString()] = row[pivotValue.ColumnName];


            }

            return result;
        }

        private void grdViewNpi_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            try
            {
                if (e.Column.AbsoluteIndex >= 7)
                {
                    string ValueCell = grdViewNpi.GetRowCellValue(e.RowHandle, grdViewNpi.Columns[e.Column.FieldName]).ToString();
                    if (ValueCell.Length > 1)
                    {
                        ValueCell = ValueCell.Substring(0, 1);
                    }
                    switch (ValueCell)
                    {
                        case "Y":
                            e.Appearance.BackColor = Color.Yellow;
                            e.Appearance.ForeColor = Color.Yellow;
                            break;
                        case "G":
                            e.Appearance.BackColor = Color.Green;
                            e.Appearance.ForeColor = Color.Green;
                            break;
                        case "R":
                            e.Appearance.BackColor = Color.Red;
                            e.Appearance.ForeColor = Color.Red;
                            break;
                        case "S":
                            e.Appearance.BackColor = Color.Silver;
                            e.Appearance.ForeColor = Color.Silver;
                            break;
                        case "B":
                            e.Appearance.BackColor = Color.Black;
                            e.Appearance.ForeColor = Color.Black;
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        #endregion

        #region Email NPI

        private void RunNPI(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                
                DataSet dsData = SEL_NPI_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;
                WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                DataTable dtData = dsData.Tables[0];
                DataTable dtHeader = dsData.Tables[1];
                DataTable dtData2 = dsData.Tables[2];
                DataTable dtHeader2 = dsData.Tables[3];
                DataTable dtExplain = dsData.Tables[4];
                DataTable dtEmail = dsData.Tables[5];

                WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                string html = GetHtmlBodyNpi(dtHeader, dtData);

                string html2 = GetHtmlBodyNpi(dtHeader2, dtData2);

                string subject = dtExplain.Rows[0]["SUBJECT"].ToString();

                string explain = dtExplain.Rows[0]["TXT"].ToString();

                CreateMail(subject, explain + html + "<br>" + html2, dtEmail);
                WriteLog("RunNPI: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }

        }

        private string GetHtmlBodyNpi(DataTable dtHeader, DataTable dtData)
        {
            try
            {
                string TableHeader = "";

                int i = 0;
                int npiCode = 0;
                string HeaderRow1 = "", HeaderRow2 = "";

                // int[] colWidth = { 55, 55, 55, 60, 66, 55, 70, 80, 70, 80, 65, 65, 60, 70, 65, 55, 75, 55 };

                int[] colWidth = { 67, 67, 67, 67, 67, 67, 67, 80, 67, 80, 67, 67, 67, 67, 67, 67, 75, 67 };

                foreach (DataRow row in dtHeader.Rows)
                {

                    int.TryParse(row["NPI_CODE"].ToString(), out npiCode);
                    if (npiCode >= 10)
                    {

                        HeaderRow1 += $"<td bgcolor = '#ff9900' style = 'color:#ffffff' align = 'center' width = '{colWidth[i]}'>{row["NPI_DATE"]}</td>";
                        HeaderRow2 += $"<td bgcolor = '#00ced1' style = 'color:#ffffff' align = 'center' width = '{colWidth[i]}'>{row["NPI_NAME"]}</td>";
                        i++;
                    }

                    else
                        HeaderRow1 += $"<td bgcolor = '#00ced1' style = 'color:#ffffff' rowspan ='2' align = 'center'>{row["NPI_NAME"]}</td>";


                }



                TableHeader = "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow1 + "</tr> " +
                              "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow2 + "</tr> ";

                //Row
                string TableRow = "", rowspan = "", rowspan2 = "", rowspan3 = "";

                for (int iRowData = 0; iRowData < dtData.Rows.Count; iRowData++)
                {
                    string plantNm = dtData.Rows[iRowData]["PLANT_NM"].ToString();
                    string lineCd = dtData.Rows[iRowData]["LINE_CD"].ToString();
                    string category = dtData.Rows[iRowData]["CATEGORY_NAME"].ToString();
                    string tdCode = dtData.Rows[iRowData]["TD_CODE"].ToString();
                    string modelName = dtData.Rows[iRowData]["MODEL_NAME"].ToString();
                    string styleCode = dtData.Rows[iRowData]["STYLE_CODE"].ToString();
                    string prodDate = dtData.Rows[iRowData]["PROD_DATE"].ToString();
                    string backColor = dtData.Rows[iRowData]["BCOLOR"].ToString();

                    string plantNm_prev = "";
                    string lineCd_prev = "";
                    string category_prev = "";
                    string tdCode_prev = "";
                    string modelName_prev = "";
                    string styleCode_prev = "";
                    string prodDate_prev = "";

                    if (iRowData > 0)
                    {
                        plantNm_prev = dtData.Rows[iRowData - 1]["PLANT_NM"].ToString();
                        lineCd_prev = dtData.Rows[iRowData - 1]["LINE_CD"].ToString();
                        category_prev = dtData.Rows[iRowData - 1]["CATEGORY_NAME"].ToString();
                        tdCode_prev = dtData.Rows[iRowData - 1]["TD_CODE"].ToString();
                        modelName_prev = dtData.Rows[iRowData - 1]["MODEL_NAME"].ToString();
                        styleCode_prev = dtData.Rows[iRowData - 1]["STYLE_CODE"].ToString();
                        prodDate_prev = dtData.Rows[iRowData - 1]["PROD_DATE"].ToString();
                    }

                    if (plantNm != plantNm_prev || lineCd != lineCd_prev || category != category_prev || tdCode != tdCode_prev ||
                        modelName != modelName_prev || styleCode != styleCode_prev || prodDate != prodDate_prev)
                    {
                        if (iRowData > 0) TableRow += "</tr> ";
                        TableRow += "<tr> " +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left'>{plantNm }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='center' >{ lineCd }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left'   >{ category }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='center' >{ tdCode }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left' >{styleCode }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left'   >{ modelName }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='center' >{ prodDate }</td>" +
                                $"<td bgcolor='{backColor}' style='color:BLACK' width = '50' align='center' ></td>"
                              ;
                    }
                    else
                    {
                        TableRow += $"<td bgcolor='{backColor}' style='color:BLACK' width = '50' align='center'></td>";
                    }
                }

                return "<table style='font-family:Calibri; font-size:15px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' with = '5000' >" +
                            TableHeader + TableRow +
                       "</table>";
            }
            catch (Exception ex)
            {
                WriteLog("GetHtmlBodyCutting: " + ex.ToString());
                return "";
            }
        }

        public DataSet SEL_NPI_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();

            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_NPI";
                MyOraDB.ReDim_Parameter(9);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_LOC";
                MyOraDB.Parameter_Name[2] = "V_P_DATE";
                MyOraDB.Parameter_Name[3] = "CV_1";
                MyOraDB.Parameter_Name[4] = "CV_2";
                MyOraDB.Parameter_Name[5] = "CV_3";
                MyOraDB.Parameter_Name[6] = "CV_4";
                MyOraDB.Parameter_Name[7] = "CV_EXPLAIN";
                MyOraDB.Parameter_Name[8] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[7] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[8] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = V_P_DATE;
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";
                MyOraDB.Parameter_Values[6] = "";
                MyOraDB.Parameter_Values[7] = "";
                MyOraDB.Parameter_Values[8] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("P_SEND_EMAIL_NPI: null");
                    }
                    return null;
                }
                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }



        public DataSet SEL_TMD_DASH_DATA(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;

            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_TMS_ORDER";
                MyOraDB.ReDim_Parameter(4);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "CV_1";
                MyOraDB.Parameter_Name[2] = "CV_2";
                MyOraDB.Parameter_Name[3] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";


                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("SEL_CUTTING_DATA: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
        public DataSet SEL_TIME_CONTRAINT_DATA(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;

            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_TIME_CONTRAINT";
                MyOraDB.ReDim_Parameter(4);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "CV_1";
                MyOraDB.Parameter_Name[2] = "CV_2";
                MyOraDB.Parameter_Name[3] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";


                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        WriteLog("SEL_CUTTING_DATA: null");
                    }
                    return null;
                }

                return ds_ret;
            }
            catch (Exception ex)
            {
                WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
        #endregion



        private string ColorNull(string argColor)
        {
            return argColor == "" ? "WHITE" : argColor;
        }

        private void WriteLog(string argText)
        {

            txtLog.BeginInvoke(new Action(() =>
            {
                txtLog.Text += argText + "\r\n";
                txtLog.SelectionStart = txtLog.TextLength;
                txtLog.ScrollToCaret();
                txtLog.Refresh();
            }));
        }

        private void btnRunTMS_Click(object sender, EventArgs e)
        {
            RunTMSDash("Q");
        }

        private void checkRunning()
        {
            foreach (Process p in Process.GetProcessesByName("Send_Email"))
            {
                p.CloseMainWindow();
            }
        }



        private void RunTMSDash(string arg_type)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_TMD_DASH_DATA(arg_type);
                if (dsData == null) return;

                DataTable dtHeader = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                string html = getHTMLBodyHeaderTMSDash(dtHeader, dtData);

                CreateMail(Emoji.ChartIncreasing + " TMS MONITORING SUMMARY", html, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void RunTimeContraint(string arg_type)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_TIME_CONTRAINT_DATA(arg_type);
                if (dsData == null) return;

                DataTable dtHeader = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                string html = getHTMLBodyHeaderTimeContraint(dtHeader, dtData);

                CreateMail(Emoji.ChartIncreasing + " TIME CONTRAINT", html, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void btnTimeContraint_Click(object sender, EventArgs e)
        {
            try
            {
                RunTimeContraint("Q");
            }
            catch
            {


            }
        }

        private string getHTMLBodyHeaderTMSDash(DataTable dtHead, DataTable dtData)
        {
            try
            {
                string style = System.IO.File.ReadAllText(Application.StartupPath + "\\tmsdashstyle.txt");
                string header1 = System.IO.File.ReadAllText(Application.StartupPath + "\\tmsdashhtml.txt");
                string body = string.Empty;
                
                object[] argsBody = new object[dtData.Columns.Count];
                for (int i = 0; i < dtData.Rows.Count; i++)
                {
                    for (int j = 0; j < dtData.Columns.Count; j++)
                    {
                        if (j > 0)
                            argsBody[j - 1] = dtData.Rows[i][j].ToString();
                    }
                    string body1 = string.Empty;
                    if (!string.IsNullOrEmpty(dtData.Rows[i]["REASON"].ToString()))
                    {
                        body1 = string.Format("<tr style='color:red;font-weight:bold;border-style:dotted;'><td colspan = '3' align = 'center'>Reason Of Plant {0} </td><td colspan = '35'>{1}</td></tr> ", dtData.Rows[i]["PLANT"], dtData.Rows[i]["REASON"]);
                    }
                    body += string.Format(@"</tr>
                               <tr align='right' style='font-weight:bold;'>
                               <td align='center'>{0}</td>
                               <td align='center'>{1}</td>
                               <td align='center' bgcolor='{40}' style='color:{41}'>{2}</td>
                               <td bgcolor='#fff4b0'>{3}</td>
                               <td>{4}</td>
                               <td bgcolor='{38}' style='color:{39}'>{5}</td>
                               <td style='color:red'>{6}</td>
                               <td>{7}</td>
                               <td bgcolor='#fff4b0'>{8}</td>
                               <td>{9}</td>
                               <td>{10}</td>
                               <td>{11}</td>
                               <td>{12}</td>
                               <td bgcolor='#fff4b0'>{13}</td>
                               <td>{14}</td>
                               <td>{15}</td>
                               <td>{16}</td>
                               <td>{17}</td>
                               <td bgcolor='#fff4b0'> {18} </td>
                               <td>{19}</td>
                               <td>{20}</td>
                               <td>{21}</td>
                               <td>{22}</td>
                               <td bgcolor='#fff4b0'>{23}</td>
                               <td>{24}</td>
                               <td>{25}</td>
                               <td>{26}</td>
                               <td>{27}</td>
                               <td bgcolor='#fff4b0'> {28}</td>
                               <td>{29}</td>
                               <td>{30}</td>
                               <td>{31}</td>
                               <td>{32}</td>
                               <td bgcolor='#fff4b0'> {33} </td>
                               <td>{34}</td>
                               <td>{35}</td>
                               <td>{36}</td>
                               <td>{37} </td>
                               <!--<td align='left'>{42} </td>-->
                              </tr>" + body1, argsBody);
                }
                object[] argsHeader = new object[dtHead.Rows.Count + 1];
                for (int i = 0; i < dtHead.Rows.Count; i++)
                {
                    argsHeader[i] = dtHead.Rows[i]["CAPTION"].ToString();
                }
                argsHeader[6] = dtHead.Rows[0]["REMARKS"].ToString();
                string end = @"</table><hr></body></html>";
                string remakeHeader = string.Format(header1, argsHeader);

                //  string remakeBody = string.Format(body, argsBody);
                return string.Concat(style, remakeHeader, body, end);

            }
            catch (Exception ex)
            {
                return null;
            }

        }


        private string getHTMLBodyHeaderTimeContraint(DataTable dtHead, DataTable dtData)
        {
            try
            {
                string sHTML = string.Empty;
                string sHeader1 = string.Empty;
                string sHeader2 = string.Empty;
                string sHeader3 = string.Empty;
                string sBody = string.Empty;
                string sBody2 = string.Empty;
                string sStyle = string.Empty;
                string EndTag = string.Empty;
                sStyle = @"<html><head>
		                                            <style type='text/css'>
			                                            .tg  {border-collapse:collapse;border-spacing:0;}
                                                        .tg td{border-color:#9dcc7a;border-style:solid;border-width:1px;background-color:#ffffff;font-family:Arial, Helvetica, sans-serif ;font-size:12px;
                                                          overflow:hidden;padding:10px 5px;word-break:normal;}
                                                        .tg th{border-color:#9dcc7a;border-style:solid;border-width:1px;background-color:#0080a0;color: #ffffff; font-family:Arial, sans-serif;font-size:12px;
                                                          font-weight:bold;overflow:hidden;padding:5px 10px;word-break:normal;}
                                                        .tg .tg-0lax{text-align:center;}
                                                        .tg .tg-1lax{text-align:left;}
                                                        .tg .tg-2lax{text-align:right;}
                                                        .tg .tg-total{text-align:right;background-color:yellow;color: #000000;}
                                                        @media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;}}</style>";
                string Colspan = dtHead.Rows[0]["COLSPAN"].ToString();
                sHeader1 += string.Format(@"
                                                    </head>
                                                  <body>
                                                            <div class='tg-wrap'><table class='tg'>
                                                            <thead>
                                                              <tr>
                                                                <th class='tg-0lax' rowspan='2' id = 'test'>Process</th>
                                                                <th class='tg-0lax' rowspan='2'>Model Name</th>
                                                                <th class='tg-0lax' rowspan='2'>Style Code</th>
                                                                <th class='tg-0lax' colspan='{0}'>Size</th>
                                                                <th class='tg-0lax' rowspan='2'>Total</th>
                                                                <th class='tg-0lax' rowspan='2'>Reason</th>
				                                              </tr><tr>", Colspan);
                for (int j = 0; j < dtHead.Rows.Count; j++)
                {
                    string SIZE_CODE = dtHead.Rows[j]["SIZE_CODE"].ToString();
                    sHeader2 += string.Format(@"<th class='tg-0lax'>{0}</th>", SIZE_CODE);
                    string iDx = (j + 5).ToString();
                    sBody2 += @"<td class='tg-2lax'>{" + iDx + "}</td>";
                }
                sHeader3 = @"</tr></thead><tbody>";
              
                DataView view = dtData.DefaultView;
                view.Sort = "SIZE_NO,PROC_CODE,STYLE_CODE";
                DataTable dt = view.ToTable();
                dt.Columns.Remove(dt.Columns["SIZE_NO"]);
                DataTable dtGrid = Pivot(dt, dt.Columns["SIZE_CODE"], dt.Columns["QTY"]);
                object[] argBodys = new object[dtGrid.Columns.Count];
                for (int iRow = 0; iRow < dtGrid.Rows.Count; iRow++)
                {
                    for (int iCol = 0; iCol < dtGrid.Columns.Count; iCol++)
                    {
                        if (iCol > 0)
                        {
                            if (dtGrid.Columns[iCol].ColumnName.Equals("TOTAL"))
                                argBodys[iCol - 1] = string.Format("{0:n0}", dtGrid.Rows[iRow][iCol]);
                            else
                                argBodys[iCol - 1] = dtGrid.Rows[iRow][iCol].ToString();
                        }

                    }
                    sBody += string.Format(@"<tr><td class='tg-0lax'>{0}</td>
				                       <td class='tg-1lax'>{1}</td>
				                       <td class='tg-0lax'>{2}</td>
				                        " + sBody2 + " <td class='tg-total'>{3}</td><td class='tg-1lax'>{4}</td></tr>", argBodys);
                }

                EndTag = @"</tbody></table></div>
                     </body>
                      </html>";
                //<script type='text/javascript'>document.getElementById('test').innerHTML = 'Hello World';</script>
                sHTML = string.Concat(sStyle, sHeader1, sHeader2, sHeader3, sBody, EndTag);
                return sHTML;
            }
            catch (Exception ex)
            {
                return null;
            }

        }



    }
}
