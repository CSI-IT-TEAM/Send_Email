using JPlatform.Client.Controls6;
using System;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
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

            chart2.Size = new Size(1950, 1035);

            tmrLoad.Enabled = true;
            this.Text = "20200730090000";
        }

        DataTable dtEmail;
        bool _isRun = false, _isRun2 = false;
        int _start_column = 0;
        //"jungbo.shim@dskorea.com", "nguyen.it@changshininc.com",
        readonly string[] _emailTest = { "jungbo.shim@dskorea.com", "nguyen.it@changshininc.com", "dien.it@changshininc.com" };

        private void tmrLoad_Tick(object sender, EventArgs e)
        {
            RunProduction("Q1");
            RunAndon("Q1");
            Run("Q1");
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

        #region Email Production

        private void RunProduction(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_PROD_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;

                DataTable dtDate = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                CreateMailProduction(dtDate, dtData, dtEmail);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
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
                                  "</tr>";
                        }
                           
                    }
                }    


                string strDate = "";
                foreach (DataRow row in dtDate.Rows)
                {
                    strDate += "<th bgcolor = '#ff9900' style = 'color:#ffffff' align = 'center' width = '70' >" + row["YMD"].ToString() + " </th >";
                }


                    string html = "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1000'>" +
                                  "<tr bgcolor='#ffe5cc'>" +
                                     " <th rowspan = '2' align='center' width='70'>Plant</th>" +
                                     " <th rowspan = '2' align='center' width='70'>Mini Line</th>" +
                                     " <th bgcolor = '#ff9900' style = 'color:#ffffff' colspan = '6' align='center'>Full time on previous day performace</th>" +
                                     " <th bgcolor = '#366cc9' style = 'color:#ffffff' colspan = '4' align='center'> Before lunch on today performace</th>" +
                                  "</tr>" +
                                  "<tr>" +
                                     strDate +
                                     "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='70'>Daily Plan</th>" +
                                     "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='70'>Real Plan</th>" +
                                     "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='70'>Actual</th>" +
                                     "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='70'>Ratio(%)</th>" +
                                  "</tr>" +
                                    rowValue +
                              "</table>";

                //string html = "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='600'>" +
                //                  "<tr bgcolor='#366cc9' style='color:#ffffff'>" +
                //                     "<th style='color:#ffffff' align='center' width='70'>Plant</th>" +
                //                     "<th style='color:#ffffff' align='center' width='70'>Mini Line</th>" +
                //                     "<th style='color:#ffffff' align='center' width='70' >Daily Plan</th>" +
                //                     "<th style='color:#ffffff' align='center' width='70'>Real Plan</th>" +
                //                     "<th style='color:#ffffff' align='center' width='70'>Actual</th>" +                                     
                //                     "<th style='color:#ffffff' align='center' width='70'>Ratio(%)</th>" +
                //                  "</tr>" +
                //                    rowValue +
                //              "</table>";



                mailItem.HTMLBody = html;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
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

                if (ds_ret == null) return null;
                return ds_ret;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region Email ANDON
        

        private void RunAndon(string argType)
        {
            DataSet dsData = SEL_ANDON_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
            if (dsData == null) return;
            DataTable dtData = dsData.Tables[0];
            DataTable dtEmail = dsData.Tables[1];
            CreateMailAndon(dtData, dtEmail);
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
                                    "Total Downtime per Line & Downtime by line = under 10 minutes is green and from 10 min to 29:59 is yellow and more than 30 min is red" +
                               "</p>" +
                              "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;'>" +
                                    "Total average measure & Downtime average by line = under 2 minutes is green and from 2 min to 4:59 is yellow and more than 5 min is red" +
                              "</p>" +
                              "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;'>" +
                                    "Machine D/T(Min) = under 96 minutes is green and from 96 min to 120 min is yellow and more than 120 min is red" +
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
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='100'>Line</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='200'>Downtime by Line</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='200'>Calling Times by Line</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='200'>Downtime Average by Line</th>" +
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
                Console.WriteLine(ex.ToString());
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
                CaptureControl(chart2, "Chart");
                CaptureControl(grdBase1, "Grid1");
                CaptureControl(grdBase2, "Grid2");
                CaptureControl(grdBase3, "Grid3");
                CaptureControl(grdBase4, "Grid4");
                CreateMailBottomInventory();
            }
            catch {lblStatus.Text = DateTime.Now.ToString() + "Do not Send!"; }
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
            chart2.Series[3].ArgumentDataMember = "LINE_NM";
            chart2.Series[3].ValueDataMembers.AddRange(new string[] { "TAR_QTY" });

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
            bmp.Save(Path + nameImg +  @".png", System.Drawing.Imaging.ImageFormat.Png); //you could ave in BPM, PNG  etc format.
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

        #region No use

        private void CreateMailAndon_BAK(DataTable dtData, DataTable dtEmail)
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
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["MACHINE_CNT_LINE"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center' " +
                                            "bgcolor='" + dtData.Rows[iRow]["BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR"].ToString() + "'>" +
                                                dtData.Rows[iRow]["DOWNTIME_LINE"].ToString() +
                                        "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["RECEIVE_LINE"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["CALLING_TIMES_LINE"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center' " +
                                            "bgcolor='" + dtData.Rows[iRow]["BG_COLOR2"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR2"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_LINE_AVG"].ToString() +
                                         "</td>" +


                                        //Mline
                                        "<td align ='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["MACHINE_CNT_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["RECEIVE_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["CALLING_TIMES_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["AVERAGE_ELAPSE_MLINE"].ToString() + "</td>" +

                                   "</tr>";
                    }
                    // bgcolor='" + dtData.Rows[iRow]["BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR"].ToString() + "'
                    if (iRow > 0)
                    {
                        if (dtData.Rows[iRow]["PLANT"].ToString() == dtData.Rows[iRow - 1]["PLANT"].ToString())
                        {
                            rowValue += "<tr>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["MACHINE_CNT_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["RECEIVE_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["CALLING_TIMES_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["AVERAGE_ELAPSE_MLINE"].ToString() + "</td>" +
                                    "</tr>";
                        }
                        else
                        {
                            rowValue += "<tr>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["RANKING"].ToString() + " </td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["MACHINE_CNT_LINE"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center' " +
                                            "bgcolor='" + dtData.Rows[iRow]["BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR"].ToString() + "'>" +
                                                dtData.Rows[iRow]["DOWNTIME_LINE"].ToString() +
                                        "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["RECEIVE_LINE"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["CALLING_TIMES_LINE"].ToString() + "</td>" +
                                        "<td rowspan='" + strRowSpan + "' align ='center' " +
                                            "bgcolor='" + dtData.Rows[iRow]["BG_COLOR2"].ToString() + "' style='color:" + dtData.Rows[iRow]["FORE_COLOR2"].ToString() + "'>" +
                                            dtData.Rows[iRow]["DOWNTIME_LINE_AVG"].ToString() +
                                         "</td>" +


                                        //Mline
                                        "<td align ='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["MACHINE_CNT_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["DOWNTIME_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["RECEIVE_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["CALLING_TIMES_MLINE"].ToString() + "</td>" +
                                        "<td align ='center'>" + dtData.Rows[iRow]["AVERAGE_ELAPSE_MLINE"].ToString() + "</td>" +

                                   "</tr>";

                        }
                    }
                }

                string text = "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;' >" +
                                    "Machine Downtime time = under 96 minutes is green and from 96 min to 120 min is yellow and more than 120 min is red" +
                               "</p>" +
                              "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;'>" +
                                    "Time average receive = under 2 minutes is green and from 2 min to 5 min is yellow and more than 5 min is red" +
                              "</p>"
                              ;

                string html = text +
                            "<br>" +
                            "<table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1000px'>" +
                                  "<tr bgcolor='#366cc9' style='color:#ffffff'>" +
                                     "<th style='color:#ffffff' align='center' width='100'>Ranking</th>" +
                                     "<th style='color:#ffffff' align='center' width='100'>Plant</th>" +
                                     "<th style='color:#ffffff' align='center' width='100'>Machine Total</th>" +
                                     "<th style='color:#ffffff' align='center' width='200'>Machine D/T(Min)</th>" +
                                     "<th style='color:#ffffff' align='center' width='200'>Time Receive Andon Total(Min)</th>" +
                                     "<th style='color:#ffffff' align='center' width='150'>Total Calling Times(Andon)</th>" +
                                     "<th style='color:#ffffff' align='center' width='200'>Time Average Receive(Min)</th>" +
                                     //Mline
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='70'>Line</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='70'>Machine Total</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='200'>Machine D/T(Min)</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='200'>Time Receive Andon Total(Min)</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='150'>Calling Times by Line(Andon)</th>" +
                                     "<th bgcolor='#f5b038' style='color:#ffffff' align='center' width='200' >Time Average Receive(Min)</th>" +
                                  "</tr>" +
                                    rowValue +
                              "</table>";



                mailItem.HTMLBody = html;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }


        }



        private void DrawControlToBitmap(Control control)
        {
            string Path = Application.StartupPath + @"\Capture\";
            Bitmap bitmap = new Bitmap(control.Width, control.Height);
            Graphics graphics = Graphics.FromImage(bitmap);
            Rectangle rect = control.RectangleToScreen(control.ClientRectangle);
            graphics.CopyFromScreen(rect.Location, Point.Empty, control.Size);
            bitmap.Save(Path + @"Capture.png", ImageFormat.Png);


        }

        private void CreateMailItem()
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
                mailItem.Subject = "This is Test Send Email";
                mailItem.To = "ngoc.it@changshininc.com";
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



        private void SaveControlImage(Control theControl)
        {
            try
            {
                string Path = Application.StartupPath + @"\Capture\";

                if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);

                Bitmap controlBitMap = new Bitmap(theControl.Width, theControl.Height);
                Graphics g = Graphics.FromImage(controlBitMap);
                g.CopyFromScreen(PointToScreen(theControl.Location), new Point(0, 0), theControl.Size);

                // example of saving to the desktop
                controlBitMap.Save(Path + @"Capture.png", ImageFormat.Png);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }


        #endregion

        
    }
}
