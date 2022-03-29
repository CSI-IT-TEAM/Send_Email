using DevExpress.Utils;
using DevExpress.XtraCharts;
using DevExpress.XtraGrid.Views.BandedGrid;
using JPlatform.Client.Controls6;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Send_Email
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();

            grdBase1.Size = new Size(1980, 1320);
            grdBase2.Size = new Size(1980, 1150);
            grdBase3.Size = new Size(1980, 1370);
            grdBase4.Size = new Size(1980, 1450);

            grdBaseNpi.Size = new Size(4000, 700);

            panel1.Size = new Size(1950, 1100);
            chart2.Size = new Size(1950, 900);

            //Phước Thêm TMS Dass
            pnTMSDassChart.Size = new Size(1700, 500);
            pnTMSDassGrid.Size = new Size(1420, 215);

            //Rework Monthly
            pnchartReworkPlant.Size = new Size(1150, 400);
            pnChartReworkReason.Size = new Size(400, 400);
            pnchartBCGrade.Size = new Size(1560, 400);
            grdRework.Size = new Size(1560, 270);

            pnMold.Size = new Size(2000, 1000);
            chartMold.Size = new Size(750, 700);
            grdMain.Size = new Size(1700, 300);

            pnChartFGA_INV.Size = new Size(1400, 400);


            tmrLoad.Enabled = true;
            //this.Text = "20210710133500";
            //this.Text = "20211008083800";
            this.Text = "Data as a service";
            var monday = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
        }

        //Phuoc.IT

        private string[] headNames = new string[] { "COMP" };
        private string[] divNames = new string[] { "Order By Set (prs)", "Total Outgoing (prs)", "Per", "", "", "", "", "", "", "DMP-Y", "IP-Y", "PU-Y", "OS-Y", "PH-Y" };
        private DataTable dtEmail;
        private bool _isRun = false, _isRun2 = false;
        private int _start_column = 0;

        //"jungbo.shim@dskorea.com", "nguyen.it@changshininc.com", "dien.it@changshininc.com", "do.it@changshininc.com"
        //, "nguyen.it@changshininc.com", "dien.it@changshininc.com", "ngoc.it@changshininc.com", "yen.it@changshininc.com"
        //readonly string[] _emailTest = {   "do.it@changshininc.com", "nguyen.it@changshininc.com", "dien.it@changshininc.com", "ngoc.it@changshininc.com", "yen.it@changshininc.com" };
        private readonly string[] _emailTest = { "nguyen.it@changshininc.com", "dien.it@changshininc.com" }; //,"nguyen.it@changshininc.com",

        #region Event

        int _iCount = 0;
        private void tmrLoad_Tick(object sender, EventArgs e)
        {
            _iCount++;
            if (_iCount < 60) return;
            _iCount = 0;

            string TimeNow = System.DateTime.Now.ToString("HH:mm");
            DateTime today = DateTime.Today;

            //if (TimeNow.Equals("09:00"))
            //    if (cmdPORegisterChk.Checked)
            //        RunPORegReport("Q");

            //12h
            Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            Debug.WriteLine(_iCount);
            if (cmdPoToChk.Checked)
                RunToPo("Q1");
            if (cmdPoToIeChk.Checked)
                RunToPoIe("Q1");
            if (cmdRunProdChk.Checked)
                RunProduction("Q1");



            //12H - Phước thêm 2021/12/07
            if (TimeNow.Equals("06:58"))
            {
                if (cmdRunAssInLineChk.Checked)
                    RunAssInLine("Q1");
            }

            //Upper Inventory
            if (btnRunUpperInvChk.Checked)
                RunUpperInv("Q", DateTime.Now.ToString("yyyyMMdd"));


            //7h
            if (cmdRunEscanChk.Checked)
                RunEScan("Q1");
            if (cmdRunAndonChk.Checked)
                RunAndon("Q1");
            if (button1Chk.Checked)
                Run("Q1");

            //8h
            if (btnRunTMSChk.Checked)
                RunTMSDash("Q1");

            if (cmdRunSumDaaSChk.Checked)
                RunSumDaaS("Q1");

            if (cmdNpiChk.Checked)
                RunNPI("Q1");
            if (cmdMoldRepairChk.Checked)
                RunMoldRepair("Q1");
            if (cmd_BudgetChk.Checked)
                RunBuget("Q1");
            if (cmd_QualityChk.Checked)
                RunQuality("Q1");
            if (cmd_Quality2Chk.Checked)
                RunQuality2("Q1");
            if (cmd_BotDefChk.Checked)
                RunBotDef("Q1");

            if (cmdCanteenChk.Checked)
                RunCanteen("Q1");

            if (cmdMoldRepairMonthChk.Checked)
                RunMoldRepairMonth("Q1");

            if (cmdMoldRepairMonthChk.Checked)
                RunMoldRepairMonthWh("Q1");
            //15h
            if (cmd_IeReliefChk.Checked)
                RunIeRelief("Q1");

            //16h
            if (cmdCuttingChk.Checked)
                RunCutting("Q1");

            //10h
            if (btnTimeContraintChk.Checked)
            {
                if (TimeNow.Equals("09:59"))
                {
                    RunTimeContraint("Bottom", "Q1"); //BOTTOM

                }
                if (TimeNow.Equals("10:03"))
                {
                    RunTimeContraint("Stockfit", "Q2"); //STOCKFIT

                }

            }

            if (cmd_HourlyProdTrackingChk.Checked)
            {
                RunEscanSituationTracking("Q1");
            }

            if (cmd_BolRrChk.Checked)
            {
                RunBolRr("Q1");
            }


            //16h

            if (TimeNow.Equals("16:00"))
            {
                if (btnRunTMSV2Chk.Checked)
                    RunTMSDashv2("Q");
                if (btnRunScadaChk.Checked)
                    RunScada("Q");
            }
            if (btnTMS_SummaryChk.Checked)
            {
                if (today.DayOfWeek == DayOfWeek.Monday && TimeNow.Equals("08:00"))
                {
                    RunTMS_Summary("Q");
                }
            }

            if (btnRunOS_RedChk.Checked)
            {
                //int min = int.Parse(TimeNow.Substring(3, 2));
                //Debug.WriteLine("OS: " + min % 5);
                //if (min % 5 == 0)
                //{
                //    RunOSRedMachine("Q", DateTime.Now.ToString("yyyyMMdd"), "10");
                //}


                switch (TimeNow)
                {
                    case "06:10":
                    case "10:10":
                    case "14:10":
                    case "18:10":
                    case "22:10":
                    case "02:10":
                        RunOSRedMachine("Q1", DateTime.Now.ToString("yyyyMMdd"), TimeNow.Substring(0, 2));

                        break;
                }
            }

            ////07 - Do IT thêm 2022/03/24
            //if (TimeNow.Equals("06:59"))
            //    if (btnRunUpperInvChk.Checked)
            //        RunUpperInv("Q", DateTime.Now.ToString("yyyyMMdd"));
        }


        private void cmdRunProd_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunProduction("Q");
        }

        private void cmdRunAndon_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunAndon("Q");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                Run("Q");
        }

        private void cmdRunEscan_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunEScan("Q");
        }

        private void cmdPoTo_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunToPo("Q");
        }

        private void cmdPoToIe_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunToPoIe("Q");
        }

        private void cmdCutting_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunCutting("Q");
        }

        private void cmdNpi_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunNPI("Q");
            //RunNPI2();
        }

        private void cmdMoldRepair_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunMoldRepair("Q");
        }

        private void cmdRunOS_Red_MC_Click(object sender, EventArgs e)
        {
            //RunOSRedMachine("Q", DateTime.Now.ToString("yyyyMMdd"), DateTime.Now.ToString("HH"));
            //Run TEst
            if (SendYN(((Button)sender).Text))
                RunOSRedMachine("Q", DateTime.Now.ToString("yyyyMMdd"), DateTime.Now.ToString("HH"));

        }
        private void btnRunOS_Monthly_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunOSMonthly("Q", DateTime.Now.ToString("yyyyMMdd"));
            //  RunOSMonthly("Q", "20220314");
        }

        private void cmdPORegister_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunPORegReport("Q");
        }


        private void cmd_Budget_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunBuget("Q");
        }
        private void btnP_Click(object sender, EventArgs e)
        {

        }
        #endregion Event

        private bool SendYN(string MailSubject)
        {
            DialogResult dl = MessageBox.Show(this, string.Concat("Bạn có muốn gửi mail", "\n", MailSubject), "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (dl == DialogResult.Yes)
                return true;
            return false;
        }

        private void CreateMailFeedBack(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\feedback.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = htmlBody + String.Format(@"<body><br><img src='cid:{0}'></body>", imgInfo);

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();

                Send_Feedback.UPD_DATA();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailFeedBack: " + ex.ToString());
            }
        }

        private void CreateMailTimeContraint(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\constraint_kr.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\constraint_vi.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo", imgInfo1 = "imgInfo1";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo1);
                mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'><br><br><img src='cid:{1}'></body>", imgInfo, imgInfo1) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        private void CreateMail(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                //Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\outsole.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                //  string imgInfo = "imgInfo";
                // oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        private void CreateMailQuality(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\quality_ko.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'></body>", imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        private void CreateMailQuality2(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\quality2_ko.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'></body>", imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        private void CreateMailBotDef(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                string shift = dtEmail.Rows[0]["SHIFT"].ToString();
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Shift" + shift + ".jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\botDef_ko.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo", imgShift = "imgShift";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgShift);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'><br><img src='cid:{1}'></body>", imgShift, imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        private void CreateMailReworkMonthly(string Subject, string htmlBody, DataTable dtEmail, DataTable argStyle)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);

                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\chartReworkMonthly.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\chartReworkReasonMonthly.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic3 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\chartBCGrade.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic4 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\gridRework.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic5 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\quality_month.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                //Explaination
                string strStyle = argStyle.Rows[0]["STYLE"].ToString();
                string strTitle1 = argStyle.Rows[0]["TITLE1"].ToString();
                string strTitle2 = argStyle.Rows[0]["TITLE2"].ToString();
                string strExplain = argStyle.Rows[0]["TXT"].ToString();


                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
                string imgChart1 = "imgChart1", imgChart2 = "imgChart2", imgChart3 = "imgChart3", gridRework = "gridRework", imgTitle = "imgTitle";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgChart1);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgChart2);
                oAttachPic3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgChart3);
                oAttachPic4.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", gridRework);
                oAttachPic5.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgTitle);
                //mailItem.HTMLBody = String.Format(
                //    @"<body><img src='cid:{0}'></body>", imgInfo) + htmlBody;
                mailItem.HTMLBody = string.Format(
                    @"
                    <html>{5}
                    <body>
                    <img src='cid:{4}'
                    {6} {7} {8} 
                    <table class='tg'>
                    <thead>
                      <tr>
                        <th class='tg-0lax'><img src='cid:{0}'></th>
                        <th class='tg-0lax'><img src='cid:{1}'></th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                       <td class='tg-0lax' colspan='2'><img src='cid:{2}'></td>
                      </tr>
                      <tr>
                       <td class='tg-0lax' colspan='2'><img src='cid:{3}'></td>
                      </tr>
                    </tbody>
                    </table>
                    </body></html>
                    ", imgChart1, imgChart2, imgChart3, gridRework, imgTitle, strStyle, strExplain, strTitle1, strTitle2
                    );
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailReworkMonthly: " + ex.ToString());
            }
        }





        private void CreateMailOs(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\outsole.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = String.Format(
                    @"<body><img src='cid:{0}'></body>", imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                WriteLog("CreateMailOS: send");
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailOS: " + ex.ToString());
            }
        }

        private void CreateMailwithImage(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                // Outlook.Attachment oAttach = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\TMSChart.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                //  Outlook.Attachment oAttachPicGrid1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\TMSGrid.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                // string imgChart = "imgChart", imgGrid1 = "imgGrid1";
                //  oAttach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgChart);
                //  oAttachPicGrid1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgGrid1);
                // string EmbedImg = string.Format(@"<table class='tftable' border='1' width='100%' cellspacing='0' cellpadding='0'><tr><td class='tftable-clax'><img src='cid:{0}'></td></tr><tr><td class='tftable-clax'><img src='cid:{1}'</td></tr></table></body></html>", imgChart, imgGrid1);
                string endTag = "</body></html>";
                mailItem.HTMLBody = htmlBody + endTag;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailwithImage: " + ex.ToString());
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
                mailItem.Subject = "POD Achievement";

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                //string[] strValue = new string[14];

                string strDate = "", strLastDate = "";
                int iDateRow = dtDate.Rows.Count;
                for (int i = 0; i < iDateRow; i++)
                {
                    if (i == iDateRow - 1)
                    {
                        strLastDate = $"<th bgcolor = '#00ced1' style = 'color:#ffffff' align = 'center' width = '80' rowspan='2' >%</th >";
                    }
                    else
                    {
                        strDate += $"<th bgcolor = '#cc66ff' style = 'color:#ffffff' align = 'center' width = '80' >{dtDate.Rows[i]["YMD"]}</th >";
                    }
                }
                foreach (DataRow row in dtDate.Rows)
                {
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
                            strLastDate +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' colspan='3' >Production</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' colspan='3' >POD</th>" +
                            "<th bgcolor='#00ced1' style='color:#ffffff' align='center' colspan='6' >Staffing Ratio(%)</th>" +
                            "<th bgcolor='#ff9900' style='color:#ffffff' align='center' rowspan='2' width = '80'>Next TO</th>" +
                            "<th bgcolor='#ff9900' style='color:#ffffff' align='center' rowspan='2' width = '80'>Staffing Ratio(%)</th>" +
                        "</tr>" +
                        "<tr>" +
                            "<th bgcolor='#ff9966' style='color:#ffffff' align='center' width = '80'>Plan</th>" +
                            "<th bgcolor='#ff9966' style='color:#ffffff' align='center' width = '80'>Actual</th>" +
                            "<th bgcolor='#0066ff' style='color:#ffffff' align='center' width = '80'>%</th>" +
                            "<th bgcolor='#ff9999' style='color:#ffffff' align='center' width = '80'>Plan</th>" +
                            "<th bgcolor='#ff9999' style='color:#ffffff' align='center' width = '80'>Actual</th>" +
                            "<th bgcolor='#0066ff' style='color:#ffffff' align='center' width = '80'>%</th>" +
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

                    string ProdPlan = dtData.Rows[iRow]["PLAN_QTY"].ToString();
                    string ProdActual = dtData.Rows[iRow]["ACTUAL_QTY"].ToString();
                    string ProdRate = dtData.Rows[iRow]["PROD_RATE"].ToString();

                    string PodPlan = dtData.Rows[iRow]["POD_PLAN"].ToString();
                    string PodActual = dtData.Rows[iRow]["POD_ACTUAL"].ToString();
                    string PodRate = dtData.Rows[iRow]["POD_RATE"].ToString();

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

                    string ProdBgColor = dtData.Rows[iRow]["BG_COLOR_PROD"].ToString();
                    string ProdForeColor = dtData.Rows[iRow]["FORE_COLOR_PROD"].ToString();

                    string PodBgColor = dtData.Rows[iRow]["BG_COLOR_POD"].ToString();
                    string PodForeColor = dtData.Rows[iRow]["FORE_COLOR_POD"].ToString();

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
                                $"<td bgcolor='{Day6BgColor}'  style='color:{Day6ForeColor}'  align='right'>{Day6}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{ProdPlan}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{ProdActual}</td>" +
                                $"<td bgcolor='{ProdBgColor}'  style='color:{ProdForeColor}'  align='right'>{ProdRate}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{PodPlan}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{PodActual}</td>" +
                                $"<td bgcolor='{PodBgColor}'   style='color:{PodForeColor}'   align='right'>{PodRate}</td>" +
                                $"<td bgcolor='{Day1BgColor}'  style='color:{Day1ForeColor}'  align='right'>{Day1}</td>" +
                                $"<td bgcolor='{Day2BgColor}'  style='color:{Day2ForeColor}'  align='right'>{Day2}</td>" +
                                $"<td bgcolor='{Day3BgColor}'  style='color:{Day3ForeColor}'  align='right'>{Day3}</td>" +
                                $"<td bgcolor='{Day4BgColor}'  style='color:{Day4ForeColor}'  align='right'>{Day4}</td>" +
                                $"<td bgcolor='{Day5BgColor}'  style='color:{Day5ForeColor}'  align='right'>{Day5}</td>" +
                                $"<td bgcolor='{DAvgBgColor}'  style='color:{DAvgForeColor}'  align='right'>{DAvg}</td>" +
                                $"<td bgcolor='{TotalBgColor}' style='color:{TotalForeColor}' align='right'>{NextTo}</td>" +
                                $"<td bgcolor='{NextBgColor}'  style='color:{NextForeColor}' align='right'>{Next}</td>" +
                            "</tr>";
                }

                string imgInfo = "imgInfo", imgInfo2 = "imgInfo2";
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\POD_ko.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\POD_vi.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                // mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'></body>", imgInfo) ;

                //string explain = String.Format(dtDate.Rows[0]["TEXT1"].ToString(), imgInfo, imgInfo2);

                string html = "<body>" + String.Format(dtDate.Rows[0]["TEXT1"].ToString(), imgInfo, imgInfo2) + "</ body > " +
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
                string process_name = "P_SEND_EMAIL_TO_PO_V3";
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

        #endregion Email TO/PO

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
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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

        #endregion Email TO/PO IE

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
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string process_name = "P_SEND_EMAIL_ESCAN_V2";

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

        #endregion Email E-SCAN

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
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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

        #endregion Email Production

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
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string style = "<style> " +
                               "   .tblBoder { " +
                               "             font-family: 'Times New Roman', Times, serif; " +
                               "             font-style: italic; " +
                               "   } " +
                               "   .tblBoder td, .tblBoder th { " +
                               "         border: 0px; " +
                               "         padding: 3px 2px; " +
                               "         white-space: nowrap; " +
                               "         border: 1px solid #c0c0c0; " +
                               "   } " +
                               "   .tblBoder tbody td { " +
                               "             font-size: 20px; " +
                               "         } " +
                               "   .tblBoder thead { " +
                               "         background: #26A1B2; " +
                               "         font-style: italic; " +
                               "         border-bottom: 0px solid #444444; " +
                               "   } " +
                               "   .tblBoder thead th { " +
                               "             font-size: 19px; " +
                               "             font-weight: bold; " +
                               "         color: #F0F0F0; " +
                               "     background: #26A1B2; " +
                               "     text-align: center; " +
                               "         } " +
                               "   .green{ " +
                               "         background: green; " +
                               "         color: white; " +
                               "         } " +
                               "   .yellow{ " +
                               "         background: yellow; " +
                               "         color: black; " +
                               "         } " +
                               "   .red{ " +
                               "         background: red; " +
                               "         color: white; " +
                               "         } " +
                               "   .orange{ " +
                               "         background: orange; " +
                               "         color: white; " +
                               "         } " +
                               "</style> ";
                string text = "<table class='tblBoder'> " +
                              "  <tr> " +
                              "    <td></ td >" +
                              "    <td class='green'  align ='center'>Green</td>" +
                              "    <td class='yellow'  align ='center'>Yellow</td>" +
                              "    <td class='orange'  align ='center'>Orange</td>" +
                              "    <td class='red'  align ='center'>Red</td>" +
                              "  </tr>" +
                              "  <tr>" +
                              "    <td>Total Downtime per Line</td>" +
                              "    <td>&lt;9 minutes</td>" +
                              "    <td>9 ~ 25 minutes</td>" +
                              "    <td>25 ~ 50 minutes</td>" +
                              "    <td>Over 50 minutes</td>" +
                              "  </tr>" +
                              "  <tr>" +
                              "      <td>Total average measure</td>" +
                              "      <td>&lt;1.5 minutes</td>" +
                              "      <td>1.5 ~ 04 minutes</td>" +
                              "      <td>04 ~ 09 minutes</td>" +
                              "      <td>Over 09 minutes</td>" +
                              "  </tr>" +
                              "  <tr>" +
                              "      <td>Downtime by line</td>" +
                              "      <td>&lt;9 minutes</td>" +
                              "      <td>9 ~ 25 minutes</td>" +
                              "      <td>25 ~ 50 minutes</td>" +
                              "      <td>Over 50 minutes</td>" +
                              "  </tr>" +
                              "  <tr>" +
                              "      <td>Downtime average by line</td>" +
                              "      <td>&lt;1.5 minutes</td>" +
                              "      <td>1.5 ~ 04 minutes</td>" +
                              "      <td>04 ~ 09 minutes</td>" +
                              "      <td>Over 09 minutes</td>" +
                              "  </tr>       " +
                              "</table>";
                //string text = "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;' >" +
                //                    "Total Downtime per Line & Downtime by line = under 10 minutes is <b style='color:green'>green </b> " +
                //                    "and from 10 min to 29:59 is <b style='background-color:black; color:yellow'>yellow</b> " +
                //                    "and from 30 min to 59:59 min is <b style='color:orange'>orange</b> " +
                //                    "and then more than 1 hour is <b style='color:red'>red</b>" +
                //               "</p>" +
                //              "<p style='font-family:Times New Roman; font-size:18px; font-style:Italic;'>" +
                //                    "Total average measure & Downtime average by line = under 2 minutes is <b style='color:green'>green </b> " +
                //                    "and from 2 min to 4:59 is <b style='background-color:black; color:yellow'>yellow</b> " +
                //                    "and from 5 min to 09:59 is <b style='color:orange'>orange</b> " +
                //                    "and then more than 10 min is <b style='color:red'>red</b>" +
                //              "</p>"
                //              ;

                string html = "<html>" +
                                "<head>" +
                                    style +
                                "</head>" +
                                "<body>" +
                                     text +
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
                                    "</table>" +
                               "</body>" +
                            "<html>"
                              ;

                mailItem.HTMLBody = html;
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailAndon: " + ex.ToString());
            }
        }



        private string GetHtmlBodyRework()
        {
            try
            {

                return "";
            }
            catch (Exception ex)
            {
                WriteLog("GetHtmlBodyRework: " + ex.ToString());
                return "";
            }
        }
        #region Rework Monthly
        private DataSet SEL_REWORK_MONTHLY_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            MyOraDB.ShowErr = true;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_QUALITY_MONTHLY";
                MyOraDB.ReDim_Parameter(8);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_TIME";
                MyOraDB.Parameter_Name[2] = "CV_DATA1";
                MyOraDB.Parameter_Name[3] = "CV_DATA2";
                MyOraDB.Parameter_Name[4] = "CV_DATA3";
                MyOraDB.Parameter_Name[5] = "CV_DATA4";
                MyOraDB.Parameter_Name[6] = "CV_SUBJECT";
                MyOraDB.Parameter_Name[7] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[7] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "2";
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";
                MyOraDB.Parameter_Values[6] = "";
                MyOraDB.Parameter_Values[7] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        // WriteLog("P_SEND_EMAIL_NPI: null");
                    }
                    return null;
                }
                return ds_ret;
            }
            catch (Exception ex)
            {
                // WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
        #endregion

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

        #endregion Email ANDON

        #region Email Mold Monthly

        private void RunMoldRepairMonth(string argType)
        {
            DataSet ds = SEL_LOAD_MOLD_DATA(argType);
            if (ds == null || ds.Tables.Count == 0) return;
            DataTable dtData = ds.Tables[0];
            if (dtData == null || dtData.Rows.Count == 0) return;

            string subject = ds.Tables[2].Rows[0]["SUBJECT"].ToString();
            DataTable dtEmail = ds.Tables[3];
            WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunMoldRepairMonth({argType}): BEGIN");
            using (Mold_Repair_Monthly frmMold = new Mold_Repair_Monthly())
            {
                frmMold._dt1 = dtData;
                frmMold._dt2 = ds.Tables[1];
                frmMold._dt3 = ds.Tables[4];
                frmMold.Show();
                frmMold.SendToBack();
                CreateMailMoldMonth(subject, "", dtEmail);
            }

            //if (LoadDataMold(dtData, dtData2))
            //{
            //    CaptureControl(pnMold, "MoldChart");
            //    CaptureControl(grdMain, "MoldGrid");
            //    CreateMailMoldMonth(subject, "", dtEmail);

            //}

            WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunMoldRepairMonth({argType}): END");
        }

        private bool LoadDataMold(DataTable argDt, DataTable argDt2)
        {
            try
            {
                DataTable dt = argDt;
                DataTable dtYMD = dt.AsEnumerable().Where(r => r.Field<string>("IS_YMD") == "Y").OrderBy(r => r.Field<string>("WORK_YMD")).CopyToDataTable();
                DataView view = new DataView(dtYMD);
                DataTable distinctValues = view.ToTable(true, "WORK_YMD");
                InitBandHeader(distinctValues);
                dt.Columns.Remove(dt.Columns["IS_YMD"]);
                DataTable dtPivot = Pivot(dt, dt.Columns["WORK_YMD"], dt.Columns["MOLD_RP_QTY"]);
                grdMain.DataSource = dtPivot;

                //SetData(grdMain, dtPivot);
                FormatGrid(grdView);
                BindingChart(dtPivot);
                SetDataChart(argDt2);

                setChartRound(chartControl2, GetDataTemp(argDt2, "PU"));
                setChartRound(chartControl3, GetDataTemp(argDt2, "IP"));
                setChartRound(chartControl4, GetDataTemp(argDt2, "DMP"));
                setChartRound(chartControl5, GetDataTemp(argDt2, "Outsole"));
                setChartRound(chartControl6, GetDataTemp(argDt2, "Phylon"));
                setChartRound(chartControl7, GetDataTemp(argDt2, "CMP"));
                // SetTreelist(argDt2);
                return true;
            }
            catch (Exception ex)
            {
                WriteLog($"  LoadDataMold: {ex.Message}");
                return false;
            }

        }

        private void SetDataChart(DataTable argDt)
        {
            try
            {
                chartControl1.DataSource = argDt;
                for (int i = 0; i < 5; i++)
                {
                    chartControl1.Series[i].ArgumentDataMember = "WORKSHOP";
                    chartControl1.Series[i].ValueDataMembers.AddRange(new string[] { "ERR" + (i + 1).ToString() });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }

        }

        private void setChartRound(ChartControl argChart, DataTable argData)
        {
            try
            {
                argChart.DataSource = argData;
                argChart.Series[0].ArgumentDataMember = "ERR_NM";
                argChart.Series[0].ValueDataMembers.AddRange(new string[] { "ERR" });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }

        }

        private void SetTreelist(DataTable argDt)
        {
            try
            {

                //DataTable dt = GetDataTemp(argDt);
                //gridControlEx1.DataSource = dt;



                //DataTable dt = GetDataTemp(argDt);
                //tlsLoction.DataSource = dt;
                //tlsLoction.KeyFieldName = "ID";
                //tlsLoction.ParentFieldName = "PARENTID";
                //tlsLoction.Columns["ID_NAME"].Visible = false;
                //Skin skin = GridSkins.GetSkin(tlsLoction.LookAndFeel);
                //skin.Properties[GridSkins.OptShowTreeLine] = true;

                //foreach (TreeListNode node in tlsLoction.Nodes)
                //{
                //    var dataRow = tlsLoction.GetDataRecordByNode(node);
                //    node.Tag = dataRow;
                //    string nodeId = node.GetValue("ID").ToString();
                //    node.Checked = true;
                //    node.Expanded = true;
                //    foreach (TreeListNode node1 in node.RootNode.Nodes)
                //    {

                //        node1.Checked = true;
                //    }
                //}
            }
            catch (Exception ex)
            { }
        }


        private DataTable GetDataTemp(DataTable argDt, string argLocation)
        {
            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("ERR_NM");
            dt.Columns.Add("ERR", typeof(int));
            try
            {


                DataTable dtData = argDt.Select($"WORKSHOP = '{argLocation}'").CopyToDataTable();

                for (int i = 1; i <= 5; i++)
                {
                    string errName = dtData.Rows[0]["ERR_NM" + i].ToString();
                    if (errName == "") continue;
                    DataRow row = dt.NewRow();
                    row["ERR_NM"] = dtData.Rows[0]["ERR_NM" + i].ToString();
                    row["ERR"] = int.Parse(dtData.Rows[0]["ERR" + i].ToString());
                    dt.Rows.Add(row);
                }


                /*
                DataTable dt = new DataTable();
                dt.Clear();
                dt.Columns.Add("WORKSHOP");
                dt.Columns.Add("ERR_NM");
                dt.Columns.Add("ERR");
                foreach (DataRow dataRow in argDt.Rows)
                {
                    for(int i = 1; i < 6; i++)
                    {
                        if (dataRow["ERR_NM" + i.ToString()].ToString() == "") continue;
                        DataRow row = dt.NewRow();
                        row["WORKSHOP"] = dataRow["WORKSHOP"].ToString();
                        row["ERR_NM"] = dataRow["ERR_NM" + i.ToString()].ToString();
                        row["ERR"] = dataRow["ERR" + i.ToString()].ToString();
                        dt.Rows.Add(row);
                    }                    
                }
                */
                /*
                DataTable dt = new DataTable();
                dt.Clear();
                dt.Columns.Add("PARENTID");
                dt.Columns.Add("ID");
                dt.Columns.Add("MENU_NM");
                foreach (DataRow dataRow in argDt.Rows)
                {
                    foreach(DataColumn dataColumn in argDt.Columns)
                    {
                        if (dataColumn.ColumnName.Contains("ERR_NM"))
                        {
                            if (dataRow[dataColumn].ToString() == "") continue;
                            DataRow row = dt.NewRow();
                            row["PARENTID"] = dataRow["WORKSHOP"].ToString();
                            row["ID"] = dataRow["WORKSHOP"].ToString() + "|" + dataRow[dataColumn].ToString();
                            row["MENU_NM"] = dataRow[dataColumn].ToString();
                            dt.Rows.Add(row);
                        }

                    }
                }*/

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());

            }
            return dt;
        }

        private void InitBandHeader(DataTable dt)
        {
            try
            {
                grdView.Bands.Clear();
                grdView.Columns.Clear();
                GridBand gridBandBottom = new GridBand();
                GridBand gridBandTotalMold = new GridBand();
                GridBand gridBandMonth = new GridBand();

                //2 band cuối
                GridBand gridBandAvgMold = new GridBand();
                GridBand gridBandPerMold = new GridBand();
                // 
                // gridBandBottom
                // 
                gridBandBottom.AppearanceHeader.Options.UseTextOptions = true;
                gridBandBottom.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridBandBottom.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                gridBandBottom.Caption = "Location";
                gridBandBottom.Name = "gridBandBottom";
                gridBandBottom.RowCount = 2;
                //  gridBandBottom.VisibleIndex = 0;
                gridBandBottom.Width = 50;
                // 
                // gridBandTotalMold
                // 
                gridBandTotalMold.AppearanceHeader.Options.UseTextOptions = true;
                gridBandTotalMold.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridBandTotalMold.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                gridBandTotalMold.Caption = "Total\nMold";
                gridBandTotalMold.Name = "gridBandTotalMold";
                // gridBandTotalMold.VisibleIndex = 1;
                gridBandTotalMold.Width = 150;

                //2 band cuối
                // 
                // gridBandavgMold
                // 
                gridBandAvgMold.AppearanceHeader.Options.UseTextOptions = true;
                gridBandAvgMold.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridBandAvgMold.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                gridBandAvgMold.Caption = "Averange\nMold";
                gridBandAvgMold.Name = "gridBandAvgMold";
                // gridBandAvgMold.VisibleIndex = 50;
                gridBandAvgMold.Width = 150;

                // 
                // gridBandPerMold
                // 
                gridBandPerMold.AppearanceHeader.Options.UseTextOptions = true;
                gridBandPerMold.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridBandPerMold.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                gridBandPerMold.Caption = "Repair\nRatio";

                gridBandPerMold.Name = "gridBandPerMold";
                //  gridBandPerMold.VisibleIndex = 51;
                gridBandPerMold.Width = 70;

                BandedGridColumn WORK_BOTTOM = new BandedGridColumn();
                BandedGridColumn WORK_PLACE_NM = new BandedGridColumn();
                BandedGridColumn TOTAL_MOLD_RP = new BandedGridColumn();

                //2 COLUMNS LAST
                BandedGridColumn AVG_MOLD = new BandedGridColumn();
                BandedGridColumn PER_MOLD = new BandedGridColumn();
                // 
                // WORK_BOTTOM
                // 
                WORK_BOTTOM.Caption = "WORK_BOTTOM";
                WORK_BOTTOM.FieldName = "WORK_BOTTOM";
                WORK_BOTTOM.Name = "WORK_BOTTOM";
                WORK_BOTTOM.Visible = true;
                // 
                // WORK_PLACE_NM
                // 
                WORK_PLACE_NM.Caption = "WORK_PLACE_NM";
                WORK_PLACE_NM.FieldName = "WORK_PLACE_NM";
                WORK_PLACE_NM.Name = "WORK_PLACE_NM";
                WORK_PLACE_NM.Visible = true;

                // 
                // TOTAL_MOLD_RP
                // 
                TOTAL_MOLD_RP.Caption = "TOTAL_MOLD_RP";
                TOTAL_MOLD_RP.FieldName = "TOTAL_MOLD_RP";
                TOTAL_MOLD_RP.Name = "TOTAL_MOLD_RP";
                TOTAL_MOLD_RP.Visible = true;

                // 
                // AVG_MOLD
                // 
                AVG_MOLD.Caption = "AVG_MOLD";
                AVG_MOLD.FieldName = "AVG_MOLD_RP";
                AVG_MOLD.Name = "AVG_MOLD";
                AVG_MOLD.Visible = true;

                // 
                // PER_MOLD
                // 
                PER_MOLD.Caption = "PER_MOLD";
                PER_MOLD.FieldName = "PER_MOLD_RP";
                PER_MOLD.Name = "PER_MOLD";
                PER_MOLD.Visible = true;

                gridBandBottom.Columns.Add(WORK_BOTTOM);
                gridBandBottom.Columns.Add(WORK_PLACE_NM);
                gridBandTotalMold.Columns.Add(TOTAL_MOLD_RP);

                gridBandAvgMold.Columns.Add(AVG_MOLD);
                gridBandPerMold.Columns.Add(PER_MOLD);

                grdView.Bands.AddRange(new GridBand[] { gridBandBottom, gridBandTotalMold, gridBandMonth, gridBandAvgMold, gridBandPerMold });
                grdView.Columns.AddRange(new BandedGridColumn[] {
                   WORK_BOTTOM,
                   WORK_PLACE_NM,
                   TOTAL_MOLD_RP,AVG_MOLD,PER_MOLD});
                // 
                // gridBandMonth
                // 
                string date = DateTime.Now.AddMonths(-1).ToString("MMM-yyyy");
                gridBandMonth.AppearanceHeader.Options.UseTextOptions = true;
                gridBandMonth.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                gridBandMonth.AppearanceHeader.TextOptions.VAlignment = VertAlignment.Center;
                gridBandMonth.Caption = date;
                gridBandMonth.Name = "gridBandMonth";
                gridBandMonth.VisibleIndex = 2;

                for (int i = 0; i < dt.Rows.Count; i++)
                {

                    GridBand gridbandDays = new GridBand();
                    // 
                    // gridbandDays
                    // 
                    string Days = dt.Rows[i]["WORK_YMD"].ToString();
                    string CaptionOfDays = Days.Substring(Days.Length - 2, 2);
                    gridbandDays.AppearanceHeader.Options.UseTextOptions = true;

                    gridbandDays.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridbandDays.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridbandDays.Caption = CaptionOfDays;
                    gridbandDays.Name = Days;
                    gridbandDays.VisibleIndex = i + 3; //Từ column 3 trở đi
                    gridBandMonth.Children.AddRange(new GridBand[] { gridbandDays });

                    BandedGridColumn ColumnsDays = new BandedGridColumn();
                    // 
                    // ColumnsDays
                    // 
                    ColumnsDays.Caption = Days;
                    ColumnsDays.FieldName = Days;
                    ColumnsDays.Name = Days;
                    ColumnsDays.Visible = true;
                    ColumnsDays.Width = 43;

                    gridbandDays.Columns.Add(ColumnsDays);
                    grdView.Columns.AddRange(new BandedGridColumn[] { ColumnsDays });
                }


            }
            catch (Exception ex)
            {

                throw;
            }
        }
        private void FormatGrid(BandedGridView grid)
        {
            try
            {
                // grdMain.Font = new Font("Calibri", 15, FontStyle.Bold);
                grdView.OptionsView.AllowCellMerge = true;
                grdView.BandPanelRowHeight = 30;

                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    if (grid.Columns[i].OwnerBand.ParentBand != null)
                    {
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Bold);
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.BackColor = Color.Orange;
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.ForeColor = Color.White;

                    }
                    grid.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Bold);
                    grid.Columns[i].AppearanceCell.Font = new Font("Calibri", 12, FontStyle.Regular);
                    grid.Columns[i].OwnerBand.AppearanceHeader.BackColor = Color.DodgerBlue;
                    grid.Columns[i].OwnerBand.AppearanceHeader.ForeColor = Color.White;
                    if (i <= 1)
                    {
                        grid.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.True;
                    }
                    else
                    {

                        grid.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.False;
                        grid.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                        grid.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                        grid.Columns[i].DisplayFormat.FormatString = "#,0.##";
                    }
                }
            }
            catch
            {

            }

        }


        //private void InitBandHeader(DataTable dt)
        //{
        //    try
        //    {

        //        grdView.Bands.Clear();
        //        grdView.Columns.Clear();

        //        GridBand gridBandBottom = new GridBand();
        //        GridBand gridBandTotalMold = new GridBand();
        //        GridBand gridBandMonth = new GridBand();

        //        //2 band cuối
        //        GridBand gridBandAvgMold = new GridBand();
        //        GridBand gridBandPerMold = new GridBand();
        //        // 
        //        // gridBandBottom
        //        // 
        //       // gridBandBottom.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Bold);
        //        gridBandBottom.AppearanceHeader.Options.UseTextOptions = true;
        //        gridBandBottom.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        //        gridBandBottom.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
        //        gridBandBottom.Caption = "Location";
        //        gridBandBottom.Name = "gridBandBottom";
        //        gridBandBottom.RowCount = 2;
        //        gridBandBottom.VisibleIndex = 0;
        //        gridBandBottom.Width = 150;
        //        // 
        //        // gridBandTotalMold
        //        // 
        //        gridBandTotalMold.AppearanceHeader.Options.UseTextOptions = true;
        //      //  gridBandTotalMold.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Bold);
        //        gridBandTotalMold.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        //        gridBandTotalMold.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
        //        gridBandTotalMold.Caption = "Total Mold";
        //        gridBandTotalMold.Name = "gridBandTotalMold";
        //        gridBandTotalMold.VisibleIndex = 1;
        //        gridBandTotalMold.Width = 100;

        //        //2 band cuối
        //        // 
        //        // gridBandavgMold
        //        // 
        //        gridBandAvgMold.AppearanceHeader.Options.UseTextOptions = true;
        //        gridBandAvgMold.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        //        gridBandAvgMold.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
        //        gridBandAvgMold.Caption = "Average\nMold";
        //        gridBandAvgMold.Name = "gridBandAvgMold";
        //        gridBandAvgMold.VisibleIndex = 50;
        //        gridBandAvgMold.Width = 100;

        //        // 
        //        // gridBandPerMold
        //        // 
        //        gridBandPerMold.AppearanceHeader.Options.UseTextOptions = true;
        //        gridBandPerMold.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        //        gridBandPerMold.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
        //        gridBandPerMold.Caption = "Repair\nRatio";

        //        gridBandPerMold.Name = "gridBandPerMold";
        //        gridBandPerMold.VisibleIndex = 51;
        //        gridBandPerMold.Width = 75;

        //        BandedGridColumn WORK_BOTTOM = new BandedGridColumn();
        //        BandedGridColumn WORK_PLACE_NM = new BandedGridColumn();
        //        BandedGridColumn TOTAL_MOLD_RP = new BandedGridColumn();

        //        //2 COLUMNS LAST
        //        BandedGridColumn AVG_MOLD = new BandedGridColumn();
        //        BandedGridColumn PER_MOLD = new BandedGridColumn();
        //        // 
        //        // WORK_BOTTOM
        //        // 
        //        WORK_BOTTOM.Caption = "WORK_BOTTOM";
        //        WORK_BOTTOM.FieldName = "WORK_BOTTOM";
        //        WORK_BOTTOM.Name = "WORK_BOTTOM";
        //        WORK_BOTTOM.Visible = true;


        //        // 
        //        // WORK_PLACE_NM
        //        // 
        //        WORK_PLACE_NM.Caption = "WORK_PLACE_NM";
        //        WORK_PLACE_NM.FieldName = "WORK_PLACE_NM";
        //        WORK_PLACE_NM.Name = "WORK_PLACE_NM";
        //        WORK_PLACE_NM.Visible = true;


        //        // 
        //        // TOTAL_MOLD_RP
        //        // 
        //        TOTAL_MOLD_RP.Caption = "TOTAL_MOLD_RP";
        //        TOTAL_MOLD_RP.FieldName = "TOTAL_MOLD_RP";
        //        TOTAL_MOLD_RP.Name = "TOTAL_MOLD_RP";
        //        TOTAL_MOLD_RP.Visible = true;
        //        TOTAL_MOLD_RP.Width = 100;

        //        // 
        //        // AVG_MOLD
        //        // 
        //        AVG_MOLD.Caption = "AVG_MOLD";
        //        AVG_MOLD.FieldName = "AVG_MOLD_RP";
        //        AVG_MOLD.Name = "AVG_MOLD";
        //        AVG_MOLD.Visible = true;
        //        AVG_MOLD.Width = 100;


        //        // 
        //        // PER_MOLD
        //        // 
        //        PER_MOLD.Caption = "PER_MOLD";
        //        PER_MOLD.FieldName = "PER_MOLD_RP";
        //        PER_MOLD.Name = "PER_MOLD";
        //        PER_MOLD.Visible = true;


        //        gridBandBottom.Columns.Add(WORK_BOTTOM);
        //        gridBandBottom.Columns.Add(WORK_PLACE_NM);
        //        gridBandTotalMold.Columns.Add(TOTAL_MOLD_RP);

        //        gridBandAvgMold.Columns.Add(AVG_MOLD);
        //        gridBandPerMold.Columns.Add(PER_MOLD);


        //        grdView.Bands.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { gridBandBottom, gridBandTotalMold, gridBandMonth, gridBandAvgMold, gridBandPerMold });
        //        grdView.Columns.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn[] {
        //           WORK_BOTTOM,
        //           WORK_PLACE_NM,
        //           TOTAL_MOLD_RP,AVG_MOLD,PER_MOLD});

        //        // 
        //        // gridBandMonth
        //        // 
        //        string date = DateTime.Now.AddMonths(-1).ToString("MMM-yyyy");
        //        gridBandMonth.AppearanceHeader.Options.UseTextOptions = true;
        //        gridBandMonth.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        //        gridBandMonth.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
        //        gridBandMonth.Caption = date;
        //        gridBandMonth.Name = "gridBandMonth";
        //        gridBandMonth.VisibleIndex = 2;

        //        for (int i = 0; i < dt.Rows.Count; i++)
        //        {

        //            GridBand gridbandDays = new GridBand();
        //            // 
        //            // gridbandDays
        //            // 
        //            string Days = dt.Rows[i]["WORK_YMD"].ToString();
        //            string CaptionOfDays = Days.Substring(Days.Length - 2, 2);
        //            gridbandDays.AppearanceHeader.Options.UseTextOptions = true;
        //           // gridbandDays.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Bold);

        //            gridbandDays.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        //            gridbandDays.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
        //            gridbandDays.Caption = CaptionOfDays;
        //            gridbandDays.Name = Days;
        //            gridbandDays.VisibleIndex = i + 3; //Từ column 3 trở đi
        //            gridBandMonth.Children.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { gridbandDays });

        //            BandedGridColumn ColumnsDays = new BandedGridColumn();
        //            // 
        //            // ColumnsDays
        //            // 

        //            ColumnsDays.Caption = Days;
        //            ColumnsDays.FieldName = Days;
        //            ColumnsDays.Name = Days;
        //            ColumnsDays.Width = 50;
        //            ColumnsDays.Visible = true;



        //            gridbandDays.Columns.Add(ColumnsDays);
        //            grdView.Columns.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn[] { ColumnsDays });
        //        }


        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //}

        //private void FormatGrid(BandedGridView grid)
        //{
        //    try
        //    {
        //       // grdMain.Font = new Font("Calibri", 15, FontStyle.Bold);
        //        grdView.OptionsView.AllowCellMerge = true;
        //        grdView.BandPanelRowHeight = 30;

        //        for (int i = 0; i < grid.Columns.Count; i++)
        //        {
        //            if (grid.Columns[i].OwnerBand.ParentBand != null)
        //            {
        //                grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Bold);
        //                grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.BackColor = Color.Orange;
        //                grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.ForeColor = Color.White;

        //            }
        //            grid.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Bold);
        //            grid.Columns[i].AppearanceCell.Font = new Font("Calibri", 12, FontStyle.Regular);
        //            grid.Columns[i].OwnerBand.AppearanceHeader.BackColor = Color.DodgerBlue;
        //            grid.Columns[i].OwnerBand.AppearanceHeader.ForeColor = Color.White;
        //            if (i <= 1)
        //            {
        //                grid.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.True;
        //            }
        //            else
        //            {

        //                grid.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.False;
        //                grid.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
        //                grid.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
        //                grid.Columns[i].DisplayFormat.FormatString = "#,0.##";
        //            }
        //        }
        //    }
        //    catch
        //    {

        //    }

        //}

        private void grdView_CustomDrawBandHeader(object sender, BandHeaderCustomDrawEventArgs e)
        {
            if (e.Band == null) return;
            if (e.Band.AppearanceHeader.BackColor != Color.Empty)
                e.Info.AllowColoring = true;


        }

        private void BindingChart(DataTable dt)
        {
            try
            {
                chartMold.DataSource = dt;
                chartMold.Series[0].ArgumentDataMember = "WORK_PLACE_NM";
                chartMold.Series[0].ValueDataMembers.AddRange(new string[] { "PER_MOLD_RP" });
            }
            catch
            {

            }
        }


        private void CreateMailMoldMonth(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\MoldChart.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\MoldGrid.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                // Outlook.Attachment oAttachPic3 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\MoldGrid2.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo", imgInfo2 = "imgInfo2", imgInfo3 = "imgInfo3";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                // oAttachPic3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo3);
                mailItem.HTMLBody = String.Format(@"<img src='cid:{0}'><br>
                                                    <img src='cid:{1}'>"
                                                , imgInfo, imgInfo2);

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        private DataSet SEL_LOAD_MOLD_DATA(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
                string process_name = "P_EMAIL_MOLD_REPAIR_MONTH_V2";
                // MyOraDB.ShowErr = true;
                MyOraDB.ReDim_Parameter(7);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_DATA";
                MyOraDB.Parameter_Name[3] = "CV_COL";
                MyOraDB.Parameter_Name[4] = "CV_SUBJECT";
                MyOraDB.Parameter_Name[5] = "CV_EMAIL";
                MyOraDB.Parameter_Name[6] = "CV_DATA2";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
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

        #endregion

        #region Email Mold Monthly By Warehouse
        private void RunMoldRepairMonthWh(string argType)
        {
            DataSet ds = SEL_LOAD_MOLD_DATA_WH(argType);
            if (ds == null || ds.Tables.Count == 0) return;
            DataTable dtData = ds.Tables[0];
            if (dtData == null || dtData.Rows.Count == 0) return;

            WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunMoldRepairMonth({argType}): BEGIN");
            using (Mold_Repair_Monthly_WH frmMold = new Mold_Repair_Monthly_WH())
            {
                frmMold._chkTest = chkTest.Checked;
                frmMold._subject = ds.Tables[2].Rows[0]["SUBJECT"].ToString();
                frmMold._dt1 = dtData;
                frmMold._dt2 = ds.Tables[1];
                frmMold._dt3 = ds.Tables[4];
                frmMold._dtEmail = ds.Tables[3];
                frmMold.Show();
                frmMold.SendToBack();
                // CreateMailMoldMonthWh(subject, "", dtEmail);
            }

            //if (LoadDataMold(dtData, dtData2))
            //{
            //    CaptureControl(pnMold, "MoldChart");
            //    CaptureControl(grdMain, "MoldGrid");
            //    CreateMailMoldMonth(subject, "", dtEmail);

            //}

            WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunMoldRepairMonth({argType}): END");
        }

        private void CreateMailMoldMonthWh(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\MoldChartWh.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                // Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\MoldGrid.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                // Outlook.Attachment oAttachPic3 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\MoldGrid2.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo", imgInfo2 = "imgInfo2", imgInfo3 = "imgInfo3";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                //  oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                //   oAttachPic3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo3);
                mailItem.HTMLBody = String.Format(@"<img src='cid:{0}'>", imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailMoldMonthWh: " + ex.ToString());
            }
        }

        private DataSet SEL_LOAD_MOLD_DATA_WH(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
                string process_name = "P_EMAIL_MOLD_REPAIR_MONTH_WH";
                // MyOraDB.ShowErr = true;
                MyOraDB.ReDim_Parameter(7);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_DATA";
                MyOraDB.Parameter_Name[3] = "CV_COL";
                MyOraDB.Parameter_Name[4] = "CV_SUBJECT";
                MyOraDB.Parameter_Name[5] = "CV_EMAIL";
                MyOraDB.Parameter_Name[6] = "CV_DATA2";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
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


        private DataSet SEL_OS_PRESS_MONTHLY(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
                string process_name = "P_EMAIL_MOLD_REPAIR_MONTH_WH";
                // MyOraDB.ShowErr = true;
                MyOraDB.ReDim_Parameter(7);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_DATA";
                MyOraDB.Parameter_Name[3] = "CV_COL";
                MyOraDB.Parameter_Name[4] = "CV_SUBJECT";
                MyOraDB.Parameter_Name[5] = "CV_EMAIL";
                MyOraDB.Parameter_Name[6] = "CV_DATA2";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
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

        #endregion


        #region Email Bottom Inventory

        private void CreateMailBottomInventory()
        {
            try
            {
                //Outlook.MailItem mailItem = (Outlook.MailItem)
                // this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: Begin send email");
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttach = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Chart.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: Chart.png ok");

                Outlook.Attachment oAttachPicGrid1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid1.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: Grid1.png ok");

                Outlook.Attachment oAttachPicGrid2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid2.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: Grid2.png ok");

                Outlook.Attachment oAttachPicGrid3 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid3.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: Grid3.png ok");

                Outlook.Attachment oAttachPicGrid4 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Grid4.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: Grid4.png ok");

                mailItem.Subject = "Bottom Inventory Set Analysis";
                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
                {
                    WriteLog(DateTime.Now.ToString() + " Bottom Inventory: email " + dtEmail.Rows.Count.ToString());
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
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: send ok ");
            }
            catch (Exception ex)
            {
                WriteLog(DateTime.Now.ToString() + " Bottom Inventory: " + ex.ToString());
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

                int iBotSet, iFssSet, iFinishSole;

                // Grid
                DataTable dtSource = new DataTable();
                if (buildHeader_detail(dtSource, dtGrid))
                {
                    if (bindingDataSource_detail(dtSource, dtGrid))
                    {
                        int.TryParse(dtSource.Rows[0]["Total Inv"].ToString(), out iBotSet);
                        int.TryParse(dtSource.Rows[1]["Total Inv"].ToString(), out iFssSet);
                        int.TryParse(dtSource.Rows[2]["Total Inv"].ToString(), out iFinishSole);
                        lblBotSetQty.Text = "Bottom Set: " + iBotSet.ToString("#,#") + " (Prs)";
                        lblFssSetQty.Text = "Stockfit Imcoming Set: " + iFssSet.ToString("#,#") + " (Prs)";
                        lblFinishSoleQty.Text = "Finised Sole Inventory: " + iFinishSole.ToString("#,#") + " (Prs)";
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

        #endregion Set Data For Chart & Grid

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

        #endregion Email Bottom Inventory

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
                        int CurrentValue;
                        int.TryParse(htLine[HtKey].ToString(), out CurrentValue);
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

        #endregion Email Cutting

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
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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

        /// <summary>
        /// Gets a Inverted DataTable
        /// </summary>
        /// <param name="table">DataTable do invert</param>
        /// <param name="columnX">X Axis Column</param>
        /// <param name="nullValue">null Value to Complete the Pivot Table</param>
        /// <param name="columnsToIgnore">Columns that should be ignored in the pivot
        /// process (X Axis column is ignored by default)</param>
        /// <returns>C# Pivot Table Method  - Felipe Sabino</returns>
        public static DataTable GetInversedDataTable(DataTable table, string columnX,
                                                     params string[] columnsToIgnore)
        {
            try
            {
                //Create a DataTable to Return
                DataTable returnTable = new DataTable();

                if (columnX == "")
                    columnX = table.Columns[0].ColumnName;

                //Add a Column at the beginning of the table

                returnTable.Columns.Add(columnX);

                //Read all DISTINCT values from columnX Column in the provided DataTale
                List<string> columnXValues = new List<string>();

                //Creates list of columns to ignore
                List<string> listColumnsToIgnore = new List<string>();
                if (columnsToIgnore.Length > 0)
                    listColumnsToIgnore.AddRange(columnsToIgnore);

                if (!listColumnsToIgnore.Contains(columnX))
                    listColumnsToIgnore.Add(columnX);

                foreach (DataRow dr in table.Rows)
                {
                    string columnXTemp = dr[columnX].ToString();
                    //Verify if the value was already listed
                    if (!columnXValues.Contains(columnXTemp))
                    {
                        //if the value id different from others provided, add to the list of
                        //values and creates a new Column with its value.
                        columnXValues.Add(columnXTemp);
                        returnTable.Columns.Add(columnXTemp);
                    }
                    else
                    {
                        //Throw exception for a repeated value
                        throw new Exception("The inversion used must have " +
                                            "unique values for column " + columnX);
                    }
                }

                //Add a line for each column of the DataTable

                foreach (DataColumn dc in table.Columns)
                {
                    if (!columnXValues.Contains(dc.ColumnName) &&
                        !listColumnsToIgnore.Contains(dc.ColumnName))
                    {
                        DataRow dr = returnTable.NewRow();
                        dr[0] = dc.ColumnName;
                        returnTable.Rows.Add(dr);
                    }
                }

                //Complete the datatable with the values
                for (int i = 0; i < returnTable.Rows.Count; i++)
                {
                    for (int j = 1; j < returnTable.Columns.Count; j++)
                    {
                        returnTable.Rows[i][j] =
                          table.Rows[j - 1][returnTable.Rows[i][0].ToString()].ToString();
                    }
                }

                return returnTable;
            }
            catch (Exception ex) { return null; }
        }

        private DataTable Pivot(DataTable dt, DataColumn pivotColumn, DataColumn pivotValue)
        {
            try
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
            catch (Exception ex) { return null; }
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

        #endregion Email NPI Capture

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
                string HeaderRow1 = "", HeaderRow2 = "", HeaderRow3 = "";

                // int[] colWidth = { 55, 55, 55, 60, 66, 55, 70, 80, 70, 80, 65, 65, 60, 70, 65, 55, 75, 55 };

                int[] colWidth = { 55, 55, 63, 63, 63, 63, 63, 65, 80, 65, 80, 63, 63, 63, 63, 63, 63, 75, 63 };

                foreach (DataRow row in dtHeader.Rows)
                {
                    int.TryParse(row["NPI_CODE"].ToString(), out npiCode);
                    if (npiCode >= 10)
                    {
                        HeaderRow1 += $"<td bgcolor = '#e9f1fb' style = 'color:#000000' align = 'center' width = '{colWidth[i]}'>{row["NPI_DATE"]}</td>";
                        HeaderRow2 += $"<td bgcolor = '#e9f1fb' style = 'color:#000000' align = 'center' width = '{colWidth[i]}'>{row["NPI_NAME"]}</td>";
                        HeaderRow3 += $"<td bgcolor = '#e9f1fb' style = 'color:#000000' align = 'center' width = '{colWidth[i]}'>{row["VALUE3"]}</td>";
                        i++;
                    }
                    else
                        HeaderRow1 += $"<td bgcolor = '#e9f1fb' style = 'color:#000000' rowspan ='3' align = 'center'>{row["NPI_NAME"]}</td>";
                }

                TableHeader = "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow1 + "</tr> " +
                              "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow2 + "</tr> " +
                              "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow3 + "</tr> ";

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
                    string foreColor = dtData.Rows[iRowData]["FCOLOR"].ToString();
                    string trackingDate = dtData.Rows[iRowData]["TRACKING_DATE"].ToString();

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
                                $"<td bgcolor='{backColor}' style='color:{foreColor}' width = '50' align='center' >{trackingDate}</td>"
                              ;
                    }
                    else
                    {
                        TableRow += $"<td bgcolor='{backColor}' style='color:{foreColor}' width = '50' align='center'>{trackingDate}</td>";
                    }
                }

                return "<table style='font-family:Calibri; font-size:15px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' with = '5000' >" +
                            TableHeader + TableRow +
                       "</table>";
            }
            catch (Exception ex)
            {
                WriteLog("GetHtmlBodyNpi: " + ex.ToString());
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

        private DataSet SEL_DATA_TMS_DAAS_CHART(string WorkType, string DateF, string DateT, string LineCd)
        {
            System.Data.DataSet retDS;
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            MyOraDB.ReDim_Parameter(13);
            MyOraDB.Process_Name = "LMES.P_GMES0266_Q_2";
            MyOraDB.ShowErr = true;
            //  for (int i = 0; i < intParm; i++)

            MyOraDB.Parameter_Type[0] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[1] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[2] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[3] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[4] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[5] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[6] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[7] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[8] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[9] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[10] = (char)OracleType.VarChar;
            MyOraDB.Parameter_Type[11] = (char)OracleType.Cursor;
            MyOraDB.Parameter_Type[12] = (char)OracleType.Cursor;

            //V_P_TYPE,V_P_OPTION
            MyOraDB.Parameter_Name[0] = "V_P_WORK_TYPE";
            MyOraDB.Parameter_Name[1] = "V_P_DATEF";
            MyOraDB.Parameter_Name[2] = "V_P_DATET";
            MyOraDB.Parameter_Name[3] = "V_P_LINE_CD";
            MyOraDB.Parameter_Name[4] = "V_P_ERROR_CODE";
            MyOraDB.Parameter_Name[5] = "V_P_ROW_COUNT";
            MyOraDB.Parameter_Name[6] = "V_P_ERROR_NOTE";
            MyOraDB.Parameter_Name[7] = "V_P_RETURN_STR";
            MyOraDB.Parameter_Name[8] = "V_P_ERROR_STR";
            MyOraDB.Parameter_Name[9] = "V_ERRORSTATE";
            MyOraDB.Parameter_Name[10] = "V_ERRORPROCEDURE";
            MyOraDB.Parameter_Name[11] = "OUT_CURSOR";
            MyOraDB.Parameter_Name[12] = "OUT_CURSOR1";

            MyOraDB.Parameter_Values[0] = WorkType;
            MyOraDB.Parameter_Values[1] = DateF;
            MyOraDB.Parameter_Values[2] = DateT;
            MyOraDB.Parameter_Values[3] = LineCd;
            MyOraDB.Parameter_Values[4] = "";
            MyOraDB.Parameter_Values[5] = "";
            MyOraDB.Parameter_Values[6] = "";
            MyOraDB.Parameter_Values[7] = "";
            MyOraDB.Parameter_Values[8] = "";
            MyOraDB.Parameter_Values[9] = "";
            MyOraDB.Parameter_Values[10] = "";
            MyOraDB.Parameter_Values[11] = "";
            MyOraDB.Parameter_Values[12] = "";

            MyOraDB.Add_Select_Parameter(true);
            retDS = MyOraDB.Exe_Select_Procedure();
            if (retDS == null) return null;
            return retDS;
        }

        public DataSet SEL_SCADA_DATA(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.SEPHIROTH;

            DataSet ds_ret;
            try
            {
                string process_name = "MES.P_SEND_EMAIL_SCADA";
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
                        WriteLog("SEL_SCADA_DATA: null");
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

        public DataSet SEL_TMS_SUMMARY_DATA(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;

            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_TMS_ORDER_SUM";
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
                        WriteLog("SEL_TMS_SUM_DATA: null");
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

        #endregion Email NPI

        #region OS Red Machine

        private void RunOSRedMachine(string argType, string argDate, string argHH)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                OS_Red_Machine OsRedMachine = new OS_Red_Machine();

                string html = OsRedMachine.Html(argType, argDate, argHH);
                if (string.IsNullOrEmpty(html)) return;
                WriteLog("RunOSRedMachine: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailOs(OsRedMachine._subject, html, OsRedMachine._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        #endregion OS Red Machine

        #region Po Relief Manual Register Report

        private DataSet SEL_DATA_PO_REG_REPORT(string V_P_TYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_POTO_REG_V2";
                MyOraDB.ReDim_Parameter(3);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "CV_DATA";
                MyOraDB.Parameter_Name[2] = "CV_EMAIL";


                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;


                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    return null;
                }
                return ds_ret;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        #endregion

        #region BOTTOM MONTHLY INVENTORY ANALYSIS
        private DataTable SEL_DATA_MONTHLY_BOTTOM_ANALYSIS(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_BT_MONTHLY_INV";
                MyOraDB.ReDim_Parameter(4);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_DATA";
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
                        //WriteLog("P_SEND_EMAIL_NPI: null");
                    }
                    return null;
                }
                return ds_ret.Tables[0];
            }
            catch (Exception ex)
            {
                // WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
        #endregion

        #region OS Red Machine Monthly
        private DataTable SEL_DATA_WEEKLY_B_C(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.SEPHIROTH;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_WEEKLY_B_C";
                MyOraDB.ReDim_Parameter(3);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";


                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        //WriteLog("P_SEND_EMAIL_NPI: null");
                    }
                    return null;
                }
                return ds_ret.Tables[0];
            }
            catch (Exception ex)
            {
                // WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
        #endregion

        #region OS Red Machine Monthly
        private DataTable SEL_DATA_OS_MACHINE_MONTHLY(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.SEPHIROTH;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_OUTSOLE_MONTHLY";
                MyOraDB.ReDim_Parameter(3);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";


                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;


                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        //WriteLog("P_SEND_EMAIL_NPI: null");
                    }
                    return null;
                }
                return ds_ret.Tables[0];
            }
            catch (Exception ex)
            {
                // WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
        #endregion

        #region Run Feedback

        private void RunFeedback(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Feedback send_Feedback = new Send_Feedback();

                string html = send_Feedback.Html(argType);
                if (string.IsNullOrEmpty(html)) return;
                WriteLog("Run Feedback: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailFeedBack(send_Feedback._subject, html, send_Feedback._email);



                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        #endregion Run Feedback

        #region Mold Repair

        private void RunMoldRepair(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Mold_Repair mold_Repair = new Mold_Repair();
                string html = mold_Repair.Html_MoldRepair(argType);
                if (html == "") return;
                WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMail(mold_Repair._subject, html, mold_Repair._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        #endregion Mold Repair

        #region Budget

        private void RunBuget(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Budget budget = new Send_Budget();
                string html = budget.Html(argType);
                if (html == "") return;
                WriteLog("RunBudget: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailBudget(budget._subject, html, budget._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        private void CreateMailBudget(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\budget_ko.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'></body>", imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                WriteLog("CreateMailProduction: " + ex.ToString());
            }
        }

        #endregion Budget

        #region IE Relief

        private void cmd_IeRelief_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunIeRelief("Q");
        }

        private void RunIeRelief(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_IE_Relief iE_Relief = new Send_IE_Relief();
                string html = iE_Relief.Html(argType);
                if (html == "") return;
                WriteLog("RunIeRelief: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailIeRelief(iE_Relief._subject, html, iE_Relief._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        private void CreateMailIeRelief(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ie_relief_kr.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ie_relief_vi.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo", imgInfo2 = "imgInfo2";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                //mailItem.HTMLBody = htmlBody;
                mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'><br><img src='cid:{1}'></body>", imgInfo, imgInfo2) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
                WriteLog("RunIeRelief: Send Ok -->" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            catch (Exception ex)
            {
                WriteLog("RunIeRelief: " + ex.Message);
            }
        }

        #endregion Budget

        #region Hourly Production Tracking

        private void cmd_EscanSituationTracking_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunEscanSituationTracking("Q");
        }

        private void RunEscanSituationTracking(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Escan_Situation_Tracking Hourly_Prod_Tracking = new Send_Escan_Situation_Tracking();
                string html = Hourly_Prod_Tracking.Html(argType);
                if (html == "") return;
                WriteLog("Escan Situation Tracking: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailEscanSituationTracking(Hourly_Prod_Tracking._subject, html, Hourly_Prod_Tracking._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        private void CreateMailEscanSituationTracking(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                //  Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ie_relief_kr.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                // Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ie_relief_vi.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                //string imgInfo = "imgInfo", imgInfo2 = "imgInfo2";
                // oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                // oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                mailItem.HTMLBody = htmlBody;
                // mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'><br><img src='cid:{1}'></body>", imgInfo, imgInfo2) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
                WriteLog("Escan Situation Tracking: Send Ok -->" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            catch (Exception ex)
            {
                WriteLog("Escan Situation Tracking: " + ex.Message);
            }
        }

        #endregion

        #region Bol RR

        private void cmd_BolRR_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunBolRr("Q");
        }

        private void RunBolRr(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Bol_Rr send_Bol_Rr = new Send_Bol_Rr();
                string html = send_Bol_Rr.Html(argType);
                if (html == "") return;
                WriteLog("Bol RR: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailRunBolRr(send_Bol_Rr._subject, html, send_Bol_Rr._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        private void CreateMailRunBolRr(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\RR_kr.JPG", Outlook.OlAttachmentType.olByValue, null, "tr");
                // Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ie_relief_vi.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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

                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
                {
                    mailItem.BCC = "nguyen.it@changshininc.com; dien.it@changshininc.com; ngoc.it@changshininc.com";
                }
                else
                {
                    mailItem.BCC = "ngoc.it@changshininc.com";
                }

                string imgInfo = "imgInfo";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                // oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                //mailItem.HTMLBody = htmlBody;
                mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'></body>", imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
                WriteLog("Bol RR: Send Ok -->" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            catch (Exception ex)
            {
                WriteLog("Bol RR: " + ex.Message);
            }
        }

        #endregion

        #region Quality

        private void RunQuality(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Quality quality = new Send_Quality();
                string html = quality.Html(argType);
                if (html == "") return;
                WriteLog("RunQuality: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailQuality(quality._subject, html, quality._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        private void RunQuality2(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Quality2 quality2 = new Send_Quality2();
                string html = quality2.Html(argType);
                if (html == "") return;
                WriteLog("RunQuality: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailQuality2(quality2._subject, html, quality2._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        private void RunQualityMonth(string argType)
        {
            try
            {
                if (_isRun2) return;
                _isRun2 = true;
                string _subject = "";
                DataSet dsData = SEL_REWORK_MONTHLY_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));

                if (dsData == null) return;

                DataTable dtData1 = dsData.Tables[0];
                DataTable dtData2 = dsData.Tables[1];
                DataTable dtData3 = dsData.Tables[2];
                DataTable dtData4 = dsData.Tables[3];
                DataTable dtExplained = dsData.Tables[4];
                DataTable dtEmail = dsData.Tables[5];
                BindingControls(dtData1, dtData2, dtData3, dtData4);
                _subject = dtExplained.Rows[0]["SUBJECT"].ToString();
                //string html = GetHtmlBodyRework();
                //if (html == "") return;
                //WriteLog("RunQualityMonth: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                //if (html.StartsWith("Error"))
                //{
                //    WriteLog(html);
                //    return;
                //}
                CreateMailReworkMonthly(_subject, string.Empty, dtEmail, dtExplained);
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




        private void BindingControls(DataTable dtData1, DataTable dtData2, DataTable dtData3, DataTable dtData4)
        {
            try
            {
                //Bar Chart Rework
                chartReworkPlant.DataSource = dtData1;
                chartReworkPlant.Series[0].ArgumentDataMember = "LINE_NM";
                chartReworkPlant.Series[0].ValueDataMembers.AddRange(new string[] { "REWORK_RATE" });

                for (int i = 0; i < dtData1.Rows.Count; i++)
                {
                    chartReworkPlant.Series[0].Points[i].Color = Color.FromName(dtData1.Rows[i]["BCOLOR_RE"].ToString());
                }
                CaptureControl(pnchartReworkPlant, "chartReworkMonthly");
                //Pie Chart Reason
                chartReworkReason.DataSource = dtData2;
                chartReworkReason.Series[0].ArgumentDataMember = "REWORK_NAME_EN";
                chartReworkReason.Series[0].ValueScaleType = ScaleType.Numerical;
                chartReworkReason.Series[0].ValueDataMembers.AddRange(new string[] { "REWORK_QTY" });
                CaptureControl(pnChartReworkReason, "chartReworkReasonMonthly");

                //BC Grade Chart
                chartBCGrade.DataSource = dtData3;
                chartBCGrade.Series[0].ArgumentDataMember = "LINE_NM";
                chartBCGrade.Series[0].ValueScaleType = ScaleType.Numerical;
                chartBCGrade.Series[0].ValueDataMembers.AddRange(new string[] { "BC_RATE" });
                for (int i = 0; i < dtData3.Rows.Count; i++)
                {
                    chartBCGrade.Series[0].Points[i].Color = Color.FromName(dtData3.Rows[i]["BCOLOR_BC"].ToString());
                }


                CaptureControl(pnchartBCGrade, "chartBCGrade");

                //Binding data Grid
                BindingGrid(dtData4);
                CaptureControl(grdRework, "gridRework");

            }
            catch (Exception ex)
            {
                WriteLog("BindingControls: " + ex.Message);

            }
        }
        private void BindingGrid(DataTable dt)
        {
            try
            {

                while (gvwView.Columns.Count > 0)
                {
                    gvwView.Columns.RemoveAt(0);
                }
                DevExpress.XtraGrid.Columns.GridColumn gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
                gridColumn1.AppearanceHeader.Options.UseTextOptions = true;
                gridColumn1.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gridColumn1.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                gridColumn1.Caption = "Division";
                gridColumn1.FieldName = "ITEMS";
                gridColumn1.Name = "gridColumn1";
                gridColumn1.Visible = true;
                gridColumn1.VisibleIndex = 0;
                gridColumn1.Width = 100;
                gvwView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] { gridColumn1 });
                int vIdx = 1;
                DataView dv = new DataView(dt);
                DataTable dtHead = dv.ToTable(true, "CAL_DATE");
                foreach (DataRow dr in dtHead.Rows)
                {
                    DevExpress.XtraGrid.Columns.GridColumn gColumns = new DevExpress.XtraGrid.Columns.GridColumn();
                    gColumns.AppearanceHeader.Options.UseTextOptions = true;
                    gColumns.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gColumns.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gColumns.Caption = dr["CAL_DATE"].ToString();
                    gColumns.FieldName = dr["CAL_DATE"].ToString();
                    gColumns.Name = dr["CAL_DATE"].ToString();
                    gColumns.Visible = true;
                    gColumns.VisibleIndex = vIdx;
                    gColumns.Width = 50;
                    vIdx++;
                    gvwView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] { gColumns });

                }
                DataTable dtSource = Pivot(dt, dt.Columns["CAL_DATE"], dt.Columns["QTY"]);
                grdRework.DataSource = dtSource;

                for (int i = 0; i < gvwView.Columns.Count; i++)
                {
                    if (i > 0)
                    {
                        gvwView.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                        gvwView.Columns[i].AppearanceCell.TextOptions.VAlignment = VertAlignment.Center;
                        //gvwView.Columns[i].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
                        //gvwView.Columns[i].DisplayFormat.FormatString = "#,#.#";
                    }
                }
                gvwView.Columns["TOTAL"].Width = 70;
            }
            catch (Exception ex)
            {

                WriteLog("BindingGrid: " + ex.Message);
            }
        }
        #endregion

        #region Bottom Defective Rate
        private void RunBotDef(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Bottom_Defective bottom_Defective = new Send_Bottom_Defective();
                string html = bottom_Defective.Html(argType);
                if (html == "") return;
                WriteLog("RunQuality: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailBotDef(bottom_Defective._subject, html, bottom_Defective._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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
        #endregion

        #region Assembly Inline Inventory
        private void RunAssInLine(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_FGA_INV_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return;
                WriteLog($"RunFGAInlineInv({argType}): BEGIN ");
                DataTable dtDate = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog("  " + dtDate.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                CreateMailFGAInlineInv(dtDate, dtData, dtEmail);
                WriteLog($"RunFGAInlineInv({argType}): END ");
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

        private void RunAssInLine_v2(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                DataSet dsData = SEL_FGA_INV_DATA_v2("Q1", DateTime.Now.ToString("yyyyMMdd")); //Get Data for HTML Table


                if (dsData == null) return;
                WriteLog($"RunFGAInlineInv({argType}): BEGIN ");
                DataTable dtData = dsData.Tables[0];
                DataTable dtChart = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                BindingFGA_INVChart(dtChart);

                CaptureControl(chartFGA_INV, "CHART_FGA_INV");

                WriteLog("  " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                CreateMailFGAInlineInv_v2(dtData, dtEmail);
                WriteLog($"RunFGAInlineInv({argType}): END ");
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

        private void CreateMailFGAInlineInv(DataTable dtDate, DataTable dtData, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\FGA_INV_KOR.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\FGA_INV_VIE.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = "Inline inventory by each assembly line";

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                mailItem.Body = "This is the message.";
                string imgInfo = "imgInfo";
                string imgInfo1 = "imgInfo1";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo1);
                string rowValue = "";

                string strRowSpan = "";

                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    strRowSpan = dtData.Rows[iRow]["CNT"].ToString();
                    if (iRow == 0)
                    {
                        rowValue += "<tr>" +
                                       "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + " </td>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D6_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D6_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D6"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D5_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D5_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D5"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D4_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D4_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D4"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D3_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D3_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D3"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D2_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D2_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D2"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D1_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D1_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D1"].ToString() + "</td>" +
                                       // "<td align='right' >" + dtData.Rows[iRow]["PLAN"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["PLAN_2H"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["TODAY_BG_COLOR"].ToString() + "'  style='color:" + dtData.Rows[iRow]["TODAY_FORE_COLOR"].ToString() + "'  align='right'>" + dtData.Rows[iRow]["INV"].ToString() + "</td>" +
                                  "</tr>";
                    }
                    else
                    {
                        if (dtData.Rows[iRow]["PLANT"].ToString() == dtData.Rows[iRow - 1]["PLANT"].ToString())
                        {
                            rowValue += "<tr>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D6_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D6_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D6"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D5_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D5_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D5"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D4_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D4_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D4"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D3_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D3_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D3"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D2_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D2_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D2"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D1_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D1_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D1"].ToString() + "</td>" +
                                       //  "<td align='right' >" + dtData.Rows[iRow]["PLAN"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["PLAN_2H"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["TODAY_BG_COLOR"].ToString() + "'  style='color:" + dtData.Rows[iRow]["TODAY_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["INV"].ToString() + "</td>" +
                                  "</tr>";
                        }
                        else
                        {
                            rowValue += "<tr>" +
                                         "<td rowspan='" + strRowSpan + "' align ='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + " </td>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["MLINE_CD"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D6_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D6_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D6"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D5_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D5_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D5"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D4_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D4_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D4"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D3_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D3_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D3"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D2_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D2_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D2"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["D1_BG_COLOR"].ToString() + "' style='color:" + dtData.Rows[iRow]["D1_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["D1"].ToString() + "</td>" +
                                       // "<td align='right' >" + dtData.Rows[iRow]["PLAN"].ToString() + "</td>" +
                                       "<td align='right'>" + dtData.Rows[iRow]["PLAN_2H"].ToString() + "</td>" +
                                       "<td bgcolor='" + dtData.Rows[iRow]["TODAY_BG_COLOR"].ToString() + "'  style='color:" + dtData.Rows[iRow]["TODAY_FORE_COLOR"].ToString() + "' align='right'>" + dtData.Rows[iRow]["INV"].ToString() + "</td>" +
                                  "</tr>";
                        }
                    }
                }

                string strDate = "";
                foreach (DataRow row in dtDate.Rows)
                {
                    strDate += "<th bgcolor = '#ff9900' style = 'color:#ffffff' align = 'center' width = '70' >" + row["YMD"].ToString() + " </th >";
                }

                string html = "<img src='cid:" + imgInfo + "'><br><img src='cid:" + imgInfo1 + "'><br><table style='font-family:Times New Roman; font-size:20px; font-style: italic;' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0'>" +
                    "<tr ><td colspan='2' align='center'><strong>Assembly Inline Inventory Target</strong></td></tr>" +
                    "<tr><td align='left'>&le; 1 Hour</td><td align='center' bgcolor = 'green' style = 'color:#ffffff'>Green</td></tr>" +
                      "<tr><td align='left'>&gt; 1 Hour</td><td align='center' bgcolor = 'red' style = 'color:#ffffff'>Red</td></tr>" +
                    "</table> <h3><strong>UNIT: PAIRS</strong></h3>" +
                    "<p></p>" +
                    "          <table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1400'>" +
                               "<tr bgcolor='#ffe5cc'>" +
                                  " <th rowspan = '2' align='center' width='70'>Plant</th>" +
                                  " <th rowspan = '2' align='center' width='70'>Mini Line</th>" +
                                  " <th bgcolor = '#ff9900' style = 'color:#ffffff' colspan = '6' align='center'>Full time on previous day inventory</th>" +
                                  " <th bgcolor = '#366cc9' style = 'color:#ffffff' colspan = '2' align='center'>Yesterday inventory</th>" +
                               "</tr>" +
                               "<tr>" +
                                  strDate +
                                  //"<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='100'>Daily Plan</th>" +
                                  "<th bgcolor='#366cc9' style='color:#ffffff' align='center' width='100'>1 Hour Plan</th>" +
                                  "<th bgcolor='#f7d231' style='color:#000000' align='center' width='100'>Inline Inventory</th>" +
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

        private void CreateMailFGAInlineInv_v2(DataTable dtData, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                // Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\FGA_INV_KOR.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                //  Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\FGA_INV_VIE.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic3 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\CHART_FGA_INV.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = "Set balance situation in front of FGA";

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                mailItem.Body = "This is the message.";
                //   string imgInfo = "imgInfo";
                //   string imgInfo1 = "imgInfo1";
                string imgInfo2 = "imgInfo2";
                //  oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                //  oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo1);
                oAttachPic3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                string rowValue = "";

                string strRowFACSpan = "", strRowPLANTSpan = "";

                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    strRowFACSpan = dtData.Rows[iRow]["CNT_FAC"].ToString();
                    strRowPLANTSpan = dtData.Rows[iRow]["CNT_PLANT"].ToString();
                    string bg_tot_color = string.Empty;
                    if (iRow == 0)
                    {

                        if (dtData.Rows[iRow]["FACTORY"].ToString().ToUpper().Equals("GRAND TOTAL"))
                        {
                            bg_tot_color = "#34c916";
                        }

                        rowValue += "<tr bgcolor='" + bg_tot_color + "'>" +
                                       "<td rowspan='" + strRowFACSpan + "' align ='center'>" + dtData.Rows[iRow]["FACTORY"].ToString() + " </td>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + "</td>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["LINE"].ToString() + "</td>" +
                                       "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["UP_QTY"]) + "</td>" +
                                       "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["FS_QTY"]) + "</td>" +
                                       "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["SET_QTY"]) + "</td>" +
                                       "<td align='right'>" + string.Format("{0:n1}", dtData.Rows[iRow]["SET_RATIO"]) + "</td>" +
                                  "</tr>";

                    }
                    else
                    {
                        if (dtData.Rows[iRow]["FACTORY"].ToString() == dtData.Rows[iRow - 1]["FACTORY"].ToString())
                        {
                            if (dtData.Rows[iRow]["PLANT"].ToString() == dtData.Rows[iRow - 1]["PLANT"].ToString())
                            {
                                rowValue += "<tr>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["LINE"].ToString() + "</td>" +
                                       "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["UP_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["FS_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["SET_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n1}", dtData.Rows[iRow]["SET_RATIO"]) + "</td>" +
                                  "</tr>";
                            }
                            else
                            {
                                rowValue += "<tr>" +
                                       "<td rowspan='" + strRowPLANTSpan + "' align='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + "</td>" +
                                       "<td align='center'>" + dtData.Rows[iRow]["LINE"].ToString() + "</td>" +
                                        "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["UP_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["FS_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["SET_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n1}", dtData.Rows[iRow]["SET_RATIO"]) + "</td>" +
                                  "</tr>";
                            }
                        }
                        else
                        {
                            if (dtData.Rows[iRow]["FACTORY"].ToString().ToUpper().Equals("GRAND TOTAL"))
                            {
                                bg_tot_color = "#34c916";
                            }
                            rowValue += "<tr bgcolor='" + bg_tot_color + "'>" +
                                     "<td rowspan='" + strRowFACSpan + "' align ='center'>" + dtData.Rows[iRow]["FACTORY"].ToString() + " </td>" +
                                     "<td rowspan='" + strRowPLANTSpan + "' align='center'>" + dtData.Rows[iRow]["PLANT"].ToString() + "</td>" +
                                     "<td align='center'>" + dtData.Rows[iRow]["LINE"].ToString() + "</td>" +
                                       "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["UP_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["FS_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n0}", dtData.Rows[iRow]["SET_QTY"]) + "</td>" +
                                           "<td align='right'>" + string.Format("{0:n1}", dtData.Rows[iRow]["SET_RATIO"]) + "</td>" +
                                "</tr>";
                        }
                    }
                }

                string html = "<img src='cid:" + imgInfo2 + "'><br>" +
                    //"<table style='font-family:Times New Roman; font-size:20px; font-style: italic;' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0'>" +
                    //"<tr ><td colspan='2' align='center'><strong>Assembly Inventory Set Balance Target</strong></td></tr>" +
                    //"<tr><td align='left'>Under 2 Hours</td><td align='center' bgcolor = 'green' style = 'color:#ffffff'>Green</td></tr>" +
                    // "<tr><td align='left'>2~3 Hours</td><td align='center' bgcolor = 'yellow' style = 'color:black'>Yellow</td></tr>" +
                    //  "<tr><td align='left'>Over 3 Hours</td><td align='center' bgcolor = 'red' style = 'color:#ffffff'>Red</td></tr>" +
                    //"</table>" +
                    " <h3><strong>Unit: pairs</strong></h3>" +
                    "<p></p>" +
                    "          <table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1000'>" +
                               "<tr bgcolor='#ffe5cc'>" +
                                  " <th bgcolor = '#0760f0' style = 'color:#ffffff' align='center' width='200'>Factory</th>" +
                                  " <th bgcolor = '#0760f0' style = 'color:#ffffff' align='center' width='200'>Plant</th>" +
                                  " <th bgcolor = '#0760f0' style = 'color:#ffffff' align='center' width='200' >Line</th>" +
                                  " <th bgcolor = '#f0ba07' align='center' width='400'>Upper Inventory</th>" +
                                  " <th bgcolor = '#f0ba07' align='center' width='400'>Finish Sole Inventory</th>" +
                                  " <th bgcolor = '#f0ba07' align='center' width='400'>Set Balance</th>" +
                                  " <th bgcolor = '#f0ba07' align='center' width='400'>Set Ratio (%)</th>" +
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
        private void BindingFGA_INVChart(DataTable dt)
        {
            try
            {
                chartFGA_INV.DataSource = dt;
                chartFGA_INV.Series[0].ArgumentDataMember = "PLANT";
                chartFGA_INV.Series[0].ValueDataMembers.AddRange(new string[] { "SET_RATIO" });
            }
            catch
            {

            }
        }
        public DataSet SEL_FGA_INV_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_FGA_INV";
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
        public DataSet SEL_FGA_INV_DATA_v2(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_INV_SET";
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

        #region Sum DaaS

        private void CreateMailSumOpenDaaS(DataTable dtData, string Subject, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\OPEN_DAAS.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\Open_Daas_Explain.jpg", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                string imgInfo = "imgInfo", imgInfo2 = "imgInfo2";

                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                string rowValueVSM = "", rowValueBot = "";

                string strRowFACSpan = "", strRowPLANTSpan = "";

                for (int iRow = 0; iRow < dtData.Rows.Count; iRow++)
                {
                    strRowFACSpan = dtData.Rows[iRow]["CNT_FAC"].ToString();
                    strRowPLANTSpan = dtData.Rows[iRow]["CNT_PLANT"].ToString();
                    string bg_tot_color = string.Empty;
                    string strFactory = dtData.Rows[iRow]["PLANT_NAME"].ToString();
                    string strPlant = dtData.Rows[iRow]["LINE_NAME"].ToString();
                    string strLine = dtData.Rows[iRow]["MLINE_CD"].ToString();

                    string strAndonDt = dtData.Rows[iRow]["ANDON_DT"].ToString();
                    string strAndonDtBc = dtData.Rows[iRow]["ANDON_DT_BC"].ToString();
                    string strAndonDtFc = dtData.Rows[iRow]["ANDON_DT_FC"].ToString();

                    string strNpi = dtData.Rows[iRow]["NPI"].ToString();
                    string strNpiBc = dtData.Rows[iRow]["NPI_BC"].ToString();
                    string strNpiFc = dtData.Rows[iRow]["NPI_FC"].ToString();

                    string strTimeContraint = dtData.Rows[iRow]["TIME_CONTRAINT"].ToString();
                    string strTimeContraintBc = dtData.Rows[iRow]["TIME_CONTRAINT_BC"].ToString();
                    string strTimeContraintFc = dtData.Rows[iRow]["TIME_CONTRAINT_FC"].ToString();

                    string strRework = dtData.Rows[iRow]["REWORK"].ToString();
                    string strReworkBc = dtData.Rows[iRow]["REWORK_BC"].ToString();
                    string strReworkFc = dtData.Rows[iRow]["REWORK_FC"].ToString();

                    string strMoldRepair = dtData.Rows[iRow]["MOLD_RP"].ToString();
                    string strMoldRepairBc = dtData.Rows[iRow]["MOLD_RP_BC"].ToString();
                    string strMoldRepairFc = dtData.Rows[iRow]["MOLD_RP_FC"].ToString();

                    string strBottomDefective = dtData.Rows[iRow]["BOTTOM_DEF_RP"].ToString();
                    string strBottomDefectiveBc = dtData.Rows[iRow]["BOTTOM_DEF_RP_BC"].ToString();
                    string strBottomDefectiveFc = dtData.Rows[iRow]["BOTTOM_DEF_RP_FC"].ToString();

                    string strIsb = dtData.Rows[iRow]["FGA_INV_SET"].ToString();
                    string strIsbBc = dtData.Rows[iRow]["FGA_INV_SET_BC"].ToString();
                    string strIsbFc = dtData.Rows[iRow]["FGA_INV_SET_FC"].ToString();

                    string strProductionTarget = dtData.Rows[iRow]["MEET_TARGET"].ToString();
                    string strProductionTargetBc = dtData.Rows[iRow]["MEET_TARGET_BC"].ToString();
                    string strProductionTargetFc = dtData.Rows[iRow]["MEET_TARGET_FC"].ToString();

                    string strTms = dtData.Rows[iRow]["TMS"].ToString();
                    string strTmsBc = dtData.Rows[iRow]["TMS_BC"].ToString();
                    string strTmsFc = dtData.Rows[iRow]["TMS_FC"].ToString();

                    string strPod = dtData.Rows[iRow]["POD"].ToString();
                    string strPodBc = dtData.Rows[iRow]["POD_BC"].ToString();
                    string strPodFc = dtData.Rows[iRow]["POD_FC"].ToString();

                    string strTopo = dtData.Rows[iRow]["TO_PO"].ToString();
                    string strTopoBc = dtData.Rows[iRow]["TO_PO_BC"].ToString();
                    string strTopoFc = dtData.Rows[iRow]["TO_PO_FC"].ToString();

                    string strAbsent = dtData.Rows[iRow]["ABS_RATE"].ToString();
                    string strAbsentBc = dtData.Rows[iRow]["ABS_RATE_BC"].ToString();
                    string strAbsentFc = dtData.Rows[iRow]["ABS_RATE_FC"].ToString();


                    if (iRow == 0)
                    {
                        rowValueVSM += "<tr>" +
                                       "<td rowspan='" + strRowFACSpan + "' align ='center'>" + strFactory + " </td>" +
                                       "<td align='center'>" + strPlant + "</td>" +
                                       "<td align='center'>" + strLine + "</td>" +
                                       "<td bgcolor='" + strAndonDtBc + "' style = 'color:" + strAndonDtFc + "' align='center' >" + strAndonDt + "</td>" +
                                       "<td bgcolor='" + strNpiBc + "' style = 'color:" + strNpiFc + "' align='right' >" + strNpi + "</td>" +
                                       "<td bgcolor='" + strTimeContraintBc + "' style = 'color:" + strTimeContraintFc + "' align='right' >" + strTimeContraint + "</td>" +
                                       "<td bgcolor='" + strReworkBc + "' style = 'color:" + strReworkFc + "' align='right' >" + strRework + "</td>" +
                                       "<td bgcolor='" + strIsbBc + "' style = 'color:" + strIsbFc + "' align='right' >" + strIsb + "</td>" +
                                       "<td bgcolor='" + strProductionTargetBc + "' style = 'color:" + strProductionTargetFc + "' align='right' >" + strProductionTarget + "</td>" +
                                       "<td bgcolor='" + strTmsBc + "' style = 'color:" + strTmsFc + "' align='right' >" + strTms + "</td>" +
                                       "<td bgcolor='" + strPodBc + "' style = 'color:" + strPodFc + "' align='right' >" + strPod + "</td>" +
                                       "<td bgcolor='" + strTopoBc + "' style = 'color:" + strTopoFc + "' align='right' >" + strTopo + "</td>" +
                                       "<td bgcolor='" + strAbsentBc + "' style = 'color:" + strAbsentFc + "' align='right' >" + strAbsent + "</td>" +
                                  "</tr>";


                    }
                    else
                    {
                        if (strFactory == dtData.Rows[iRow - 1]["PLANT_NAME"].ToString())
                        {
                            if (strPlant == dtData.Rows[iRow - 1]["LINE_NAME"].ToString())
                            {

                                if (strFactory.Contains("Bottom"))
                                {
                                    rowValueBot += "<tr>" +

                                                        "<td bgcolor='" + strTimeContraintBc + "' style = 'color:" + strTimeContraintFc + "' align='right' >" + strTimeContraint + "</td>" +
                                                        "<td bgcolor='" + strMoldRepairBc + "' style = 'color:" + strMoldRepairFc + "' align='right' >" + strMoldRepair + "</td>" +
                                                        "<td bgcolor='" + strBottomDefectiveBc + "' style = 'color:" + strBottomDefectiveFc + "' align='right' >" + strBottomDefective + "</td>" +

                                                        "<td bgcolor='" + strTmsBc + "' style = 'color:" + strTmsFc + "' align='right' >" + strTms + "</td>" +
                                                        "<td bgcolor='" + strPodBc + "' style = 'color:" + strPodFc + "' align='right' >" + strPod + "</td>" +
                                                        "<td bgcolor='" + strTopoBc + "' style = 'color:" + strTopoFc + "' align='right' >" + strTopo + "</td>" +
                                                        "<td bgcolor='" + strAbsentBc + "' style = 'color:" + strAbsentFc + "' align='right' >" + strAbsent + "</td>" +
                                                    "</tr>";
                                }
                                else
                                {
                                    rowValueVSM += "<tr>" +
                                                    "<td align='center'>" + strLine + "</td>" +
                                                    "<td bgcolor='" + strAndonDtBc + "' style = 'color:" + strAndonDtFc + "' align='center' >" + strAndonDt + "</td>" +
                                                    "<td bgcolor='" + strNpiBc + "' style = 'color:" + strNpiFc + "' align='left' >" + strNpi + "</td>" +
                                                    "<td bgcolor='" + strTimeContraintBc + "' style = 'color:" + strTimeContraintFc + "' align='right' >" + strTimeContraint + "</td>" +
                                                    "<td bgcolor='" + strReworkBc + "' style = 'color:" + strReworkFc + "' align='right' >" + strRework + "</td>" +
                                                    "<td bgcolor='" + strIsbBc + "' style = 'color:" + strIsbFc + "' align='right' >" + strIsb + "</td>" +
                                                    "<td bgcolor='" + strProductionTargetBc + "' style = 'color:" + strProductionTargetFc + "' align='right' >" + strProductionTarget + "</td>" +
                                                    "<td bgcolor='" + strTmsBc + "' style = 'color:" + strTmsFc + "' align='right' >" + strTms + "</td>" +
                                                    "<td bgcolor='" + strPodBc + "' style = 'color:" + strPodFc + "' align='right' >" + strPod + "</td>" +
                                                    "<td bgcolor='" + strTopoBc + "' style = 'color:" + strTopoFc + "' align='right' >" + strTopo + "</td>" +
                                                    "<td bgcolor='" + strAbsentBc + "' style = 'color:" + strAbsentFc + "' align='right' >" + strAbsent + "</td>" +
                                                "</tr>";
                                }

                            }
                            else
                            {

                                if (strFactory.Contains("Bottom"))
                                {
                                    rowValueBot += "<tr>" +
                                                    "<td rowspan='" + strRowPLANTSpan + "' align='center'>" + strPlant + "</td>" +

                                                    "<td bgcolor='" + strTimeContraintBc + "' style = 'color:" + strTimeContraintFc + "' align='right' >" + strTimeContraint + "</td>" +
                                                    "<td bgcolor='" + strMoldRepairBc + "' style = 'color:" + strMoldRepairFc + "' align='right' >" + strMoldRepair + "</td>" +
                                                    "<td bgcolor='" + strBottomDefectiveBc + "' style = 'color:" + strBottomDefectiveFc + "' align='right' >" + strBottomDefective + "</td>" +

                                                    "<td bgcolor='" + strTmsBc + "' style = 'color:" + strTmsFc + "' align='right' >" + strTms + "</td>" +
                                                    "<td bgcolor='" + strPodBc + "' style = 'color:" + strPodFc + "' align='right' >" + strPod + "</td>" +
                                                    "<td bgcolor='" + strTopoBc + "' style = 'color:" + strTopoFc + "' align='right' >" + strTopo + "</td>" +
                                                    "<td bgcolor='" + strAbsentBc + "' style = 'color:" + strAbsentFc + "' align='right' >" + strAbsent + "</td>" +
                                                "</tr>";
                                }
                                else
                                {
                                    rowValueVSM += "<tr>" +
                                                    "<td  rowspan='" + strRowPLANTSpan + "' align='center'>" + strPlant + "</td>" +
                                                    "<td align='center'>" + strLine + "</td>" +
                                                    "<td bgcolor='" + strAndonDtBc + "' style = 'color:" + strAndonDtFc + "' align='center' >" + strAndonDt + "</td>" +
                                                    "<td bgcolor='" + strNpiBc + "' style = 'color:" + strNpiFc + "' align='left' >" + strNpi + "</td>" +
                                                    "<td bgcolor='" + strTimeContraintBc + "' style = 'color:" + strTimeContraintFc + "' align='right' >" + strTimeContraint + "</td>" +
                                                    "<td bgcolor='" + strReworkBc + "' style = 'color:" + strReworkFc + "' align='right' >" + strRework + "</td>" +
                                                    "<td bgcolor='" + strIsbBc + "' style = 'color:" + strIsbFc + "' align='right' >" + strIsb + "</td>" +
                                                    "<td bgcolor='" + strProductionTargetBc + "' style = 'color:" + strProductionTargetFc + "' align='right' >" + strProductionTarget + "</td>" +
                                                    "<td bgcolor='" + strTmsBc + "' style = 'color:" + strTmsFc + "' align='right' >" + strTms + "</td>" +
                                                    "<td bgcolor='" + strPodBc + "' style = 'color:" + strPodFc + "' align='right' >" + strPod + "</td>" +
                                                    "<td bgcolor='" + strTopoBc + "' style = 'color:" + strTopoFc + "' align='right' >" + strTopo + "</td>" +
                                                    "<td bgcolor='" + strAbsentBc + "' style = 'color:" + strAbsentFc + "' align='right' >" + strAbsent + "</td>" +
                                                "</tr>";
                                }

                            }
                        }
                        else
                        {


                            if (strFactory.Contains("Bottom"))
                            {
                                rowValueBot += "<tr>" +
                                              "<td rowspan='" + strRowFACSpan + "' align ='center'>" + strFactory + " </td>" +
                                              "<td rowspan='" + strRowPLANTSpan + "' align='center'>" + strPlant + "</td>" +

                                              "<td bgcolor='" + strTimeContraintBc + "' style = 'color:" + strTimeContraintFc + "' align='right' >" + strTimeContraint + "</td>" +
                                              "<td bgcolor='" + strMoldRepairBc + "' style = 'color:" + strMoldRepairFc + "' align='right' >" + strMoldRepair + "</td>" +
                                              "<td bgcolor='" + strBottomDefectiveBc + "' style = 'color:" + strBottomDefectiveFc + "' align='right' >" + strBottomDefective + "</td>" +

                                              "<td bgcolor='" + strTmsBc + "' style = 'color:" + strTmsFc + "' align='right' >" + strTms + "</td>" +
                                              "<td bgcolor='" + strPodBc + "' style = 'color:" + strPodFc + "' align='right' >" + strPod + "</td>" +
                                              "<td bgcolor='" + strTopoBc + "' style = 'color:" + strTopoFc + "' align='right' >" + strTopo + "</td>" +
                                              "<td bgcolor='" + strAbsentBc + "' style = 'color:" + strAbsentFc + "' align='right' >" + strAbsent + "</td>" +
                                            "</tr>";
                            }
                            else
                            {
                                rowValueVSM += "<tr>" +
                                              "<td rowspan='" + strRowFACSpan + "' align ='center'>" + strFactory + " </td>" +
                                              "<td rowspan='" + strRowPLANTSpan + "' align='center'>" + strPlant + "</td>" +
                                              "<td align='center'>" + strLine + "</td>" +
                                              "<td bgcolor='" + strAndonDtBc + "' style = 'color:" + strAndonDtFc + "' align='center' >" + strAndonDt + "</td>" +
                                              "<td bgcolor='" + strNpiBc + "' style = 'color:" + strNpiFc + "' align='left' >" + strNpi + "</td>" +
                                              "<td bgcolor='" + strTimeContraintBc + "' style = 'color:" + strTimeContraintFc + "' align='right' >" + strTimeContraint + "</td>" +
                                              "<td bgcolor='" + strReworkBc + "' style = 'color:" + strReworkFc + "' align='right' >" + strRework + "</td>" +
                                              "<td bgcolor='" + strIsbBc + "' style = 'color:" + strIsbFc + "' align='right' >" + strIsb + "</td>" +
                                              "<td bgcolor='" + strProductionTargetBc + "' style = 'color:" + strProductionTargetFc + "' align='right' >" + strProductionTarget + "</td>" +
                                              "<td bgcolor='" + strTmsBc + "' style = 'color:" + strTmsFc + "' align='right' >" + strTms + "</td>" +
                                              "<td bgcolor='" + strPodBc + "' style = 'color:" + strPodFc + "' align='right' >" + strPod + "</td>" +
                                              "<td bgcolor='" + strTopoBc + "' style = 'color:" + strTopoFc + "' align='right' >" + strTopo + "</td>" +
                                              "<td bgcolor='" + strAbsentBc + "' style = 'color:" + strAbsentFc + "' align='right' >" + strAbsent + "</td>" +
                                           "</tr>";
                            }
                        }
                    }
                }

                string style = "<style> " +
                              "   .tblBoder { " +
                              "             font-family: 'Times New Roman', Times, serif; " +
                              "             font-style: italic; " +
                              "   } " +
                              "   .tblBoder td, .tblBoder th { " +
                              "         border: 0px; " +
                              "         padding: 3px 2px; " +
                              "         white-space: nowrap; " +
                              "         border: 1px solid #c0c0c0; " +
                              "   } " +
                              "   .tblBoder tbody td { " +
                              "             font-size: 20px; " +
                              "         } " +
                              ".title {" +
                              "         font-family: 'Times New Roman';" +
                              "         font-style: italic;" +
                              "         font-size: 30px;" +
                              "        }" +
                              "   .tblBoder thead { " +
                              "         background: #26A1B2; " +
                              "         font-style: italic; " +
                              "         border-bottom: 0px solid #444444; " +
                              "   } " +
                              "   .tblBoder thead th { " +
                              "             font-size: 19px; " +
                              "             font-weight: bold; " +
                              "         color: #F0F0F0; " +
                              "     background: #26A1B2; " +
                              "     text-align: center; " +
                              "         } " +
                              "   .green{ " +
                              "         background: green; " +
                              "         color: white; " +
                              "         } " +
                              "   .yellow{ " +
                              "         background: yellow; " +
                              "         color: black; " +
                              "         } " +
                              "   .red{ " +
                              "         background: red; " +
                              "         color: white; " +
                              "         } " +
                              "   .orange{ " +
                              "         background: orange; " +
                              "         color: white; " +
                              "         } " +
                              "   .gray{ " +
                              "         background: silver; " +
                              "         color: black; " +
                              "         } " +
                              "</style> ";
                string text = "<img src='cid:" + imgInfo + "'><br>" +
                              "<img src='cid:" + imgInfo2 + "'>";


                string html = "<head>" + style + "</head>" +
                               "<body>" + text + "<br>" +
                               "<br>" +
                               "<b class = 'title'>VSM</b>" +
                        "           <table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1920'>" +
                                   "<tr bgcolor='#ffe5cc'>" +
                                      " <th rowspan= '2'  bgcolor = '#18213C' style = 'color:#ffffff' align='center' width='100'>Factory</th>" +
                                      " <th rowspan= '2'  bgcolor = '#18213C' style = 'color:#ffffff' align='center' width='100'>Plant</th>" +
                                      " <th rowspan= '2'  bgcolor = '#18213C' style = 'color:#ffffff' align='center' width='100'>Line</th>" +
                                      " <th colspan = '4' bgcolor = '#18213C' style = 'color:#ffffff' align='center' >Quality</th>" +
                                      " <th colspan = '4' bgcolor = '#18213C' style = 'color:#ffffff' align='center' >Production/Inventory/Logistics</th>" +
                                      " <th colspan = '2' bgcolor = '#18213C' style = 'color:#ffffff' align='center' >HR</th>" +
                                   "</tr>" +
                                   "<tr bgcolor='#ffe5cc'>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100'>Andon<br>D/T</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100'>NPI</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Time<br>Contraint By<br>Stockfit</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Rework</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >ISB</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Production<br> Target</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >TMS</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >POD</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >TO/PO&nbsp;</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Absenteeism</th>" +
                                   "</tr>" +
                                   "<tr bgcolor='#ffe49c'>" +
                                      " <th colspan = '3' bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100'>Unit</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100'>Time</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100'>Day</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >Prs</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                   "</tr>" +
                                   rowValueVSM +
                               "</table>" +
                           "<br>" +
                           "<b class = 'title'>Bottom</b>" +
                           "           <table style='font-family:Calibri; font-size:20px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' width='1500'>" +
                                   "<tr bgcolor='#ffe5cc'>" +
                                      " <th rowspan= '2'  bgcolor = '#18213C' style = 'color:#ffffff' align='center' width='100'>Factory</th>" +
                                      " <th rowspan= '2'  bgcolor = '#18213C' style = 'color:#ffffff' align='center' width='100'>Plant</th>" +

                                      " <th colspan = '3' bgcolor = '#18213C' style = 'color:#ffffff' align='center' ' >Quality</th>" +
                                      " <th colspan = '2' bgcolor = '#18213C' style = 'color:#ffffff' align='center' '>Production/Inventory/Logistics</th>" +
                                      " <th colspan = '2' bgcolor = '#18213C' style = 'color:#ffffff' align='center' '>HR</th>" +
                                   "</tr>" +
                                   "<tr bgcolor='#ffe5cc'>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Time<br>Contraint By<br>Bottom</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Mold Repair</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Bottom<br>Defective</th>" +


                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >TMS</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >POD</th>" +

                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >TO/PO&nbsp;</th>" +
                                      " <th bgcolor = '#CCCCCE' style = 'color:#000' align='center' width='100' >Absenteeism</th>" +
                                   "</tr>" +
                                   "<tr bgcolor='#ffe49c'>" +
                                      " <th colspan = '2' bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100'>Unit</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >Prs</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +


                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +

                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                      " <th bgcolor = '#ffe49c' style = 'color:#000' align='center' width='100' >%</th>" +
                                   "</tr>" +
                                   rowValueBot +
                               "</table>" +
                           "</body>";

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

        public DataSet SEL_SUM_OPEN_DAAS(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            DataSet ds_ret;
            //  MyOraDB.ShowErr = true;
            try
            {
                string process_name = "P_SEND_EMAIL_OPEN_DAAS";
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

        #region Email Canteen
        private void cmdCanteen_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunCanteen("Q");
        }

        private void RunCanteen(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                //if (dsData == null) return;
                Send_Canteen Canteen = new Send_Canteen();
                string html = Canteen.Html(argType);
                if (html == "") return;
                WriteLog("RunCateen: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                if (html.StartsWith("Error"))
                {
                    WriteLog(html);
                    return;
                }
                // WriteLog("RunMoldRepair: Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                CreateMailCanteen(Canteen._subject, html, Canteen._email);
                //  WriteLog("RunMoldRepair: End --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            catch (Exception ex)
            {
                WriteLog("RunCateen: " + ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void CreateMailCanteen(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                // Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ie_relief_kr.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                // Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ie_relief_vi.png", Outlook.OlAttachmentType.olByValue, null, "tr");
                mailItem.Subject = Subject;

                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;

                //Get List Send email
                if (app.Session.CurrentUser.AddressEntry.Address.ToUpper().Contains("IT.DAAS"))
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
                //string imgInfo = "imgInfo", imgInfo2 = "imgInfo2";
                // oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                // oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                mailItem.HTMLBody = htmlBody;
                //  mailItem.HTMLBody = String.Format(@"<body><img src='cid:{0}'><br><img src='cid:{1}'></body>", imgInfo, imgInfo2) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
                WriteLog("Mail Canteen: Send Ok -->" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            catch (Exception ex)
            {
                WriteLog("Mail Canteen: " + ex.Message);
            }
        }
        #endregion

        #region Upper Inventory
        private DataSet SEL_DATA_UPPER_INV(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_UPPER_INV";
                MyOraDB.ReDim_Parameter(6);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "CV_1";
                MyOraDB.Parameter_Name[3] = "CV_2";
                MyOraDB.Parameter_Name[4] = "CV_3";
                MyOraDB.Parameter_Name[5] = "CV_EMAIL";


                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                        //WriteLog("P_SEND_EMAIL_NPI: null");
                    }
                    return null;
                }
                return ds_ret;
            }
            catch (Exception ex)
            {
                // WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }

        private void RunUpperInv(string argType, string argDate)
        {
            if (argType == "Q" || argType == "Q1")
            {
                DataSet ds = SEL_DATA_UPPER_INV(argType, argDate);//UPPER INVENTORY

                if (ds == null)
                    return;
                WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunUpperInv({argType}): BEGIN");

                using (frmUpper_Inventory frmUpperInv = new frmUpper_Inventory())
                {
                    frmUpperInv._chkTest = chkTest.Checked;
                    frmUpperInv._subject = "Upper Inventory";
                    frmUpperInv._dt1 = ds.Tables[0];
                    frmUpperInv._dt2 = ds.Tables[1];
                    frmUpperInv._dt3 = ds.Tables[2];
                    frmUpperInv._dt4 = ds.Tables[3];
                    frmUpperInv.Show();
                    frmUpperInv.SendToBack();
                }

                WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunUpperInv({argType}): END");
            }
        }


        #endregion

        private string ColorNull(string argColor)
        {
            return argColor == "" ? "WHITE" : argColor;
        }


        public void WriteLog(string argText)
        {
            try
            {
                txtLog.Invoke((MethodInvoker)(() =>
                {
                    txtLog.Text += argText + "\r\n";
                    txtLog.SelectionStart = txtLog.TextLength;
                    txtLog.ScrollToCaret();
                    txtLog.Refresh();
                }));
            }
            catch (Exception ex)
            {
            }

        }

        private void btnRunTMS_Click(object sender, EventArgs e)
        {
            RunTMSDash("Q");
        }

        private void btnRunTMSV2_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunTMSDashv2("Q");
        }

        private void checkRunning()
        {
            foreach (Process p in Process.GetProcessesByName("Send_Email"))
            {
                p.CloseMainWindow();
            }
        }

        private void RunTMSDashv2(string arg_type)
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

                string html = getHTMLBodyHeaderTMSDashv3(dtHeader, dtData);
                DataSet dsHead = SEL_DATA_TMS_DAAS_CHART("QH", DateTime.Now.AddDays(-6).ToString("yyyyMMdd"), DateTime.Now.ToString("yyyyMMdd"), "ALL");
                DataTable dtDays = dsHead.Tables[0];
                DataTable dtDates = dsHead.Tables[1];
                string DateF, DateT;
                DateF = dtDates.Rows[0]["DATEF"].ToString();
                DateT = dtDates.Rows[0]["DATET"].ToString();
                DataSet ds = SEL_DATA_TMS_DAAS_CHART("Q", DateF, DateT, "ALL");
                DataTable dtSource = new DataTable();
                DataTable dtDataGrid = ds.Tables[0];
                BuildHeader(dtSource, dtDays);
                BindingDataSource(dtSource, dtDays, dtDataGrid);
                grdBase.DataSource = dtSource;
                Format_Grid();
                DataTable dtChart = ds.Tables[1];
                loadchart(dtChart);

                //CaptureControl(pnTMSDassChart, "TMSChart");
                //CaptureControl(pnTMSDassGrid, "TMSGrid");
                CreateMailwithImage(Emoji.ChartIncreasing + " TMS MONITORING DETAIL", html, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            finally
            {
                _isRun2 = false;
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

        private void RunTimeContraint(string DivTag, string arg_type)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_TIME_CONTRAINT_DATA(arg_type);
                if (dsData == null) return;

                WriteLog("RunTimeContraint: " + DivTag + " Run --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                DataTable dtHeader = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                string html = getHTMLBodyHeaderTimeContraint(arg_type, dtHeader, dtData);

                CreateMailTimeContraint(Emoji.ChartIncreasing + " Time Constraint By " + DivTag, html, dtEmail);
                WriteLog("RunTimeContraint: " + DivTag + " end --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
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

        private void RunScada(string arg_type)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_SCADA_DATA(arg_type);
                if (dsData == null) return;

                DataTable dtHeader = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());
                string html = getHTMLBodyHeaderScada(dtHeader, dtData);
                CreateMailwithImage(Emoji.ChartIncreasing + "Top 50 Over Temperature", html, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void RunTMS_Summary(string arg_type)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;
                DataSet dsData = SEL_TMS_SUMMARY_DATA(arg_type);
                if (dsData == null) return;

                DataTable dtHeader = dsData.Tables[0];
                DataTable dtData = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];

                WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());
                string html = getHTMLBodyHeaderTMSSummary(dtHeader, dtData);
                CreateMailwithImage(Emoji.ChartIncreasing + "Weekly TMS Performance Summary", html, dtEmail);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            finally
            {
                _isRun2 = false;
            }
        }

        private void Format_Grid()
        {
            double order_qty = 0;
            double actual_qty = 0;

            #region replace

            gvwBase.BeginUpdate();

            for (int i = 0; i <= gvwBase.RowCount - 1; i++)
            {
                if (gvwBase.GetRowCellValue(i, gvwBase.Columns[0]).ToString() == "OSP")
                {
                    gvwBase.SetRowCellValue(i, gvwBase.Columns[0], "OS");
                }

                if (gvwBase.GetRowCellValue(i, gvwBase.Columns[0]).ToString() == "PHP")
                {
                    gvwBase.SetRowCellValue(i, gvwBase.Columns[0], "PH");
                }
                if (gvwBase.GetRowCellValue(i, gvwBase.Columns[0]).ToString() == "PUP")
                {
                    gvwBase.SetRowCellValue(i, gvwBase.Columns[0], "PU");
                }
                if (gvwBase.GetRowCellValue(i, gvwBase.Columns[0]).ToString() == "IPP")
                {
                    gvwBase.SetRowCellValue(i, gvwBase.Columns[0], "IP");
                }
                if (gvwBase.GetRowCellValue(i, gvwBase.Columns[0]).ToString() == "DMP")
                {
                    gvwBase.SetRowCellValue(i, gvwBase.Columns[0], "DMP");
                }

                if (gvwBase.GetRowCellValue(i, gvwBase.Columns[0]).ToString() == "TOTAL")
                {
                    gvwBase.SetRowCellValue(i, gvwBase.Columns[0], "Bottom Performance");
                }
            }

            for (int i = 0; i < gvwBase.Columns.Count; i++)
            {
                gvwBase.Columns[i].AppearanceCell.Options.UseTextOptions = true;
                gvwBase.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                gvwBase.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.True;
                gvwBase.Columns[i].OptionsFilter.AllowFilter = false;
                gvwBase.Columns[i].OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.False;
                gvwBase.Columns[i].OptionsColumn.AllowEdit = false;

                gvwBase.ColumnPanelRowHeight = 25;
                gvwBase.RowHeight = 25;
                gvwBase.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                if (i < 1)
                {
                    gvwBase.Columns[i].Width = 150;
                    gvwBase.Columns[i].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gvwBase.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                }
                else
                {
                    gvwBase.Columns[i].Width = 60;
                    gvwBase.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                    //  gvwBase.Columns[i].DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;

                    //  gvwBase.Columns[j].DisplayFormat.FormatString = "#,###";
                }
            }

            for (int i = 0; i <= gvwBase.RowCount - 1; i++)
            {
                for (int j = 1; j <= gvwBase.Columns.Count - 1; j += 3)
                {
                    double.TryParse(gvwBase.GetRowCellDisplayText(i, gvwBase.Columns[j]).ToString(), out order_qty);
                    double.TryParse(gvwBase.GetRowCellDisplayText(i, gvwBase.Columns[j + 1]).ToString(), out actual_qty);

                    if (order_qty > 0 && actual_qty == 0)
                    {
                        gvwBase.SetRowCellValue(i, gvwBase.Columns[j + 1], "0");
                    }
                }
            }

            gvwBase.Appearance.Row.Font = new System.Drawing.Font("DotumChe", 10F, System.Drawing.FontStyle.Regular);
            //   gvwBase.BestFitColumns();

            // gvwBase.OptionsView.ColumnAutoWidth = false;
            gvwBase.EndUpdate();

            #endregion replace
        }

        private void BuildHeader(DataTable dtSource, DataTable dtDays)
        {
            GridBandEx band, bandChild;
            string col_name = "";
            try
            {
                // Reset band header.

                while (gvwBase.Bands.Count > 0)
                {
                    gvwBase.Bands.RemoveAt(0);
                }
                while (gvwBase.Columns.Count > 0)
                {
                    gvwBase.Columns.RemoveAt(0);
                }

                for (int i = 0; i < headNames.Length; i++)
                {
                    band = new GridBandEx();
                    band.AppearanceHeader.Font = new System.Drawing.Font("DotumChe", 9F);
                    if (headNames[i].Equals("COMP"))
                    {
                        band.Caption = "Process";
                    }
                    else
                    {
                        band.Caption = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(headNames[i].ToLower().Replace("_", " "));
                    }

                    //band.MinWidth = 80;//TextRenderer.MeasureText(band.Caption, band.AppearanceHeader.Font).Width;
                    band.AppearanceHeader.Options.UseTextOptions = true;
                    band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                    gvwBase.Bands.Add(band);

                    AddColumn(band, headNames[i]);
                    dtSource.Columns.Add(headNames[i], typeof(string));
                }
                for (int i = 0; i < dtDays.Rows.Count; i++)
                {
                    band = new GridBandEx();
                    band.AppearanceHeader.Font = new System.Drawing.Font("DotumChe", 9F);
                    band.Caption = dtDays.Rows[i]["DATE_STRING"].ToString();
                    //if (i >= 3)
                    //{
                    //    band.MinWidth = TextRenderer.MeasureText(band.Caption, band.AppearanceHeader.Font).Width + 20;
                    //}
                    //band.MinWidth = 20;
                    band.AppearanceHeader.Options.UseTextOptions = true;
                    band.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                    for (int j = 1; j <= 3; j++)
                    {
                        if (j == 1) col_name = "Order";
                        else if (j == 2) col_name = "Actual";
                        else if (j == 3) col_name = "Percent";

                        bandChild = new GridBandEx();
                        bandChild.AppearanceHeader.Font = new System.Drawing.Font("DotumChe", 9F);
                        bandChild.Caption = col_name;
                        //bandChild.MinWidth = 20;
                        bandChild.AppearanceHeader.Options.UseTextOptions = true;
                        bandChild.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                        // TextRenderer.MeasureText(bandChild.Caption, bandChild.AppearanceHeader.Font).Width + 20;

                        band.Children.Add(bandChild);

                        AddColumn(bandChild, dtDays.Rows[i]["THEDATE"].ToString() + "_" + col_name);
                        dtSource.Columns.Add(dtDays.Rows[i]["THEDATE"].ToString() + "_" + col_name, typeof(double));
                    }
                    gvwBase.Bands.Add(band);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("BuildHeader(): " + ex.Message);
            }
        }

        private void BindingDataSource(DataTable dtSource, DataTable dtDays, DataTable dtData)
        {
            double Order = 0;
            double Actual = 0;
            double Percent = 0;
            double total_row = 0;
            double total_tmp = 0;

            int row_opcd = 0;
            try
            {
                dtSource.Rows.Add("OSP");
                dtSource.Rows.Add("PHP");
                dtSource.Rows.Add("PUP");
                dtSource.Rows.Add("IPP");
                dtSource.Rows.Add("DMP");
                //   dtSource.Rows.Add("TOTAL");

                for (int col = 1; col <= dtSource.Columns.Count - 1; col++)
                {
                    for (int i = 0; i <= dtData.Rows.Count - 1; i++)
                    {
                        double.TryParse(dtData.Rows[i]["Order"].ToString(), out Order);
                        double.TryParse(dtData.Rows[i]["Actual"].ToString(), out Actual);
                        double.TryParse(dtData.Rows[i]["Percent"].ToString(), out Percent);

                        //--tim dong opcd--
                        for (int row = 0; row <= 4; row++)
                        {
                            if (dtData.Rows[i]["OP_CD"].ToString() == dtSource.Rows[row][0].ToString())
                            {
                                row_opcd = row;
                                break;
                            }
                        }

                        if (dtData.Rows[i]["D"].ToString() == dtSource.Columns[col].Caption.Split('_')[0].ToString())//date
                        {
                            if (dtSource.Columns[col].Caption.Split('_')[1].ToString() == "Order")
                            {
                                dtSource.Rows[row_opcd][col] = Order;
                            }
                            if (dtSource.Columns[col].Caption.Split('_')[1].ToString() == "Actual")
                            {
                                dtSource.Rows[row_opcd][col] = Actual;
                            }
                            if (dtSource.Columns[col].Caption.Split('_')[1].ToString() == "Percent")
                            {
                                dtSource.Rows[row_opcd][col] = Percent;
                            }
                        }
                    }
                }

                // --TOTAL TRONG LUOI---
                dtSource.Rows.Add("TOTAL");
                for (int j = 1; j <= dtSource.Columns.Count - 1; j++)
                {
                    if (dtSource.Columns[j].Caption.Split('_')[1].ToString() != "Percent")
                    {
                        total_row = 0;
                        for (int i = 0; i <= dtSource.Rows.Count - 1; i++)
                        {
                            double.TryParse(dtSource.Rows[i][j].ToString(), out total_tmp);
                            total_row = total_row + total_tmp;
                        }
                        dtSource.Rows[dtSource.Rows.Count - 1][j] = total_row.ToString();
                    }
                    else if (dtSource.Columns[j].Caption.Split('_')[1].ToString() == "Percent")
                    {
                        if (!string.IsNullOrEmpty(dtSource.Rows[dtSource.Rows.Count - 1][j - 2].ToString()))
                        {
                            if (Convert.ToDouble(dtSource.Rows[dtSource.Rows.Count - 1][j - 2]) > 0)
                                dtSource.Rows[dtSource.Rows.Count - 1][j] = Math.Round(Convert.ToDouble(dtSource.Rows[dtSource.Rows.Count - 1][j - 1]) / Convert.ToDouble(dtSource.Rows[dtSource.Rows.Count - 1][j - 2]) * 100, 0);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("BindingDataSource(): " + ex.Message);
            }
        }

        private void AddColumn(GridBandEx band, string fieldName)
        {
            BandedGridColumnEx col = new BandedGridColumnEx();
            col.FieldName = fieldName;
            col.Visible = true;
            col.OptionsColumn.AllowEdit = false;
            col.OptionsColumn.ReadOnly = true;
            col.OptionsFilter.AllowFilter = false;
            col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            gvwBase.Columns.Add(col);
            col.OwnerBand = band;
        }

        private bool loadchart(DataTable dtchart)
        {
            try
            {
                chart.DataSource = dtchart;
                chart.Series[0].ArgumentDataMember = "YMD"; // cot X
                chart.Series[0].ValueDataMembers.AddRange(new string[] { "OS_PER" }); //COT Y

                chart.Series[1].ArgumentDataMember = "YMD"; // cot X
                chart.Series[1].ValueDataMembers.AddRange(new string[] { "PH_PER" }); //COT Y

                chart.Series[2].ArgumentDataMember = "YMD"; // cot X
                chart.Series[2].ValueDataMembers.AddRange(new string[] { "PU_PER" }); //COT Y

                chart.Series[3].ArgumentDataMember = "YMD"; // cot X
                chart.Series[3].ValueDataMembers.AddRange(new string[] { "IP_PER" }); //COT Y

                chart.Series[4].ArgumentDataMember = "YMD"; // cot X
                chart.Series[4].ValueDataMembers.AddRange(new string[] { "DMP_PER" }); //COT Y

                chart.Series[5].ArgumentDataMember = "YMD"; // cot X
                chart.Series[5].ValueDataMembers.AddRange(new string[] { "TOTAL_PER" }); //COT Y

                chart.Series[6].ArgumentDataMember = "YMD"; // cot X
                chart.Series[6].ValueDataMembers.AddRange(new string[] { "TAR_PER" }); //COT Y

                ((XYDiagram)chart.Diagram).AxisX.Label.Staggered = false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private void btnTimeContraint_Click(object sender, EventArgs e)
        {
            try
            {
                if (SendYN(((Button)sender).Text))
                {

                    RunTimeContraint("Stockfit", "Q2"); //STOCKFIT
                }
            }
            catch
            {
            }
        }

        private void btnTimeContraintBottom_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
            {
                RunTimeContraint("Bottom", "Q1"); //BOTTOM
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
                    string nRowSpan = "<td align='center'>{0}</td>";
                    if (!string.IsNullOrEmpty(dtData.Rows[i]["REASON"].ToString()))
                    {
                        nRowSpan = "<td align='center' rowspan = '2'>{0}</td>";
                        body1 = string.Format("<tr style='color:red;font-weight:bold;border-style:dotted;'><td colspan = '2' align = 'center'>Reason Of Plant {0} </td><td colspan = '35'>{1}</td></tr> ", dtData.Rows[i]["PLANT"], dtData.Rows[i]["REASON"]);
                    }
                    body += string.Format(@"</tr>
                               <tr align='right' style='font-weight:bold;'>"
                                + nRowSpan + @"
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

        private string getHTMLBodyHeaderTMSDashv2(DataTable dtHead, DataTable dtData)
        {
            try
            {
                string style = System.IO.File.ReadAllText(Application.StartupPath + "\\tmsdashstylev2.txt");
                string header1 = System.IO.File.ReadAllText(Application.StartupPath + "\\tmsdashhtmlv2.txt");
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
                    string nRowSpan = "<td align='center'>{0}</td>";
                    if (!string.IsNullOrEmpty(dtData.Rows[i]["REASON"].ToString()))
                    {
                        nRowSpan = "<td align='center' rowspan = '2'>{0}</td>";
                        body1 = string.Format("<tr style='color:red;font-weight:bold;border-style:dotted;'><td colspan = '2' align = 'center'>Reason Of Plant {0} </td><td colspan = '35'>{1}</td></tr> ", dtData.Rows[i]["PLANT"], dtData.Rows[i]["REASON"]);
                    }
                    body += string.Format(@"</tr>
                               <tr align='right' style='font-weight:bold;'><td align='center'>{0}</td>
                               <td align='center'>{1}</td>
                               <td align='center' bgcolor='{40}' style='color:{41}'>{2}</td>
                               <td bgcolor='#fff4b0'>{3}</td>
                               <td>{4}</td>
                               <td bgcolor='{38}' style='color:{39}'>{5}</td>
                               <td style='color:red'>{6}</td>
                               <td>{7}</td>

                               <td align='left'>{42} </td>
                              </tr>", argsBody);
                }
                object[] argsHeader = new object[dtHead.Rows.Count + 1];
                for (int i = 0; i < dtHead.Rows.Count; i++)
                {
                    argsHeader[i] = dtHead.Rows[i]["CAPTION"].ToString();
                }
                argsHeader[6] = dtHead.Rows[0]["REMARKS"].ToString();
                string end = @"</table><hr>";
                string remakeHeader = string.Format(header1, argsHeader);

                string HTML = string.Concat(style, remakeHeader, body, end);
                return HTML;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private string getHTMLBodyHeaderTMSDashv3(DataTable dtHead, DataTable dtData)
        {
            try
            {
                string style = System.IO.File.ReadAllText(Application.StartupPath + "\\TMS_DAAS_CSS.txt");
                string headertable = System.IO.File.ReadAllText(Application.StartupPath + "\\TMS_DAAS_HEADER_TABLE.txt");
                object[] argsHeader = new object[dtHead.Columns.Count * dtHead.Rows.Count];

                int HeaderiDx = 0;
                string headerTotal = string.Empty, headerDivided = string.Empty;
                for (int iCol = 0; iCol < dtHead.Columns.Count; iCol++)
                {
                    if (iCol > 0)
                        for (int iRow = 0; iRow < dtHead.Rows.Count; iRow++)
                        {
                            argsHeader[HeaderiDx] = dtHead.Rows[iRow][iCol].ToString();
                            HeaderiDx++;
                        }
                }
                DataTable dtHeaderTotal = new DataTable();
                //Create Header Total
                if (dtData.Select("FACTORY = 'TOTAL' AND PLANT = 'TOTAL'").Count() > 0)
                    dtHeaderTotal = dtData.Select("FACTORY = 'TOTAL' AND PLANT = 'TOTAL'").CopyToDataTable();

                for (int iHead = 0; iHead < dtHeaderTotal.Columns.Count; iHead++)
                {
                    if (iHead > 3 && iHead < dtHeaderTotal.Columns["REASON"].Ordinal)
                    {
                        if (dtHeaderTotal.Columns[iHead].ToString().Equals("EMPTY_COL"))
                        {
                            headerTotal += $"<td style='width: 1px; padding: 1px;border: blue;' bgcolor='blue'></td>";
                        }
                        else
                        {
                            if (dtHeaderTotal.Columns[iHead].ToString().Contains("ORDR"))
                                headerTotal += $"<td style='color:{dtHeaderTotal.Rows[0][dtHeaderTotal.Columns[string.Concat(dtHeaderTotal.Columns[iHead].ColumnName, "_F_COLOR")]]}' bgcolor='{dtHeaderTotal.Rows[0][dtHeaderTotal.Columns[string.Concat(dtHeaderTotal.Columns[iHead].ColumnName, "_B_COLOR")]]}' class='tftable-rlax'>{string.Format("{0:n0}", dtHeaderTotal.Rows[0][iHead])}</td>";
                            else
                                headerTotal += $"<td class='tftable-rlax'>{string.Format("{0:n0}", dtHeaderTotal.Rows[0][iHead])}</td>";
                        }
                    }

                    if (iHead < dtHeaderTotal.Columns["REASON"].Ordinal)
                        headerDivided += "<td style='height: 1px; padding: 1px;border: blue;' bgcolor='blue'></td>";
                }
                argsHeader[38] = @"<tr>" + headerDivided + "</tr>";
                argsHeader[39] = @"<tr>" + headerTotal + "</tr>";
                headertable = string.Format(headertable, argsHeader);
                string body = string.Empty;

                //Create Body
                DataTable dtDataFillter = new DataTable();
                if (dtData.Select("FACTORY <> 'TOTAL' AND PLANT <> 'TOTAL'").Count() > 0)
                    dtDataFillter = dtData.Select("FACTORY <> 'TOTAL' AND PLANT <> 'TOTAL'").CopyToDataTable();
                for (int iRowData = 0; iRowData < dtDataFillter.Rows.Count; iRowData++)
                {
                    string bodyTD = string.Empty;
                    for (int iColData = 0; iColData < dtDataFillter.Columns.Count; iColData++)
                    {
                        if (iColData > 0 && iColData <= dtDataFillter.Columns["REASON"].Ordinal)
                        {
                            if (dtDataFillter.Columns[iColData].ToString().Equals("EMPTY_COL"))
                                bodyTD += "<td style='width: 1px; padding: 1px;border-width: 0px;border: blue;' bgcolor='blue'></td>";
                            else
                            {
                                if (dtDataFillter.Columns[iColData].ToString().Contains("ORDR") || dtDataFillter.Columns[iColData].ToString().Equals("THIS_RANK"))
                                    if (dtDataFillter.Columns[iColData].ToString().Equals("THIS_RANK"))
                                        bodyTD += $"<td style='color:{dtDataFillter.Rows[iRowData][dtDataFillter.Columns[string.Concat(dtHeaderTotal.Columns[iColData].ColumnName, "_F_COLOR")]]}' bgcolor='{dtDataFillter.Rows[iRowData][dtDataFillter.Columns[string.Concat(dtDataFillter.Columns[iColData].ColumnName, "_B_COLOR")]]}' class='tftable-clax'>{string.Format("{0:n0}", dtDataFillter.Rows[iRowData][iColData])}</td>";
                                    else
                                        bodyTD += $"<td style='color:{dtDataFillter.Rows[iRowData][dtDataFillter.Columns[string.Concat(dtHeaderTotal.Columns[iColData].ColumnName, "_F_COLOR")]]}' bgcolor='{dtDataFillter.Rows[iRowData][dtDataFillter.Columns[string.Concat(dtDataFillter.Columns[iColData].ColumnName, "_B_COLOR")]]}' class='tftable-rlax'>{string.Format("{0:n0}", dtDataFillter.Rows[iRowData][iColData])}</td>";
                                else
                                if (dtDataFillter.Columns[iColData].ToString().Equals("PLANT") || dtDataFillter.Columns[iColData].ToString().Equals("LAST_RANK"))
                                    bodyTD += $"<td class='tftable-clax'>{string.Format("{0:n0}", dtDataFillter.Rows[iRowData][iColData])}</td>";
                                else if (dtDataFillter.Columns[iColData].ToString().Equals("REASON"))
                                    bodyTD += $"<td class='tftable-llax'>{string.Format("{0:n0}", dtDataFillter.Rows[iRowData][iColData])}</td>";
                                else
                                    bodyTD += $"<td class='tftable-rlax'>{string.Format("{0:n0}", dtDataFillter.Rows[iRowData][iColData])}</td>";
                            }
                        }
                    }
                    body += string.Format("<tr>{0}</tr>", bodyTD);
                }
                string end = "</tbody></table><hr>";
                string HTML = string.Concat(style, headertable, body, end);
                return HTML;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private string getHTMLBodyHeaderTMSSummary(DataTable dtHead, DataTable dtData)
        {
            try
            {
                string style, sdate = string.Empty, headertable, body = string.Empty;
                style = System.IO.File.ReadAllText(Application.StartupPath + "\\TMS_DAAS_CSS.txt");
                body = string.Empty;
                bool isAppend = false;
                int iDx = 0;
                foreach (DataRow dr in dtData.Rows)
                {
                    if (iDx == 5)
                    {
                        body += "<tr><td style='border:none;padding: 3px;' bgcolor=silver colspan =14></td></tr>";
                        body += $"<tr><td>{dr["PLANT"]}</td>" +
                             $"<td bgcolor='{dr["I01_BCOLOR"]}' style ='color:{dr["I01_FCOLOR"]}'>{dr["I01"]}</td>" +
                             $"<td bgcolor='{dr["I02_BCOLOR"]}' style ='color:{dr["I02_FCOLOR"]}'>{dr["I02"]}</td>" +
                             $"<td bgcolor='{dr["I03_BCOLOR"]}' style ='color:{dr["I03_FCOLOR"]}'>{dr["I03"]}</td>" +
                             $"<td bgcolor='{dr["I04_BCOLOR"]}' style ='color:{dr["I04_FCOLOR"]}'>{dr["I04"]}</td>" +
                             $"<td bgcolor='{dr["I05_BCOLOR"]}' style ='color:{dr["I05_FCOLOR"]}'>{dr["I05"]}</td>" +
                             $"<td bgcolor='{dr["I06_BCOLOR"]}' style ='color:{dr["I06_FCOLOR"]}'>{dr["I06"]}</td>" +
                             $"<td width=5px bgcolor=silver></td>" +
                             $"<td bgcolor='{dr["W01_BCOLOR"]}' style ='color:{dr["W01_FCOLOR"]}'>{dr["W01"]}</td>" +
                             $"<td bgcolor='{dr["W02_BCOLOR"]}' style ='color:{dr["W02_FCOLOR"]}'>{dr["W02"]}</td>" +
                             $"<td bgcolor='{dr["W03_BCOLOR"]}' style ='color:{dr["W03_FCOLOR"]}'>{dr["W03"]}</td>" +
                             $"<td bgcolor='{dr["W04_BCOLOR"]}' style ='color:{dr["W04_FCOLOR"]}'>{dr["W04"]}</td>" +
                             $"<td bgcolor='{dr["W05_BCOLOR"]}' style ='color:{dr["W05_FCOLOR"]}'>{dr["W05"]}</td>" +
                             $"<td bgcolor='{dr["W06_BCOLOR"]}' style ='color:{dr["W06_FCOLOR"]}'>{dr["W06"]}</td></tr>";
                    }
                    else
                    {
                        body += $"<tr><td>{dr["PLANT"]}</td>" +
                             $"<td bgcolor='{dr["I01_BCOLOR"]}' style ='color:{dr["I01_FCOLOR"]}'>{dr["I01"]}</td>" +
                             $"<td bgcolor='{dr["I02_BCOLOR"]}' style ='color:{dr["I02_FCOLOR"]}'>{dr["I02"]}</td>" +
                             $"<td bgcolor='{dr["I03_BCOLOR"]}' style ='color:{dr["I03_FCOLOR"]}'>{dr["I03"]}</td>" +
                             $"<td bgcolor='{dr["I04_BCOLOR"]}' style ='color:{dr["I04_FCOLOR"]}'>{dr["I04"]}</td>" +
                             $"<td bgcolor='{dr["I05_BCOLOR"]}' style ='color:{dr["I05_FCOLOR"]}'>{dr["I05"]}</td>" +
                             $"<td bgcolor='{dr["I06_BCOLOR"]}' style ='color:{dr["I06_FCOLOR"]}'>{dr["I06"]}</td>" +
                             $"<td width=5px bgcolor=silver></td>" +
                             $"<td bgcolor='{dr["W01_BCOLOR"]}' style ='color:{dr["W01_FCOLOR"]}'>{dr["W01"]}</td>" +
                             $"<td bgcolor='{dr["W02_BCOLOR"]}' style ='color:{dr["W02_FCOLOR"]}'>{dr["W02"]}</td>" +
                             $"<td bgcolor='{dr["W03_BCOLOR"]}' style ='color:{dr["W03_FCOLOR"]}'>{dr["W03"]}</td>" +
                             $"<td bgcolor='{dr["W04_BCOLOR"]}' style ='color:{dr["W04_FCOLOR"]}'>{dr["W04"]}</td>" +
                             $"<td bgcolor='{dr["W05_BCOLOR"]}' style ='color:{dr["W05_FCOLOR"]}'>{dr["W05"]}</td>" +
                             $"<td bgcolor='{dr["W06_BCOLOR"]}' style ='color:{dr["W06_FCOLOR"]}'>{dr["W06"]}</td></tr>";
                    }
                    iDx++;
                }
                string DivTemp = "I";
                foreach (DataRow dr in dtHead.Rows)
                {
                    if (!DivTemp.Equals(dr["DIV"].ToString()))
                    {
                        sdate += $"<th></th><th>{dr["CAPTION"]}</th>";
                        DivTemp = dr["DIV"].ToString();
                    }
                    else
                    {
                        sdate += $"<th>{dr["CAPTION"]}</th>";
                        DivTemp = dr["DIV"].ToString();
                    }
                }
                string InOrder = string.Empty, WithOutOrder = string.Empty;
                InOrder = dtHead.Rows[0]["OIO"].ToString();
                WithOutOrder = dtHead.Rows[0]["OWO"].ToString();
                headertable = @"
                                    <table class='greyGridTable'>
                                    <thead>
                                    <tr>
                                    <th rowspan='2'>Plant</th>
                                    <th colspan='6'>
                                        <table class='infoTable' >
                                        <tbody >
                                        <tr><td style='border:none; border-bottom: 2px solid #0043fa;' colspan='4'><p style='color:#0043fa'>Outgoing In Order</p></td></tr>
                                        <tr>" + InOrder + @"</tr>
                                        </tbody>
                                    </table></th>
                                    <th></th>
                                    <th colspan='6'><table class='infoTable' >
                                        <tbody >
                                        <tr><td style='border:none;border-bottom: 2px solid #0043fa;' colspan='4'><p style='color:#ff0000'>Without Order</p></td></tr>
                                        <tr>" + WithOutOrder + @"</tr>
                                        </tbody>
                                    </table></th>
                                    </tr>
                                    <tr>" + sdate + @"
                                    </tr>
                                    </thead>

                                    <tbody> " +
                                    body + @"
                                    </tbody>
                                    </table></html>";

                string HTML = string.Concat(style, headertable);
                return HTML;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private string getHTMLBodyHeaderScada(DataTable dtHead, DataTable dtData)
        {
            try
            {
                string style, headertable, body, end;
                style = System.IO.File.ReadAllText(Application.StartupPath + "\\TMS_DAAS_CSS.txt");
                headertable = @"<body><div style='-webkit-border-radius: 5px; -moz-border-radius: 5px; border-radius: 5px; color: #029c3d; display: block;'>
                                          <span>Top 50 Scada Machine <b>Abnormal Temperature</b></span><br>
                                      <hr><div class='info' style = 'color:black'>
                                          <b>Thiết bị nóng</b><br>
                                            <ul>
                                          <li>Tiêu chuẩn nhiệt độ cài đặt là <b>+-3</b></li>
                                          <li>Nếu nhiệt độ vượt quá tiêu chuẩn, hệ thống sẽ thông báo <b style='color:yellow; background-color:black'>màu vàng</b></li>
                                          <li>Nếu nhiệt độ vượt quá thời gian 3 phút, hệ thống sẽ thông báo <b style='color:red; background-color:black'>màu đỏ</b></li>
                                           </ul>
                                          <b>Thiết bị lạnh</b><br>
                                          <ul>
                                          <li>Tiêu chuẩn nhiệt độ cài đặt là <b>+10</b> và <b>-12</b></li>
                                          <li>Nếu nhiệt độ vượt quá tiêu chuẩn, hệ thống sẽ thông báo <b style='color:yellow; background-color:black'>màu vàng</b></li>
                                          <li>Nếu nhiệt độ vượt quá thời gian 5 phút, hệ thống sẽ thông báo <b style='color:red; background-color:black'>màu đỏ</b></li>
                                          </ul>
                                        </div>
                                        <hr>
                                      </div>
                                    <table class='tftable2' border='1' width='100%' cellspacing='0' cellpadding='0'>
                                    <thead>
                                    <tr style='font-weight: bold;'>
                                    <td style='width: auto; font-weight: bolder; font-size: 18px;'   align='center'>Date</td>
                                    <td style='width: auto; font-weight: bolder; font-size: 18px;'   align='center'>Plant</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Line</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Process</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Machine Name</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Machine Code</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Machine ID</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Description</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>PV</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Min</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Max</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>I/F Data Count</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Abnormal</td>
                                    <td style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Abnormal Ratio</td>
                                    <td class='bg-red' style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Alert (Minutes)</td>
                                    <td class='bg-red' style='width: auto; font-weight: bolder;  font-size: 18px;'  align='center'>Alert (Times)</td>
                                    </tr>
                                    </thead>
                                    <tbody>";
                body = string.Empty;
                int iDx = 0;
                foreach (DataRow dr in dtData.Rows)
                {
                    string rowColor = string.Empty;
                    if (iDx % 2 == 0)
                        rowColor = "#ededed";
                    else
                        rowColor = "white";
                    body += $"<tr style='background-color:{rowColor}'>" +
                              $"<td class='tftable2-clax'>{dr["YMD"]}</td>" +
                              $"<td class='tftable2-clax'>{dr["PLANT"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["LINE"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["OP_CD"]}</td>" +
                              $"   <td class='tftable2-llax'>{dr["MACHINE_NAME"]}</td>" +
                              $"   <td class='tftable2-llax'>{dr["MACHINE_CODE"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["MACHINE_ID"]}</td>" +
                              $"   <td class='tftable2-llax'>{dr["DESCR"]}</td>" +
                              $"   <td class='tftable2-clax'>{ dr["PV"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["MIN_VALUE"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["MAX_VALUE"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["TOTAL"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["OVER"]}</td>" +
                              $"   <td class='tftable2-clax'>{string.Concat(dr["RATE"], "%")}</td>" +
                              $"   <td class='tftable2-clax'>{dr["ALARM_TIME_M"]}</td>" +
                              $"   <td class='tftable2-clax'>{dr["ALARM_TIME_C"]}</td></tr>";
                    iDx++;
                }

                end = "</tbody></table><hr></body></html>";

                string HTML = string.Concat(style, headertable, body, end);
                return HTML;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void gvwBase_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.ColumnHandle >= 1)
            {
                if (e.CellValue.ToString() == "") return;
                e.DisplayText = double.Parse(e.CellValue.ToString()).ToString("###,##0");
                //e.DisplayText = (((int)(Convert.ToDecimal(e.CellValue) * 100))).ToString() + "%";
            }
        }

        private void gvwBase_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            try
            {
                double temp = 0.0;
                if (gvwBase.GetRowCellValue(e.RowHandle, gvwBase.Columns[0]).ToString().Contains("Bottom Performance"))
                {
                    e.Appearance.BackColor = Color.Cyan;
                }

                if ((e.Column.FieldName.Contains("Percent")) && (e.CellValue != null))
                {
                    double.TryParse(e.CellValue.ToString(), out temp); //out

                    if (temp > 0 && temp < 50)
                    {
                        e.Appearance.BackColor = Color.Black;
                        e.Appearance.ForeColor = Color.White;
                    }
                    else if (temp >= 50 && temp < 80)
                    {
                        e.Appearance.BackColor = Color.Red;
                        e.Appearance.ForeColor = Color.White;
                    }
                    else if (temp >= 80 && temp < 95)
                    {
                        e.Appearance.BackColor = Color.Yellow;
                        e.Appearance.ForeColor = Color.Black;
                    }
                    else if (temp >= 95)
                    {
                        e.Appearance.BackColor = Color.LightGreen;
                        e.Appearance.ForeColor = Color.Black;
                    }

                    if (e.CellValue.ToString() == "0")
                    {
                        e.Appearance.BackColor = Color.Red;
                        e.Appearance.ForeColor = Color.White;
                    }
                }
            }
            catch
            {
            }
        }

        private void tmrLoad2_Tick(object sender, EventArgs e)
        {
            //Each 5 minutes.

            RunFeedback("Q"); //RunFeedback("U");
        }

        private void btnRunScada_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunScada("Q");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunTMS_Summary("Q");
        }

        private void btnFeedback_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
            {
                RunFeedback("Q");
            }
        }

        private void cmd_Quality_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunQuality("Q");
        }

        private void cmd_Quality2_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunQuality2("Q");
        }

        private void cmd_BotDef_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunBotDef("Q");
        }

        private void cmdMoldRepairMonth_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunMoldRepairMonth("Q");
        }

        private void cmdMoldRepairMonthWh_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunMoldRepairMonthWh("Q");
        }

        private void btnRunQualityMonth_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunQualityMonth("Q");
        }

        private void gvwView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            try
            {

                string sTotItems = gvwView.GetRowCellValue(e.RowHandle, gvwView.Columns["ITEMS"]).ToString().ToUpper();
                if (sTotItems.Equals("PPM"))
                {
                    e.Appearance.BackColor = Color.FromArgb(70, 244, 250);
                    e.Appearance.ForeColor = Color.Black;
                }
                if (sTotItems.Equals("B/C GRADE"))
                {
                    e.Appearance.BackColor = Color.FromArgb(252, 217, 98);
                    e.Appearance.ForeColor = Color.Black;
                }
                if (sTotItems.Equals("REWORK RATE(%)"))
                {
                    e.Appearance.BackColor = Color.FromArgb(252, 217, 98);
                    e.Appearance.ForeColor = Color.Black;
                }
                if (sTotItems.Equals("PRODUCTION"))
                {
                    e.Appearance.BackColor = Color.FromArgb(212, 255, 254);
                    e.Appearance.ForeColor = Color.Black;
                }
                if (e.Column.FieldName.ToUpper().Equals("TOTAL"))
                {
                    e.Appearance.BackColor = Color.FromArgb(252, 252, 194);
                }


            }
            catch (Exception)
            {

                throw;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cmdAssInline_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                //RunAssInLine_v2("Q");
                RunAssInLine("Q");
        }

        private void cmdAssInvSet_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunAssInLine_v2("Q");
        }

        private void cmdSumDaaS_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunSumDaaS("Q");
        }

        private void RunSumDaaS(string argType)
        {
            try
            {
                if (_isRun2) return;

                _isRun2 = true;

                DataSet dsData = SEL_SUM_OPEN_DAAS(argType, DateTime.Now.ToString("yyyyMMdd")); //Get Data for HTML Table

                if (dsData == null) return;
                WriteLog($"RunSumDaaS({argType}): BEGIN ");
                DataTable dtData = dsData.Tables[0];
                DataTable dtData2 = dsData.Tables[1];
                DataTable dtEmail = dsData.Tables[2];
                string Subject = dtData2.Rows[0]["SUBJECT"].ToString();
                WriteLog("  " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                CreateMailSumOpenDaaS(dtData, Subject, dtEmail);
                WriteLog($"RunSumDaaS({argType}): END ");
            }
            catch (Exception ex)
            {
                WriteLog($"  RunSumDaaS({argType}) " + ex.ToString());
            }
            finally
            {
                _isRun2 = false;
            }
        }
        private void RunPORegReport(string argType)
        {
            DataSet ds = SEL_DATA_PO_REG_REPORT(argType);//MACHINE TIMES
            if (ds == null) return;
            using (Releif_AVSM_Report_Daily f = new Releif_AVSM_Report_Daily())
            {
                f._chkTest = chkTest.Checked;
                f._subject = "Workforce Management Systems Report Daily (" + DateTime.Now.ToString("yyyy/MM/dd") + ")";
                f._dtData = ds.Tables[0];
                f._dtEmail = ds.Tables[1];
                f.Show();
                f.SendToBack();
            }

            WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunPORegReport({argType}): END");
        }
        public static int GetIso8601WeekOfYear(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        #region RunWeeklyBC
        private void RunWeeklyBC(string argType, string argDate)
        {
            DataTable dtChart1 = SEL_DATA_WEEKLY_B_C("CHART1", argDate);//BOTTOM CONSTRAINT
            DataTable dtChart2 = SEL_DATA_WEEKLY_B_C("CHART2", argDate);//LINE
            DataTable dtData2 = SEL_DATA_WEEKLY_B_C("GRID2", argDate);
            DataTable dtData21 = SEL_DATA_WEEKLY_B_C("GRID21", argDate);
            //DataTable dtChart3 = SEL_DATA_OS_MACHINE_MONTHLY("CHART3", argDate);//REASON
            //DataTable dtChart4 = SEL_DATA_OS_MACHINE_MONTHLY("CHART4", argDate);// HOURS
            //DataTable dtChart5 = SEL_DATA_OS_MACHINE_MONTHLY("CHART5", argDate);//SHIFT
            //DataTable dtChart6 = SEL_DATA_OS_MACHINE_MONTHLY("CHART6", argDate);//DAILY
            //DataTable dtEmail = SEL_DATA_OS_MACHINE_MONTHLY("EMAIL", argDate); //Email Send

            //if (dtChart1 == null || dtChart2 == null || dtChart3 == null || dtChart4 == null || dtChart5 == null || dtChart6 == null || dtEmail == null)
            //    return;
            // WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunOSMonthly({argType}): BEGIN");

            using (frmWeekly_Bottom_Constraint frmWeeklyBC = new frmWeekly_Bottom_Constraint())
            {
                frmWeeklyBC._chkTest = chkTest.Checked;
                frmWeeklyBC._subject = "Weekly Bottom Constraint (" + DateTime.Now.AddMonths(-1).ToString("yyyy") + "/" + DateTime.Now.ToString("MMM") + "/" + GetIso8601WeekOfYear(DateTime.Now.AddDays(-2)).ToString() + ")";
                frmWeeklyBC._dtChart1 = dtChart1;
                frmWeeklyBC._dtChart2 = dtChart2;
                frmWeeklyBC.dtGrid2 = dtData2;
                frmWeeklyBC._dtChart21 = dtData2;
                frmWeeklyBC.dtGrid21 = dtData21;
                //frmWeeklyBC._dtChart5 = dtChart5;
                //frmWeeklyBC._dtChart6 = dtChart6;
                //frmWeeklyBC._dtEmail = dtEmail;
                frmWeeklyBC.Show();
                frmWeeklyBC.SendToBack();

            }

            // frmWeekly_Bottom_Constraint frmWeeklyBC = new frmWeekly_Bottom_Constraint();

            // frmWeeklyBC._chkTest = chkTest.Checked;
            // // frmWeeklyBC._subject = "Weekly Bottom Constraint (" + DateTime.Now.AddMonths(-1).ToString("yyyy") + "/" + DateTime.Now.ToString("MMM") + "/" + GetIso8601WeekOfYear(DateTime.Now.AddDays(-2)).ToString() + ")";
            // frmWeeklyBC._dtChart1 = dtChart1;
            // frmWeeklyBC._dtChart2 = dtChart2;
            // //frmWeeklyBC._dtChart3 = dtChart3;
            // //frmWeeklyBC._dtChart4 = dtChart4;
            // //frmWeeklyBC._dtChart5 = dtChart5;
            // //frmWeeklyBC._dtChart6 = dtChart6;
            // //frmWeeklyBC._dtEmail = dtEmail;
            // frmWeeklyBC.Show();
            //// frmWeeklyBC.SendToBack();



            // WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunMoldRepairMonth({argType}): END");
        }
        #endregion RunWeeklyBC
        private void RunOSMonthly(string argType, string argDate)
        {
            DataTable dtChart1 = SEL_DATA_OS_MACHINE_MONTHLY("CHART1", argDate);//MACHINE TIMES
            DataTable dtChart2 = SEL_DATA_OS_MACHINE_MONTHLY("CHART2", argDate);//LINE
            DataTable dtChart3 = SEL_DATA_OS_MACHINE_MONTHLY("CHART3", argDate);//REASON
            DataTable dtChart4 = SEL_DATA_OS_MACHINE_MONTHLY("CHART4", argDate);// HOURS
            DataTable dtChart5 = SEL_DATA_OS_MACHINE_MONTHLY("CHART5", argDate);//SHIFT
            DataTable dtChart6 = SEL_DATA_OS_MACHINE_MONTHLY("CHART6", argDate);//DAILY
            DataTable dtEmail = SEL_DATA_OS_MACHINE_MONTHLY("EMAIL", argDate); //Email Send

            if (dtChart1 == null || dtChart2 == null || dtChart3 == null || dtChart4 == null || dtChart5 == null || dtChart6 == null || dtEmail == null)
                return;
            WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunOSMonthly({argType}): BEGIN");

            using (frmWeekly_Bottom_Constraint frmOsMonthly = new frmWeekly_Bottom_Constraint())
            {
                frmOsMonthly._chkTest = chkTest.Checked;
                frmOsMonthly._subject = "Monthly Os Press Machine Drawback (" + DateTime.Now.AddMonths(-1).ToString("yyyy") + "/" + DateTime.Now.ToString("MMM") + "/" + GetIso8601WeekOfYear(DateTime.Now.AddDays(-2)).ToString() + ")";
                frmOsMonthly._dtChart1 = dtChart1;
                frmOsMonthly._dtChart2 = dtChart2;
                frmOsMonthly._dtChart3 = dtChart3;
                frmOsMonthly._dtChart4 = dtChart4;
                frmOsMonthly._dtChart5 = dtChart5;
                frmOsMonthly._dtChart6 = dtChart6;
                frmOsMonthly._dtEmail = dtEmail;
                frmOsMonthly.Show();
                frmOsMonthly.SendToBack();
            }

            WriteLog($"{DateTime.Now:yyyy-MM-dd hh:mm:ss} RunMoldRepairMonth({argType}): END");
        }




        private void btnRunWeekly_Bottom_Constraint_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunWeeklyBC("Q", "20220314");
        }

        private void btnUpperInv_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
            {
                //Upper Inventory
                if (btnRunUpperInvChk.Checked)
                    RunUpperInv("Q1", DateTime.Now.ToString("yyyyMMdd"));
            }
        }

        private void btnMonthlyBottomAnalysis_Click(object sender, EventArgs e)
        {
            if (SendYN(((Button)sender).Text))
                RunMonthBottomAnalysis("Q");


        }

        private void RunMonthBottomAnalysis(string ARG_TYPE)
        {
            try
            {
                //Prepairing Data
                string ARG_DATE = DateTime.Now.ToString("yyyyMMdd");
                DataTable dtChart = SEL_DATA_MONTHLY_BOTTOM_ANALYSIS(ARG_TYPE, ARG_DATE);//BOTTOM CONSTRAINT
                Monthly_Bottom_Analysis f = new Monthly_Bottom_Analysis();
                f._dtChart = dtChart;
                f._subject = "Phuoc Test Email Bottom Inventory Analysis by Monthly";
                f.Show();
                //using (Monthly_Bottom_Analysis f = new Monthly_Bottom_Analysis())
                //{
                //    f._dtChart = dtChart;
                //    f._subject = "Phuoc Test Email Bottom Inventory Analysis by Monthly";
                //    f.Show();
                //    f.SendToBack();
                //}
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }
        }



        private string getHTMLBodyHeaderTimeContraint(string Qtype, DataTable dtHead, DataTable dtData)
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
			                                            .tg  {font-size:12px; width:100%;border-width: 1px; border-collapse: collapse;}
                                                        .tg td{font-size:12px;font-family:Calibri;bgcolor:#ffffff;border-width: 1px;padding: 8px;border-style: solid;border-color: #9dcc7a;white-space: nowrap}
                                                        .tg th{border-color:#9dcc7a;border-style:solid;border-width:1px;background-color:#0080a0;color: #ffffff; font-family:Arial, sans-serif;font-size:12px;
                                                          font-weight:bold;overflow:hidden;padding:5px 10px;word-break:normal;}
                                                        .tg .tg-0lax{text-align:center;}
                                                        .tg .tg-1lax{text-align:left;}
                                                        .tg .tg-2lax{text-align:right;}
                                                        .tg .tg-eslapse{text-align:center;background-color:yellow;color: #000000;}
                                                        span {
                                                        color: #2e5f82;
                                                        display: inline-block;
                                                        padding: 3px 10px;
                                                        border-radius: 5px;
                                                        font-family: 'Times New Roman', Times, serif;
                                                        font-size: 25px;
                                                        font-style: italic;
                                                        }
                                                         </style>";
                string Colspan = dtData.Rows[0]["COLSPAN"].ToString();
                sHeader1 += string.Format(@"
                                                    </head>
                                                  <body>

                                             <hr>
                                                            <table class='tg'>
                                                            <thead>
                                                              <tr>
                                                                <th class='tg-0lax' rowspan='2'>Process</th>
                                                                <th class='tg-0lax' rowspan='2'>Assembly Date</th>
                                                                <th class='tg-0lax' rowspan='2'>Production Date</th>
                                                                <th class='tg-0lax' rowspan='2'>Elapsed Time</th>
                                                                <th class='tg-0lax' rowspan='2'>Division</th>
                                                                <th class='tg-0lax' rowspan='2'>Plant</th>
                                                                <th class='tg-0lax' rowspan='2'>Assembly Line</th>
                                                                <th class='tg-0lax' rowspan='2'>Mini Line</th>
                                                                <th class='tg-0lax' rowspan='2'>Style Name</th>
                                                                <th class='tg-0lax' rowspan='2'>Style Code</th>
                                                                <th class='tg-0lax' rowspan='2'>Location</th>
                                                                <th class='tg-0lax' colspan='{0}'>Size</th>
                                                                <th class='tg-0lax' rowspan='2'>Total</th>
                                                                <th class='tg-0lax' rowspan='2'>Reason</th>
				                                              </tr><tr>", Colspan);
                DataView dtView = new DataView(dtData);
                dtView.Sort = "SIZE_NO";
                dtHead = dtView.ToTable(true, "SIZE_CODE");
                int iDx = 0;
                for (int j = 0; j < dtHead.Rows.Count; j++)
                {
                    sHeader2 += $"<th class='tg-0lax'>{dtHead.Rows[j]["SIZE_CODE"]}</th>";
                    iDx = j + 16;
                    sBody2 += @"<td class='tg-2lax' style = 'font-weight:bold;background-color:{99}'>{" + iDx + "}</td>";
                }
                sHeader3 = @"</tr></thead><tbody>";

                DataView view = dtData.DefaultView;
                view.Sort = "SIZE_NO,PROC_CODE,STYLE_CODE";
                DataTable dt = view.ToTable();
                dt.Columns.Remove(dt.Columns["SIZE_NO"]);
                DataTable dtGrid = Pivot(dt, dt.Columns["SIZE_CODE"], dt.Columns["QTY"]);
                DataView dtViewGrid = new DataView(dtGrid);
                if (Qtype.Equals("Q1"))
                    dtViewGrid.Sort = "PROC_CODE, ELAPSE_TIME DESC, FA_DATE, PLANT_CODE,ERP_FA_WC_CD,ERP_FA_MLINE_CD,STYLE_CODE,LOCATE";
                else
                    dtViewGrid.Sort = "PROC_CODE, PROC_NAME, ELAPSE_TIME DESC, FA_DATE, PLANT_CODE,ERP_FA_WC_CD,ERP_FA_MLINE_CD,STYLE_CODE,LOCATE";

                dtGrid = dtViewGrid.ToTable();
                object[] argBodys = new object[100];

                string ProcCodeTmp = string.Empty;
                string PlantCodeTmp = string.Empty;
                //Loop Rows
                for (int iRow = 0; iRow < dtGrid.Rows.Count; iRow++)
                {
                    //Loop Cols
                    for (int iCol = 0; iCol < dtGrid.Columns.Count; iCol++)
                    {
                        if (iCol >= 2)
                        {
                            //if (dtGrid.Columns[iCol].ColumnName.Equals("TOTAL"))
                            argBodys[iCol - 2] = string.Format("{0:n0}", dtGrid.Rows[iRow][iCol]);
                            //else
                            //    argBodys[iCol - 2] = dtGrid.Rows[iRow][iCol].ToString();
                        }
                    }
                    string BodyProcName = string.Empty;
                    string BodyPlantName = string.Empty;
                    if (dtGrid.Rows[iRow]["PROC_NAME"].ToString().ToUpper().Equals("TOTAL"))
                    {
                        argBodys[99] = "#ffffc9";
                        BodyProcName = string.Format("<td class='tg-0lax' colspan = '11' style='font-weight:bold;background-color:#0080a0;color: #ffffff;'>{0}</td>", dtGrid.Rows[iRow]["PROC_NAME"]);
                        ProcCodeTmp = string.Empty;
                        PlantCodeTmp = string.Empty;
                    }
                    else
                    {
                        argBodys[99] = "";
                        //  BodyProcName = string.Format("<td class='tg-0lax' rowspan='1'>{1}</td>", dtGrid.Rows[iRow]["ROWSPAN1"], dtGrid.Rows[iRow]["PROC_NAME"]);
                        if (!ProcCodeTmp.Equals(dtGrid.Rows[iRow]["PROC_NAME"].ToString()))
                        {
                            BodyProcName = string.Format("<td class='tg-0lax' rowspan='{0}'>{1}</td>", dtGrid.Rows[iRow]["ROWSPAN1"], dtGrid.Rows[iRow]["PROC_NAME"]);
                        }
                        if (!dtGrid.Rows[iRow]["ROWSPAN1"].ToString().Equals("1"))
                            ProcCodeTmp = dtGrid.Rows[iRow]["PROC_NAME"].ToString();
                        else
                            ProcCodeTmp = string.Empty;
                    }

                    if (!PlantCodeTmp.Equals(dtGrid.Rows[iRow]["FA_WC_CD"].ToString()))
                    {
                        BodyPlantName = string.Format("<td class='tg-0lax' rowspan='{0}'>{1}</td>", dtGrid.Rows[iRow]["ROWSPAN2"], dtGrid.Rows[iRow]["FA_WC_CD"]);
                    }
                    if (!dtGrid.Rows[iRow]["ROWSPAN2"].ToString().Equals("1"))
                        PlantCodeTmp = dtGrid.Rows[iRow]["FA_WC_CD"].ToString();
                    else
                        PlantCodeTmp = string.Empty;

                    if (dtGrid.Rows[iRow]["PROC_NAME"].ToString().ToUpper().Equals("TOTAL"))
                    {
                        sBody += string.Format(@"<tr>" + BodyProcName + @"
				                        " + sBody2 + "<td class='tg-2lax' style='font-weight:bold;background-color:#ffffc9'>{11}</td><td class='tg-1lax'>{12}</td></tr>", argBodys);
                    }
                    else
                    {
                        sBody += string.Format(@"<tr>" + BodyProcName + @"
				                        <td class='tg-0lax'>{1}</td>
                                        <td class='tg-0lax'>{2}</td>
                                        <td class='tg-eslapse'>{3}</td>
                                        <td class='tg-0lax'>{4}</td>" +
                                            BodyPlantName + @"
                                        <td class='tg-0lax'>{6}</td>
                                        <td class='tg-0lax'>{7}</td>
                                        <td class='tg-1lax'>{8}</td>
				                        <td class='tg-0lax'>{9}</td>
                                        <td class='tg-0lax'>{10}</td>
				                        " + sBody2 + " <td class='tg-2lax' style='font-weight:bold;background-color:#ffffc9'>{11}</td><td class='tg-1lax'>{12}</td></tr>", argBodys);
                    }
                }

                EndTag = @"</tbody></table></body></html>";
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