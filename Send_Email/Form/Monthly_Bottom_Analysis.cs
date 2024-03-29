﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace Send_Email
{
    public partial class Monthly_Bottom_Analysis : Form
    {
        public Monthly_Bottom_Analysis()
        {
            InitializeComponent();
            pnMain.Size = new Size(4000, 2000);
        }
        private readonly string[] _emailTest = {  "MAN.SPT@changshininc.com" };
        Main frmMain = new Main();
        public DataTable _dtChart, _dtEmail;
        public bool _chkTest = false;
        public string _subjectSend = "";

        private void Monthly_Bottom_Analysis_Load(object sender, EventArgs e)
        {
            try
            {
                if (
                BindingDataForChart(_dtChart))
                {
                    CaptureControl(pnMain,"BT_INV_ANALYSIS");
                  //  CreateMail(_subjectSend, "", _dtEmail);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private bool BindingDataForChart(DataTable dt)
        {
            //Bottom & Stockfit Set
            try
            {
                chartBottomSet.DataSource = dt;
                chartBottomSet.Series[0].ArgumentDataMember = "FA_WC_NM";
                chartBottomSet.Series[0].ValueDataMembers.AddRange(new string[] { "BT_HOURS" });
                chartBottomSet.Series[1].ArgumentDataMember = "FA_WC_NM";
                chartBottomSet.Series[1].ValueDataMembers.AddRange(new string[] { "STK_HOURS" });
            }
            catch (Exception ex)
            {
                return false;
            }

            //Finished Sole Set
            try
            {
                chartFS.DataSource = dt;
                chartFS.Series[0].ArgumentDataMember = "FA_WC_NM";
                chartFS.Series[0].ValueDataMembers.AddRange(new string[] { "FS_HOURS" });
            }
            catch (Exception ex)
            {
                return false;
            }


            //Top 5 bottom inventory sets
            try
            {
                DataTable dtBTChart = dt.Select("BT_HOURS_SEQ <=5", "BT_HOURS_SEQ").CopyToDataTable();
                chartTop5BT.DataSource = dtBTChart;
                chartTop5BT.Series[0].ArgumentDataMember = "FA_WC_NM";
                chartTop5BT.Series[0].ValueDataMembers.AddRange(new string[] { "BT_HOURS" });
            }
            catch (Exception ex)
            {

                return false;
            }

            //Top 5 stockfit inventory sets
            try
            {
                DataTable dtSTKChart = dt.Select("STK_HOURS_SEQ <=5", "STK_HOURS_SEQ").CopyToDataTable();
                chartTop5STK.DataSource = dtSTKChart;
                chartTop5STK.Series[0].ArgumentDataMember = "FA_WC_NM";
                chartTop5STK.Series[0].ValueDataMembers.AddRange(new string[] { "STK_HOURS" });
            }
            catch (Exception ex)
            {
                return false;
            }

            //Top 5 finised sole-upper sets
            try
            {
                DataTable dtFSUPChart = dt.Select("FS_UP_HOURS_SEQ <=5", "FS_UP_HOURS_SEQ").CopyToDataTable();
                chartTop5FSUP.DataSource = dtFSUPChart;
                chartTop5FSUP.Series[0].ArgumentDataMember = "FA_WC_NM";
                chartTop5FSUP.Series[0].ValueDataMembers.AddRange(new string[] { "FS_UP_HOURS" });
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;

        }

        private void CaptureControl(Control control, string nameImg)
        {
            try
            {
                string Path = Application.StartupPath + @"\Capture\";
                Bitmap bmp = new Bitmap(control.Width, control.Height);
                if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);
                control.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, control.Width, control.Height));
                bmp.Save(Path + nameImg + @".png", System.Drawing.Imaging.ImageFormat.Png);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        private void CreateMail(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\BT_INV_ANALYSIS.png", Outlook.OlAttachmentType.olByValue, null, "tr");
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

                if (_chkTest)
                {
                    for (int i = 0; i < _emailTest.Length; i++)
                    {
                        Microsoft.Office.Interop.Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }
                oRecips = null;
                mailItem.BCC = "phuoc.it@changshininc.com";
                string imgInfo = "imgInfo";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                mailItem.HTMLBody = String.Format(@"<img src='cid:{0}'>", imgInfo) + htmlBody;

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                // WriteLog("CreateMailMoldMonthWh: " + ex.ToString());
            }
        }

    }
}
