using DevExpress.Utils;
using System;
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
    public partial class Releif_AVSM_Report_Daily : Form
    {
        public Releif_AVSM_Report_Daily()
        {
            InitializeComponent();
        }
        public bool _chkTest = false;
        public string _subject = "";
        private string _subjectSend = "";
        private readonly string[] _emailTest = { "dien.it@changshininc.com" };
        Main frmMain = new Main();
        public DataTable _dtData, _dtEmail;

        private void Releif_AVSM_Report_Daily_Load(object sender, EventArgs e)
        {
            try
            {
                SetData(_dtData);
                CaptureControl(panel1, "ReliefReg");
                CreateMail(_subject, "", _dtEmail);
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void SetData(DataTable dtData)
        {
            try
            {
                grdMain2.DataSource = dtData;
                formatGrid();
            }
            catch
            {
                
            }
        }

        private void formatGrid()
        {
            try
            {
                for (int i = 0; i < grdView2.Columns.Count; i++)
                {
                    grdView2.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                    grdView2.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    grdView2.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    if (i >= 7)
                    {
                        grdView2.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                        grdView2.Columns[i].DisplayFormat.FormatString = "{0:n0}";
                    }
                    grdView2.Columns["AREA_NAME"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    grdView2.Columns["PROCESS_NAME"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    grdView2.Columns["PLANT_NAME"].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.True;
                    grdView2.Columns["LINE_CD"].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.True;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void CaptureControl(Control control, string nameImg)
        {
            try
            {
                panel1.Size = new Size(1100,  (grdView2.RowCount * grdView2.RowHeight) + 105);
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

        private void grdView2_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            try
            {
                string cellValue = grdView2.GetRowCellValue(e.RowHandle, grdView2.Columns["PLANT_NAME"]).ToString();
                if (cellValue.ToUpper().Equals("TOTAL"))
                {
                    e.Appearance.BackColor = Color.FromArgb(250, 249, 200);
                }
            }
            catch
            {

                throw;
            }
        }

        private void CreateMail(string Subject, string htmlBody, DataTable dtEmail)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ReliefReg.png", Outlook.OlAttachmentType.olByValue, null, "tr");
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
