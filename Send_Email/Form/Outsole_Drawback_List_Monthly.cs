using DevExpress.Utils;
using DevExpress.XtraGrid.Views.BandedGrid;
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
    public partial class Outsole_Drawback_List_Monthly : Form
    {
        public Outsole_Drawback_List_Monthly()
        {
            InitializeComponent();
            tblOutsole.Size = new Size(4000, 2000);
        }
        public bool _chkTest = false;
        public string _subject = "";
        private string _subjectSend = "";
        private readonly string[] _emailTest = { "dien.it@changshininc.com", "ngoc.it@changshininc.com", "MAN.SPT@changshininc.com" };
        Main frmMain = new Main();
        public DataTable _dtChart1, _dtChart2, _dtChart3, _dtChart4, _dtChart5, _dtChart6, _dtEmail;

        private void Outsole_Drawback_List_Monthly_Load(object sender, EventArgs e)
        {
            try
            {
                LoadDataOusoleMonthly(_dtChart1, _dtChart2, _dtChart3, _dtChart4, _dtChart5, _dtChart6);
                CaptureControl(tblOutsole, "OSMonthly");
                CreateMail(_subject, "", _dtEmail);
            }
            catch (Exception)
            {

                throw;
            }
        }

        private bool LoadDataOusoleMonthly(DataTable argChart1Dt,
                                           DataTable argChart2Dt,
                                           DataTable argChart3Dt,
                                           DataTable argChart4Dt,
                                           DataTable argChart5Dt,
                                           DataTable argChart6Dt)
        {
            try
            {
                SetChart("CHART1", argChart1Dt);
                SetChart("CHART2", argChart2Dt);
                SetChart("CHART3", argChart3Dt);
                SetChart("CHART4", argChart4Dt);
                SetChart("CHART5", argChart5Dt);
                SetChart("CHART6", argChart6Dt);
                //SetChart2(argDt, argWh);
                //SetChart3(argDt, argWh);
                //SetChart4(argDt, argWh);
                //SetChart5(argDt, argWh);
                //SetChart6(argDt, argWh);
                //DataTable dt = argDt2.Select($"WORK_PLACE ={((int)argWh)}").CopyToDataTable();
                //grdMain.DataSource = dt;
                //FormatGrid(grdView, "TOTAL MOLD REPAIR 2022");

                //grdMain2.DataSource = argDt3.Select($"WORK_PLACE ={((int)argWh)}").CopyToDataTable();
                //FormatGrid2(grdView2, "MOLD REPAIR BY LOCATION 2022");

                //_subjectSend = _subject.Replace("{WH}", dt.Rows[0]["WORK_PLACE_NM"].ToString());
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //frmMain.WriteLog($"  LoadDataMold: {ex.Message}");
                return false;
            }

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
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\OSMonthly.png", Outlook.OlAttachmentType.olByValue, null, "tr");
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
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }
                oRecips = null;
                mailItem.BCC = "ngoc.it@changshininc.com";
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

        #region Set Data Chart
        private void SetChart(string argType, DataTable argDt)
        {
            try
            {
                switch (argType)
                {
                    case "CHART1":
                        chartControl1.DataSource = argDt;
                        chartControl1.Series[0].ArgumentDataMember = "RESOURCE_CD";
                        chartControl1.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });
                        chartControl1.Series[1].ArgumentDataMember = "RESOURCE_CD";
                        chartControl1.Series[1].ValueDataMembers.AddRange(new string[] { "DRAWBACK_PROD" });
                        break;
                    case "CHART2":
                        chartControl2.DataSource = argDt;
                        chartControl2.Series[0].ArgumentDataMember = "LINE";
                        chartControl2.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });

                        break;
                    case "CHART3":
                        chartControl3.DataSource = argDt;
                        chartControl3.Series[0].ArgumentDataMember = "REASON_NAME";
                        chartControl3.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });

                        break;
                    case "CHART4":
                        chartControl4.DataSource = argDt;
                        chartControl4.Series[0].ArgumentDataMember = "RESOURCE_CD";
                        chartControl4.Series[0].ValueDataMembers.AddRange(new string[] { "HOURS" });

                        break;
                    case "CHART5":
                        chartControl5.DataSource = argDt;
                        chartControl5.Series[0].ArgumentDataMember = "SHIFT";
                        chartControl5.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });

                        break;
                    case "CHART6":
                        chartControl6.DataSource = argDt;
                        chartControl6.Series[0].ArgumentDataMember = "YMD";
                        chartControl6.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });

                      //  DataTable dt = Pivot(argDt, argDt.Columns["YMD"], argDt.Columns["CNT"]);
                        //createGrid(dt);
                        break;
                    default:
                        break;
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //frmMain.WriteLog($"  RepairMonthWh[Chart1]: {ex.Message}");
            }

        }

        private void createGrid(DataTable argDt)
        {
            try
            {
                //grdMain.DataSource = argDt;
                //for (int i = 0; i < grdView.Columns.Count; i++)
                //{
                //    grdView.Columns[i].OptionsColumn.AllowEdit = false;
                //    grdView.Columns[i].OptionsColumn.ReadOnly = true;
                //    grdView.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                //    grdView.Columns[i].AppearanceCell.TextOptions.VAlignment =  VertAlignment.Center;
                //    grdView.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 20, FontStyle.Bold);
                //    grdView.Columns[i].AppearanceCell.Font = new Font("Calibri", 20, FontStyle.Regular);
                //    if (i > 0)
                //    {
                //        grdView.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                //        grdView.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                //        grdView.Columns[i].DisplayFormat.FormatString = "{0:n0}";
                //    }
                //}
            }
            catch
            {

            }
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
        #endregion
    }
}
