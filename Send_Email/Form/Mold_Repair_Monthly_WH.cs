using DevExpress.Utils;
using DevExpress.XtraCharts;
using DevExpress.XtraGrid.Views.BandedGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Send_Email
{
    public partial class Mold_Repair_Monthly_WH : Form
    {
        public Mold_Repair_Monthly_WH()
        {
            InitializeComponent();

            tblMold.Size = new Size(4000, 2000);

        }
        enum WH
        {
            IP = 30,
            PU = 50,
            DMP = 90,
            OUTSOLE = 20,
            PHYLON = 70,
            CMP = 40,
        }

        public bool _chkTest = false;
        public DataTable _dtEmail = null;
        public string _subject = "";
        private string _subjectSend = "";

        //"jungbo.shim@dskorea.com", "nguyen.it@changshininc.com", "dien.it@changshininc.com"
        private readonly string[] _emailTest = { "nguyen.it@changshininc.com", "dien.it@changshininc.com" }; 

        Main frmMain = new Main();
        public DataTable _dt1, _dt2,_dt3;
        private void Mold_Repair_Monthly_Load(object sender, EventArgs e)
        {
            //foreach (WH warehouse in Enum.GetValues(typeof(WH)))
            //{
            //    LoadDataMold(_dt1, _dt2, _dt3, warehouse);
            //    CaptureControl(tblMold, "MoldChartWh");

            //    CreateMailMoldMonthWh(_subjectSend, "", _dtEmail);
            //}

            LoadDataMold(_dt1, _dt2, _dt3, WH.OUTSOLE);
            CaptureControl(tblMold, "MoldChartWh");

            CreateMailMoldMonthWh(_subjectSend, "", _dtEmail);
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
                Debug.WriteLine(ex.Message);
                // WriteLog("CreateMailMoldMonthWh: " + ex.ToString());
            }
        }



        private void CaptureControl(Control control, string nameImg)
        {
            try
            {
                //  MemoryStream ms = new MemoryStream();
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

        private bool LoadDataMold(DataTable argDt, DataTable argDt2, DataTable argDt3, WH argWh)
        {
            try
            {
                SetChart1(argDt, argWh);
                SetChart2(argDt, argWh);
                SetChart3(argDt, argWh);
                SetChart4(argDt, argWh);
                SetChart5(argDt, argWh);
                SetChart6(argDt, argWh);
                DataTable dt = argDt2.Select($"WORK_PLACE ={((int)argWh)}").CopyToDataTable();
                grdMain.DataSource = dt;
                FormatGrid(grdView, "TOTAL MOLD REPAIR 2022");

                grdMain2.DataSource = argDt3.Select($"WORK_PLACE ={((int)argWh)}").CopyToDataTable();
                FormatGrid2(grdView2, "MOLD REPAIR BY LOCATION 2022");

                _subjectSend = _subject.Replace("{WH}", dt.Rows[0]["WORK_PLACE_NM"].ToString());
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //frmMain.WriteLog($"  LoadDataMold: {ex.Message}");
                return false;
            }

        }

        private void FormatGrid(BandedGridView grid, string argText)
        {
            try
            {
                // grdMain.Font = new Font("Calibri", 15, FontStyle.Bold);
                //grid.OptionsView.AllowCellMerge = true;
                //grid.BandPanelRowHeight = 30;
                int width = grdMain.Width / 12;
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    if (grid.Columns[i].OwnerBand.ParentBand != null)
                    {
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.Font = new Font("Calibri", 18, FontStyle.Bold);
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.BackColor = Color.FromArgb(30, 84, 111);
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.ForeColor = Color.White;
                        grid.Columns[i].OwnerBand.ParentBand.Caption = argText;

                    }
                    grid.Columns[i].OwnerBand.AppearanceHeader.BackColor = Color.FromArgb(30, 84, 111);
                    grid.Columns[i].OwnerBand.AppearanceHeader.ForeColor = Color.White;
                    grid.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 18, FontStyle.Bold);
                    grid.Columns[i].OwnerBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                    grid.Columns[i].OwnerBand.AppearanceHeader.TextOptions.VAlignment = VertAlignment.Center;

                    grid.Columns[i].Width = width;
                    grid.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                    grid.Columns[i].AppearanceCell.TextOptions.VAlignment = VertAlignment.Center;


                    grid.Columns[i].AppearanceCell.Font = new Font("Calibri", 20, FontStyle.Regular);

                    grid.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.False;
                    grid.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                    grid.Columns[i].DisplayFormat.FormatString = "#,0.##";
                }
            }
            catch
            {

            }

        }

        private void FormatGrid2(BandedGridView grid, string argText)
        {
            try
            {
                // grdMain.Font = new Font("Calibri", 15, FontStyle.Bold);
                //grid.OptionsView.AllowCellMerge = true;
                //grid.BandPanelRowHeight = 30;
                int width = (grdMain2.Width -150) / 12;
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    if (grid.Columns[i].OwnerBand.ParentBand != null)
                    {
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.Font = new Font("Calibri", 18, FontStyle.Bold);
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.BackColor = Color.FromArgb(30, 84, 111);
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.ForeColor = Color.White;
                        grid.Columns[i].OwnerBand.ParentBand.Caption = argText;

                    }
                    grid.Columns[i].OwnerBand.AppearanceHeader.BackColor = Color.FromArgb(30, 84, 111);
                    grid.Columns[i].OwnerBand.AppearanceHeader.ForeColor = Color.White;
                    grid.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 18, FontStyle.Bold);
                    grid.Columns[i].OwnerBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                    grid.Columns[i].OwnerBand.AppearanceHeader.TextOptions.VAlignment = VertAlignment.Center;

                    if (i == 0)
                    {
                        grid.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Near;
                        grid.Columns[i].Width = 150;
                    }                       
                    else
                    {
                        grid.Columns[i].AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                        grid.Columns[i].Width = width;
                    }

                    grid.Columns[i].AppearanceCell.TextOptions.VAlignment = VertAlignment.Center;

                    grid.Columns[i].AppearanceCell.Font = new Font("Calibri", 20, FontStyle.Regular);

                    grid.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.False;
                    grid.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                    grid.Columns[i].DisplayFormat.FormatString = "#,0.##";
                }
            }
            catch
            {

            }

        }


        private void grdView_CustomDrawBandHeader(object sender, BandHeaderCustomDrawEventArgs e)
        {
            if (e.Band == null) return;
            if (e.Band.AppearanceHeader.BackColor != Color.Empty)
                e.Info.AllowColoring = true;


        }

        private void SetChart1(DataTable argDt, WH argWh)
        {
            try
            {
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {1}", "RN").CopyToDataTable();

                chart1.DataSource = dt;
                chart1.Series[0].ArgumentDataMember = "TXT";
                chart1.Series[0].ValueDataMembers.AddRange(new string[] { "VAL" });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //frmMain.WriteLog($"  RepairMonthWh[Chart1]: {ex.Message}");
            }
           
        }
        private void SetChart2(DataTable argDt, WH argWh)
        {
            try
            {
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {2}", "RN DESC").CopyToDataTable();

                chart2.DataSource = dt;
                chart2.Series[0].ArgumentDataMember = "TXT";
                chart2.Series[0].ValueDataMembers.AddRange(new string[] { "VAL" });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //frmMain.WriteLog($"  RepairMonthWh[Chart2]: {ex.Message}");
            }

        }
        private void SetChart3(DataTable argDt, WH argWh)
        {
            try
            {
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {3}", "RN").CopyToDataTable();

                chart3.DataSource = dt;
                chart3.Series[0].ArgumentDataMember = "TXT";
                chart3.Series[0].ValueDataMembers.AddRange(new string[] { "VAL" });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //frmMain.WriteLog($"  RepairMonthWh[Chart3]: {ex.Message}");
            }

        }
        private void SetChart4(DataTable argDt, WH argWh)
        {
            try
            {
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {4}", "RN DESC").CopyToDataTable();

                chart4.DataSource = dt;
                chart4.Series[0].ArgumentDataMember = "TXT";
                chart4.Series[0].ValueDataMembers.AddRange(new string[] { "VAL" });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                // frmMain.WriteLog($"  RepairMonthWh[Chart4]: {ex.Message}");
            }

        }
        private void SetChart5(DataTable argDt, WH argWh)
        {
            try
            {
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {5}", "RN DESC").CopyToDataTable();

                chart5.DataSource = dt;
                chart5.Series[0].ArgumentDataMember = "TXT";
                chart5.Series[0].ValueDataMembers.AddRange(new string[] { "VAL" });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                // frmMain.WriteLog($"  RepairMonthWh[Chart5]: {ex.Message}");
            }

        }
        private void SetChart6(DataTable argDt, WH argWh)
        {
            try
            {
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {6}", "RN").CopyToDataTable();

                chart6.DataSource = dt;
                chart6.Series[0].ArgumentDataMember = "TXT";
                chart6.Series[0].ValueDataMembers.AddRange(new string[] { "VAL" });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                // frmMain.WriteLog($"  RepairMonthWh[Chart6]: {ex.Message}");
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




    }
}
