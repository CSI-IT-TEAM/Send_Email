using DevExpress.Utils;
using DevExpress.XtraCharts;
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

namespace Send_Email
{
    public partial class Mold_Repair_Monthly_WH : Form
    {
        public Mold_Repair_Monthly_WH()
        {
            InitializeComponent();

            tblMold.Size = new Size(3000, 1000);

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

        Main frmMain = new Main();
        public DataTable _dt1, _dt2,_dt3;
        private void Mold_Repair_Monthly_Load(object sender, EventArgs e)
        {
            LoadDataMold(_dt1, _dt2, _dt3);
            CaptureControl(tblMold, "MoldChartWh");
         //   CaptureControl(grdMain, "MoldGrid");
          //  CaptureControl(pnMold2, "MoldGrid2");

        }

        private void CaptureControl(Control control, string nameImg)
        {
            //  MemoryStream ms = new MemoryStream();
            string Path = Application.StartupPath + @"\Capture\";
            Bitmap bmp = new Bitmap(control.Width, control.Height);
            if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);
            control.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, control.Width, control.Height));
            bmp.Save(Path + nameImg + @".png", System.Drawing.Imaging.ImageFormat.Png); 
        }

        private bool LoadDataMold(DataTable argDt, DataTable argDt2, DataTable argDt3)
        {
            try
            {
                SetChart1(argDt, WH.IP);
                SetChart2(argDt, WH.IP);
                SetChart3(argDt, WH.OUTSOLE);
                SetChart4(argDt, WH.IP);
                SetChart5(argDt, WH.IP);
                SetChart6(argDt, WH.IP);

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                //frmMain.WriteLog($"  LoadDataMold: {ex.Message}");
                return false;
            }

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
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {2}", "RN").CopyToDataTable();

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
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {4}", "RN").CopyToDataTable();

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
                DataTable dt = argDt.Select($"WORK_PLACE ={((int)argWh)} and CHART = {5}", "RN").CopyToDataTable();

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


       
        
    }
}
