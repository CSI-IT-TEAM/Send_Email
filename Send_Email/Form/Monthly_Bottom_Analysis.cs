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
    public partial class Monthly_Bottom_Analysis : Form
    {
        public Monthly_Bottom_Analysis()
        {
            InitializeComponent();
            pnMain.Size = new Size(4000, 2000);
        }
        private readonly string[] _emailTest = {  "MAN.SPT@changshininc.com" };
        Main frmMain = new Main();
        public DataTable _dtChart;
        public string _subject = "";
        private string _subjectSend = "";

        private void Monthly_Bottom_Analysis_Load(object sender, EventArgs e)
        {
            try
            {
                if (
                BindingDataForChart(_dtChart))
                {
                    CaptureControl(pnMain,"BT_INV_ANALYSIS");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private bool BindingDataForChart(DataTable dt)
        {
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

    }
}
