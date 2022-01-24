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
    public partial class Mold_Repair_Monthly2 : Form
    {
        public Mold_Repair_Monthly2()
        {
            InitializeComponent();

            chart1.Size = new Size(750, 700);

        }

        Main frmMain = new Main();
        public DataTable _dt1, _dt2,_dt3;
        private void Mold_Repair_Monthly_Load(object sender, EventArgs e)
        {
            LoadDataMold(_dt1, _dt2, _dt3);
          //  CaptureControl(pnMold, "MoldChart");
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
                SetChart1



                return true;
            }
            catch (Exception ex)
            {
                frmMain.WriteLog($"  LoadDataMold: {ex.Message}");
                return false;
            }

        }

        private void SetChart1(DataTable argDt)
        {
            try
            {
                DataTable dt = argDt.Select($"CHART = {1}", "RN").CopyToDataTable();

                chart1.DataSource = dt;
                chart1.Series[0].ArgumentDataMember = "TXT";
                chart1.Series[0].ValueDataMembers.AddRange(new string[] { "VAL" });
            }
            catch (Exception ex)
            {
                frmMain.WriteLog($"  RepairMonthWh: {ex.Message}");
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




        private void BindingChart(DataTable dt)
        {
            try
            {
                chart1.DataSource = dt;
                chart1.Series[0].ArgumentDataMember = "WORK_PLACE_NM";
                chart1.Series[0].ValueDataMembers.AddRange(new string[] { "PER_MOLD_RP" });
            }
            catch
            {

            }
        }

       
        
    }
}
