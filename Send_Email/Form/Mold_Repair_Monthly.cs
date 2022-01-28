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


namespace Send_Email
{
    public partial class Mold_Repair_Monthly : Form
    {
        public Mold_Repair_Monthly()
        {
            InitializeComponent();

            pnMold.Size = new Size(2000, 1000);
            chartMold.Size = new Size(1000, 1000);
            tblGrid.Size = new Size(2200, 260);
           // pnMold2.Size = new Size(500, 300);
        }

        Main frmMain = new Main();
        public DataTable _dt1, _dt2,_dt3;
        private void Mold_Repair_Monthly_Load(object sender, EventArgs e)
        {
            LoadDataMold(_dt1, _dt2, _dt3);
            CaptureControl(pnMold, "MoldChart");
            CaptureControl(tblGrid, "MoldGrid");
           // CaptureControl(pnMold2, "MoldGrid2");

        }

        

        
        DataTable _dtPivot = null;
        private bool LoadDataMold(DataTable argDt, DataTable argDt2, DataTable argDt3)
        {
            try
            {
                DataTable dt = argDt;
                DataTable dtYMD = dt.AsEnumerable().Where(r => r.Field<string>("IS_YMD") == "Y").OrderBy(r => r.Field<string>("WORK_YMD")).CopyToDataTable();
                DataView view = new DataView(dtYMD);
                DataTable distinctValues = view.ToTable(true, "WORK_YMD");
                InitBandHeader(distinctValues);
                dt.Columns.Remove(dt.Columns["IS_YMD"]);
                _dtPivot = Pivot(dt, dt.Columns["WORK_YMD"], dt.Columns["MOLD_RP_QTY"]);
                grdMain.DataSource = _dtPivot;
                grdMain2.DataSource = argDt3;

                //SetData(grdMain, dtPivot);
                FormatGrid(grdView);
                FormatGrid2(grdView2);
                BindingChart(_dtPivot);
                //SetDataChart(argDt2);

                setChartRound(chartPu, GetDataTemp(argDt2, "PU"));
                setChartRound(chartIp, GetDataTemp(argDt2, "IP"));
                setChartRound(chartDmp, GetDataTemp(argDt2, "DMP"));
                setChartRound(chartOutsole, GetDataTemp(argDt2, "Outsole"));
                setChartRound(chartPhylon, GetDataTemp(argDt2, "Phylon"));
                setChartRound(chartCmp, GetDataTemp(argDt2, "CMP"));
                // SetTreelist(argDt2);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return false;
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
                gridBandTotalMold.Width = 100;

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
                WORK_BOTTOM.Width = 50;
                WORK_BOTTOM.AppearanceHeader.BorderColor = Color.White;

                // 
                // WORK_PLACE_NM
                // 
                WORK_PLACE_NM.Caption = "WORK_PLACE_NM";
                WORK_PLACE_NM.FieldName = "WORK_PLACE_NM";
                WORK_PLACE_NM.Name = "WORK_PLACE_NM";
                WORK_PLACE_NM.Visible = true;
                WORK_PLACE_NM.Width = 50;
                WORK_PLACE_NM.AppearanceHeader.BorderColor = Color.White;

                // 
                // TOTAL_MOLD_RP
                // 
                TOTAL_MOLD_RP.Caption = "TOTAL_MOLD_RP";
                TOTAL_MOLD_RP.FieldName = "TOTAL_MOLD_RP";
                TOTAL_MOLD_RP.Name = "TOTAL_MOLD_RP";
                TOTAL_MOLD_RP.Visible = true;
                TOTAL_MOLD_RP.AppearanceHeader.BorderColor = Color.White;

                // 
                // AVG_MOLD
                // 
                AVG_MOLD.Caption = "AVG_MOLD";
                AVG_MOLD.FieldName = "AVG_MOLD_RP";
                AVG_MOLD.Name = "AVG_MOLD";
                AVG_MOLD.Visible = true;
                AVG_MOLD.Width = 100;
                AVG_MOLD.AppearanceHeader.BorderColor = Color.White;

                // 
                // PER_MOLD
                // 
                PER_MOLD.Caption = "PER_MOLD";
                PER_MOLD.FieldName = "PER_MOLD_RP";
                PER_MOLD.Name = "PER_MOLD";
                PER_MOLD.Visible = true;
                PER_MOLD.AppearanceHeader.BorderColor = Color.White;

                gridBandBottom.Columns.Add(WORK_BOTTOM);
                gridBandBottom.Columns.Add(WORK_PLACE_NM);
                gridBandTotalMold.Columns.Add(TOTAL_MOLD_RP);
                gridBandTotalMold.AppearanceHeader.BorderColor = Color.White;
                gridBandBottom.AppearanceHeader.BorderColor = Color.White;

                gridBandAvgMold.Columns.Add(AVG_MOLD);
                gridBandPerMold.Columns.Add(PER_MOLD);
                gridBandPerMold.AppearanceHeader.BorderColor = Color.White;

                grdView.Bands.AddRange(new GridBand[] { gridBandBottom, gridBandTotalMold, gridBandMonth, gridBandAvgMold, gridBandPerMold });
                grdView.Columns.AddRange(new BandedGridColumn[] {
                   WORK_BOTTOM,
                   WORK_PLACE_NM,
                   TOTAL_MOLD_RP,AVG_MOLD,PER_MOLD});
                // 
                // gridBandMonth
                // 
                DateTime date1;
                DateTime.TryParseExact(dt.Rows[0]["WORK_YMD"].ToString(),"yyyyMMdd"
                                      , System.Globalization.CultureInfo.InvariantCulture
                                      , System.Globalization.DateTimeStyles.None 
                                      , out date1);
                string date = date1.ToString("MMM-yyyy");
                gridBandMonth.AppearanceHeader.Options.UseTextOptions = true;
                gridBandMonth.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                gridBandMonth.AppearanceHeader.TextOptions.VAlignment = VertAlignment.Center;
                gridBandMonth.Caption = date;
                gridBandMonth.Name = "gridBandMonth";
                gridBandMonth.VisibleIndex = 2;
                gridBandMonth.AppearanceHeader.BorderColor = Color.White;

                
               _numDays = dt.Rows.Count;
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
                    ColumnsDays.Width = 40;
                    ColumnsDays.AppearanceHeader.BorderColor = Color.White;
                    gridbandDays.Columns.Add(ColumnsDays);
                    grdView.Columns.AddRange(new BandedGridColumn[] { ColumnsDays });
                }
                grdView.PaintStyleName = "Flat";

            }
            catch (Exception ex)
            {

                throw;
            }
        }
        int _numDays;
        private void FormatGrid(BandedGridView grid)
        {
            try
            {
                // grdMain.Font = new Font("Calibri", 15, FontStyle.Bold);
                grdView.OptionsView.AllowCellMerge = true;
                // grdView.BandPanelRowHeight = 30;
                int gridWidth = 0;
                int gridColCount = grid.Columns.Count;
                int width = (grdMain.Width - (52 + (82 *3) + 112)) / _numDays;

                
                for (int i = 0; i < gridColCount; i++)
                {
                    if (grid.Columns[i].OwnerBand.ParentBand != null)
                    {
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.Font = new Font("Calibri", 15, FontStyle.Bold);
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.BackColor = Color.FromArgb(30, 84, 111);
                        grid.Columns[i].OwnerBand.ParentBand.AppearanceHeader.ForeColor = Color.White;

                    }
                    grid.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 15, FontStyle.Bold);
                    grid.Columns[i].AppearanceCell.Font = new Font("Calibri", 16, FontStyle.Regular);
                    grid.Columns[i].OwnerBand.AppearanceHeader.BackColor = Color.FromArgb(30, 84, 111);
                    grid.Columns[i].OwnerBand.AppearanceHeader.ForeColor = Color.White;

                    switch (grid.Columns[i].Name)
                    {
                        case "WORK_BOTTOM":
                            grid.Columns[i].Width = 52;                           
                            break;

                        case "PER_MOLD":
                        case "TOTAL_MOLD_RP":
                        case "WORK_PLACE_NM":
                            grid.Columns[i].Width = 82;
                            break;
                        case "AVG_MOLD":
                            grid.Columns[i].Width = 112;
                            break;
                        default:
                            grid.Columns[i].Width = width;
                            break;
                    }
                    gridWidth += grid.Columns[i].Width;

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


        private void FormatGrid2(BandedGridView grid)
        {
            try
            {
                // grdMain.Font = new Font("Calibri", 15, FontStyle.Bold);
                grdView.OptionsView.AllowCellMerge = true;
                grdView.BandPanelRowHeight = 30;

                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    grid.Columns[i].OwnerBand.AppearanceHeader.Font = new Font("Calibri", 14, FontStyle.Bold);
                    grid.Columns[i].OwnerBand.AppearanceHeader.BackColor = Color.FromArgb(30, 84, 111);
                    grid.Columns[i].OwnerBand.AppearanceHeader.ForeColor = Color.White;

                    if (i==1) 
                        grid.Columns[i].Width = 80;
                    else
                        grid.Columns[i].Width = 93;
                    grid.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.False;
                    if (i == 0)
                        grid.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    else
                        grid.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                    grid.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                    grid.Columns[i].DisplayFormat.FormatString = "#,0.##";

                    grid.Columns[i].AppearanceCell.Font = new Font("Calibri", 16);
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

        private void chartMold_CustomDrawSeriesPoint(object sender, CustomDrawSeriesPointEventArgs e)
        {
            if (e.SeriesPoint.Values.Count() < 1) return;
            if (e.SeriesPoint.Values[0] > 0)
            {
                for (int i = 0; i < _dtPivot.Rows.Count; i++)
                {
                    if (e.SeriesPoint.Argument.ToString().Equals(_dtPivot.Rows[i]["WORK_PLACE_NM"].ToString()))
                    {
                        e.LabelText = $"Ratio: { _dtPivot.Rows[i]["PER_MOLD_RP"].ToString().TrimEnd('0')}%\n" +
                                      $"Avg Mold: {_dtPivot.Rows[i]["AVG_MOLD_RP"].ToString().TrimEnd('0')}";
                    }
                }
            }
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

        

    }
}
