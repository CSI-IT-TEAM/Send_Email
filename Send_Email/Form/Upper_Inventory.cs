using DevExpress.Utils;
using DevExpress.XtraCharts;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Grid;
using JPlatform.Client.Controls6;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Send_Email
{
    public partial class frmUpper_Inventory : Form
    {
        public frmUpper_Inventory()
        {
            InitializeComponent();
            //tblChart.Size = new Size(3000, 1000);
            //tblAssy.Size = new Size(3000, 500);
            //tblGridInv.Size = new Size(3000, 300);

            tblChart.Size = new Size(2500, 600);
            tblAssy.Size = new Size(2600, 130);
            //tblGridInv.Size = new Size(3000, 2800);
            tblGridInv.Size = new Size(2500, 1500);
            tblGridInv2.Size = new Size(2500, 1500);
        }
        public bool _chkTest = false;
        public string _subject = "";
        private string _subjectSend = "";
        public string _type = "";
        private readonly string[] _emailTest = { "dien.it@changshininc.com", "ngoc.it@changshininc.com", "MAN.SPT@changshininc.com" };
        Main frmMain = new Main();
        public DataTable _dt1, _dt2, _dt3, _dt4;

        private void pnOSMain_Paint(object sender, PaintEventArgs e)
        {

        }

        private void grdViewDOS_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {            

            try
            {
                Rectangle rect = e.Bounds;
                rect.Inflate(new Size(1, 1));

                Brush brush = new SolidBrush(e.Appearance.BackColor);
                e.Graphics.FillRectangle(brush, rect);
                Pen pen_vertical = new Pen(Color.White, 3F);

                //raw bottom
                e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y + rect.Height, rect.X + rect.Width, rect.Y + rect.Height);
                
                // draw right
                e.Graphics.DrawLine(pen_vertical, rect.X + rect.Width - 1, rect.Y, rect.X + rect.Width - 1, rect.Y + rect.Height);
                if (e.Column.ColumnHandle == 0)
                {
                    // draw left
                    e.Graphics.DrawLine(pen_vertical, rect.X + 1, rect.Y, rect.X + 1, rect.Y + rect.Height);
                }

                e.Graphics.DrawString(e.DisplayText, new System.Drawing.Font("Calibri", 12, FontStyle.Regular), new SolidBrush(Color.White), rect, e.Appearance.GetStringFormat());

                e.Handled = true;
            }
            catch
            {

            }
        }

        private void grdViewOS_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                GridViewEx gvw = sender as GridViewEx;
                Rectangle rect = e.Bounds;
                rect.Inflate(new Size(1, 1));

                Brush brush = new SolidBrush(e.Appearance.BackColor);
                e.Graphics.FillRectangle(brush, rect);
                Pen pen_vertical = new Pen(Color.White, 3F);


                if (e.RowHandle == 0)
                {
                    // raw top
                    e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y  + 1, rect.X + rect.Width, rect.Y + 1);
                }                

                //raw bottom
                e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y + rect.Height - 1, rect.X + rect.Width, rect.Y + rect.Height - 1);
                // draw left
                e.Graphics.DrawLine(pen_vertical, rect.X + 1, rect.Y, rect.X + 1, rect.Y + rect.Height);
                if (e.Column.ColumnHandle == gvw.Columns.Count - 1)
                {

                    // draw right
                    e.Graphics.DrawLine(pen_vertical, rect.X + rect.Width - 1, rect.Y, rect.X + rect.Width - 1, rect.Y + rect.Height);
                }

                e.Graphics.DrawString(e.DisplayText, new System.Drawing.Font("Calibri", 12, FontStyle.Regular), new SolidBrush(Color.White), rect, e.Appearance.GetStringFormat());

                e.Handled = true;
            }
            catch
            {

            }
        }

        private void grdViewDOS_CustomDrawColumnHeader(object sender, DevExpress.XtraGrid.Views.Grid.ColumnHeaderCustomDrawEventArgs e)
        {
            try
            {
                if (e.Column == null) return;
                Rectangle rect = e.Bounds;
                rect.Inflate(new Size(1, 1));

                Brush brush = new SolidBrush(e.Appearance.BackColor);
                e.Graphics.FillRectangle(brush, rect);
                Pen pen_vertical = new Pen(Color.White, 3F);
                Pen line = new Pen(Color.White, 2F);
                string[] ls = e.Column.Caption.Split('\n');
                // draw right
                e.Graphics.DrawLine(pen_vertical, rect.X + rect.Width - 3, rect.Y, rect.X + rect.Width - 3, rect.Y + rect.Height);


                //draw top
                e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y + 2, rect.X + rect.Width - 2, rect.Y + 2);

                //raw bottom
                e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y + rect.Height - 2, rect.X + rect.Width, rect.Y + rect.Height - 2);

                if (e.Column.ColumnHandle == 0)
                {
                    // draw left
                    e.Graphics.DrawLine(pen_vertical, rect.X + 2, rect.Y, rect.X + 2, rect.Y + rect.Height);
                }
                else
                {
                    // draw left
                    e.Graphics.DrawLine(line, rect.X, rect.Y, rect.X, rect.Y + rect.Height);
                }

                e.Graphics.DrawString(CultureInfo.CurrentCulture.TextInfo.ToTitleCase(e.Column.GetCaption().ToLower()), e.Appearance.GetFont(), new SolidBrush(e.Appearance.GetForeColor()), rect, e.Appearance.GetStringFormat());
                e.Handled = true;
            }
            catch { }
        }

        private DataTable SEL_DATA_WEEKLY_CONSTRAINT(string V_P_TYPE, string V_P_DIV, string V_P_DATE, string V_P_ITEM)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.SEPHIROTH;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_WEEKLY_CONSTRAINT";
                MyOraDB.ReDim_Parameter(5);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DIV";
                MyOraDB.Parameter_Name[2] = "V_P_DATE";
                MyOraDB.Parameter_Name[3] = "V_P_ITEM";
                MyOraDB.Parameter_Name[4] = "CV_1";


                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;


                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DIV;
                MyOraDB.Parameter_Values[2] = V_P_DATE;
                MyOraDB.Parameter_Values[3] = V_P_ITEM;
                MyOraDB.Parameter_Values[4] = "";

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

        private void frmUpper_Inventory_Load(object sender, EventArgs e)
        {
            try
            {
                //RunWeeklyBC("Q", "20220314");
                LoadUpperInv(_dt1, _dt2, _dt3);
                CaptureControl(tblChart, "ChartUpperInv");
                CaptureControl(tblAssy, "Assembly");
                CaptureControl(tblGridInv, "GridUpperInv");
                CaptureControl(tblGridInv2, "GridUpperInv2");
                CreateMail(_subject, "", _dt4, "ChartUpperInv", "Assembly", "GridUpperInv", "GridUpperInv2");
            }
            catch (Exception)
            {

                throw;
            }
        }


        private bool LoadUpperInv(DataTable argdt1,
                                DataTable argdt2,
                                DataTable argdt3
            )
        {
            try
            {
                if (argdt1.Rows.Count > 0)
                {
                    this._subject = argdt1.Rows[0]["SUBJECT"].ToString();


                    lblVCInv.Text = "Vinh Cuu Inventory : " + argdt1.Rows[0]["TOT_VC"].ToString() + " Pairs";
                    lblLTInv.Text = "Long Thanh Inventory : " + argdt1.Rows[0]["TOT_LT"].ToString() + " Pairs";
                    lblTPInv.Text = "Tan Phu Inventory : " + argdt1.Rows[0]["TOT_TP"].ToString() + " Pairs";
                    chartINV.DataSource = argdt1;
                    chartINV.Series[0].ArgumentDataMember = "LINE";
                    chartINV.Series[0].ValueDataMembers.AddRange(new string[] { "INV_VC" });
                    chartINV.Series[1].ArgumentDataMember = "LINE";
                    chartINV.Series[1].ValueDataMembers.AddRange(new string[] { "INV_LT" });
                    chartINV.Series[2].ArgumentDataMember = "LINE";
                    chartINV.Series[2].ValueDataMembers.AddRange(new string[] { "INV_TP" });
                    chartINV.Series[3].ArgumentDataMember = "LINE";
                    chartINV.Series[3].ValueDataMembers.AddRange(new string[] { "TARGET" });

                    chartINV.Titles[0].Font = new Font("Times New Roman", 14, FontStyle.Bold ^ FontStyle.Italic);
                    chartINV.Legend.Font = new Font("Times New Roman", 12, FontStyle.Bold ^ FontStyle.Italic);
                    ((XYDiagram)chartINV.Diagram).AxisX.Label.Font = new Font("Calibri", 8, FontStyle.Bold);
                    ((XYDiagram)chartINV.Diagram).AxisY.Label.Font = new Font("Calibri", 8, FontStyle.Bold);
                    chartINV.Series[0].Label.Font = new Font("Tahoma", 7, FontStyle.Bold);
                    chartINV.Series[1].Label.Font = new Font("Tahoma", 7, FontStyle.Bold);
                    chartINV.Series[2].Label.Font = new Font("Tahoma", 7, FontStyle.Bold);
                    chartINV.Series[3].Label.Font = new Font("Tahoma", 7, FontStyle.Bold);
                }
                if (argdt2.Rows.Count > 0)
                {
                    DataTable dtpivot = Pivot(argdt2, argdt2.Columns["LINE"], argdt2.Columns["PROD"]);
                    grdAssy.DataSource = dtpivot;
                    gvwAssy.ColumnPanelRowHeight = 25;
                    gvwAssy.RowHeight = 25;
                    for (int i = 0; i < gvwAssy.Columns.Count; i++)
                    {
                        GridColumn col = gvwAssy.Columns[i];
                        col.AppearanceHeader.Font = new Font("Calibri", 9, FontStyle.Regular);
                        col.AppearanceCell.Font = new Font("Calibri", 9, FontStyle.Regular);
                        if (i == 0) col.Visible = false;
                        col.Width = 35;
                        col.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                        col.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                        col.DisplayFormat.FormatType = FormatType.Numeric;
                        col.DisplayFormat.FormatString = "#,#";

                    }
                }

                if (argdt3.Rows.Count > 0)
                {
                    DataTable dtpivot = Pivot(argdt3, argdt3.Columns["CS_SIZE"], argdt3.Columns["INV"]).Select("", "Factory, Plant, LINE, GRP").CopyToDataTable();
                    if (dtpivot.Select("Factory in ('Factory 1','Factory 2','Factory 3')").Count() > 0)
                    {
                        grdInv.DataSource = dtpivot.Select("Factory in ('Factory 1','Factory 2','Factory 3')", "Factory, Plant, LINE, GRP").CopyToDataTable();
                        gvwInv.ColumnPanelRowHeight = 25;
                        gvwInv.RowHeight = 22;
                        for (int i = 0; i < gvwInv.Columns.Count; i++)
                        {
                            GridColumn col = gvwInv.Columns[i];
                            col.OptionsColumn.AllowMerge = DefaultBoolean.False;
                            col.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                            col.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                            col.Width = 90;
                            if (i < 2)
                                col.Visible = false;
                            else if (i < 6)
                            {
                                col.Width = 90;
                                if (i <= 4)
                                    col.OptionsColumn.AllowMerge = DefaultBoolean.True;
                            }
                            else
                            {
                                col.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                                if(i==6)
                                {
                                    col.Width = 75;
                                }
                                else if(i==7)
                                {
                                    col.Width = 75;
                                }
                                else
                                {
                                    col.Width = 55;
                                }
                                col.DisplayFormat.FormatType = FormatType.Numeric;
                                col.DisplayFormat.FormatString = "#,#";
                            }

                            col.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Regular);
                            col.AppearanceCell.Font = new Font("Calibri", 12, FontStyle.Regular);



                        }
                    }

                    if (dtpivot.Select("Factory not in ('Factory 1','Factory 2','Factory 3')").Count() > 0)
                    {
                        grdInv2.DataSource = dtpivot.Select("Factory not in ('Factory 1','Factory 2','Factory 3')", "Factory, Plant, LINE, GRP").CopyToDataTable();
                        gvwInv2.ColumnPanelRowHeight = 25;
                        gvwInv2.RowHeight = 22;
                        for (int i = 0; i < gvwInv2.Columns.Count; i++)
                        {
                            GridColumn col = gvwInv2.Columns[i];
                            col.OptionsColumn.AllowMerge = DefaultBoolean.False;
                            col.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                            col.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Center;
                            col.Width = 90;
                            if (i < 2)
                                col.Visible = false;
                            else if (i < 6)
                            {
                                col.Width = 90;
                                if (i <= 4)
                                    col.OptionsColumn.AllowMerge = DefaultBoolean.True;
                            }
                            else
                            {
                                col.AppearanceCell.TextOptions.HAlignment = HorzAlignment.Far;
                                if (i == 6)
                                {
                                    col.Width = 75;
                                }
                                else if (i == 7)
                                {
                                    col.Width = 75;
                                }
                                else
                                {
                                    col.Width = 55;
                                }
                                col.DisplayFormat.FormatType = FormatType.Numeric;
                                col.DisplayFormat.FormatString = "#,#";
                            }

                            col.AppearanceHeader.Font = new Font("Calibri", 12, FontStyle.Regular);
                            col.AppearanceCell.Font = new Font("Calibri", 12, FontStyle.Regular);
                        }
                    }
                }

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

        

        private void CreateMail(string Subject, string htmlBody, DataTable dtEmail, string nameImg1, string nameImg2, string nameImg3, string nameImg4)
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\" + nameImg1 +   ".png", Outlook.OlAttachmentType.olByValue, null, "tr");

                //WriteLog(DateTime.Now.ToString() + " Chart Upper Inventory: " + nameImg1 + " ok");

                Outlook.Attachment oAttachPic2 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\" + nameImg2 + ".png", Outlook.OlAttachmentType.olByValue, null, "tr");
                //WriteLog(DateTime.Now.ToString() + " Yesterday Assembly Production: " + nameImg2 + " ok");

                Outlook.Attachment oAttachPic3 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\" + nameImg3 + ".png", Outlook.OlAttachmentType.olByValue, null, "tr");
                //WriteLog(DateTime.Now.ToString() + " Grid Upper Inventory: " + nameImg3 + " ok");

                Outlook.Attachment oAttachPic4 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\" + nameImg4 + ".png", Outlook.OlAttachmentType.olByValue, null, "tr");
                //WriteLog(DateTime.Now.ToString() + " Grid Upper Inventory: " + nameImg3 + " ok");

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
                        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(_emailTest[i]);
                        oRecip.Resolve();
                    }
                }
                oRecips = null;
                mailItem.BCC = "do.it@changshininc.com";
                string imgInfo1 = "imgInfo1";
                string imgInfo2 = "imgInfo2";
                string imgInfo3 = "imgInfo3";
                string imgInfo4 = "imgInfo4";
                oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo1);
                oAttachPic2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo2);
                oAttachPic3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo3);
                oAttachPic4.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo4);

                //mailItem.HTMLBody = String.Format(@"<img src='cid:{0}'>", imgInfo) + htmlBody;
                mailItem.HTMLBody = String.Format(
                   "<body>" +
                            "<img src=\"cid:{0}\">" +
                       "<br><br><img src=\"cid:{1}\">" +
                       "<br><br><img src=\"cid:{2}\">" +
                       "<br><img src=\"cid:{3}\">" +
                   "</body>",
                   imgInfo1, imgInfo2, imgInfo3, imgInfo4);

                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                // WriteLog("CreateMailMoldMonthWh: " + ex.ToString());
            }
        }

        private void gvwInv_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                GridView ex = sender as GridView;               

                if (e.Column.ColumnHandle > 2 && ex.GetRowCellValue(e.RowHandle, "Plant").ToString() != "Total")
                {
                    if (e.CellValue == null || e.CellValue.ToString() == "") return;                   
                    if (e.Column.FieldName.Contains("Total Inv"))
                    {
                        if (Convert.ToDouble(e.CellValue.ToString().Replace(",", "")) < 0)
                        {
                            e.Appearance.BackColor = Color.Red;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else if (Convert.ToDouble(e.CellValue.ToString().Replace(",", "")) <= Convert.ToDouble(ex.GetRowCellValue(e.RowHandle, "Target").ToString().Replace(",", "")))
                        {
                            e.Appearance.BackColor = Color.Green;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else if (Convert.ToDouble(e.CellValue.ToString().Replace(",", "")) > Convert.ToDouble(ex.GetRowCellValue(e.RowHandle, "Target").ToString().Replace(",", "")))
                        {
                            e.Appearance.BackColor = Color.Yellow;
                        }
                    }
                }
                if (e.Column.ColumnHandle > 2)
                {
                    if (ex.GetRowCellValue(e.RowHandle, "Plant").ToString() == "Total")
                    {
                        e.Appearance.BackColor = Color.LightCyan;
                        e.Appearance.ForeColor = Color.Coral;
                    }
                }
            }
            catch { }
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
    }
}
