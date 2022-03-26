using DevExpress.Utils;
using DevExpress.XtraGrid.Views.BandedGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
    public partial class frmWeekly_Bottom_Constraint : Form
    {
        public frmWeekly_Bottom_Constraint()
        {
            InitializeComponent();
            tblOutsole.Size = new Size(4000, 2000);
        }
        public bool _chkTest = false;
        public string _subject = "";
        private string _subjectSend = "";
        private readonly string[] _emailTest = { "dien.it@changshininc.com", "ngoc.it@changshininc.com", "MAN.SPT@changshininc.com" };
        Main frmMain = new Main();
        public DataTable _dtChart1, _dtChart2, _dtChart21, _dtChart3, _dtChart4, _dtChart5, _dtChart6, _dtEmail, dtGrid2, dtGrid21;

        private void grdViewDOS_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                Rectangle rect = e.Bounds;
                rect.Inflate(new Size(1, 1));

                Brush brush = new SolidBrush(e.Appearance.BackColor);
                e.Graphics.FillRectangle(brush, rect);
                Pen pen_vertical = new Pen(Color.White, 2F);
                if (e.Column.ColumnHandle == 0)
                {
                    // draw left
                    e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y, rect.X, rect.Y + rect.Height);

                    e.Graphics.DrawString(e.DisplayText, new System.Drawing.Font("Calibri", 20F, FontStyle.Regular), new SolidBrush(Color.White), rect, e.Appearance.GetStringFormat());

                    e.Handled = true;
                }
            }
            catch { }
        }

        private void grdViewOS_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                Rectangle rect = e.Bounds;
                rect.Inflate(new Size(1, 1));

                Brush brush = new SolidBrush(e.Appearance.BackColor);
                e.Graphics.FillRectangle(brush, rect);
                Pen pen_vertical = new Pen(Color.White, 2F);


                if (e.RowHandle == 0)
                {
                    // raw top
                    e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y, rect.X + rect.Width, rect.Y);
                }

                if (e.RowHandle == grdViewOS.RowCount - 1)
                {
                    //raw bottom
                    //e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y + rect.Height - 1, rect.X + rect.Width, rect.Y + rect.Height - 1);
                }
                if (e.Column.ColumnHandle == 0)
                {
                    // draw right
                    //e.Graphics.DrawLine(pen_vertical, rect.X + rect.Width - 1, rect.Y, rect.X + rect.Width - 1, rect.Y + rect.Height);
                    // draw left
                    e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y, rect.X, rect.Y + rect.Height);
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
            Rectangle rect = e.Bounds;
            rect.Inflate(new Size(1, 1));

            Brush brush = new SolidBrush(e.Appearance.BackColor);
            e.Graphics.FillRectangle(brush, rect);
            Pen pen_vertical = new Pen(Color.White, 2F);
            Pen line = new Pen(Color.White, 2F);
            string[] ls = e.Column.Caption.Split('\n');
            // draw right
            e.Graphics.DrawLine(pen_vertical, rect.X + rect.Width - 2, rect.Y, rect.X + rect.Width - 2, rect.Y + rect.Height);

            //draw top
            e.Graphics.DrawLine(pen_vertical, rect.X + 1, rect.Y + 2, rect.X + rect.Width, rect.Y + 2);

            //raw bottom
            e.Graphics.DrawLine(pen_vertical, rect.X, rect.Y + rect.Height - 2, rect.X + rect.Width, rect.Y + rect.Height - 2);

            if(e.Column.ColumnHandle == 0)
            {
                // draw left
                e.Graphics.DrawLine(pen_vertical, rect.X + 2, rect.Y, rect.X + 2, rect.Y + rect.Height);
            }

            e.Graphics.DrawString(CultureInfo.CurrentCulture.TextInfo.ToTitleCase(e.Column.GetCaption().ToLower()), e.Appearance.GetFont(), new SolidBrush(e.Appearance.GetForeColor()), rect, e.Appearance.GetStringFormat());
            e.Handled = true;
        }

        private void frmWeekly_Bottom_Constraint_Load(object sender, EventArgs e)
        {
            try
            {
                LoadWeeklyBC(_dtChart1, _dtChart2, dtGrid2, dtGrid21, _dtChart3, _dtChart4, _dtChart5, _dtChart6);
                CaptureControl(tblOutsole, "OSMonthly");
                CreateMail(_subject, "", _dtEmail);
            }
            catch (Exception)
            {

                throw;
            }
        }

        private bool LoadWeeklyBC(DataTable argChart1Dt,
                                DataTable argChart2Dt,
                                DataTable argGrid2,
                                DataTable argGrid21,
                                DataTable argChart3Dt,
                                DataTable argChart4Dt,
                                DataTable argChart5Dt,
                                DataTable argChart6Dt)
        {
            try
            {
                SetChart("CHART1", argChart1Dt);
                SetChart("CHART2", argChart2Dt);
                SetData("GRID2",argGrid2);
                SetChart("CHART21", argGrid2);
                SetData("GRID21",argGrid21);
                //SetChart("CHART4", argChart4Dt);
                //SetChart("CHART5", argChart5Dt);
                //SetChart("CHART6", argChart6Dt);

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
                        chartControl1.Series[0].ArgumentDataMember = "DIV";
                        chartControl1.Series[0].ValueDataMembers.AddRange(new string[] { "CON_QTY" });
                        chartControl1.Series[1].ArgumentDataMember = "DIV";
                        chartControl1.Series[1].ValueDataMembers.AddRange(new string[] { "PER" });
                        break;
                    case "CHART2":
                        chartControl2.DataSource = argDt;
                        chartControl2.Series[0].ArgumentDataMember = "MODEL_NAME";
                        chartControl2.Series[0].ValueDataMembers.AddRange(new string[] { "QTY" });
                        break;
                    case "CHART21":
                        chartControl21.DataSource = argDt;
                        chartControl21.Series[0].ArgumentDataMember = "DAY_DIV";
                        chartControl21.Series[0].ValueDataMembers.AddRange(new string[] { "PER_1" });
                        break;
                    //case "CHART3":
                    //    chartControl3.DataSource = argDt;
                    //    chartControl3.Series[0].ArgumentDataMember = "REASON_NAME";
                    //    chartControl3.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });

                    //    break;
                    //case "CHART4":
                    //    chartControl4.DataSource = argDt;
                    //    chartControl4.Series[0].ArgumentDataMember = "RESOURCE_CD";
                    //    chartControl4.Series[0].ValueDataMembers.AddRange(new string[] { "HOURS" });

                    //    break;
                    //case "CHART5":
                    //    chartControl5.DataSource = argDt;
                    //    chartControl5.Series[0].ArgumentDataMember = "SHIFT";
                    //    chartControl5.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });

                    //    break;
                    //case "CHART6":
                    //    chartControl6.DataSource = argDt;
                    //    chartControl6.Series[0].ArgumentDataMember = "YMD";
                    //    chartControl6.Series[0].ValueDataMembers.AddRange(new string[] { "CNT" });

                    //  //  DataTable dt = Pivot(argDt, argDt.Columns["YMD"], argDt.Columns["CNT"]);
                    //    //createGrid(dt);
                    //    break;
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

        private void SetData(string grid_type, DataTable dtData)
        {
            try
            {
                if (grid_type == "GRID2")
                {
                    while (grdViewOS.Columns.Count > 0)
                    {
                        grdViewOS.Columns.RemoveAt(0);
                    }


                    grdMainOS.DataSource = dtData;
                    formatGrid("GRID2");
                }
                else if (grid_type == "GRID21")
                {
                    while (grdViewDOS.Columns.Count > 0)
                    {
                        grdViewDOS.Columns.RemoveAt(0);
                    }

                    grdDayOS.DataSource = dtData;
                    formatGrid("GRID21");
                }
            }
            catch
            {

            }
        }

        private void formatGrid(string strgrid_type)
        {
            try
            {
                if (strgrid_type == "GRID2")
                {
                    for (int i = 0; i < grdViewOS.Columns.Count; i++)
                    {
                        grdViewOS.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                        grdViewOS.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        grdViewOS.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                        grdViewOS.Columns[i].Width = 150;
                        //if (i >= 7)
                        //{
                        //    grdViewOS.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                        //    grdViewOS.Columns[i].DisplayFormat.FormatString = "{0:n0}";
                        //}

                    }
                    grdViewOS.Columns[0].Visible = false;
                    grdViewOS.Columns[1].Visible = false;
                    grdViewOS.Columns[2].Width = 250;
                    grdViewOS.RowHeight = 70;
                    grdViewOS.Appearance.Row.Font = new System.Drawing.Font("Calibri", 30F, System.Drawing.FontStyle.Regular);
                }
                else if (strgrid_type == "GRID21")
                {
                    //grdViewDOS.BeginUpdate();
                    for (int i = 0; i < grdViewDOS.Columns.Count ; i++)
                    {
                        grdViewDOS.Columns[i].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                        grdViewDOS.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        grdViewDOS.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                        grdViewDOS.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        grdViewDOS.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                        grdViewDOS.Columns[i].Width = 220;

                        //grdViewDOS.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                        //grdViewDOS.Columns[i].DisplayFormat.FormatString = "{0:n0}";
                        grdViewDOS.Columns[i].AppearanceHeader.Font = new System.Drawing.Font("Calibri", 20F, System.Drawing.FontStyle.Regular);
                    //    grdViewOS.Columns[i].AppearanceCell.Font = new System.Drawing.Font("Calibri", 30F, System.Drawing.FontStyle.Regular);

                    }
                    grdViewDOS.ColumnPanelRowHeight = 30;
                    grdViewDOS.RowHeight = 30;
                    grdViewDOS.Appearance.Row.Font = new System.Drawing.Font("Calibri", 20F, System.Drawing.FontStyle.Regular);
                    //grdViewDOS.EndUpdate();

                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
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
