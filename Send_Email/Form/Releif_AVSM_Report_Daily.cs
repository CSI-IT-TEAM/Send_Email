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
        private readonly string[] _emailTest = { "nguyen.it@changshininc.com", "jungbo.shim@dskorea.com" };
        Main frmMain = new Main();
        public DataTable _dtData, _dtEmail;

        private string GetHTML(DataTable dt)
        {
            try
            {
                string style = @"<head>
                                <style>
                                    table {
                                        font-family: 'Times New Roman', Times, serif;
                                        text-align: center;
                                        font-style: italic;
                                      }
                                      table td, table th {
                                        border: 0px;
                                        padding: 3px 2px;
                                        white-space: nowrap;
                                      }
                                      table tbody td {
                                        font-size: 20px;
                                      }
                                      table thead {
            
                                        font-style: italic;
                                        border-bottom: 0px solid #444444;
                                      }
                                      .TOTAL{
                                            background: #fff2b0;
                                        }
                                        .GTOTAL{
                                            background: #bdff7a;
                                        }
.RED{
    background:red;
    color:white;
}
.GREEN{
    background:green;
    color:white;
}
.YELLOW{
    background:yellow;
    color:black;
}
          
                                      table thead th {
                                        font-size: 19px;
                                        font-weight: bold;
                                        color: #F0F0F0;
                                        background: #26A1B2;
                                        text-align: center;
                                      }
                                      .tblBoder td, .tblBoder th{
                                        border: 1px solid #c0c0c0;
                                      }
                                      .bcGrade {
                                        background: #f5ba25;
                                        color: white;
                                      }
                                      .rework {
                                        background: #7260f7;
                                        color: white;
                                      }
                                      .info{
                                        font-family: 'Times New Roman', Times, serif;
                                        font-style: italic;
                                        font-weight: bold;
                                        font-size: 20px;
                                        color: #1f497d;
                                      }
                                      .name{
                                        font-family: 'Times New Roman', Times, serif;
                                        font-style: italic;
                                        font-weight: bold;
                                        font-size: 30px;
                                        color: BLACK;
                                      }
          
                                      .date{
                                          background:#ff9900;
                                          color:#ffffff
                                      }
                                </style>
                            </head>";
                string header = @"<tr>" +
                                  $"<th align='center' rowspan = '2'> Plant </th>" +
                                  $"<th align='center' rowspan = '2'> Line </th>" +
                                  $"<th align='center' rowspan = '2'> Area </th> " +
                                  $"<th align='center' rowspan = '2'> Process </th> " +
                                  $"<th align='center' rowspan = '2'> TO </th> " +
                                  $"<th align='center' colspan = '5' >PO</th> " +
                                  $"<th align='center' rowspan = '2'> Balance</th>" +
                                  $"</tr>" +
                                  $"<tr>" +
                                  $"<th align='center'> Workshop </th>" +
                                  $"<th align='center'> Relief </th>" +
                                  $"<th align='center'> Other Line </th>" +
                                  $"<th align='center'> Material Handler </th>" +
                                  $"<th align='center'> Total</th>" +
                                  $"</tr>";
                string body = string.Empty;
                string _PLANT_NAME_TEMP = string.Empty;
                string _LINE_CD_TEMP = string.Empty;
                string _AREA_NAME_TEMP = string.Empty;
                foreach (DataRow dr in dt.Rows)
                {
                    if (!dr["PLANT_NAME"].Equals(_PLANT_NAME_TEMP))
                    {
                        body += @"<tr class=" + dr["TOTAL_STYLESHEET"] + ">" +
                            $"<td  width: 80' align='center' rowspan ={dr["PLANT_NAME_CNT"]}>{dr["PLANT_NAME"]}</td>" +
                            $"<td  width: 80' align='center' rowspan ={dr["LINE_CD_CNT"]}>{dr["LINE_CD"]}</td>" +
                            $"<td  width: 80' align='center' rowspan ={dr["AREA_NAME_CNT"]}>{dr["AREA_NAME"]}</td>" +
                            $"<td  width: 80' align='center'>{dr["PROCESS_NAME"]}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["TO_QTY"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}",dr["PO_WS"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}",dr["PO_RELIEF"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}",dr["PO_OTHER_LINE"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}",dr["PO_MAT_HANDLER"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}",dr["PO_TOTAL"])}</td>" +
                            $"<td class={dr["BAL_COLOR"]} width: 80' align='center'>{string.Format("{0:n0}", dr["BALANCE"])}</td>" +
                            $"</tr>";
                        _PLANT_NAME_TEMP = dr["PLANT_NAME"].ToString();
                        _LINE_CD_TEMP = dr["LINE_CD"].ToString();
                        _AREA_NAME_TEMP = dr["AREA_NAME"].ToString();
                    }
                    else
                    {
                        if (!dr["LINE_CD"].Equals(_LINE_CD_TEMP))
                        {
                            body += @"
                          <tr class=" + dr["TOTAL_STYLESHEET"] + ">" +
                          $"<td  width: 80' align='center' rowspan ={dr["LINE_CD_CNT"]}>{dr["LINE_CD"]}</td>" +
                          $"<td  width: 80' align='center'>{dr["AREA_NAME"]}</td>" +
                          $"<td  width: 80' align='center'>{dr["PROCESS_NAME"]}</td>" +
                          $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["TO_QTY"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_WS"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_RELIEF"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_OTHER_LINE"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_MAT_HANDLER"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_TOTAL"])}</td>" +
                            $"<td class={dr["BAL_COLOR"]} width: 80' align='center'>{string.Format("{0:n0}", dr["BALANCE"])}</td>" +
                          $"</tr>";
                            _LINE_CD_TEMP = dr["LINE_CD"].ToString();
                            _AREA_NAME_TEMP = dr["AREA_NAME"].ToString();
                        }
                        else
                        {
                            if (!dr["AREA_NAME"].Equals(_AREA_NAME_TEMP))
                            {
                                body += @"
                          <tr class=" + dr["TOTAL_STYLESHEET"] + ">" +
                             $"<td  width: 80' align='center' rowspan ={dr["AREA_NAME_CNT"]}>{dr["AREA_NAME"]}</td>" +
                             $"<td  width: 80' align='center'>{dr["PROCESS_NAME"]}</td>" +
                             $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["TO_QTY"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_WS"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_RELIEF"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_OTHER_LINE"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_MAT_HANDLER"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_TOTAL"])}</td>" +
                            $"<td class={dr["BAL_COLOR"]} width: 80' align='center'>{string.Format("{0:n0}", dr["BALANCE"])}</td>" +
                             $"</tr>";
                                _AREA_NAME_TEMP = dr["AREA_NAME"].ToString();

                            }
                            else
                            {
                                body += @"
                            <tr class=" + dr["TOTAL_STYLESHEET"] + ">" +
                            $"<td  width: 80' align='center'>{dr["PROCESS_NAME"]}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["TO_QTY"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_WS"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_RELIEF"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_OTHER_LINE"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_MAT_HANDLER"])}</td>" +
                            $"<td  width: 80' align='center'>{string.Format("{0:n0}", dr["PO_TOTAL"])}</td>" +
                            $"<td class={dr["BAL_COLOR"]} width: 80' align='center'>{string.Format("{0:n0}", dr["BALANCE"])}</td>" +
                            $"</tr>";
                            }
                        }

                        _PLANT_NAME_TEMP = dr["PLANT_NAME"].ToString();
                    }
                }

                return @"<html>" + style + "" +
                                  $"<table class= 'tblBoder'>" +
                                  $"<thead>" + header + "</thead>" +
                                  $"<tbody>" + body + "</tbody>" +
                                $"</table>" +
                                $"</html>";
            }
            catch (Exception ex)
            { return string.Empty; }
        }

        private void Releif_AVSM_Report_Daily_Load(object sender, EventArgs e)
        {
            try
            {
                //SetData(_dtData);
                //CaptureControl(panel1, "ReliefReg");
                string _htmlBody = GetHTML(_dtData);
                CreateMail(_subject, _htmlBody, _dtEmail);
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
                    grdView2.Columns["AREA_NAME"].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.True;
                    // grdView2.Columns["PROCESS_NAME"].OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.True;
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
                panel1.Size = new Size(1100, (grdView2.RowCount * grdView2.RowHeight) + 130);
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
                string PlantNameValue = grdView2.GetRowCellValue(e.RowHandle, grdView2.Columns["PLANT_NAME"]).ToString();
                string AreaNameValue = grdView2.GetRowCellValue(e.RowHandle, grdView2.Columns["AREA_NAME"]).ToString();
                string BalColor = grdView2.GetRowCellValue(e.RowHandle, "BAL_COLOR").ToString();
                string BalTextColor = grdView2.GetRowCellValue(e.RowHandle, "BAL_TEXT_COLOR").ToString();
                if (PlantNameValue.ToUpper().Equals("G-TOTAL"))
                {
                    e.Appearance.BackColor = Color.FromArgb(250, 249, 200);
                    e.Appearance.Font = new Font("Calibri", 12F, FontStyle.Bold);
                }
                if (AreaNameValue.ToUpper().Equals("TOTAL"))
                {
                    e.Appearance.BackColor = Color.FromArgb(255, 237, 173);
                    e.Appearance.Font = new Font("Calibri", 12F, FontStyle.Bold);
                }
                if (e.Column.FieldName.Equals("BALANCE"))
                {
                    e.Appearance.BackColor = Color.FromName(BalColor);
                    e.Appearance.ForeColor = Color.FromName(BalTextColor);
                    e.Appearance.Font = new Font("Calibri", 12F, FontStyle.Bold);
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
                // Outlook.Attachment oAttachPic1 = mailItem.Attachments.Add(Application.StartupPath + @"\Capture\ReliefReg.png", Outlook.OlAttachmentType.olByValue, null, "tr");
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
                //  string imgInfo = "imgInfo";
                // oAttachPic1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo);
                //mailItem.HTMLBody = String.Format(@"<img src='cid:{0}'>", imgInfo) + htmlBody;

                mailItem.HTMLBody = htmlBody;
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
