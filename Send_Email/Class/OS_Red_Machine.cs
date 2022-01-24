using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Send_Email
{
    class OS_Red_Machine
    {
        public string _subject = "";
        public DataTable _email;
        public string Html(string argType,string argDate,string argHH)
        {
            try
            {
                string htmlReturn = "";

                DataSet dsData = SEL_DATA(argType, argDate, argHH);
                if (dsData == null || dsData.Tables[1].Rows.Count <=0) return "";
                //WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                DataTable dtData = dsData.Tables[1];
                DataTable dtHeader = dsData.Tables[0];
                _email = dsData.Tables[2];

                // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                htmlReturn = GetHtmlBody(dtHeader, dtData);

                
                _subject = "Outsole press machine drawback list";
                //_subject = "(Test Email) Outsole press machine drawback list";
                return htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.ToString();
            }
            
        }

        private string GetHtmlBody(DataTable dtHeader, DataTable dtData)
        {
            try
            {
                string StyleSheet = @"<html><head><style>table.OSPTable {
                                  font-family: 'Times New Roman', Times, serif;
                                  text-align: center;
                                }
                                table.OSPTable td, table.OSPTable th {
                                  border: 1px solid #c0c0c0;
                                  padding: 3px 2px;
                                  white-space: nowrap;
                                }
                                table.OSPTable tbody td {
                                  font-size: 20px;
                                }
                                table.OSPTable thead {
                                  background: #26A1B2;
                                  font-style: italic;
                                  border-bottom: 0px solid #444444;
                                }
                                table.OSPTable tbody {
                                  font-style: italic;
                                }
                                table.OSPTable thead th {
                                  font-size: 19px;
                                  font-weight: bold;
                                  color: #F0F0F0;
                                  background: #26A1B2;
                                  text-align: center;
                                }
                                .info{
                                  font-family: 'Times New Roman', Times, serif;
                                  font-style: italic;
                                  font-size: 24px;
                                  color: #1a6e79;
                                }
                                table.OSPTable thead th.pic{
                                   color: black;
                                   background-color: orange;
                                   width: 230px;
                                }
                                table.OSPTable thead th.reason{
                                   color: rgb(255, 255, 255);
                                   background-color: rgb(255, 0, 0);
                                   width: 500px;
                                }
                                table.OSPTable tbody td.pic {
                                  font-size: 20px;
                                  font-weight: bold;
                                }
                                </style></head>";

                string TableHeader = string.Format(@"<body>
                                       <!-- <div class='info'>
                                        4시간 동안 아웃솔 프레스 실적 이 interface 되지 않으면 자동 으로 메일이 담당자 들에게 발송 처리가 된다.</div></br>

                                        <div class='info'>Without pressing actual during 4 hours were mailed out to PIC promptly.</div></br> -->

                                        <div class='info'>&nbsp;Số lượng sản xuất không đạt Target trong 4 giờ sẽ gửi Email</div></br>
                             

                                        <table class='OSPTable'>
                                        <thead>
                                        <tr>
                                        <th rowspan='2'>Line</th>
                                        <th rowspan='2'>Machine</th>
                                        <th colspan='2'>{0}</th>
                                        <th colspan='2'>{1}</th>
                                        <th rowspan='2' class='pic'>PIC</th>
                                        <th rowspan='2' class ='reason'>Reason</th>
                                        </tr>
                                        <tr>
                                        <th width='80px'>Plan</th>
                                        <th width='80px'>Actual</th>
                                        <th width='80px'>Plan</th>
                                        <th width='80px'>Actual</th>
                                        </tr>
                                        </thead><tbody>", dtHeader.Rows[0][0], dtHeader.Rows[0][1]);
                //Row
                string TableRow = "";
                foreach (DataRow row in dtData.Rows)
                {
                    TableRow += $"<tr><td>{row["LINE"]}</td><td > {row["MC"]} </td><td>{string.Format("{0:n0}",row["COL1_PLAN"].ToString())} </td><td> {string.Format("{0:n0}", row["COL1_ACT"].ToString())}</ td><td> {string.Format("{0:n0}", row["COL2_PLAN"].ToString())} </td ><td> {string.Format("{0:n0}", row["COL2_ACT"].ToString())} </td ><td class='pic'>{row["STATUS"].ToString()}</td><td>{row["REASON"].ToString()}</td></tr>";
                }

                string EndTag = "</tbody></table></body></html>";
                string CompletedHTML = string.Concat(StyleSheet, TableHeader, TableRow, EndTag);
                return CompletedHTML;
            }
            catch (Exception ex)
            {
                // WriteLog("GetHtmlBodyCutting: " + ex.ToString());
                Debug.WriteLine(ex);
                return "";
            }
        }

        private DataSet SEL_DATA(string V_P_TYPE, string V_P_DATE, string V_P_HH)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.SEPHIROTH;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_OUTSOLE_PRESS";
                MyOraDB.ReDim_Parameter(6);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_DATE";
                MyOraDB.Parameter_Name[2] = "V_P_HH";
                MyOraDB.Parameter_Name[3] = "CV_1";
                MyOraDB.Parameter_Name[4] = "CV_2";
                MyOraDB.Parameter_Name[5] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = V_P_DATE;
                MyOraDB.Parameter_Values[2] = V_P_HH;
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";

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
                return ds_ret;
            }
            catch (Exception ex)
            {
                // WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
    }
}
