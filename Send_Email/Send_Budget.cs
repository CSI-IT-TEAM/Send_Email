using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Send_Email
{
    class Send_Budget
    {
        public string _subject = "";
        public DataTable _email;
        public string Html(string argType)
        {
            try
            {
                string htmlReturn = "";

                DataSet dsData = SEL_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return "";
                //WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                DataTable dtData = dsData.Tables[0];
                DataTable dtHeader = dsData.Tables[1];
                DataTable dtExplain = dsData.Tables[2];
                _email = dsData.Tables[3];

                // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());



                htmlReturn = GetHtml(dtHeader, dtData, dtExplain.Rows[0]["STYLE"].ToString());

                _subject = dtExplain.Rows[0]["SUBJECT"].ToString();

                string explain = dtExplain.Rows[0]["TXT"].ToString();

                return explain + htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.ToString();
            }
            
        }

        private string GetHtml(DataTable arg_DtHeader, DataTable arg_DtData, string arg_Style)
        {
            try
            {
                string TableHeader = "";

                string HeaderRow1 = "";
                string HeaderRow2 = "";
                string[] headerArray = new string[arg_DtHeader.Rows.Count];
                //int iArray = 0;
                HeaderRow1 = "<tr>" +
                                $"<th bgcolor='#000099' style='color:#ffffff' rowspan = '2' align='center'>Head of Group</th>" +
                                $"<th bgcolor='#ff9900' style='color:#ffffff' colspan = '3' align='center'>2021-March</th>" +
                             "</tr>";
                HeaderRow2 = "<tr>" +
                                 $"<th bgcolor='#000099' style='color:#ffffff' align='center'> Target </th>" +
                                 $"<th bgcolor='#000099' style='color:#ffffff' align='center'> Actual </th>" +
                                 $"<th bgcolor='#000099' style='color:#ffffff' align='center'> Rate </th> " +
                             "</tr>";
                //foreach (DataRow row in dtHeader.Rows)
                //{
                //    HeaderRow1 += $"<th bgcolor = '{row["BCOLOR"]}' style = 'color:{row["FCOLOR"]}' align = 'center' width = '{row["WIDTH"]}'>{row["CAPTION"]}</th>";
                //    headerArray[iArray] = row["FIELD_NAME"].ToString();
                //    iArray++;
                //}

                TableHeader = "<thead>" + HeaderRow1 + HeaderRow2 + "</thead>";

                //Row
                string TableRow = "", bColor, fColor,bColorRate, fColorRate, dept, target, actual, rate;
                TableRow = "<tbody>";
                foreach (DataRow rowData in arg_DtData.Rows)
                {

                    bColor = rowData["BCOLOR"].ToString();
                    fColor = rowData["FCOLOR"].ToString();
                    bColorRate = rowData["BCOLOR_RATE"].ToString();
                    fColorRate = rowData["FCOLOR_RATE"].ToString();
                    dept = rowData["DEPT"].ToString();
                    target = rowData["PLAN_QTY"].ToString();
                    actual = rowData["ACTUAL_QTY"].ToString();
                    rate = rowData["RATE"].ToString();

                    TableRow += "<tr> ";
                    TableRow += $"<td bgcolor='{bColor}' style='color:{fColor}; width: 150' align='left'>{dept}</td>" +
                                $"<td bgcolor='{bColor}' style='color:{fColor}; width: 100' align='right'>{target}</td>" +
                                $"<td bgcolor='{bColor}' style='color:{fColor}; width: 100' align='right'>{actual}</td>" +
                                $"<td bgcolor='{bColorRate}' style='color:{fColorRate}; width: 100' align='right'>{rate}</td>";
                    TableRow += "</tr> ";
                }


                //foreach (DataRow rowData in dtData.Rows)
                //{
                //    TableRow += "<tr> ";
                //    foreach (DataRow rowHeader in dtHeader.Rows)
                //    {
                //        if (rowHeader["FIELD_NAME"].ToString() == "SCAN_FINISHED")
                //        {
                //            TableRow += $"<td bgcolor='{rowData["STATUS_BCOLOR"]}' style='color:{rowData["STATUS_FCOLOR"]}' align='{rowHeader["ALIGN"]}'>" +
                //                        $"{rowData[rowHeader["FIELD_NAME"].ToString()]}" +
                //                    $"</td>";
                //        }
                //        else
                //        {
                //            TableRow += $"<td bgcolor='WHITE' style='color:BLACK' align='{rowHeader["ALIGN"]}'>" +
                //                        $"{rowData[rowHeader["FIELD_NAME"].ToString()]}" +
                //                    $"</td>";
                //        }

                //    }

                //    TableRow += "</tr> ";

                //}

                    return "<Html>" +
                            arg_Style +
                             "<body>" +
                                "<table>" +
                                TableHeader + TableRow +
                                "</table>" +
                             "</body>"+
                         "<html>"
                       ;
            }
            catch (Exception ex)
            {
                // WriteLog("GetHtmlBodyCutting: " + ex.ToString());
                Debug.WriteLine(ex);
                return "";
            }
        }



        private DataSet SEL_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            MyOraDB.ShowErr = true;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_BUDGET";
                MyOraDB.ReDim_Parameter(7);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_LOC";
                MyOraDB.Parameter_Name[2] = "V_P_DATE";
                MyOraDB.Parameter_Name[3] = "CV_DATA";
                MyOraDB.Parameter_Name[4] = "CV_COL";
                MyOraDB.Parameter_Name[5] = "CV_SUBJECT";
                MyOraDB.Parameter_Name[6] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = V_P_DATE;
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";
                MyOraDB.Parameter_Values[6] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (V_P_TYPE == "Q")
                    {
                       // WriteLog("P_SEND_EMAIL_NPI: null");
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
