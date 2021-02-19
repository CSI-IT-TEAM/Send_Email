using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Send_Email
{
    class Mold_Repair
    {
        public string _subject = "";
        public DataTable _email;
        public string Html_MoldRepair(string argType)
        {
            try
            {
                string htmlReturn = "";

                DataSet dsData = SEL_MOLD_REPAIR(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null) return "";
                //WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                DataTable dtData = dsData.Tables[0];
                DataTable dtHeader = dsData.Tables[1];
                DataTable dtExplain = dsData.Tables[2];
                _email = dsData.Tables[3];

                // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                htmlReturn = GetHtmlBodyMoldRepair(dtHeader, dtData);

                _subject = dtExplain.Rows[0]["SUBJECT"].ToString();

                string explain = dtExplain.Rows[0]["TXT"].ToString();

                return explain + htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.ToString();
            }
            
        }

        private string GetHtmlBodyMoldRepair(DataTable dtHeader, DataTable dtData)
        {
            try
            {
                string TableHeader = "";

               // int i = 0;
                //int npiCode = 0;
                string HeaderRow1 = "";

                // int[] colWidth = { 67, 67, 67, 67, 67, 67, 67, 80, 67, 80, 67, 67, 67, 67, 67, 67, 75, 67 };
                string[] headerArray = new string[dtHeader.Rows.Count];
                int iArray = 0;
                foreach (DataRow row in dtHeader.Rows)
                {
                    HeaderRow1 += $"<th bgcolor = '{row["BCOLOR"]}' style = 'color:{row["FCOLOR"]}' align = 'center' width = '{row["WIDTH"]}'>{row["CAPTION"]}</th>";
                    headerArray[iArray] = row["FIELD_NAME"].ToString();
                    iArray++;
                }

                TableHeader = "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow1 + "</tr> " ;

                //Row
                string TableRow = "";

                foreach (DataRow rowData in dtData.Rows)
                {
                    TableRow += "<tr> ";
                    foreach (DataRow rowHeader in dtHeader.Rows)
                    {
                        if (rowHeader["FIELD_NAME"].ToString() == "SCAN_FINISHED")
                        {
                            TableRow += $"<td bgcolor='{rowData["STATUS_BCOLOR"]}' style='color:{rowData["STATUS_FCOLOR"]}' align='{rowHeader["ALIGN"]}'>" +
                                        $"{rowData[rowHeader["FIELD_NAME"].ToString()]}" +
                                    $"</td>";
                        }
                        else
                        {
                            TableRow += $"<td bgcolor='WHITE' style='color:BLACK' align='{rowHeader["ALIGN"]}'>" +
                                        $"{rowData[rowHeader["FIELD_NAME"].ToString()]}" +
                                    $"</td>";
                        }
                        
                    }
                    
                    TableRow += "</tr> ";

                }

                return "<table style='font-family:Calibri; font-size:15px' bgcolor='#f5f3ed' border='1' cellpadding='2' cellspacing='0' >" +
                            TableHeader + TableRow +
                       "</table>";
            }
            catch (Exception ex)
            {
                // WriteLog("GetHtmlBodyCutting: " + ex.ToString());
                Debug.WriteLine(ex);
                return "";
            }
        }



        private DataSet SEL_MOLD_REPAIR(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();

            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_MOLD_REPAIR";
                MyOraDB.ReDim_Parameter(7);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_LOC";
                MyOraDB.Parameter_Name[2] = "V_P_DATE";
                MyOraDB.Parameter_Name[3] = "CV_1";
                MyOraDB.Parameter_Name[4] = "CV_2";
                MyOraDB.Parameter_Name[5] = "CV_EXPLAIN";
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
