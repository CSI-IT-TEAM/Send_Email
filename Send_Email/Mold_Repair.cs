using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Linq;
using System.Text;

namespace Send_Email
{
    class Mold_Repair
    {
        public string Html_MoldRepair(string argType)
        {
            string htmlReturn = "";

            DataSet dsData = SEL_MOLD_REPAIR(argType, DateTime.Now.ToString("yyyyMMdd"));
            if (dsData == null) return "";
            //WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            DataTable dtData = dsData.Tables[0];
            DataTable dtHeader = dsData.Tables[1];
            DataTable dtExplain = dsData.Tables[2];
            DataTable dtEmail = dsData.Tables[3];

           // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

            string html = GetHtmlBodyNpi(dtHeader, dtData);

            string subject = dtExplain.Rows[0]["SUBJECT"].ToString();

            string explain = dtExplain.Rows[0]["TXT"].ToString();

            return htmlReturn;
        }

        private string GetHtmlBodyNpi(DataTable dtHeader, DataTable dtData)
        {
            try
            {
                string TableHeader = "";

                int i = 0;
                int npiCode = 0;
                string HeaderRow1 = "", HeaderRow2 = "";

                int[] colWidth = { 67, 67, 67, 67, 67, 67, 67, 80, 67, 80, 67, 67, 67, 67, 67, 67, 75, 67 };

                foreach (DataRow row in dtHeader.Rows)
                {
                    int.TryParse(row["NPI_CODE"].ToString(), out npiCode);
                    if (npiCode >= 10)
                    {

                        HeaderRow1 += $"<td bgcolor = '#ff9900' style = 'color:#ffffff' align = 'center' width = '{colWidth[i]}'>{row["NPI_DATE"]}</td>";
                        HeaderRow2 += $"<td bgcolor = '#00ced1' style = 'color:#ffffff' align = 'center' width = '{colWidth[i]}'>{row["NPI_NAME"]}</td>";
                        i++;
                    }

                    else
                        HeaderRow1 += $"<td bgcolor = '#00ced1' style = 'color:#ffffff' rowspan ='2' align = 'center'>{row["NPI_NAME"]}</td>";
                }

                TableHeader = "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow1 + "</tr> " +
                              "<tr style='font-family:Calibri; font-size:14px'> " + HeaderRow2 + "</tr> ";

                //Row
                string TableRow = "";

                for (int iRowData = 0; iRowData < dtData.Rows.Count; iRowData++)
                {
                    string plantNm = dtData.Rows[iRowData]["PLANT_NM"].ToString();
                    string lineCd = dtData.Rows[iRowData]["LINE_CD"].ToString();
                    string category = dtData.Rows[iRowData]["CATEGORY_NAME"].ToString();
                    string tdCode = dtData.Rows[iRowData]["TD_CODE"].ToString();
                    string modelName = dtData.Rows[iRowData]["MODEL_NAME"].ToString();
                    string styleCode = dtData.Rows[iRowData]["STYLE_CODE"].ToString();
                    string prodDate = dtData.Rows[iRowData]["PROD_DATE"].ToString();
                    string backColor = dtData.Rows[iRowData]["BCOLOR"].ToString();

                    string plantNm_prev = "";
                    string lineCd_prev = "";
                    string category_prev = "";
                    string tdCode_prev = "";
                    string modelName_prev = "";
                    string styleCode_prev = "";
                    string prodDate_prev = "";

                    if (iRowData > 0)
                    {
                        plantNm_prev = dtData.Rows[iRowData - 1]["PLANT_NM"].ToString();
                        lineCd_prev = dtData.Rows[iRowData - 1]["LINE_CD"].ToString();
                        category_prev = dtData.Rows[iRowData - 1]["CATEGORY_NAME"].ToString();
                        tdCode_prev = dtData.Rows[iRowData - 1]["TD_CODE"].ToString();
                        modelName_prev = dtData.Rows[iRowData - 1]["MODEL_NAME"].ToString();
                        styleCode_prev = dtData.Rows[iRowData - 1]["STYLE_CODE"].ToString();
                        prodDate_prev = dtData.Rows[iRowData - 1]["PROD_DATE"].ToString();
                    }

                    if (plantNm != plantNm_prev || lineCd != lineCd_prev || category != category_prev || tdCode != tdCode_prev ||
                        modelName != modelName_prev || styleCode != styleCode_prev || prodDate != prodDate_prev)
                    {
                        if (iRowData > 0) TableRow += "</tr> ";
                        TableRow += "<tr> " +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left'>{plantNm }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='center' >{ lineCd }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left'   >{ category }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='center' >{ tdCode }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left' >{styleCode }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='left'   >{ modelName }</td>" +
                                $"<td bgcolor='WHITE' style='color:BLACK' align='center' >{ prodDate }</td>" +
                                $"<td bgcolor='{backColor}' style='color:BLACK' width = '50' align='center' ></td>"
                              ;
                    }
                    else
                    {
                        TableRow += $"<td bgcolor='{backColor}' style='color:BLACK' width = '50' align='center'></td>";
                    }
                }

                return "<table style='font-family:Calibri; font-size:15px' bgcolor='#f5f3ed' border='1' cellpadding='0' cellspacing='0' with = '5000' >" +
                            TableHeader + TableRow +
                       "</table>";
            }
            catch (Exception ex)
            {
               // WriteLog("GetHtmlBodyCutting: " + ex.ToString());
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
