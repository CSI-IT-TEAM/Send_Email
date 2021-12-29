using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;

namespace Send_Email
{
    class Send_Hourly_Prod_Tracking
    {
        public string _subject = "";
        public DataTable _email;
        public string Html(string argType)
        {
            try
            {
                string htmlReturn = "";

                DataSet dsData = SEL_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null || dsData.Tables.Count <= 1) return "";
                //WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                DataTable dtData1 = dsData.Tables[0];
                DataTable dtData2 = dsData.Tables[1];
                DataTable dtData3 = dsData.Tables[2];
                DataTable dtData4 = dsData.Tables[3];

                DataTable dtHtml = dsData.Tables[4];
                _email = dsData.Tables[5];

                // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                htmlReturn = GetHtml(dtData1, dtData2,  dtData3, dtData4, dtHtml);

                _subject = dtHtml.Rows[0]["ATTRIB1"].ToString();

                

                return htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }

        }
        private string GetHtml(DataTable arg_DtData1, DataTable arg_DtData2, DataTable arg_DtData3, DataTable arg_DtData4, DataTable arg_DtHtml)
        {
            try
            {
                string htmlReturn = arg_DtHtml.Rows[0]["TEXT1"].ToString();

                string strTBody1 = GetDataTboby(arg_DtData1, arg_DtHtml, 1);
                string strTBody2= GetDataTboby(arg_DtData2, arg_DtHtml, 1);
                string strTBody3 = GetDataTboby(arg_DtData3, arg_DtHtml, 1);
                string strTBody4 = GetDataTboby(arg_DtData4, arg_DtHtml, 1);

                htmlReturn = htmlReturn.Replace("{tbody1}", strTBody1);
                htmlReturn = htmlReturn.Replace("{tbody2}", strTBody2);
                htmlReturn = htmlReturn.Replace("{tbody3}", strTBody3);
                htmlReturn = htmlReturn.Replace("{tbody4}", strTBody4);

                return htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        private string GetDataTboby(DataTable argDtData, DataTable argDtHtml, int row = 1)
        {
            string strTbodyRtn = "";
            try
            {
                int rowSpanFactory = 1;
                int rowSpanPlan = 1;
                string factoryPre = "", plantPre = "";
                string factory = "", plant = "";
                string strRow = "";

                string rowCol1Span = argDtHtml.Rows[row]["TEXT1"].ToString();
                string rowCol2Span = argDtHtml.Rows[row]["TEXT2"].ToString();
                string rowColMerge = argDtHtml.Rows[row]["TEXT3"].ToString();
                string rowRowSpan = argDtHtml.Rows[row]["TEXT4"].ToString();

                foreach (DataRow rowData in argDtData.Rows)
                {
                    factory = rowData["FACTORY"].ToString();
                    plant = rowData["PLANT"].ToString();

                    if (factoryPre == "" || factory != factoryPre)
                    {
                        factoryPre = factory;
                        plantPre = plant;
                        strRow = rowCol1Span;

                        rowSpanFactory = (int)argDtData.Compute("COUNT(FACTORY)", $"FACTORY ='{factory}' ");
                        rowSpanPlan = (int)argDtData.Compute("COUNT(PLANT)", $" PLANT ='{plant}'");
                        fnReplace(ref strRow, "{COL1_SPAN}", rowSpanFactory.ToString());
                        fnReplace(ref strRow, "{COL2_SPAN}", rowSpanPlan.ToString());
                       // fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                       // fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                        strTbodyRtn += fnReplaceRow(strRow, rowData);
                    }
                    else
                    {
                        if (plantPre == "" || plant != plantPre)
                        {
                            factoryPre = factory;
                            plantPre = plant;
                            strRow = rowCol2Span;

                            rowSpanPlan = (int)argDtData.Compute("COUNT(PLANT)", $" PLANT ='{plant}'");
                            fnReplace(ref strRow, "{COL2_SPAN}", rowSpanPlan.ToString());
                           // fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                           // fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                            strTbodyRtn += fnReplaceRow(strRow, rowData);
                        }
                        else
                        {
                            factoryPre = factory;
                            plantPre = plant;
                            strRow = rowColMerge;
                           // fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                           // fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                            strTbodyRtn += fnReplaceRow(strRow, rowData);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return strTbodyRtn;
        }

        private void fnReplace(ref string argText, string argOldChar, string argNewChar)
        {
            argText = argText.Replace(argOldChar, argNewChar);
        }

        private string fnReplaceRow(string argText, DataRow argDtRow)
        {
            try
            {
                string strReturn = argText;
                foreach (DataColumn column in argDtRow.Table.Columns)
                {
                    strReturn = strReturn.Replace("{" + column.ColumnName + "}", argDtRow[column.ColumnName].ToString());
                }

                return strReturn;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return null;
                
            }
            
        }



        private DataSet SEL_DATA(string V_P_TYPE, string V_P_DATE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ShowErr = true;
            DataSet ds_ret;
            try
            {
                string process_name = "P_EMAIL_HOUR_PROD_TRACKING";
                MyOraDB.ReDim_Parameter(9);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "V_P_LOC";
                MyOraDB.Parameter_Name[2] = "V_P_DATE";
                MyOraDB.Parameter_Name[3] = "CV_DATA";
                MyOraDB.Parameter_Name[4] = "CV_DATA2";
                MyOraDB.Parameter_Name[5] = "CV_DATA3";
                MyOraDB.Parameter_Name[6] = "CV_DATA4";
                MyOraDB.Parameter_Name[7] = "CV_SUBJECT";
                MyOraDB.Parameter_Name[8] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[6] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[7] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[8] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = V_P_DATE;
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";
                MyOraDB.Parameter_Values[6] = "";
                MyOraDB.Parameter_Values[7] = "";
                MyOraDB.Parameter_Values[8] = "";

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
