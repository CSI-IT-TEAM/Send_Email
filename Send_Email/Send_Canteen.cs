using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;

namespace Send_Email
{
    class Send_Canteen
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
                DataTable dtDataVj1 = dsData.Tables[0];
                DataTable dtDataBot = dsData.Tables[1];
                DataTable dtDataVj2 = dsData.Tables[2];
                DataTable dtDataVj3 = dsData.Tables[3];

                DataTable dtHtml = dsData.Tables[4];
                _email = dsData.Tables[5];

                // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                htmlReturn = GetHtml(dtDataBot, dtDataVj1, dtDataVj2, dtDataVj3, dtHtml);

                _subject = dtHtml.Rows[0]["ATTRIB1"].ToString();

                

                return htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }

        }
        private string GetHtml(DataTable arg_DtDataBot, DataTable arg_DtDataVj1, DataTable arg_DtDataVj2, DataTable arg_DtDataVj3, DataTable arg_DtHtml)
        {
            try
            {
                string htmlReturn = arg_DtHtml.Rows[0]["TEXT1"].ToString();

                string rowCol1Span = arg_DtHtml.Rows[1]["TEXT1"].ToString();
                string rowCol2Span = arg_DtHtml.Rows[1]["TEXT2"].ToString();
                string rowColMerge = arg_DtHtml.Rows[1]["TEXT3"].ToString();
                string rowRowSpan = arg_DtHtml.Rows[1]["TEXT4"].ToString();

                int rowSpanFactory = 1;
                int rowSpanPlan = 1;
                string factoryPre = "", plantPre = "";
                string factory ="", plant = "";
                string strRow = "";

                string strTBody1 = "";
                string strTBody2= "";
                string strTBody3 = "";

                //DataTable dtDataVsm = arg_DtDataVj1.Select("FACTORY in ('F1','F2','F3','F4','F5')").CopyToDataTable();
               // DataTable dtDataVj2 = arg_DtDataVj1.Select("FACTORY = 'VJ2'").CopyToDataTable();

                foreach (DataRow rowData in arg_DtDataVj1.Rows)
                {
                    factory = rowData["FACTORY"].ToString();
                    plant = rowData["PLANT"].ToString();

                    if (factoryPre == "" || factory != factoryPre)
                    {
                        factoryPre = factory;
                        plantPre = plant;
                        strRow = rowCol1Span;

                        rowSpanFactory = (int)arg_DtDataVj1.Compute("COUNT(FACTORY)", $"FACTORY ='{factory}' ");
                        rowSpanPlan = (int)arg_DtDataVj1.Compute("COUNT(PLANT)", $" PLANT ='{plant}'");                        
                        fnReplace(ref strRow, "{COL1_SPAN}", rowSpanFactory.ToString());
                        fnReplace(ref strRow, "{COL2_SPAN}", rowSpanPlan.ToString());
                        fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                        fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                        strTBody1 += fnReplaceRow(strRow, rowData);
                    }
                    else
                    {
                        if (plantPre == "" || plant != plantPre)
                        {
                            factoryPre = factory;
                            plantPre = plant;
                            strRow = rowCol2Span;

                            rowSpanPlan = (int)arg_DtDataVj1.Compute("COUNT(PLANT)", $" PLANT ='{plant}'");
                            fnReplace(ref strRow, "{COL2_SPAN}", rowSpanPlan.ToString());
                            fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                            fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                            strTBody1 += fnReplaceRow(strRow, rowData);
                        }
                        else
                        {
                            factoryPre = factory;
                            plantPre = plant;
                            strRow = rowColMerge;
                            fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                            fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                            strTBody1 += fnReplaceRow(strRow, rowData);
                        }
                    }
                }


                foreach (DataRow rowData in arg_DtDataVj2.Rows)
                {
                    factory = rowData["FACTORY"].ToString();
                    plant = rowData["PLANT"].ToString();

                    if (factoryPre == "" || factory != factoryPre)
                    {
                        factoryPre = factory;
                        plantPre = plant;
                        strRow = rowCol1Span;

                        rowSpanFactory = (int)arg_DtDataVj2.Compute("COUNT(FACTORY)", $"FACTORY ='{factory}' ");
                        rowSpanPlan = (int)arg_DtDataVj2.Compute("COUNT(PLANT)", $" PLANT ='{plant}'");
                        fnReplace(ref strRow, "{COL1_SPAN}", rowSpanFactory.ToString());
                        fnReplace(ref strRow, "{COL2_SPAN}", rowSpanPlan.ToString());
                        fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                        fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                        strTBody3 += fnReplaceRow(strRow, rowData);
                    }
                    else
                    {
                        if (plantPre == "" || plant != plantPre)
                        {
                            factoryPre = factory;
                            plantPre = plant;
                            strRow = rowCol2Span;

                            rowSpanPlan = (int)arg_DtDataVj2.Compute("COUNT(PLANT)", $" PLANT ='{plant}'");
                            fnReplace(ref strRow, "{COL2_SPAN}", rowSpanPlan.ToString());
                            fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                            fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                            strTBody3 += fnReplaceRow(strRow, rowData);
                        }
                        else
                        {
                            factoryPre = factory;
                            plantPre = plant;
                            strRow = rowColMerge;
                            fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                            fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                            strTBody3 += fnReplaceRow(strRow, rowData);
                        }
                    }
                }

                rowCol1Span = arg_DtHtml.Rows[2]["TEXT1"].ToString();
                rowColMerge = arg_DtHtml.Rows[2]["TEXT3"].ToString();
                rowRowSpan = arg_DtHtml.Rows[2]["TEXT4"].ToString();

                foreach (DataRow rowData in arg_DtDataBot.Rows)
                {
                    factory = rowData["FACTORY"].ToString();
                    plant = rowData["PLANT"].ToString();

                    if (factoryPre == "" || factory != factoryPre)
                    {
                        factoryPre = factory;
                        strRow = rowCol1Span;

                        rowSpanFactory = (int)arg_DtDataBot.Compute("COUNT(FACTORY)", $"FACTORY ='{factory}' ");
                        fnReplace(ref strRow, "{COL1_SPAN}", rowSpanFactory.ToString());
                        fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                        fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                        strTBody2 += fnReplaceRow(strRow, rowData);
                    }
                    else
                    {
                        factoryPre = factory;
                        strRow = rowColMerge;
                        fnReplace(ref strRow, "{BCOLOR}", rowData["BCOLOR"].ToString());
                        fnReplace(ref strRow, "{FCOLOR}", rowData["FCOLOR"].ToString());
                        strTBody2 += fnReplaceRow(strRow, rowData);
                    }
                }
                
                htmlReturn = htmlReturn.Replace("{ABS_VSM_CNT}", arg_DtDataVj1.Rows[0]["ABS_VSM_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{IE_VSM_CNT}", arg_DtDataVj1.Rows[0]["IE_VSM_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{ABS_BOT_CNT}", arg_DtDataVj1.Rows[0]["ABS_BOT_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{IE_BOT_CNT}", arg_DtDataVj1.Rows[0]["IE_BOT_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{ABS_VJ2_CNT}", arg_DtDataVj1.Rows[0]["ABS_VJ2_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{IE_VJ2_CNT}", arg_DtDataVj1.Rows[0]["IE_VJ2_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{ABS_TOT_CNT}", arg_DtDataVj1.Rows[0]["ABS_TOT_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{IE_TOT_CNT}", arg_DtDataVj1.Rows[0]["IE_TOT_CNT"].ToString());
                htmlReturn = htmlReturn.Replace("{tbody1}", strTBody1);
                htmlReturn = htmlReturn.Replace("{tbody2}", strTBody2);
                htmlReturn = htmlReturn.Replace("{tbody3}", strTBody3);

                return htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
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
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            MyOraDB.ShowErr = true;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_CANTEEN";
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
