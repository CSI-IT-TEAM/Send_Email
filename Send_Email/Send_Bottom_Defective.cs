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
    class Send_Bottom_Defective
    {
        public string _subject = "";
        public DataTable _email;
        public string Html(string argType)
        {
            try
            {
                string htmlReturn = "";

                DataSet dsData = SEL_DATA(argType, DateTime.Now.ToString("yyyyMMdd"));
                if (dsData == null || dsData.Tables.Count <=1) return "";
                //WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                DataTable dtData = dsData.Tables[0];
                DataTable dtData2 = dsData.Tables[1];
                DataTable dtHeader = dsData.Tables[2];
                DataTable dtExplain = dsData.Tables[3];
                _email = dsData.Tables[4];

                // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());



                htmlReturn = GetHtml(dtHeader, dtData, dtData2, dtExplain);

                _subject = dtExplain.Rows[0]["SUBJECT"].ToString();

                

                return htmlReturn;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.ToString();
            }
            
        }

        private string GetHtml(DataTable arg_DtHeader, DataTable arg_DtData, DataTable arg_DtData2, DataTable argStyle)
        {
            try
            {
               
                string TableRewBcg = "", TableRewBcgHeader = "", HeaderRewBcgRow1 = "", HeaderRewBcgRow2 ="", TableRewBcgRow = "";
                string bColor, fColor, bColorRew, fColorRew, bColorBcg, fColorBcg;
                
                string strStyle = argStyle.Rows[0]["STYLE"].ToString();
                string strTitle1 = argStyle.Rows[0]["TITLE1"].ToString();
                string strTitle2 = argStyle.Rows[0]["TITLE2"].ToString();
                string strExplain = argStyle.Rows[0]["TXT"].ToString();

                
                ///
                ///Bottom Part Defective Rate
                ///
                HeaderRewBcgRow1 = "<tr>" +
                                     $"<th rowspan = '2' align='center'> Plant </th>" +
                                     $"<th rowspan = '2' align='center'> Model Name </th>" +
                                     $"<th rowspan = '2' align='center'> Production<br>Quantity </th> " +
                                     $"<th rowspan = '2' align='center'> Defective </th> " +
                                     $"<th rowspan = '2' align='center'> Rate </th> " +
                                     $"<th colspan = '4' align='center'> Defective Type </th> " +
                                                                       
                                 "</tr>";
                HeaderRewBcgRow2 = "<tr>" +
                                    $"<th align='center'> Rank 1 </th> " +
                                    $"<th align='center'> Rank 2 </th> " +
                                    $"<th align='center'> Rank 3 </th> " +
                                    $"<th align='center'> Others </th> " +
                                 "</tr>";

                TableRewBcgHeader = "<thead>" + HeaderRewBcgRow1 + HeaderRewBcgRow2 + "</thead>";
                

               
                string plant, plantPre ="", prod, osd, rate, colorRate, colorRow, styleCd, styleNm, reason1, reason2, reason3, reason4;
                int rowspanPlant = 1;
                TableRewBcgRow += "<tbody>";
                foreach (DataRow rowData in arg_DtData.Rows)
                {
                    plant = rowData["OP_NM"].ToString();

                    colorRate = rowData["COLOR_RATE"].ToString();
                    colorRow = rowData["COLOR_ROW"].ToString();
                    prod = rowData["PROD_QTY"].ToString();
                    osd = rowData["OSD_QTY"].ToString();
                    rate = rowData["RATE"].ToString();
                    styleCd = rowData["STYLE_CD"].ToString();
                    styleNm = rowData["STYLE_NM"].ToString();
                    reason1 = rowData["REASON_NM1"].ToString();
                    reason2 = rowData["REASON_NM2"].ToString();
                    reason3 = rowData["REASON_NM3"].ToString();
                    reason4 = rowData["REASON_NM4"].ToString();


                    TableRewBcgRow += "<tr>";

                    if (plantPre == "" || plantPre != plant)
                    {
                        rowspanPlant = (int)arg_DtData.Compute("COUNT(OP_NM)", $"OP_NM ='{plant}'");
                        TableRewBcgRow += $"<td rowspan = '{rowspanPlant}' col class= 'white' style=' width: 80' align='left'>{plant}</td>";
                    }
                    TableRewBcgRow += $"<td col class= '{colorRow}' style=' width: 300' align='left'>{styleNm}</td>" +
                                      $"<td col class= '{colorRow}' style=' width: 80' align='right'>{prod}</td>" +
                                      $"<td col class= '{colorRow}' style=' width: 80' align='right'>{osd}</td>" +
                                      $"<td col class= '{colorRate}' style=' width: 80' align='right'>{rate}</td>" +
                                      $"<td col class= '{colorRow}' style=' width: 250' align='left'>{reason1}</td>" +
                                      $"<td col class= '{colorRow}' style=' width: 250' align='left'>{reason2}</td>" +
                                      $"<td col class= '{colorRow}' style=' width: 250' align='left'>{reason3}</td>" +
                                      $"<td col class= '{colorRow}' style=' width: 250' align='left'>{reason4}</td>"
                                      ;

                    TableRewBcgRow += "</tr>";
                    plantPre = plant;
                }
                
                                    
                TableRewBcgRow += "</tbody>";
                TableRewBcg = "<table class = 'tblBoder'>" + TableRewBcgHeader + TableRewBcgRow + "</table>";

                /*
                string rate1 = "", rate2 = "", rate3 = "", rate4 = "", rate5 = "", rate6 = "";
                string clr1 = "", clr2 = "", clr3 = "", clr4 = "", clr5 = "", clr6 = "";
                try
                {
                    DataTable dtRate = arg_DtData.Select("ORD_SORT = '10000'", "RN").CopyToDataTable();
                    rate1 = dtRate.Rows[0]["RATE"].ToString();
                    rate2 = dtRate.Rows[1]["RATE"].ToString();
                    rate3 = dtRate.Rows[2]["RATE"].ToString();
                    rate4 = dtRate.Rows[3]["RATE"].ToString();
                    rate5 = dtRate.Rows[4]["RATE"].ToString();
                    rate6 = dtRate.Rows[5]["RATE"].ToString();
                    clr1 = dtRate.Rows[0]["COLOR_RATE"].ToString();
                    clr2 = dtRate.Rows[1]["COLOR_RATE"].ToString();
                    clr3 = dtRate.Rows[2]["COLOR_RATE"].ToString();
                    clr4 = dtRate.Rows[3]["COLOR_RATE"].ToString();
                    clr5 = dtRate.Rows[4]["COLOR_RATE"].ToString();
                    clr6 = dtRate.Rows[5]["COLOR_RATE"].ToString();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }

                strExplain = string.Format(strExplain, rate1, rate2, rate3, rate4, rate5, rate6
                                                     , clr1, clr2, clr3, clr4, clr5, clr6);
                */

                
                try
                {
                    string[] arrValue ;
                    string ord = arg_DtData.Rows[0]["COL_ORDER"].ToString();
                    DataTable dtRate = arg_DtData.Select("ORD_SORT = '10000'", ord).CopyToDataTable();
                    int iMaxRow = dtRate.Rows.Count;
                    if (iMaxRow < 6) iMaxRow = 6;
                    arrValue = new string[iMaxRow *6];
                    int i = 0;
                    string[] arrValue2 = new string[iMaxRow * 2];
                    for (int iArr = 0; iArr < arrValue2.Length; iArr++)
                    {
                        // arrValue2[iArr] = " ";
                        if (iArr % 2 == 0)
                            arrValue2[iArr] = "0";
                        else
                            arrValue2[iArr] = "trans";
                    }
                    foreach (DataRow row in dtRate.Rows)
                    {
                        switch (row["RN"].ToString())
                        {
                            case "1":
                                arrValue2[0] = row["RATE"].ToString();
                                arrValue2[1] = row["COLOR_RATE"].ToString();
                                break;
                            case "2":
                                arrValue2[2] = row["RATE"].ToString();
                                arrValue2[3] = row["COLOR_RATE"].ToString();
                                break;
                            case "3":
                                arrValue2[4] = row["RATE"].ToString();
                                arrValue2[5] = row["COLOR_RATE"].ToString();
                                break;
                            case "4":
                                arrValue2[6] = row["RATE"].ToString();
                                arrValue2[7] = row["COLOR_RATE"].ToString();
                                break;
                            case "5":
                                arrValue2[8] = row["RATE"].ToString();
                                arrValue2[9] = row["COLOR_RATE"].ToString();
                                break;
                            case "6":
                                arrValue2[10] = row["RATE"].ToString();
                                arrValue2[11] = row["COLOR_RATE"].ToString();
                                break;
                            default:
                                break;

                        }
                        //arrValue[i] = row["OP_NM2"].ToString();
                        //i++;
                        //arrValue[i] = row["GREEN"].ToString();
                        //i++;
                        //arrValue[i] = row["YELLOW"].ToString();
                        //i++;
                        //arrValue[i] = row["RED"].ToString();
                        //i++;
                        //arrValue[i] = row["RATE"].ToString();
                        //i++;
                        //arrValue[i] = row["COLOR_RATE"].ToString();
                        //i++;
                    }

                    strExplain = string.Format(strExplain, arrValue2);
                }
                
                catch (Exception ex)
                {
                    Debug.WriteLine(ex); 
                }
                
                return "<Html>" +
                            strStyle +
                            "<body>" +
                                strExplain +
                                "<table>" +
                                 "<tr> " +
                                      strTitle1 +  
                                 "</tr>" +
                                    
                                 "<tr> " +
                                    "<td>" + TableRewBcg + " </td>" +
                                 "</tr>" +

                                "</table>" +
                               // strTitle2 + 
                               // TableExtOsd +
                            "</body>" +
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
           // MyOraDB.ConnectName = COM.OraDB.ConnectDB.SEPHIROTH;
            MyOraDB.ShowErr = true;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_BOTTOM_DEFECTIVE";
                MyOraDB.ReDim_Parameter(6);
                MyOraDB.Process_Name = process_name;

                MyOraDB.Parameter_Name[0] = "V_P_TYPE";
                MyOraDB.Parameter_Name[1] = "CV_DATA";
                MyOraDB.Parameter_Name[2] = "CV_DATA2";
                MyOraDB.Parameter_Name[3] = "CV_COL";
                MyOraDB.Parameter_Name[4] = "CV_SUBJECT";
                MyOraDB.Parameter_Name[5] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[4] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[5] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = V_P_TYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";
                MyOraDB.Parameter_Values[3] = "";
                MyOraDB.Parameter_Values[4] = "";
                MyOraDB.Parameter_Values[5] = "";

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
