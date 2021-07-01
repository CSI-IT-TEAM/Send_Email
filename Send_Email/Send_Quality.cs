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
    class Send_Quality
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
               // string TableExtOsd = "", TableExtOsdHeader = "", HeaderExtOsdRow1 = "", HeaderExtOsdRow2 = "", TableExtOsdRow = "";
                string TableRewBcg = "", TableRewBcgHeader = "", HeaderRewBcgRow1 = "", TableRewBcgRow = "";
                string bColor, fColor, bColorOsd, fColorOsd, bColorRew, fColorRew, bColorBcg, fColorBcg;
                string factory, plant, line, production,  rework, reworkRate, bQty, cQty, bcQty, bcRate, qty;
                string strStyle = argStyle.Rows[0]["STYLE"].ToString();
                string strTitle1 = argStyle.Rows[0]["TITLE1"].ToString();
                string strTitle2 = argStyle.Rows[0]["TITLE2"].ToString();
                string strExplain = argStyle.Rows[0]["TXT"].ToString();
                ///
                ///Rework & BC Grade
                ///
                HeaderRewBcgRow1 = "<tr>" +
                                     $"<th align='center'> Factory </th>" +
                                     $"<th align='center'> Plant </th>" +
                                     $"<th align='center'> Line </th> " +
                                     $"<th align='center'> Production Quantity </th> " +
                                     $"<th align='center' class='bcGrade' > B - Grade </th> " +
                                     $"<th align='center' class='bcGrade'> C - Grade </th> " +
                                     $"<th align='center' class='bcGrade'> Total B,C Grade </th> " +
                                     $"<th align='center' class='bcGrade'> PPM </th> " +
                                     $"<th align='center' class='rework'> Rework </th> " +
                                     $"<th align='center' class='rework'> Rate(%) </th> " +
                                 "</tr>";

                TableRewBcgHeader = "<thead>" + HeaderRewBcgRow1 + "</thead>";


                TableRewBcgRow = "<tbody>";

                
                int rowspanFactory = 1, rowspanPlant = 1;
                string factoryPre = "", plantPre = ""; 
                foreach (DataRow rowData in arg_DtData.Rows)
                {

                    bColor = rowData["BCOLOR"].ToString();
                    fColor = rowData["FCOLOR"].ToString();
                    bColorRew = rowData["BCOLOR_RE"].ToString();
                    fColorRew = rowData["FCOLOR_RE"].ToString();
                    bColorBcg = rowData["BCOLOR_BC"].ToString();
                    fColorBcg = rowData["FCOLOR_BC"].ToString();
                    
                    factory = rowData["FACTORY"].ToString();
                    plant = rowData["LINE_NM"].ToString();
                    line = rowData["MLINE_CD"].ToString();
                    production = rowData["PROD_QTY"].ToString();
                    bQty = rowData["B_QTY"].ToString();
                    cQty = rowData["C_QTY"].ToString();
                    bcQty = rowData["BC_QTY"].ToString();
                    bcRate = rowData["BC_RATE"].ToString();
                    rework = rowData["REWORK_QTY"].ToString();
                    reworkRate = rowData["REWORK_RATE"].ToString();

                    if (bcRate == "")
                    {
                        bColorBcg = bColor;
                        fColorBcg = fColor;
                    }

                    if (reworkRate == "")
                    {
                        bColorRew = bColor;
                        fColorRew = fColor;
                    }


                    TableRewBcgRow += "<tr> ";
                    if (factoryPre == "" || factory != factoryPre)
                    {
                        rowspanFactory = (int)arg_DtData.Compute("COUNT(FACTORY)", $"FACTORY ='{factory}'");
                        TableRewBcgRow +=  $"  <td rowspan = '{rowspanFactory}' bgcolor='{bColor}' style='color:{fColor}; width: 100' align='left'>{factory}</td>";
                    }

                    if (plantPre == "" || plant != plantPre)
                    {
                        rowspanPlant = (int)arg_DtData.Compute("COUNT(LINE_NM)", $"LINE_NM ='{plant}'");
                        TableRewBcgRow += $"  <td rowspan = '{rowspanPlant}' bgcolor='{bColor}' style='color:{fColor}; width: 100' align='left'>{plant}</td>";
                    }


                    TableRewBcgRow += $"  <td bgcolor='{bColor}'    style='color:{fColor};    width: 100' align='left'>{line}</td>" +
                                        $"<td bgcolor='{bColor}'    style='color:{fColor};    width: 100' align='right'>{production}</td>" +
                                        $"<td bgcolor='{bColor}'    style='color:{fColor};    width: 100' align='right'>{bQty}</td>" +
                                        $"<td bgcolor='{bColor}'    style='color:{fColor};    width: 100' align='right'>{cQty}</td>" +
                                        $"<td bgcolor='{bColor}'    style='color:{fColor};    width: 100' align='right'>{bcQty}</td>" +
                                        $"<td bgcolor='{bColorBcg}' style='color:{fColorBcg}; width: 100' align='right'>{bcRate}</td>" +
                                        $"<td bgcolor='{bColor}'    style='color:{fColor};    width: 100' align='right'>{rework}</td>" +
                                        $"<td bgcolor='{bColorRew}' style='color:{fColorRew}; width: 100' align='right'>{reworkRate}</td>";
                    TableRewBcgRow += "</tr> ";

                    factoryPre = factory;
                    plantPre = plant;
                }
                TableRewBcgRow += "</tbody>";
                TableRewBcg = "<table class = 'tblBoder'>" + TableRewBcgHeader + TableRewBcgRow + "</table>";


                ///
                ///OS&D External
                ///
                /*
                string[] arrHead = { "Outsole", "Phylon", "IP-A", "IP-B", "PU", "F#1", "Plant N", "Plant C", "Plant LE" };
                HeaderExtOsdRow1 = "<tr>" +
                                    $"<th  rowspan = '2' align='center'></th>" +
                                    $"<th class='date' colspan = '2' align='center'>Bottom#1</th>" +
                                    $"<th class='date' colspan = '3' align='center'>Bottom#2</th>" +
                                    $"<th class='date' colspan = '2' align='center'>Long Thanh</th>" +
                                    $"<th class='date' colspan = '3' align='center'>VJ3</th>" +
                                    $"<th  rowspan = '2' align='center'>Total</th>" +
                                 "</tr>";
                HeaderExtOsdRow2 = "<tr>" +
                                     $"<th align='center'> Outsole </th>" +
                                     $"<th align='center'> Phylon </th>" +
                                     $"<th align='center'> IP-A </th> " +
                                     $"<th align='center'> IP-B </th> " +
                                     $"<th align='center'> PU </th> " +
                                     $"<th align='center'> F#1 </th> " +
                                     $"<th align='center'> Plant N </th> " +
                                     $"<th align='center'> Plant C </th> " +
                                     $"<th align='center'> Plant N </th> " +
                                     $"<th align='center'> Plant LE </th> " +
                                 "</tr>";

                TableExtOsdHeader = "<thead>" + HeaderExtOsdRow1 + HeaderExtOsdRow2 + "</thead>";
                TableExtOsdRow = "<tbody>";

                string divPre = "", divCur ="", type = "";
                int iColMax = 11, iCol =0;

                foreach (DataRow rowData in arg_DtData2.Rows)
                {
                    divCur = rowData["DIV"].ToString();
                    type = rowData["O_TYPE"].ToString();
                    bColor = rowData["BCOLOR"].ToString();
                    fColor = rowData["FCOLOR"].ToString();
                    bColorOsd = rowData["BCOLOR_RATE"].ToString();
                    fColorOsd = rowData["FCOLOR_RATE"].ToString();
                    qty = rowData["QTY"].ToString();
                    
                    if (divPre == "")
                    {
                        TableExtOsdRow += "<tr> ";
                        TableExtOsdRow += $"<td bgcolor='{bColor}' style='color:{fColor}; width: 150' align='left'>{type}</td>";                       
                    }
                    else if (divCur != divPre && divCur == "2")
                    {                        
                        TableExtOsdRow += "</tr> ";                     
                        TableExtOsdRow += "<tr> ";
                        iCol = 0;
                        TableExtOsdRow += $"<td bgcolor='{bColor}' style='color:{fColor}; width: 90' align='left'>{type}</td>";
                        iCol++;
                    }
                    else if (divCur != divPre && divCur == "3")
                    {
                        if (iCol < iColMax)
                        {
                            for (int i = iCol; i <= iColMax; i++)
                            {
                                TableExtOsdRow += $"<td bgcolor='WHITE' style='color:WHITE; width: 100' align='right'></td>";
                            }
                        }
                        TableExtOsdRow += "</tr> ";
                        TableExtOsdRow += $"<td bgcolor='{bColor}' style='color:{fColor}; width: 90' align='left'>{type}</td>";
                        
                        TableExtOsdRow += "<tr> ";
                        iCol = 1;
                    }
                    if (divCur == "3")
                        TableExtOsdRow += $"<td bgcolor='{bColorOsd}' style='color:{fColorOsd}; width: 100' align='right'>{qty}</td>";
                    else
                        TableExtOsdRow += $"<td bgcolor='{bColor}' style='color:{fColor}; width: 100' align='right'>{qty}</td>";
                    divPre = divCur;
                    iCol++;
                }
                TableExtOsdRow += "</tr> ";
                TableExtOsdRow += "</tbody>";
                TableExtOsd = "<table>" + TableExtOsdHeader + TableExtOsdRow + "</table>";
                */


                return "<Html>" +
                            strStyle +
                            "<body>" +
                                strExplain +
                                strTitle1 +
                                TableRewBcg +
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
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            MyOraDB.ShowErr = true;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_QUALITY";
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
