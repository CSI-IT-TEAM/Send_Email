﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Send_Email
{
    class Send_Feedback
    {
        public string _subject = "";
        public DataTable _email;
        public static string _Hms, _Ymd, _User;
        public string Html(string argType)
        {
            try
            {
                string htmlReturn = "";
                //string msg = EncryptExtend.ToDecryptString("pFtsI2Z5oYLg1ir9wDtd1LX/ycQqKlIUJfbez9AMpklbTJkBE8p6SDm/rluGBZOlb7C5L5rEclw0bJg2fn/9xQ==");
                DataSet dsData = SEL_DATA(argType);
                if (argType.Equals("Q"))
                {
                    if (dsData.Tables[0] == null || dsData.Tables[0].Rows.Count == 0) return "";
                    //WriteLog("RunNPI: Start --> " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    DataTable dtData = dsData.Tables[0];
                   
                    foreach (DataRow dr in dtData.Rows)
                    {
                        dr["TITLE"] = EncryptExtend.DescryptString(dr["TITLE"].ToString());
                        dr["CONTENTS"] = EncryptExtend.DescryptString(dr["CONTENTS"].ToString());
                        _Hms = dr["HMS"].ToString();
                        _Ymd = dr["YMD"].ToString();
                        _User = dr["ADD_USER"].ToString();
                        saveImg(dr["IMG"]);
                    }
                    
                    DataTable dtHeader = dsData.Tables[0];
                    _email = dsData.Tables[1];

                    // WriteLog(dtHeader.Rows.Count.ToString() + " " + dtData.Rows.Count.ToString() + " " + dtEmail.Rows.Count.ToString());

                    htmlReturn = GetHtmlBody(dtHeader, dtData);

                    _subject = "Digital Twin Help Desk";
                }
                //_subject = "(Test Email) Outsole press machine drawback list";
                return htmlReturn;
            }
            catch// (Exception ex)
            {
                return null;
               // return "Error: " + ex.ToString();
            }
            
        }

        private void saveImg(object argImg)
        {
            try
            {
                string path = Application.StartupPath + @"\Capture\feedback.png";

                byte[] MyData = (byte[])argImg;
                int ArraySize = MyData.GetUpperBound(0) + 1;
                FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);

                fs.Write(MyData, 0, ArraySize);
                fs.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
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
                                  text-align: left;
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
                                  font-size: 24px;
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

                                        <div class='info'>Without pressing actual during 4 hours were mailed out to PIC promptly.</div></br>

                                        <div class='info'>&nbsp;Số lượng sản xuất không đạt Target trong 4 giờ sẽ gửi Email</div></br>
                                        -->

                                        <table class='OSPTable'>
                                        <thead>
                                        <tr>
                                        <th>User Request </th>
                                        <th>Title</th>
                                        <th>Content</th>
                                        </tr>
                                        </thead><tbody>");
                //Row
                string TableRow = "";
                foreach (DataRow row in dtData.Rows)
                {
                   
                    TableRow += $"<tr><td>{row["REG_USER"].ToString().Replace("\r", "<br>")}</td><td>{row["TITLE"].ToString().Replace("\r", "<br>")} </td><td>{row["CONTENTS"].ToString().Replace("\r", "<br>")}</td></tr>";
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

        private DataSet SEL_DATA(string ARG_QTYPE)
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_FEEDBACK";
                MyOraDB.ReDim_Parameter(3);
                MyOraDB.Process_Name = process_name;
                MyOraDB.Parameter_Name[0] = "ARG_QTYPE";
                MyOraDB.Parameter_Name[1] = "CV_1";
                MyOraDB.Parameter_Name[2] = "CV_EMAIL";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.Cursor;
                MyOraDB.Parameter_Type[2] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = ARG_QTYPE;
                MyOraDB.Parameter_Values[1] = "";
                MyOraDB.Parameter_Values[2] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();

                if (ds_ret == null)
                {
                    if (ARG_QTYPE == "Q")
                    {
                      //  WriteLog("P_SEND_EMAIL_NPI: null");
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

        public static DataTable UPD_DATA()
        {
            COM.OraDB MyOraDB = new COM.OraDB();
            MyOraDB.ConnectName = COM.OraDB.ConnectDB.LMES;
            DataSet ds_ret;
            try
            {
                string process_name = "P_SEND_EMAIL_FEEDBACK_UPDATE";
                MyOraDB.ReDim_Parameter(4);
                MyOraDB.Process_Name = process_name;
                MyOraDB.Parameter_Name[0] = "ARG_YMD";
                MyOraDB.Parameter_Name[1] = "ARG_HMS";
                MyOraDB.Parameter_Name[2] = "ARG_REG_USER";
                MyOraDB.Parameter_Name[3] = "CV_1";

                MyOraDB.Parameter_Type[0] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[1] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[2] = (int)OracleType.VarChar;
                MyOraDB.Parameter_Type[3] = (int)OracleType.Cursor;

                MyOraDB.Parameter_Values[0] = _Ymd;
                MyOraDB.Parameter_Values[1] = _Hms;
                MyOraDB.Parameter_Values[2] = _User;
                MyOraDB.Parameter_Values[3] = "";

                MyOraDB.Add_Select_Parameter(true);

                ds_ret = MyOraDB.Exe_Select_Procedure();
            
                return ds_ret.Tables[0];
            }
            catch (Exception ex)
            {
                // WriteLog("SEL_CUTTING_DATA: " + ex.ToString());
                return null;
            }
        }
    }
}
