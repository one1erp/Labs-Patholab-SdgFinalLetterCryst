using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using LSEXT;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using System.Diagnostics;
using Oracle.ManagedDataAccess.Client;
using Patholab_Common;

//if sdg is authorize ,then print from u_pdf_path'
// if has revision check with ziv
//

namespace SdgFinalLetterCryst
{

    [ComVisible(true)]
    [ProgId("SdgFinalLetterCryst.SdgFinalLetterCrystCls")]
    public class SdgFinalLetterCrystCls : IWorkflowExtension
    {


        #region private members
        INautilusServiceProvider _sp;
        SdgInfo _sdgInfo = null;
        OracleConnection oraCon = null;
        OracleCommand cmd = null;
        ReportDocument cr = null;
        //   private NautilusDBConnection _ntlsUser;
        private string SdgID;
        private long wnid;
        private string wnName;
        #endregion

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {

                Patholab_Common.Logger.WriteLogFile("Starting Print Pdf Letter Event " + "SdgFinalLetterCryst program");

                _sp = Parameters["SERVICE_PROVIDER"];
                string role = Parameters["ROLE_NAME"];


                bool debug = (role.ToUpper() == "DEBUG");
                if (debug) Debugger.Launch();
                string tableName = Parameters["TABLE_NAME"];

                wnid = Parameters["WORKFLOW_NODE_ID"];

                INautilusDBConnection _ntlsCon = null;

                if (_sp != null)
                {
                    _ntlsCon = _sp.QueryServiceProvider("DBConnection") as NautilusDBConnection;
                }

                if (_ntlsCon != null)
                {
                    // _username= dbConnection.GetUsername();
                    oraCon = GetConnection(_ntlsCon);

                    //set oracleCommand's connection
                    cmd = oraCon.CreateCommand();
                }

                var records = Parameters["RECORDS"];
                records.MoveLast();

                var recordId = records.Fields[tableName + "_ID"].Value;

                SdgID = recordId.ToString();// --Authorised
                string sqlwnName = "select pn.name from lims_sys.WORKFLOW_NODE cn,lims_sys.WORKFLOW_NODE pn " +
                             "where cn.WORKFLOW_NODE_ID='" + wnid + "' and pn.WORKFLOW_NODE_ID=cn.PARENT_ID";
                cmd.CommandText = sqlwnName;
                var res = cmd.ExecuteScalar();
                wnName = res.ToString();

                _sdgInfo = GetSdgInfo();

                if (!string.IsNullOrEmpty(_sdgInfo.PdfPath))
                {
                    Logger.WriteLogFile("Print " + _sdgInfo.PdfPath);
                    SendToPrinter(_sdgInfo.PdfPath);
                    return;
                }

                string sql =
                    string.Format(
                        "select parameter_2 from lims_sys.workflow_node where parent_id={0} and NAME='Comment' AND long_name='Path'",
                        wnid);


                cmd = new OracleCommand(sql, oraCon);
                var path = cmd.ExecuteScalar();
                string rptPath = path.ToString();


                sql =
              string.Format(
                  "select parameter_2 from lims_sys.workflow_node where parent_id={0} and NAME='Comment' AND long_name='Param'",
                  wnid);
                cmd.CommandText = sql;
                string param = (string)cmd.ExecuteScalar();

                //LDAP TODO
                string serverName;
                string nautilusUserName;
                string nautilusPassword;

                serverName = _ntlsCon.GetServerDetails();
                var IsProxy = _ntlsCon.GetServerIsProxy();
                if (IsProxy)
                {
                    nautilusUserName = "";
                    nautilusPassword = "";
                }
                else
                {
                    nautilusUserName = _ntlsCon.GetUsername();
                    nautilusPassword = _ntlsCon.GetPassword();
                }

                cr = new ReportDocument();
                TableLogOnInfo crTableLoginInfo;
                var crConnectionInfo = new ConnectionInfo();
                crConnectionInfo.ServerName = serverName;
                string p2 = path.ToString();
                cr.Load(p2);
                if (IsProxy)
                {
                    crConnectionInfo.IntegratedSecurity = true;
                }
                else
                {
                    crConnectionInfo.UserID = nautilusUserName;
                    crConnectionInfo.Password = nautilusPassword;
                }

                cr.SetParameterValue(param, recordId.ToString());
                var CrTables = cr.Database.Tables;
                foreach (Table CrTable in CrTables)
                {
                    crTableLoginInfo = CrTable.LogOnInfo;
                    crTableLoginInfo.ConnectionInfo = crConnectionInfo;
                    CrTable.ApplyLogOnInfo(crTableLoginInfo);
                }




                if (wnName == "Authorised" || (string.IsNullOrEmpty(_sdgInfo.PdfPath) && _sdgInfo.Status == "A"))//if event is raised by authorized evnt just save
                {
                    Patholab_Common.Logger.WriteLogFile("Save Pdf " + _sdgInfo.PatholabNbr);

                    SavePdf();
                }
                else
                {
                    Patholab_Common.Logger.WriteLogFile("Print by crystal");

                    cr.PrintToPrinter(1, true, 0, 0);
                }

                // ashi 29.8.18 for leaking memory


            }
            catch (Exception e)
            {
                //    MessageBox.Show("Err At Final Letter report :  " + e.Message, "Nautilus - Final Letter");
                Patholab_Common.Logger.WriteLogFile("Err At Final Letter report :  " + e.Message + " Nautilus - Final Letter");

            }
            finally
            {
                Patholab_Common.Logger.WriteLogFile("Ending Print Pdf Letter Event " + "SdgFinalLetterCryst program");

                if (cmd != null) cmd.Dispose();
                if (oraCon != null) oraCon.Close();


                if (cr != null)
                {

                    cr.Close();
                    cr.Dispose();
                }
                // ashi 15.8.18 for leaking memory
                GC.Collect();
                GC.SuppressFinalize(this);

            }
        }
        private SdgInfo GetSdgInfo()
        {
            string sdgSql = "SELECT DU.U_PDF_PATH,DU.U_PATHOLAB_NUMBER,D.STATUS FROM LIMS_SYS.SDG D,LIMS_SYS.SDG_USER DU " +
                            " WHERE D.SDG_ID=DU.SDG_ID AND D.SDG_ID='" + SdgID + "'";
            cmd.CommandText = sdgSql;

            var reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                _sdgInfo = new SdgInfo()
                {
                    PatholabNbr = reader["U_PATHOLAB_NUMBER"].ToString(),
                    Status = reader["Status"].ToString(),
                    PdfPath = reader["U_PDF_PATH"].ToString()
                };
                return _sdgInfo;
            }
            else
            {
                Patholab_Common.Logger.WriteLogFile("Error on get sdg details");
                //MessageBox.Show("Error on get sdg details");
                return null;
            }

        }
        private const string phraseHeaderName = "System Parameters";
        private const string PDF_Directory_Test = "PDF Directory2";
        private void SavePdf()
        {

            // cmd.Dispose();

            string sqlPhrase = "select phrase_description from lims_sys.phrase_entry where phrase_id in" +
                               "(select phrase_id from lims_sys.phrase_header  where name ='" + phraseHeaderName +
                               "') and phrase_name='" + PDF_Directory_Test + "'";
            cmd.CommandText = sqlPhrase;
            var path2Save = cmd.ExecuteScalar();

            if (path2Save != null)
            {
                //  Logger.WriteLogFile("Export To Disk by crystal");
                Patholab_Common.Logger.WriteLogFile(path2Save + SdgID + "_" + wnid.ToString() + ".pdf  saved by SdgFinalLetterCryst.dll");

                cr.ExportToDisk(ExportFormatType.PortableDocFormat, path2Save + SdgID + "_" + wnid.ToString() + ".pdf");

            }
            else
            {
                Logger.WriteLogFile("Error on get pdf path");
            }
        }
        private void SendToPrinter(string pdfPath)
        {
            Process pp = new Process();
            pp.StartInfo = new ProcessStartInfo()
            {
                CreateNoWindow = true,
                Verb = "print",
                FileName = pdfPath
            };
            pp.Start();
        }
        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {
            OracleConnection connection = null;
            if (ntlsCon != null)
            {
                //initialize variables
                string rolecommand;
                //try catch block
                try
                {

                    string connectionString;
                    string server = ntlsCon.GetServerDetails();
                    string user = ntlsCon.GetUsername();
                    string password = ntlsCon.GetPassword();

                    connectionString =
                        string.Format("Data Source={0};User ID={1};Password={2};", server, user, password);
                    var username = ntlsCon.GetUsername();
                    if (string.IsNullOrEmpty(username))
                    {
                        var serverDetails = ntlsCon.GetServerDetails();
                        connectionString = "User Id=/;Data Source=" + serverDetails + ";";
                    }
                    Logger.WriteLogFile(connectionString);
                    //create connection
                    connection = new OracleConnection(connectionString);

                    //open the connection
                    connection.Open();

                    //get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    //set role lims user
                    if (limsUserPassword == "")
                    {
                        //lims_user is not password protected 
                        rolecommand = "set role lims_user";
                    }
                    else
                    {
                        //lims_user is password protected
                        rolecommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    //set the oracle user for this connection
                    OracleCommand command = new OracleCommand(rolecommand, connection);

                    //try/catch block
                    try
                    {
                        //execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        //throw the exeption
                        MessageBox.Show("Inconsistent role Security : " + f.Message);
                    }

                    //get session id
                    double sessionId = ntlsCon.GetSessionId();

                    //connect to the same session 
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    //Build the command 
                    command = new OracleCommand(sSql, connection);

                    //execute the command
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    //throw the exeption
                    MessageBox.Show("Err At GetConnection: " + e.Message);
                }
            }
            return connection;
        }

    }

}