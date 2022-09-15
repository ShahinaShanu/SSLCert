using System;
using System.Net;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Net.Security;
using System.Configuration;

namespace SSL
{
    public partial class _Default : Page
    {
        public static string ExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
        public static int count = 0;
        protected void Page_Load(object sender, EventArgs e)
        {
            ServicePointManager.ServerCertificateValidationCallback = new
            RemoteCertificateValidationCallback
            (
                 delegate { return true; }
            );
            DomainInfo info = new DomainInfo();
            info.DaysLeftUntilExpiration = 0;
            info.error = "";
            HttpWebResponse response = null;
            //The following line enables TLS 1.1 (in case other requests need to support TLS 1.1) and also TLS 1.2 for Quik! 
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
            string connectionString = "";
            string path = Server.MapPath(ExcelPath);
            try
            {
                if (path.Contains(".xlsx"))
                {
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                }
                else if (path.Contains(".xls"))
                {
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                }
                string query = "SELECT [Sites Domain] FROM [Sheet1$]";
                //Providing connection
                OleDbConnection conn = new OleDbConnection(connectionString);
                //checking that connection state is closed or not if closed the 
                //open the connection
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                //create command object
                OleDbCommand cmd = new OleDbCommand(query, conn);
                // create a data adapter and get the data into dataadapter
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                //fill the excel data to data set
                da.Fill(ds);
                ds.Tables[0].Columns.Add("Days Left", typeof(System.Int32));
                ds.Tables[0].Columns.Add("Created Date", typeof(System.String));
                ds.Tables[0].Columns.Add("Expiry Date", typeof(System.String));
                ds.Tables[0].Columns.Add("IP Address", typeof(System.String));
                ds.Tables[0].Columns.Add("Status", typeof(System.String));
                if (ds.Tables != null && ds.Tables.Count > 0)
                {
                    count= ds.Tables[0].Rows.Count;
                     for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            if (ds.Tables[0].Columns[0].ToString() == "Sites Domain")
                            {
                                //HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://"+Domainlines[i]);
                                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://" + ds.Tables[0].Rows[i].Field<string>(0).ToString());
                                //dataSet.Tables["MyTable"].Rows[index]["MyColumn"]
                                response = (HttpWebResponse)request.GetResponse();
                                response.Close();
                                if (response.StatusCode == HttpStatusCode.OK)
                                {
                                //IPAddress[] addresslist = Dns.GetHostAddresses("https://" + ds.Tables[0].Rows[i].Field<string>(0).ToString());
                                 //retrieve the ssl cert and assign it to an X509Certificate object
                                X509Certificate cert = request.ServicePoint.Certificate;
                                    if (cert != null)
                                    {
                                        //convert the X509Certificate to an X509Certificate2 object by passing it into the constructor
                                        X509Certificate2 cert2 = new X509Certificate2(cert);
                                        // string cn = cert2.GetIssuerName();
                                        // string cedate = cert2.GetExpirationDateString();
                                        // string cpub = cert2.GetPublicKeyString();
                                        info.DaysLeftUntilExpiration = (Convert.ToDateTime(cert2.GetExpirationDateString()) - DateTime.Now).Days;
                                        info.SSLCreatedDate = cert2.GetEffectiveDateString();
                                        info.SSLExpirationDate = cert2.GetExpirationDateString();
                                        ds.Tables[0].Rows[i]["Days Left"] = info.DaysLeftUntilExpiration;
                                        ds.Tables[0].Rows[i]["Created Date"] = info.SSLCreatedDate;
                                        ds.Tables[0].Rows[i]["Expiry Date"] = info.SSLExpirationDate;
                                        ds.Tables[0].Rows[i]["IP Address"] = Dns.GetHostEntry(ds.Tables[0].Rows[i].Field<string>(0).ToString()).AddressList.First(
                                                                            addr => addr.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);

                                    if (info.DaysLeftUntilExpiration> 40)
                                    {
                                        ds.Tables[0].Rows[i]["Status"] = "Active";
                                    }
                                    else
                                    {
                                        ds.Tables[0].Rows[i]["Status"] = "Not Active";
                                    }
                                   //gvExcelFile.Rows[i].BackColor = System.Drawing.Color.Red;
                                    ds.Tables[0].AcceptChanges();
                                    }
                                    else
                                    {
                                    ds.Tables[0].Rows[i]["Status"] = "Failed";
                                    //Console.WriteLine(string.Format("{0}: {1}", response.StatusCode.ToString(), "Error"));
                                    }

                                }
                                else
                                {
                                    Console.WriteLine(string.Format("{0}: {1}", response.StatusCode.ToString(), "Error"));
                                }

                            }
                            
                        }
                    }
                

                gvExcelFile.DataSource = ds.Tables[0];
                //binding the gridview
                gvExcelFile.DataBind();
                //close the connection
                conn.Close();

}
         

            catch (WebException ex)
            {
                info.error = ex.Message.ToString();
            }

            catch (FileNotFoundException ex)
            {
                info.error = ex.Message.ToString();
            }
            catch (Exception ex)
            {
                info.error = ex.Message.ToString();
            }
            finally
            {
                if (response != null)
                    response.Close();
            }
            
        }
        protected void gvExcelFile_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            
            //if (e.Row.RowType == DataControlRowType.DataRow)
            //{
            //for (int i = 0; i < count; i++) { 
            //    if (e.Row.Cells[2].Text.Contains(""))
            //{
            //    //e.Row.Font.Bold = false;
            //    //e.Row.BackColor = System.Drawing.Color.Red;
            //    e.Row.Cells[2].BackColor = System.Drawing.Color.Red;
            //}
            //}
            //}
        }
    }
   
    internal class DomainInfo
    {
        internal string SSLCreatedDate;

        public int DaysLeftUntilExpiration { get; internal set; }
        public string error { get; internal set; }
        public string SSLExpirationDate { get; internal set; }
    }
   
}

