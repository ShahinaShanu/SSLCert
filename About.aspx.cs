using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SSL
{
    public partial class About : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
    }
}

//using System;
//using System.Net;
//using System.Reflection;
//using System.Collections.Generic;
//using System.Linq;
//using System.Web;
//using System.Web.UI;
//using System.Web.UI.WebControls;
//using System.Security.Cryptography.X509Certificates;
//using System.IO;
//using System.Data.OleDb;
//using System.Data;
//using System.Net.Security;

//namespace SSL
//{
//    public partial class _Default : Page
//    {
//        public static string path = @"D:\shahina\SSL\SSL\DomainList\";
//        public static string filename = "SiteRenewalList.xlsx";
//        protected void Page_Load(object sender, EventArgs e)
//        {
//            ServicePointManager.ServerCertificateValidationCallback = new
//            RemoteCertificateValidationCallback
//            (
//                 delegate { return true; }
//            );
//            DomainInfo info = new DomainInfo();
//            info.DaysLeftUntilExpiration = 0;
//            info.error = "";
//            HttpWebResponse response = null;
//            //The following line enables TLS 1.1 (in case other requests need to support TLS 1.1) and also TLS 1.2 for Quik! 
//            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
//            string connectionString = "";
//            //string strFileType = "Type";
//            //string path = @"C:\Users\UserName\Downloads\";
//            //string path = Server.MapPath(@"~/DomainList/");
//            //string filename = "SiteRenewalList.xlsx";
//            try
//            {
//                if (filename.Contains(".xlsx"))
//                {
//                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + filename + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
//                }
//                else if (filename.Contains(".xls"))
//                {
//                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + filename + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
//                }
//                string query = "SELECT [Sites Domain] FROM [Sheet1$]";
//                //Providing connection
//                OleDbConnection conn = new OleDbConnection(connectionString);
//                //checking that connection state is closed or not if closed the 
//                //open the connection
//                if (conn.State == ConnectionState.Closed)
//                {
//                    conn.Open();
//                }
//                //create command object
//                OleDbCommand cmd = new OleDbCommand(query, conn);
//                // create a data adapter and get the data into dataadapter
//                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
//                DataSet ds = new DataSet();
//                //fill the excel data to data set
//                da.Fill(ds);
//                ds.Tables[0].Columns.Add("Days", typeof(System.Int32));
//                ds.Tables[0].Columns.Add("CreateDate", typeof(System.String));
//                if (ds.Tables != null && ds.Tables.Count > 0)
//                {
//                    foreach (DataRow row in ds.Tables[0].Rows)
//                    {
//                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
//                        {
//                            //if (ds.Tables[0].Columns[1].ToString() == "Sites Domain" && ds.Tables[0].Columns[2].ToString() == "Status")
//                            if (ds.Tables[0].Columns[0].ToString() == "Sites Domain")
//                            {
//                                //HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://"+Domainlines[i]);
//                                //HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://" + ds.Tables[0].Columns[0].ToString());
//                                //HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://" + ds.Tables[0].Rows[1].ToString());
//                                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://" + ds.Tables[0].Rows[i].Field<string>(0).ToString());
//                                //dataSet.Tables["MyTable"].Rows[index]["MyColumn"]
//                                response = (HttpWebResponse)request.GetResponse();
//                                response.Close();
//                                if (response.StatusCode == HttpStatusCode.OK)
//                                {
//                                    //retrieve the ssl cert and assign it to an X509Certificate object
//                                    X509Certificate cert = request.ServicePoint.Certificate;
//                                    if (cert != null)
//                                    {
//                                        //convert the X509Certificate to an X509Certificate2 object by passing it into the constructor
//                                        X509Certificate2 cert2 = new X509Certificate2(cert);
//                                        // string cn = cert2.GetIssuerName();
//                                        // string cedate = cert2.GetExpirationDateString();
//                                        // string cpub = cert2.GetPublicKeyString();
//                                        info.DaysLeftUntilExpiration = (Convert.ToDateTime(cert2.GetExpirationDateString()) - DateTime.Now).Days;
//                                        info.SSLCreatedDate = cert2.GetEffectiveDateString();
//                                        info.SSLExpirationDate = cert2.GetExpirationDateString();

//                                        //for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
//                                        //{
//                                        //int[] days = new int[10];
//                                        //days[j] = info.DaysLeftUntilExpiration;
//                                        row["Days"] = info.DaysLeftUntilExpiration;
//                                        row["CreateDate"] = info.SSLCreatedDate;
//                                        row.AcceptChanges();
//                                        //}
//                                        //need to set value to NewColumn column
//                                        // or set it to some other value

//                                    }
//                                    else
//                                    {
//                                        Console.WriteLine(string.Format("{0}: {1}", response.StatusCode.ToString(), "Error"));
//                                    }

//                                }
//                                else
//                                {
//                                    Console.WriteLine(string.Format("{0}: {1}", response.StatusCode.ToString(), "Error"));
//                                }

//                            }

//                            //else if (ds.Tables[0].Rows[0][i].ToString().ToUpper() == "NAME")
//                            //{

//                            //}
//                            //else if (ds.Tables[0].Rows[0][i].ToString().ToUpper() == "EMAIL")
//                            //{

//                            //}
//                        }
//                    }
//                }


//                //set data source of the grid view
//                //BoundField Days = new BoundField();
//                //Days.DataField = info.DaysLeftUntilExpiration.ToString();
//                //Days.HeaderText = "Days";
//                //gvExcelFile.Columns.Add(Days);
//                gvExcelFile.DataSource = ds.Tables[0];
//                //binding the gridview
//                gvExcelFile.DataBind();
//                //close the connection
//                conn.Close();


//            }
//            //var path = Server.MapPath(@"~/DomainList/Domains.txt");
//            //var path = Server.MapPath(@"~/DomainList/SiteRenewalList.xlsx");

//            catch (WebException ex)
//            {
//                info.error = ex.Message.ToString();
//            }

//            catch (FileNotFoundException ex)
//            {
//                info.error = ex.Message.ToString();
//            }
//            catch (Exception ex)
//            {
//                info.error = ex.Message.ToString();
//            }
//            finally
//            {
//                if (response != null)
//                    response.Close();
//            }

//        }
//    }

//    internal class DomainInfo
//    {
//        internal string SSLCreatedDate;

//        public int DaysLeftUntilExpiration { get; internal set; }
//        public string error { get; internal set; }
//        public string SSLExpirationDate { get; internal set; }
//    }
//}
