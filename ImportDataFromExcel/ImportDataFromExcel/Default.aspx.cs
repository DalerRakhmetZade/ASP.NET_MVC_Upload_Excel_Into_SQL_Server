using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ImportDataFromExcel
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                populateDatabaseData();
                
            }
        }
        private void populateDatabaseData()
        {
            using (ClientDataEntities1 dc = new ClientDataEntities1())
            {
                gvData.DataSource = dc.ClientMasters.ToList();
                gvData.DataBind();
            }
        }
        protected void btnImportFromCSV_Click(object sender, EventArgs e)
        {
            if (FileUpload1.PostedFile.ContentType == "application/vnd.ms-excel" ||
                FileUpload1.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {

                try
                {
                    string fileName = System.IO.Path.Combine(Server.MapPath("~/UploadDocuments"), Guid.NewGuid().ToString() + Path.GetExtension(FileUpload1.PostedFile.FileName));
                    FileUpload1.PostedFile.SaveAs(fileName);

                    string conString = "";
                    string ext = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    if (ext.ToLower() == ".xls")
                    {
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    else if (ext.ToLower() == ".xlsx")
                    {
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }

                    string query = "Select [ClientID], [First_Name], [Last_Name], [Employer], [Title], [Phone_Number], [Zip] from [ClientList$]";

                    OleDbConnection con = new OleDbConnection(conString);
                    if (con.State == System.Data.ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    OleDbCommand cmd = new OleDbCommand(query, con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    da.Dispose();
                    con.Close();
                    con.Dispose();

                    // Update database data
                    using (ClientDataEntities1 dc = new ClientDataEntities1())
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            string empID = dr["ClientID"].ToString();
                            var v = dc.ClientMasters.Where(a => a.ClientID.Equals(empID)).FirstOrDefault();
                            if (v != null)
                            {
                                v.First_Name = dr["First_Name"].ToString();
                                v.Last_Name = dr["Last_Name"].ToString();
                                v.Employer = dr["Employer"].ToString();
                                v.Title = dr["Title"].ToString();
                                v.Phone_Number = dr["Phone_Number"].ToString();
                                v.Zip = dr["Zip"].ToString();

                            }
                            else
                            {
                                dc.ClientMasters.Add(new ClientMaster
                                {
                                    ClientID = dr["ClientID"].ToString(),
                                    First_Name = dr["First_Name"].ToString(),
                                    Last_Name = dr["Last_Name"].ToString(),
                                    Employer = dr["Employer"].ToString(),
                                    Title = dr["Title"].ToString(),
                                    Phone_Number = dr["Phone_Number"].ToString(),
                                    Zip = dr["Zip"].ToString(),
                                });
                            }
                        }

                        dc.SaveChanges();

                        // populate updated data 
                        populateDatabaseData();
                        lblMessage.Text = "Excel file has been imported Successfully!";
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        protected void gvData_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}