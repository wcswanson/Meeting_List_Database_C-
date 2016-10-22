using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;

namespace waDEIGcsharp.ListUpdates
{
    public partial class ListUpdate : System.Web.UI.Page
    {
        string dash = "-";
        string zero = "0";
        string comma = ", ";
        // enum day of week and remove it from db.

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
              
               ddlDOW.SelectedIndex = 1;

            }
        }

        protected void lbGet_Click(object sender, EventArgs e)
        {
            
            StringBuilder sbGet = new StringBuilder();
            
            string strDow = ddlDOW.SelectedValue.ToString();
            string strTime = ddlTime.SelectedValue.ToString();
            string strTown = ddlTown.SelectedValue.ToString();

            // get day selection
            if (strDow != zero)
            {
                sbGet.Append(string.Format(" DOW = {0}", strDow));
            }
            else 
            {
                sbGet.Append(String.Format(" DOW > {0}", 0));
            }

            // Get time selection
            if (strTime != zero)
            {
                if (sbGet.Length > 0)
                {
                    sbGet.Append(String.Format(" AND TimeID = {0}", strTime));
                }
            }

            // Get Town selection
            if (strTown != dash)
            {
                if (sbGet.Length > 0)
                {
                    sbGet.Append(String.Format(" AND Town = '{0}'", strTown));
                }                
            }
            else
            {
                sbGet.Append(String.Format(" AND Town like '{0}'", "%"));
            }

            // Set the filter
            if (sbGet.Length > 0)
            {
                SqlDsList.FilterExpression = sbGet.ToString();
            }
            else
            {
                SqlDsList.FilterExpression = null;
            }
            lblInfo.Text = "Filter: " + SqlDsList.FilterExpression.ToString();
 
        }

        protected void lnkInsert_Click(object sender, EventArgs e)
        {
            if (PnlAdd.Visible)
            {
                PnlAdd.Visible = false;
            }
            else
            {
                PnlAdd.Visible = true;   
            }
        }

        protected void lnkSave_Click(object sender, EventArgs e)
        {
            // Get the day
            int rc = -1;
            int dow = Int32.Parse(ddlDOWInsert.SelectedValue);

            // Get the time
            int timeId = Int32.Parse(ddlTimeInsert.SelectedValue);

            // Get town
            string strTown = txtTown.Text;

            // Get Group
            string strGroupName = tbGroupName.Text;

            // Get information
            string strInformation = tbInfo.Text;

            // Get Location
            string strLocation = tbLocation.Text;

            // Get type
            string strType = tbType.Text;

            // Get connection string
            string cnStr = ConfigurationManager.ConnectionStrings["cnDEIG"].ConnectionString;
            SqlConnection cn = new SqlConnection(cnStr);

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText= string.Format("INSERT INTO [List] ([DOW], [TimeID], [Town], [GroupName], [Information], [Location], [Type]) VALUES ({0}, {1}, '{2}', '{3}', '{4}', '{5}', '{6}')", dow, timeId, strTown, strGroupName, strInformation, strLocation, strType);
            cmd.CommandType = CommandType.Text;
            cn.Open();

            try
            {
                rc = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                lblInfo.Text = ex.ToString();
                throw;
            }

            cn.Close();
        }
        
        protected void LnkExport_Click(object sender, EventArgs e)
        {
            int rc = -1;
            string cnStr = ConfigurationManager.ConnectionStrings["cnDEIG"].ConnectionString;
            SqlConnection cn = new SqlConnection(cnStr);
            SqlCommand cmd = new SqlCommand("GetMeetingList");
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter da = new SqlDataAdapter();

            // Get the connection into the SqlCommand object
            cmd.Connection = cn;

            da.SelectCommand = cmd;

            DataTable dt = new DataTable();

            try
            {
                rc = da.Fill(dt);
            }
            catch (InvalidCastException ex)
            {
                lblInfo.Text = ex.ToString();
            }

            // Read the table to create a file for export
            StringBuilder sb = new StringBuilder();

            foreach (DataColumn column in dt.Columns)
            {
                // Add the Header row for the text file
                sb.Append(column.ColumnName + comma);
            }
            // Add new line
            sb.AppendLine();

            // Get the data
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    sb.Append(row[column.ColumnName].ToString() + comma);
                }

                // Add new line
                sb.AppendLine();
            }

            // Download the text file
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=MeetingList.csv");
            Response.Charset = "";
            Response.ContentType = "application/text";
            Response.Output.Write(sb.ToString());
            Response.Flush();
            Response.End();

        }       
     }
}