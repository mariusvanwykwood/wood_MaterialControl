using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Wood_MaterialControl
{
    public partial class DownloadFileHelper : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                int fileID;
                if (int.TryParse(Request.QueryString["fileID"], out fileID))
                {
                    DownloadFileData(fileID);
                }
            }
        }

        private void DownloadFileData(int fileID)
        {
            string connStr = DataClass.conMat;

            using (SqlConnection conn = new SqlConnection(connStr))
            {
                string query = "SELECT FinalFileName, FileData FROM SPMAT_FileExports WHERE FinalFileID = @FileID";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@FileID", fileID);

                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    string fileName = reader["FinalFileName"].ToString();
                    byte[] fileData = (byte[])reader["FileData"];

                    Response.Clear();
                    Response.ContentType = "application/vnd.ms-excel";
                    Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
                    Response.BinaryWrite(fileData);
                    Response.Flush();
                    Response.End();
                }
            }
        }

    }
}