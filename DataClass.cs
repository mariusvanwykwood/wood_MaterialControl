
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Reflection;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Configuration;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data.OleDb;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Office.Word;
using System.Diagnostics;
using static Wood_MaterialControl.DataClass;




namespace Wood_MaterialControl
{
    public class DataClass
    {
        public static string tmpcon = @"Data Source=EDC09-SQL01\DAILYDIARY;Initial Catalog=SQL_Engdata_DEV;User ID=SQL_EngData;Password=SQL_EngData;";
        public static string conUsers = @"Data Source=EDC09-SQL01\DAILYDIARY;Initial Catalog=WOOD_EMPLocations;User ID=LoactionLogger;Password=LocationLogger@2023;";

        #region SQL EngData
        #region Classes
        public class UserLookupData
        {
            public int UserID { get; set; }

            public string UserName { get; set; }
        }
        public class DDLList
        {
            public string DDLList_ID { get; set; }
            public string DDLListName { get; set; }
            public string DDLID { get; set; } = "";
        }
        public class ParsedDecimal
        {
            public decimal Value { get; set; } = 0m;
            public bool HasError { get; set; } = false;
        }
        #endregion
        #region Functions
        public static ParsedDecimal DecParse(string value)
        {
            ParsedDecimal parsed = new ParsedDecimal();
            if (string.IsNullOrWhiteSpace(value))
            {
                parsed.HasError = true;
                return parsed;
            }
            string cleanedValue = value.Replace(" ", "").Trim();
            string[] separators = { ",", "." };
            foreach (var sep in separators)
            {
                try
                {
                    var culture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
                    culture.NumberFormat.NumberDecimalSeparator = sep;
                    parsed.Value = decimal.Parse(cleanedValue, NumberStyles.Float, culture);
                    return parsed; // Success
                }
                catch
                {
                    // Try next separator
                }
            }
            // Parsing failed
            parsed.HasError = true;
            parsed.Value = 0m;
            return parsed;
        }
        public static string ReplaceFirst(string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
        public static System.Data.DataTable ToDataTable<T>(List<T> items)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
        #endregion
        #region DataCalls
        internal static List<DataClass.UserLookupData> GetAllUsersLookup()
        {
            List<DataClass.UserLookupData> allUsersLookup = new List<DataClass.UserLookupData>();
            SqlConnection connection = (SqlConnection)null;
            try
            {
                string cmdText = "SELECT DISTINCT [EID], [Surname] +' , '+[PreferredName]+'  ( '+[JobTitle] + ' - ' + [EmployeeEmail] +' )' FROM [WOOD_EMPLocations].[dbo].[Employees] where [IsDeleted]=0 and EID in(Select distinct [fld_EID] from [WOOD_MaterialControl_DEV].[dbo].[tbl_UserAccess]) order by 2";
                using (connection = new SqlConnection(DataClass.conUsers))
                {
                    SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
                    sqlCommand.CommandType = CommandType.Text;
                    connection.Open();
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                    while (sqlDataReader.Read())
                    {
                        DataClass.UserLookupData userLookupData = new DataClass.UserLookupData();
                        userLookupData.UserID = int.Parse(sqlDataReader[0].ToString());
                        userLookupData.UserName = sqlDataReader[1].ToString();
                        if (!allUsersLookup.Contains(userLookupData))
                            allUsersLookup.Add(userLookupData);
                    }
                    sqlDataReader.Close();
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                var x = ex.Message;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                GC.Collect();
            }
            return allUsersLookup;
        }

        internal static bool UserHasAccess(int uid)
        {
            bool hasAccess = false;
            SqlConnection connection = (SqlConnection)null;
            try
            {
                string cmdText = "SELECT TOP (1) [fld_HasAccess]  FROM [dbo].[tbl_UserAccess] where fld_EID = " + uid.ToString();
                using (connection = new SqlConnection(DataClass.conMat))
                {
                    SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
                    sqlCommand.CommandType = CommandType.Text;
                    connection.Open();
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                    while (sqlDataReader.Read())
                    {
                        hasAccess = bool.Parse(sqlDataReader[0].ToString());
                    }
                    sqlDataReader.Close();
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                var x = ex.Message;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                GC.Collect();
            }
            return hasAccess;
        }
        #endregion
        #endregion

        #region Material
        public static string conMat = @"Data Source=EDC09-SQL01\DAILYDIARY;Initial Catalog=WOOD_MaterialControl;User ID=MaterialControl;Password=MaterialControl;";
        #region Classes
        public class SpecData
        {
            public string Project { get; set; }
            public string Lineclass { get; set; }
            public string Shortcode { get; set; }
            public string Ident { get; set; }
            public string Commodity_code { get; set; }
            public string Short_desc { get; set; }
            public string Description { get; set; }
            public string Size_sch1 { get; set; }
            public string Size_sch2 { get; set; }
            public string Size_sch3 { get; set; }
            public string Size_sch4 { get; set; }
            public string Size_sch5 { get; set; }
            public string Spec_revision { get; set; }
            public string Published { get; set; }
            public string Published_date { get; set; }
            public string Option_code { get; set; }
            public string Option_code_desc { get; set; }
        }


        [Serializable]
        public class GridData
        {
            public int MaterialID { get; set; }
            public int ProjectID { get; set; }
            public string Discipline { get; set; }
            public string Area { get; set; }
            public string Unit { get; set; }
            public string Phase { get; set; }
            public string Const_Area { get; set; }
            public string ISO { get; set; }
            public string Component_Type { get; set; }
            public string Spec { get; set; }
            public string Shortcode { get; set; }
            public string Ident_no { get; set; }
            public string IsoShortDescription { get; set; }
            public string Size_sch1 { get; set; }
            public string Size_sch2 { get; set; }
            public string Size_sch3 { get; set; }
            public string Size_sch4 { get; set; }
            public string Size_sch5 { get; set; }
            public string qty { get; set; }
            public string qty_unit { get; set; }
            public string Fabrication_Type { get; set; }
            public string Source { get; set; } = "P";
            public string IsoRevisionDate { get; set; } = "";
            public string IsoRevision { get; set; } = "";
            public string IsLocked { get; set; } = "";
            public int IsoUniqeRevID { get; set; } = 0;

        }
        public class SPMATData
        {
            public int MTOID { get; set; }
            public string Discipline { get; set; }
            public string Area { get; set; }
            public string Unit { get; set; }
            public string Phase { get; set; }
            public string Const_Area { get; set; }
            public string ISO { get; set; }
            public string Ident_no { get; set; }
            public decimal qty { get; set; }
            public string qty_unit { get; set; }
            public string Fabrication_Type { get; set; }
            public string Spec { get; set; }
            public string Pos { get; set; }
            public string IsoRevisionDate { get; set; }
            public string IsoRevision { get; set; }
            public string IsLocked { get; set; }
            public string Code { get; set; }
            public string ImportStatus { get; set; }
            public int IsoUniqeRevID { get; set; }
        }
        public class SPMATIntrimData
        {
            public int INTID { get; set; }
            public int MaterialID { get; set; }
            public int ProjectID { get; set; }
            public string Discipline { get; set; }
            public string Area { get; set; }
            public string Unit { get; set; }
            public string Phase { get; set; }
            public string Const_Area { get; set; }
            public string ISO { get; set; }
            public string Ident_no { get; set; }
            public decimal qty { get; set; }
            public string qty_unit { get; set; }
            public string Fabrication_Type { get; set; }
            public string Spec { get; set; }
            public string Pos { get; set; }
            public string IsoRevisionDate { get; set; }
            public string IsoRevision { get; set; }
            public string IsLocked { get; set; }
            public string Code { get; set; }
            public int IsoUniqeRevID { get; set; }
        }
        public class SPMATDBData
        {
            public int MaterialID { get; set; }
            public int ProjectID { get; set; }
            public string Discipline { get; set; }
            public string Area { get; set; }
            public string Unit { get; set; }
            public string Phase { get; set; }
            public string Const_Area { get; set; }
            public string ISO { get; set; }
            public string Ident_no { get; set; }
            public string qty { get; set; }
            public string qty_unit { get; set; }
            public string Fabrication_Type { get; set; }
            public string Spec { get; set; }
            public string IsoRevisionDate { get; set; }
            public string IsoRevision { get; set; }
            public string Lock { get; set; }
            public string Code { get; set; }
            public int IsoUniqeRevID { get; set; }
        }
        public class IsoRevisionData
        {
            public int MTOID { get; set; }
            public int MaterialID { get; set; }
            public int ProjectID { get; set; }
            public string ISO { get; set; }
            public string Ident_no { get; set; }
            public string qty { get; set; }
            public string qty_unit { get; set; }
            public string Fabrication_Type { get; set; }
            public bool ReleasedMaterial { get; set; } = false;
        }
        public class SPMATFinalDBData
        {
            public int MaterialID { get; set; }
            public int ProjectID { get; set; }
            public string Discipline { get; set; }
            public string Area { get; set; }
            public string Unit { get; set; }
            public string Phase { get; set; }
            public string Const_Area { get; set; }
            public string ISO { get; set; }
            public string Ident_no { get; set; }
            public string qty { get; set; }
            public string qty_unit { get; set; }
            public string Fabrication_Type { get; set; }
            public string Spec { get; set; }
            public string IsoRevisionDate { get; set; }
            public string IsoRevision { get; set; }
            public string Lock { get; set; }
            public string Code { get; set; }
            public int? FileID { get; set; }
            public string ImportedStatus { get; set; }
            public int IsoUniqeRevID { get; set; }
        }
        public class SPMATDeletedData
        {
            public int DelID { get; set; }
            public int MaterialID { get; set; }
            public string Discipline { get; set; }
            public string Area { get; set; }
            public string Unit { get; set; }
            public string Phase { get; set; }
            public string Const_Area { get; set; }
            public string ISO { get; set; }
            public string Ident_no { get; set; }
            public decimal qty { get; set; }
            public string qty_unit { get; set; }
            public string Fabrication_Type { get; set; }
            public string Spec { get; set; }
            public string IsoRevisionDate { get; set; }
            public string IsoRevision { get; set; }
            public string IsLocked { get; set; }
            public string Code { get; set; }
            public string ImportStatus { get; set; }
            public string Changes { get; set; }
            public int MTOID { get; set; }
            public int IsoUniqeRevID { get; set; }
        }
        public class ExportedFiles
        {
            public int FinalFileID { get; set; }
            public string FinalFileName { get; set; }
            public string FileMTOIDs { get; set; }
            public string FileCompleted { get; set; }
            public string Import { get; set; }
            public byte[] FileData { get; set; }
        }
        #endregion
        #region DataCalls
        public static List<SpecData> LoadSpectDataFromDB(string mainProjectID)
        {
            List<SpecData> dbspec = new List<SpecData>();
            try
            {
                string constr = "Data Source = (DESCRIPTION = (ADDRESS = (PROTOCOL = tcp)(HOST =jbg1-ora5)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = SDBFT))); User Id = m_sys; Password = manager1;";

                string cmdtext = @"
SELECT 
    M_SYS.M_SPEC_ITEMS.PROJ_ID AS Project,
    M_SYS.M_SPEC_HEADERS.SPEC_CODE AS LineClass,
    M_SYS.M_SPEC_ITEMS.SHORT_CODE AS ShortCode,
    M_SYS.M_IDENTS.IDENT_CODE AS Ident,
    M_SYS.M_COMMODITY_CODES.COMMODITY_CODE,
    M_SYS.M_COMMODITY_CODE_NLS.SHORT_DESC,
    M_SYS.M_COMMODITY_CODE_NLS.DESCRIPTION,
    M_SYS.M_IDENTS.INPUT_1 AS Size_sch1,
    M_SYS.M_IDENTS.INPUT_2 AS Size_sch2,
    M_SYS.M_IDENTS.INPUT_3 AS Size_sch3,
    M_SYS.M_IDENTS.INPUT_4 AS Size_sch4,
    M_SYS.M_IDENTS.INPUT_5 AS Size_sch5,
    M_SYS.M_SPEC_HEADERS.XREV AS Spec_Revision,
    M_SYS.M_SPEC_HEADERS.PUBLISHED,
    M_SYS.M_SPEC_HEADERS.DATE_PUBLISHED AS Published_Date,
    M_SYS.M_SPEC_ITEMS.OPTION_CODE,
    M_SYS.M_SHORT_CODE_OPTION_CODES.SCOC_COMMENT AS Option_Code_Desc
FROM M_SYS.M_SPEC_ITEMS
INNER JOIN M_SYS.M_RELEASED_SPEC_IDENTS
    ON M_SYS.M_SPEC_ITEMS.SPEC_ITEM_ID = M_SYS.M_RELEASED_SPEC_IDENTS.SPEC_ITEM_ID
    AND M_SYS.M_SPEC_ITEMS.PROJ_ID = M_SYS.M_RELEASED_SPEC_IDENTS.PROJ_ID
INNER JOIN M_SYS.M_IDENTS
    ON M_SYS.M_RELEASED_SPEC_IDENTS.IDENT = M_SYS.M_IDENTS.IDENT
INNER JOIN M_SYS.M_SPEC_HEADERS
    ON M_SYS.M_RELEASED_SPEC_IDENTS.SPEC_HEADER_ID = M_SYS.M_SPEC_HEADERS.SPEC_HEADER_ID
INNER JOIN M_SYS.M_COMMODITY_CODES
    ON M_SYS.M_IDENTS.COMMODITY_ID = M_SYS.M_COMMODITY_CODES.COMMODITY_ID
INNER JOIN M_SYS.M_PROJECTS
    ON M_SYS.M_COMMODITY_CODES.PROJ_ID = M_SYS.M_PROJECTS.PROJ_ID
INNER JOIN M_SYS.M_COMMODITY_CODE_NLS
    ON M_SYS.M_PROJECTS.PROJ_ID = M_SYS.M_COMMODITY_CODE_NLS.PROJ_ID
    AND M_SYS.M_COMMODITY_CODES.COMMODITY_ID = M_SYS.M_COMMODITY_CODE_NLS.COMMODITY_ID
INNER JOIN M_SYS.M_COMMODITY_GROUPS
    ON M_SYS.M_COMMODITY_CODES.GROUP_ID = M_SYS.M_COMMODITY_GROUPS.GROUP_ID
INNER JOIN M_SYS.M_PARTS
    ON M_SYS.M_PARTS.PART_ID = M_SYS.M_COMMODITY_CODES.PART_ID
INNER JOIN M_SYS.M_SHORT_CODE_OPTION_CODES
    ON M_SYS.M_SPEC_ITEMS.OPTION_CODE = M_SYS.M_SHORT_CODE_OPTION_CODES.OPTION_CODE
    AND M_SYS.M_SPEC_ITEMS.SHORT_CODE = M_SYS.M_SHORT_CODE_OPTION_CODES.SHORT_CODE
WHERE M_SYS.M_COMMODITY_CODE_NLS.NLS_ID = 1
  AND M_SYS.M_SHORT_CODE_OPTION_CODES.SCOC_USAGE = 'SP3D'
  AND M_SYS.M_RELEASED_SPEC_IDENTS.PROJ_ID = :mainProjectID
  AND (M_SYS.M_SPEC_HEADERS.SPEC_CODE, M_SYS.M_SPEC_HEADERS.XREV) IN (
      SELECT SPEC_CODE, MAX(XREV)
      FROM M_SYS.M_SPEC_HEADERS
      WHERE Date_Published is not null and PROJ_ID = :mainProjectID
      GROUP BY SPEC_CODE
  )
ORDER BY LineClass, ShortCode";


                using (OracleConnection conn = new OracleConnection(constr))
                using (OracleCommand cmd = new OracleCommand(cmdtext, conn))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Add(new OracleParameter("mainProjectID", mainProjectID));
                    cmd.InitialLOBFetchSize = -1;
                    conn.Open();
                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {
                        dr.FetchSize = dr.RowSize * 1000;
                        while (dr.Read())
                        {
                            SpecData item = new SpecData();
                            item.Project = dr.GetValue(0).ToString();
                            item.Lineclass = dr.GetValue(1).ToString();
                            item.Shortcode = dr.GetValue(2).ToString();
                            item.Ident = dr.GetValue(3).ToString();
                            item.Commodity_code = dr.GetValue(4).ToString();
                            item.Short_desc = dr.GetValue(5).ToString();
                            item.Description = dr.GetValue(6).ToString();
                            item.Size_sch1 = dr.GetValue(7).ToString();
                            item.Size_sch2 = dr.GetValue(8).ToString();
                            item.Size_sch3 = dr.GetValue(9).ToString();
                            item.Size_sch4 = dr.GetValue(10).ToString();
                            item.Size_sch5 = dr.GetValue(11).ToString();
                            item.Spec_revision = dr.GetValue(12).ToString();
                            item.Published = dr.GetValue(13).ToString();
                            item.Published_date = dr.GetValue(14).ToString();
                            item.Option_code = dr.GetValue(15).ToString();
                            item.Option_code_desc = dr.GetValue(16).ToString();
                            dbspec.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var err = ex.Message;
            }
            return dbspec;
        }
        public static List<DDLList> LoadProjectsSpecs(string Company)
        {
            List<DDLList> projlst = new List<DDLList>();
            try
            {
                // Please replace the connection string attribute settings
                string constr = "Data Source = (DESCRIPTION = (ADDRESS = (PROTOCOL = tcp)(HOST =jbg1-ora5)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = SDBFT))); User Id = m_sys; Password = manager1;";
                //OracleConnection con = new OracleConnection(constr);
                var cmdtext = "Select Distinct PROJ_ID, TO_CHAR(PROJ_ID)|| ' - ' || DESCRIPTION from M_PROJECTS where PROJ_ID = '" + Company + "'";
                using (OracleConnection conn = new OracleConnection(constr))
                using (OracleCommand cmd = new OracleCommand(cmdtext, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {

                        while (dr.Read())
                        {
                            DDLList proj = new DDLList();
                            proj.DDLList_ID = dr[0].ToString();
                            proj.DDLListName = dr[1].ToString();
                            if (!projlst.Contains(proj))
                            {
                                projlst.Add(proj);
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                var err = ex.Message;
            }
            return projlst;
        }

        public static List<DDLList> LoadSubProjects(string Spec)
        {
            List<DDLList> sublst = new List<DDLList>();
            try
            {
                // Please replace the connection string attribute settings
                string constr = "Data Source = (DESCRIPTION = (ADDRESS = (PROTOCOL = tcp)(HOST =jbg1-ora5)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = SDBFT))); User Id = m_sys; Password = manager1;";
                //OracleConnection con = new OracleConnection(constr);
                var cmdtext = "Select Distinct PROJECT_ID, PROJECT_ID||' - '||PROJECT_SHORT_DESC from W_SUB_PROJECT where M_PROJECT_ID ='" + Spec.Trim() + "'";
                using (OracleConnection conn = new OracleConnection(constr))
                using (OracleCommand cmd = new OracleCommand(cmdtext, conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {

                        while (dr.Read())
                        {
                            DDLList items = new DDLList();
                            items.DDLList_ID = dr.GetValue(0).ToString();
                            items.DDLListName = dr.GetValue(1).ToString();
                            if (!sublst.Contains(items))
                            {
                                sublst.Add(items);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                var err = ex.Message;
            }
            return sublst;
        }
        public static List<DDLList> LoadCients()
        {
            List<DDLList> clientlst = new List<DDLList>();
            try
            {

                // Please replace the connection string attribute settings
                string constr = "Data Source = (DESCRIPTION = (ADDRESS = (PROTOCOL = tcp)(HOST =jbg1-ora5)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = SDBFT))); User Id = m_sys; Password = manager1;";
                //OracleConnection con = new OracleConnection(constr);
                using (OracleConnection conn = new OracleConnection(constr))
                using (OracleCommand cmd = new OracleCommand("Select distinct PGR_ID,PGR_Code from M_PROJECT_GROUPS order by PGR_Code", conn))
                {
                    conn.Open();

                    // reader is IDisposable and should be closed
                    using (OracleDataReader dr = cmd.ExecuteReader())
                    {


                        while (dr.Read())
                        {
                            DDLList item = new DDLList();
                            item.DDLList_ID = dr.GetValue(0).ToString();
                            item.DDLListName = dr.GetValue(1).ToString();
                            if (!clientlst.Contains(item))
                            {
                                clientlst.Add(item);
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                var err = ex.Message;
            }
            return clientlst;
        }
        internal static List<DDLList> GetAllRefClients()
        {
            List<DDLList> u = new List<DDLList>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT distinct [fld_RefDataProject],[fld_ClientTitle],[fld_ID] FROM [dbo].[tbl_Clients] where [fld_RefDataProject] is not null order by fld_ClientTitle";
                using (cn = new System.Data.SqlClient.SqlConnection(tmpcon))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        DDLList Client = new DDLList();
                        Client.DDLList_ID = dr[0].ToString();
                        Client.DDLListName = dr[1].ToString();
                        Client.DDLID = dr[2].ToString();
                        if (!u.Contains(Client))
                        {
                            u.Add(Client);
                        }
                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return u;
        }
        internal static List<DDLList> GetRefProjects(string ClientID)
        {
            List<DDLList> projlst = new List<DDLList>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT Distinct [fld_ID],[fld_ProjectNo]+ ' - ' +[fld_ProjectDescription],[fld_ID] FROM [dbo].[tbl_ProjectInformation] where [fld_ClientID]= " + ClientID.Trim();
                using (cn = new System.Data.SqlClient.SqlConnection(tmpcon))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        DDLList proj = new DDLList();
                        proj.DDLList_ID = dr[0].ToString();
                        proj.DDLListName = dr[1].ToString();
                        proj.DDLID = dr[2].ToString();
                        if (!projlst.Contains(proj))
                        {
                            projlst.Add(proj);
                        }
                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return projlst;
        }

        internal static List<DDLList> GetProjectISO(string projid, bool All = true)
        {
            List<DDLList> isolst = new List<DDLList>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "";
                if (All)
                {
                    query = "SELECT Distinct [ISO],[ISO]+'- Rev: '+ [IsoRevision] FROM [dbo].[SPMAT_REQData] where ProjectID=" + projid + " and Deleted=0  and isnull(IsLocked,'False')='True' and (IsoRevisionDate is not null and  IsoRevisionDate<>'') order by 1";
                }
                else
                {
                    query = "SELECT Distinct [ISO],[ISO]+'- Rev: '+ [IsoRevision] FROM [dbo].[SPMAT_REQData] where ProjectID=" + projid + " and Checked=0 and Deleted=0 and isnull(IsLocked,'False')='True' and (IsoRevisionDate is not null and  IsoRevisionDate<>'')  order by 1";
                }
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        DDLList iso = new DDLList();
                        iso.DDLList_ID = dr[0].ToString();
                        iso.DDLListName = dr[1].ToString();

                        if (!isolst.Contains(iso))
                        {
                            isolst.Add(iso);
                        }
                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return isolst;
        }

        internal static List<SPMATDBData> GetIsoSheetMTOData(string isosheet, string ProjectID, bool All = true)
        {
            List<SPMATDBData> isodata = new List<SPMATDBData>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "";

                if (All)
                {
                    query = $@"
    SELECT [MaterialID],[ProjectID],[Discipline],[Area],[Unit],[Phase],[Const_Area],[ISO],
           [Ident_no],[qty],[qty_unit],[Fabrication_Type],[Spec],[IsoRevisionDate],
           [IsoRevision],[IsLocked],[Code],IsoUniqeRevID
    FROM [dbo].[SPMAT_REQData]
    WHERE Moved = 0 and Deleted=0 AND LTRIM(RTRIM(ISO)) = '{isosheet.Trim()}' AND ProjectID = {ProjectID.Trim()}

    UNION ALL

    SELECT [MaterialID],[ProjectID],[Discipline],[Area],[Unit],[Phase],[Const_Area],[ISO],
           [Ident_no],[qty],[qty_unit],[Fabrication_Type],[Spec],[IsoRevisionDate],
           [IsoRevision],[IsLocked],[Code],IsoUniqeRevID
    FROM [dbo].[SPMAT_REQData_Temp]
    WHERE LTRIM(RTRIM(ISO)) = '{isosheet.Trim()}' AND ProjectID = {ProjectID.Trim()}";
                }
                else
                {
                    query = $@"
    SELECT [MaterialID],[ProjectID],[Discipline],[Area],[Unit],[Phase],[Const_Area],[ISO],
           [Ident_no],[qty],[qty_unit],[Fabrication_Type],[Spec],[IsoRevisionDate],
           [IsoRevision],[IsLocked],[Code],IsoUniqeRevID
    FROM [dbo].[SPMAT_REQData]
    WHERE LTRIM(RTRIM(ISO)) = '{isosheet.Trim()}' AND Checked = 0 and Deleted=0 AND ProjectID = {ProjectID.Trim()}

    UNION ALL

    SELECT [MaterialID],[ProjectID],[Discipline],[Area],[Unit],[Phase],[Const_Area],[ISO],
           [Ident_no],[qty],[qty_unit],[Fabrication_Type],[Spec],[IsoRevisionDate],
           [IsoRevision],[IsLocked],[Code],IsoUniqeRevID
    FROM [dbo].[SPMAT_REQData_Temp]
    WHERE LTRIM(RTRIM(ISO)) = '{isosheet.Trim()}' AND Checked = 0 AND ProjectID = {ProjectID.Trim()}";
                }

                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        SPMATDBData spdata = new SPMATDBData();
                        spdata.MaterialID = Convert.ToInt32(dr[0]);
                        spdata.ProjectID = Convert.ToInt32(dr[1]);
                        spdata.Discipline = dr[2].ToString();
                        spdata.Area = dr[3].ToString();
                        spdata.Unit = dr[4].ToString();
                        spdata.Phase = dr[5].ToString();
                        spdata.Const_Area = dr[6].ToString();
                        spdata.ISO = dr[7].ToString();
                        spdata.Ident_no = dr[8].ToString();
                        spdata.qty = dr[9].ToString().Trim();
                        spdata.qty_unit = dr[10].ToString();
                        spdata.Fabrication_Type = dr[11].ToString();
                        spdata.Spec = dr[12].ToString();
                        spdata.IsoRevisionDate = dr[13].ToString();
                        spdata.IsoRevision = dr[14].ToString();
                        spdata.Lock = dr[15].ToString();
                        spdata.Code = dr[16].ToString();
                        spdata.IsoUniqeRevID=int.Parse(dr[17].ToString());

                        if (!isodata.Contains(spdata))
                        {
                            isodata.Add(spdata);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return isodata;
        }
        internal static List<SPMATDBData> GetWorkingMTOData(string ProjectID)
        {
            List<SPMATDBData> isodata = new List<SPMATDBData>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = " SELECT [MaterialID],[ProjectID],[Discipline],[Area],[Unit],[Phase],[Const_Area],[ISO],[Ident_no],[qty],[qty_unit],[Fabrication_Type],[Spec],[IsoRevisionDate],[IsoRevision],[IsLocked],[Code],IsoUniqeRevID  FROM [dbo].[SPMAT_REQData] where Checked = 1 and Moved = 0 and Deleted=0 and ProjectID = " + ProjectID.Trim() +
                               " and MaterialID not in (SELECT distinct [MaterialID] FROM[dbo].[SPMAT_REQData_Temp] where Checked = 1 and Moved = 0 and Deleted=0 and ProjectID = " + ProjectID.Trim() + " ) " +
                               " UNION ALL " +
                               " SELECT [MaterialID],[ProjectID],[Discipline],[Area],[Unit],[Phase],[Const_Area],[ISO],[Ident_no],[qty],[qty_unit],[Fabrication_Type],[Spec],[IsoRevisionDate],[IsoRevision],[IsLocked],[Code],IsoUniqeRevID FROM [dbo].[SPMAT_REQData_Temp] where Checked = 1 and Moved = 0 and ProjectID = " + ProjectID.Trim();
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        SPMATDBData spdata = new SPMATDBData();
                        spdata.MaterialID = Convert.ToInt32(dr[0]);
                        spdata.ProjectID = Convert.ToInt32(dr[1]);
                        spdata.Discipline = dr[2].ToString();
                        spdata.Area = dr[3].ToString();
                        spdata.Unit = dr[4].ToString();
                        spdata.Phase = dr[5].ToString();
                        spdata.Const_Area = dr[6].ToString();
                        spdata.ISO = dr[7].ToString();
                        spdata.Ident_no = dr[8].ToString();
                        spdata.qty = dr[9].ToString().Trim();
                        spdata.qty_unit = dr[10].ToString();
                        spdata.Fabrication_Type = dr[11].ToString();
                        spdata.Spec = dr[12].ToString();
                        spdata.IsoRevisionDate = dr[13].ToString();
                        spdata.IsoRevision = dr[14].ToString();
                        spdata.Lock = dr[15].ToString();
                        spdata.Code = dr[16].ToString();
                        spdata.IsoUniqeRevID = int.Parse(dr[17].ToString());

                        if (!isodata.Contains(spdata))
                        {
                            isodata.Add(spdata);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return isodata;
        }

        //internal static void UpdateQtyInDatabase(SPMATDBData toUpdate)
        //{
        //    System.Data.SqlClient.SqlConnection cn = null;
        //    using (cn = new System.Data.SqlClient.SqlConnection(conMat))
        //    {
        //        string query = @"
        //    UPDATE [dbo].[SPMAT_REQData]
        //    SET 
        //        qty = @qty,
        //        Unit = @Unit,
        //        Phase = @Phase,
        //        Const_Area = @ConstArea,
        //        Spec = @Spec,
        //        Ident_no = @IdentNo,
        //        Fabrication_Type = @FabricationType,
        //        IsoRevision = @IsoRevision,
        //        IsoRevisionDate = @IsoRevisionDate,
        //        IsLocked = @IsLocked,
        //        Code = @Code
        //    WHERE MaterialID = @MaterialID";

        //        SqlCommand cmd = new SqlCommand(query, cn);
        //        cmd.Parameters.AddWithValue("@qty", DecParse(toUpdate.qty.ToString()).Value);
        //        cmd.Parameters.AddWithValue("@Unit", toUpdate.Unit ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@Phase", toUpdate.Phase ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@ConstArea", toUpdate.Const_Area ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@Spec", toUpdate.Spec ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@IdentNo", toUpdate.Ident_no ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@FabricationType", toUpdate.Fabrication_Type ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@IsoRevision", toUpdate.IsoRevision ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@IsoRevisionDate", toUpdate.IsoRevisionDate ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@IsLocked", toUpdate.Lock ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@Code", toUpdate.Code ?? (object)DBNull.Value);
        //        cmd.Parameters.AddWithValue("@MaterialID", toUpdate.MaterialID);

        //        cn.Open();
        //        cmd.ExecuteNonQuery();
        //    }
        //}
        internal static void FinalizeMaterialUpdate(SPMATDBData item)
        {
            using (var cn = new SqlConnection(conMat))
            {
                cn.Open();

                // 1. Update Checked field
                var updateCmd = new SqlCommand("UPDATE [dbo].[SPMAT_REQData] SET [Checked] = 1 WHERE [MaterialID] = @MaterialID", cn);
                updateCmd.Parameters.AddWithValue("@MaterialID", item.MaterialID);
                updateCmd.ExecuteNonQuery();

                // 2. Update Checked field in temp table
                var updateTempCmd = new SqlCommand("UPDATE [dbo].[SPMAT_REQData_Temp] SET [Checked] = 1 WHERE [MaterialID] = @MaterialID", cn);
                updateTempCmd.Parameters.AddWithValue("@MaterialID", item.MaterialID);
                updateTempCmd.ExecuteNonQuery();


            }
        }



        internal static List<string> GetUnitsByProject(string projid)
        {
            List<string> units = new List<string>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT  distinct [Unit]  FROM [dbo].[SPMAT_REQData] where ProjectID=" + projid.Trim();
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        var un = dr[0].ToString();
                        if (!units.Contains(un))
                        {
                            units.Add(un);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return units;
        }

        internal static List<string> GetPhasesByProject(string projid)
        {
            List<string> Phase = new List<string>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT  distinct [Phase]  FROM [dbo].[SPMAT_REQData] where ProjectID=" + projid.Trim();
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        var ph = dr[0].ToString();
                        if (!Phase.Contains(ph))
                        {
                            Phase.Add(ph);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return Phase;
        }

        internal static List<string> GetConstAreasByProject(string projid)
        {
            List<string> constarea = new List<string>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT  distinct [Const_Area]  FROM [dbo].[SPMAT_REQData] where ProjectID=" + projid.Trim();
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        var ca = dr[0].ToString();
                        if (!constarea.Contains(ca))
                        {
                            constarea.Add(ca);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return constarea;
        }

        public static Dictionary<string, List<string>> GetUnitPhaseMap(string projectId)
        {
            var result = new Dictionary<string, List<string>>();
            System.Data.SqlClient.SqlConnection cn = null;
            using (cn = new System.Data.SqlClient.SqlConnection(conMat))
            {
                string query = @"
            SELECT DISTINCT Unit, Phase
            FROM [dbo].[SPMAT_REQData]
            WHERE ProjectID = @ProjectID AND Unit IS NOT NULL AND Phase IS NOT NULL";

                SqlCommand cmd = new SqlCommand(query, cn);
                cmd.Parameters.AddWithValue("@ProjectID", projectId);

                cn.Open();
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string unit = reader["Unit"].ToString();
                    string phase = reader["Phase"].ToString();

                    if (!result.ContainsKey(unit))
                        result[unit] = new List<string>();

                    if (!result[unit].Contains(phase))
                        result[unit].Add(phase);
                }
            }

            return result;
        }

        internal static Dictionary<string, Dictionary<string, List<string>>> GetUnitPhaseConstAreaMap(string projid)
        {
            var result = new Dictionary<string, Dictionary<string, List<string>>>();
            using (var cn = new System.Data.SqlClient.SqlConnection(conMat))
            {
                string query = @"
            SELECT DISTINCT Unit, Phase, Const_Area
            FROM [dbo].[SPMAT_REQData]
            WHERE ProjectID = @ProjectID 
              AND Unit IS NOT NULL 
              AND Phase IS NOT NULL 
              AND Const_Area IS NOT NULL";

                SqlCommand cmd = new SqlCommand(query, cn);
                cmd.Parameters.AddWithValue("@ProjectID", projid);

                cn.Open();
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string unit = reader["Unit"].ToString();
                    string phase = reader["Phase"].ToString();
                    string constArea = reader["Const_Area"].ToString();

                    if (!result.ContainsKey(unit))
                        result[unit] = new Dictionary<string, List<string>>();

                    if (!result[unit].ContainsKey(phase))
                        result[unit][phase] = new List<string>();

                    if (!result[unit][phase].Contains(constArea))
                        result[unit][phase].Add(constArea);
                }
            }

            return result;

        }

        internal static List<string> GetSpecsByProject(string projid)
        {
            List<string> specs = new List<string>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT  distinct [Spec]  FROM [dbo].[SPMAT_REQData] where ProjectID=" + projid.Trim();
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        var sp = dr[0].ToString();
                        if (!specs.Contains(sp))
                        {
                            specs.Add(sp);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return specs;
        }

        public static bool AreObjectsDifferentExcept<T>(T obj1, T obj2, params string[] excludedFields)
        {
            if (obj1 == null || obj2 == null)
                throw new ArgumentNullException("Objects cannot be null");

            var type = typeof(T);
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                 .Where(p => !excludedFields.Contains(p.Name));

            foreach (var property in properties)
            {
                var value1 = property.GetValue(obj1);
                var value2 = property.GetValue(obj2);

                // Skip comparison if value2 is null or an empty string
                if (value2 == null || (value2 is string str && string.IsNullOrEmpty(str)))
                {
                    continue;
                }

                if (!Equals(value1, value2))
                    return true;
            }

            return false;
        }

        internal static List<SPMATData> GetMTOData(string projid,bool Completed=false)
        {
            List<SPMATData> mtodata = new List<SPMATData>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "";
                if (Completed)
                {
                    query = "SELECT MTOID,Discipline,Area,Unit,Phase,Const_Area,ISO,Ident_no,qty,qty_unit,Fabrication_Type,Spec,'' as Pos,IsoRevisionDate,IsoRevision,IsLocked,Code,CAST(CASE WHEN ImportedStatus IS NULL THEN 'Not Imported' ELSE ImportedStatus END AS NVARCHAR(50)) as ImportedStatus,IsoUniqeRevID  FROM [dbo].[SPMAT_MTOData] where ProjectID=" + projid + " and (Imported =1 or FileID is not null) and IsDeleted=0 order by ISO";
                }
                else
                {

                    query = "SELECT MTOID,Discipline,Area,Unit,Phase,Const_Area,ISO,Ident_no,qty,qty_unit,Fabrication_Type,Spec,'' as Pos,IsoRevisionDate,IsoRevision,IsLocked,Code,CAST(CASE WHEN ImportedStatus IS NULL THEN 'Not Imported' ELSE ImportedStatus END AS NVARCHAR(50)) as ImportedStatus,IsoUniqeRevID  FROM [dbo].[SPMAT_MTOData] where ProjectID=" + projid + " and (Imported <>1 and FileID is null)  and IsDeleted=0 order by ISO";
                }


                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        SPMATData mto = new SPMATData();
                        mto.MTOID = int.Parse(dr[0].ToString());
                        mto.Discipline = dr[1].ToString();
                        mto.Area = dr[2].ToString(); ;
                        mto.Unit = dr[3].ToString();
                        mto.Phase = dr[4].ToString();
                        mto.Const_Area = dr[5].ToString();
                        mto.ISO = dr[6].ToString();
                        mto.Ident_no = dr[7].ToString();
                        mto.qty = DecParse(dr[8].ToString().Trim()).Value;
                        mto.qty_unit = dr[9].ToString();
                        mto.Fabrication_Type = dr[10].ToString();
                        mto.Spec = dr[11].ToString();
                        mto.Pos = dr[12].ToString();
                        mto.IsoRevisionDate = dr[13].ToString();
                        mto.IsoRevision = dr[14].ToString();
                        mto.IsLocked = dr[15].ToString();
                        mto.Code = dr[16].ToString();
                        mto.ImportStatus = dr[17].ToString();
                        if (dr[18] != DBNull.Value)
                        {
                            mto.IsoUniqeRevID = int.Parse(dr[18].ToString());
                        }

                        if (!mtodata.Contains(mto))
                        {
                            mtodata.Add(mto);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return mtodata;
        }

        public static void MarkMaterialAsChecked(int materialID)
        {
            using (var cn = new SqlConnection(conMat))
            {
                cn.Open();
                var cmd = new SqlCommand("UPDATE [dbo].[SPMAT_REQData] SET [Checked] = 1, Deleted=1 WHERE [MaterialID] = @MaterialID", cn);
                cmd.Parameters.AddWithValue("@MaterialID", materialID);
                cmd.ExecuteNonQuery();
            }
        }
        public static void MarkTempMaterialAsChecked(int materialID)
        {
            using (var cn = new SqlConnection(conMat))
            {
                cn.Open();
                var cmd = new SqlCommand("UPDATE [dbo].[SPMAT_REQData_Temp] SET [Checked] = 1 WHERE [MaterialID] = @MaterialID", cn);
                cmd.Parameters.AddWithValue("@MaterialID", materialID);
                cmd.ExecuteNonQuery();
            }
        }
        public static void DeleteTempMaterial(int materialID)
        {
            using (var cn = new SqlConnection(conMat))
            {
                cn.Open();
                var cmd = new SqlCommand("Delete from [dbo].[SPMAT_REQData_Temp] WHERE [MaterialID] = @MaterialID", cn);
                cmd.Parameters.AddWithValue("@MaterialID", materialID);
                cmd.ExecuteNonQuery();
            }
        }

        public static void DeleteMTOEntry(string Iso)
        {
            using (var cn = new SqlConnection(conMat))
            {
                cn.Open();
                var cmd = new SqlCommand(@"
            DELETE FROM SPMAT_IntrimData
            WHERE ISO = @ISO ", cn);
                cmd.Parameters.AddWithValue("@ISO", Iso);
                cmd.ExecuteNonQuery();
            }
        }

        public static void UncheckREQEntry(string iso)
        {
            using (var cn = new SqlConnection(conMat))
            {
                cn.Open();
                var cmd = new SqlCommand(@"
            UPDATE SPMAT_REQData
            SET Checked = 0, Moved = 0, Deleted=0 
            WHERE ISO = @ISO ", cn);
                cmd.Parameters.AddWithValue("@ISO", iso);
                cmd.ExecuteNonQuery();

                var cmd2 = new SqlCommand(@"
            DELETE FROM SPMAT_REQData_Temp
            WHERE ISO = @ISO and Processed=0 ", cn);
                cmd2.Parameters.AddWithValue("@ISO", iso);
                cmd2.ExecuteNonQuery();
            }
        }

        //public static int SaveExportRecord(string fileName, List<int> materialIDs)
        //{
        //    string ids = string.Join(",", materialIDs);
        //    int FileID = 0;
        //    using (SqlConnection conn = new SqlConnection(conMat))
        //    {
        //        conn.Open();

        //        // Insert into SPMAT_FIleExports
        //        string insertQuery = @"
        //    INSERT INTO SPMAT_FileExports (FinalFileName, FileMTOIDs)
        //    VALUES (@FileName, @MTOIDs) SELECT SCOPE_IDENTITY()";
        //        using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
        //        {
        //            insertCmd.Parameters.AddWithValue("@FileName", fileName);
        //            insertCmd.Parameters.AddWithValue("@MTOIDs", ids);
        //            FileID = Convert.ToInt32(insertCmd.ExecuteScalar());

        //        }

        //        // Update SPMAT_IntrimData
        //        string updateQuery = $@"
        //    UPDATE SPMAT_MTOData
        //    SET FileID = @FileID
        //    WHERE MTOID IN ({ids})";
        //        using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
        //        {
        //            updateCmd.Parameters.AddWithValue("@FileID", FileID);
        //            updateCmd.ExecuteNonQuery();
        //        }
        //    }
        //    return FileID;
        //}
        public static void SaveExportRecordFile(int FileID, byte[] filedata)
        {
            using (SqlConnection conn = new SqlConnection(conMat))
            {
                conn.Open();

                // Update SPMAT_IntrimData
                string updateQuery = $@"
            UPDATE SPMAT_FileExports
            SET FileData = @FileData
            WHERE FinalFileID =@FileID";
                using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                {
                    updateCmd.Parameters.AddWithValue("@FileID", FileID);
                    updateCmd.Parameters.AddWithValue("@FileData", filedata);
                    updateCmd.ExecuteNonQuery();
                }
            }
        }
        public static void MoveToIntrim(SPMATDBData data)
        {

            using (SqlConnection conn = new SqlConnection(conMat))
            {

                conn.Open();
                string markDeletedQuery = @"
            UPDATE [dbo].[SPMAT_IntrimData]
            SET IsDeleted = 1
            WHERE MaterialID = @MaterialID";

                using (SqlCommand markDeletedCmd = new SqlCommand(markDeletedQuery, conn))
                {
                    markDeletedCmd.Parameters.AddWithValue("@MaterialID", data.MaterialID);
                    markDeletedCmd.ExecuteNonQuery();
                }

                // Insert into SPMAT_FIleExports
                string insertQuery = @"
IF NOT EXISTS (
    SELECT 1 FROM [dbo].[SPMAT_IntrimData]
    WHERE MaterialID = @MaterialID and IsDeleted=0
)
BEGIN
    INSERT INTO [dbo].[SPMAT_IntrimData]
        ([MaterialID], [ProjectID], [Discipline], [Area], [Unit], [Phase],
         [Const_Area], [ISO], [Ident_no], [qty], [qty_unit], [Fabrication_Type],
         [Spec], [IsoRevisionDate], [IsoRevision], [IsLocked], [Code], [Checked],IsoUniqeRevID)
    VALUES
        (@MaterialID, @ProjectID, @Discipline, @Area, @Unit, @Phase,
         @Const_Area, @ISO, @Ident_no, @qty, @qty_unit, @Fabrication_Type,
         @Spec, @IsoRevisionDate, @IsoRevision, @IsLocked, @Code, @Checked,@IsoUniqeRevID)
END";


                using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                {
                    // Add parameters with example values
                    cmd.Parameters.AddWithValue("@MaterialID", data.MaterialID);
                    cmd.Parameters.AddWithValue("@ProjectID", data.ProjectID);
                    cmd.Parameters.AddWithValue("@Discipline", data.Discipline);
                    cmd.Parameters.AddWithValue("@Area", data.Area);
                    cmd.Parameters.AddWithValue("@Unit", data.Unit);
                    cmd.Parameters.AddWithValue("@Phase", data.Phase);
                    cmd.Parameters.AddWithValue("@Const_Area", data.Const_Area);
                    cmd.Parameters.AddWithValue("@ISO", data.ISO);
                    cmd.Parameters.AddWithValue("@Ident_no", data.Ident_no);
                    cmd.Parameters.AddWithValue("@qty", DecParse(data.qty.Trim()).Value);
                    cmd.Parameters.AddWithValue("@qty_unit", data.qty_unit);
                    cmd.Parameters.AddWithValue("@Fabrication_Type", data.Fabrication_Type);
                    cmd.Parameters.AddWithValue("@Spec", data.Spec);
                    cmd.Parameters.AddWithValue("@IsoRevisionDate", data.IsoRevisionDate);
                    cmd.Parameters.AddWithValue("@IsoRevision", data.IsoRevision);
                    cmd.Parameters.AddWithValue("@IsLocked", data.Lock);
                    cmd.Parameters.AddWithValue("@Code", data.Code);
                    cmd.Parameters.AddWithValue("@Checked", true);
                    cmd.Parameters.AddWithValue("@IsoUniqeRevID", data.IsoUniqeRevID);


                    cmd.ExecuteNonQuery();
                }


                // Update SPMAT_IntrimData
                string updateQuery = $@"
            UPDATE SPMAT_REQData
            SET Moved = 1,
                MovedDate = GETDATE()
            WHERE MaterialID =" + data.MaterialID.ToString();
                using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                {
                    updateCmd.ExecuteNonQuery();
                }
                string updateQuery2 = $@"
            UPDATE SPMAT_REQData_Temp
            SET Moved = 1,
                MovedDate = GETDATE()
            WHERE MaterialID =" + data.MaterialID.ToString();
                using (SqlCommand updateCmd2 = new SqlCommand(updateQuery2, conn))
                {
                    updateCmd2.ExecuteNonQuery();
                }
            }
        }
        internal static List<SPMATIntrimData> GetMTOIntrimData(string projid)
        {
            List<SPMATIntrimData> intrimdata = new List<SPMATIntrimData>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT INTID,MaterialID,ProjectID,Discipline,Area,Unit,Phase,Const_Area,ISO,Ident_no,qty,qty_unit,Fabrication_Type,Spec,'' as Pos,IsoRevisionDate,IsoRevision,IsLocked,Code,IsoUniqeRevID FROM [dbo].[SPMAT_IntrimData] where ProjectID=" + projid + "  and MovedToFinal =0 and IsDeleted=0 order by ISO";


                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        SPMATIntrimData intrim = new SPMATIntrimData();
                        intrim.INTID = int.Parse(dr[0].ToString());
                        intrim.MaterialID = int.Parse(dr[1].ToString());
                        intrim.ProjectID = int.Parse(dr[2].ToString());
                        intrim.Discipline = dr[3].ToString();
                        intrim.Area = dr[4].ToString(); ;
                        intrim.Unit = dr[5].ToString();
                        intrim.Phase = dr[6].ToString();
                        intrim.Const_Area = dr[7].ToString();
                        intrim.ISO = dr[8].ToString();
                        intrim.Ident_no = dr[9].ToString();
                        intrim.qty = DecParse(dr[10].ToString().Trim()).Value;
                        intrim.qty_unit = dr[11].ToString();
                        intrim.Fabrication_Type = dr[12].ToString();
                        intrim.Spec = dr[13].ToString();
                        intrim.Pos = dr[14].ToString();
                        intrim.IsoRevisionDate = dr[15].ToString();
                        intrim.IsoRevision = dr[16].ToString();
                        intrim.IsLocked = dr[17].ToString();
                        intrim.Code = dr[18].ToString();
                        intrim.IsoUniqeRevID = int.Parse(dr[19].ToString());

                        if (!intrimdata.Contains(intrim))
                        {
                            intrimdata.Add(intrim);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return intrimdata;
        }

        internal static void MoveToFinal(SPMATIntrimData data)
        {
            using (SqlConnection conn = new SqlConnection(conMat))
            {
                //Check for revs if exist on different rev/. Mark as deleted

                conn.Open();

                string markDeletedQuery = @"
            UPDATE [dbo].[SPMAT_MTOData]
            SET IsDeleted = 1
            WHERE MaterialID = @MaterialID";

                using (SqlCommand markDeletedCmd = new SqlCommand(markDeletedQuery, conn))
                {
                    markDeletedCmd.Parameters.AddWithValue("@MaterialID", data.MaterialID);
                    markDeletedCmd.ExecuteNonQuery();
                }

                // Insert into SPMAT_FIleExports
                string insertQuery = @"
            INSERT INTO [dbo].[SPMAT_MTOData]
                ([MaterialID], [ProjectID], [Discipline], [Area], [Unit], [Phase],
                 [Const_Area], [ISO], [Ident_no], [qty], [qty_unit], [Fabrication_Type],
                 [Spec], [IsoRevisionDate], [IsoRevision], [IsLocked], [Code],IsoUniqeRevID)
            VALUES
                (@MaterialID, @ProjectID, 'MATCON_PIPING', @Area, @Unit, @Phase,
                 @Const_Area, @ISO, @Ident_no, @qty, @qty_unit, @Fabrication_Type,
                 @Spec, @IsoRevisionDate, @IsoRevision, @IsLocked, @Code,@IsoUniqeRevID)";

                using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                {
                    // Add parameters with example values
                    cmd.Parameters.AddWithValue("@MaterialID", data.MaterialID);
                    cmd.Parameters.AddWithValue("@ProjectID", data.ProjectID);
                    cmd.Parameters.AddWithValue("@Discipline", data.Discipline);
                    cmd.Parameters.AddWithValue("@Area", data.Area);
                    cmd.Parameters.AddWithValue("@Unit", data.Unit);
                    cmd.Parameters.AddWithValue("@Phase", data.Phase);
                    cmd.Parameters.AddWithValue("@Const_Area", data.Const_Area);
                    cmd.Parameters.AddWithValue("@ISO", data.ISO);
                    cmd.Parameters.AddWithValue("@Ident_no", data.Ident_no);
                    cmd.Parameters.AddWithValue("@qty", DecParse(data.qty.ToString().Trim()).Value);
                    cmd.Parameters.AddWithValue("@qty_unit", data.qty_unit);
                    cmd.Parameters.AddWithValue("@Fabrication_Type", data.Fabrication_Type);
                    cmd.Parameters.AddWithValue("@Spec", data.Spec);
                    cmd.Parameters.AddWithValue("@IsoRevisionDate", data.IsoRevisionDate);
                    cmd.Parameters.AddWithValue("@IsoRevision", data.IsoRevision);
                    cmd.Parameters.AddWithValue("@IsLocked", data.IsLocked);
                    cmd.Parameters.AddWithValue("@Code", data.Code);
                    cmd.Parameters.AddWithValue("@Checked", true);
                    cmd.Parameters.AddWithValue("@IsoUniqeRevID", data.IsoUniqeRevID);
                    



                    cmd.ExecuteNonQuery();
                }


                // Update SPMAT_IntrimData
                string updateQuery = $@"
            UPDATE SPMAT_IntrimData
            SET MovedToFinal = 1,
                MovedToFinalDate = GETDATE()
            WHERE INTID =" + data.INTID.ToString();
                using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                {
                    updateCmd.ExecuteNonQuery();
                }

                // Update Temp
                string updateTempQuery = $@"
            UPDATE SPMAT_REQData_Temp
            SET Processed = 1 WHERE MaterialID =" + data.MaterialID.ToString();
                using (SqlCommand updatetempCmd = new SqlCommand(updateTempQuery, conn))
                {
                    updatetempCmd.ExecuteNonQuery();
                }
            }

        }

        internal static List<ExportedFiles> GetExportedFiles(string projid)
        {
            List<ExportedFiles> filedata = new List<ExportedFiles>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT Distinct [FinalFileID],[FinalFileName] ,[FileMTOIDs],[FileCompleted],Case when [FileCompleted]=1 then (Select top 1 ImportedStatus from  [dbo].[SPMAT_MTOData] where FileID=[FinalFileID]) END  Import,[FileData] FROM [dbo].[SPMAT_FileExports] " +
                               " where FinalFileID in(Select  distinct FIleID FROM [dbo].[SPMAT_MTOData] where ProjectID = " + projid + " and FileID is not null)";


                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        ExportedFiles file = new ExportedFiles();
                        file.FinalFileID = int.Parse(dr[0].ToString());
                        file.FinalFileName = dr[1].ToString();
                        file.FileMTOIDs = dr[2].ToString(); ;
                        file.FileCompleted = dr[3].ToString();
                        file.Import = dr[4].ToString();
                        if (dr[5] != DBNull.Value)
                        {
                            file.FileData = (byte[])dr[5];
                        }
                        if (!filedata.Contains(file))
                        {
                            filedata.Add(file);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return filedata;
        }

        internal static void UpdateSPMAT_FileExports(int fileId, string fileMTOIDs, string importCode)
        {
            string ids = fileMTOIDs;
            int FileID = fileId;
            using (SqlConnection conn = new SqlConnection(conMat))
            {
                conn.Open();

                // Insert into SPMAT_FIleExports
                string fileQuery = @"
            Update SPMAT_FileExports Set [FileCompleted]=1 where [FinalFileID]=@FileID";
                using (SqlCommand fileCmd = new SqlCommand(fileQuery, conn))
                {
                    fileCmd.Parameters.AddWithValue("@FileID", FileID);
                    fileCmd.ExecuteNonQuery();
                }

                // Update SPMAT_IntrimData
                string updateQuery = $@"
            UPDATE SPMAT_MTOData
                        SET Imported = 1,
                            ImportedDate = GETDATE(),
                            ImportedStatus = @Status
                        WHERE MTOID IN ({ids})";
                using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                {
                    updateCmd.Parameters.AddWithValue("@Status", importCode);
                    updateCmd.ExecuteNonQuery();
                }
                // 3. Process each MaterialID
                foreach (string id in ids.Split(','))
                {
                    if (int.TryParse(id.Trim(), out int materialID))
                    {
                        string holdingSelectQuery = @"
                    SELECT TOP 1 *
                    FROM SPMAT_REQData_Holding
                    WHERE MaterialID = @MaterialID
                    ORDER BY IsoUniqeRevID ASC, InsertedDate ASC";

                        using (SqlCommand selectCmd = new SqlCommand(holdingSelectQuery, conn))
                        {
                            selectCmd.Parameters.AddWithValue("@MaterialID", materialID);
                            using (SqlDataReader reader = selectCmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    int holdingID = Convert.ToInt32(reader["HoldingID"]);

                                    // Copy current REQData to Deleted
                                    string copyQuery = @"
                                INSERT INTO SPMAT_REQData_Deleted (
                                    MaterialID, ProjectID, Discipline, Area, Unit, Phase, Const_Area, ISO, Ident_no,
                                    qty, qty_unit, Fabrication_Type, Spec, IsoRevisionDate, IsoRevision, IsLocked,
                                    Code, Checked, Moved, MovedDate, Processed,IsoUniqeRevID
                                )
                                SELECT MaterialID, ProjectID, Discipline, Area, Unit, Phase, Const_Area, ISO, Ident_no,
                                       qty, qty_unit, Fabrication_Type, Spec, IsoRevisionDate, IsoRevision, IsLocked,
                                       Code, Checked, Moved, MovedDate, 0,IsoUniqeRevID
                                FROM SPMAT_REQData
                                WHERE MaterialID = @MaterialID";
                                    using (SqlCommand copyCmd = new SqlCommand(copyQuery, conn))
                                    {
                                        copyCmd.Parameters.AddWithValue("@MaterialID", materialID);
                                        copyCmd.ExecuteNonQuery();
                                    }

                                    // Update REQData with values from Holding
                                    string updateReqQuery = @"
                                UPDATE SPMAT_REQData
                                SET Discipline = @Discipline,
                                    Area = @Area,
                                    Unit = @Unit,
                                    Phase = @Phase,
                                    Const_Area = @Const_Area,
                                    ISO = @ISO,
                                    Ident_no = @Ident_no,
                                    qty = @qty,
                                    qty_unit = @qty_unit,
                                    Fabrication_Type = @Fabrication_Type,
                                    Spec = @Spec,
                                    IsoRevisionDate = @IsoRevisionDate,
                                    IsoRevision = @IsoRevision,
                                    IsLocked = @IsLocked,
                                    Code = @Code,
                                    Checked = 0,
                                    Moved = 0,
                                    MovedDate = NULL,
                                    Deleted=0,
                                    IsoUniqeRevID= @IsoUniqeRevID
                                WHERE MaterialID = @MaterialID";
                                    using (SqlCommand updateCmd = new SqlCommand(updateReqQuery, conn))
                                    {
                                        updateCmd.Parameters.AddWithValue("@MaterialID", materialID);
                                        updateCmd.Parameters.AddWithValue("@Discipline", reader["Discipline"]);
                                        updateCmd.Parameters.AddWithValue("@Area", reader["Area"]);
                                        updateCmd.Parameters.AddWithValue("@Unit", reader["Unit"]);
                                        updateCmd.Parameters.AddWithValue("@Phase", reader["Phase"]);
                                        updateCmd.Parameters.AddWithValue("@Const_Area", reader["Const_Area"]);
                                        updateCmd.Parameters.AddWithValue("@ISO", reader["ISO"]);
                                        updateCmd.Parameters.AddWithValue("@Ident_no", reader["Ident_no"]);
                                        updateCmd.Parameters.AddWithValue("@qty", reader["qty"]);
                                        updateCmd.Parameters.AddWithValue("@qty_unit", reader["qty_unit"]);
                                        updateCmd.Parameters.AddWithValue("@Fabrication_Type", reader["Fabrication_Type"]);
                                        updateCmd.Parameters.AddWithValue("@Spec", reader["Spec"]);
                                        updateCmd.Parameters.AddWithValue("@IsoRevisionDate", reader["IsoRevisionDate"]);
                                        updateCmd.Parameters.AddWithValue("@IsoRevision", reader["IsoRevision"]);
                                        updateCmd.Parameters.AddWithValue("@IsLocked", reader["IsLocked"]);
                                        updateCmd.Parameters.AddWithValue("@Code", reader["Code"]);
                                        updateCmd.Parameters.AddWithValue("@IsoUniqeRevID", reader["IsoUniqeRevID"]);
                                        
                                        updateCmd.ExecuteNonQuery();
                                    }

                                    reader.Close();

                                    // Delete from Holding using HoldingID
                                    string deleteHoldingQuery = "DELETE FROM SPMAT_REQData_Holding WHERE HoldingID = @HoldingID";
                                    using (SqlCommand deleteCmd = new SqlCommand(deleteHoldingQuery, conn))
                                    {
                                        deleteCmd.Parameters.AddWithValue("@HoldingID", holdingID);
                                        deleteCmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }

        internal static List<SPMATDeletedData> GetMaintenanceData(string projid)
        {
            List<SPMATDeletedData> deldata = new List<SPMATDeletedData>();
            using (SqlConnection cn = new SqlConnection(conMat))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("sp_GetMaintenanceData", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@ProjectID", int.Parse(projid));
                    cn.Open();
                    SqlDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        SPMATDeletedData del = new SPMATDeletedData
                        {
                            DelID = int.Parse(dr[0].ToString()),
                            MaterialID = int.Parse(dr[1].ToString()),
                            Discipline = dr[2].ToString(),
                            Area = dr[3].ToString(),
                            Unit = dr[4].ToString(),
                            Phase = dr[5].ToString(),
                            Const_Area = dr[6].ToString(),
                            ISO = dr[7].ToString(),
                            Ident_no = dr[8].ToString(),
                            qty = DecParse(dr[9].ToString().Trim()).Value,
                            qty_unit = dr[10].ToString(),
                            Fabrication_Type = dr[11].ToString(),
                            Spec = dr[12].ToString(),
                            IsoRevisionDate = dr[13].ToString(),
                            IsoRevision = dr[14].ToString(),
                            IsLocked = dr[15].ToString(),
                            Code = dr[16].ToString(),
                            ImportStatus = dr[17].ToString(),
                            Changes = dr[18].ToString(),
                            MTOID = dr[19] != DBNull.Value ? int.Parse(dr[19].ToString()) : 0
                        };

                        if (!deldata.Contains(del))
                            deldata.Add(del);
                    }
                    dr.Close();
                }
                catch (Exception ex)
                {
                    var e = ex.Message;
                }
                finally
                {
                    if (cn.State == ConnectionState.Open)
                        cn.Close();
                    GC.Collect();
                }
            }
            return deldata;
        }

        internal class DownloadFile
        {
            public string filename { get; set; }
            public string contenttype { get; set; }
            public byte[] filedata { get; set; }
        }
        public static void InsertIntoSPMAT_REQData_Temp(SPMATDBData data)
        {
            using (SqlConnection conn = new SqlConnection(conMat))
            {

                string insertQuery = @"
IF NOT EXISTS (
    SELECT 1 FROM [dbo].[SPMAT_REQData_Temp]
    WHERE [MaterialID] = @MaterialID AND [ProjectID] = @ProjectID AND [Discipline] = @Discipline
      AND [Area] = @Area AND [Unit] = @Unit AND [Phase] = @Phase AND [Const_Area] = @Const_Area
      AND [ISO] = @ISO AND [Ident_no] = @Ident_no AND [qty] = @qty AND [qty_unit] = @qty_unit
      AND [Fabrication_Type] = @Fabrication_Type AND [Spec] = @Spec
      AND [IsoRevisionDate] = @IsoRevisionDate AND [IsoRevision] = @IsoRevision
      AND [IsLocked] = @IsLocked AND [Code] = @Code AND IsoUniqeRevID=@IsoUniqeRevID
)
BEGIN
    INSERT INTO [dbo].[SPMAT_REQData_Temp]
    ([MaterialID], [ProjectID], [Discipline], [Area], [Unit], [Phase], [Const_Area], [ISO],
     [Ident_no], [qty], [qty_unit], [Fabrication_Type], [Spec], [IsoRevisionDate], [IsoRevision],
     [IsLocked], [Code], [Checked], [Moved], [MovedDate],IsoUniqeRevID)
    VALUES
    (@MaterialID, @ProjectID, @Discipline, @Area, @Unit, @Phase, @Const_Area, @ISO,
     @Ident_no, @qty, @qty_unit, @Fabrication_Type, @Spec, @IsoRevisionDate, @IsoRevision,
     @IsLocked, @Code, @Checked, @Moved, @MovedDate,@IsoUniqeRevID)
END";



                SqlCommand cmd = new SqlCommand(insertQuery, conn);
                cmd.Parameters.AddWithValue("@MaterialID", data.MaterialID);
                cmd.Parameters.AddWithValue("@ProjectID", data.ProjectID);
                cmd.Parameters.AddWithValue("@Discipline", data.Discipline ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Area", data.Area ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Unit", data.Unit ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Phase", data.Phase ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Const_Area", data.Const_Area ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@ISO", data.ISO ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Ident_no", data.Ident_no ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@qty", DecParse(data.qty).Value);
                cmd.Parameters.AddWithValue("@qty_unit", data.qty_unit ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Fabrication_Type", data.Fabrication_Type ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Spec", data.Spec ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@IsoRevisionDate", data.IsoRevisionDate ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@IsoRevision", data.IsoRevision ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@IsLocked", data.Lock ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Code", data.Code ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@Checked", false);
                cmd.Parameters.AddWithValue("@Moved", false);
                cmd.Parameters.AddWithValue("@MovedDate", DBNull.Value);
                cmd.Parameters.AddWithValue("@IsoUniqeRevID", data.IsoUniqeRevID );
                

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        public static int GetNextNegativeMaterialID(int projectID)
        {
            int nextID = -1; // Default starting point if no records exist
            using (var cn = new SqlConnection(conMat))
            {
                string query = @"
            SELECT MIN(MaterialID)
            FROM [dbo].[SPMAT_REQData_Temp]
            WHERE ProjectID = @ProjectID AND MaterialID < 0";

                SqlCommand cmd = new SqlCommand(query, cn);
                cmd.Parameters.AddWithValue("@ProjectID", projectID);
                cn.Open();
                var result = cmd.ExecuteScalar();
                if (result != DBNull.Value && result != null)
                {
                    nextID = Convert.ToInt32(result) - 1;
                }
            }
            return nextID;
        }

        internal static void RemoveFromFinal(string MTOID, string ISO)
        {


            using (SqlConnection conn = new SqlConnection(conMat))
            {
                conn.Open();

                // Insert into SPMAT_FIleExports
                string UpdateQuery = @"
             Update [dbo].[SPMAT_REQData] set Checked = 0, Moved = 0,Deleted=0 where ISO in (@ISO)
            delete from[SPMAT_REQData_Temp] where ISO in (@ISO)
            delete from[SPMAT_IntrimData] where ISO in (@ISO)
            Delete from[SPMAT_MTOData] where ISO in (@ISO)";
                using (SqlCommand UpdateCmd = new SqlCommand(UpdateQuery, conn))
                {
                    UpdateCmd.Parameters.AddWithValue("@ISO", ISO);
                    UpdateCmd.ExecuteNonQuery();
                }
            }
        }

        #endregion
        #endregion
        #region ACCESS
        public class ISOData
        {
            public string Drawing_Number { get; set; }
            public string Revision { get; set; }
            public DateTime? RevisionDate { get; set; }
            public bool IsoLock { get; set; }
        }
        internal static List<ISOData> GetIsoData(object SourceAcessDB)
        {
            List<ISOData> isodata = new List<ISOData>();
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SourceAcessDB + ";";
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                OleDbDataReader reader = null;
                //                var cmdtext = "SELECT ISODATA.DRAWING_NUMBER, ISODATA.HAZ_CAT,  qry_s3d_isometric_str_Attribute.oid FROM qry_s3d_isometric_str_Attribute RIGHT JOIN ISODATA ON qry_s3d_isometric_str_Attribute.DRAWING_NUMBER = ISODATA.DRAWING_NUMBER GROUP BY ISODATA.DRAWING_NUMBER, ISODATA.HAZ_CAT, qry_s3d_isometric_str_Attribute.oid HAVING(((ISODATA.DRAWING_NUMBER)= 'DR-056-3137_001'));";
                var cmdtext = "SELECT DISTINCT ISODATA.DRAWING_NUMBER, [Revision History].Date, [Revision History].REVISION, [Production History].[Ready for AFC] FROM (ISODATA INNER JOIN [Revision History] ON ISODATA.ID = [Revision History].ID) INNER JOIN [Production History] ON ISODATA.ID = [Production History].ID ORDER BY ISODATA.DRAWING_NUMBER, [Revision History].Date DESC;";
                OleDbCommand command = new OleDbCommand(cmdtext, connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    ISOData iso = new ISOData();
                    iso.Drawing_Number = reader[0].ToString();
                    if (string.IsNullOrEmpty(reader[1].ToString()))
                    {
                        iso.RevisionDate = DateTime.MinValue;
                    }
                    else
                    {
                        iso.RevisionDate = DateTime.Parse(reader[1].ToString());
                    }
                    iso.Revision = reader[2].ToString();
                    if (string.IsNullOrEmpty(reader[3].ToString()) && iso.RevisionDate == DateTime.MinValue)
                    {
                        iso.IsoLock = false;
                    }
                    else
                    {
                        iso.IsoLock = true;
                    }
                    isodata.Add(iso);

                }
            }
            return isodata;
        }

        internal static string GetIsoAccess(string projid)
        {
            string accesspath = "";
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT  top 1 [IsoControlPath]  FROM [dbo].[SPMAT_ISOControl] where ProjectID=" + projid.Trim();
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        accesspath = dr[0].ToString();
                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return accesspath;
        }

        internal static void RefreshISO(ISOData i)
        {
            using (SqlConnection cn = new SqlConnection(conMat))
            {
                try
                {
                    string query = @"
IF EXISTS (
    SELECT 1 FROM [dbo].[SPMAT_REQData]
    WHERE ISO = @ISO
      AND (IsoRevisionDate <> @IsoRevisionDate OR IsoRevision <> @IsoRevision OR IsLocked <> @IsLocked)
)
BEGIN
    -- Move to Deleted table
    INSERT INTO [dbo].[SPMAT_REQData_Deleted]
    (MaterialID, ProjectID, Discipline, Area, Unit, Phase, Const_Area, ISO, Ident_no, qty, qty_unit,
     Fabrication_Type, Spec, IsoRevisionDate, IsoRevision, IsLocked, Code, Checked, Moved, MovedDate, Processed)
    SELECT MaterialID, ProjectID, Discipline, Area, Unit, Phase, Const_Area, ISO, Ident_no, qty, qty_unit,
           Fabrication_Type, Spec, IsoRevisionDate, IsoRevision, IsLocked, Code, Checked, Moved, MovedDate, 0
    FROM [dbo].[SPMAT_REQData]
    WHERE ISO = @ISO;

    -- Update the record
    UPDATE [dbo].[SPMAT_REQData]
    SET IsoRevisionDate = @IsoRevisionDate,
        IsoRevision = @IsoRevision,
        IsLocked = @IsLocked,
        Checked = 0,
        Moved = 0,
        MovedDate = NULL
    WHERE ISO = @ISO;
END";

                    using (SqlCommand cmd = new SqlCommand(query, cn))
                    {
                        cmd.Parameters.AddWithValue("@ISO", i.Drawing_Number ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@IsoRevisionDate", i.RevisionDate ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@IsoRevision", i.Revision ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@IsLocked", i.IsoLock);

                        cn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Error in RefreshISO: " + ex.Message);
                }
            }
        }

        public static void CleanRecords(string projectId, List<int> finalMTOIDs, List<int> materialIDs)
        {
            using (var conn = new SqlConnection(conMat))
            using (var cmd = new SqlCommand("CleanRecords", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@ProjectID", projectId);

                var finalMTOParam = new SqlParameter("@FinalMTOIDs", SqlDbType.Structured)
                {
                    TypeName = "IntList",
                    Value = ConvertToDataTable(finalMTOIDs)
                };
                cmd.Parameters.Add(finalMTOParam);

                var materialIDParam = new SqlParameter("@MaterialIDs", SqlDbType.Structured)
                {
                    TypeName = "IntList",
                    Value = ConvertToDataTable(materialIDs)
                };
                cmd.Parameters.Add(materialIDParam);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }


        internal static List<int> GetOtherFileIDs(string projid)
        {
            List<int> fileids = new List<int>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT Distinct [FinalFileID]  from [dbo].[SPMAT_FileExports] where FinalFileID in (Select  distinct FIleID FROM [dbo].[SPMAT_MTOData] where ProjectID = " + projid + " and FileID is not null) and [FileCompleted]=1 ";


                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        var fid= int.Parse(dr[0].ToString());
                        if (!fileids.Contains(fid))
                        {
                            fileids.Add(fid);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return fileids;
        }



        public static List<int> GetMTOIDsToDelete(string projectId, List<int> finalMTOIDs, List<int> otherFileIDs)
        {
            var mtoIDsToDelete = new List<int>();

            using (var conn = new SqlConnection(conMat))
            using (var cmd = new SqlCommand("GetMTOIDsToDelete", conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@ProjectID", projectId);

                var finalMTOParam = new SqlParameter("@FinalMTOIDs", SqlDbType.Structured)
                {
                    TypeName = "IntList",
                    Value = ConvertToDataTable(finalMTOIDs)
                };
                cmd.Parameters.Add(finalMTOParam);

                var otherFileParam = new SqlParameter("@OtherFileIDs", SqlDbType.Structured)
                {
                    TypeName = "IntList",
                    Value = ConvertToDataTable(otherFileIDs)
                };
                cmd.Parameters.Add(otherFileParam);

                conn.Open();
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        mtoIDsToDelete.Add(reader.GetInt32(0));
                    }
                }
            }

            return mtoIDsToDelete;
        }
        public static List<(int MaterialID, int MTOID)> GetObsoleteMTOs(string projectId)
        {
            var result = new List<(int, int)>();
            using (SqlConnection conn = new SqlConnection(conMat))
            {
                using (SqlCommand cmd = new SqlCommand("dbo.GetObsoleteMTOs", conn)) // if wrapped in SP
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@ProjectID", projectId);

                    conn.Open();
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add((reader.GetInt32(0), reader.GetInt32(1)));
                        }
                    }
                }
            }
            return result;
        }


        private static DataTable ConvertToDataTable(List<int> ids)
        {
            var table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            foreach (var id in ids)
            {
                table.Rows.Add(id);
            }
            return table;
        }

        internal static int InsertExportRecord(string projid, List<int> mTOIDs)
        {
            string ids = string.Join(",", mTOIDs);
            int FileID = 0;
            using (SqlConnection conn = new SqlConnection(conMat))
            {
                conn.Open();

                // Insert into SPMAT_FIleExports
                string insertQuery = @"
            INSERT INTO SPMAT_FileExports (FileMTOIDs)
            VALUES (@MTOIDs) SELECT SCOPE_IDENTITY()";
                using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
                {
                    insertCmd.Parameters.AddWithValue("@MTOIDs", ids);
                    FileID = Convert.ToInt32(insertCmd.ExecuteScalar());

                }

                // Update SPMAT_IntrimData
                string updateQuery = $@"
            UPDATE SPMAT_MTOData
            SET FileID = @FileID
            WHERE MTOID IN ({ids})";
                using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                {
                    updateCmd.Parameters.AddWithValue("@FileID", FileID);
                    updateCmd.ExecuteNonQuery();
                }
            }
            return FileID;
        }

        internal static void UpdateExportRecord(int fileID, string fileName)
        {
            using (SqlConnection conn = new SqlConnection(conMat))
            {
                conn.Open();
                // Update SPMAT_IntrimData
                string updateQuery = $@"
            UPDATE SPMAT_FileExports
            SET [FinalFileName] = @FileName
            WHERE FinalFileID=@FileID";
                using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                {
                    updateCmd.Parameters.AddWithValue("@FileID", fileID);
                    updateCmd.Parameters.AddWithValue("@FileName", fileName);
                    updateCmd.ExecuteNonQuery();
                }
            }
        }

        internal static List<DDLList> GetProjectISORevData(string projid)
        {
            List<DDLList> isolst = new List<DDLList>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = "SELECT Distinct [ISO],[ISO]+'- Rev: '+ [IsoRevision] + ' : Records ( '+convert(varchar(20) ,count(IsoRevision))+' )' FROM [dbo].SPMAT_MTOData where ProjectID=" + projid + " and IsDeleted=0 group by ISO,IsoRevision order by 1";
               
                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        DDLList iso = new DDLList();
                        iso.DDLList_ID = dr[0].ToString();
                        iso.DDLListName = dr[1].ToString();

                        if (!isolst.Contains(iso))
                        {
                            isolst.Add(iso);
                        }
                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return isolst;
        }

        internal static List<IsoRevisionData> GetIsoReviewMTOData(string isosheet, string projid)
        {
            List<IsoRevisionData> isodata = new List<IsoRevisionData>();
            System.Data.SqlClient.SqlConnection cn = null;
            try
            {
                string query = $@"
    SELECT [MTOID],[MaterialID],[ProjectID],[ISO],[Ident_no],[qty],[qty_unit],[Fabrication_Type],[ReleasedMaterial]
    FROM [dbo].[SPMAT_MTOData]
    WHERE LTRIM(RTRIM(ISO)) = '{isosheet.Trim()}' AND IsDeleted=0 AND ProjectID = {projid.Trim()}";

                using (cn = new System.Data.SqlClient.SqlConnection(conMat))
                {
                    System.Data.SqlClient.SqlCommand Command = new System.Data.SqlClient.SqlCommand(query, cn);
                    Command.CommandType = System.Data.CommandType.Text;
                    cn.Open();
                    System.Data.SqlClient.SqlDataReader dr = Command.ExecuteReader();
                    while (dr.Read())
                    {
                        IsoRevisionData spdata = new IsoRevisionData();
                        spdata.MTOID = Convert.ToInt32(dr[0]);
                        spdata.MaterialID = Convert.ToInt32(dr[1]);
                        spdata.ProjectID = Convert.ToInt32(dr[2]);
                        spdata.ISO = dr[3].ToString();
                        spdata.Ident_no = dr[4].ToString();
                        spdata.qty = dr[5].ToString().Trim();
                        spdata.qty_unit = dr[6].ToString();
                        spdata.Fabrication_Type = dr[7].ToString();
                        if (dr[8] != DBNull.Value)
                        {
                            spdata.ReleasedMaterial = bool.Parse(dr[8].ToString());
                        }
                        if (!isodata.Contains(spdata))
                        {
                            isodata.Add(spdata);
                        }

                    }
                    dr.Close();
                    cn.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open)
                {
                    cn.Close();
                    cn = null;
                }
                System.GC.Collect();
            }
            return isodata;
        }

        public static void UpdateReleasedMaterialStatus(int mtoid, bool released)
        {
            using (SqlConnection conn = new SqlConnection(conMat))
            {
                string query = "UPDATE SPMAT_MTOData SET ReleasedMaterial = @Released WHERE MTOID = @MTOID";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Released", released);
                cmd.Parameters.AddWithValue("@MTOID", mtoid);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }


        #endregion
    }
}