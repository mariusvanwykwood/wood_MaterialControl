using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using static Wood_MaterialControl.DataClass;

namespace Wood_MaterialControl
{
    public partial class MainHome : System.Web.UI.Page
    {

        protected void Page_Init(object sender, EventArgs e)
        {
            if (Session["GridData"] != null && Session["CurrentIso"] != null && Session["CurrentArea"] != null)
            {
                var griddata = Session["GridData"] as List<GridData>;
                var area = Session["CurrentArea"].ToString();
                var iso = Session["CurrentIso"].ToString();

                // Rebuild columns only, no data binding
                BindExcelGridView(griddata, area, iso, bindData: false);
            }
        }


        private List<SPMATData> CachedMTOData
        {
            get => Session["MTOData"] as List<SPMATData>;
            set => Session["MTOData"] = value;
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                ExcelGridView.RowCommand += ExcelGridView_RowCommand;
                SetNoCache();
                if (!IsPostBack)
                {
                    var EID = -1;
                    try
                    {
                        EID = int.Parse(Session["EID"].ToString());
                    }
                    catch { }
                    if (EID > -1 && EID != 0)
                    {
                        Session["EID"] = EID;
                    }
                    else
                    {
                        Session.Clear();
                        Response.Redirect("Default.aspx?UF=1");
                    }

                    List<DDLList> clientlist = DataClass.GetAllRefClients();
                    Session["ClientList"] = clientlist;
                    ddlclient.DataSource = clientlist;
                    ddlclient.DataTextField = "DDLListName";
                    ddlclient.DataValueField = "DDLList_ID";
                    ddlclient.DataBind();
                    ddlclient.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- Select --", ""));

                    // Bind grid only on first load
                    if (Session["GridData"] != null)
                    {
                        var griddata = Session["GridData"] as List<GridData>;
                        var area = Session["CurrentArea"].ToString();
                        var iso = Session["CurrentIso"].ToString();
                        BindExcelGridView(griddata, area, iso, bindData: true);
                    }

                    // Always set HttpContext items for templates
                    HttpContext.Current.Items["Units"] = Session["Units"];
                    HttpContext.Current.Items["Phases"] = Session["Phases"];
                    HttpContext.Current.Items["ConstAreas"] = Session["ConstAreas"];

                    var Projid = ddlprojsel.SelectedValue;
                    var data = DataClass.GetMTOData(Projid, true);
                    CachedMTOData = data;
                    gvExported.DataSource = data;
                    gvExported.DataBind();



                }
                else if (IsPostBack)
                {
                    HttpContext.Current.Items["Units"] = Session["Units"];
                    HttpContext.Current.Items["Phases"] = Session["Phases"];
                    HttpContext.Current.Items["ConstAreas"] = Session["ConstAreas"];
                    string eventTarget = Request["__EVENTTARGET"];
                    string clickedButton = Request.Form["btnAddNew"] ?? Request.Form["btnSaveNew"] ?? Request.Form["btnCancelAdd"];

                    string changedDropdown =
                     Request.Form["ddlAddUnit"] ??
                     Request.Form["ddlAddPhase"] ??
                     Request.Form["ddlAddSpec"] ??
                     Request.Form["ddlAddShortcode"];

                    if ((!string.IsNullOrEmpty(clickedButton) || !string.IsNullOrEmpty(eventTarget) || !string.IsNullOrEmpty(changedDropdown)) &&
                        (
                        eventTarget.Contains("ddlUnit") || eventTarget.Contains("ddlSpec") || eventTarget.Contains("ddlShortcode") || eventTarget.Contains("ddlPhase") || eventTarget.Contains("ddlConstArea") || eventTarget.Contains("ddlIdent") || eventTarget.Contains("chkChecked") || eventTarget.Contains("txtQty") ||
                        eventTarget.Contains("btnAddNew") || eventTarget.Contains("btnSaveNew") || eventTarget.Contains("btnCancelAdd") || !string.IsNullOrEmpty(clickedButton) || !string.IsNullOrEmpty(changedDropdown)
                         )
                        )
                    {
                        var griddata = Session["GridData"] as List<GridData>;
                        var area = Session["CurrentArea"].ToString();
                        var iso = Session["CurrentIso"].ToString();
                        BindExcelGridView(griddata, area, iso, bindData: true);
                    }



                }
            }
            catch
            {
            }
        }
        private void BindGrid(DataTable dt)
        {
            gvExported.DataSource = dt;
            gvExported.DataBind();
        }

        protected void ExcelGridView_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "DeleteRow")
            {
                int rowIndex = Convert.ToInt32(e.CommandArgument);
                int materialID = Convert.ToInt32(ExcelGridView.DataKeys[rowIndex].Value);
                // Confirm deletion
                ClientScript.RegisterStartupScript(this.GetType(), "confirmDelete",
                    $"if(confirm('Are you sure you want to delete this item?')) {{ __doPostBack('{ExcelGridView.UniqueID}', 'DeleteConfirmed${rowIndex}'); }}", true);
            }
            else if (e.CommandName.StartsWith("DeleteConfirmed"))
            {
                int rowIndex = Convert.ToInt32(e.CommandArgument);
                int materialID = Convert.ToInt32(ExcelGridView.DataKeys[rowIndex].Value);
                var griddata = Session["GridData"] as List<GridData> ?? new List<GridData>();
                if (griddata.Where(x => x.MaterialID == materialID).Any())
                {
                    var itemToremove = griddata.FirstOrDefault(x => x.MaterialID == materialID);

                    if (itemToremove != null)
                    {
                        griddata.Remove(itemToremove);
                    }

                }
                Session["GridData"] = griddata;
                // Mark as Checked in DB
                DataClass.MarkMaterialAsChecked(materialID);
                DataClass.DeleteTempMaterial(materialID);


                // Reload grid
                // btnloadiso_Click(null, null);
                ForceButtons("btnloadiso", false);
            }
        }

        private void SetNoCache()
        {
            HttpContext.Current.Response.Cache.SetAllowResponseInBrowserHistory(false);
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.Cache.SetNoStore();
            Response.Cache.SetExpires(DateTime.Now);
            Response.Cache.SetValidUntilExpires(true);
        }

        protected void ddlclient_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlclient.SelectedValue != "")
            {
                //Get Data for Iso DLL
                List<DDLList> projects = DataClass.LoadProjectsSpecs(ddlclient.SelectedValue.ToString().Trim());
                ddlprojects.DataSource = projects;
                ddlprojects.DataTextField = "DDLListName";
                ddlprojects.DataValueField = "DDLList_ID";
                ddlprojects.DataBind();
                ddlprojects.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- Select --", ""));
            }
        }
        protected void ddlprojects_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlprojects.SelectedValue != "")
            {
                List<DDLList> clientlist = (List<DDLList>)Session["ClientList"];
                var Value = clientlist.Where(x => x.DDLList_ID == ddlprojects.SelectedValue).Select(x => x.DDLID).First();
                List<DDLList> subprojects = DataClass.GetRefProjects(Value.Trim());
                ddlprojsel.DataSource = subprojects;
                ddlprojsel.DataTextField = "DDLListName";
                ddlprojsel.DataValueField = "DDLList_ID";
                ddlprojsel.DataBind();
                ddlprojsel.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- Select --", ""));
            }
        }

        protected void ddlprojsel_SelectedIndexChanged(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            if (ddlprojsel.SelectedValue != "")
            {
                //Get Data for Iso DLL
                var projid = ddlprojsel.SelectedValue;
                Session["ENGPROJECTID"] = projid;
                List<DDLList> iso = DataClass.GetProjectISO(projid, false);
                if (iso.Count > 0)
                {
                    ddliso.DataSource = iso;
                    ddliso.DataTextField = "DDLListName";
                    ddliso.DataValueField = "DDLList_ID";
                    ddliso.DataBind();
                    // btnviewspmat.CssClass = "shown";
                    btnviewInterim.CssClass = "shown";
                    btnViewFinal.CssClass = "shown";
                }
                else
                {
                    diverror.Style["display"] = "block";
                    lblerror.Text = "No Iso in correct state.";
                    lblerror.ForeColor = System.Drawing.Color.Red;
                    return;
                }
            }

        }

        protected void btnloadiso_Click(object sender, EventArgs e)
        {
            ClearSession();
            ExcelGridView.DataSource = null;
            lblerror.Text = "";
            diverror.Style["display"] = "none";
            if (string.IsNullOrEmpty(ddlprojects.SelectedValue) ||
             string.IsNullOrEmpty(ddlprojsel.SelectedValue) ||
             string.IsNullOrEmpty(ddliso.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select a Spec, Project and ISO before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
            pnlExcelUpload.Style["display"] = "block";
            var spec = DataClass.LoadSpectDataFromDB(ddlprojects.SelectedValue.ToString().Trim());
            Session["LoadedSpecs"] = spec;
            var isosheet = ddliso.SelectedValue;
            var projid = ddlprojsel.SelectedValue;
            var SelectedISO = ddliso.SelectedItem.Text;

            string isorev = "";
            var match = System.Text.RegularExpressions.Regex.Match(SelectedISO, @"\bRev:\s*([^\s:]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success)
            {
                isorev = match.Groups[1].Value.Trim(); // e.g., "0", "0.1", "S0", "R0", etc.
            }
            List<SPMATDBData> sptest = DataClass.GetIsoSheetMTOData(isosheet, projid, isorev, false);
            if (sptest == null || sptest.Count <= 0)
            {
                sptest = DataClass.GetPreviousIsoSheetMTOData(isosheet, projid, isorev);
                if (sptest != null && sptest.Count > 0)
                {
                    Session["IsPreviousData"] = true;
                }
            }
            if (sptest.Count > 0)
            {
                var area = sptest.Select(x => x.Area).Distinct().First().ToString();
                var disipline = sptest.Select(x => x.Discipline).Distinct().First().ToString();

                var distinctSpecslist = spec.GroupBy(s => new { s.Lineclass, s.Ident }).Select(g => g.First()).ToList();

                var griddata = PopulateGridData(sptest, distinctSpecslist).Where(x => x.ISO == isosheet).ToList();
                var dbdata = griddata.Select(item => new GridData
                {
                    MaterialID = item.MaterialID,
                    ProjectID = item.ProjectID,
                    Discipline = item.Discipline,
                    Area = item.Area,
                    Unit = item.Unit,
                    Phase = item.Phase,
                    Const_Area = item.Const_Area,
                    ISO = item.ISO,
                    Component_Type = item.Component_Type,
                    Spec = item.Spec,
                    Shortcode = item.Shortcode,
                    Ident_no = item.Ident_no,
                    IsoShortDescription = item.IsoShortDescription,
                    Size_sch1 = item.Size_sch1,
                    Size_sch2 = item.Size_sch2,
                    Size_sch3 = item.Size_sch3,
                    Size_sch4 = item.Size_sch4,
                    Size_sch5 = item.Size_sch5,
                    qty = item.qty,
                    qty_unit = item.qty_unit,
                    Fabrication_Type = item.Fabrication_Type,
                    Source = item.Source,
                    IsoRevision = item.IsoRevision,
                    IsoRevisionDate = item.IsoRevisionDate,
                    IsLocked = item.IsLocked,
                    IsoUniqeRevID = item.IsoUniqeRevID
                }).ToList();

                List<string> distinctUnits = DataClass.GetUnitsByProject(projid);
                List<string> distinctPhases = DataClass.GetPhasesByProject(projid);
                List<string> distinctConstAreas = DataClass.GetConstAreasByProject(projid);

                List<string> distinctSpecs = DataClass.GetSpecsByProject(projid);
                List<string> distinctShortCodes = spec.Select(x => x.Shortcode).Distinct().ToList();
                List<string> distinctIdents = spec.Select(x => x.Ident).Distinct().ToList();

                // Spec → Shortcodes
                var specShortcodeMap = spec.GroupBy(x => x.Lineclass).ToDictionary(g => g.Key, g => g.Select(x => x.Shortcode).Distinct().ToList());

                // Shortcode → Idents
                //Dictionary<string, Dictionary<string, List<string>>>
                var shortCodeIdentMap = spec.GroupBy(x => x.Lineclass).ToDictionary(g => g.Key, g => g.GroupBy(s => s.Shortcode).ToDictionary(sg => sg.Key, sg => sg.Select(s => s.Ident).Distinct().ToList()));

                // Optionally: store combinations for cascading logic
                var unitPhaseMap = DataClass.GetUnitPhaseMap(projid);

                var unitConstAreaMap = DataClass.GetUnitPhaseConstAreaMap(projid);
                Session["GridData"] = griddata;
                Session["OriginalDBData"] = dbdata;
                Session["CurrentIso"] = isosheet;
                Session["CurrentArea"] = area;
                Session["CurrentDisipline"] = disipline;
                Session["Units"] = distinctUnits;
                Session["Phases"] = distinctPhases;
                Session["ConstAreas"] = distinctConstAreas;
                Session["UnitPhaseMap"] = unitPhaseMap;
                Session["UnitConstAreaMap"] = unitConstAreaMap;
                Session["Specs"] = distinctSpecs;
                Session["ShortCodes"] = distinctShortCodes;
                Session["Idents"] = distinctIdents;
                Session["SpecShortCodeMap"] = specShortcodeMap;
                Session["ShortCodeIdentMap"] = shortCodeIdentMap;

                HttpContext.Current.Items["Units"] = distinctUnits;
                HttpContext.Current.Items["Phases"] = distinctPhases;
                HttpContext.Current.Items["ConstAreas"] = distinctConstAreas;
                HttpContext.Current.Items["Specs"] = distinctSpecs;
                HttpContext.Current.Items["ShortCodes"] = distinctShortCodes;
                HttpContext.Current.Items["Idents"] = distinctIdents;

                BindExcelGridView(griddata, area, isosheet, bindData: true);
                btnSubmit.Style["display"] = "block";
                btnAddNew.Style["display"] = "block";
            }
            else
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "No Current or Previous data found for the selected ISO.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
        }

        private void ClearSession()
        {
            Session["GridData"] = null;
            Session["OriginalDBData"] = null;
            Session["CurrentIso"] = null;
            Session["CurrentArea"] = null;
            Session["CurrentDisipline"] = null;
            Session["Units"] = null;
            Session["Phases"] = null;
            Session["ConstAreas"] = null;
            Session["UnitPhaseMap"] = null;
            Session["UnitConstAreaMap"] = null;
            Session["Specs"] = null;
            Session["ShortCodes"] = null;
            Session["Idents"] = null;
            Session["SpecShortCodeMap"] = null;
            Session["ShortCodeIdentMap"] = null;
        }

        public List<GridData> PopulateGridData(List<SPMATDBData> spmatList, List<SpecData> specList)
        {
            var gridDataList = (from spmat in spmatList
                                join spec in specList
                                on new { Spec = spmat.Spec, Ident = spmat.Ident_no }
                                equals new { Spec = spec.Lineclass, Ident = spec.Ident }
                                select new GridData
                                {
                                    MaterialID = spmat.MaterialID,
                                    ProjectID = spmat.ProjectID,
                                    Discipline = spmat.Discipline,
                                    Area = spmat.Area,
                                    Unit = spmat.Unit,
                                    Phase = spmat.Phase,
                                    Const_Area = spmat.Const_Area,
                                    ISO = spmat.ISO,
                                    Component_Type = spec.Commodity_code, // or spec.Short_desc
                                    Spec = spmat.Spec,
                                    Shortcode = spec.Shortcode,
                                    Ident_no = spmat.Ident_no,
                                    IsoShortDescription = spec.Description,
                                    Size_sch1 = spec.Size_sch1,
                                    Size_sch2 = spec.Size_sch2,
                                    Size_sch3 = spec.Size_sch3,
                                    Size_sch4 = spec.Size_sch4,
                                    Size_sch5 = spec.Size_sch5,
                                    qty = spmat.qty.ToString(),
                                    qty_unit = spmat.qty_unit,
                                    Fabrication_Type = spmat.Fabrication_Type,
                                    Source = spmat.Code,
                                    IsoRevision = spmat.IsoRevision,
                                    IsoRevisionDate = spmat.IsoRevisionDate,
                                    IsLocked = spmat.Lock,
                                    IsoUniqeRevID = spmat.IsoUniqeRevID

                                }).ToList();

            return gridDataList;
        }

        public DataTable ConvertToDataTable(List<GridData> gridDataList)
        {


            DataTable table = new DataTable();

            // Only use properties from GridData
            var properties = typeof(GridData).GetProperties();

            foreach (var prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (var item in gridDataList)
            {
                var row = table.NewRow();
                foreach (var prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;

        }

        public DataTable ConvertSPMATDataToDataTable(List<SPMATData> dataList)
        {
            DataTable table = new DataTable();

            // Define columns based on SPMATData properties
            table.Columns.Add("MTOID", typeof(int));
            table.Columns.Add("Discipline", typeof(string));
            table.Columns.Add("Area", typeof(string));
            table.Columns.Add("Unit", typeof(string));
            table.Columns.Add("Phase", typeof(string));
            table.Columns.Add("Const_Area", typeof(string));
            table.Columns.Add("ISO", typeof(string));
            table.Columns.Add("Ident_no", typeof(string));
            table.Columns.Add("qty", typeof(decimal));
            table.Columns.Add("qty_unit", typeof(string));
            table.Columns.Add("Fabrication_Type", typeof(string));
            table.Columns.Add("Spec", typeof(string));
            table.Columns.Add("Pos", typeof(string));
            table.Columns.Add("IsoRevisionDate", typeof(string));
            table.Columns.Add("IsoRevision", typeof(string));
            table.Columns.Add("IsLocked", typeof(string));
            table.Columns.Add("Code", typeof(string));
            table.Columns.Add("ImportStatus", typeof(string));
            table.Columns.Add("IsoUniqeRevID", typeof(int));

            foreach (var item in dataList)
            {
                table.Rows.Add(
                    item.MTOID,
                    item.Discipline,
                    item.Area,
                    item.Unit,
                    item.Phase,
                    item.Const_Area,
                    item.ISO,
                    item.Ident_no,
                    item.qty,
                    item.qty_unit,
                    item.Fabrication_Type,
                    item.Spec,
                    item.Pos,
                    item.IsoRevisionDate,
                    item.IsoRevision,
                    item.IsLocked,
                    item.Code,
                    item.ImportStatus,
                    item.IsoUniqeRevID
                );
            }
            return table;
        }

        public List<SPMATData> ConvertToSPMATDataList(DataTable table)
        {
            var list = new List<SPMATData>();

            foreach (DataRow row in table.Rows)
            {
                var item = new SPMATData
                {
                    Discipline = row["Discipline"]?.ToString(),
                    Area = row["Area"]?.ToString(),
                    Unit = row["Unit"]?.ToString(),
                    Phase = row["Phase"]?.ToString(),
                    Const_Area = row["Const_Area"]?.ToString(),
                    ISO = row["ISO"]?.ToString(),
                    Ident_no = row["Ident_no"]?.ToString(),
                    qty = row["qty"] != DBNull.Value ? Convert.ToDecimal(row["qty"]) : 0,
                    qty_unit = row["qty_unit"]?.ToString(),
                    Fabrication_Type = row["Fabrication_Type"]?.ToString(),
                    Spec = row["Spec"]?.ToString(),
                    Pos = row["Pos"]?.ToString(),
                    IsoRevisionDate = row["IsoRevisionDate"]?.ToString(),
                    IsoRevision = row["IsoRevision"]?.ToString(),
                    IsLocked = row["IsLocked"]?.ToString(),
                    IsoUniqeRevID = Convert.ToInt32(row["IsoUniqeRevID"].ToString())

                };

                list.Add(item);
            }

            return list;
        }

        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            bool allChecked = true;
            List<SPMATDBData> updatedItems = new List<SPMATDBData>();
            var SelectedISO = ddliso.SelectedItem.Text;

            string isorev = "";
            var match = System.Text.RegularExpressions.Regex.Match(SelectedISO, @"\bRev:\s*([^\s:]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success)
            {
                isorev = match.Groups[1].Value.Trim(); // e.g., "0", "0.1", "S0", "R0", etc.
            }


            var loadedSpecs = Session["LoadedSpecs"] as List<SpecData>;
            var griddata = Session["GridData"] as List<GridData>;
            var dbdata = Session["OriginalDBData"] as List<GridData>;
            var area = Session["CurrentArea"]?.ToString();
            var disipline = Session["CurrentDisipline"]?.ToString();
            var iso = Session["CurrentIso"]?.ToString();
            BindExcelGridView(griddata, area, iso, bindData: true);
            var chklist = Session["CheckedMaterialIDs"] as List<int> ?? new List<int>();
            var changedUnits = Session["ChangedUnits"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var changedPhases = Session["ChangedPhases"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var changedConstAreas = Session["ChangedConstAreas"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var changedSpecs = Session["ChangedSpecs"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var changedShortcodes = Session["ChangedShortcodes"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var changedIdents = Session["ChangedIdents"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var changedText = Session["ChangedText"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            foreach (GridViewRow row in ExcelGridView.Rows)
            {
                int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
                if (!chklist.Contains(materialID))
                {
                    allChecked = false;
                    break;
                }
                // Retrieve controls
                var txtQty = row.FindControl("txtQty") as System.Web.UI.WebControls.TextBox;
                var ddlUnit = row.FindControl("ddlUnit") as DropDownList;
                var ddlPhase = row.FindControl("ddlPhase") as DropDownList;
                var ddlConstArea = row.FindControl("ddlConstArea") as DropDownList;
                var ddlSpec = row.FindControl("ddlSpec") as DropDownList;
                var ddlShortcode = row.FindControl("ddlShortcode") as DropDownList;
                var ddlIdent = row.FindControl("ddlIdent") as DropDownList;

                //if (txtQty == null || !decimal.TryParse(txtQty.Text, out decimal newQty))
                //    continue;

                string selectedText = changedText.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? txtQty?.Text;
                string selectedUnit = changedUnits.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlUnit?.SelectedValue;
                string selectedPhase = changedPhases.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlPhase?.SelectedValue;
                string selectedConstArea = changedConstAreas.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlConstArea?.SelectedValue;
                string selectedSpec = changedSpecs.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlSpec?.SelectedValue;
                string selectedShortcode = changedShortcodes.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlShortcode?.SelectedValue;
                string selectedIdent = changedIdents.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlIdent?.SelectedValue;

                if (DecParse(selectedText).HasError)
                {
                    lblMessage.Text = "Please enter a valid quantity.";
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    return;
                }

                string qtyUnit = "pc"; // default

                if (loadedSpecs != null && !string.IsNullOrEmpty(selectedShortcode))
                {
                    var matchedSpec = loadedSpecs.FirstOrDefault(s => s.Shortcode == selectedShortcode);
                    if (matchedSpec != null && matchedSpec.Shortcode == "PIP")
                        qtyUnit = "m";
                }
                //compare to original, if no changes keep iso and code
                var orgdbdata = dbdata?.FirstOrDefault(g => g.MaterialID == materialID);
                var originalData = griddata?.FirstOrDefault(g => g.MaterialID == materialID);
                var havediffs = false;
                if (orgdbdata != null && originalData != null)
                {
                    havediffs = AreObjectsDifferentExcept(orgdbdata, originalData);
                }
                else if (orgdbdata == null)
                {
                    havediffs = true;
                }
                if (originalData != null)
                {
                    var orgToSPMAT = new SPMATDBData
                    {
                        MaterialID = originalData.MaterialID,
                        ProjectID = originalData.ProjectID,
                        Discipline = originalData.Discipline,
                        Area = originalData.Area,
                        Unit = originalData.Unit,
                        Phase = originalData.Phase,
                        Const_Area = originalData.Const_Area,
                        ISO = originalData.ISO,
                        Ident_no = originalData.Ident_no,
                        qty = originalData.qty,
                        qty_unit = originalData.qty_unit,
                        Fabrication_Type = originalData.Fabrication_Type,
                        Spec = originalData.Spec,
                        IsoRevisionDate = originalData.IsoRevisionDate,
                        IsoRevision = originalData.IsoRevision,
                        Lock = originalData.IsLocked,
                        Code = originalData.Source,
                        IsoUniqeRevID = originalData.IsoUniqeRevID
                    };

                    var updatedCompare = new SPMATDBData
                    {
                        MaterialID = originalData.MaterialID,
                        ProjectID = originalData.ProjectID,
                        Discipline = string.IsNullOrEmpty(originalData.Discipline) ? disipline : originalData?.Discipline,
                        Area = string.IsNullOrEmpty(originalData.Area) ? area : originalData?.Area,
                        Unit = selectedUnit,
                        Phase = selectedPhase,
                        Const_Area = selectedConstArea,
                        ISO = originalData.ISO,
                        Ident_no = selectedIdent,
                        qty = selectedText,
                        qty_unit = qtyUnit,
                        Fabrication_Type = havediffs ? "Undefined" : orgdbdata?.Fabrication_Type,
                        Spec = havediffs ? selectedSpec : orgdbdata?.Spec,
                        IsoRevision = !String.IsNullOrEmpty(isorev) ? isorev : orgdbdata?.IsoRevision,
                        IsoRevisionDate = havediffs ? "" : orgdbdata?.IsoRevisionDate,
                        Lock = havediffs ? "" : orgdbdata?.IsLocked,
                        Code = havediffs ? "M" : orgdbdata?.Source,
                        IsoUniqeRevID = orgdbdata.IsoUniqeRevID
                    };
                    if (orgToSPMAT != null && updatedCompare != null)
                    {
                        havediffs = AreObjectsDifferentExcept(orgToSPMAT, updatedCompare);
                    }
                    var updated = new SPMATDBData
                    {
                        MaterialID = originalData.MaterialID,
                        ProjectID = originalData.ProjectID,
                        Discipline = string.IsNullOrEmpty(originalData.Discipline) ? disipline : originalData?.Discipline,
                        Area = string.IsNullOrEmpty(originalData.Area) ? area : originalData?.Area,
                        Unit = selectedUnit,
                        Phase = selectedPhase,
                        Const_Area = selectedConstArea,
                        ISO = originalData.ISO,
                        Ident_no = selectedIdent,
                        qty = selectedText,
                        qty_unit = qtyUnit,
                        Fabrication_Type = havediffs ? "Undefined" : orgdbdata?.Fabrication_Type,
                        Spec = havediffs ? selectedSpec : orgdbdata?.Spec,
                        IsoRevision = !String.IsNullOrEmpty(isorev) ? isorev : orgdbdata?.IsoRevision,
                        IsoRevisionDate = havediffs ? "" : orgdbdata?.IsoRevisionDate,
                        Lock = havediffs ? "" : orgdbdata?.IsLocked,
                        Code = havediffs ? "M" : orgdbdata?.Source,
                        IsoUniqeRevID = orgdbdata.IsoUniqeRevID
                    };
                    updatedItems.Add(updated);

                    if (havediffs)
                    {
                        DataClass.InsertIntoSPMAT_REQData_Temp(updated);
                    }
                }

            }

            if (!allChecked)
            {
                lblMessage.Text = "Please check all rows before submitting.";
                lblMessage.ForeColor = System.Drawing.Color.Red;
                return;
            }

            // Update database
            bool IsPreviousData = bool.TryParse(Session["IsPreviousData"]?.ToString(), out var result) && result;

            foreach (var item in updatedItems)
            {
                DataClass.FinalizeMaterialUpdate(item, IsPreviousData);
            }

            lblMessage.ForeColor = System.Drawing.Color.Green;
            lblMessage.Text = "All quantities updated successfully.";
            var eid = 0;
            List<DDLList> clientlist = new List<DDLList>();
            try
            {
                eid = int.Parse(Session["EID"].ToString());
                clientlist = (List<DDLList>)Session["ClientList"];
            }
            catch { }
            Session.Clear();
            Session.Clear();
            Session["EID"] = eid;
            Session["ClientList"] = clientlist;
            ddliso.ClearSelection();
            ExcelGridView.DataSource = null;
            ExcelGridView.DataBind();
            lblFileName.Text = "";
            btnSubmit.Style["display"] = "none";
            btnAddNew.Style["display"] = "none";
            pnlExcelUpload.Style["display"] = "none";
            btnCancelAdd_Click(null, null);
            ReloadIso();
            //btnviewspmat_Click(null, null);
            //ForceButtons("btnviewspmat",true);
            var Projid = ddlprojsel.SelectedValue;
            List<SPMATDBData> mtoData = DataClass.GetWorkingMTOData(Projid);
            Session["SPMATExportData"] = mtoData;
            ForceButtons("btnExportSPMAT", true);
            ForceButtons("btnviewInterim", true);

        }

        private void ReloadIso()
        {
            var projid = ddlprojsel.SelectedValue;
            List<DDLList> iso = DataClass.GetProjectISO(projid, false);
            ddliso.DataSource = iso;
            ddliso.DataTextField = "DDLListName";
            ddliso.DataValueField = "DDLList_ID";
            ddliso.DataBind();
        }

        protected void ExcelGridView_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                var dataItem = (DataRowView)e.Row.DataItem;

                var specList = Session["LoadedSpecs"] as List<SpecData>;

                // Get dropdowns
                DropDownList ddlUnit = (DropDownList)e.Row.FindControl("ddlUnit");
                DropDownList ddlPhase = (DropDownList)e.Row.FindControl("ddlPhase");
                DropDownList ddlConstArea = (DropDownList)e.Row.FindControl("ddlConstArea");
                DropDownList ddlSpec = (DropDownList)e.Row.FindControl("ddlSpec");
                DropDownList ddlShortcode = (DropDownList)e.Row.FindControl("ddlShortcode");
                DropDownList ddlIdent = (DropDownList)e.Row.FindControl("ddlIdent");



                System.Web.UI.WebControls.CheckBox chk = (System.Web.UI.WebControls.CheckBox)e.Row.FindControl("chkChecked");
                System.Web.UI.WebControls.TextBox txtQty = (System.Web.UI.WebControls.TextBox)e.Row.FindControl("txtQty");
                int materialID = Convert.ToInt32(dataItem["MaterialID"]);
                var chklist = Session["CheckedMaterialIDs"] as List<int> ?? new List<int>();
                // Get current values from data
                string currentText = dataItem["qty"]?.ToString();
                string currentUnit = dataItem["Unit"]?.ToString();
                string currentPhase = dataItem["Phase"]?.ToString();
                string currentConstArea = dataItem["Const_Area"]?.ToString();
                string currentSpec = dataItem["Spec"]?.ToString();
                string currentShortcode = dataItem["Shortcode"]?.ToString();
                string currentIdent = dataItem["Ident_no"]?.ToString();
                bool IsChecked = false;
                var rowindex = -1;

                var changedText = Session["ChangedText"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
                var changedUnits = Session["ChangedUnits"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
                var changedPhases = Session["ChangedPhases"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
                var changedConstAreas = Session["ChangedConstAreas"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
                var changedSpecs = Session["ChangedSpecs"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
                var changedShortcodes = Session["ChangedShortcodes"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
                var changedIdents = Session["ChangedIdents"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
                string selectedSpec = "";
                string selectedShortcode = "";
                string selectedIdent = "";
                string selectedUnit = "";
                string selectedPhase = "";
                string selectedConstArea = "";
                string selectedText = "";

                if (Session["rowindex"] != null)
                {
                    rowindex = Convert.ToInt32(Session["rowindex"]);
                }
                else
                {
                    rowindex = e.Row.RowIndex;
                }

                if (rowindex != -1)
                {
                    if (chklist.Contains(materialID))
                    {
                        if (e.Row.RowIndex == rowindex)
                        {
                            chk.Checked = true;
                            IsChecked = true;
                        }
                    }
                    var existingSpecEntry = changedSpecs.FirstOrDefault(d => d.ContainsKey(materialID));
                    if (existingSpecEntry != null)
                    {
                        string spec = existingSpecEntry[materialID];
                        selectedSpec = spec;
                    }

                    var existingShortcodeEntry = changedShortcodes.FirstOrDefault(d => d.ContainsKey(materialID));
                    if (existingShortcodeEntry != null)
                    {
                        string shortcode = existingShortcodeEntry[materialID];
                        selectedShortcode = shortcode;
                    }

                    var existingIdentEntry = changedIdents.FirstOrDefault(d => d.ContainsKey(materialID));
                    if (existingIdentEntry != null)
                    {
                        string ident = existingIdentEntry[materialID];
                        selectedIdent = ident;
                    }

                    var existingUnitEntry = changedUnits.FirstOrDefault(d => d.ContainsKey(materialID));
                    if (existingUnitEntry != null)
                    {
                        string unit = existingUnitEntry[materialID];
                        selectedUnit = unit;
                    }

                    var existingPhaseEntry = changedPhases.FirstOrDefault(d => d.ContainsKey(materialID));
                    if (existingPhaseEntry != null)
                    {
                        string phase = existingPhaseEntry[materialID];
                        selectedPhase = phase;
                    }

                    var existingConstAreaEntry = changedConstAreas.FirstOrDefault(d => d.ContainsKey(materialID));
                    if (existingConstAreaEntry != null)
                    {
                        string constArea = existingConstAreaEntry[materialID];
                        selectedConstArea = constArea;
                    }
                    var existingText = changedText.FirstOrDefault(d => d.ContainsKey(materialID));
                    if (existingText != null)
                    {
                        string text = existingText[materialID];
                        selectedText = text;
                    }

                    Session["rowindex"] = null;

                }

                // Get mappings from Session
                var unitPhaseMap = Session["UnitPhaseMap"] as Dictionary<string, List<string>>;
                var unitConstAreaMap = Session["UnitConstAreaMap"] as Dictionary<string, Dictionary<string, List<string>>>;
                var specShortCodeMap = Session["SpecShortCodeMap"] as Dictionary<string, List<string>>;
                var shortCodeIdentMap = Session["ShortCodeIdentMap"] as Dictionary<string, Dictionary<string, List<string>>>;

                var allUnits = Session["Units"] as List<string>;
                var allSpecs = Session["Specs"] as List<string>;

                if (ddlUnit != null)
                {
                    ddlUnit.DataSource = allUnits;
                    try
                    {
                        ddlUnit.DataBind();
                    }
                    catch { }
                    ddlUnit.AutoPostBack = true;
                    ddlUnit.SelectedIndexChanged += ddlUnit_SelectedIndexChanged;
                    ddlUnit.Items.Insert(0, new ListItem("-- Select --", ""));
                    if (!string.IsNullOrEmpty(selectedUnit))
                    {
                        if (currentUnit != selectedUnit)
                        {
                            currentUnit = selectedUnit;
                            ddlUnit.SelectedValue = currentUnit;
                        }
                    }
                }
                if (ddlPhase != null && unitPhaseMap.ContainsKey(currentUnit) && !string.IsNullOrEmpty(selectedPhase))
                {
                    var validPhases = unitPhaseMap[currentUnit];
                    ddlPhase.DataSource = validPhases;
                    try
                    {
                        ddlPhase.DataBind();
                    }
                    catch { }
                    ddlPhase.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlPhase.AutoPostBack = true;
                    ddlPhase.SelectedIndexChanged += ddlPhase_SelectedIndexChanged;

                    if (!string.IsNullOrEmpty(selectedPhase))
                    {
                        if (currentPhase != selectedPhase)
                        {
                            currentPhase = selectedPhase;
                        }
                        if (validPhases.Contains(currentPhase))
                        {
                            ddlPhase.SelectedValue = currentPhase;
                        }
                    }
                }
                else if (ddlPhase != null && unitPhaseMap.ContainsKey(currentUnit))
                {
                    var validPhases = unitPhaseMap[currentUnit];
                    ddlPhase.DataSource = validPhases;
                    try
                    {
                        ddlPhase.DataBind();
                    }
                    catch { }
                    ddlPhase.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlPhase.AutoPostBack = true;
                    ddlPhase.SelectedIndexChanged += ddlPhase_SelectedIndexChanged;
                    if (validPhases.Contains(currentPhase))
                    {
                        ddlPhase.SelectedValue = currentPhase;
                    }
                }
                try
                {
                    if (ddlConstArea != null && !string.IsNullOrEmpty(selectedUnit) && !string.IsNullOrEmpty(selectedPhase) && unitConstAreaMap[selectedUnit].ContainsKey(selectedPhase))
                    {
                        ddlConstArea.SelectedIndex = 0;
                        var validConstAreas = unitConstAreaMap[selectedUnit][selectedPhase];
                        ddlConstArea.DataSource = validConstAreas;
                        try
                        {
                            ddlConstArea.DataBind();
                        }
                        catch { }

                        ddlConstArea.Items.Insert(0, new ListItem("-- Select --", ""));
                        ddlConstArea.AutoPostBack = true;
                        ddlConstArea.SelectedIndexChanged += ddlConstArea_SelectedIndexChanged;


                        if (!string.IsNullOrEmpty(selectedConstArea))
                        {
                            if (currentConstArea != selectedConstArea)
                            {
                                currentConstArea = selectedConstArea;
                            }
                            if (validConstAreas.Contains(currentConstArea))
                            {
                                ddlConstArea.SelectedValue = currentConstArea;
                            }
                        }
                    }
                    else if (ddlConstArea != null && !string.IsNullOrEmpty(currentUnit) && !string.IsNullOrEmpty(currentPhase) && unitConstAreaMap[currentUnit].ContainsKey(currentPhase))
                    {
                        ddlConstArea.SelectedIndex = 0;
                        var validConstAreas = unitConstAreaMap[currentUnit][currentPhase];
                        ddlConstArea.DataSource = validConstAreas;
                        try
                        {
                            ddlConstArea.DataBind();
                        }
                        catch { }
                        ddlConstArea.Items.Insert(0, new ListItem("-- Select --", ""));
                        ddlConstArea.AutoPostBack = true;
                        ddlConstArea.SelectedIndexChanged += ddlConstArea_SelectedIndexChanged;
                        if (validConstAreas.Contains(currentConstArea))
                        {
                            ddlConstArea.SelectedValue = currentConstArea;
                        }
                    }
                }
                catch (Exception ae)
                {
                    var xx = ae.Message;
                }
                if (ddlSpec != null && !string.IsNullOrEmpty(selectedSpec))
                {
                    ddlSpec.DataSource = allSpecs;
                    try
                    {
                        ddlSpec.DataBind();
                    }
                    catch { }
                    ddlSpec.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlSpec.AutoPostBack = true;
                    ddlSpec.SelectedIndexChanged += ddlSpec_SelectedIndexChanged;

                    if (!string.IsNullOrEmpty(selectedSpec) && allSpecs.Contains(selectedSpec))
                    {
                        if (currentSpec != selectedSpec)
                        {
                            currentSpec = selectedSpec;
                        }
                        ddlSpec.SelectedValue = currentSpec;
                    }

                }
                else if (ddlSpec != null && !string.IsNullOrEmpty(currentSpec) && allSpecs.Contains(currentSpec))
                {
                    ddlSpec.DataSource = allSpecs;
                    try
                    {
                        ddlSpec.DataBind();
                    }
                    catch { }
                    ddlSpec.AutoPostBack = true;
                    ddlSpec.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlSpec.SelectedIndexChanged += ddlSpec_SelectedIndexChanged;

                    ddlSpec.SelectedValue = currentSpec;
                }

                if (ddlShortcode != null && !string.IsNullOrEmpty(currentSpec) && !string.IsNullOrEmpty(selectedShortcode))
                {
                    var validShortcodes = specShortCodeMap[currentSpec];
                    ddlShortcode.DataSource = validShortcodes;
                    try
                    {
                        ddlShortcode.DataBind();
                    }
                    catch { }
                    ddlShortcode.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlShortcode.AutoPostBack = true;
                    ddlShortcode.SelectedIndexChanged += ddlShortcode_SelectedIndexChanged;
                    if (!string.IsNullOrEmpty(selectedShortcode) && validShortcodes.Contains(selectedShortcode))
                    {
                        if (currentShortcode != selectedShortcode)
                        {
                            currentShortcode = selectedShortcode;
                        }
                        ddlShortcode.SelectedValue = currentShortcode;
                    }
                }

                else if (ddlShortcode != null && !string.IsNullOrEmpty(currentSpec) && !string.IsNullOrEmpty(currentShortcode) && specShortCodeMap.ContainsKey(currentSpec))
                {
                    var validShortcodes = specShortCodeMap[currentSpec];
                    ddlShortcode.DataSource = validShortcodes;
                    try
                    {
                        ddlShortcode.DataBind();
                    }
                    catch { }
                    ddlShortcode.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlShortcode.AutoPostBack = true;
                    ddlShortcode.SelectedIndexChanged += ddlShortcode_SelectedIndexChanged;

                    if (validShortcodes.Contains(currentShortcode))
                    {
                        ddlShortcode.SelectedValue = currentShortcode;
                    }
                }
                if (ddlIdent != null
                 && !string.IsNullOrEmpty(currentSpec)
                 && !string.IsNullOrEmpty(currentShortcode)
                 && shortCodeIdentMap.ContainsKey(currentSpec)
                 && shortCodeIdentMap[currentSpec].ContainsKey(currentShortcode)
                 && !string.IsNullOrEmpty(selectedIdent))
                {
                    var validIdents = shortCodeIdentMap[currentSpec][currentShortcode];
                    ddlIdent.DataSource = validIdents;
                    try
                    {
                        ddlIdent.DataBind();
                    }
                    catch { }
                    ddlIdent.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlIdent.AutoPostBack = true;
                    ddlIdent.SelectedIndexChanged += ddlIdent_SelectedIndexChanged;

                    if (!string.IsNullOrEmpty(selectedIdent))
                    {
                        if (currentIdent != selectedIdent)
                        {
                            currentIdent = selectedIdent;
                        }
                        if (validIdents.Contains(selectedIdent))
                        {
                            ddlIdent.SelectedValue = currentIdent;
                            //need to set 

                            var match = specList?.FirstOrDefault(s => s.Lineclass == currentSpec && s.Ident == selectedIdent);

                            if (match != null)
                            {
                                SetLabelText(e.Row, "lblComponent_Type", match.Commodity_code);
                                SetLabelText(e.Row, "lblIsoShortDescription", match.Description);
                                SetLabelText(e.Row, "lblSize_sch1", match.Size_sch1);
                                SetLabelText(e.Row, "lblSize_sch2", match.Size_sch2);
                                SetLabelText(e.Row, "lblSize_sch3", match.Size_sch3);
                                SetLabelText(e.Row, "lblSize_sch4", match.Size_sch4);
                                SetLabelText(e.Row, "lblSize_sch5", match.Size_sch5);
                                string qtyUnitText = "pc"; // default

                                if (!string.IsNullOrEmpty(match.Shortcode))
                                {
                                    if (match.Shortcode == "PIP")
                                        qtyUnitText = "m";
                                }
                                SetLabelText(e.Row, "lblqty_unit", qtyUnitText);
                            }

                        }
                    }
                }

                else if (ddlIdent != null)
                {
                    var validIdents = shortCodeIdentMap[currentSpec][currentShortcode];
                    ddlIdent.DataSource = validIdents;
                    try
                    {
                        ddlIdent.DataBind();
                    }
                    catch { }
                    ddlIdent.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlIdent.AutoPostBack = true;
                    ddlIdent.SelectedIndexChanged += ddlIdent_SelectedIndexChanged;

                    if (validIdents.Contains(currentIdent))
                    {
                        ddlIdent.SelectedValue = currentIdent;
                        var match = specList?.FirstOrDefault(s => s.Lineclass == currentSpec && s.Ident == currentIdent);

                        if (match != null)
                        {
                            SetLabelText(e.Row, "lblComponent_Type", match.Commodity_code);
                            SetLabelText(e.Row, "lblIsoShortDescription", match.Description);
                            SetLabelText(e.Row, "lblSize_sch1", match.Size_sch1);
                            SetLabelText(e.Row, "lblSize_sch2", match.Size_sch2);
                            SetLabelText(e.Row, "lblSize_sch3", match.Size_sch3);
                            SetLabelText(e.Row, "lblSize_sch4", match.Size_sch4);
                            SetLabelText(e.Row, "lblSize_sch5", match.Size_sch5);
                            string qtyUnitText = "pc"; // default

                            if (!string.IsNullOrEmpty(match.Shortcode))
                            {
                                if (match.Shortcode == "PIP")
                                    qtyUnitText = "m";
                            }
                            SetLabelText(e.Row, "lblqty_unit", qtyUnitText);
                        }
                    }
                }
                if (ddlIdent != null && specList != null && !string.IsNullOrEmpty(currentSpec))
                {
                    foreach (ListItem item in ddlIdent.Items)
                    {
                        string identValue = item.Value;

                        var match = specList.FirstOrDefault(s => s.Lineclass == currentSpec && s.Ident == identValue);
                        if (match != null && !string.IsNullOrEmpty(match.Description))
                        {
                            item.Attributes["title"] = match.Description;
                        }
                        else
                        {
                            item.Attributes["title"] = "No description available";
                        }
                    }
                }



                if (chk != null && chklist.Count > 0)
                {
                    chk.AutoPostBack = true;
                    chk.Checked = IsChecked;
                    chk.CheckedChanged += Chk_CheckedChanged;
                }
                else
                {
                    chk.AutoPostBack = true;
                    chk.CheckedChanged += Chk_CheckedChanged;

                }
                if (txtQty != null)
                {
                    if (currentText != selectedText && selectedText != "")
                    {
                        txtQty.Text = selectedText;
                    }
                    else
                    {
                        txtQty.Text = currentText;
                    }
                    txtQty.AutoPostBack = true;
                    txtQty.Attributes.Add("style", "width:80px;");
                    txtQty.TextChanged += txtQty_TextChanged;
                }
            }
        }

        private void BindExcelGridView(List<GridData> griddata, string area, string isosheet, bool bindData = true)
        {
            var gridDataTable = ConvertToDataTable(griddata);
            if (Session["GridData"] != null)
            {
                var tmpgd = Session["GridData"] as List<GridData>;
                var existid = griddata.Select(x => x.MaterialID).Distinct().ToList();
                var notinlistalready = tmpgd.Where(x => !existid.Contains(x.MaterialID)).ToList();
                if (notinlistalready.Count > 0)
                {
                    var tmpgdtable = ConvertToDataTable(notinlistalready);
                    gridDataTable.Merge(tmpgdtable);
                }

            }

            lblFileName.Text = "Area: " + area + " -> ISO: " + isosheet;

            ExcelGridView.Columns.Clear();

            // Recreate the checkbox column
            TemplateField checkboxColumn = new TemplateField
            {
                HeaderText = "Checked",
                HeaderStyle = { Wrap = false },
                ItemStyle = { Wrap = false },
                ItemTemplate = new CheckBoxTemplate("chkChecked")
            };
            ExcelGridView.Columns.Add(checkboxColumn);

            ButtonField deleteButton = new ButtonField
            {
                ButtonType = ButtonType.Button,
                Text = "Delete",
                CommandName = "DeleteRow",
                HeaderText = "Delete",
            };
            ExcelGridView.Columns.Add(deleteButton);

            // Add columns dynamically
            var hiddenColumns = new HashSet<string> { "MaterialID", "ProjectID", "Discipline", "Area", "ISO", "Fabrication_Type", "IsoRevision", "IsoRevisionDate", "IsLocked", "IsoUniqeRevID" };
            foreach (DataColumn col in gridDataTable.Columns)
            {

                if (col.ColumnName != "Checked")
                {
                    TemplateField templateField = null;

                    switch (col.ColumnName)
                    {
                        case "qty":
                            templateField = new TemplateField
                            {
                                HeaderText = "qty",
                                ItemTemplate = new QtyTextBoxTemplate("txtQty"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;

                        case "Unit":
                            templateField = new TemplateField
                            {
                                HeaderText = "Unit",
                                ItemTemplate = new UnitDropDownTemplate("ddlUnit"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;

                        case "Phase":
                            templateField = new TemplateField
                            {
                                HeaderText = "Phase",
                                ItemTemplate = new PhaseDropDownTemplate("ddlPhase"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;

                        case "Const_Area":
                            templateField = new TemplateField
                            {
                                HeaderText = "Const_Area",
                                ItemTemplate = new ConstAreaDropDownTemplate("ddlConstArea"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;
                        case "Spec":
                            templateField = new TemplateField
                            {
                                HeaderText = "Spec",
                                ItemTemplate = new SpecDropDownTemplate("ddlSpec"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;
                        case "Shortcode":
                            templateField = new TemplateField
                            {
                                HeaderText = "Shortcode",
                                ItemTemplate = new ShortCodeDropDownTemplate("ddlShortcode"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;
                        case "Ident_no":
                            templateField = new TemplateField
                            {
                                HeaderText = "Ident_no",
                                ItemTemplate = new IdentDropDownTemplate("ddlIdent"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;
                        case "Component_Type":
                            templateField = new TemplateField
                            {
                                HeaderText = "Component_Type",
                                ItemTemplate = new LabelTemplate("lblComponent_Type", "Component_Type"),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;
                        case "IsoShortDescription":
                            templateField = new TemplateField
                            {
                                HeaderText = "IsoShortDescription",
                                ItemTemplate = new LabelISoTemplate("lblIsoShortDescription", "IsoShortDescription"),
                                ItemStyle = { Wrap = true },
                                HeaderStyle = { Wrap = false }
                            };
                            break;
                        case "Size_sch1":
                        case "Size_sch2":
                        case "Size_sch3":
                        case "Size_sch4":
                        case "Size_sch5":
                        case "qty_unit":
                            templateField = new TemplateField
                            {
                                HeaderText = col.ColumnName,
                                ItemTemplate = new LabelTemplate("lbl" + col.ColumnName, col.ColumnName),
                                ItemStyle = { Wrap = false },
                                HeaderStyle = { Wrap = false }
                            };
                            break;

                        default:
                            BoundField boundField = new BoundField
                            {
                                DataField = col.ColumnName,
                                HeaderText = col.ColumnName
                            };
                            //if (col.ColumnName == "IsoShortDescription")
                            //{
                            //    boundField.ItemStyle.Wrap = true;
                            //    boundField.HeaderStyle.Wrap = false;
                            //    boundField.ItemStyle.Width = Unit.Pixel(250);
                            //}
                            //else
                            //{
                            boundField.ItemStyle.Wrap = false;
                            boundField.HeaderStyle.Wrap = false;
                            //}

                            ExcelGridView.Columns.Add(boundField);
                            break;
                    }

                    if (templateField != null)
                    {
                        ExcelGridView.Columns.Add(templateField);
                    }
                }
            }

            if (bindData)
            {
                ExcelGridView.DataSource = gridDataTable;
                ExcelGridView.DataBind();
            }

            ExcelGridView.Columns
                .Cast<DataControlField>()
                .Where(c => hiddenColumns.Contains(c.HeaderText))
                .ToList()
                .ForEach(c => c.Visible = false);

            ExcelGridView.HeaderStyle.BackColor = System.Drawing.Color.LightBlue;
            ExcelGridView.Visible = true;
            pnlExcelUpload.Visible = true;
        }

        private void Chk_CheckedChanged(object sender, EventArgs e)
        {
            System.Web.UI.WebControls.CheckBox chk = (System.Web.UI.WebControls.CheckBox)sender;
            GridViewRow row = (GridViewRow)chk.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var list = Session["CheckedMaterialIDs"] as List<int> ?? new List<int>();
            if (chk.Checked)
            {
                if (!list.Contains(materialID))
                {
                    list.Add(materialID);
                }
            }
            else
            {
                if (list.Contains(materialID))
                {
                    list.Remove(materialID);
                }
            }
            Session["rowindex"] = row.RowIndex;
            Session["CheckedMaterialIDs"] = list;
            chk.Focus();
        }

        private void txtQty_TextChanged(object sender, EventArgs e)
        {
            lblMessage.Text = "";
            System.Web.UI.WebControls.TextBox txt = (System.Web.UI.WebControls.TextBox)sender;
            GridViewRow row = (GridViewRow)txt.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var griddata = Session["GridData"] as List<GridData>;
            var itemToUpdate = new GridData();
            if (griddata != null)
            {
                itemToUpdate = griddata.FirstOrDefault(g => g.MaterialID == materialID);
            }

            // Get the linked ddlShortcode
            DropDownList ddlShortcode = (DropDownList)row.FindControl("ddlShortcode");
            string selectedShortcode = ddlShortcode?.SelectedValue;

            string inputQty = txt.Text.Trim();
            bool isValid = false;

            if (selectedShortcode == "PIP")
            {
                // Allow decimal
                isValid = !DecParse(inputQty).HasError;
            }
            else
            {
                // Allow only whole numbers
                isValid = int.TryParse(inputQty, out _);
            }

            if (!isValid)
            {
                lblMessage.Text = selectedShortcode == "PIP"
                    ? "Please enter a valid decimal quantity for PIP items."
                    : "Please enter a whole number quantity.";
                lblMessage.ForeColor = System.Drawing.Color.Red;
                return;
            }

            // Save the changed quantity
            var list = Session["ChangedText"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var existingEntry = list.FirstOrDefault(d => d.ContainsKey(materialID));
            if (existingEntry != null)
            {
                existingEntry[materialID] = inputQty;
            }
            else
            {
                list.Add(new Dictionary<int, string> { { materialID, inputQty } });
            }
            if (itemToUpdate != null && itemToUpdate.MaterialID == materialID)
            {
                itemToUpdate.qty = inputQty;
                Session["GridData"] = griddata;
            }
            Session["ChangedText"] = list;
        }
        protected void ddlUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlUnit = (DropDownList)sender;
            GridViewRow row = (GridViewRow)ddlUnit.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var changedUnitsList = Session["ChangedUnits"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var existingEntry = changedUnitsList.FirstOrDefault(d => d.ContainsKey(materialID));
            if (existingEntry != null)
            {
                existingEntry[materialID] = ddlUnit.SelectedValue;
            }
            else
            {
                changedUnitsList.Add(new Dictionary<int, string> { { materialID, ddlUnit.SelectedValue } });
            }
            Session["ChangedUnits"] = changedUnitsList;
            Session["rowindex"] = row.RowIndex;
        }
        protected void ddlSpec_SelectedIndexChanged(object sender, EventArgs e)
        {

            var ddl = (DropDownList)sender;
            var row = (GridViewRow)ddl.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var changedSpecsList = Session["ChangedSpecs"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var existingEntry = changedSpecsList.FirstOrDefault(d => d.ContainsKey(materialID));
            if (existingEntry != null)
            {
                existingEntry[materialID] = ddl.SelectedValue;
            }
            else
            {
                changedSpecsList.Add(new Dictionary<int, string> { { materialID, ddl.SelectedValue } });
            }
            Session["ChangedSpecs"] = changedSpecsList;
            Session["rowindex"] = row.RowIndex;
        }
        protected void ddlShortcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            var ddl = (DropDownList)sender;
            var row = (GridViewRow)ddl.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var changedShortcodesList = Session["ChangedShortcodes"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var existingEntry = changedShortcodesList.FirstOrDefault(d => d.ContainsKey(materialID));
            if (existingEntry != null)
            {
                existingEntry[materialID] = ddl.SelectedValue;
            }
            else
            {
                changedShortcodesList.Add(new Dictionary<int, string> { { materialID, ddl.SelectedValue } });
            }
            Session["ChangedShortcodes"] = changedShortcodesList;
            Session["rowindex"] = row.RowIndex;
        }
        private void ddlIdent_SelectedIndexChanged(object sender, EventArgs e)
        {
            var ddl = (DropDownList)sender;
            var row = (GridViewRow)ddl.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);

            var changedIdentsList = Session["ChangedIdents"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            //var existingEntry = changedIdentsList.FirstOrDefault(d => d.ContainsKey(materialID));
            //if (existingEntry != null)
            //{

            int index = changedIdentsList.FindIndex(d => d.ContainsKey(materialID));
            if (index >= 0)
            {
                changedIdentsList[index][materialID] = ddl.SelectedValue;
            }
            else
            {
                changedIdentsList.Add(new Dictionary<int, string> { { materialID, ddl.SelectedValue } });
            }

            //}
            //else
            //{
            //    changedIdentsList.Add(new Dictionary<int, string> { { materialID, ddl.SelectedValue } });
            //}
            Session["ChangedIdents"] = changedIdentsList;
            Session["rowindex"] = row.RowIndex;

        }
        private void ddlPhase_SelectedIndexChanged(object sender, EventArgs e)
        {
            var ddl = (DropDownList)sender;
            var row = (GridViewRow)ddl.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var changedPhasesList = Session["ChangedPhases"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var existingEntry = changedPhasesList.FirstOrDefault(d => d.ContainsKey(materialID));
            if (existingEntry != null)
            {
                existingEntry[materialID] = ddl.SelectedValue;
            }
            else
            {
                changedPhasesList.Add(new Dictionary<int, string> { { materialID, ddl.SelectedValue } });
            }
            Session["ChangedPhases"] = changedPhasesList;
            Session["rowindex"] = row.RowIndex;
        }
        private void ddlConstArea_SelectedIndexChanged(object sender, EventArgs e)
        {
            var ddl = (DropDownList)sender;
            var row = (GridViewRow)ddl.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var changedConstAreasList = Session["ChangedConstAreas"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
            var existingEntry = changedConstAreasList.FirstOrDefault(d => d.ContainsKey(materialID));
            if (existingEntry != null)
            {
                existingEntry[materialID] = ddl.SelectedValue;
            }
            else
            {
                changedConstAreasList.Add(new Dictionary<int, string> { { materialID, ddl.SelectedValue } });
            }
            Session["ChangedConstAreas"] = changedConstAreasList;
            Session["rowindex"] = row.RowIndex;
        }
        private void SetLabelText(GridViewRow row, string controlId, string text)
        {
            var label = row.FindControl(controlId) as System.Web.UI.WebControls.Label;
            if (label != null)
            {
                label.Text = text;
            }
        }
        #region Add Form
        protected void btnAddNew_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            lblMessage.Text = "";
            pnlAddForm.Visible = true;

            ddlAddUnit.DataSource = Session["Units"] as List<string>;
            ddlAddUnit.DataBind();
            ddlAddUnit.Items.Insert(0, new ListItem("-- Select --", ""));

            ddlAddSpec.DataSource = Session["Specs"] as List<string>;
            ddlAddSpec.DataBind();
            ddlAddSpec.Items.Insert(0, new ListItem("-- Select --", ""));

            // Clear dependent dropdowns
            ddlAddPhase.Items.Clear();
            ddlAddConstArea.Items.Clear();
            ddlAddShortcode.Items.Clear();
            ddlAddIdent.Items.Clear();
            txtAddQty.Text = "";
            txtAddQty.Attributes["style"] = "background-color:#ffffcc; border:2px solid #ff9900;";
            txtAddQty.Attributes["placeholder"] = "Enter quantity/Length";



        }

        protected void btnCancelAdd_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            lblMessage.Text = "";
            ddlAddUnit.ClearSelection();
            ddlAddPhase.ClearSelection();
            ddlAddConstArea.ClearSelection();
            ddlAddSpec.ClearSelection();
            ddlAddShortcode.ClearSelection();
            ddlAddIdent.ClearSelection();
            txtAddQty.Text = "";
            pnlAddForm.Visible = false;
        }

        protected void btnSaveNew_Click(object sender, EventArgs e)
        {
            lblMessage.Text = "";
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            string shortcode = ddlAddShortcode.SelectedValue;
            string qtyText = txtAddQty.Text.Trim();
            bool isValidQty = false;
            var SelectedISO = ddliso.SelectedItem.Text;

            string isorev = "";
            var match = System.Text.RegularExpressions.Regex.Match(SelectedISO, @"\bRev:\s*([^\s:]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success)
            {
                isorev = match.Groups[1].Value; // e.g., "0", "0.1", "S0", "R0", etc.
            }

            if (shortcode == "PIP")
            {
                isValidQty = !DecParse(qtyText.Trim()).HasError;
            }
            else
            {
                isValidQty = int.TryParse(qtyText, out _);
            }

            if (!isValidQty)
            {

                lblMessage.Text = shortcode == "PIP"
                 ? "Please enter a valid decimal quantity for PIP items."
                 : "Please enter a whole number quantity.";
                lblMessage.ForeColor = System.Drawing.Color.Red;
                pnlAddForm.Visible = true;
                return;

            }

            // Validate all required fields
            if (string.IsNullOrWhiteSpace(ddlAddUnit.SelectedValue) ||
         string.IsNullOrWhiteSpace(ddlAddPhase.SelectedValue) ||
         string.IsNullOrWhiteSpace(ddlAddConstArea.SelectedValue) ||
         string.IsNullOrWhiteSpace(ddlAddSpec.SelectedValue) ||
         string.IsNullOrWhiteSpace(ddlAddShortcode.SelectedValue) ||
         string.IsNullOrWhiteSpace(ddlAddIdent.SelectedValue) ||
         string.IsNullOrWhiteSpace(txtAddQty.Text) ||
         DecParse(txtAddQty.Text.Trim()).HasError)
            {
                lblMessage.Text = "Please fill in all fields with valid values.";
                lblMessage.ForeColor = System.Drawing.Color.Red;
                pnlAddForm.Visible = true;
                return;
            }

            var griddata = Session["GridData"] as List<GridData> ?? new List<GridData>();
            var specList = Session["LoadedSpecs"] as List<SpecData> ?? new List<SpecData>();
            var tempgriddata = new List<GridData>();


            if (Session["TempMaterialID"] == null)
            {

                int projectID = int.Parse(ddlprojsel.SelectedValue);
                Session["TempMaterialID"] = DataClass.GetNextNegativeMaterialID(projectID); ;
            }
            int tempID = (int)Session["TempMaterialID"];
            Session["TempMaterialID"] = tempID - 1;

            // Create a temporary SPMATDBData object
            var testArea = Session["CurrentArea"] as string ?? "";
            var testDisipline = Session["CurrentDisipline"] as string ?? "";

            decimal parsedQty;
            bool isValidNumber = decimal.TryParse(txtAddQty.Text.Trim(), out parsedQty);
            bool hasFraction = isValidNumber && parsedQty % 1 != 0;

            string selectedShortcode = ddlAddShortcode.SelectedValue?.Trim();
            string qtyUnit = (string.Equals(selectedShortcode, "PIP", StringComparison.OrdinalIgnoreCase))
             ? "m"
             : (hasFraction ? "m" : "pc");


            var newSPMAT = new SPMATDBData
            {
                MaterialID = tempID, // Temporary ID
                ProjectID = int.Parse(ddlprojsel.SelectedValue),
                Discipline = testDisipline, // Optional
                Area = testArea, // Optional
                Unit = ddlAddUnit.SelectedValue,
                Phase = ddlAddPhase.SelectedValue,
                Const_Area = ddlAddConstArea.SelectedValue,
                ISO = ddliso.SelectedValue,
                Ident_no = ddlAddIdent.SelectedValue,
                qty = txtAddQty.Text.Trim(),
                qty_unit = qtyUnit,
                Fabrication_Type = "Undefined",
                Spec = ddlAddSpec.SelectedValue,
                IsoRevision = string.IsNullOrEmpty(isorev) ? "" : isorev,
                IsoRevisionDate = "",
                Lock = "",
                Code = "M",
                IsoUniqeRevID = 0
            };

            // Use PopulateGridData to enrich the entry
            var enriched = PopulateGridData(new List<SPMATDBData> { newSPMAT }, specList);
            if (enriched.Any())
            {
                var newItem = enriched.First();
                DataClass.InsertIntoSPMAT_REQData_Temp(new SPMATDBData
                {
                    MaterialID = newItem.MaterialID,
                    ProjectID = newItem.ProjectID,
                    Discipline = newItem.Discipline,
                    Area = newItem.Area,
                    Unit = newItem.Unit,
                    Phase = newItem.Phase,
                    Const_Area = newItem.Const_Area,
                    ISO = newItem.ISO,
                    Ident_no = newItem.Ident_no,
                    qty = newItem.qty,
                    qty_unit = newItem.qty_unit,
                    Fabrication_Type = newItem.Fabrication_Type,
                    Spec = newItem.Spec,
                    IsoRevisionDate = newItem.IsoRevisionDate,
                    IsoRevision = newItem.IsoRevision,
                    Lock = newItem.IsLocked,
                    Code = newItem.Source,
                    IsoUniqeRevID = newItem.IsoUniqeRevID
                });

                griddata.Add(newItem);
            }


            Session["GridData"] = griddata;
            tempgriddata = griddata;

            // Clear form fields for next entry
            ddlAddUnit.ClearSelection();
            ddlAddPhase.ClearSelection();
            ddlAddConstArea.ClearSelection();
            ddlAddSpec.ClearSelection();
            ddlAddShortcode.ClearSelection();
            ddlAddIdent.ClearSelection();
            txtAddQty.Text = "";
            lblidentdesc.Text = "";
            // Keep form visible for next entry
            btnCancelAdd_Click(null, null);
            pnlAddForm.Visible = false;


            BindExcelGridView(griddata, Session["CurrentArea"].ToString(), Session["CurrentIso"].ToString(), bindData: true);


        }

        protected void ddlAddUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            var unit = ddlAddUnit.SelectedValue;
            var unitPhaseMap = Session["UnitPhaseMap"] as Dictionary<string, List<string>>;
            if (unitPhaseMap != null && unitPhaseMap.ContainsKey(unit))
            {
                ddlAddPhase.DataSource = unitPhaseMap[unit];
                ddlAddPhase.DataBind();
                ddlAddPhase.Items.Insert(0, new ListItem("-- Select --", ""));
            }

            ddlAddConstArea.Items.Clear();
        }
        protected void ddlAddPhase_SelectedIndexChanged(object sender, EventArgs e)
        {
            var unit = ddlAddUnit.SelectedValue;
            var phase = ddlAddPhase.SelectedValue;
            var unitConstAreaMap = Session["UnitConstAreaMap"] as Dictionary<string, Dictionary<string, List<string>>>;
            if (unitConstAreaMap != null && unitConstAreaMap.ContainsKey(unit) && unitConstAreaMap[unit].ContainsKey(phase))
            {
                ddlAddConstArea.DataSource = unitConstAreaMap[unit][phase];
                ddlAddConstArea.DataBind();
                ddlAddConstArea.Items.Insert(0, new ListItem("-- Select --", ""));
            }
        }
        protected void ddlAddSpec_SelectedIndexChanged(object sender, EventArgs e)
        {
            var spec = ddlAddSpec.SelectedValue;
            var specShortCodeMap = Session["SpecShortCodeMap"] as Dictionary<string, List<string>>;
            if (specShortCodeMap != null && specShortCodeMap.ContainsKey(spec))
            {
                ddlAddShortcode.DataSource = specShortCodeMap[spec];
                ddlAddShortcode.DataBind();
                ddlAddShortcode.Items.Insert(0, new ListItem("-- Select --", ""));
                lblidentdesc.Text = "";
            }

            ddlAddIdent.Items.Clear();
        }
        protected void ddlAddShortcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            var spec = ddlAddSpec.SelectedValue;
            var shortcode = ddlAddShortcode.SelectedValue;
            var shortCodeIdentMap = Session["ShortCodeIdentMap"] as Dictionary<string, Dictionary<string, List<string>>>;
            if (shortCodeIdentMap != null && shortCodeIdentMap.ContainsKey(spec) && shortCodeIdentMap[spec].ContainsKey(shortcode))
            {
                ddlAddIdent.DataSource = shortCodeIdentMap[spec][shortcode];
                ddlAddIdent.DataBind();
                ddlAddIdent.Items.Insert(0, new ListItem("-- Select --", ""));
                lblidentdesc.Text = "";
            }
        }

        #endregion
        //protected void btnviewspmat_Click(object sender, EventArgs e)
        //{
        //    diverror.Style["display"] = "none";
        //    lblerror.Text = "";
        //    if (string.IsNullOrEmpty(ddlprojsel.SelectedValue))
        //    {
        //        diverror.Style["display"] = "block";
        //       lblerror.Text = "Please select a Spec and  Project before loading.";
        //        lblerror.ForeColor = System.Drawing.Color.Red;
        //        return;
        //    }
        //    var Projid = ddlprojsel.SelectedValue;
        //    List<SPMATDBData> mtoData = DataClass.GetWorkingMTOData(Projid);

        //    gvSPMAT.EmptyDataText = "No Working Data";
        //    gvSPMAT.DataSource = mtoData;
        //    gvSPMAT.DataBind();
        //    pnlSPMATView.Visible = true;
        //    if (mtoData.Count > 0)
        //    {
        //        btnExportSPMAT.Style["Display"] = "block";
        //    }
        //    else
        //    {
        //        btnExportSPMAT.Style["Display"] = "none";
        //    }

        //    ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionTwo", "var el = document.getElementById('collapseTwo'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);

        //    // Store in session for export
        //    Session["SPMATExportData"] = mtoData;

        //    if (Session["GridData"] != null && Session["CurrentIso"] != null && Session["CurrentArea"] != null)
        //    {
        //        var griddata = Session["GridData"] as List<GridData>;
        //        var area = Session["CurrentArea"].ToString();
        //        var iso = Session["CurrentIso"].ToString();
        //        BindExcelGridView(griddata, area, iso, bindData: true);
        //    }


        //}

        protected void btnExportSPMAT_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            var data = Session["SPMATExportData"] as List<SPMATDBData>;
            if (data == null || !data.Any()) return;
            List<int> materialIDs = data.Select(d => d.MaterialID).Distinct().ToList();
            MoveToIntrim(data, materialIDs);
            // gvSPMAT.EmptyDataText = "No Working Data";
            // gvSPMAT.DataSource = null;
            // gvSPMAT.DataBind();
            // pnlSPMATView.Visible = false;
            // btnExportSPMAT.Style["Display"] = "none";
            //btnviewspmat_Click(null, null);
            //ForceButtons("btnviewspmat", false);
            // btnviewInterim_Click(null, null);
            // ForceButtons("btnviewInterim",true);


        }
        private void MoveToIntrim(List<SPMATDBData> data, List<int> materialIDs)
        {
            foreach (SPMATDBData s in data)
            {
                DataClass.MoveToIntrim(s);
            }
        }
        protected void ddlAddIdent_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedSpec = ddlAddSpec.SelectedValue;
            string selectedIdent = ddlAddIdent.SelectedValue;
            if (ddlAddIdent.SelectedIndex > 0)
            {
                var specList = Session["LoadedSpecs"] as List<SpecData>;
                if (specList != null)
                {
                    var match = specList.FirstOrDefault(s => s.Lineclass == selectedSpec && s.Ident == selectedIdent);
                    if (match != null)
                    {
                        lblidentdesc.Text = match.Description;
                    }
                    else
                    {
                        lblidentdesc.Text = "Description not found.";
                    }
                }
            }
            else
            {
                lblidentdesc.Text = "";

            }

        }
        //private int SaveExportMetadata(string fileName, List<int> MTOIDs)
        //{
        //    return DataClass.SaveExportRecord(fileName, MTOIDs);
        //}
        protected void btnMoveToFinal_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            if (string.IsNullOrEmpty(ddlprojsel.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select a Spec and  Project before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
            var data = Session["SPMATIntrimData"] as List<SPMATIntrimData>;
            if (data == null || !data.Any()) return;
            MoveToFinal(data);
            //gvSPMAT.EmptyDataText = "No Working Data";
            //gvSPMAT.DataSource = null;
            //gvSPMAT.DataBind();
            //pnlSPMATView.Visible = false;

            //btnExportSPMAT.Style["Display"] = "none";
            //btnviewspmat_Click(null, null);
            //ForceButtons("btnviewspmat", false);
            //btnviewInterim_Click(null, null);
            ForceButtons("btnviewInterim", false);
            //btnViewFinal_Click(null, null);
            ForceButtons("btnViewFinal", true);
        }

        private void MoveToFinal(List<SPMATIntrimData> data)
        {
            bool IsPreviousData = bool.TryParse(Session["IsPreviousData"]?.ToString(), out var result) && result;
            foreach (SPMATIntrimData s in data)
            {
                DataClass.MoveToFinal(s, IsPreviousData);
            }
            var Projid = ddlprojsel.SelectedValue;
            List<SPMATIntrimData> mtoData = DataClass.GetMTOIntrimData(Projid);
            gvInterim.EmptyDataText = "No Working MTO Data";
            gvInterim.DataSource = mtoData;
            gvInterim.DataBind();
            pnlViewInterimData.Visible = true;
            if (mtoData.Count > 0)
            {
                lblintrim.Style["Display"] = "block";
                btnMoveToFinal.Style["Display"] = "block";
            }
            else
            {
                lblintrim.Style["Display"] = "none";
                btnMoveToFinal.Style["Display"] = "none";
            }
            // Store in session for export
            Session["SPMATIntrimData"] = mtoData;

            if (Session["GridData"] != null && Session["CurrentIso"] != null && Session["CurrentArea"] != null)
            {
                var griddata = Session["GridData"] as List<GridData>;
                var area = Session["CurrentArea"].ToString();
                var iso = Session["CurrentIso"].ToString();
                BindExcelGridView(griddata, area, iso, bindData: true);
            }
            // btnviewInterim_Click(null, null);
            ForceButtons("btnViewExported", true);
        }

        protected void btnviewInterim_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            if (string.IsNullOrEmpty(ddlprojsel.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select a Spec and  Project before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
            var Projid = ddlprojsel.SelectedValue;
            List<SPMATIntrimData> mtoData = DataClass.GetMTOIntrimData(Projid);

            gvInterim.EmptyDataText = "No Working MTO Data";
            gvInterim.DataSource = mtoData;
            gvInterim.DataBind();
            pnlViewInterimData.Visible = true;
            if (mtoData.Count > 0)
            {
                lblintrim.Style["Display"] = "block";
                btnMoveToFinal.Style["Display"] = "block";
            }
            else
            {
                lblintrim.Style["Display"] = "none";
                btnMoveToFinal.Style["Display"] = "none";
            }
            // Store in session for export
            ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionThree", "var el = document.getElementById('collapseThree'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
            Session["SPMATIntrimData"] = mtoData;

            if (Session["GridData"] != null && Session["CurrentIso"] != null && Session["CurrentArea"] != null)
            {
                var griddata = Session["GridData"] as List<GridData>;
                var area = Session["CurrentArea"].ToString();
                var iso = Session["CurrentIso"].ToString();
                BindExcelGridView(griddata, area, iso, bindData: true);
            }
        }

        protected void gvInterim_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "RemoveRow")
            {
                int fileId = Convert.ToInt32(e.CommandArgument);
                GridViewRow row = ((Button)e.CommandSource).NamingContainer as GridViewRow;

                var keys = gvInterim.DataKeys[row.RowIndex];
                string INTID = keys["INTID"].ToString();
                string MaterialID = keys["MaterialID"].ToString();
                string ISO = keys["ISO"].ToString();
                int Uniquerev = int.Parse(keys["IsoUniqeRevID"].ToString());

                // Refactored: Use DataClass methods instead of inline SQL
                DataClass.DeleteMTOEntry(ISO);
                DataClass.UncheckREQEntry(ISO, Uniquerev);
                Session["IsPreviousData"] = null;
                ReloadIso();
                // Refresh the grid
                //btnviewInterim_Click(null, null);
                ForceButtons("btnviewInterim", false);
            }

        }

        protected void btnViewFinal_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            if (string.IsNullOrEmpty(ddlprojsel.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select a Spec and  Project before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
            var Projid = ddlprojsel.SelectedValue;
            List<SPMATData> mtoData = DataClass.GetMTOData(Projid, false);

            gvFinal.EmptyDataText = "No MTO Export Data";
            gvFinal.DataSource = mtoData;
            gvFinal.DataBind();
            pnlFinalMTO.Visible = true;
            if (mtoData.Count > 0)
            {

                btnExportFinalMTO.Style["Display"] = "block";
            }
            else
            {
                btnExportFinalMTO.Style["Display"] = "none";
            }
            ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionFour", "var el = document.getElementById('collapseFour'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
            // Store in session for export
            Session["SPMATFinalData"] = mtoData;

            if (Session["GridData"] != null && Session["CurrentIso"] != null && Session["CurrentArea"] != null)
            {
                var griddata = Session["GridData"] as List<GridData>;
                var area = Session["CurrentArea"].ToString();
                var iso = Session["CurrentIso"].ToString();
                BindExcelGridView(griddata, area, iso, bindData: true);
            }
        }

        protected void btnExportFinalMTO_Click(object sender, EventArgs e)
        {
            btnExportFinalMTO.Style["display"] = "none";
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            if (string.IsNullOrEmpty(ddlprojsel.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select a Spec and  Project before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
            List<SPMATData> datacurrent = Session["SPMATFinalData"] as List<SPMATData>;
            var Projid = ddlprojsel.SelectedValue;
            List<SPMATData> mtoData = DataClass.GetMTOData(Projid, true);

            if (mtoData.Where(x => x.ImportStatus.ToLower().Trim() == "not imported").Any())
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Cannot Export MTO while there is an MTO File that have not been Imported.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
            if (datacurrent == null || !datacurrent.Any())
            {

                ForceButtons("btnViewExported", false);
                ForceButtons("btnViewFinal", true);
                return;
            }
            List<int> MTOIDs = datacurrent.Select(d => d.MTOID).Distinct().ToList();
            int FileID = DataClass.InsertExportRecord(Projid, MTOIDs); // returns the new FileExport ID
                                                                       // List<int> OtherFiles = DataClass.GetOtherFileIDs(Projid); // All existing files

            // List<int> mtoIDsToDelete = DataClass.GetMTOIDsToDelete(Projid, MTOIDs, OtherFiles);
            List<(int MaterialID, int MTOID)> itemsToDelete = DataClass.GetObsoleteMTOs(Projid);

            datacurrent.AddRange(mtoData);
            List<SPMATData> data = datacurrent.Where(d => !itemsToDelete.Any(del => del.MTOID == d.MTOID)).ToList();

            ClosedXML.Excel.XLWorkbook workbook = new ClosedXML.Excel.XLWorkbook();

            IXLWorksheet worksheet = workbook.Worksheets.Add("SPMAT_Export");

            // Define headers
            var headers = new[]
            {
                "Discipline", "Area", "Unit", "Phase", "Const_Area", "ISO", "Ident_no",
                "qty", "qty_unit", "Fabrication_Type", "Spec", "Pos",
                "IsoRevisionDate", "IsoRevision", "IsLocked","IsoUniqeRevID" };


            // Add headers
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = worksheet.Cell(1, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            // Add data
            int row = 2;
            foreach (var item in data)
            {
                worksheet.Cell(row, 1).Value = item.Discipline ?? "";
                worksheet.Cell(row, 2).Value = item.Area ?? "";
                worksheet.Cell(row, 3).Value = item.Unit ?? "";
                worksheet.Cell(row, 4).Value = item.Phase ?? "";
                worksheet.Cell(row, 5).Value = item.Const_Area ?? "";
                worksheet.Cell(row, 6).Value = item.ISO ?? "";
                worksheet.Cell(row, 7).Value = item.Ident_no ?? "";
                worksheet.Cell(row, 8).Value = item.qty.ToString().Replace(",", ".");
                worksheet.Cell(row, 9).Value = item.qty_unit ?? "";
                worksheet.Cell(row, 10).Value = item.Fabrication_Type ?? "";
                worksheet.Cell(row, 11).Value = item.Spec ?? "";
                worksheet.Cell(row, 12).Value = item.Pos ?? "";
                worksheet.Cell(row, 13).Value = item.IsoRevisionDate ?? "";
                worksheet.Cell(row, 14).Value = item.IsoRevision ?? "";
                worksheet.Cell(row, 15).Value = item.IsLocked ?? "";
                worksheet.Cell(row, 16).Value = item.IsoUniqeRevID.ToString();

                for (int col = 1; col <= headers.Length; col++)
                {
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Cell(row, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }

                row++;
            }
            worksheet.CellsUsed().Style.NumberFormat.Format = "@";
            // Adjust column widths
            worksheet.Columns().AdjustToContents();
            var FileName = Server.UrlEncode(ddlprojsel.SelectedItem.Text.Split('-')[0].Trim() + "_MTO_ALL_SPMAT_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx");
            DataClass.UpdateExportRecord(FileID, FileName);
            var mtoIDs = itemsToDelete.Select(x => x.MTOID).Distinct().ToList();
            var materialIDs = itemsToDelete.Select(x => x.MaterialID).Distinct().ToList();
            DataClass.CleanRecords(Projid, mtoIDs, materialIDs);

            ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionSix", "var el = document.getElementById('collapseSix'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
            Session["SPMATFinalData"] = null;
            ForceButtons("btnViewFinal", true);
            ForceButtons("btnViewExported", false);
            DownloadFile efile = new DownloadFile();
            efile.filename = FileName;
            efile.contenttype = "application/vnd.ms-excel";
            // Export to browser
            using (System.IO.MemoryStream stream = GetStreamXL(workbook))
            {
                try
                {
                    stream.Position = 0;
                    using (BinaryReader br = new BinaryReader(stream))
                    {
                        byte[] bytes = br.ReadBytes((Int32)stream.Length);
                        efile.filedata = bytes;
                    }
                }
                catch { }
                DataClass.SaveExportRecordFile(FileID, efile.filedata);
                Session.Add(FileName, efile);
                btnSaveFile.CssClass = "shown";
                btnSaveFile.ToolTip = FileName;
            }
        }

        private void ForceButtons(string btnID, bool expand)
        {
            var Projid = ddlprojsel.SelectedValue;
            switch (btnID)
            {
                case "btnViewFinal":
                    List<SPMATData> mtoData = DataClass.GetMTOData(Projid, false);
                    CachedMTOData = mtoData;
                    gvFinal.EmptyDataText = "No MTO Export Data";
                    gvFinal.DataSource = mtoData;
                    gvFinal.DataBind();
                    pnlFinalMTO.Visible = true;
                    if (mtoData.Count > 0)
                    {

                        btnExportFinalMTO.Style["Display"] = "block";
                    }
                    else
                    {
                        btnExportFinalMTO.Style["Display"] = "none";
                    }
                    Session["SPMATFinalData"] = mtoData;
                    if (expand)
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionFour", "var el = document.getElementById('collapseFour'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
                    }
                    break;
                case "btnViewExported":
                    List<SPMATData> mtoData2 = DataClass.GetMTOData(Projid, true);
                    Session["CurrentFiltered"] = mtoData2;
                    Session["SelectedFilter"] = "";
                    pnlFinalMTO.Visible = true;
                    gvExported.EmptyDataText = "No Exported MTO Data";
                    gvExported.DataSource = mtoData2;
                    gvExported.DataBind();

                    List<ExportedFiles> fileData = DataClass.GetExportedFiles(Projid);
                    grFiles.EmptyDataText = "No Exported Files";
                    grFiles.DataSource = fileData;
                    grFiles.DataBind();
                    if (expand)
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionFive", "var el = document.getElementById('collapseFive'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
                    }
                    break;
                case "btnloadiso":
                    //Session["GridData"] = null;
                    ExcelGridView.DataSource = null;
                    lblerror.Text = "";
                    diverror.Style["display"] = "none";
                    if (string.IsNullOrEmpty(ddlprojects.SelectedValue) ||
                     string.IsNullOrEmpty(ddlprojsel.SelectedValue) ||
                     string.IsNullOrEmpty(ddliso.SelectedValue))
                    {
                        diverror.Style["display"] = "block";
                        lblerror.Text = "Please select a Spec, Project and ISO before loading.";
                        lblerror.ForeColor = System.Drawing.Color.Red;
                        return;
                    }
                    pnlExcelUpload.Style["display"] = "block";
                    var spec = Session["LoadedSpecs"] != null ? (List<SpecData>)Session["LoadedSpecs"] : DataClass.LoadSpectDataFromDB(ddlprojects.SelectedValue.ToString().Trim());
                    //Session["LoadedSpecs"] = spec;
                    var isosheet = ddliso.SelectedValue;
                    var projid = ddlprojsel.SelectedValue;
                    var SelectedISO = ddliso.SelectedItem.Text;
                    string isorev = "";
                    var match = System.Text.RegularExpressions.Regex.Match(SelectedISO, @"\bRev:\s*([^\s:]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        isorev = match.Groups[1].Value.Trim(); // e.g., "0", "0.1", "S0", "R0", etc.
                    }
                    List<SPMATDBData> sptest = DataClass.GetIsoSheetMTOData(isosheet, projid, isorev, false);
                    var area = "";
                    var disipline = "";
                    if (sptest.Count > 0)
                    {
                        area = sptest.Select(x => x.Area).Distinct().First().ToString();
                        disipline = sptest.Select(x => x.Discipline).Distinct().First().ToString();
                    }
                    var distinctSpecslist = spec.GroupBy(s => new { s.Lineclass, s.Ident }).Select(g => g.First()).ToList();

                    var griddata = Session["GridData"] != null ? (List<GridData>)Session["GridData"] : PopulateGridData(sptest, distinctSpecslist).Where(x => x.ISO == isosheet).ToList();
                    var orgdbdata = (List<GridData>)Session["OriginalDBData"];
                    List<string> distinctUnits = DataClass.GetUnitsByProject(projid);
                    List<string> distinctPhases = DataClass.GetPhasesByProject(projid);
                    List<string> distinctConstAreas = DataClass.GetConstAreasByProject(projid);

                    List<string> distinctSpecs = DataClass.GetSpecsByProject(projid);
                    List<string> distinctShortCodes = spec.Select(x => x.Shortcode).Distinct().ToList();
                    List<string> distinctIdents = spec.Select(x => x.Ident).Distinct().ToList();

                    // Spec → Shortcodes
                    var specShortcodeMap = spec.GroupBy(x => x.Lineclass).ToDictionary(g => g.Key, g => g.Select(x => x.Shortcode).Distinct().ToList());

                    // Shortcode → Idents
                    //Dictionary<string, Dictionary<string, List<string>>>
                    var shortCodeIdentMap = spec.GroupBy(x => x.Lineclass).ToDictionary(g => g.Key, g => g.GroupBy(s => s.Shortcode).ToDictionary(sg => sg.Key, sg => sg.Select(s => s.Ident).Distinct().ToList()));

                    // Optionally: store combinations for cascading logic
                    var unitPhaseMap = DataClass.GetUnitPhaseMap(projid);

                    var unitConstAreaMap = DataClass.GetUnitPhaseConstAreaMap(projid);
                    if (griddata.Count > 0)
                    {
                        Session["GridData"] = griddata;
                        Session["OriginalDBData"] = orgdbdata;
                        Session["CurrentIso"] = isosheet;
                        Session["CurrentArea"] = area;
                        Session["CurrentDisipline"] = disipline;
                        Session["Units"] = distinctUnits;
                        Session["Phases"] = distinctPhases;
                        Session["ConstAreas"] = distinctConstAreas;
                        Session["UnitPhaseMap"] = unitPhaseMap;
                        Session["UnitConstAreaMap"] = unitConstAreaMap;
                        Session["Specs"] = distinctSpecs;
                        Session["ShortCodes"] = distinctShortCodes;
                        Session["Idents"] = distinctIdents;
                        Session["SpecShortCodeMap"] = specShortcodeMap;
                        Session["ShortCodeIdentMap"] = shortCodeIdentMap;

                        HttpContext.Current.Items["Units"] = distinctUnits;
                        HttpContext.Current.Items["Phases"] = distinctPhases;
                        HttpContext.Current.Items["ConstAreas"] = distinctConstAreas;
                        HttpContext.Current.Items["Specs"] = distinctSpecs;
                        HttpContext.Current.Items["ShortCodes"] = distinctShortCodes;
                        HttpContext.Current.Items["Idents"] = distinctIdents;
                        BindExcelGridView(griddata, area, isosheet, bindData: true);
                        btnSubmit.Style["display"] = "block";
                        btnAddNew.Style["display"] = "block";
                    }
                    else
                    {
                        Session["GridData"] = null;
                        Session["OriginalDBData"] = null;
                        isosheet = "";
                        Session["CurrentIso"] = isosheet;
                        Session["CurrentArea"] = area;
                        Session["CurrentDisipline"] = disipline;
                        Session["Units"] = distinctUnits;
                        Session["Phases"] = distinctPhases;
                        Session["ConstAreas"] = distinctConstAreas;
                        Session["UnitPhaseMap"] = unitPhaseMap;
                        Session["UnitConstAreaMap"] = unitConstAreaMap;
                        Session["Specs"] = distinctSpecs;
                        Session["ShortCodes"] = distinctShortCodes;
                        Session["Idents"] = distinctIdents;
                        Session["SpecShortCodeMap"] = specShortcodeMap;
                        Session["ShortCodeIdentMap"] = shortCodeIdentMap;
                        ddliso.ClearSelection();
                        HttpContext.Current.Items["Units"] = distinctUnits;
                        HttpContext.Current.Items["Phases"] = distinctPhases;
                        HttpContext.Current.Items["ConstAreas"] = distinctConstAreas;
                        HttpContext.Current.Items["Specs"] = distinctSpecs;
                        HttpContext.Current.Items["ShortCodes"] = distinctShortCodes;
                        HttpContext.Current.Items["Idents"] = distinctIdents;
                        BindExcelGridView(griddata, area, isosheet, bindData: true);
                        btnSubmit.Style["display"] = "none";
                        btnAddNew.Style["display"] = "none";
                        ExcelGridView.Visible = false;
                        pnlExcelUpload.Visible = false;
                        ReloadIso();
                    }


                    break;
                //case "btnviewspmat":
                //    List<SPMATDBData> mtoData3 = DataClass.GetWorkingMTOData(Projid);

                //    gvSPMAT.EmptyDataText = "No Working Data";
                //    gvSPMAT.DataSource = mtoData3;
                //    gvSPMAT.DataBind();
                //    pnlSPMATView.Visible = true;
                //    if (mtoData3.Count > 0)
                //    {
                //        btnExportSPMAT.Style["Display"] = "block";
                //    }
                //    else
                //    {
                //        btnExportSPMAT.Style["Display"] = "none";
                //    }
                //    Session["SPMATExportData"] = mtoData3;
                //    if (expand)
                //    {
                //        ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionTwo", "var el = document.getElementById('collapseTwo'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
                //    }
                //    break;
                case "btnviewInterim":
                    List<SPMATIntrimData> mtoData4 = DataClass.GetMTOIntrimData(Projid);

                    gvInterim.EmptyDataText = "No Working MTO Data";
                    gvInterim.DataSource = mtoData4;
                    gvInterim.DataBind();
                    pnlViewInterimData.Visible = true;
                    if (mtoData4.Count > 0)
                    {
                        lblintrim.Style["Display"] = "block";
                        btnMoveToFinal.Style["Display"] = "block";
                    }
                    else
                    {
                        lblintrim.Style["Display"] = "none";
                        btnMoveToFinal.Style["Display"] = "none";
                    }
                    Session["SPMATIntrimData"] = mtoData4;
                    if (expand)
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionThree", "var el = document.getElementById('collapseThree'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
                    }
                    break;
                case "btnMTOMaintenance":
                    List<SPMATDeletedData> mtoData5 = DataClass.GetMaintenanceData(Projid);

                    gvmtorev.EmptyDataText = "No Maintenance Data";
                    gvmtorev.DataSource = mtoData5;
                    gvmtorev.DataBind();
                    if (expand)
                    {
                        ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionSix", "var el = document.getElementById('collapseSix'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
                    }
                    break;
                case "btnExportSPMAT":
                    diverror.Style["display"] = "none";
                    lblerror.Text = "";
                    var data = Session["SPMATExportData"] as List<SPMATDBData>;
                    if (data == null || !data.Any()) return;
                    List<int> materialIDs = data.Select(d => d.MaterialID).Distinct().ToList();
                    MoveToIntrim(data, materialIDs);
                    // gvSPMAT.EmptyDataText = "No Working Data";
                    // gvSPMAT.DataSource = null;
                    // gvSPMAT.DataBind();
                    // pnlSPMATView.Visible = false;
                    // btnExportSPMAT.Style["Display"] = "none";
                    //btnviewspmat_Click(null, null);
                    //ForceButtons("btnviewspmat", false);
                    // btnviewInterim_Click(null, null);
                    break;

            }


        }

        protected void btnViewExported_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            pnlFileMan.Visible = true;
            var Projid = ddlprojsel.SelectedValue;
            List<SPMATData> mtoData = DataClass.GetMTOData(Projid, true);
            Session["CurrentFiltered"] = mtoData;
            Session["SelectedFilter"] = "";
            pnlFinalMTO.Visible = true;
            CachedMTOData = mtoData;
            gvExported.EmptyDataText = "No Exported MTO Data";
            gvExported.DataSource = mtoData;
            gvExported.DataBind();
            PopulateIsoDropdown();
            List<ExportedFiles> fileData = DataClass.GetExportedFiles(Projid);
            grFiles.EmptyDataText = "No Exported Files";
            grFiles.DataSource = fileData;
            grFiles.DataBind();

            ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionFive", "var el = document.getElementById('collapseFive'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
            // Store in session for export
            if (Session["GridData"] != null && Session["CurrentIso"] != null && Session["CurrentArea"] != null)
            {
                var griddata = Session["GridData"] as List<GridData>;
                var area = Session["CurrentArea"].ToString();
                var iso = Session["CurrentIso"].ToString();
                BindExcelGridView(griddata, area, iso, bindData: true);
            }
        }

        protected void grFiles_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "Import")
            {
                diverror.Style["display"] = "none";
                lblerror.Text = "";
                int fileId = Convert.ToInt32(e.CommandArgument);
                GridViewRow row = ((Button)e.CommandSource).NamingContainer as GridViewRow;
                TextBox txtCode = (TextBox)row.FindControl("txtImportCode");

                string importCode = txtCode.Text.Trim();

                if (string.IsNullOrEmpty(importCode) || importCode.Length > 5)
                {
                    diverror.Style["display"] = "block";
                    lblerror.Text = "Please Enter a Valid Import Status (P Code from MTO Import).";
                    lblerror.ForeColor = System.Drawing.Color.Red;
                    return;
                }

                string fileMTOIDs = grFiles.DataKeys[row.RowIndex]["FileMTOIDs"].ToString();


                if (!string.IsNullOrEmpty(importCode) && importCode.Length <= 5)
                {
                    // TODO: Update the database with the import code and mark FileCompleted = true
                    // Example:
                    var projid = Session["ENGPROJECTID"].ToString();
                   var ue= DataClass.UpdateSPMAT_FileExports(fileId, fileMTOIDs, importCode, int.Parse(projid));
                    if (!string.IsNullOrEmpty(ue))
                    {
                        diverror.Style["display"] = "block";
                        lblerror.Text = "Error on Updating Data: "+ue.Trim();
                        lblerror.ForeColor = System.Drawing.Color.Red;
                        return;
                    }
                    // Rebind the GridView after update
                    //btnViewFinal_Click(null, null);
                    //btnViewExported_Click(null,null);
                    ForceButtons("btnViewFinal", false);
                    ForceButtons("btnViewExported", true);

                }
            }
        }

        protected void btnMTOMaintenance_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            if (string.IsNullOrEmpty(ddlprojsel.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select a Spec and  Project before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }
            var Projid = ddlprojsel.SelectedValue;
            List<SPMATDeletedData> mtoData = DataClass.GetMaintenanceData(Projid);

            gvmtorev.EmptyDataText = "No Maintenance Data";
            gvmtorev.DataSource = mtoData;
            gvmtorev.DataBind();

            ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionSix", "var el = document.getElementById('collapseSix'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
            // Store in session for export
            if (Session["GridData"] != null && Session["CurrentIso"] != null && Session["CurrentArea"] != null)
            {
                var griddata = Session["GridData"] as List<GridData>;
                var area = Session["CurrentArea"].ToString();
                var iso = Session["CurrentIso"].ToString();
                BindExcelGridView(griddata, area, iso, bindData: true);
            }
        }

        protected void gvmtorev_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "RemoveRow")
            {
                int rowIndex = Convert.ToInt32(e.CommandArgument);
                GridViewRow row = gvInterim.Rows[rowIndex];

                var keys = gvInterim.DataKeys[row.RowIndex];
                string MTOID = keys["MTOID"].ToString();
                //Maybe History Table
                //Delete from SPMAT_MTOData


                // Refresh the grid
                //btnMTOMaintenance_Click(null, null);
                ForceButtons("btnMTOMaintenance", true);
            }
        }

        protected void btnSaveFile_Click(object sender, EventArgs e)
        {
            ForceButtons("btnViewFinal", true);
            var sesname = btnSaveFile.ToolTip.ToString().Trim();
            if (string.IsNullOrEmpty(sesname))
            {
                lblMessage.Text = "No file to download or file already downloaded.";
                lblMessage.ForeColor = System.Drawing.Color.Red;
                btnSaveFile.CssClass = "hidden";
                return;
            }
            DataClass.DownloadFile dl = (DataClass.DownloadFile)Session[sesname];
            if (dl == null)
            {
                lblMessage.Text = "No file to download or file already downloaded.";
                lblMessage.ForeColor = System.Drawing.Color.Red;
                btnSaveFile.CssClass = "hidden";
                return;
            }
            try
            {
                Session.Remove(sesname);
                Response.Clear();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=" + dl.filename);
                Response.ContentType = dl.contenttype;
                Response.BinaryWrite(dl.filedata.ToArray());
                Response.Flush();
                Response.End();
            }
            catch { }

        }

        public MemoryStream GetStreamXL(IXLWorkbook excelWorkbook)
        {
            MemoryStream fs = new MemoryStream();
            excelWorkbook.SaveAs(fs);
            fs.Position = 0;
            return fs;
        }

        //protected void btnrefreshIso_Click(object sender, EventArgs e)
        //{
        //    diverror.Style["display"] = "none";
        //    lblerror.Text = "";
        //    var Projid = ddlprojsel.SelectedValue;
        //    var spec = DataClass.LoadSpectDataFromDB(ddlprojects.SelectedValue.ToString().Trim());
        //    string isoaccesspath = DataClass.GetIsoAccess(Projid);
        //    var IsoData = DataClass.GetIsoData(isoaccesspath);
        //    foreach (var i in IsoData)
        //    {
        //        DataClass.RefreshISO(i);
        //    }
        //    var projid = ddlprojsel.SelectedValue;
        //    List<DDLList> iso = DataClass.GetProjectISO(projid, false);
        //    if (iso.Count > 0)
        //    {
        //        ddliso.DataSource = iso;
        //        ddliso.DataTextField = "DDLListName";
        //        ddliso.DataValueField = "DDLList_ID";
        //        ddliso.DataBind();
        //        // btnviewspmat.CssClass = "shown";
        //        btnviewInterim.CssClass = "shown";
        //        btnViewFinal.CssClass = "shown";
        //        diverror.Style["display"] = "block";
        //        lblerror.Text = "Iso Data Refreshed.";
        //        lblerror.ForeColor = System.Drawing.Color.Blue;
        //    }
        //    else
        //    {
        //        diverror.Style["display"] = "block";
        //        lblerror.Text = "No Iso in correct state.";
        //        lblerror.ForeColor = System.Drawing.Color.Red;
        //        return;
        //    }
        //}

        protected void gvFinal_RowDataBound(object sender, GridViewRowEventArgs e)
        {

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                // Replace with your actual username check
                var EID = int.Parse(Session["EID"].ToString());
                bool btnvis = false;
                if (EID == 447 || EID == 240)
                {
                    btnvis = true;
                }
                else
                {
                    btnvis = false;
                }
                Button btnRemove = (Button)e.Row.FindControl("btnRemoveFinal");
                if (btnRemove != null)
                {
                    btnRemove.Visible = btnvis;
                }

            }
        }

        protected void gvFinal_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "RemoveFinalRow")
            {
                int fileId = Convert.ToInt32(e.CommandArgument);
                GridViewRow row = ((Button)e.CommandSource).NamingContainer as GridViewRow;

                var keys = gvFinal.DataKeys[row.RowIndex];
                string MTOID = keys["MTOID"].ToString();
                string ISO = keys["ISO"].ToString();
                int IsoUniqeRevID = int.Parse(keys["IsoUniqeRevID"].ToString());
                DataClass.RemoveFromFinal(MTOID, ISO, IsoUniqeRevID);

                ReloadIso();
                // Refresh the grid
                //btnviewInterim_Click(null, null);
                ForceButtons("btnViewFinal", true);
            }
        }

        protected void grFiles_RowDataBound(object sender, GridViewRowEventArgs e)
        {

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                var fileData = DataBinder.Eval(e.Row.DataItem, "FileData");
                var lnkDownload = (HyperLink)e.Row.FindControl("lnkDownload");

                if (fileData != DBNull.Value && fileData != null)
                {
                    lnkDownload.Visible = true;
                }
            }

        }

        protected void gvExported_RowCreated(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.Header)
            {
                DropDownList ddlIsoFilter = (DropDownList)e.Row.FindControl("ddlIsoFilter");
                if (ddlIsoFilter != null)
                {
                    if (CachedMTOData != null)
                    {
                        var isoList = CachedMTOData
                            .Select(x => x.ISO)
                            .Where(s => !string.IsNullOrEmpty(s))
                            .Distinct()
                            .OrderBy(s => s)
                            .ToList();

                        ddlIsoFilter.Items.Clear();
                        ddlIsoFilter.Items.Add(new ListItem("-- All --", ""));
                        foreach (var iso in isoList)
                        {
                            ddlIsoFilter.Items.Add(new ListItem(iso, iso));
                        }
                    }

                }
            }
        }

        private void PopulateIsoDropdown(string selectedValue = "")
        {
            if (gvExported.HeaderRow != null)
            {
                DropDownList ddlIsoFilter = (DropDownList)gvExported.HeaderRow.FindControl("ddlIsoFilter");
                if (ddlIsoFilter != null && CachedMTOData != null)
                {
                    var isoList = CachedMTOData
                        .Select(x => x.ISO)
                        .Where(s => !string.IsNullOrEmpty(s))
                        .Distinct()
                        .OrderBy(s => s)
                        .ToList();

                    ddlIsoFilter.Items.Clear();
                    ddlIsoFilter.Items.Add(new ListItem("-- All --", ""));
                    foreach (var iso in isoList)
                    {
                        ddlIsoFilter.Items.Add(new ListItem(iso, iso));
                    }
                    // Set selected value
                    if (!string.IsNullOrEmpty(selectedValue) && ddlIsoFilter.Items.FindByValue(selectedValue) != null)
                    {
                        ddlIsoFilter.SelectedValue = selectedValue;
                    }
                }
            }
        }

        protected void ddlIsoFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddl = (DropDownList)sender;
            string selectedIso = ddl.SelectedValue;

            var filtered = string.IsNullOrEmpty(selectedIso)
                ? CachedMTOData
                : CachedMTOData.Where(x => x.ISO == selectedIso).ToList();
            Session["CurrentFiltered"] = filtered;
            Session["SelectedFilter"] = selectedIso;
            gvExported.DataSource = filtered;
            gvExported.DataBind();
            PopulateIsoDropdown(selectedIso); // Re-populate dropdown
            // Re-expand accordion
            ScriptManager.RegisterStartupScript(this, GetType(), "expandAccordionFive",
                "var el = document.getElementById('collapseFive'); var bsCollapse = new bootstrap.Collapse(el, {toggle: true});", true);
        }

        protected void btnExportFiltered_Click(object sender, EventArgs e)
        {
            var filteredData = (List<SPMATData>)Session["CurrentFiltered"]; // Or use filtered list if stored separately
            var filtername = string.IsNullOrEmpty(Session["SelectedFilter"].ToString()) ? "" : Session["SelectedFilter"].ToString();
            if (filteredData == null || !filteredData.Any())
                return;

            string FileName = "MTOExport";
            string sheetname = "MTOData";
            if (!string.IsNullOrEmpty(filtername))
            {
                FileName = "Filtered_" + FileName + "_" + filtername;
                sheetname += "_" + filtername;
            }
            FileName += "_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx";
            // Convert to DataTable
            DataTable dt = ConvertSPMATDataToDataTable(filteredData);

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, sheetname);

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    stream.Position = 0;

                    Response.Clear();
                    Response.Buffer = true;
                    Response.AddHeader("content-disposition", "attachment; filename=" + FileName);
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.BinaryWrite(stream.ToArray());
                    Response.Flush();
                    Response.End();
                }
            }

        }
    }
}