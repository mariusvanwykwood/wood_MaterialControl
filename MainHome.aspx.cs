using System;
using System.Security.Principal;
using System.Web;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI.WebControls;
using static Wood_MaterialControl.DataClass;
using ClosedXML.Excel;
using System.Data;
using System.IO;


namespace Wood_MaterialControl
{
    public partial class MainHome : System.Web.UI.Page
    {
        protected void Page_Init(object sender, EventArgs e)
        {
            if (ViewState["GridData"] != null && ViewState["CurrentIso"] != null && ViewState["CurrentArea"] != null)
            {
                var griddata = ViewState["GridData"] as List<GridData>;
                var area = ViewState["CurrentArea"].ToString();
                var iso = ViewState["CurrentIso"].ToString();

                // Rebuild columns only, no data binding
                BindExcelGridView(griddata, area, iso, bindData: false);
            }
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
                    if (ViewState["GridData"] != null)
                    {
                        var griddata = ViewState["GridData"] as List<GridData>;
                        var area = ViewState["CurrentArea"].ToString();
                        var iso = ViewState["CurrentIso"].ToString();
                        BindExcelGridView(griddata, area, iso, bindData: true);
                    }

                    // Always set HttpContext items for templates
                    HttpContext.Current.Items["Units"] = ViewState["Units"];
                    HttpContext.Current.Items["Phases"] = ViewState["Phases"];
                    HttpContext.Current.Items["ConstAreas"] = ViewState["ConstAreas"];


                }
                else if (IsPostBack)
                {
                    HttpContext.Current.Items["Units"] = ViewState["Units"];
                    HttpContext.Current.Items["Phases"] = ViewState["Phases"];
                    HttpContext.Current.Items["ConstAreas"] = ViewState["ConstAreas"];
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
                        eventTarget.Contains("btnAddNew") || eventTarget.Contains("btnSaveNew") || eventTarget.Contains("btnCancelAdd")|| !string.IsNullOrEmpty(clickedButton) || !string.IsNullOrEmpty(changedDropdown)
                         )
                        )
                    {
                        var griddata = ViewState["GridData"] as List<GridData>;
                        var area = ViewState["CurrentArea"].ToString();
                        var iso = ViewState["CurrentIso"].ToString();
                        BindExcelGridView(griddata, area, iso, bindData: true);
                    }



                }
            }
            catch
            {
            }
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
                var tempgriddata =ViewState["TmpGridData"] as List<GridData> ?? new List<GridData>();
                if(tempgriddata.Where(x=>x.MaterialID==materialID).Any())
                {
                    var itemToremove = tempgriddata.SingleOrDefault(x => x.MaterialID == materialID);
                    tempgriddata.Remove(itemToremove);
                }
                ViewState["TmpGridData"] = tempgriddata;
                // Mark as Checked in DB
                DataClass.MarkMaterialAsChecked(materialID);

                // Reload grid
                btnloadiso_Click(null, null);
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
            if (ddlprojsel.SelectedValue != "")
            {
                //Get Data for Iso DLL
                var projid = ddlprojsel.SelectedValue;
                List<DDLList> iso = DataClass.GetProjectISO(projid,false);
                ddliso.DataSource = iso;
                ddliso.DataTextField = "DDLListName";
                ddliso.DataValueField = "DDLList_ID";
                ddliso.DataBind();
                btnviewspmat.CssClass = "shown";
            }

        }

        protected void btnloadiso_Click(object sender, EventArgs e)
        {
            ViewState["GridData"] = null;
            ExcelGridView.DataSource = null;
            lblerror.Text = "";
            diverror.Style["display"] = "none";
            if (string.IsNullOrEmpty(ddlprojects.SelectedValue) ||
             string.IsNullOrEmpty(ddlprojsel.SelectedValue) ||
             string.IsNullOrEmpty(ddliso.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select a project, subproject, and ISO before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }

            var spec = DataClass.LoadSpectDataFromDB(ddlprojects.SelectedValue.ToString().Trim());
            Session["LoadedSpecs"] = spec;
            var isosheet = ddliso.SelectedValue;
            var projid = ddlprojsel.SelectedValue;
            List<SPMATDBData> sptest = DataClass.GetIsoSheetMTOData(isosheet, projid,false);
            var area = sptest.Select(x => x.Area).Distinct().First().ToString();
            var disipline = sptest.Select(x => x.Discipline).Distinct().First().ToString();

            var distinctSpecslist = spec.GroupBy(s => new { s.Lineclass, s.Ident }).Select(g => g.First()).ToList();

            var griddata = PopulateGridData(sptest, distinctSpecslist).Where(x => x.ISO == isosheet).ToList();
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
            ViewState["GridData"] = griddata;
            ViewState["OriginalDBData"] = griddata;
            ViewState["CurrentIso"] = isosheet;
            ViewState["CurrentArea"] = area;
            ViewState["CurrentDisipline"] = disipline;
            ViewState["Units"] = distinctUnits;
            ViewState["Phases"] = distinctPhases;
            ViewState["ConstAreas"] = distinctConstAreas;
            ViewState["UnitPhaseMap"] = unitPhaseMap;
            ViewState["UnitConstAreaMap"] = unitConstAreaMap;
            ViewState["Specs"] = distinctSpecs;
            ViewState["ShortCodes"] = distinctShortCodes;
            ViewState["Idents"] = distinctIdents;
            ViewState["SpecShortCodeMap"] = specShortcodeMap;
            ViewState["ShortCodeIdentMap"] = shortCodeIdentMap;

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
                                    IsLocked=spmat.Lock

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
                    IsLocked = row["IsLocked"]?.ToString()
                };

                list.Add(item);
            }

            return list;
        }

        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            bool allChecked = true;
            List<SPMATDBData> updatedItems = new List<SPMATDBData>();

            var loadedSpecs = Session["LoadedSpecs"] as List<SpecData>;
            var griddata = ViewState["GridData"] as List<GridData>;
            var dbdata = ViewState["OriginalDBData"] as List<GridData>;
            var area = ViewState["CurrentArea"]?.ToString();
            var disipline =   ViewState["CurrentDisipline"]?.ToString();
            var iso = ViewState["CurrentIso"]?.ToString();
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

                string selectedText=changedText.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? txtQty?.Text;
                string selectedUnit = changedUnits.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlUnit?.SelectedValue;
                string selectedPhase = changedPhases.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlPhase?.SelectedValue;
                string selectedConstArea = changedConstAreas.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlConstArea?.SelectedValue;
                string selectedSpec = changedSpecs.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlSpec?.SelectedValue;
                string selectedShortcode = changedShortcodes.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlShortcode?.SelectedValue;
                string selectedIdent = changedIdents.FirstOrDefault(d => d.ContainsKey(materialID))?[materialID] ?? ddlIdent?.SelectedValue;

                if(DecParse(selectedText).HasError)
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
                var orgdbdata= dbdata?.FirstOrDefault(g => g.MaterialID == materialID);
                var originalData = griddata?.FirstOrDefault(g => g.MaterialID == materialID);
                var havediffs = false;
                if (orgdbdata != null && originalData != null)
                {
                    havediffs=AreObjectsDifferentExcept(orgdbdata, originalData);
                }
                else if(orgdbdata==null)
                {
                    havediffs = true;
                }
                if (originalData != null)
                {
                    updatedItems.Add(new SPMATDBData
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
                        IsoRevision = havediffs ? "" : orgdbdata?.IsoRevision,
                        IsoRevisionDate = havediffs ? "" : orgdbdata?.IsoRevisionDate,
                        Lock = havediffs ? "" : orgdbdata?.IsLocked,
                        Code = havediffs ? "M" : orgdbdata?.Source
                    });
                }
            }

            if (!allChecked)
            {
                lblMessage.Text = "Please check all rows before submitting.";
                lblMessage.ForeColor = System.Drawing.Color.Red;
                return;
            }

            // Update database
            foreach (var item in updatedItems)
            {
                DataClass.FinalizeMaterialUpdate(item);
            }

            lblMessage.ForeColor = System.Drawing.Color.Green;
            lblMessage.Text = "All quantities updated successfully.";
            var eid = 0;
            List<DDLList> clientlist = new List<DDLList>();
            try
            {
                eid = int.Parse(Session["EID"].ToString());
                clientlist =  (List<DDLList>)Session["ClientList"];
            }
            catch { }
            Session.Clear();
            ViewState.Clear();
            Session["EID"] = eid;
            Session["ClientList"] = clientlist;
            ddliso.ClearSelection();
            ExcelGridView.DataSource = null;
            ExcelGridView.DataBind();
            lblFileName.Text = "";
            btnSubmit.Style["display"] = "none";
            btnAddNew.Style["display"] = "none";
            btnCancelAdd_Click(null, null);
            ReloadIso();
            

        }

        private void ReloadIso()
        {
            var projid = ddlprojsel.SelectedValue;
            List<DDLList> iso = DataClass.GetProjectISO(projid,false);
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

                // Get mappings from ViewState
                var unitPhaseMap = ViewState["UnitPhaseMap"] as Dictionary<string, List<string>>;
                var unitConstAreaMap = ViewState["UnitConstAreaMap"] as Dictionary<string, Dictionary<string, List<string>>>;
                var specShortCodeMap = ViewState["SpecShortCodeMap"] as Dictionary<string, List<string>>;
                var shortCodeIdentMap = ViewState["ShortCodeIdentMap"] as Dictionary<string, Dictionary<string, List<string>>>;

                var allUnits = ViewState["Units"] as List<string>;
                var allSpecs = ViewState["Specs"] as List<string>;

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
                    }catch { }
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
                    }catch { }
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
                       catch {}
                        ddlConstArea.Items.Insert(0, new ListItem("-- Select --", ""));
                        ddlConstArea.AutoPostBack = true;
                        ddlConstArea.SelectedIndexChanged += ddlConstArea_SelectedIndexChanged;
                        if (validConstAreas.Contains(currentConstArea))
                        {
                            ddlConstArea.SelectedValue = currentConstArea;
                        }
                    }
                }
                catch(Exception ae)
                {
                    var xx = ae.Message;
                }
                if (ddlSpec != null && !string.IsNullOrEmpty(selectedSpec))
                {
                    ddlSpec.DataSource = allSpecs;
                    try
                    {
                        ddlSpec.DataBind();
                    } catch {}
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
                    }catch {}
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
                    }catch {}
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
                    }catch { }
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
                if(txtQty != null)
                {
                    if(currentText!=selectedText && selectedText!="")
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
            if (ViewState["TmpGridData"] != null)
            {
                var tmpgd = ViewState["TmpGridData"] as List<GridData>;
                var existid=griddata.Select(x => x.MaterialID).Distinct().ToList();
                var notinlistalready=tmpgd.Where(x => !existid.Contains(x.MaterialID)).ToList();
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
            var hiddenColumns = new HashSet<string> { "MaterialID", "ProjectID", "Discipline", "Area", "ISO", "Fabrication_Type","IsoRevision","IsoRevisionDate", "IsLocked" };
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
            System.Web.UI.WebControls.TextBox txt = (System.Web.UI.WebControls.TextBox)sender;
            GridViewRow row = (GridViewRow)txt.NamingContainer;
            int materialID = Convert.ToInt32(ExcelGridView.DataKeys[row.RowIndex].Value);
            var list = Session["ChangedText"] as List<Dictionary<int, string>> ?? new List<Dictionary<int, string>>();
           var tmptxt=txt.Text.Trim();
            var existingEntry = list.FirstOrDefault(d => d.ContainsKey(materialID));
            if (existingEntry != null)
            {
                existingEntry[materialID] = tmptxt;
            }
            else
            {
                list.Add(new Dictionary<int, string> { { materialID, tmptxt } });
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
            lblMessage.Text = "";
            pnlAddForm.Visible = true;

            ddlAddUnit.DataSource = ViewState["Units"] as List<string>;
            ddlAddUnit.DataBind();
            ddlAddUnit.Items.Insert(0, new ListItem("-- Select --", ""));

            ddlAddSpec.DataSource = ViewState["Specs"] as List<string>;
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

            var griddata = ViewState["GridData"] as List<GridData> ?? new List<GridData>();
            var specList = Session["LoadedSpecs"] as List<SpecData> ?? new List<SpecData>();
            var tempgriddata = new List<GridData>();


            if (Session["TempMaterialID"] == null)
            {
                Session["TempMaterialID"] = -1;
            }
            int tempID = (int)Session["TempMaterialID"];
            Session["TempMaterialID"] = tempID - 1;

            // Create a temporary SPMATDBData object
            var newSPMAT = new SPMATDBData
            {
                MaterialID = tempID, // Temporary ID
                ProjectID = int.Parse(ddlprojsel.SelectedValue),
                Discipline = "", // Optional
                Area = "", // Optional
                Unit = ddlAddUnit.SelectedValue,
                Phase = ddlAddPhase.SelectedValue,
                Const_Area = ddlAddConstArea.SelectedValue,
                ISO = ddliso.SelectedValue,
                Ident_no = ddlAddIdent.SelectedValue,
                qty = txtAddQty.Text.Trim(),
                qty_unit = "pc",
                Fabrication_Type = "Undefined",
                Spec = ddlAddSpec.SelectedValue,
                IsoRevision = "",
                IsoRevisionDate = "",
                Lock = "",
                Code = "M"
            };

            // Use PopulateGridData to enrich the entry
            var enriched = PopulateGridData(new List<SPMATDBData> { newSPMAT }, specList);
            if (enriched.Any())
            {
                griddata.Add(enriched.First());
            }

            ViewState["GridData"] = griddata;
            tempgriddata = griddata;
            ViewState["TmpGridData"] = tempgriddata;

            // Clear form fields for next entry
            ddlAddUnit.ClearSelection();
            ddlAddPhase.ClearSelection();
            ddlAddConstArea.ClearSelection();
            ddlAddSpec.ClearSelection();
            ddlAddShortcode.ClearSelection();
            ddlAddIdent.ClearSelection();
            txtAddQty.Text = "";

            // Keep form visible for next entry
            pnlAddForm.Visible = true;


            BindExcelGridView(griddata, ViewState["CurrentArea"].ToString(), ViewState["CurrentIso"].ToString(), bindData: true);

        }
    
        protected void ddlAddUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            var unit = ddlAddUnit.SelectedValue;
            var unitPhaseMap = ViewState["UnitPhaseMap"] as Dictionary<string, List<string>>;
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
            var unitConstAreaMap = ViewState["UnitConstAreaMap"] as Dictionary<string, Dictionary<string, List<string>>>;
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
            var specShortCodeMap = ViewState["SpecShortCodeMap"] as Dictionary<string, List<string>>;
            if (specShortCodeMap != null && specShortCodeMap.ContainsKey(spec))
            {
                ddlAddShortcode.DataSource = specShortCodeMap[spec];
                ddlAddShortcode.DataBind();
                ddlAddShortcode.Items.Insert(0, new ListItem("-- Select --", ""));
            }

            ddlAddIdent.Items.Clear();
        }
        protected void ddlAddShortcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            var spec = ddlAddSpec.SelectedValue;
            var shortcode = ddlAddShortcode.SelectedValue;
            var shortCodeIdentMap = ViewState["ShortCodeIdentMap"] as Dictionary<string, Dictionary<string, List<string>>>;
            if (shortCodeIdentMap != null && shortCodeIdentMap.ContainsKey(spec) && shortCodeIdentMap[spec].ContainsKey(shortcode))
            {
                ddlAddIdent.DataSource = shortCodeIdentMap[spec][shortcode];
                ddlAddIdent.DataBind();
                ddlAddIdent.Items.Insert(0, new ListItem("-- Select --", ""));
            }
        }

        #endregion

        protected void btnviewspmat_Click(object sender, EventArgs e)
        {
            var Projid= ddlprojsel.SelectedValue;
            List<SPMATData> mtoData = DataClass.GetMTOData(Projid);

            gvSPMAT.DataSource = mtoData;
            gvSPMAT.DataBind();
            pnlSPMATView.Visible = true;

            // Store in session for export
            Session["SPMATExportData"] = mtoData;

        }

        protected void btnExportSPMAT_Click(object sender, EventArgs e)
        {
            var data = Session["SPMATExportData"] as List<SPMATData>;
            if (data == null || !data.Any()) return;

            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("SPMAT_Export");

                // Define headers
                var headers = new[]
                {
            "Discipline", "Area", "Unit", "Phase", "Const_Area", "ISO", "Ident_no",
            "qty", "qty_unit", "Fabrication_Type", "Spec", "Pos",
            "IsoRevisionDate", "IsoRevision", "IsLocked"
        };

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
                    worksheet.Cell(row, 1).Value = item.Discipline;
                    worksheet.Cell(row, 2).Value = item.Area;
                    worksheet.Cell(row, 3).Value = item.Unit;
                    worksheet.Cell(row, 4).Value = item.Phase;
                    worksheet.Cell(row, 5).Value = item.Const_Area;
                    worksheet.Cell(row, 6).Value = item.ISO;
                    worksheet.Cell(row, 7).Value = item.Ident_no;
                    worksheet.Cell(row, 8).Value = item.qty;
                    worksheet.Cell(row, 9).Value = item.qty_unit;
                    worksheet.Cell(row, 10).Value = item.Fabrication_Type;
                    worksheet.Cell(row, 11).Value = item.Spec;
                    worksheet.Cell(row, 12).Value = item.Pos;
                    worksheet.Cell(row, 13).Value = item.IsoRevisionDate;
                    worksheet.Cell(row, 14).Value = item.IsoRevision;
                    worksheet.Cell(row, 15).Value = item.IsLocked;

                    for (int col = 1; col <= headers.Length; col++)
                    {
                        worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        worksheet.Cell(row, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }

                    row++;
                }

                // Adjust column widths
                worksheet.Columns().AdjustToContents();
                var FileName = ddlprojsel.SelectedItem.Text.Split('-')[0].Trim() + "_MTO_ALL_SPMAT_" + DateTime.Now.ToString("yyMMdd") + ".xlsx";
                // Export to browser
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    Response.Clear();
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment;filename="+FileName);
                    Response.BinaryWrite(stream.ToArray());
                    Response.End();
                }
            }
        }

        protected void gvSPMAT_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName == "RemoveRow")
            {
                int rowIndex = Convert.ToInt32(e.CommandArgument);
                GridViewRow row = gvSPMAT.Rows[rowIndex];

                var keys = gvSPMAT.DataKeys[row.RowIndex];
                string discipline = keys["Discipline"].ToString();
                string area = keys["Area"].ToString();
                string unit = keys["Unit"].ToString();
                string phase = keys["Phase"].ToString();
                string constArea = keys["Const_Area"].ToString();
                string iso = keys["ISO"].ToString();
                string ident = keys["Ident_no"].ToString();
                string spec = keys["Spec"].ToString();

                // Refactored: Use DataClass methods instead of inline SQL
                DataClass.DeleteMTOEntry(discipline, area, unit, phase, constArea, iso, ident, spec);
                DataClass.UncheckREQEntry(discipline, area, unit, phase, constArea, iso, ident, spec);

                ReloadIso();
                // Refresh the grid
                btnviewspmat_Click(null, null);
            }
        }

    }
}