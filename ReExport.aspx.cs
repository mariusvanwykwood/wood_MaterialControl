using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ClosedXML.Excel;
using static Wood_MaterialControl.DataClass;

namespace Wood_MaterialControl
{
    public partial class ReExport : System.Web.UI.Page
    {
        protected void Page_Init(object sender, EventArgs e)
        {
            if (Session["IsoReviewData"] != null )
            {
                var griddata = Session["IsoReviewData"] as List<SPMATReExportData>;
               
                // Rebuild columns only, no data binding
                BindReExportGrid(griddata);
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                SetNoCache();
                if (!IsPostBack)
                {
                    var EID = -1;
                    try { EID = int.Parse(Session["EID"].ToString()); } catch { }
                    if (EID > -1 && EID != 0)
                    {
                        Session["EID"] = EID;
                        if (EID != 240) Response.Redirect("Default.aspx");
                    }
                    else
                    {
                        Session.Clear();
                        Response.Redirect("Default.aspx?UF=1");
                    }

                    var projid = Session["ENGPROJECTID"].ToString();
                    List<DDLList> iso = DataClass.GetProjectISORevData(projid);
                    if (iso.Count > 0)
                    {
                        ddliso.DataSource = iso;
                        ddliso.DataTextField = "DDLListName";
                        ddliso.DataValueField = "DDLList_ID";
                        ddliso.DataBind();
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
            catch { }
        }

        private void SetNoCache()
        {
            HttpContext.Current.Response.Cache.SetAllowResponseInBrowserHistory(false);
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.Cache.SetNoStore();
            Response.Cache.SetExpires(DateTime.Now);
            Response.Cache.SetValidUntilExpires(true);
        }

        protected void btnloadiso_Click(object sender, EventArgs e)
        {
            lblerror.Text = "";
            diverror.Style["display"] = "none";
            if (string.IsNullOrEmpty(ddliso.SelectedValue))
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "Please select an ISO before loading.";
                lblerror.ForeColor = System.Drawing.Color.Red;
                return;
            }

            pnlisodata.Style["display"] = "block";
            var isosheet = ddliso.SelectedValue;
            var projid = Session["ENGPROJECTID"].ToString();
            Session["REISOSHEET"] = isosheet;
            List<SPMATReExportData> sptest = DataClass.GetReExportMTOData(isosheet, projid);
            Session["IsoReviewData"] = sptest;

            if (sptest != null && sptest.Count > 0)
            {
                List<string> distinctUnits = DataClass.GetUnitsByProject(projid);
                List<string> distinctPhases = DataClass.GetPhasesByProject(projid);
                List<string> distinctConstAreas = DataClass.GetConstAreasByProject(projid);
                List<string> distinctRevisions = DataClass.GetRevisionsByProject(projid);
                Session["Units"] = distinctUnits;
                Session["Phases"] = distinctPhases;
                Session["ConstAreas"] = distinctConstAreas;
                Session["Revisions"] = distinctRevisions;
                Session["UnitPhaseMap"] = DataClass.GetUnitPhaseMap(projid);
                Session["UnitConstAreaMap"] = DataClass.GetUnitPhaseConstAreaMap(projid);
                HttpContext.Current.Items["Units"] = distinctUnits;
                HttpContext.Current.Items["Phases"] = distinctPhases;
                HttpContext.Current.Items["ConstAreas"] = distinctConstAreas;
                HttpContext.Current.Items["Revisions"] = distinctRevisions;

                BindReExportGrid(sptest);
                lblFileName.Text = $"Review for ISO: {isosheet}";
                grisoreview.Visible = true;
                pnlisodata.Visible = true;
                btnUpdate_ISO.Visible = true;


            }
            else
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "No revision data found for selected ISO.";
                lblerror.ForeColor = System.Drawing.Color.Red;
            }
        }

        private void BindReExportGrid(List<SPMATReExportData> data)
        {
            var dt = ConvertReExportDataToDataTable(data);
            grisoreview.Columns.Clear();

            foreach (DataColumn col in dt.Columns)
            {
                TemplateField templateField = null;
                switch (col.ColumnName)
                {
                    case "Unit":
                        templateField = new TemplateField
                        {
                            HeaderText = "Unit",
                            ItemTemplate = new UnitDropDownTemplate("ddlUnit")
                        };
                        break;
                    case "Phase":
                        templateField = new TemplateField
                        {
                            HeaderText = "Phase",
                            ItemTemplate = new PhaseDropDownTemplate("ddlPhase")
                        };
                        break;
                    case "Const_Area":
                        templateField = new TemplateField
                        {
                            HeaderText = "Const Area",
                            ItemTemplate = new ConstAreaDropDownTemplate("ddlConstArea")
                        };
                        break;
                    case "qty":
                        templateField = new TemplateField
                        {
                            HeaderText = "Quantity",
                            ItemTemplate = new QtyTextBoxTemplate("txtQty")
                        };
                        break;
                    case "IsoRevision":
                        templateField = new TemplateField
                        {
                            HeaderText = "IsoRevision",
                            ItemTemplate = new RevisionsDropDownTemplate("ddlRevisions")
                        };
                        break;
                    default:
                        BoundField boundField = new BoundField
                        {
                            DataField = col.ColumnName,
                            HeaderText = col.ColumnName
                        };
                        boundField.ItemStyle.Wrap = false;
                        boundField.HeaderStyle.Wrap = false;
                        grisoreview.Columns.Add(boundField);
                        break;
                }

                if (templateField != null)
                {
                    templateField.ItemStyle.Wrap = false;
                    templateField.HeaderStyle.Wrap = false;
                    grisoreview.Columns.Add(templateField);
                }
            }

            grisoreview.DataSource = dt;
            grisoreview.DataBind();
        }

        private DataTable ConvertReExportDataToDataTable(List<SPMATReExportData> data)
        {
            DataTable table = new DataTable();
            var properties = typeof(SPMATReExportData).GetProperties();
            foreach (var prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }
            foreach (var item in data)
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

        protected void grisoreview_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                var dataItem = (DataRowView)e.Row.DataItem;

                DropDownList ddlUnit = (DropDownList)e.Row.FindControl("ddlUnit");
                DropDownList ddlPhase = (DropDownList)e.Row.FindControl("ddlPhase");
                DropDownList ddlConstArea = (DropDownList)e.Row.FindControl("ddlConstArea");
                DropDownList ddlSpec = (DropDownList)e.Row.FindControl("ddlSpec");
                DropDownList ddlShortcode = (DropDownList)e.Row.FindControl("ddlShortcode");
                DropDownList ddlIdent = (DropDownList)e.Row.FindControl("ddlIdent");
                DropDownList ddlRevisions = (DropDownList)e.Row.FindControl("ddlRevisions");
                TextBox txtQty = (TextBox)e.Row.FindControl("txtQty");

                var allUnits = Session["Units"] as List<string>;
                var allSpecs = Session["Specs"] as List<string>;
                var unitPhaseMap = Session["UnitPhaseMap"] as Dictionary<string, List<string>>;
                var unitConstAreaMap = Session["UnitConstAreaMap"] as Dictionary<string, Dictionary<string, List<string>>>;
                var specShortCodeMap = Session["SpecShortCodeMap"] as Dictionary<string, List<string>>;
                var shortCodeIdentMap = Session["ShortCodeIdentMap"] as Dictionary<string, Dictionary<string, List<string>>>;

                if (ddlUnit != null && allUnits != null)
                {
                    ddlUnit.DataSource = allUnits;
                    ddlUnit.DataBind();
                    ddlUnit.Items.Insert(0, new ListItem("-- Select --", ""));
                    ddlUnit.SelectedValue = dataItem["Unit"].ToString();
                }

                // Similar cascading logic for Phase, ConstArea, Spec, Shortcode, Ident
                if (txtQty != null) txtQty.Text = dataItem["qty"].ToString();
            }
        }

        protected void btnUpdate_ISO_Click(object sender, EventArgs e)
        {
            var isosheet = Session["REISOSHEET"].ToString();
            var updatedItems = new List<SPMATReExportData>();
            foreach (GridViewRow row in grisoreview.Rows)
            {
                int mtoId = Convert.ToInt32(grisoreview.DataKeys[row.RowIndex].Value);
                var txtQty = row.FindControl("txtQty") as TextBox;
                var ddlUnit = row.FindControl("ddlUnit") as DropDownList;
                var ddlPhase = row.FindControl("ddlPhase") as DropDownList;
                var ddlConstArea = row.FindControl("ddlConstArea") as DropDownList;
                var ddlRevisions = row.FindControl("ddlRevisions") as DropDownList;

                decimal qtyValue = 0;
                decimal.TryParse(txtQty?.Text, out qtyValue);

                var updated = new SPMATReExportData
                {
                    MTOID = mtoId,
                    qty = qtyValue,
                    Unit = ddlUnit?.SelectedValue,
                    Phase = ddlPhase?.SelectedValue,
                    Const_Area = ddlConstArea?.SelectedValue,
                    IsoRevision=ddlRevisions?.SelectedValue,
                };
                updatedItems.Add(updated);
            }
            var projid = Session["ENGPROJECTID"].ToString();
            List<int> MTOIDs = updatedItems.Select(d => d.MTOID).Distinct().ToList();
            int FileID = DataClass.InsertExportRecord(projid, MTOIDs);
            Session["ReExportFileID"] = FileID;
            foreach (var item in updatedItems)
            {
               DataClass.UpdateReExportData(item,FileID,isosheet);
            }
            grisoreview.DataSource = null;
            grisoreview.DataBind();
            lblFileName.Text = "Re-Export data updated successfully.";
            lblFileName.ForeColor = System.Drawing.Color.Green;
            btnReExportMTO.Visible = true;
            btnUpdate_ISO.Visible = false;
            ddliso.ClearSelection();
            Session["REISOSHEET"] = null;
            List<DDLList> iso = DataClass.GetProjectISORevData(projid);
            if (iso.Count > 0)
            {
                ddliso.DataSource = iso;
                ddliso.DataTextField = "DDLListName";
                ddliso.DataValueField = "DDLList_ID";
                ddliso.DataBind();
            }
        }

        protected void btnReExportMTO_Click(object sender, EventArgs e)
        {
            //var data = Session["IsoReviewData"] as List<SPMATReExportData>;
            var FileID= int.Parse(Session["ReExportFileID"].ToString());
            var projid = Session["ENGPROJECTID"].ToString();
            var projdesc = Session["ENGPROJDESC"].ToString();
            List<SPMATData> mtoData = DataClass.GetMTOData(projid, true);
            //DataClass.UpdateExportRecord(FileID, FileName);
            if (mtoData == null || !mtoData.Any()) return;

            ClosedXML.Excel.XLWorkbook workbook = new ClosedXML.Excel.XLWorkbook();

            IXLWorksheet worksheet = workbook.Worksheets.Add("SPMAT_ReExport");

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
            foreach (var item in mtoData)
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
            var FileName = Server.UrlEncode(projdesc + "_REEXPORT_MTO_ALL_SPMAT_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx");
            DataClass.UpdateExportRecord(FileID, FileName);
            DownloadFile efile = new DownloadFile();
            efile.filename = FileName;
            efile.contenttype = "application/vnd.ms-excel";
            using (System.IO.MemoryStream stream = GetStreamXL(workbook))
            {
                try
                {
                    stream.Position = 0;
                    using (BinaryReader br = new BinaryReader(stream))
                    {
                        byte[] bytes = br.ReadBytes((Int32)stream.Length);
                        efile.filedata = bytes;
                        DataClass.SaveExportRecordFile(FileID, efile.filedata);
                        Response.Clear();
                        Response.Buffer = true;
                        Response.AddHeader("content-disposition", "attachment; filename=" + efile.filename);
                        Response.ContentType = efile.contenttype;
                        Response.BinaryWrite(efile.filedata.ToArray());
                        Response.Flush();
                        Response.End();
                    }
                }
                catch { }
                DataClass.SaveExportRecordFile(FileID, efile.filedata);
               
            }
            //DataTable dt = ConvertReExportDataToDataTable(data);
            //using (XLWorkbook wb = new XLWorkbook())
            //{
            //    wb.Worksheets.Add(dt, "ReExportMTO");
            //    using (MemoryStream stream = new MemoryStream())
            //    {
            //        wb.SaveAs(stream);
            //        stream.Position = 0;
            //        Response.Clear();
            //        Response.Buffer = true;
            //        Response.AddHeader("content-disposition", "attachment; filename=ReExportMTO_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            //        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //        Response.BinaryWrite(stream.ToArray());
            //        Response.Flush();
            //        Response.End();
            //    }
            //}
        }

        public MemoryStream GetStreamXL(IXLWorkbook excelWorkbook)
        {
            MemoryStream fs = new MemoryStream();
            excelWorkbook.SaveAs(fs);
            fs.Position = 0;
            return fs;
        }
    }
}