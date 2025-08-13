using System;
using System.Security.Principal;
using System.Web;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI.WebControls;
using static Wood_MaterialControl.DataClass;
using ClosedXML.Excel;
using System.Data;


namespace Wood_MaterialControl
{
    public partial class ReviewIso : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {

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
                        if(EID!=240)
                        {
                            Response.Redirect("Default.aspx");
                        }
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

                }
            }
            catch
            {
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

        protected void btnloadiso_Click(object sender, EventArgs e)
        {
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
            pnlisodata.Style["display"] = "block";
            var isosheet = ddliso.SelectedValue;
            var projid = ddlprojsel.SelectedValue;
            List<IsoRevisionData> sptest = DataClass.GetIsoReviewMTOData(isosheet, projid);
            Session["IsoReviewData"] = sptest;

            if (sptest != null && sptest.Count > 0)
            {
                grisoreview.DataSource = sptest;
                grisoreview.DataBind();
                pnlisodata.Visible = true;
                lblFileName.Text = $"Review for ISO: {isosheet}";
            }
            else
            {
                diverror.Style["display"] = "block";
                lblerror.Text = "No revision data found for selected ISO.";
                lblerror.ForeColor = System.Drawing.Color.Red;
            }

        }

  

        protected void chkReleased_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox chk = (CheckBox)sender;
            GridViewRow row = (GridViewRow)chk.NamingContainer;

            int mtoid = Convert.ToInt32(grisoreview.DataKeys[row.RowIndex].Value);

            bool isChecked = chk.Checked;

            // Call your data update method
            DataClass.UpdateReleasedMaterialStatus(mtoid, isChecked);

            // Get the full object from session
            var reviewData = Session["IsoReviewData"] as List<IsoRevisionData>;
            var item = reviewData?.FirstOrDefault(x => x.MTOID == mtoid);

            string iso = item?.ISO ?? "[Unknown]";
            string ident = item?.Ident_no ?? "[Unknown]";

            // Refresh the grid
            string projid = ddlprojsel.SelectedValue;
            string isosheet = ddliso.SelectedValue;

            reviewData = DataClass.GetIsoReviewMTOData( isosheet,projid);
            Session["IsoReviewData"] = reviewData;
            grisoreview.DataSource = reviewData;
            grisoreview.DataBind();
       
            // Optional: show confirmation
            lblerror.Text = $"Iso: {iso} ⇒ Ident: {ident} marked as {(isChecked ? "Released" : "Unreleased")}.";
            lblerror.ForeColor = System.Drawing.Color.Green;
            lblerror.ForeColor = isChecked ? System.Drawing.Color.Green : System.Drawing.Color.Orange;
            diverror.Style["display"] = "block";

        }

    }
}