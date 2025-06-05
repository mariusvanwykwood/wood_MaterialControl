using System;
using System.Security.Principal;
using System.Web;

namespace Wood_MaterialControl
{
    public partial class Default : System.Web.UI.Page
    {
        
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                SetNoCache();
                if (!IsPostBack)
                {
                    var userfailure = 0;
                    try
                    {
                        userfailure = int.Parse(Request.QueryString["UF"]);

                    }
                    catch { }
                    if (userfailure > 0)
                    {
                        diverror.Style["display"] = "block";
                        if (userfailure == 1)
                        {
                            lblerror.Text = "You need to select your user before you can continue";
                        }
                        else if (userfailure == 2)
                        {
                            lblerror.Text = "You Are not Authorized to use the selected Functionality";
                        }
                        else
                        {
                            lblerror.Text = "You do not have access";
                        }
                    }
                    ddlemployees.DataSource = DataClass.GetAllUsersLookup();
                    ddlemployees.DataTextField = "UserName";
                    ddlemployees.DataValueField = "UserID";
                    ddlemployees.DataBind();
                    ddlemployees.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- Select --", ""));
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


        protected void btnLogin_Click(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            var uid = int.Parse(ddlemployees.SelectedValue);
            if (uid != 0)
            {
                if (DataClass.UserHasAccess(uid))
                {

                    Session["EID"]= uid;
                    Response.Redirect("MainHome.aspx");
                }
                else
                {
                    diverror.Style["display"] = "block";
                    lblerror.Text = "Sorry , you do not currently have access to this sytem";
                }
                    
            }

        }

        protected void ddlemployees_SelectedIndexChanged(object sender, EventArgs e)
        {
            diverror.Style["display"] = "none";
            lblerror.Text = "";
            var uid = int.Parse(ddlemployees.SelectedValue);
            if (uid != 0)
            {
                if (DataClass.UserHasAccess(uid))
                {

                    Session.Add("EID", uid);
                    btnLogin.CssClass = "shown";
                }
                else
                {
                    diverror.Style["display"] = "block";
                    lblerror.Text = "You do not have access";
                    btnLogin.CssClass = "hidden";
                }

            }
        }
    }
}