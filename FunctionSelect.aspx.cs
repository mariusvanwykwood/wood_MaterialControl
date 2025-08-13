using System;
using System.Security.Principal;
using System.Web;

namespace Wood_MaterialControl
{
    public partial class FunctionSelect : System.Web.UI.Page
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
                        if(EID==75)
                        {
                            btnMatcon.Visible = false;
                        }
                    }
                    else
                    {
                        Session.Clear();
                        Response.Redirect("Default.aspx?UF=1");
                    }
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

        

        protected void btnMatcon_Click(object sender, EventArgs e)
        {
            Response.Redirect("MainHome.aspx");
        }

        protected void btnIsoReview_Click(object sender, EventArgs e)
        {
            Response.Redirect("ReviewIso.aspx");
        }
    }
}