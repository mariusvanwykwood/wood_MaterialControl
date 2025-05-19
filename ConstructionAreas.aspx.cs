using System;
using System.Security.Principal;
using System.Web;
using System.Collections.Generic;
using System.Linq;

namespace Wood_MaterialControl
{
    public partial class ConstructionAreas : System.Web.UI.Page
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
    }
}