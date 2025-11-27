using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Wood_MaterialControl
{

    public class CheckBoxTemplate : ITemplate
    {
        private string _id;

        public CheckBoxTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            CheckBox chk = new CheckBox();
            chk.ID = _id;
            chk.AutoPostBack = true;
            chk.EnableViewState = true;
            chk.Attributes["onchange"] = $"__doPostBack('chkChecked', '')";
            container.Controls.Add(chk);
        }
    }

    public class QtyTextBoxTemplate : ITemplate
    {
        private string _id;

        public QtyTextBoxTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            TextBox txt = new TextBox();
            txt.ID = _id;
            txt.AutoPostBack = true;
            txt.EnableViewState= true;
            txt.Attributes.Add("style", "width:80px;");
            txt.Attributes["onchange"] = $"__doPostBack('txtQty', '')";
            // Bind the value from the DataItem
            txt.DataBinding += (sender, e) =>
            {
                TextBox t = (TextBox)sender;
                GridViewRow row = (GridViewRow)t.NamingContainer;

                if (row != null && row.DataItem != null)
                {

                    object dataValue = DataBinder.Eval(row.DataItem, "qty");
                    if (dataValue != null)
                    {
                        t.Text = dataValue.ToString();
                    }
                }
            };

            container.Controls.Add(txt);
        }
    }

    public class UnitDropDownTemplate : ITemplate
    {
        private string _id;

        public UnitDropDownTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            DropDownList ddl = new DropDownList();
            ddl.ID = _id;
            ddl.AutoPostBack = true;
            ddl.EnableViewState = true;
            ddl.Attributes["onchange"] = $"__doPostBack('ddlUnit', '')";
            var units = HttpContext.Current.Items["Units"] as List<string> ?? new List<string>();
            ddl.DataBinding += (sender, e) =>
            {
                DropDownList d = (DropDownList)sender;
                GridViewRow row = (GridViewRow)d.NamingContainer;

                if (row != null && row.DataItem != null)
                {

                    string currentValue = DataBinder.Eval(row.DataItem, "Unit")?.ToString();
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        ListItem item = d.Items.FindByValue(currentValue);
                        if (item != null)
                            d.SelectedValue = currentValue;
                    }
                }
            };
            ddl.DataSource = units;
            ddl.DataBind();
            container.Controls.Add(ddl);
        }

    }

    public class PhaseDropDownTemplate : ITemplate
    {
        private string _id;

        public PhaseDropDownTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            DropDownList ddl = new DropDownList();
            ddl.ID = _id;
            ddl.AutoPostBack = true;
            var phases = (List<string>)HttpContext.Current.Items["Phases"];
            ddl.Attributes["onchange"] = $"__doPostBack('ddlPhase', '')";
          
            ddl.DataBinding += (sender, e) =>
            {
                DropDownList d = (DropDownList)sender;
                GridViewRow row = (GridViewRow)d.NamingContainer;


                if (row != null && row.DataItem != null)
                {

                    string currentValue = DataBinder.Eval(row.DataItem, "Phase")?.ToString();
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        ListItem item = d.Items.FindByValue(currentValue);
                        if (item != null)
                            d.SelectedValue = currentValue;
                    }
                }
            };
            ddl.DataSource = phases;
            ddl.DataBind();
            container.Controls.Add(ddl);
        }
    }
    public class ConstAreaDropDownTemplate : ITemplate
    {
        private string _id;

        public ConstAreaDropDownTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            DropDownList ddl = new DropDownList();
            ddl.ID = _id;
            ddl.AutoPostBack = true;
            ddl.Attributes["onchange"] = $"__doPostBack('ddlConstArea', '')";
            var constAreas = (List<string>)HttpContext.Current.Items["ConstAreas"];
           
            ddl.DataBinding += (sender, e) =>
            {
                DropDownList d = (DropDownList)sender;
                GridViewRow row = (GridViewRow)d.NamingContainer;

                if (row != null && row.DataItem != null)
                {



                    string currentValue = DataBinder.Eval(row.DataItem, "Const_Area")?.ToString();
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        ListItem item = d.Items.FindByValue(currentValue);
                        if (item != null)
                            d.SelectedValue = currentValue;
                    }
                }
            };
            ddl.DataSource = constAreas;
            ddl.DataBind();
            container.Controls.Add(ddl);
        }
    }
    public class RevisionsDropDownTemplate : ITemplate
    {
        private string _id;

        public RevisionsDropDownTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            DropDownList ddl = new DropDownList();
            ddl.ID = _id;
            ddl.AutoPostBack = true;
            ddl.Attributes["onchange"] = $"__doPostBack('ddlConstArea', '')";
            var revisions = (List<string>)HttpContext.Current.Items["Revisions"];

            ddl.DataBinding += (sender, e) =>
            {
                DropDownList d = (DropDownList)sender;
                GridViewRow row = (GridViewRow)d.NamingContainer;

                if (row != null && row.DataItem != null)
                {



                    string currentValue = DataBinder.Eval(row.DataItem, "IsoRevision")?.ToString();
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        ListItem item = d.Items.FindByValue(currentValue);
                        if (item != null)
                            d.SelectedValue = currentValue;
                    }
                }
            };
            ddl.DataSource = revisions;
            ddl.DataBind();
            container.Controls.Add(ddl);
        }
    }

    public class SpecDropDownTemplate : ITemplate
    {
        private string _id;

        public SpecDropDownTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            DropDownList ddl = new DropDownList();
            ddl.ID = _id;
            ddl.AutoPostBack = true;
            ddl.EnableViewState = true;
            ddl.Attributes["onchange"] = $"__doPostBack('ddlSpec', '')";
            var specs = HttpContext.Current.Items["Specs"] as List<string> ?? new List<string>();


            ddl.DataBinding += (sender, e) =>
            {
                DropDownList d = (DropDownList)sender;
                GridViewRow row = (GridViewRow)d.NamingContainer;

                if (row != null && row.DataItem != null)
                {

                    string currentValue = DataBinder.Eval(row.DataItem, "Spec")?.ToString();
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        ListItem item = d.Items.FindByValue(currentValue);
                        if (item != null)
                        {
                            d.SelectedValue = currentValue;
                        }
                    }
                }
            };
            ddl.DataSource = specs;
            ddl.DataBind();
            container.Controls.Add(ddl);
        }

    }

    public class ShortCodeDropDownTemplate : ITemplate
    {
        private string _id;

        public ShortCodeDropDownTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            DropDownList ddl = new DropDownList();
            ddl.ID = _id;
            ddl.AutoPostBack = true;
            ddl.EnableViewState = true;
            ddl.Attributes["onchange"] = $"__doPostBack('ddlShortcode', '')";
            var shortcodes = (List<string>)HttpContext.Current.Items["ShortCodes"];
           
           
            ddl.DataBinding += (sender, e) =>
            {
                DropDownList d = (DropDownList)sender;
                GridViewRow row = (GridViewRow)d.NamingContainer;

                if (row != null && row.DataItem != null)
                {


                    string currentValue = DataBinder.Eval(row.DataItem, "Shortcode")?.ToString();
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        ListItem item = d.Items.FindByValue(currentValue);
                        if (item != null)
                        {
                            d.SelectedValue = currentValue;
                        }
                    }
                }
            };
            ddl.DataSource = shortcodes;
            ddl.DataBind();
            container.Controls.Add(ddl);
        }

    }
    public class IdentDropDownTemplate : ITemplate
    {
        private string _id;

        public IdentDropDownTemplate(string id)
        {
            _id = id;
        }

        public void InstantiateIn(Control container)
        {
            DropDownList ddl = new DropDownList();
            ddl.ID = _id;
            ddl.AutoPostBack = true;
            var idents = (List<string>)HttpContext.Current.Items["Idents"];
            ddl.Attributes["onchange"] = $"__doPostBack('ddlIdent', '')";
          
            ddl.DataBinding += (sender, e) =>
            {
                DropDownList d = (DropDownList)sender;
                GridViewRow row = (GridViewRow)d.NamingContainer;


                if (row != null && row.DataItem != null)
                {


                    string currentValue = DataBinder.Eval(row.DataItem, "Ident_no")?.ToString();
                    if (!string.IsNullOrEmpty(currentValue))
                    {
                        ListItem item = d.Items.FindByValue(currentValue);
                        if (item != null)
                            d.SelectedValue = currentValue;
                    }
                }
            };
            ddl.DataSource = idents;
            ddl.DataBind();
            container.Controls.Add(ddl);
        }
    }

    public class LabelTemplate : ITemplate
    {
        private readonly string _id;
        private readonly string _dataField;

        public LabelTemplate(string id, string dataField)
        {
            _id = id;
            _dataField = dataField;
        }

        public void InstantiateIn(Control container)
        {
            Label lbl = new Label { ID = _id };
            lbl.DataBinding += (sender, e) =>
            {
                Label l = (Label)sender;
                GridViewRow row = (GridViewRow)l.NamingContainer;
                object dataValue = DataBinder.Eval(row.DataItem, _dataField);
                l.Text = dataValue?.ToString() ?? string.Empty;
            };
            container.Controls.Add(lbl);
        }
    }
    public class LabelISoTemplate : ITemplate
    {
        private readonly string _id;
        private readonly string _dataField;

        public LabelISoTemplate(string id, string dataField)
        {
            _id = id;
            _dataField = dataField;
        }

        public void InstantiateIn(Control container)
        {
            Label lbl = new Label { ID = _id };
            lbl.Width = Unit.Pixel(550);
            lbl.DataBinding += (sender, e) =>
            {
                Label l = (Label)sender;
                GridViewRow row = (GridViewRow)l.NamingContainer;
                object dataValue = DataBinder.Eval(row.DataItem, _dataField);
                l.Text = dataValue?.ToString() ?? string.Empty;
            };
            container.Controls.Add(lbl);
        }
    }

}