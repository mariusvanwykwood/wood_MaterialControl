<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MainHome.aspx.cs" MaintainScrollPositionOnPostBack="true" Inherits="Wood_MaterialControl.MainHome" EnableEventValidation="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        #pnlExcelUpload {
            padding: 10px;
            border: 1px solid #ccc;
            background-color: #f9f9f9;
        }

        .excel-grid th, .excel-grid td {
            border: 1px solid #ccc;
            padding: 5px;
        }

        .excel-grid th {
            background-color: #d9edf7;
        }

        .excel-grid tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .hidden{
            display:none;
        }
        .shown{
    display:block;
}
        .wrap-label {
    display: block;
    min-width:250px;
    max-width: 500px;
    word-wrap: break-word;
    white-space: normal;
}


    </style>



    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link rel="stylesheet" href="Content/bootstrap.css"/>
<link rel="stylesheet" href="Content/Site.css"/>
<link rel="stylesheet" href="https://use.typekit.net/pjp8xxm.css"/>
<link rel="stylesheet" href="css/daily-diary.css"/>
<link href="wood-theme-css/App.css" rel="stylesheet"/>
<script src="assets/vendor/popper/popper.min.js"></script>
<script src="assets/vendor/bootstrap/js/bootstrap.min.js"></script>
  <meta name="viewport" content = "width =device-width, initial-scale = 1.0, maximum-scale = 1.0, user-scalable = no"/>
</head>
<body>
    <header>
    <nav>
        <div class="navigation-header">
            <a href="MainHome.aspx" title="Material Control Main Home"  >
            <img class="large-logo" src="images/dd-logo.png" style="padding-left: 12px; margin-top: 16px;" />
            <img class="small-logo" src="images/dd-small-logo.png" style="padding-left: 24px; margin-top: 16px; width: 48px;"/></a>
        </div>
    </nav>
</header>
    <form id="form1" runat="server">
        <div id="divstartup" runat="server" style="padding-left: 20px;">
            <h2 runat="server" id="mainh1" style="text-align:left;">Material Control</h2> 
            <br />
         </div>   
<div style="display: flex; align-items: flex-start; gap: 20px;">
    
    <div >
        <asp:Label ID="lblClient" runat="server" Text="Choose Client"></asp:Label>
   <asp:DropDownList ID="ddlclient" runat="server" OnSelectedIndexChanged="ddlclient_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
        <asp:Label ID="lblproj" runat="server" Text="Choose Spec"></asp:Label>
           <asp:DropDownList ID="ddlprojects" runat="server" OnSelectedIndexChanged="ddlprojects_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
         <asp:Label ID="lblprojsel" runat="server" Text="Choose Project"></asp:Label>
    <asp:DropDownList ID="ddlprojsel" runat="server" OnSelectedIndexChanged="ddlprojsel_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
        <br />
        <br />
        <asp:Label ID="lbliso" runat="server" Text="Select Iso"></asp:Label>
        <asp:DropDownListChosen  ID="ddliso" runat="server" NoResultsText="No results match."  DataPlaceHolder="Search..." AllowSingleDeselect="true" Width="250px" ></asp:DropDownListChosen>&nbsp;<asp:Button ID="btnloadiso" runat="server" OnClick="btnloadiso_Click" Text="Load Iso" />
        <br />
        <br />
      
        <asp:Panel ID="pnlExcelUpload" runat="server" Visible="false" style="flex-grow: 1;padding: 10px; border: 1px solid #ccc; background-color: #f9f9f9;">
    <br />
    <asp:Label ID="lblFileName" runat="server" Font-Bold="true" Font-Size="Large" />
    <br />
    <asp:GridView ID="ExcelGridView" EnableViewState="true" DataKeyNames="MaterialID" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered" OnRowDataBound="ExcelGridView_RowDataBound">
    </asp:GridView>
            <br />
<asp:Button ID="btnAddNew" runat="server" Text="Add New Entry" OnClick="btnAddNew_Click" BackColor="Wheat"/>
<asp:Panel ID="pnlAddForm" runat="server" Visible="false" >
    <table class="table table-bordered">
        <tr>
            <td>Unit:</td>
            <td>
                <asp:DropDownList ID="ddlAddUnit" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAddUnit_SelectedIndexChanged" />
            </td>
        </tr>
        <tr>
            <td>Phase:</td>
            <td>
                <asp:DropDownList ID="ddlAddPhase" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAddPhase_SelectedIndexChanged" />
            </td>
        </tr>
        <tr>
            <td>Const Area:</td>
            <td>
                <asp:DropDownList ID="ddlAddConstArea" runat="server" />
            </td>
        </tr>
        <tr>
            <td>Spec:</td>
            <td>
                <asp:DropDownList ID="ddlAddSpec" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAddSpec_SelectedIndexChanged" />
            </td>
        </tr>
        <tr>
            <td>Shortcode:</td>
            <td>
                <asp:DropDownList ID="ddlAddShortcode" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAddShortcode_SelectedIndexChanged" />
            </td>
        </tr>
        <tr>
            <td>Ident:</td>
            <td>
                <asp:DropDownList ID="ddlAddIdent" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAddIdent_SelectedIndexChanged"/>&nbsp;<asp:Label ID="lblidentdesc" runat="server" CssClass="wrap-label"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>Quantity/Length:</td>
            <td>
                <asp:TextBox ID="txtAddQty" runat="server" CssClass="form-control" />
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <asp:Button ID="btnSaveNew" runat="server" Text="Save Entry" OnClick="btnSaveNew_Click" />
                <asp:Button ID="btnCancelAdd" runat="server" Text="Cancel/Hide" OnClick="btnCancelAdd_Click"  />
            </td>
        </tr>
    </table>
</asp:Panel>

<br />
            <asp:Label ID="lblMessage" runat="server" ForeColor="Red" />
            <br />
<asp:Button ID="btnSubmit" runat="server" Text="Save IsoData"
    OnClick="btnSubmit_Click" />
</asp:Panel>
        <br />
        <asp:Button ID="btnviewspmat" runat="server" Text="View Working Data" OnClick="btnviewspmat_Click" CssClass="hidden" />  
        
 <br />
        
<asp:Panel ID="pnlSPMATView" runat="server" Visible="false">
<asp:GridView ID="gvSPMAT" runat="server" AutoGenerateColumns="False"
    CssClass="table table-bordered"
    DataKeyNames="Discipline,Area,Unit,Phase,Const_Area,ISO,Ident_no,Spec">
    
    <Columns>
        <asp:BoundField DataField="MaterialID" HeaderText="MaterialID"  Visible="false"/>
        <asp:BoundField DataField="ProjectID" HeaderText="ProjectID"  Visible="false"/>
        <asp:BoundField DataField="Discipline" HeaderText="Discipline" />
        <asp:BoundField DataField="Area" HeaderText="Area" />
        <asp:BoundField DataField="Unit" HeaderText="Unit" />
        <asp:BoundField DataField="Phase" HeaderText="Phase" />
        <asp:BoundField DataField="Const_Area" HeaderText="Const Area" />
        <asp:BoundField DataField="ISO" HeaderText="ISO" />
        <asp:BoundField DataField="Ident_no" HeaderText="Ident No" />
        <asp:BoundField DataField="qty" HeaderText="Quantity" />
        <asp:BoundField DataField="qty_unit" HeaderText="Unit" />
        <asp:BoundField DataField="Spec" HeaderText="Spec" />
        <asp:BoundField DataField="Fabrication_Type" HeaderText="Fabrication Type" />
        <asp:BoundField DataField="IsoRevisionDate" HeaderText="Revision Date" />
        <asp:BoundField DataField="IsoRevision" HeaderText="Revision" />
        <asp:BoundField DataField="Lock" HeaderText="Locked" />
        <asp:BoundField DataField="Code" HeaderText="Code" />
       <%-- <asp:TemplateField HeaderText="Actions">
            <ItemTemplate>
                <asp:Button ID="btnRemove" runat="server" Text="Remove"
                    CommandName="RemoveRow"
                    CommandArgument='<%# Container.DataItemIndex %>'/>
            </ItemTemplate>
        </asp:TemplateField>--%>
    </Columns>
</asp:GridView>

    <br />
    <asp:Button ID="btnExportSPMAT" runat="server" Text="Confirm For MTO" OnClick="btnExportSPMAT_Click"/>
</asp:Panel>
        <br />
 <asp:Button ID="btnviewInterim" runat="server" Text="View PreFinal MTo Data" OnClick="btnviewInterim_Click" CssClass="hidden" />  
        <br />
        <asp:Panel ID="pnlViewInterimData" runat="server" Visible="false">
     <asp:Label ID="lblintrim" runat="server" Font-Bold="true" Font-Size="Large" Text="Removing an Item will Remove All for that ISO" ForeColor="Red" />
<asp:GridView ID="gvInterim" runat="server" AutoGenerateColumns="False"
    CssClass="table table-bordered" OnRowCommand="gvInterim_RowCommand"
    DataKeyNames="INTID,MaterialID,ISO">
    
    <Columns>
        <asp:BoundField DataField="INTID" HeaderText="INTID"  Visible="false"/>
        <asp:BoundField DataField="MaterialID" HeaderText="MaterialID" Visible="false"/>
        <asp:BoundField DataField="Discipline" HeaderText="Discipline" />
        <asp:BoundField DataField="Area" HeaderText="Area" />
        <asp:BoundField DataField="Unit" HeaderText="Unit" />
        <asp:BoundField DataField="Phase" HeaderText="Phase" />
        <asp:BoundField DataField="Const_Area" HeaderText="Const Area" />
        <asp:BoundField DataField="ISO" HeaderText="ISO" />
        <asp:BoundField DataField="Ident_no" HeaderText="Ident No" />
        <asp:BoundField DataField="qty" HeaderText="Quantity" />
        <asp:BoundField DataField="qty_unit" HeaderText="Unit" />
        <asp:BoundField DataField="Spec" HeaderText="Spec" />
        <asp:BoundField DataField="Fabrication_Type" HeaderText="Fabrication Type" />
        <asp:BoundField DataField="IsoRevisionDate" HeaderText="Revision Date" />
        <asp:BoundField DataField="IsoRevision" HeaderText="Revision" />
        <asp:BoundField DataField="IsLocked" HeaderText="Locked" />
        <asp:TemplateField HeaderText="Actions">
            <ItemTemplate>
                <asp:Button ID="btnRemove" runat="server" Text="Remove"
                    CommandName="RemoveRow"
                    CommandArgument='<%# Container.DataItemIndex %>'/>
            </ItemTemplate>
        </asp:TemplateField>
    </Columns>
</asp:GridView>

    <br />
    <asp:Button ID="btnMoveToFinal" runat="server" Text="Move To Final" OnClick="btnMoveToFinal_Click"/>
</asp:Panel>
               <br />
<asp:Button ID="btnViewFinal" runat="server" Text="View Pre Export MTo Data" OnClick="btnViewFinal_Click" CssClass="hidden" />  
       <br />
        <asp:Panel ID="pnlFinalMTO" runat="server" Visible="false">
<asp:GridView ID="gvFinal" runat="server" AutoGenerateColumns="False"
    CssClass="table table-bordered" OnRowCommand="gvInterim_RowCommand"
    DataKeyNames="MTOID">
    
    <Columns>
        <asp:BoundField DataField="MTOID" HeaderText="INTID"  Visible="false"/>
        <asp:BoundField DataField="Discipline" HeaderText="Discipline" />
        <asp:BoundField DataField="Area" HeaderText="Area" />
        <asp:BoundField DataField="Unit" HeaderText="Unit" />
        <asp:BoundField DataField="Phase" HeaderText="Phase" />
        <asp:BoundField DataField="Const_Area" HeaderText="Const Area" />
        <asp:BoundField DataField="ISO" HeaderText="ISO" />
        <asp:BoundField DataField="Ident_no" HeaderText="Ident No" />
        <asp:BoundField DataField="qty" HeaderText="Quantity" />
        <asp:BoundField DataField="qty_unit" HeaderText="Unit" />
        <asp:BoundField DataField="Spec" HeaderText="Spec" />
        <asp:BoundField DataField="Fabrication_Type" HeaderText="Fabrication Type" />
        <asp:BoundField DataField="IsoRevisionDate" HeaderText="Revision Date" />
        <asp:BoundField DataField="IsoRevision" HeaderText="Revision" />
        <asp:BoundField DataField="IsLocked" HeaderText="Locked" />
        <asp:BoundField DataField="Code" HeaderText="Code" />
        <asp:BoundField DataField="ImportStatus" HeaderText="Imported Status"  ItemStyle-Wrap="false"/>
    </Columns>
</asp:GridView>

    <br />
    <asp:Button ID="btnExportFinalMTO" runat="server" Text="Export MTO File" OnClick="btnExportFinalMTO_Click"/>
</asp:Panel>
        <br />
        <asp:Panel ID="pnlFileMan" runat="server">
        Need File List here once Exported, from FileExports Table.
        Need a Button for Imported and Text Field for Imported Status
        </asp:Panel>
        </div>
           </div>
        <div id="diverror" runat="server" style="display:none;padding-left: 20px;">
            <asp:Label ID="lblerror" runat="server" CssClass="errorlabel"></asp:Label>
        </div>
    </form>
</body>
</html>
