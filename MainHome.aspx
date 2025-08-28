
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
        .errorlabel {
            color: red;
            font-weight: bold;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link rel="stylesheet" href="Content/bootstrap.css"/>
    <link rel="stylesheet" href="Content/Site.css"/>
    <link rel="stylesheet" href="https://use.typekit.net/pjp8xxm.css"/>
    <link rel="stylesheet" href="css/daily-diary.css"/>
    <link href="wood-theme-css/App.css" rel="stylesheet"/>
    <script src="assets/vendor/popper/popper.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no"/>
</head>
<body>
    <header>
        <nav>
            <div class="navigation-header">
                <a href="Default.aspx" title="Material Control Main Home">
                    <img class="large-logo" src="images/dd-logo.png" style="padding-left: 12px; margin-top: 16px;" />
                    <img class="small-logo" src="images/dd-small-logo.png" style="padding-left: 24px; margin-top: 16px; width: 48px;"/>
                </a>
            </div>
        </nav>
    </header>
    <form id="form1" runat="server">
        
<asp:ScriptManager ID="ScriptManager1" runat="server" />

        <div class="accordion" id="panelAccordion">
            <div class="accordion-item">
                <h2 class="accordion-header" id="headingOne">
                    <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseOne">
                       Main Selections/ Step 1
                    </button>
                </h2>
                <div id="collapseOne" class="accordion-collapse collapse show" >
                    <div class="accordion-body">
                        <div id="divstartup" runat="server" style="padding-left: 20px;">
                            <h2 runat="server" id="mainh1" style="text-align:left;">Material Control</h2>
                            <br />
                        </div>
                        <div style="display: flex; align-items: flex-start; gap: 20px;">
                            <div>
                                <asp:Label ID="lblClient" runat="server" Text="Choose Client" AssociatedControlID="ddlclient"></asp:Label>
                                <asp:DropDownList ID="ddlclient" runat="server" OnSelectedIndexChanged="ddlclient_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
                                <asp:Label ID="lblproj" runat="server" Text="Choose Spec" AssociatedControlID="ddlprojects"></asp:Label>
                                <asp:DropDownList ID="ddlprojects" runat="server" OnSelectedIndexChanged="ddlprojects_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
                                <asp:Label ID="lblprojsel" runat="server" Text="Choose Project" AssociatedControlID="ddlprojsel"></asp:Label>
                                <asp:DropDownList ID="ddlprojsel" runat="server" OnSelectedIndexChanged="ddlprojsel_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList>
                                <br /><br />
                                <asp:Label ID="lbliso" runat="server" Text="Select Iso" AssociatedControlID="ddliso"></asp:Label>
                                <asp:DropDownListChosen ID="ddliso" runat="server" NoResultsText="No results match." DataPlaceHolder="Search..." AllowSingleDeselect="true" Width="250px"></asp:DropDownListChosen>
                                <asp:Button ID="btnloadiso" runat="server" OnClick="btnloadiso_Click" Text="Load Iso" /> 
                                <br /><br />
                                <asp:Panel ID="pnlExcelUpload" runat="server" Visible="false" style="flex-grow: 1;">
                                    <br />
                                    <asp:Label ID="lblFileName" runat="server" Font-Bold="true" Font-Size="Large" />
                                    <br />
                                    <asp:GridView ID="ExcelGridView" EnableViewState="true" DataKeyNames="MaterialID" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered" OnRowDataBound="ExcelGridView_RowDataBound">
                                    </asp:GridView>
                                    <br />
                                    <asp:Button ID="btnAddNew" runat="server" Text="Add New Entry" OnClick="btnAddNew_Click" BackColor="Wheat"/>
                                    <asp:Panel ID="pnlAddForm" runat="server" Visible="false">
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
                                                    <asp:DropDownList ID="ddlAddIdent" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlAddIdent_SelectedIndexChanged"/>
                                                    <asp:Label ID="lblidentdesc" runat="server" CssClass="wrap-label"></asp:Label>
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
                                                    <asp:Button ID="btnCancelAdd" runat="server" Text="Cancel/Hide" OnClick="btnCancelAdd_Click" />
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <br />
                                    <asp:Label ID="lblMessage" runat="server" ForeColor="Red" />
                                    <br />
                                    <asp:Button ID="btnSubmit" runat="server" Text="Save & Confirm IsoData" OnClick="btnSubmit_Click" />
                                </asp:Panel>
                                
                            </div>
                        </div>
                    </div>
                </div>
            </div>


            <div class="accordion-item">
                     <h2 class="accordion-header" id="headingThree">
                     <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseThree">
                      Step 2: Working Data
                     </button>
                    </h2>
                <div id="collapseThree" class="accordion-collapse collapse" >
                    <div class="accordion-body">
                        <br />
<asp:Button ID="btnviewInterim" runat="server" Text="Load Working Data" OnClick="btnviewInterim_Click" CssClass="hidden" />
<br />
                        <asp:Panel ID="pnlViewInterimData" runat="server" Visible="false">
                            <asp:Label ID="lblintrim" runat="server" Font-Bold="true" Font-Size="Large" Text="Removing an Item will Remove All for that ISO" ForeColor="Red" />
                            <asp:GridView ID="gvInterim" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered" OnRowCommand="gvInterim_RowCommand" DataKeyNames="INTID,MaterialID,ISO,IsoUniqeRevID">
                                <Columns>
                                    <asp:BoundField DataField="INTID" HeaderText="INTID" Visible="false"/>
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
                                    <asp:BoundField DataField="IsoUniqeRevID" HeaderText="IsoUniqeRevID" />
                                    
                                    <asp:TemplateField HeaderText="Actions">
                                        <ItemTemplate>
                                            <asp:Button ID="btnRemove" runat="server" Text="Remove" CommandName="RemoveRow" CommandArgument='<%# Eval("INTID") %>' OnClientClick="return confirm('Are you sure you want to remove this item?');"/>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                            <br />
                            <asp:Button ID="btnMoveToFinal" runat="server" Text="Move To Final" OnClick="btnMoveToFinal_Click"/>
                        </asp:Panel>
                        
                    </div>
                </div>
            </div>

            <div class="accordion-item">
                <h2 class="accordion-header" id="headingFour">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseFour">
                        Step 3: MTO Export Data
                    </button>
                </h2>
                <div id="collapseFour" class="accordion-collapse collapse" >
                    <div class="accordion-body">
                        <br />
<asp:Button ID="btnViewFinal" runat="server" Text="Load MTO Export Data" OnClick="btnViewFinal_Click" CssClass="hidden" />
<br />
                        <asp:Panel ID="pnlFinalMTO" runat="server" Visible="false">
                            <asp:GridView ID="gvFinal" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered" OnRowCommand="gvFinal_RowCommand" DataKeyNames="MTOID,ISO,IsoUniqeRevID" OnRowDataBound="gvFinal_RowDataBound">
                                <Columns>
                                    <asp:BoundField DataField="MTOID" HeaderText="MTOID" Visible="false"/>
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
                                    <asp:BoundField DataField="ImportStatus" HeaderText="Imported Status" Visible="false"/>
                                    <asp:BoundField DataField="IsoUniqeRevID" HeaderText="IsoUniqeRevID" />
                                    <asp:TemplateField HeaderText="Actions">
    <ItemTemplate>
        <asp:Button ID="btnRemoveFinal" runat="server" Text="Remove" 
            CommandName="RemoveFinalRow" 
            CommandArgument='<%# Eval("MTOID") %>' 
            Visible="false"
            OnClientClick="return confirm('Are you sure you want to remove this item?');" />
    </ItemTemplate>
</asp:TemplateField>

                                </Columns>
                            </asp:GridView>
                            <br />
                            <asp:Button ID="btnExportFinalMTO" runat="server" Text="Export MTO File" OnClick="btnExportFinalMTO_Click"/><br />
                            <asp:Button ID="btnSaveFile" runat="server" Text="Save Exported File" OnClick="btnSaveFile_Click" CssClass="hidden"/><br />
                            
                        </asp:Panel>
                        <br />
                                </div>
    </div>
</div>

            <div class="accordion-item">
    <h2 class="accordion-header" id="headingFive">
        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseFive">
            Step 4: Final MTO Data
        </button>
    </h2>
    <div id="collapseFive" class="accordion-collapse collapse" >
        <div class="accordion-body">
            <br />
<asp:Button ID="btnViewExported" runat="server" Text="Load Final MTO Data" OnClick="btnViewExported_Click"/><br />
                        <asp:Panel ID="pnlFileMan" runat="server" Visible="false">
                             <asp:GridView ID="gvExported" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered"
                                 OnRowCreated="gvExported_RowCreated"  >
     <Columns>
         <asp:BoundField DataField="MTOID" HeaderText="MTOID" Visible="false"/>
         <asp:BoundField DataField="Discipline" HeaderText="Discipline" />
         <asp:BoundField DataField="Area" HeaderText="Area" />
         <asp:BoundField DataField="Unit" HeaderText="Unit" />
         <asp:BoundField DataField="Phase" HeaderText="Phase" />
         <asp:BoundField DataField="Const_Area" HeaderText="Const Area" />
        <asp:TemplateField HeaderText="ISO">
            <HeaderTemplate>
                <span>ISO  </span><asp:DropDownList ID="ddlIsoFilter" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlIsoFilter_SelectedIndexChanged" />
            </HeaderTemplate>
            <ItemTemplate>
                <%# Eval("ISO") %>
            </ItemTemplate>
        </asp:TemplateField>
         <asp:BoundField DataField="Ident_no" HeaderText="Ident No" />
         <asp:BoundField DataField="qty" HeaderText="Quantity" />
         <asp:BoundField DataField="qty_unit" HeaderText="Unit" />
         <asp:BoundField DataField="Spec" HeaderText="Spec" />
         <asp:BoundField DataField="Fabrication_Type" HeaderText="Fabrication Type" />
         <asp:BoundField DataField="IsoRevisionDate" HeaderText="Revision Date" />
         <asp:BoundField DataField="IsoRevision" HeaderText="Revision" />
         <asp:BoundField DataField="IsLocked" HeaderText="Locked" />
         <asp:BoundField DataField="Code" HeaderText="Code" />
         <asp:BoundField DataField="IsoUniqeRevID" HeaderText="IsoUniqeRevID" />
         
         <asp:BoundField DataField="ImportStatus" HeaderText="Imported Status" />
     </Columns>
 </asp:GridView>
 <br />
<asp:Button ID="btnExportFiltered" runat="server" Text="Export Filtered MTO Data" OnClick="btnExportFiltered_Click" />
<br />
<br />
    <asp:GridView ID="grFiles" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered" OnRowCommand="grFiles_RowCommand" DataKeyNames="FileMTOIDs" OnRowDataBound="grFiles_RowDataBound" >
    <Columns>
        <asp:BoundField DataField="FinalFileID" HeaderText="FinalFileID" Visible="false"/>
        <asp:BoundField DataField="FinalFileName" HeaderText="FinalFileName" />
        <asp:TemplateField HeaderText="Download">
    <ItemTemplate>
        <asp:HyperLink ID="lnkDownload" runat="server"
            NavigateUrl='<%# "DownloadFileHelper.aspx?fileID=" + Eval("FinalFileID") %>'
            Text="Download"
            Target="_blank"
            Visible="false" />
    </ItemTemplate>
</asp:TemplateField>


        <asp:BoundField DataField="FileMTOIDs" HeaderText="FileMTOIDs"  Visible="false"/>
        <asp:TemplateField HeaderText="File Imported">
    <ItemTemplate>
        <%# Eval("FileCompleted").ToString().ToLower() == "true" ? "Yes" : "" %>
         <asp:Panel ID="pnlImport" runat="server" Visible='<%# Eval("FileCompleted").ToString().ToLower() == "false" %>'>
            <asp:TextBox ID="txtImportCode" runat="server" MaxLength="5" Width="60px" />
            <asp:Button ID="btnImport" runat="server" Text="Imported" CommandName="Import" CommandArgument='<%# Eval("FinalFileID") %>' />
        </asp:Panel>
    </ItemTemplate>
</asp:TemplateField>
        <asp:BoundField DataField="Import" HeaderText="Import"  />
        
    </Columns>
</asp:GridView>
                        </asp:Panel>
                    </div>
                </div>
          </div>

            <div class="accordion-item" style="display:none;">
    <h2 class="accordion-header" id="headingSix">
        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSix" style="background-color:navajowhite">
            <b>!!</b> Data Revision/Removal/Possible Issues For Reference and Questions NOT PART OF FLOW - Can form part of Reporting <b>!!</b>
        </button>
    </h2>
    <div id="collapseSix" class="accordion-collapse collapse" >
        <div class="accordion-body">
                                        <br />
<asp:Button ID="btnviewmain" runat="server" Text="Load Maintenance Data" OnClick="btnMTOMaintenance_Click"/><br />
<br />
                        <asp:Panel ID="pnlmtorev" runat="server">
                             <asp:GridView ID="gvmtorev" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered" OnRowCommand="gvmtorev_RowCommand" 
                                 DataKeyNames="DelID,MaterialID,MTOID">
     <Columns>
         <asp:BoundField DataField="DelID" HeaderText="DELID" Visible="false"/>
         <asp:BoundField DataField="MaterialID" HeaderText="MaterialID" Visible="false"/>
         <asp:BoundField DataField="Discipline" HeaderText="Discipline" ItemStyle-Wrap="false"  HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Area" HeaderText="Area" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Unit" HeaderText="Unit" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Phase" HeaderText="Phase" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Const_Area" HeaderText="Const_Area" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="ISO" HeaderText="ISO" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Ident_no" HeaderText="Ident_No" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="qty" HeaderText="Qty" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="qty_unit" HeaderText="Qty_Unit" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Spec" HeaderText="Spec" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Fabrication_Type" HeaderText="Fabrication_Type" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="IsoRevisionDate" HeaderText="ISORevDate" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="IsoRevision" HeaderText="ISORev" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="IsLocked" HeaderText="Locked" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Code" HeaderText="Code" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="ImportStatus" HeaderText="Imported Status" ItemStyle-Wrap="false" HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="Changes" HeaderText="Changes"  HeaderStyle-Wrap="false"/>
         <asp:BoundField DataField="MTOID" HeaderText="MTOID" Visible="false" />
         <asp:TemplateField HeaderText="Actions">
    <ItemTemplate>
        <asp:Button ID="btnRemove" runat="server" Text="Actions ?" CommandName="RemoveRow" CommandArgument='<%# Eval("MTOID") %>' OnClientClick="return confirm('Are you sure you want to remove this item?');"/>
    </ItemTemplate>
</asp:TemplateField>
     </Columns>
 </asp:GridView>
 <br />
                        </asp:Panel>
                    </div>
                </div>
          </div>
        <div id="diverror" runat="server" style="display:none;padding-left: 20px;">
            <asp:Label ID="lblerror" runat="server" CssClass="errorlabel"></asp:Label>
        </div>
        </div>
    </form>
</body>
</html>
