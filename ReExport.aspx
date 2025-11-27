
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ReExport.aspx.cs" MaintainScrollPositionOnPostBack="true" Inherits="Wood_MaterialControl.ReExport" EnableEventValidation="false" %>

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
                       Modify ISO for Re-Export
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
                               <br />
                                <asp:Label ID="lbliso" runat="server" Text="Select Iso" AssociatedControlID="ddliso"></asp:Label>
                                <asp:DropDownListChosen ID="ddliso" runat="server" NoResultsText="No results match." DataPlaceHolder="Search..." AllowSingleDeselect="true" Width="250px"></asp:DropDownListChosen>
                                <asp:Button ID="btnloadiso" runat="server" OnClick="btnloadiso_Click" Text="Load Iso" />
                                <br /><br />
                                <asp:Panel ID="pnlisodata" runat="server" Visible="false" style="flex-grow: 1;">
                                    <br />
                                    <asp:Label ID="lblFileName" runat="server" Font-Bold="true" Font-Size="Large" />
                                    <br />
                                      <asp:GridView ID="grisoreview" runat="server" AutoGenerateColumns="False" CssClass="table table-bordered"
                    DataKeyNames="MTOID" OnRowDataBound="grisoreview_RowDataBound">
                </asp:GridView>
<br />
                                    <asp:Button ID="btnUpdate_ISO" runat="server" OnClick="btnUpdate_ISO_Click" Text="Update MTO for Re-Export" />
                                    <br />
                                    <br />
<asp:Button ID="btnReExportMTO" runat="server" Text="Re-Export MTO"   OnClick="btnReExportMTO_Click" Visible="false" />

                                    <br />
                                </asp:Panel>
                            </div>
                        </div>
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
