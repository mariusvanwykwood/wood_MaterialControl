<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ConstructionAreas.aspx.cs" Inherits="Wood_MaterialControl.ConstructionAreas" EnableEventValidation="false" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
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
            <h2 runat="server" id="mainh1" style="text-align:left;">Iso Control</h2> 
 <br />
            <asp:Label ID="lblMainHeader" runat="server"  Text="Construction Areas"></asp:Label><br />
<br />
            What Now  ??:D
           </div>
        <div id="diverror" runat="server" style="display:none;padding-left: 20px;">
            <asp:Label ID="lblerror" runat="server" CssClass="errorlabel"></asp:Label>
        </div>
    </form>
</body>
</html>
