<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="Williams.Lea.Symphony.Support" %>
<%@ Import Namespace="Williams.Lea.Symphony.Support.SupportFunctions" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Williams.Lea.Symphony.SymphonyWebService" %>
<%@ Import Namespace="Williams.Lea.Symphony.InterfaceLib" %>
<%@ Import Namespace="Williams.Lea.Symphony.UIObjects" %>
<%@ Register Assembly="Williams.Lea.WebControls" Namespace="Williams.Lea.WebControlLibrary.UIWebControls"
    TagPrefix="cc1" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register Assembly="EO.Web" Namespace="EO.Web" TagPrefix="eo" %>

<%@ Page ClassName="RPT_CSFB_PERSONNEL" Inherits="Williams.Lea.Symphony.UIObjects.BasePage"
    Language="vb" StylesheetTheme="Default" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>RPT_CSFB_PERSONNEL</title>
    <script runat="server">
        
        Public Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
           
            labelCriteria.Text = String.Format("Personnel Report")
            txtFrom.FormatString = Session("DATE_FORMAT")
            txtTo.FormatString = Session("DATE_FORMAT")
            
            'Was this a post request
            If Not Me.IsPostBack() Then
                
                Dim dsSite As DataSet
                Dim dsJobType As DataSet
                
                txtFrom.Text = Me.currentDate.Month & "/1/" & Me.currentDate.Year()
                txtTo.Text = Me.currentDate.Date().ToString
                	           
                'Load Site List based on USER
                dsSite = websrv.GetUserSites(Williams.Lea.Symphony.Support.SupportFunctions.GetSAK, User.Identity.Name)
                If Not dsSite Is Nothing Then
                    ddSite.DataSource = dsSite.Tables(0).DefaultView
                    ddSite.DataTextField = "SITE"
                    ddSite.DataValueField = "SITE"
                    ddSite.DataBind()
                End If
                
                'Load Type List based on site
                dsJobType = websrv.GetSiteJobTypes(Williams.Lea.Symphony.Support.SupportFunctions.GetSAK, Session("SITE"))
                If Not dsJobType Is Nothing Then
                    ddType.DataSource = dsJobType.Tables(0).DefaultView
                    ddType.DataTextField = "JOB_TYPE"
                    ddType.DataValueField = "JOB_TYPE"
                    ddType.DataBind()
                End If
            Else
                
                If txtFrom.Text.Trim.Length = 0 Then txtFrom.Text = Me.currentDate.Date()
                If txtTo.Text.Trim.Length = 0 Then txtTo.Text = txtFrom.Text
               
                'Swap dates if needed
                If CDate(txtFrom.Text) > CDate(txtTo.Text) Then
                    Dim tmp As String = txtFrom.Text
                    txtFrom.Text = txtTo.Text
                    txtTo.Text = tmp
                End If
                
                If (IsNothing(Request.Form("ButtonExecute")) AndAlso (Not IsNothing(Request.Form("__EVENTTARGET"))) AndAlso (Request.Form("__EVENTTARGET").IndexOf("ToolbarMainItemClicked", StringComparison.OrdinalIgnoreCase) > -1)) Then
                    ToolBar1_ItemClicked(Request("__EVENTARGUMENT"))
                End If
                
            End If
            
            ClientScript.RegisterStartupScript(Me.GetType(), "SCROLL_CHECK", Williams.Lea.Symphony.InterfaceLib.GenerateScrollScript())
            
        End Sub
						
		           
        Private Function GenerateData() As DataSet
           
            Dim ds As DataSet
            Dim sql As StringBuilder = New StringBuilder
            Dim frmd8 As Date = Date.Parse(txtFrom.Text)
            Dim tod8 As Date = Date.Parse(txtTo.Text)
            Dim txtSite As String = ""
            Dim txtType As String = ddType.SelectedValue
           
            For Each li As ListItem In ddSite.Items
                If li.Selected = True Then
                    If txtSite = "" Then
                        txtSite = li.Text
                    Else
                        txtSite = txtSite + ", " + li.Text
                    End If
                End If
            Next
            If txtSite = "" Then
                Me.PopupMessage("Please select atleast one site")
                Return Nothing
            End If
            labelCriteria1.Text = String.Format("Started between {0} and {1}", frmd8.ToString(Session("DATE_FORMAT").ToString.Substring(0, 10)), tod8.ToString(Session("DATE_FORMAT").ToString.Substring(0, 10)))
            labelCriteria2.Text = String.Format("Site(s) : {0}  /  Job Type : {1} ", txtSite, txtType)

            'Get production data.
            sql.AppendFormat("RPT_CSFB_PERSONNEL  '{0}', '{1}', '{2}', '{3}', '{4}', '{5}'", txtSite.Replace("'", "").Replace(" ", ""), txtType.Replace("'", "").Replace(",", "/"), Me.websrv.AdjustToServerTimeZone(GetSAK(), txtFrom.Text & " 00:00:00"), Me.websrv.AdjustToServerTimeZone(Williams.Lea.Symphony.Support.SupportFunctions.GetSAK, txtTo.Text & " 23:59:59.998"), Me.User.Identity.Name, Me.Page.ToString)
            Try
                ds = Me.websrv.ExecuteQuery(Williams.Lea.Symphony.Support.SupportFunctions.GetSAK, Williams.Lea.Symphony.SymphonyWebService.Databases.JOBS, sql.ToString)
            Catch ex As Exception
                Me.PopupMessage(ex.Message)
                Me.websrv.LogEvent(Williams.Lea.Symphony.Support.SupportFunctions.GetSAK, Session("SITE"), Me.User.Identity.Name, Me.Page.ToString, "Personnel Report Ran.".ToUpper)
                Return Nothing
            End Try
            
            Return ds
    
        End Function
        
    
        Private Sub ButtonExecute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
           
            Dim ds As DataSet = GenerateData()

            If Not ds Is Nothing Then
                grdData.DataSource = ds.Tables(0).DefaultView
                grdData.DataBind()
            End If

        End Sub
        

        Protected Sub ToolBar1_ItemClicked(ByVal item As String)

            Select Case item.ToUpper
                
                Case "EXPORT"
                    ExportReport()
                    
            End Select

        End Sub

        
        Private Sub ExportReport()
           
            Dim ds As DataSet = GenerateData()
           
            If Not ds Is Nothing Then
                grdData.DataSource = ds.Tables(0).DefaultView
                grdData.DataBind()
                ExportDatasetToExcel.DataGridToExcel(grdData, ReportTitle)
            End If
            
        End Sub
        
	
        Private Sub GrdData_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
          
            Select Case e.Item.ItemType
                
                Case ListItemType.AlternatingItem, ListItemType.Item, ListItemType.SelectedItem
                   
                    For Each cell As TableCell In e.Item.Cells
                        cell.Wrap = False
                        cell.BorderStyle = BorderStyle.Solid
                        cell.BorderColor = Color.LightGray
                    Next
                
            
            End Select

        End Sub

		
    </script>
    <script language="javascript" type="text/javascript">

        function PrintReport() {
            if (document.all("grdData")) {
                document.all("PanelToolbar").style.visibility = 'hidden';
                document.all("ReportTitle").style.top = 0;
                document.all("ResultData").style.top = 48;
                window.print();
                document.all("ResultData").style.top = 180;
                document.all("ReportTitle").style.top = 125;
                document.all("PanelToolbar").style.visibility = 'visible';
            }
            else
                alert("No Data Found");
        }

        function MainItemClicked(toolBar, toolBarItem) {

            switch (toolBarItem.getCommandName()) {
                case 'Print':
                    PrintReport()
                    break;
                case 'Export':
                    __doPostBack('ToolbarMainItemClicked', toolBarItem.getCommandName());
                    break;
                case 'Exit':
                    window.close();
                    break;
            }
        }
    </script>
</head>
<body>
    <form id="FormMain" runat="server" method="post">
        <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="600">
        </asp:ScriptManager>
        <asp:Panel ID="PanelToolbar" SkinID="pnlHeaderBGColor" runat="server" Font-Bold="false"
            Height="120px" Style="z-index: 1030; left: 0px; position: absolute; top: 0px">
            <div style="clear: left; border-bottom-style: solid; border-bottom-width: 1px; font-family: Arial;"
                class="left">
                <eo:ToolBar ID="ToolBar1" runat="server" TextAlign="Underneath" class="SideContainer"
                    ClientSideOnItemClick="MainItemClicked" Height="40px">
                    <Items>
                        <eo:ToolBarItem CommandName="Print" ImageUrl="~/Images/print.jpg" Text="Print" ToolTip="Print the report"
                            Type="Button">
                            <NormalStyle CssText="width: 50px" />
                        </eo:ToolBarItem>
                        <eo:ToolBarItem CommandName="Export" ImageUrl="~/Images/excel.jpg" Text="Export"
                            ToolTip="Export to excel file" Type="Button">
                            <NormalStyle CssText="width: 50px" />
                        </eo:ToolBarItem>
                        <eo:ToolBarItem CommandName="Exit" ImageUrl="~/Images/cancel.gif" Text="Exit" ToolTip="Exit Report"
                            Type="Button">
                            <NormalStyle CssText="width: 50px" />
                        </eo:ToolBarItem>
                    </Items>
                    <ItemTemplates>
                        <eo:ToolBarItem Type="DropDownMenu">
                            <NormalStyle CssClass="mTopNormalStyle" />
                            <HoverStyle CssClass="mTopHoverStyle" />
                            <DownStyle CssText="background-color:transparent;border-bottom-style:none;border-left-style:none;border-right-style:none;border-top-style:none;padding-bottom:1px;padding-left:2px;padding-right:2px;padding-top:2px;" />
                        </eo:ToolBarItem>
                        <eo:ToolBarItem>
                            <NormalStyle CssClass="mTopNormalStyle" />
                            <HoverStyle CssClass="mTopHoverStyle" />
                        </eo:ToolBarItem>
                    </ItemTemplates>
                    <NormalStyle CssClass="mTopNormalStyle" />
                    <HoverStyle CssClass="mTopHoverStyle" />
                    <DownStyle CssText="BORDER-RIGHT: #335ea8 1px solid; PADDING-RIGHT: 2px; BORDER-TOP: #335ea8 1px solid; PADDING-LEFT: 4px; BACKGROUND-POSITION-X: right; BACKGROUND-IMAGE: url(00100305); PADDING-BOTTOM: 1px; BORDER-LEFT: #335ea8 1px solid; CURSOR: hand; PADDING-TOP: 3px; BORDER-BOTTOM: #335ea8 1px solid;" />
                </eo:ToolBar>
            </div>
            <table border="0" cellpadding="0" cellspacing="0" style="margin: 0; position: relative; top: 10px"
                width="100%">
                <tr style="vertical-align: middle">
                    <td style="text-align: right">
                        <asp:Label runat="server" ID="label1" Text="Site" Font-Bold="true" Width="60px"></asp:Label>
                    </td>
                    <td style="width: 100px; height: 30px;">
                        <asp:CheckBoxList ID="ddSite" RepeatColumns="5" RepeatDirection="Horizontal" RepeatLayout="Table" runat="server" CellPadding="1" CellSpacing="1" TextAlign="Right"></asp:CheckBoxList>

                    </td>
                    <td style="text-align: right">
                        <asp:Label runat="server" ID="label2" Text="Type" Font-Bold="true" Width="60px"></asp:Label>
                    </td>
                    <td style="width: 90px;">
                        <asp:DropDownList ID="ddType" runat="server" Font-Size="9" Height="18px" Width="88px">
                        </asp:DropDownList>
                    </td>
                    <td style="width: 80px; text-align: right">
                        <asp:Label runat="server" ID="label3" Text="From" Font-Bold="true" Width="60px"></asp:Label>
                    </td>
                    <td style="width: 96px;">
                        <cc1:DateTimePicker ID="txtFrom" runat="server" Font-Size="9" Height="18px" Width="85px" />
                    </td>
                    <td style="width: 80px; text-align: right">
                        <asp:Label runat="server" ID="label4" Text="To" Font-Bold="true" Width="60px"></asp:Label>
                    </td>
                    <td style="width: 94px;">
                        <cc1:DateTimePicker ID="txtTo" runat="server" Font-Size="9" Height="18px" Width="85px" />
                    </td>
                    <td>
                        <telerik:RadButton runat="server" ID="ButtonExecute" Text="Execute" ToolTip="Execute"
                            OnClick="ButtonExecute_Click" Height="22px" Width="70px">
                        </telerik:RadButton>
                    </td>
                    <td style="width: 100%"></td>
                </tr>
            </table>
        </asp:Panel>
        <div id="ReportTitle" runat="server" style="display: inline; font-weight: bold; font-size: 12pt; z-index: 114; left: 0px; width: 100%; position: absolute; top: 125px; height: 40px; text-align: center">
            <asp:Label ID="labelCriteria" runat="server" Font-Bold="True" Font-Size="11pt"></asp:Label><br />
            <asp:Label ID="labelCriteria1" runat="server" Font-Bold="True" Font-Size="9pt"></asp:Label><br />
            <asp:Label ID="labelCriteria2" runat="server" Font-Bold="True" Font-Size="9pt"></asp:Label><br />
        </div>
        <div id="ResultData" runat="server" style="z-index: 104; left: 0px; width: 100%; position: absolute; top: 180px; height: 95%">
            <asp:DataGrid ID="grdData" runat="server" AutoGenerateColumns="true" SkinID="gv"
                BorderWidth="1px" Height="72px" OnItemDataBound="grdData_ItemDataBound" Style="z-index: 112; left: 0px; position: absolute; top: 1px"
                Width="100%" EnableViewState="false">
            </asp:DataGrid>
        </div>
    </form>
</body>
</html>
