<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="waDEIGcsharp._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
 <asp:SqlDataSource ID="SqlDsList" runat="server" DataSourceMode="DataSet" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" SelectCommand="SELECT [DOW], [TimeID], [Town], [GroupName], [Information], [Location], [Type] FROM [List] ORDER BY [DOW], [TimeID], [Town]" />

        <div class="content-wrapper">
            <hgroup class="title">
                <h1><%: Title %></h1>
            </hgroup>
           
        </div>
        <div class="centerHeading" >
            <table  class="centerHeading">
	        <tr>
		        <td class="head"><asp:Label ID="lblDay" runat="server" Text="Day of Week"></asp:Label></td>
		        <td class="head"> <asp:Label ID="lblTime" runat="server" Text="Time"></asp:Label>  </td>
		        <td class="head"><asp:Label ID="lblTown" runat="server" Text="Town"></asp:Label></td>
                <td class="head"><asp:Label ID="lblSelect" runat="server" Text="Select"></asp:Label></td>
	        </tr>
	        <tr>
		        <td>    
                    <asp:DropDownList ID="ddlDOW" runat="server" DataSourceID="SQLDsDow" DataTextField="DayName" DataValueField="Id" Font-Bold="True" Font-Size="Medium"  >
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SQLDsDow" runat="server" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" SelectCommand="SELECT [Id], [DayName] FROM [DOW] ORDER BY [Id]">
                    </asp:SqlDataSource>
                </td>
                <td>                  
                    <asp:DropDownList ID="ddlTime" runat="server" DataSourceID="SqlDsTime" DataTextField="xTime" DataValueField="TimeID">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDsTime" runat="server" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" SelectCommand="SELECT [TimeID], [xTime] FROM [aTime] ORDER BY [TimeID]"></asp:SqlDataSource>
                    </td>
                <td>
                    <asp:DropDownList ID="ddlTown" runat="server" DataSourceID="SqlDsTown" DataTextField="Town" DataValueField="Town">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDsTown" runat="server" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" SelectCommand="SELECT [Town] FROM [vTown] ORDER BY [Town]"></asp:SqlDataSource>    
                </td>
                <td>
                    <asp:LinkButton ID="lbGet" runat="server" OnClick="lbGet_Click" >Get Meetings</asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:Label ID="lblFilter" runat="server" Text="Feed Back:"></asp:Label></td>
            </tr>
         </table>
        </div>            
    </section>
    <section>       
         <div class="center">

            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4" DataSourceID="SqlDsList" ForeColor="#333333" GridLines="None">
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <Columns>
                    <asp:TemplateField HeaderText="Day" SortExpression="DOW">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("DOW") %>'></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:DropDownList ID="DropDownList2" runat="server" DataSourceID="SQLDsDow" DataTextField="DayName" DataValueField="Id" SelectedValue='<%# Bind("DOW") %>' BackColor="White" Enabled="False" Font-Bold="True" Font-Size="Large">
                            </asp:DropDownList>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Time" SortExpression="TimeID">
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox2" runat="server" Text='<%# Bind("TimeID") %>'></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:DropDownList ID="DropDownList5" runat="server" DataSourceID="SqlDsTime" DataTextField="xTime" DataValueField="TimeID" SelectedValue='<%# Bind("TimeID") %>' BackColor="White" Enabled="False" Font-Bold="True" Font-Size="Medium">
                            </asp:DropDownList>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="Town" HeaderText="Town" SortExpression="Town" />
                    <asp:BoundField DataField="GroupName" HeaderText="Group Name" SortExpression="GroupName" />
                    <asp:BoundField DataField="Information" HeaderText="Information" SortExpression="Information" />
                    <asp:BoundField DataField="Location" HeaderText="Location" SortExpression="Location" />
                    <asp:BoundField DataField="Type" HeaderText="Type" SortExpression="Type" />
                </Columns>
                <EditRowStyle BackColor="#999999" />
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                <SortedAscendingCellStyle BackColor="#E9E7E2" />
                <SortedAscendingHeaderStyle BackColor="#506C8C" />
                <SortedDescendingCellStyle BackColor="#FFFDF8" />
                <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
            </asp:GridView>
            <br />
            <asp:LinkButton ID="lnkExport" runat="server" OnClick="lnkExport_Click">Export List to text file</asp:LinkButton>
        </div>
    </section>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    </asp:Content>
