<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ListUpdate.aspx.cs" Inherits="waDEIGcsharp.ListUpdates.ListUpdate" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="FeaturedContent" runat="server">
      <asp:SqlDataSource ID="SqlDsDOW" runat="server" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" SelectCommand="SELECT [Id], [DayName] FROM [DOW] ORDER BY [Id]"></asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDsTown" runat="server" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" SelectCommand="SELECT [Town] FROM [vTown] ORDER BY [Town]"></asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDsTime" runat="server" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" SelectCommand="SELECT TimeID, xTime FROM aTime ORDER BY TimeID"></asp:SqlDataSource>
    <Section class="inputs">    
        <table>
	        <tr>
		        <td class="head"><asp:Label ID="lblDay" runat="server" Text="Day of Week"></asp:Label></td>
		        <td class="head"> <asp:Label ID="lblTime" runat="server" Text="Time"></asp:Label>  </td>
		        <td class="head"><asp:Label ID="lblTown" runat="server" Text="Town"></asp:Label></td>
                <td class="head"><asp:Label ID="lblSelect" runat="server" Text="Select"></asp:Label></td>
	        </tr>
	        <tr>
		        <td>    
                    <asp:DropDownList ID="ddlDOW" runat="server" DataSourceID="SQLDsDow" DataTextField="DayName" DataValueField="Id" Font-Bold="True" Font-Size="Medium" >
                    </asp:DropDownList>              
                </td>
                <td>                  
                    <asp:DropDownList ID="ddlTime" runat="server" DataSourceID="SqlDsTime" DataTextField="xTime" DataValueField="TimeID">
                    </asp:DropDownList>                 
                    </td>
                <td>
                    <asp:DropDownList ID="ddlTown" runat="server" DataSourceID="SqlDsTown" DataTextField="Town" DataValueField="Town">
                    </asp:DropDownList>                 
                </td>
                <td>
                    <asp:LinkButton ID="lbGet" runat="server" OnClick="lbGet_Click" >Get Meetings</asp:LinkButton>
                </td>
            </tr>
         </table>
             <asp:Label ID="lblInfo" runat="server" Text="Feedback:"></asp:Label>
            <br />
            <asp:LinkButton ID="lnkInsert" runat="server" OnClick="lnkInsert_Click">Add New Meeting</asp:LinkButton>
         </Section>
       
      
        <asp:Panel ID="PnlAdd" runat="server" Visible="False" >
            <section class="inputsBorder">
            <h3>Add New Meeting</h3>
            <asp:Table ID="Table2" runat="server">
                <asp:TableRow>
                    <asp:TableHeaderCell>Day</asp:TableHeaderCell>
                    <asp:TableHeaderCell>Time</asp:TableHeaderCell>
                    <asp:TableHeaderCell>Town</asp:TableHeaderCell>
                </asp:TableRow>
                <asp:TableRow>
                     <asp:TableCell>
                        <asp:DropDownList ID="ddlDOWInsert" runat="server" DataSourceID="SqlDsDOW" DataTextField="DayName" DataValueField="Id">
                                </asp:DropDownList>
                     </asp:TableCell>
                     <asp:TableCell>
                        <asp:DropDownList ID="ddlTimeInsert" runat="server" DataSourceID="SqlDsTime" DataTextField="xTime" DataValueField="TimeID" >
                        </asp:DropDownList>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:TextBox ID="txtTown" runat="server"></asp:TextBox>                       
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableHeaderCell>Group Name</asp:TableHeaderCell>
                     <asp:TableHeaderCell>Information</asp:TableHeaderCell>
                    <asp:TableHeaderCell>Location</asp:TableHeaderCell>
                </asp:TableRow>
                <asp:TableRow>                    
                    <asp:TableCell>
                        <asp:TextBox ID="tbGroupName" runat="server"></asp:TextBox>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:TextBox ID="tbInfo" runat="server"></asp:TextBox>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:TextBox ID="tbLocation" runat="server"></asp:TextBox>
                    </asp:TableCell>
                </asp:TableRow>               
                <asp:TableRow>
                    <asp:TableHeaderCell>Type</asp:TableHeaderCell>
                    <asp:TableHeaderCell>Update</asp:TableHeaderCell>
                    <asp:TableHeaderCell></asp:TableHeaderCell>
                </asp:TableRow>               
                <asp:TableRow>  
                    <asp:TableCell>
                        <asp:TextBox ID="tbType" runat="server"></asp:TextBox>
                    </asp:TableCell>                   
                    <asp:TableCell>
                        <asp:LinkButton ID="lnkSave" runat="server" OnClick="lnkSave_Click">Save</asp:LinkButton>
                    </asp:TableCell>  
                    <asp:TableCell>
                       
                    </asp:TableCell>                  
                </asp:TableRow>
            </asp:Table>
                
            </section>            
        </asp:Panel>
  
    <section class="inputs">
    <asp:SqlDataSource ID="SqlDsList" runat="server" ConnectionString="<%$ ConnectionStrings:cnDEIG %>" DeleteCommand="DELETE FROM [List] WHERE [ListID] = @ListID" InsertCommand="INSERT INTO [List] ([DOW], [TimeID], [Town], [GroupName], [Information], [Location], [Type]) VALUES (@DOW, @TimeID, @Town, @GroupName, @Information, @Location, @Type)" SelectCommand="SELECT [ListID], [DOW], [TimeID], [Town], [GroupName], [Information], [Location], [Type] FROM [List] ORDER BY [DOW], [TimeID], [Town]" UpdateCommand="UPDATE [List] SET [DOW] = @DOW, [TimeID] = @TimeID, [Town] = @Town, [GroupName] = @GroupName, [Information] = @Information, [Location] = @Location, [Type] = @Type WHERE [ListID] = @ListID">
            <DeleteParameters>
                <asp:Parameter Name="ListID" Type="Int32" />
            </DeleteParameters>
            <InsertParameters>
                <asp:Parameter Name="DOW" Type="Int32" />
                <asp:Parameter Name="TimeID" Type="Int32" />
                <asp:Parameter Name="Town" Type="String" />
                <asp:Parameter Name="GroupName" Type="String" />
                <asp:Parameter Name="Information" Type="String" />
                <asp:Parameter Name="Location" Type="String" />
                <asp:Parameter Name="Type" Type="String" />
            </InsertParameters>
            <UpdateParameters>
                <asp:Parameter Name="DOW" Type="Int32" />
                <asp:Parameter Name="TimeID" Type="Int32" />
                <asp:Parameter Name="Town" Type="String" />
                <asp:Parameter Name="GroupName" Type="String" />
                <asp:Parameter Name="Information" Type="String" />
                <asp:Parameter Name="Location" Type="String" />
                <asp:Parameter Name="Type" Type="String" />
                <asp:Parameter Name="ListID" Type="Int32" />
            </UpdateParameters>
        </asp:SqlDataSource>
        <asp:GridView ID="GridView1" runat="server" AllowSorting="True" AutoGenerateColumns="False" CellPadding="4" DataKeyNames="ListID" DataSourceID="SqlDsList" ForeColor="#333333" GridLines="None">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <Columns>
                <asp:CommandField ShowEditButton="True" ShowSelectButton="True" />
                <asp:TemplateField HeaderText="Day" SortExpression="DOW">
                    <EditItemTemplate>
                        <asp:DropDownList ID="ddlDOWedit" runat="server" AutoPostBack="True" DataSourceID="SqlDsDOW" DataTextField="DayName" DataValueField="Id" SelectedValue='<%# Bind("DOW") %>'>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:DropDownList ID="ddlDOWdiplay" runat="server" DataSourceID="SqlDsDOW" DataTextField="DayName" DataValueField="Id" Enabled="False" SelectedValue='<%# Bind("DOW") %>' BackColor="White" Font-Bold="True">
                        </asp:DropDownList>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Time" SortExpression="TimeID">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList3" runat="server" AutoPostBack="True" DataSourceID="SqlDsTime" DataTextField="xTime" DataValueField="TimeID" SelectedValue='<%# Bind("TimeID") %>' >
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:DropDownList ID="ddlTimeDisplay" runat="server" DataSourceID="SqlDsTime" DataTextField="xTime" DataValueField="TimeID" Enabled="False" SelectedValue='<%# Bind("TimeID") %>' BackColor="White" Font-Bold="True">
                        </asp:DropDownList>
                    </ItemTemplate>
                    <ItemStyle Width="55px" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Town" SortExpression="Town">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList4" runat="server" DataSourceID="SqlDsTown" DataTextField="Town" DataValueField="Town" SelectedValue='<%# Bind("Town") %>'>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("Town") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle Width="60px" />
                </asp:TemplateField>
                <asp:BoundField DataField="GroupName" HeaderText="GroupName" SortExpression="GroupName" />
                <asp:BoundField DataField="Information" HeaderText="Information" SortExpression="Information" />
                <asp:BoundField DataField="Location" HeaderText="Location" SortExpression="Location" />
                <asp:BoundField DataField="Type" HeaderText="Type" SortExpression="Type" />
                <asp:CommandField ShowDeleteButton="True" />
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
        <asp:LinkButton ID="LnkExport" runat="server" OnClick="LnkExport_Click">Export List to text file</asp:LinkButton>
    </section>  
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">
</asp:Content>
