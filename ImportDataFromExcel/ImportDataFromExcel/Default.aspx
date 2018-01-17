<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ImportDataFromExcel._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <h3>Prospect Client List</h3>
    <hr />
    <div>
        
        <div>
            <table>
                <tr>
                    <td>Select File: </td>
                    <td>
                        <asp:FileUpload ID="FileUpload1" runat="server" CssClass="btn btn-primary" /></td>
                    <td>
                        <asp:Button ID="btnImportFromCSV" runat="server" Text="Import" CssClass="btn btn-primary" OnClick="btnImportFromCSV_Click" />
                    &nbsp;&nbsp; 
                    </td>
                </tr>
            </table>
            <div>
                <em>
                <asp:Label ID="lblMessage" runat="server" style="color: #33CC33; font-size: small" />
                        </em>
                <br />
                <br />
                <br />
                <asp:GridView ID="gvData" runat="server" AutoGenerateColumns="false" Height="170px" OnSelectedIndexChanged="gvData_SelectedIndexChanged" Width="592px">
                    <EmptyDataTemplate>
                        <div style="padding: 10px;">
                            <h4>Please import Excel spread sheet</h4>
                        </div>
                    </EmptyDataTemplate>
                    <Columns>
                        <asp:BoundField HeaderText="Client ID" HeaderStyle-BackColor="Black" HeaderStyle-ForeColor="AntiqueWhite" HeaderStyle-Font-Size="Medium" DataField="ClientID" />
                        <asp:BoundField HeaderText="First Name" HeaderStyle-BackColor="Black" HeaderStyle-ForeColor="AntiqueWhite" HeaderStyle-Font-Size="Medium" DataField="First_Name" />
                        <asp:BoundField HeaderText="Last Name" HeaderStyle-BackColor="Black" HeaderStyle-ForeColor="AntiqueWhite" HeaderStyle-Font-Size="Medium" DataField="Last_Name" />
                        <asp:BoundField HeaderText="Employer" HeaderStyle-BackColor="Black" HeaderStyle-ForeColor="AntiqueWhite" HeaderStyle-Font-Size="Medium" DataField="Employer" />
                        <asp:BoundField HeaderText="Title" HeaderStyle-BackColor="Black" HeaderStyle-ForeColor="AntiqueWhite" HeaderStyle-Font-Size="Medium" DataField="Title" />
                        <asp:BoundField HeaderText="Phone Number" HeaderStyle-BackColor="Black" HeaderStyle-ForeColor="AntiqueWhite" HeaderStyle-Font-Size="Medium" DataField="Phone_Number" />
                        <asp:BoundField HeaderText="Zip" HeaderStyle-BackColor="Black" HeaderStyle-ForeColor="AntiqueWhite" HeaderStyle-Font-Size="Medium" DataField="Zip" />
                    </Columns>
                </asp:GridView>
                <br />

            </div>
        </div>
    </div>

</asp:Content>

