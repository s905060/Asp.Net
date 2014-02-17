<%@ Page Language="VB" AutoEventWireup="true" CodeFile="Default.aspx.vb" Inherits="_Default" %>
<%@ Import Namespace="System.data" %>
<%@ Import Namespace="System.Data.OleDb"%>

<script runat="server" language="vbscript">
       
    Public authorization As Integer
    Public name As String
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Session("Login") = True Then
            'Response.Write("您無權進入此網頁，勿存僥倖的心!")
            Response.Redirect("login.aspx")
            Response.End()
        End If
        
        GridView1.Visible = False
        GridView2.Visible = False
        GridView3.Visible = False
        GridView4.Visible = False
        GridView21.Visible = False
        
        DetailsView1.Visible = False
        DetailsView2.Visible = False
        DetailsView3.Visible = False
        DetailsView4.Visible = False
        DetailsView21.Visible = False
        
        Try
            Dim connStr, Showcmd As String
            connStr = "provider=microsoft.jet.oledb.4.0; data source= C:\Inetpub\wwwroot\WEBSITE\App_Data\CEMCL.mdb"
            Showcmd = "select max(報價單號碼) from Quotation"
            Dim cmd As OleDbCommand, conn As OleDbConnection
            conn = New OleDbConnection(connStr)
            conn.Open()
            cmd = New OleDbCommand(Showcmd, conn)
            Label6.Text = cmd.ExecuteScalar
            conn.Close()
            msg.Text = Session("name") + " 歡迎您登入 "
            Label2.Text = Date.Today
        Catch ex As Exception
            msg.Text = " 檢查錯誤 "
        End Try
                
        If Session("authorization") = 1 Then
            GridView1.Visible = True
            DetailsView1.Visible = True
        ElseIf Session("authorization") = 2 Then
            GridView2.Visible = True
            DetailsView2.Visible = True
        ElseIf Session("authorization") = 3 Then
            GridView3.Visible = True
            DetailsView3.Visible = True
        ElseIf Session("authorization") = 4 Then
            GridView4.Visible = True
            DetailsView4.Visible = True
        ElseIf Session("authorization") = 21 Then
            GridView21.Visible = True
            DetailsView21.Visible = True
        End If
    End Sub
    
    Protected Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Dim connStr, Showcmd As String
            connStr = "provider=microsoft.jet.oledb.4.0; data source= C:\Inetpub\wwwroot\WEBSITE\App_Data\CEMCL.mdb"
            Showcmd = "select max(報價單號碼) from Quotation"
            Dim cmd As OleDbCommand, conn As OleDbConnection
            conn = New OleDbConnection(connStr)
            conn.Open()
            cmd = New OleDbCommand(Showcmd, conn)
            Label6.Text = cmd.ExecuteScalar
            conn.Close()
            msg.Text = Session("name") + " 歡迎您登入 "
            Label2.Text = Date.Today
        Catch ex As Exception
            msg.Text = " 檢查錯誤 "
        End Try
    End Sub   

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session.RemoveAll()
        Session.Clear()
        Session.Abandon()
        Response.Redirect("login.aspx")
    End Sub
</script>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>青華企業報價單登錄系統</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />
        <div style="text-align: left">
            <table>
                <tr>
                    <td rowspan="1" style="width: 289px; height: 21px; text-align: left;">
                        <asp:Image ID="Image1" runat="server" ImageUrl="~/CEMCL logo.jpg" /><asp:Label ID="Label3"
                            runat="server" Font-Bold="True" ForeColor="RoyalBlue" Text="青華企業報價單登錄系統 Beta 3"
                            Width="250px"></asp:Label><br />
                        <br />
                        <asp:Label ID="msg" runat="server" EnableViewState="False" Font-Bold="True" Font-Size="Medium"
                            ForeColor="RoyalBlue"></asp:Label><br />
                        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                <Triggers>
                                    <asp:AsyncPostBackTrigger ControlID="Timer1" EventName="Tick" />
                                </Triggers>
                                <ContentTemplate>
                        <asp:Label ID="Label5" runat="server" Font-Bold="True" Font-Size="Medium" ForeColor="RoyalBlue"
                            Text="最末筆報價單號為:"></asp:Label>
                                    <asp:Label ID="Label6" runat="server" Text="Label" Font-Bold="True" ForeColor="Red" Font-Size="XX-Large"></asp:Label>
                                </ContentTemplate>
                            </asp:UpdatePanel>
                        <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="Medium" ForeColor="RoyalBlue"
                            Text="今天是:"></asp:Label>
                        <asp:Label ID="Label2" runat="server" Font-Bold="True" ForeColor="Red" Font-Size="XX-Large"></asp:Label><br />
                        <asp:Timer ID="Timer1" runat="server" Interval="100">
                        </asp:Timer>
                        <br />
                        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="系統登出" /></td>
                    <td rowspan="1" style="width: 544px; height: 21px">
                        <asp:DetailsView ID="DetailsView1" runat="server" AutoGenerateRows="False" CellPadding="4"
                            DataKeyNames="報價單號碼" DataSourceID="AccessDataSource1" DefaultMode="Insert" ForeColor="#333333"
                            GridLines="None" Height="50px" Width="280px">
                            <Fields>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼" />
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期" />
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱" />
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱" />
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目" />
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者" />
                                <asp:CommandField ButtonType="Button" InsertText="新增" ShowCancelButton="False" ShowInsertButton="True">
                                    <ItemStyle ForeColor="Red" />
                                </asp:CommandField>
                            </Fields>
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <CommandRowStyle BackColor="#D1DDF1" Font-Bold="True" />
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FieldHeaderStyle BackColor="#DEE8F5" Font-Bold="True" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" />
                            <AlternatingRowStyle BackColor="White" />
                            <InsertRowStyle Wrap="False" />
                        </asp:DetailsView>
                        <asp:DetailsView ID="DetailsView2" runat="server" AutoGenerateRows="False" CellPadding="4"
                            DataKeyNames="報價單號碼" DataSourceID="AccessDataSource2" DefaultMode="Insert" ForeColor="#333333"
                            GridLines="None" Height="50px" Width="280px">
                            <Fields>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼" />
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期" />
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱" />
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱" />
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目" />
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者" />
                                <asp:CommandField ButtonType="Button" InsertText="新增" ShowCancelButton="False" ShowInsertButton="True">
                                    <ItemStyle ForeColor="Red" />
                                </asp:CommandField>
                            </Fields>
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <CommandRowStyle BackColor="#D1DDF1" Font-Bold="True" />
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FieldHeaderStyle BackColor="#DEE8F5" Font-Bold="True" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" />
                            <AlternatingRowStyle BackColor="White" />
                            <InsertRowStyle Wrap="False" />
                        </asp:DetailsView>
                        <asp:DetailsView ID="DetailsView3" runat="server" AutoGenerateRows="False" CellPadding="4"
                            DataKeyNames="報價單號碼" DataSourceID="AccessDataSource3" DefaultMode="Insert" ForeColor="#333333"
                            GridLines="None" Height="50px" Width="280px">
                            <Fields>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼" />
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期" />
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱" />
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱" />
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目" />
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者" />
                                <asp:CommandField ButtonType="Button" InsertText="新增" ShowCancelButton="False" ShowInsertButton="True">
                                    <ItemStyle ForeColor="Red" />
                                </asp:CommandField>
                            </Fields>
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <CommandRowStyle BackColor="#D1DDF1" Font-Bold="True" />
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FieldHeaderStyle BackColor="#DEE8F5" Font-Bold="True" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" />
                            <AlternatingRowStyle BackColor="White" />
                            <InsertRowStyle Wrap="False" />
                        </asp:DetailsView>
                        <asp:DetailsView ID="DetailsView4" runat="server" AutoGenerateRows="False" CellPadding="4"
                            DataKeyNames="報價單號碼" DataSourceID="AccessDataSource4" DefaultMode="Insert" ForeColor="#333333"
                            GridLines="None" Height="50px" Width="280px">
                            <Fields>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼" />
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期" />
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱" />
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱" />
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目" />
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者" />
                                <asp:CommandField ButtonType="Button" InsertText="新增" ShowCancelButton="False" ShowInsertButton="True">
                                    <ItemStyle ForeColor="Red" />
                                </asp:CommandField>
                            </Fields>
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <CommandRowStyle BackColor="#D1DDF1" Font-Bold="True" />
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FieldHeaderStyle BackColor="#DEE8F5" Font-Bold="True" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#2461BF" />
                            <AlternatingRowStyle BackColor="White" />
                            <InsertRowStyle Wrap="False" />
                        </asp:DetailsView>
                        <asp:DetailsView ID="DetailsView21" runat="server" AutoGenerateRows="False" CellPadding="4"
                            DataKeyNames="報價單號碼" DataSourceID="AccessDataSource5" DefaultMode="Insert" ForeColor="#333333"
                            GridLines="None" Height="50px" Width="280px">
                            <Fields>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼" />
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期" />
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱" />
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱" />
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目" />
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者" />
                                <asp:CommandField ButtonType="Button" InsertText="新增" ShowCancelButton="False" ShowInsertButton="True">
                                    <ItemStyle ForeColor="Red" />
                                </asp:CommandField>
                            </Fields>
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <CommandRowStyle BackColor="#D1DDF1" Font-Bold="True" />
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FieldHeaderStyle BackColor="#DEE8F5" Font-Bold="True" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#2461BF" />
                            <AlternatingRowStyle BackColor="White" />
                            <InsertRowStyle Wrap="False" />
                        </asp:DetailsView>
                        <asp:Label ID="Label4" runat="server" Font-Bold="True" Font-Size="Small" ForeColor="Red"
                            Text="請注意登錄者名字大小寫!"></asp:Label></td>
                </tr>
                <tr>
                    <td colspan="2">
                        &nbsp;<asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
                            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="報價單號碼" DataSourceID="AccessDataSource1"
                            ForeColor="#333333" GridLines="None" Width="800px">
                            <Columns>
                                <asp:CommandField CancelText="跳回" EditText="修改" ShowEditButton="True" UpdateText="確定">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" ForeColor="Red" Wrap="False" />
                                </asp:CommandField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="LinkButton5" runat="server" CommandName="Delete" Font-Bold="False"
                                            Font-Size="Small" ForeColor="Red" OnClientClick="return confirm('確認刪除?');">刪除</asp:LinkButton>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="False" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="100px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                            </Columns>
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" Font-Size="Small" ForeColor="#333333"
                                Wrap="False" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" Font-Size="Small" ForeColor="White"
                                Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" Font-Size="Small" Wrap="False" />
                            <AlternatingRowStyle BackColor="White" />
                        </asp:GridView>
                        <asp:GridView ID="GridView2" runat="server" AllowPaging="True" AllowSorting="True"
                            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="報價單號碼" DataSourceID="AccessDataSource2"
                            ForeColor="#333333" GridLines="None" Width="800px">
                            <Columns>
                                <asp:CommandField CancelText="跳回" EditText="修改" ShowEditButton="True" UpdateText="確定">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" ForeColor="Red" Wrap="False" />
                                </asp:CommandField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="LinkButton4" runat="server" CommandName="Delete" Font-Size="Small"
                                            ForeColor="Red" OnClientClick="return confirm('確認刪除?');">刪除</asp:LinkButton>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="False" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="100px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                            </Columns>
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" Font-Size="Small" ForeColor="White"
                                Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" Font-Size="Small" Wrap="False" />
                            <AlternatingRowStyle BackColor="White" />
                        </asp:GridView>
                        <asp:GridView ID="GridView3" runat="server" AllowPaging="True" AllowSorting="True"
                            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="報價單號碼" DataSourceID="AccessDataSource3"
                            ForeColor="#333333" GridLines="None" Width="800px">
                            <Columns>
                                <asp:CommandField CancelText="跳回" EditText="修改" ShowEditButton="True" UpdateText="確定">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" ForeColor="Red" Wrap="False" />
                                </asp:CommandField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="LinkButton3" runat="server" CommandName="Delete" Font-Size="Small"
                                            ForeColor="Red" OnClientClick="return confirm('確認刪除?');">刪除</asp:LinkButton>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="False" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="100px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                            </Columns>
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" Font-Size="Small" ForeColor="White"
                                Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" Font-Size="Small" Wrap="False" />
                            <AlternatingRowStyle BackColor="White" />
                        </asp:GridView>
                        <asp:GridView ID="GridView4" runat="server" AllowPaging="True" AllowSorting="True"
                            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="報價單號碼" DataSourceID="AccessDataSource4"
                            ForeColor="#333333" GridLines="None" Width="800px">
                            <Columns>
                                <asp:CommandField CancelText="跳回" EditText="修改" ShowEditButton="True" UpdateText="確定">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" ForeColor="Red" Wrap="False" />
                                    <ControlStyle Font-Size="Small" />
                                </asp:CommandField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="LinkButton1" runat="server" CommandName="Delete" Font-Bold="False"
                                            Font-Size="Small" ForeColor="Red" OnClientClick="return confirm('確認刪除?');">刪除</asp:LinkButton>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="False" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="100px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                            </Columns>
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" Font-Size="Small" ForeColor="#333333"
                                Wrap="False" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" Font-Size="Small" ForeColor="White"
                                Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" Font-Size="Small" Wrap="False" />
                            <AlternatingRowStyle BackColor="White" Font-Size="Small" Wrap="False" />
                        </asp:GridView>
                        <asp:GridView ID="GridView21" runat="server" AllowPaging="True" AllowSorting="True"
                            AutoGenerateColumns="False" CellPadding="4" DataKeyNames="報價單號碼" DataSourceID="AccessDataSource5"
                            ForeColor="#333333" GridLines="None" Width="800px">
                            <Columns>
                                <asp:CommandField CancelText="跳回" EditText="修改" ShowEditButton="True" UpdateText="確定">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" ForeColor="Red" Wrap="False" />
                                </asp:CommandField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="LinkButton2" runat="server" CommandName="Delete" Font-Size="Small"
                                            ForeColor="Red" OnClientClick="return confirm('確認刪除?');">刪除</asp:LinkButton>
                                    </ItemTemplate>
                                    <ItemStyle Wrap="False" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="報價單號碼" HeaderText="報價單號碼" ReadOnly="True" SortExpression="報價單號碼">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="日期" HeaderText="日期" SortExpression="日期">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="100px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="客戶名稱" HeaderText="客戶名稱" SortExpression="客戶名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="供應商名稱" HeaderText="供應商名稱" SortExpression="供應商名稱">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="0px" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="項目" HeaderText="項目" SortExpression="項目">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Wrap="False" />
                                </asp:BoundField>
                                <asp:BoundField DataField="登錄者" HeaderText="登錄者" SortExpression="登錄者">
                                    <HeaderStyle Font-Size="Small" Wrap="False" />
                                    <ItemStyle Font-Size="Small" Width="50px" Wrap="False" />
                                </asp:BoundField>
                            </Columns>
                            <RowStyle BackColor="#EFF3FB" Font-Size="Small" Wrap="False" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" Font-Size="Small" ForeColor="White"
                                Wrap="False" />
                            <EditRowStyle BackColor="#2461BF" Font-Size="Small" Wrap="False" />
                            <AlternatingRowStyle BackColor="White" />
                        </asp:GridView>
                    </td>
                </tr>
                </table>
        </div>
        <div>
            <div style="text-align: left">
                <asp:AccessDataSource ID="AccessDataSourceID" runat="server" DataFile="~/App_Data/CEMCL_ID.mdb"
                    SelectCommand="SELECT * FROM [ID]"></asp:AccessDataSource>
                <asp:AccessDataSource ID="AccessDataSource1" runat="server" DataFile="~/App_Data/CEMCL.mdb"
                    DeleteCommand="DELETE FROM [QUOTATION] WHERE [報價單號碼] = ?" InsertCommand="INSERT INTO [QUOTATION] ([報價單號碼], [日期], [客戶名稱], [供應商名稱], [項目], [登錄者]) VALUES (?, ?, ?, ?, ?, ?)"
                    SelectCommand="SELECT * FROM [QUOTATION] ORDER BY [報價單號碼] DESC" 
                    UpdateCommand="UPDATE [QUOTATION] SET [日期] = ?, [客戶名稱] = ?, [供應商名稱] = ?, [項目] = ?, [登錄者] = ? WHERE [報價單號碼] = ?">
                    <DeleteParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </DeleteParameters>
                    <UpdateParameters>
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </UpdateParameters>
                    <InsertParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                    </InsertParameters>
                </asp:AccessDataSource>
                <asp:AccessDataSource ID="AccessDataSource2" runat="server" DataFile="~/App_Data/CEMCL.mdb"
                    DeleteCommand="DELETE FROM [QUOTATION] WHERE [報價單號碼] = ?" InsertCommand="INSERT INTO [QUOTATION] ([報價單號碼], [日期], [客戶名稱], [供應商名稱], [項目], [登錄者]) VALUES (?, ?, ?, ?, ?, ?)"
                    SelectCommand="SELECT 報價單號碼, 日期, 客戶名稱, 供應商名稱, 項目, 登錄者 FROM QUOTATION WHERE (登錄者 IN ('Daniel', 'Nova'))
 ORDER BY [報價單號碼] DESC"
                    
                    UpdateCommand="UPDATE [QUOTATION] SET [日期] = ?, [客戶名稱] = ?, [供應商名稱] = ?, [項目] = ?, [登錄者] = ? WHERE [報價單號碼] = ?">
                    <DeleteParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </DeleteParameters>
                    <UpdateParameters>
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </UpdateParameters>
                    <InsertParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                    </InsertParameters>
                </asp:AccessDataSource>
                <asp:AccessDataSource ID="AccessDataSource3" runat="server" DataFile="~/App_Data/CEMCL.mdb"
                    DeleteCommand="DELETE FROM [QUOTATION] WHERE [報價單號碼] = ?" InsertCommand="INSERT INTO [QUOTATION] ([報價單號碼], [日期], [客戶名稱], [供應商名稱], [項目], [登錄者]) VALUES (?, ?, ?, ?, ?, ?)"
                    SelectCommand="SELECT * FROM [QUOTATION] WHERE ([登錄者] LIKE '%' + ? + '%')
ORDER BY [報價單號碼] DESC" 
                    UpdateCommand="UPDATE [QUOTATION] SET [日期] = ?, [客戶名稱] = ?, [供應商名稱] = ?, [項目] = ?, [登錄者] = ? WHERE [報價單號碼] = ?">
                    <SelectParameters>
                        <asp:Parameter DefaultValue="%Maxwell%" Name="登錄者2" Type="String" />
                    </SelectParameters>
                    <DeleteParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </DeleteParameters>
                    <UpdateParameters>
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </UpdateParameters>
                    <InsertParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                    </InsertParameters>
                </asp:AccessDataSource>
                <asp:AccessDataSource ID="AccessDataSource4" runat="server" DataFile="~/App_Data/CEMCL.mdb"
                    DeleteCommand="DELETE FROM [QUOTATION] WHERE [報價單號碼] = ?" InsertCommand="INSERT INTO [QUOTATION] ([報價單號碼], [日期], [客戶名稱], [供應商名稱], [項目], [登錄者]) VALUES (?, ?, ?, ?, ?, ?)"
                    OldValuesParameterFormatString="original_{0}" SelectCommand="SELECT * FROM [QUOTATION] WHERE ([登錄者] LIKE '%' + ? + '%')
ORDER BY [報價單號碼] DESC"
                    
                    UpdateCommand="UPDATE [QUOTATION] SET [日期] = ?, [客戶名稱] = ?, [供應商名稱] = ?, [項目] = ?, [登錄者] = ? WHERE [報價單號碼] = ?">
                    <SelectParameters>
                        <asp:Parameter DefaultValue="%Jash%" Name="登錄者2" Type="String" />
                    </SelectParameters>
                    <DeleteParameters>
                        <asp:Parameter Name="original_報價單號碼" Type="Int32" />
                    </DeleteParameters>
                    <UpdateParameters>
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                        <asp:Parameter Name="original_報價單號碼" Type="Int32" />
                    </UpdateParameters>
                    <InsertParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                    </InsertParameters>
                </asp:AccessDataSource>
                <asp:AccessDataSource ID="AccessDataSource5" runat="server" DataFile="~/App_Data/CEMCL.mdb"
                    DeleteCommand="DELETE FROM [QUOTATION] WHERE [報價單號碼] = ?" InsertCommand="INSERT INTO [QUOTATION] ([報價單號碼], [日期], [客戶名稱], [供應商名稱], [項目], [登錄者]) VALUES (?, ?, ?, ?, ?, ?)"
                    SelectCommand="SELECT * FROM [QUOTATION] WHERE ([登錄者] LIKE '%' + ? + '%')
ORDER BY [報價單號碼] DESC" 
                    UpdateCommand="UPDATE [QUOTATION] SET [日期] = ?, [客戶名稱] = ?, [供應商名稱] = ?, [項目] = ?, [登錄者] = ? WHERE [報價單號碼] = ?">
                    <SelectParameters>
                        <asp:Parameter DefaultValue="%Nova%" Name="登錄者2" Type="String" />
                    </SelectParameters>
                    <DeleteParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </DeleteParameters>
                    <UpdateParameters>
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                    </UpdateParameters>
                    <InsertParameters>
                        <asp:Parameter Name="報價單號碼" Type="Int32" />
                        <asp:Parameter Name="日期" Type="String" />
                        <asp:Parameter Name="客戶名稱" Type="String" />
                        <asp:Parameter Name="供應商名稱" Type="String" />
                        <asp:Parameter Name="項目" Type="String" />
                        <asp:Parameter Name="登錄者" Type="String" />
                    </InsertParameters>
                </asp:AccessDataSource>
            </div>
        </div>
    </form>
</body>
</html>
