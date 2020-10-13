<%@ Page Language = "VB" %>
<%@ Import Namespace = "System.Data.OleDb" %>
<!DOCTYPE html>
<html xmlns = "http://www.w3.org/1999/xhtml">
<head id = "Head1" runat = "server">
<title>Connection</title>
<script runat = "server">
    Sub Create_Click(Src As Object, E As EventArgs)
        Try
            'Connect to the Database
            Dim cnAccess As New OleDbConnection(
            "Provider = Microsoft.Jet.OLEDB.4.0;" &
            "Data Source = C:\Users\yatid\Documents\HigginsHotelSystem.mdb")
            Dim sSelectSQL As String = "CREATE TABLE Guests"
            sSelectSQL &= "([GuestID] Number, [LName] TEXT(20),"
            sSelectSQL &= "[FName] TEXT(20), [ZipCode] Number,"
            sSelectSQL &= "[StateID] TEXT(20),[CreditCardNo] TEXT(20))"

            Dim cmdSelect As New OleDbCommand(sSelectSQL, cnAccess)
            cnAccess.Open()
            cmdSelect.ExecuteNonQuery()
            cnAccess.Close()
            msg.Text = "Table Created!"

        Catch ex As Exception
            msg.Text = ex.Message
            ' Response.Write("Table Exists or Connection Failed")
        End Try
    End Sub

    Sub GoTo_Click(Src As Object, E As EventArgs)
        Response.Redirect("hotel2.aspx")
    End Sub
</script>
</head>
<body style = "font-family:Tahoma;">
<h3>Higgins Hotel Systems</h3>
<form runat = "server" id = "form1">
<asp:Button Text = "Create Table" OnClick = "Create_Click"
runat = "server" ID = "Button1" />
<p><asp:Label id = "msg" runat = "server" /></p>
<br />
<asp:Button Text = "Insert Records" OnClick = "GoTo_Click"
runat = "server" ID = "Button2" />
</form>
</body>
</html>

