<%@ Page Language = "VB" %>
<%@ Import Namespace = "System.Data.OleDb" %>
<!DOCTYPE html>
<html xmlns = "http://www.w3.org/1999/xhtml">
<head id="Head1" runat = "server">
<title>Connection</title>
<script runat = "server">
    Sub Insert_Click(Src As Object, E As EventArgs)
        Try
            'Connect to the Database
            Dim cnAccess As New OleDbConnection(
            "Provider = Microsoft.Jet.OLEDB.4.0;" &
             "Data Source = C:\Users\yatid\Documents\HigginsHotelSystem.mdb")


            cnAccess.Open()
            Dim sID, sFName, sLName, sZip, sState, sCreditCardNo, sInsertSQL As String
            sID = GuestID.Text
            sFName = FName.Text
            sLName = LName.Text
            sZip = Zip.Text
            sState = State.Text
            sCreditCardNo = CreditCardNo.Text


            'Construct the insert statement
            sInsertSQL = "INSERT INTO Guests(" &
" [GuestID], [LName], [FName], [ZipCode], [StateID], [CreditCardNo]) VALUES" &
" (" & sID & ",'" & sLName & "','" & sFName & "'," & sZip & ",'" & sState & "','" & sCreditCardNo & "');"


            'Construct the OleDbCommand object
            Dim cmdInsert As New OleDbCommand(sInsertSQL, cnAccess)



            'since this is not a query, we do not expect to return data 
            cmdInsert.ExecuteNonQuery()



            Response.Write("Data Recorded!")
        Catch ex As Exception
            Response.Write(ex.Message)
            Response.Write("Connection Failed")
        End Try



    End Sub


    Sub GoTo_Click(Src As Object, E As EventArgs)
        Response.Redirect("hotel3.aspx")
    End Sub
</script>
</head>
<body style = "font-family:Tahoma;">
<h3>Enter Guest Details</h3>
<form runat = "server" id = "form1">

<table>
<tr>
<td>ID</td>
<td><asp:Textbox id = "GuestID" runat="server" /></td>
</tr>
<tr>
<td>Last Name: </td>
<td><asp:Textbox id = "LName" runat = "server" /></td>
</tr>
<tr>
<td>First Name: </td>
<td><asp:Textbox id = "FName" runat = "server" /></td>
</tr>
<tr>
 <td>ZipCode</td>
                <td>
                    <asp:TextBox ID="Zip" runat="server" />
                </td>

                <td>
                    <div>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator2"
                            ControlToValidate="Zip"
                            ValidationExpression="[0-9]{5}"
                            Display="dynamic"
                            ErrorMessage="ZipCode - Must be 5 digits"
                            runat="server" />

                    </div>
                </td>
</tr>
<tr>
<td>StateID: </td>
<td><asp:Textbox id = "State" runat = "server" /></td>
</tr>
<tr>
    <td>CreditCardNo</td>
                <td>
                    <asp:TextBox ID="CreditCardNo" runat="server" />
                </td>
<td>
                    <div>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator1"
                            ControlToValidate="CreditCardNo"
                            ValidationExpression="[0-9]{16}"
                            Display="dynamic"
                            ErrorMessage="CreditCardNo - Must be 16 digits"
                            runat="server" />

                    </div>
                </td>
</tr>
</table>
<br />
<asp:Button Text = "Insert" OnClick = "Insert_Click"
runat = "server" ID = "Button1" />
<p>
<asp:Label id = "msg" runat = "server" />
</p>
<br />
<asp:Button Text = "Retrieve Records" OnClick = "GoTo_Click"
runat = "server" ID = "Button2" />
</form>

<div></div>
</body>
</html>