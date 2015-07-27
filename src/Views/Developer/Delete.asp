Delete Developer
<br />
<%=Html.ActionLink("Index", "Home", "Index" , "") %>
<br />
<%=Html.ActionLink("About", "Home", "About" , "") %>
<br />
<%=Html.ActionLink("Abandon session", "Home", "AbandonSession" , "") %>
<br />
<form action="?controller=Developer&action=DeletePost" id="DeletePost" method="post">
<table>
    <tr>
        <td>id</td>
        <td>
            <%=Html.Encode(Model.Id)%>
            <%=Html.Hidden("id",Model.Id) %>
        </td>
    </tr>

    <tr>
        <td>FirstName</td>
        <td><%=Html.Encode(Model.FirstName)%></td>
    </tr>

    <tr>
        <td>LastName</td>
        <td><%=Html.Encode(Model.LastName)%></td>
    </tr>

</table>
<button type="submit">Submit</button>
<%=Html.ActionLink("Back to list", "Developer", "List" , "") %> <br/>
</form>

