This is an DETAILS page
<br />
<%=Html.ActionLink("Index", "Home", "Index" , "") %>
<br />
<%=Html.ActionLink("About", "Home", "About" , "") %>
<br />
<%=Html.ActionLink("Abandon session", "Home", "AbandonSession" , "") %>
<br />
<br />
<table>
    <tr>
        <td>
            ID:
        </td>
        <td>
            <%=Html.Encode(Model.Id) %>
        </td>
    </tr>
    <tr>
        <td>
            First Name:
        </td>
        <td>
            <%=Html.Encode(Model.FirstName) %>
        </td>
    </tr>
    <tr>
        <td>
            Last Name:
        </td>
        <td>
            ID:
        </td>
        <td>
            <%=Html.Encode(Model.LastName) %>
        </td>
    </tr>
</table>
