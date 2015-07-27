List Developers
<br />
<%=Html.ActionLink("Index", "Home", "Index" , "") %>
<br />
<%=Html.ActionLink("About", "Home", "About" , "") %>
<br />
<%=Html.ActionLink("Abandon session", "Home", "AbandonSession" , "") %>
<br />
<br />
<%=Html.ActionLink("Create new Developer", "Developer", "Create" , "") %> <br/>

<table>
    <tr>

        <td>FirstName</td>

        <td>LastName</td>

        <td></td>
    </tr>

    <%
    if  IsNothing(Model) then
        %> <tr><td colspan="4">No records</td> </tr><%
    Else
        Dim obj
        For each obj in Model.Items
        %>

        <tr>

            <td><%=Html.Encode(obj.FirstName) %></td>

            <td><%=Html.Encode(obj.LastName) %></td>

            <td>
                <%=Html.ActionLink("Edit", "Developer", "Edit" , "id=" + CStr(obj.Id)) %> |
                <%=Html.ActionLink("Delete", "Developer", "Delete" , "id=" + CStr(obj.Id)) %> |
                <%=Html.ActionLink("Details", "Developer", "Details" , "id=" + CStr(obj.Id)) %>
                
            </td>
        </tr>
        <% 
        Next
     End If
     %>
</table>

