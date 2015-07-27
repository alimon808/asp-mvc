Edit Developer
<br />
<%=Html.ActionLink("Index", "Home", "Index" , "") %>
<br />
<%=Html.ActionLink("About", "Home", "About" , "") %>
<br />
<%=Html.ActionLink("Abandon session", "Home", "AbandonSession" , "") %>
<br />
<form action="?controller=Developer&action=EditPost" id="EditPost" method="post">
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
        <td> 
		<%=Html.TextBox("FirstName", Model.FirstName)%>
		<!--Html.TextArea("FirstName", Model.FirstName, 60, 15)-->
		<!--Html.TextBox("FirstName", Model.FirstName)-->
		<!--Html.CheckBox("FirstName", Model.FirstName)-->
		<!--Html.DropDownList("FirstName", Model.FirstName, new FirstNameHelper.SelectAll.Items, "Id", "FirstNameName")-->
		<!--String-->
	</td>
    </tr>

    <tr>
        <td>LastName</td>
        <td> 
		<%=Html.TextBox("LastName", Model.LastName)%>
		<!--Html.TextArea("LastName", Model.LastName, 60, 15)-->
		<!--Html.TextBox("LastName", Model.LastName)-->
		<!--Html.CheckBox("LastName", Model.LastName)-->
		<!--Html.DropDownList("LastName", Model.LastName, new LastNameHelper.SelectAll.Items, "Id", "LastNameName")-->
		<!--String-->
	</td>
    </tr>




</table>
<button type="submit">Submit</button>
<%=Html.ActionLink("Back to list", "Developer", "List" , "") %> <br/>

</form>

