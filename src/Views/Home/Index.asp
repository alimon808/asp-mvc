<!-- #include virtual="/views/view.asp" -->
<h1>Classic ASP using MVC Pattern</h1>
This is an INDEX page
<br />
<%=Html.ActionLink("Index", "Home", "Index" , "") %>
<br />
<%=Html.ActionLink("About", "Home", "About" , "") %>
<br />
<%=Html.ActionLink("Abandon session", "Home", "AbandonSession" , "") %>
<br />
<br />
Hello
<%=Html.Encode(viewData.Item("name"))%>