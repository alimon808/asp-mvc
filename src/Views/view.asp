<!--#include virtual="/utils/utils.asp" -->
<!--#include virtual="/models/models.asp" -->
<%
Dim viewData : Set viewData = Session("viewData")
Session.Contents.Remove("viewData")
Dim Model
If viewData.Exists("model") Then
	Set Model = viewData.Item("model")
Else
	Set Model = Nothing
End If
%>

