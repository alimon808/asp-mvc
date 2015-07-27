<%
class cHTMLHelper

	Public Function Encode (elValue)
	    if Not IsNothing(elValue) Then
			Encode = Server.HTMLEncode(elValue)
		End If
	End Function 

	Public Function Label (elValue, objfor)
		Label = "<label for='" + objfor + "'>" + Encode(elValue) + "</label>"
	End Function
	
	Public Function Hidden ( elID, elValue)
	    Hidden =  "<input id='" + elID + "' name='" + elID + "' type='hidden' value='" + Encode(elValue) + "' />"
	End Function

	Public Function TextBox (elId, elValue)
		TextBox = "<input id='" + elID + "' name='" + elID + "' type='text' value='" + Encode(elValue) + "' />"
	End Function
	
	Public Function TextArea (elId, elValue, cols, rows)
		TextArea = "<textarea id='" + elID + "' name='" + elID + "' cols='" + Encode(cols) + "' rows='" + Encode(rows) + "'>" & _
		 Encode(elValue) & _
		 "</textarea>"
	End Function

	Public Function DropDownList (elId, elValue, list , idName, valueName)
		Dim resStr, lisItem
		resStr= "<select id='" + elID + "' name='" + elID + "'>" 
		For Each listItem in List
		    Dim optValue, optText
		    optValue = Eval("listItem." + idName)
		    optText  = Eval("listItem." + valueName)
			if elValue = optValue Then
				resStr = resStr + "<option selected='selected' value='" + Encode(optValue) + "'>" + Encode(optText) + "</option>"
			else
				resStr = resStr + "<option value='" + Encode(optValue) + "'>" + Encode(optText) + "</option>"
			End If
		Next
		resStr = resStr & "</select>"
		DropDownList = resStr
	End Function
    
    Public Function ActionLink(linkText, linkController, linkAction , linkVars) 
		'Stop
		Dim strVars
      
		If IsEmpty(linkVars) or IsNull(linkVars) then
			strVars = ""
		else
			if IsArray(linkVars) then
				for each varPair in linkVars
					strVars  = strVars + "&" + varPair
				next
			else
				strVars = linkVars
			end if
		end if
      
		if (strVars<>"") then 
			strVars = "&" + strVars
		End if
   
		ActionLink = "<a href='?controller=" + Encode(linkController)+ "&action=" + Encode(linkAction)  +   Encode(strVars) + "'>" + Encode(linkText) + "</a>"
	End Function

	Public Function CheckBox( elId, elValue)
	    Dim checked 
	    if (elValue = 1) or (elValue = true) or (LCase(elValue) = "true") Then
			checked = "CHECKED"
		else
			checked = ""
		end If
		TextArea = "<input type='checkbox' id='" + elID + "' name='" + elID + "' " + checked + " />" 
		Encode(elValue) 
	End Function

	Public Sub RenderControllerAction(controller,action,vars)
		Dim controllerName
        controllerName = controller + "Controller"
        Dim controllerInstance 
        Set controllerInstance = Eval ( " new " +  controllerName)
        Dim actionCallString 
		If Not (IsNothing(vars)) then
                actionCallString = " controllerInstance." + action + "(vars)"
            Else
                actionCallString = " controllerInstance." + action + "()"
            End If
            Eval (actionCallString)
	End Sub

    
end class

public Html
set Html = new cHTMLHelper
%>