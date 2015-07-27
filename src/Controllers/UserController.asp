<% 

class UserController
 Dim Model
 Dim ViewData
  
 private sub Class_Initialize()
  Set ViewData = Server.CreateObject("Scripting.Dictionary")
 end sub

 private sub Class_Terminate()
 end sub

 public Sub List()
    Dim u 
    set u = new UserHelper
    set Model = u.SelectAll
 
    %>   <!--#include virtual="/views/User/List.asp" --> <%
 End Sub
 
 public Sub Create()
    set Model = new User
    %>   <!--#include virtual="/views/User/Create.asp" --> <%
 End Sub
 
  public Sub CreatePost(args)
    Dim obj, objh
    set objh = new UserHelper
    set obj = new User

    obj.FirstName = args("FirstName")
    obj.LastName = args("LastName")
    obj.UserName = args("UserName")
    obj.ProjectID = args("ProjectID")
    'form values should be cleaned from injections
    'checkboxes shoud use the syntax: obj.ProjectID = (args("ProjectID") = "on")
    obj.Id = objh.Insert(obj)
 
    Response.Redirect("?controller=User&action=list")
 End Sub

 
 public Sub Edit(vars)

    Dim u
    set u = new UserHelper
    set Model = u.SelectById(vars("id"))
 
   %>   <!--#include virtual="/views/User/Edit.asp" --> <%
 End Sub
 
 public Sub EditPost(args)
    Dim obj, objh
    set objh = new UserHelper
    set obj = objh.SelectById(args("id"))

    obj.FirstName = args("FirstName")
    obj.LastName = args("LastName")
    obj.UserName = args("UserName")
    obj.ProjectID = args("ProjectID")
    'form values should be cleaned from injections
    'checkboxes shoud use the syntax: obj.ProjectID = (args("ProjectID") = "on")
    objh.Update(obj)
 
    Response.Redirect("?controller=User&action=list")
 End Sub

 
 public Sub Delete(vars)
     Dim u
    set u = new UserHelper
    set Model = u.SelectById(vars("id"))
   %>   <!--#include virtual="/views/User/Delete.asp" --> <%
 End Sub
 
 
  public Sub DeletePost(args)
    Dim res, objh
    set objh = new UserHelper
    res = objh.Delete(args("id"))
    if  res then
        Response.Redirect("?controller=User&action=list")
    else
        Response.Redirect("?controller=User&action=Delete&id=" + CStr(args("id")))
    end if
 End Sub


 public Sub Details(vars)
     Dim u
    set u = new UserHelper
    set Model = u.SelectById(vars("id"))

   %>   <!--#include virtual="/views/User/Details.asp" --> <%
 End Sub
 
 End Class
%>
