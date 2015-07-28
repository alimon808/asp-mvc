class UserController
 Dim Model
 Dim ViewData
  
 private sub Class_Initialize()
  Set ViewData = Server.CreateObject("Scripting.Dictionary")
 end sub

 private sub Class_Terminate()
 end sub

 public Sub List()
    Dim u : set u = new UserHelper
    ViewData.Add "model", u.SelectAll
    View("list")
 End Sub
 
 public Sub Create()
    ViewData.Add "model", new User
    View("create")
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
    Dim u : set u = new UserHelper
    ViewData.Add "model", u.SelectById(vars("id"))
    View("edit")
 End Sub
 
 public Sub EditPost(args)
    Dim objh : set objh = new UserHelper
    Dim obj : set obj = objh.SelectById(args("id"))
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
    Dim u : set u = new UserHelper
    ViewData.Add "model", u.SelectById(vars("id"))
    View("delete")
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
        Dim u : set u = new UserHelper
        ViewData.Add "model", u.SelectById(vars("id"))
        View("details")
    End Sub

    Private Sub View(action)
        Dim vw : vw = "/views/user/" & action & ".asp"
        Set Session("viewData") = ViewData
        Server.Transfer(vw)
    End Sub 
 End Class
