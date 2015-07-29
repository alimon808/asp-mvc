Class UserController
	Dim ViewData

	Private Sub Class_Initialize()
		Set ViewData = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Sub List()
		Dim u : set u = new UserHelper
		ViewData.Add "model", u.SelectAll
		View("list")
	End Sub

	Public Sub Create()
		ViewData.Add "model", new User
		View("create")
	End Sub

	Public Sub CreatePost(args)
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
		Response.Redirect("?controller=User&action=List")
	End Sub

	Public Sub Edit(vars)
		Dim u : set u = new UserHelper
		ViewData.Add "model", u.SelectById(vars("id"))
		View("edit")
	End Sub

	Public Sub EditPost(args)
		Dim objh : set objh = new UserHelper
		Dim obj : set obj = objh.SelectById(args("id"))
		obj.FirstName = args("FirstName")
		obj.LastName = args("LastName")
		obj.UserName = args("UserName")
		obj.ProjectID = args("ProjectID")
		'form values should be cleaned from injections
		'checkboxes shoud use the syntax: obj.ProjectID = (args("ProjectID") = "on")
		objh.Update(obj)
		Response.Redirect("?controller=User&action=List")
	End Sub

	Public Sub Delete(vars)
		Dim u : set u = new UserHelper
		ViewData.Add "model", u.SelectById(vars("id"))
		View("delete")
	End Sub

	Public Sub DeletePost(args)
		Dim objh : set objh = new UserHelper
		Call objh.Delete(args("id"))
		Response.Redirect("?controller=User&action=list")
	End Sub

	Public Sub Details(vars)
		Dim u : set u = new UserHelper
		ViewData.Add "model", u.SelectById(vars("id"))
		View("details")
	End Sub

	Private Sub View(action)
		Dim vw : vw = "/aspmvc/views/user/" & action & ".asp"
		Set Session("viewData") = ViewData
		Server.Transfer(vw)
	End Sub 
End Class
