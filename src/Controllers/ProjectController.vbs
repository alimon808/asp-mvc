Class ProjectController
	Dim Model
	Dim ViewData

	private sub Class_Initialize()
		Set ViewData = Server.CreateObject("Scripting.Dictionary")
	end sub

	private sub Class_Terminate()
	end sub

	public Sub List()
		Dim u : Set u = new ProjectHelper
		Call AddViewData("model", u.SelectAll)
		View("list")
	End Sub

	public Sub Create()
		Call AddViewData("model", new Project)
		View("create")
	End Sub

	public Sub CreatePost(args)
		Dim objh : set objh = new ProjectHelper
		Dim obj : set obj = new Project
		obj.ProjectName = args("ProjectName")
		obj.Active = (args("Active") = "on")
		obj.POP3Address = args("POP3Address")
		'form values should be cleaned from injections
		'checkboxes shoud use the syntax: obj.booleanProperty = (args("checkboxname") = "on")
		result = objh.Insert(obj)
		Response.Redirect("?controller=Project&action=list")
	End Sub

	public Sub Edit(vars)
		Dim u : Set u = new ProjectHelper
		Call AddViewData("model", u.SelectById(vars("id")))
		View("edit")
	End Sub

	public Sub EditPost(args)
		Dim objh : set objh = new ProjectHelper
		Dim obj : set obj = objh.SelectById(args("id"))
		obj.ProjectName = args("ProjectName")
		obj.Active = (args("Active") = "on")
		obj.POP3Address = args("POP3Address")

		'form values should be cleaned from injections
		'checkboxes shoud use the syntax: obj.booleanProperty = (args("checkboxname") = "on")
		objh.Update(obj)
		Response.Redirect("?controller=Project&action=list")
	End Sub

	public Sub Delete(vars)
		Dim u : Set u = new ProjectHelper
		Call AddViewData("model", u.SelectById(vars("id")))
		View("delete")
	End Sub

	public Sub DeletePost(args)
		Dim objh : set objh = new ProjectHelper
		Call objh.Delete(args("id"))
		Response.Redirect("?controller=Project&action=list")
	End Sub

	public Sub Details(vars)
		Dim u : Set u = new ProjectHelper
		Call AddViewData("model", u.SelectById(vars("id")))
		View("details")
	End Sub

	Private Function AddViewData(key, val)
		ViewData.Add key, val
	End Function
	
	Private Sub View(action)
		Dim vw : vw = "/aspmvc/views/project/" & action & ".asp"
		Set Session("viewData") = ViewData
		Server.Transfer(vw) 
	End Sub
End Class