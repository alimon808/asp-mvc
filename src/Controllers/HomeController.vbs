class HomeController
 Dim Model
 Dim ViewData
   
 private sub Class_Initialize()
    Set ViewData = Server.CreateObject("Scripting.Dictionary")
 end sub

 private sub Class_Terminate()
 end sub
 
 public Sub Index()
    ViewData.Add "name", "dude"
    View("index")
 End Sub
 
 public Sub About()
    if Session("sessionCounter")="" then
       Session("sessionCounter") = 1
    Else
        Session("sessionCounter") = Session("sessionCounter") + 1
    End If
    ViewData.Add "sessionCounter", Session("sessionCounter")
    View("About")
 End Sub
 
 Public Sub AbandonSession()
   Session.Abandon()
   Response.Redirect("?controller=Home&action=About")
 End Sub

 Private Sub View(action)
  Dim vw : vw = "/aspmvc/views/home/" & action & ".asp"
  Set Session("viewData") = ViewData
  Server.Transfer(vw)
 End Sub
End Class