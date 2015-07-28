Class View
	Public Function Html(controller, action)
		Dim filename : filename = Server.MapPath("/views/" & controller & "/" & action & ".asp")
		Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
		Dim oFile : set oFile = fso.OpenTextFile(filename)
		sText = oFile.ReadAll
		response.write sText
		response.end
		oFile.close 
		ExecuteGlobal sText
	End Function
End Class