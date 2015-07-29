class ProjectHelper
	Dim selectSQL

	Sub Class_Initialize()
		selectSQL = " SELECT * FROM [Project] "
	End Sub

	Sub Class_Terminate()
	End Sub

	Public Function Insert(ByRef obj)
		Dim sql
		sql = " Insert into [Project] ( ProjectName , Active , POP3Address)" 
		sql = sql & " Values (?  , ?  , ? ); " 
		sql = sql & " SELECT SCOPE_IDENTITY()  "
		Dim result : Set result = DbExecute(sql, Array( obj.ProjectName ,  obj.Active ,  obj.POP3Address ))
		obj.Id = CInt(result(0))
	End  Function

	' Update the Project
	Public function Update(obj)
		Dim strSQL : strSQL= "Update [Project] set ProjectName=?  , Active=?  , POP3Address=?  Where Id = ? " 
		Call DbExecute(strSQL, Array(obj.ProjectName, obj.Active, obj.POP3Address,  obj.Id))
	End Function

	Public function Delete(Id)
		Dim sql : sql= "DELETE FROM [Project] WHERE Id = ?"
		Dim result : Set result = DbExecute(sql, Array(id))
	End Function

		' Select the Project by ID
		' return Project object - if successful, Nothing otherwise
	Public function SelectById(id)
		Dim sql : sql = selectSQL & " WHERE id=?"
		Dim record : Set record = DbExecute(sql, Array(id))
		Set SelectById = PopulateObjectFromRecord(record)
		record.Close
		set record = nothing
	End Function

	Public function SelectAll()
		Dim records : set records = DbExecute(selectSQL, empty)
		If records.eof Then
			Set SelectAll = Nothing
		Else
			Dim obj, record
			Dim results : Set results = Server.CreateObject("Scripting.Dictionary")
			While Not records.eof
				Set obj = PopulateObjectFromRecord(records)
				results.Add obj.Id, obj
				records.movenext
			Wend
			Set SelectAll = results
			records.Close
		End If
		Set records = Nothing
	End Function

	' Select all Projects into a Dictionary
	' return a Dictionary of Project objects - if successful, Nothing otherwise
	Public function SelectByField(fieldName, value)
		Dim records
		set objCommand=Server.CreateObject("ADODB.command")
		objCommand.ActiveConnection=DbOpenConnection()
		objCommand.NamedParameters = False
		objCommand.CommandText = selectSQL + " where " + fieldName + "=?"
		objCommand.CommandType = adCmdText
		If DbAddParameters(objCommand, array(value)) Then
			set records = objCommand.Execute
			if records.eof then
				Set SelectByField = Nothing
			else
				Dim results, obj, record
				Set results = Server.CreateObject("Scripting.Dictionary")
				while not records.eof
					set obj = PopulateObjectFromRecord(records)
					results.Add obj.Id, obj
					records.movenext
				wend
				set SelectByField = results
				records.Close
			End If
			set records = nothing
		Else
			set SelectByField = Nothing
		End If
	End Function

	private function PopulateObjectFromRecord(record)
		if record.eof then
		else
			Dim obj
			set obj = new Project
			obj.Id = record("Id")

			obj.ProjectName = record("ProjectName")
			obj.Active = record("Active")
			obj.POP3Address = record("POP3Address")
			set PopulateObjectFromRecord = obj
		end if
	End Function

	Private Function DbAddParameters(byref objCommand, values)
		DbAddParameters = False
	End Function

	Private Function DbExecute(sql, values)
		Dim db : Set db = new Database
		If IsArray(values) Then
			Set DbExecute = db.ExecuteWithParams(sql, values)
		Else
			Set DbExecute = db.Execute(sql)
		End If
		Set db = Nothing
	End Function
end class