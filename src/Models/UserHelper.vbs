Class UserHelper
  Dim selectSQL

  private sub Class_Initialize()
	selectSQL = " SELECT * FROM [User] "
  end sub

  private sub Class_Terminate()
  end sub

 '=============================
 'public Functions



  ' Create a new User. 
  ' true - if successful, false otherwise
  public function Insert(obj)
	  Dim strSQL 
	  strSQL=   " Insert into [User] ( FirstName, LastName, UserName, ProjectID) Values ('{0}','{1}','{2}',{3}) " & _
				" SELECT SCOPE_IDENTITY();  " 
	  obj.Id = DbExecute (StringFormat(strSQL, array( obj.FirstName, obj.LastName, obj.UserName, obj.ProjectID)))
	  Insert = true
  end function

  ' Update the User
  public function Update(obj)
	 Dim strSQL 
	 strSQL= "Update [User] set  FirstName = '{0}', LastName = '{1}', UserName = '{2}', ProjectID = {3}  Where Id = {4} " 
	 DbExecute StringFormat(strSQL, array( obj.FirstName, obj.LastName, obj.UserName, obj.ProjectID, obj.Id))
	 Update = tru
  end function
  
  ' Delete the User
  public function Delete(Id)
	 Dim strSQL 
	 strSQL= "DELETE FROM [User] WHERE Id = {0}"
	 DbExecute StringFormat(strSQL, array(Id))
	 Delete = true
  end function
  
  ' Select the User by ID
  ' return User object - if successful, Nothing otherwise
  public function SelectById(id)
	Dim record
	set record = DbExecute(StringFormat(selectSQL + " WHERE Id = {0}", array(id)))
	Set SelectById = PopulateObjectFromRecord(record)
	record.Close
	set record = nothing
  end function
  
  ' Select all Users into a Dictionary
  ' return a Dictionary of User objects - if successful, Nothing otherwise
  public function SelectAll()
	Dim records
	set records = DbExecute(selectSQL)
	if records.eof then 
		Set SelectAll = Nothing
	else
		Dim results, obj, record
		Set results = Server.CreateObject("Scripting.Dictionary")
		while not records.eof
			set obj = PopulateObjectFromRecord(records)
			results.Add obj.Id, obj
			records.movenext 
		wend 
		set SelectAll = results
		records.Close
	End If
	set record = nothing
  end function
 
  ' Select all Users into a Dictionary
  ' return a Dictionary of User objects - if successful, Nothing otherwise
  public function SelectByField(fieldName, value)
	Dim records, txtSQL
	txtSQL = StringFormat(selectSQL + " where " + fieldName + "={0}" , array(value))
	set records = DbExecute(txtSQL)
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
	set record = nothing
  end function

  private function PopulateObjectFromRecord(record)
	if record.eof then
	   Set PopulateObjectFromRecord = Nothing
	else
		Dim obj
		set obj = new User
		obj.Id                       = record("Id") 
		obj.FirstName = record("FirstName")  
		obj.LastName = record("LastName")  
		obj.UserName = record("UserName")  
		obj.ProjectID = record("ProjectID") 
		set PopulateObjectFromRecord = obj
	end if
  end function

  Private Function DbExecute(sql)
	Dim db : Set db = new Database
	Set DbExecute = db.Execute(sql)
	Set db = Nothing
  End Function
end class
