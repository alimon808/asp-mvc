const adCmdText =1
const adCmdStoredProc = 4
const adCmdUnknown = 8

Class Database
	Private mConnection
	Private mCommand

	Public Property Get Connection()
		OpenConnection()
		Set Connection = mConnection
	End Property

	Public Property Get Command()
		InitializeCommand()
		Set Command = mCommand
	End Property

	Sub Class_Initialize()
	End Sub

	Sub Class_Terminate()
	End Sub

	Public Function Execute(sql)
		InitializeCommand()
		mCommand.CommandText = sql
		Set Execute = mCommand.Execute
	End Function

	Public Function ExecuteWithParams(sql, values)
		InitializeCommand()
		mCommand.CommandText = sql
		AddParameters(values)
		Set ExecuteWithParams = mCommand.Execute
	End Function

	Public function AddParameters(values)
		If Not IsArray(values) Then
			Exit Function
		End If
		If mCommand.Parameters.Count = UBound (values)+1 Then
			For i=0 to mCommand.Parameters.Count -1
				mCommand.Parameters(i) = DBSafeValue  (mCommand.Parameters(i), values(i))
			Next
		End If
	End Function

	Private Function DBSafeValue(param, value)
		If Not TypeName(param) = "Parameter" Then
			Exit Function
		End If
		If IsNothing(value)or (value = "")  Then
			DBSafeValue = null
		Else
			Select Case param.Type
				Case 129,130,200,201,202,203:
					if (param.Size<>-1) and (Len(CStr(value))> param.Size) Then
						Err.Raise 8, "db utilites: DBSafeAssign ", "string is too long(" & value & ")"
					End If
					DBSafeValue = CStr(value)
				Case 72:
					DBSafeValue = CStr(value)
				Case  7,135:
					DBSafeValue = CDate(value)
				Case 20,3,131,2,17 :
					DBSafeValue = CLng(value)
				Case 4,5,14,6 :
					DBSafeValue = CDbl(value)
				Case 11:
					DBSafeValue = CBool(value)
				Case Else:
					Err.Raise 8, "db utilites: DBSafeAssign ", "unsupported type(" & param.Type & ") of database field"
		  	End Select
		End If
	End Function

	Private Sub OpenConnection()
		Dispose(mConnection)
		Set mConnection = Server.CreateObject("ADODB.Connection")
		mConnection.Mode = 3
		mConnection.ConnectionString = Application("connectionString")
		mConnection.open
	End Sub

	Private Sub CloseConnection()
		If Not connection Is Nothing Then
			connection.Close
			Set connection = Nothing
		End If
	End Sub

	Private Sub InitializeCommand()
		Set mCommand = Server.CreateObject("ADODB.Command")
		mCommand.ActiveConnection = Connection()
		mCommand.NamedParameters = False
		mCommand.CommandText = adCmdText
	End Sub

	Private Sub Dispose(obj)
		If IsObject(obj) Then
			Set obj = Nothing
		End If
	End Sub
End Class