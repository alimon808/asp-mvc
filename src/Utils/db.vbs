Class Database
   Private mConnection

   Public Property Get Connection()
      OpenConnection()
      Set Connection = mConnection
   End Property

   Sub Class_Initialize()

   End Sub

   Sub Class_Terminate()

   End Sub

   Public Function Execute(sql)
      OpenConnection()
      Set Execute = mConnection.Execute(sql)
   End Function

   Private Sub OpenConnection()
      Dispose(mConnection)
      Set mConnection = Server.CreateObject("ADODB.Connection")
      mConnection.Mode = 3
      mConnection.ConnectionString = Application("connectionString")
      mConnection.open
   End Sub

   Private Sub Dispose(obj)
      If IsObject(obj) Then
         Set obj = Nothing
      End If
   End Sub
End Class