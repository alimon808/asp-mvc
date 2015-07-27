<%

dim oConn 
set oConn = Nothing


public function DbOpenConnection()
   if oConn is nothing then
      set oConn = CreateObject("ADODB.Connection")
      oConn.Mode = 3
      oConn.ConnectionString = connectionString
      oConn.Open
   end if
   set DbOpenConnection = oConn
end function

public function DbCloseConnection()
   if not oConn is nothing Then           
      oConn.Close                     
      set oConn = nothing
   End If
   set DbCloseConnection = oConn
end function

public function DbExecute(sql)
   dim oConnection
   set oConnection = DbOpenConnection()
   set DbExecute = oConnection.execute(sql)
end function



%>