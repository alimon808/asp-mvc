Class User
	private mId
	private mFirstName
	private mLastName
	private mUserName
	private mProjectID
  
	private sub Class_Initialize()
	end sub

	private sub Class_Terminate()
	end sub

	Public Property Get Id()
		Id = mId
	End Property
 
	Public Property let Id(val)
		mId = val
	End Property
  
	Public Property Get FirstName()
		FirstName = mFirstName
	End Property  

	Public Property let FirstName(val)
		mFirstName = val
	End Property  

	Public Property Get LastName()
		LastName = mLastName
	End Property  

	Public Property let LastName(val)
		mLastName = val
	End Property  

	Public Property Get UserName()
		UserName = mUserName
	End Property  

	Public Property let UserName(val)
		mUserName = val
	End Property  

	Public Property Get ProjectID()
		ProjectID = mProjectID
	End Property  

	Public Property let ProjectID(val)
		mProjectID = val
	End Property  
End Class