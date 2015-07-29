class Project
	private mMetadata
	private mId
	private mProjectName
	private mActive
	private mPOP3Address

	private sub Class_Initialize()
		mMetadata = Array("Id",  "ProjectName"  ,  "Active"  ,  "POP3Address" )
	end sub

	private sub Class_Terminate()
	end sub

	public property get Id()
		Id = mId
	end property

	public property let Id(val)
		mId = val
	end property

	public property get ProjectName()
		ProjectName = mProjectName
	end property

	public property let ProjectName(val)
		mProjectName = val
	end property

	public property get Active()
		Active = mActive
	end property

	public property let Active(val)
		mActive = val
	end property

	public property get POP3Address()
		POP3Address = mPOP3Address
	end property

	public property let POP3Address(val)
		mPOP3Address = val
	end property

	public property get metadata()
		metadata = mMetadata
	end property
end class