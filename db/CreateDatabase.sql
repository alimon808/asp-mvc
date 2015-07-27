Create database test
GO

Use test

-- User --------------
CREATE TABLE [User]
(
[id] [int] IDENTITY(1,1) NOT NULL,
[FirstName] nvarchar(200)  NULL ,
[LastName] nvarchar(200)  NULL ,
[UserName] nvarchar(200)  NULL ,
[ProjectID] int  NULL 

, CONSTRAINT [PK_User] PRIMARY KEY CLUSTERED 
   (
	  [id] ASC
   )WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO




-- User --------------

Insert into [User]
(FirstName , LastName , UserName , ProjectID )
Values 
('Bhaskara','Ramachandra','BRamachandra',1)
GO


Insert into [User]
(FirstName , LastName , UserName , ProjectID )
Values 
('Padmavati','Pitambara','PPitambara',1)
GO



Insert into [User]
(FirstName , LastName , UserName , ProjectID )
Values 
('Rundhati','Shripati','RShripati',1)
GO
