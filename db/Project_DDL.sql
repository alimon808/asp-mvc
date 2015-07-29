
-- Project --------------
CREATE TABLE [Project]
(
  [id] [int] IDENTITY(1,1) NOT NULL, 
  [ProjectName]    nvarchar(200)   NULL  , 
  [Active]    bit  NULL  , 
  [POP3Address]    nvarchar(200)   NULL 
, CONSTRAINT [PK_Project] PRIMARY KEY CLUSTERED ([id] ASC)WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]) ON [PRIMARY]
GO
    