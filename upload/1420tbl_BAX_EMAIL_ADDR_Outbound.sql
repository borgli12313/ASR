
if exists (select * from dbo.sysobjects where id = object_id(N'dbo.BAX_EMAIL_ADDR_Outbound') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table dbo.BAX_EMAIL_ADDR_Outbound
GO


CREATE TABLE dbo.BAX_EMAIL_ADDR_Outbound (
	Consigneekey    Char(20) NOT NULL ,	-- substring(ConsigneeName, 1, 20) 
	ConsigneeName	varchar(45) NOT NULL, 	
	EmailAddress 	varchar (50)  NOT NULL ,
	AddressType 	char (5)  NOT NULL ,		-- From , CC, To 
	ActiveFlag 		int NOT NULL ,				-- 1 = active, 0 - inactive
	AddWho			varchar(18) DEFAULT user_name(),
	AddDate			datetime DEFAULT getdate(),
   	LOB varchar(10) NULL ,
	CONSTRAINT PK_BAX_EMAIL_ADDR_outbound PRIMARY KEY  CLUSTERED 
	( ConsigneeKey, EmailAddress, AddressType ) 
) ON [PRIMARY]
GO


GRANT ALL On dbo.BAX_EMAIL_ADDR_Outbound to nSQL 
GO 


