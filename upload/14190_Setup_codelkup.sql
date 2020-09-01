
	Select * from codelist where listname = 'HUBOUTLOB' 

select * from codelkup where listname = 'HUBOUTLOB'
 
 
	INSERT INTO Codelist ( Listname, Description) 
	VALUES ( 'HUBOUTLOB' , 'HUB Outbound LOB' ) 

	INSERT INTO codelkup ( ListName, Code, Description) 
		VALUES ( 'HUBOUTLOB' , 'CHAN_DOM', 'Channel Domestic' ) 
	INSERT INTO codelkup ( ListName, Code, Description) 
		VALUES ( 'HUBOUTLOB' , 'CHAN_INT', 'Channel International' ) 

	INSERT INTO codelkup ( ListName, Code, Description) 
		VALUES ( 'HUBOUTLOB' , 'CHAN_INDIA', 'Channel International India' ) 

	INSERT INTO codelkup ( ListName, Code, Description) 
		VALUES ( 'HUBOUTLOB' , 'FF_INT', 'Channel Forwarder International' ) 

/* 
KPIAPLOB  
CHAN-IN   	Channel Intl - IN	KPIAPLOB  
CHAN_DOM  	Channel Local	KPIAPLOB  
CHAN_INT  	Channel International	KPIAPLOB  
HUB       	Non specific LOB	KPIAPLOB  
OTHER     	Not defined	KPIAPLOB  
RETAIL    	Retail	KPIAPLOB  
STOR_DOM  	Apple Store Local	KPIAPLOB  
STOR_INT  	Apple Store International	

*/ 