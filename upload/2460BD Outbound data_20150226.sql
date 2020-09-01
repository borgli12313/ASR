-- BD Outbound @2015.02.25 

	Select Sku = OD.Sku , orderDate = DelvDt , qtypicked = OD.BaseQty , Location = OD.Location , 
	OrderNo = O.DNNumber , OrderLineNumber  = OD.TOLineNumber 
	from  SCH_TransferOrder O, SCH_TransferOrderDetail OD 
	Where O.Orderkey = OD.Orderkey 
	and exists ( select 1 from BAX_PACK_DTL pd where pd.OrderKey = O.Orderkey) 
	
	and DATEDIFF( d, DelvDt , '2014-10-01'  ) < 0 
	order by 2, 1
	
 
--Sku	Order date	qty	Picked location	Order no. 	Order line


	select status, * from SCH_TransferOrder O where AddDate > '2014-10-02' 
	 and Status not in ( 'C' ) and DelvDt is null 
	and exists ( select 1 from BAX_PACK_DTL pd where pd.OrderKey = O.Orderkey) 
	
	
	sp_help SCH_TransferOrderDetail
	

	


