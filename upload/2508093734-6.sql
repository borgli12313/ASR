select status,doreplenish,* from pickdetail
where dropid = '0000003198'


select status,* from orders
where orderkey in
(select distinct orderkey from pickdetail
where dropid = '0000003198')

select status, * from bax_kewill_header
where header_id = '0000028912'


select status, * from bax_kewill_details
where header_id = '0000028912'


select status,* from bax_hub_orders
where order_number = 8093734
and ship_set_number = 6


insert into bax_hub_orders
values('0000028661','',8093734,6,'GB107328000000','2003-06-15 00:00:00.000','2003-06-12 00:00:00.000','320483','CISCO SYSTEMS INTERNATIONAL BV','DE DIEZE 17','UPS GLOBAL LOGISTICS WAREHOUSE','','','BEST','','5684PR','NL','','BXS','EXP6','EXP6',3,60,'CIP/Customer Premises/Duty Paid/Add','CIP/Customer Premises/Duty Paid/Add','','','','Y','Y',3,'0','TC','JOHN KIRBY','0','2003-06-12 13:49:05.607','CIS.STEVEN','' ,'','N','2003-06-12 13:49:05.607','Add')

select * from bax_hub_orders
where HUB_ORD_KEY = '0000019269'

sp_who 123


sp_lock

select * from ncounter


update ncounter set keycount = 28661
where KEYNAME = 'HUB_ORDKEY' 

select status,* from bax_hub_orderdetail


select * from bax_hub_orders

select status,* from bax_hub_orderdetail

where order_number = 8093734
and ship_set_number = 6

update bax_hub_orderdetail
set HUB_ORD_KEY = '0000028661'
where CARTON_NUMBER in ('29903907','29903909','29903911')

select status,dropid,* from pickdetail
where caseid = '0000004638'



select * from 