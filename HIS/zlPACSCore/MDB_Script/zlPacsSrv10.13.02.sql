alter table 强制结果 add 默认强制结果 text(100);

--Scheduled Procedure Step 预定过程步骤


update 强制结果 set 默认强制结果='RemoteAE' where 组号='40' and 元素号='1';
update 强制结果 set 默认强制结果='CT' where 组号='8' and 元素号='60';

	    
--Requested Procedure 请求的过程
update 强制结果 set 默认强制结果='12345' where 组号='20' and 元素号='D';

	   	  	    	    
--Image Service Request 图像服务请求
update 强制结果 set 默认强制结果='54321' where 组号='8' and 元素号='50';

	    
--Patient Identification  病人标识
update 强制结果 set 默认强制结果='CESHI' where 组号='10' and 元素号='10';
update 强制结果 set 默认强制结果='11111' where 组号='10' and 元素号='20';
	    
update 版本表 set 版本号='10.13.02';