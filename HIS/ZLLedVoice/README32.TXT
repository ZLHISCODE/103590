

五、WINDOWS 95，WINDOWS 98及WINDOWS NT调用说明
包括程序名：
CKY32.DLL,CKY32.H。
CKY32.DLL是一个32位的动态库兼容Windows 95和Windows NT
它包括一个可在任何Windows 95和Windows NT开发平台中调用
的函数，函数原型说明在头文件CKY32.H中。

二、函数说明：
函数原型：
integer dsbdll(integer COMPort,LPSTR OutString);
函数的第一个参数用来规定串行口，有效值如下：
1  ：    计算机串口1；
2  ：    计算机串口2；
用户应根据计算机串口资源分配的具体情况作一个选择。
函数的第二个参数是上位计算机向下位语音器发送的具体数据，
数据格式说明如下：
OutString是一个长字符串指针，如果要使语音器讲“预收12345.60元 ”，则在
应用程序中作如下调用(假定使用串口1)：
dsbdll(1,(LPSTR)"12345.60Y");
注意：1、OutString所代表的数值的最大值不要一次大于十万，但当数值是整数时，
	小数点可有可无，小数点以后的两个零也可有可无，这个函数会全部自动添加；
	另外，如果小数点后的位数超过两位，则舍去两位以后的小数位。
	这里"Y"表示“预收”。
      2、每个OutString都以一个字母结束，而且一次调用只能用一个特征字母，
	 字母涵义在第四条说明。

三、 用于Windows 95平台 POWERBUILDER调用动态库 
     把W32目录中的CKY32.DLL拷贝到当前目录或Windows 95的系统目录中，然后在
     PowerBuilder中作全局函数说明，说明格式为：
     FUNCTION INTEGER dsbdll(INTEGER COMPOYT,
     STRING OUTSTRING) LIBRAT“CKY32.DLL”


四、特征字母说明：

1.  输入门诊号或病人姓名后加
    DSBDLL(1,'F') (复位清零) 
    DSBDLL(1,'W') (您好,请稍等)
2.  计算应收药费后加DSBDLL(1,XX+'J') (请您付款***元)
3.  预收款后加   DSBDLL(1,XX+'Y')  (预收***元)
    然后将预收与应收款比较,如相等则加DSBDLL(1,'X')  (谢谢)
      如不相等则加  DSBDLL(1,XX+'Z')  (找零***元)
			DSBDLL(1,'D')  (请当面点清谢谢)
      三、礼貌用语键
	      收费语音                     
    DSBDLL(1,'a')  --请问您的姓名。          
    DSBDLL(1,'b')  --请您到中药房批价。    
    DSBDLL(1,'c')  --请您到X光室批价。     
    DSBDLL(1,'d')  --请您到注射室做皮试。     
    DSBDLL(1,'e')  --请您到门诊办公室审核盖章。
    DSBDLL(1,'f')  --请您到挂号室输入门诊号。  
    DSBDLL(1,'g')  --请您到B超室批价。         
    DSBDLL(1,'h')  --请您把病历卡拿出来。      
    DSBDLL(1,'i')  --找零请当面点清,谢谢。     

		    