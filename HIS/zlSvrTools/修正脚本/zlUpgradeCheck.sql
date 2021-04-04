Create Or Replace Function zlUpgradeCheck(
	v_CurVer In Varchar2,
	v_NewVer In Varchar2
) Return Varchar2 Is
  v_Message Varchar2(100);
  v_Error   Varchar2(2000);
  v_Name    Varchar2(100);
Begin
  --检查是否已升级到要求的版本
  Select 内容 Into v_Message From zlRegInfo Where 项目 = '版本号';
  If Zlverdiff(v_Message, '9.25') = -1 Then
    v_Error := v_Error || Chr(10) || '因为管理工具9.25在自动升级工具改进前已经发布，因此本次自动升级需要管理工具先升级到9.25；管理工具从9.24升级到9.25需要手工进行。';
  Elsif Substr(v_Message, 1, 4) = '9.36' Then
    Select Max(Table_Name) Into v_Name From User_Tables Where Table_Name In ('门诊费用记录', '住院费用记录');
    If Not v_Name Is Null Then
      v_Error := v_Error || Chr(10) || '9.36升级到9.37时会将[病人费用记录]表拆分为[门诊费用记录]和[住院费用记录],检查发现已存在这样的表,请先删除或改名.';
    End If;
    v_Name := Null;
    Select Max(Trigger_Name) Into v_Name From User_Triggers Where Table_Name = '病人费用记录' And Status = 'ENABLED';
    If Not v_Name Is Null Then
      v_Error := v_Error || Chr(10) || '本次升级会将表[病人费用记录]改名并转移门诊费用,经检查发现该表上存在触发器,请先删除或禁用.';
    End If;
  End If;

  --检查必要的对象是否存在
  v_Message := Null;
  Begin
    Select Object_Name
    Into v_Message
    From User_Objects
    Where Object_Name = Upper('p_Reg_Apply') And Object_Type = 'PROCEDURE';
  Exception
    When Others Then
      Null;
  End;
  If v_Message Is Null Then
    v_Error := v_Error || Chr(10) || '管理工具数据库中缺少授权管理所必要的对象，请检查9.25及之前的升级是否正确。';
  End If;

  --整理返回信息
  If v_Error Is Not Null Then
    v_Error := Substr(v_Error, 2);
  End If;
  Return v_Error;
Exception
  When Others Then
    v_Error := v_Error || Chr(10) || '管理工具升级检查失败。';
    If v_Error Is Not Null Then
      v_Error := Substr(v_Error, 2);
    End If;
    Return v_Error;
End;
/
