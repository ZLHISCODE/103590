Create Or Replace Function zlUpgradeCheck(
	v_CurVer In Varchar2,
	v_NewVer In Varchar2
) Return Varchar2 Is
  v_Message Varchar2(100);
  v_Error   Varchar2(2000);
  v_Name    Varchar2(100);
Begin
  --����Ƿ���������Ҫ��İ汾
  Select ���� Into v_Message From zlRegInfo Where ��Ŀ = '�汾��';
  If Zlverdiff(v_Message, '9.25') = -1 Then
    v_Error := v_Error || Chr(10) || '��Ϊ������9.25���Զ��������߸Ľ�ǰ�Ѿ���������˱����Զ�������Ҫ��������������9.25�������ߴ�9.24������9.25��Ҫ�ֹ����С�';
  Elsif Substr(v_Message, 1, 4) = '9.36' Then
    Select Max(Table_Name) Into v_Name From User_Tables Where Table_Name In ('������ü�¼', 'סԺ���ü�¼');
    If Not v_Name Is Null Then
      v_Error := v_Error || Chr(10) || '9.36������9.37ʱ�Ὣ[���˷��ü�¼]����Ϊ[������ü�¼]��[סԺ���ü�¼],��鷢���Ѵ��������ı�,����ɾ�������.';
    End If;
    v_Name := Null;
    Select Max(Trigger_Name) Into v_Name From User_Triggers Where Table_Name = '���˷��ü�¼' And Status = 'ENABLED';
    If Not v_Name Is Null Then
      v_Error := v_Error || Chr(10) || '���������Ὣ��[���˷��ü�¼]������ת���������,����鷢�ָñ��ϴ��ڴ�����,����ɾ�������.';
    End If;
  End If;

  --����Ҫ�Ķ����Ƿ����
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
    v_Error := v_Error || Chr(10) || '���������ݿ���ȱ����Ȩ��������Ҫ�Ķ�������9.25��֮ǰ�������Ƿ���ȷ��';
  End If;

  --��������Ϣ
  If v_Error Is Not Null Then
    v_Error := Substr(v_Error, 2);
  End If;
  Return v_Error;
Exception
  When Others Then
    v_Error := v_Error || Chr(10) || '�������������ʧ�ܡ�';
    If v_Error Is Not Null Then
      v_Error := Substr(v_Error, 2);
    End If;
    Return v_Error;
End;
/
