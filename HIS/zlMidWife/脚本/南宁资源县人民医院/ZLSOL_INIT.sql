--����������ƻ�ʿ�˺ţ��ɸ���ʵ���������SQL��ѯ��ZLSOLΪ����ʿ�Ĺ��������ƣ�123Ϊ��¼����
--���Ե���100����Ҫ1���ӵ�ʱ��
--��zlsol��¼ִ�����½ű���conn zlsol/his@ORA_SOL
Declare
  Cursor c_Sql Is
    Select Distinct a.�û���, Substr(d.����, 1, 1) ��, Substr(d.����, 2) ��, d.����
    From �ϻ���Ա��@Zlhis_Dbl a, ��������˵��@Zlhis_Dbl b, ������Ա@Zlhis_Dbl c, ��Ա��@Zlhis_Dbl d
    Where b.����id = c.����id And c.��Աid = d.Id And a.��Աid = c.��Աid And b.�������� = '����';
  n_group_id number(18);
Begin
  For r In c_Sql Loop
    n_group_id:=Apex_Util.find_security_group_id(p_workspace => 'ZLSOL');
    Apex_Util.set_security_group_id(p_security_group_id => n_group_id);
    Apex_Util.Create_User(p_User_Name => r.�û���, p_First_Name => r.��, p_Last_Name => r.��, p_Web_Password => '123',p_Change_Password_On_First_Use => 'N');
  End Loop;
End;
