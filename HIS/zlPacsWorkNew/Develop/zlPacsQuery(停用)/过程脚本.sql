--��ȡͼ����Դ
CREATE OR REPLACE Function Zl_Ӱ���ѯ_��ȡͼ��
(
  ����_In  In Ӱ���ѯ��Դ.��Դ����%Type,
  Pos_In In Number
) Return Varchar2 Is
  v_Buffer Varchar2(32767);
  l_Blob   Blob;
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin
  Select ͼ�� Into l_Blob From Ӱ���ѯ��Դ Where ��Դ���� = ����_In;
  
  n_Offset := n_Offset + Pos_In * n_Amount;
  If l_Blob Is Null Then
    v_Buffer := Null;
  Else
    Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
  End If;
  
  Return v_Buffer;
Exception
  When No_Data_Found Then
    Return Null;
End Zl_Ӱ���ѯ_��ȡͼ��;
/

--����û���������
Create Or Replace Procedure zl_Ӱ���ѯ_�������
(
   �û�ID_In      Ӱ���ѯ����.�û�ID%Type
) Is
Begin
    Delete From Ӱ���ѯ���� Where �û�ID=�û�ID_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--�����û�������ѯ����
Create Or Replace Procedure zl_Ӱ���ѯ_���¹���
(
   �û�ID_In      Ӱ���ѯ����.�û�ID%Type,
   ����ID_In      Ӱ���ѯ����.��ѯ����ID%Type,
   �Ƿ�Ĭ��_In    Ӱ���ѯ����.�Ƿ�Ĭ��%Type,
   �Ƿ���_In    Ӱ���ѯ����.�Ƿ���%Type,
   ����վ��_In    Ӱ���ѯ����.����վ��%Type
) Is
Begin
    Insert Into Ӱ���ѯ����(ID,�û�ID,��ѯ����ID,�Ƿ�Ĭ��,�Ƿ���,����վ��)
    Values(Ӱ���ѯ����_ID.NEXTVAL,�û�ID_In,����ID_In,�Ƿ�Ĭ��_In,�Ƿ���_In,����վ��_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_Ӱ���ѯ_���¹���;
/ 

--�����û����˴���¼������
Create Or Replace Procedure zl_Ӱ���ѯ_��������
(
   �û�ID_In      Ӱ���ѯ����.�û�ID%Type,
   ��ѯ����ID_In  Ӱ���ѯ����.��ѯ����ID%Type,
   ��������_In    Ӱ���ѯ����.��������%Type
) Is
Begin
    Update Ӱ���ѯ���� 
    Set ��������=��������_In
    Where �û�ID=�û�ID_In And ��ѯ����ID=��ѯ����ID_In;
   
    If Sql%RowCount <=0 Then
        Insert Into 
               Ӱ���ѯ����(ID, �û�ID, ��ѯ����ID, ��������)
        Values
               	(Ӱ���ѯ����_ID.NEXTVAL, �û�ID_In, ��ѯ����ID_In, ��������_In);
    End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_Ӱ���ѯ_��������;
/
--********************************************************************************

Create Or Replace Procedure Zl_Ӱ���ѯ_�༭��������
(
  Id_In   In Ӱ���ѯ����.Id%Type,
  Text_In In Ӱ���ѯ����.��������%Type
) Is
  l_Clob Clob;
Begin
  Update Ӱ���ѯ���� Set �������� = Empty_Clob() Where Id = Id_In;
  Select �������� Into l_Clob From Ӱ���ѯ���� Where Id = Id_In For Update;
  Dbms_Lob.Writeappend(l_Clob, Length(Text_In), Text_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_�༭��������;
/
--------------------------------------------------------------------------------------------------------------

Create Or Replace Procedure Zl_Ӱ���ѯ_ɾ��ͼ��(��Դ����_In In Ӱ���ѯ��Դ.��Դ���� %Type) Is
Begin
  Delete From Ӱ���ѯ��Դ Where ��Դ���� = ��Դ����_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_ɾ��ͼ��;

/

Create Or Replace Procedure Zl_Ӱ���ѯ_����ͼ��
(
  ��Դ����_In In Ӱ���ѯ��Դ.��Դ����%Type,
  ��Դ����_In In Ӱ���ѯ��Դ.��Դ����%Type
) Is
  n_Id Number(18);
Begin
  Select Ӱ���ѯ��Դ_Id.Nextval Into n_Id From Dual;
  Insert Into Ӱ���ѯ��Դ (Id, ��Դ����, ��Դ����) Values (n_Id, ��Դ����_In, ��Դ����_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_����ͼ��;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_����ͼ��
(
  ��Դ����_In In Ӱ���ѯ��Դ.��Դ����%Type,
  ͼ��_In     In Varchar2, --16���Ƶ��ļ�Ƭ�λ�����Ƭ�� 
  Cls_In      In Number := 0 --�Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0 
) Is
  l_Blob Blob;
Begin

  If Cls_In = 1 Then
    Update Ӱ���ѯ��Դ Set ͼ�� = Empty_Blob() Where ��Դ���� = ��Դ����_In;
  End If;
  Select ͼ�� Into l_Blob From Ӱ���ѯ��Դ Where ��Դ���� = ��Դ����_In For Update;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(ͼ��_In)) / 2, Hextoraw(ͼ��_In));
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_����ͼ��;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_��������
(
  ��������_In In Ӱ���ѯ����.��������%Type,
  ����˵��_In In Ӱ���ѯ����.����˵��%Type,
  �Ƿ�Ĭ��_In In Ӱ���ѯ����.�Ƿ�Ĭ��%Type,
  ʹ��״̬_In In Ӱ���ѯ����.��������%Type,
  �Ƿ���_In In Ӱ���ѯ����.�Ƿ���%Type,
  ����ģ��_In In Ӱ���ѯ����.����ģ��%Type,
  ��������_In In Ӱ���ѯ����.��������%Type
) Is
  n_Id       Number(18);
  n_������� Number(18);
Begin
  Select Ӱ���ѯ����_Id.Nextval Into n_Id From Dual;
  Select Nvl(Max(�������), 0) + 1 Into n_������� From Ӱ���ѯ���� Where ����ģ�� = ����ģ��_In;

  Insert Into Ӱ���ѯ����
    (Id, ��������, ����˵��, �Ƿ�Ĭ��, ʹ��״̬, �������, �Ƿ���, ����ģ��, �汾)
  Values
    (n_Id, ��������_In, ����˵��_In, �Ƿ�Ĭ��_In, ʹ��״̬_In, n_�������, �Ƿ���_In, ����ģ��_In, '1');

  Zl_Ӱ���ѯ_�༭��������(n_Id, ��������_In);

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_��������;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_���·���
(
  Id_In       In Ӱ���ѯ����.Id%Type,
  ��������_In In Ӱ���ѯ����.��������%Type,
  ����˵��_In In Ӱ���ѯ����.����˵��%Type,
  �Ƿ�Ĭ��_In In Ӱ���ѯ����.�Ƿ�Ĭ��%Type,
  ʹ��״̬_In In Ӱ���ѯ����.��������%Type,
  �Ƿ���_In In Ӱ���ѯ����.�Ƿ���%Type,
  ����ģ��_In In Ӱ���ѯ����.����ģ��%Type,
  ��������_In In Ӱ���ѯ����.��������%Type
) Is
Begin

  Update Ӱ���ѯ����
  Set �������� = ��������_In, ����˵�� = ����˵��_In, �Ƿ�Ĭ�� = �Ƿ�Ĭ��_In, ʹ��״̬ = ʹ��״̬_In, �Ƿ��� = �Ƿ���_In, ����ģ�� = ����ģ��_In, �汾 = �汾 + 1
  Where Id = Id_In;

  Zl_Ӱ���ѯ_�༭��������(Id_In, ��������_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_���·���;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_ɾ������(Id_In In Ӱ���ѯ����.Id%Type) Is
Begin
  Delete From Ӱ���ѯ���� Where Id = Id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_ɾ������;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_�ƶ�����
(
  ����id_In   In Ӱ���ѯ����.Id%Type,
  �����_In   In Ӱ���ѯ����.�������%Type,
  ����ģ��_In In Ӱ���ѯ����.����ģ��%Type
) Is
  n_Order Number;
Begin
  Begin
    Select ������� Into n_Order From Ӱ���ѯ���� Where Id = ����id_In And ����ģ�� = ����ģ��_In;
  Exception
    When Others Then
      Return;
  End;

  If �����_In < n_Order Then
    Update Ӱ���ѯ����
    Set ������� = ������� + 1
    Where ������� >= �����_In And ������� < n_Order And ����ģ�� = ����ģ��_In;
  
    Update Ӱ���ѯ���� Set ������� = �����_In Where Id = ����id_In And ����ģ�� = ����ģ��_In;
  Else
    Update Ӱ���ѯ����
    Set ������� = ������� - 1
    Where ������� > n_Order And ������� <= �����_In And ����ģ�� = ����ģ��_In;
  
    Update Ӱ���ѯ���� Set ������� = �����_In Where Id = ����id_In And ����ģ�� = ����ģ��_In;
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_�ƶ�����;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_Ĭ�Ϸ���
(
  ����id_In   In Ӱ���ѯ����.Id%Type,
  �Ƿ�Ĭ��_In In Ӱ���ѯ����.�Ƿ�Ĭ��%Type,
  ����ģ��_In In Ӱ���ѯ����.����ģ��%Type
) Is
Begin
  Update Ӱ���ѯ���� Set �Ƿ�Ĭ�� = 0 Where �Ƿ�Ĭ�� = 1 And ����ģ�� = ����ģ��_In;
  Update Ӱ���ѯ���� Set �Ƿ�Ĭ�� = �Ƿ�Ĭ��_In Where Id = ����id_In And ����ģ�� = ����ģ��_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_Ĭ�Ϸ���;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_���÷���
(
  ����id_In   Ӱ���ѯ����.Id%Type,
  �Ƿ���_In Ӱ���ѯ����.�Ƿ�ϵͳ��ѯ%Type
) Is
Begin
  Update Ӱ���ѯ���� Set �Ƿ��� = �Ƿ���_In Where Id = ����id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_���÷���;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_���÷���
(
  ����id_In   Ӱ���ѯ����.Id%Type,
  ʹ��״̬_In Ӱ���ѯ����.ʹ��״̬%Type
) Is
Begin
  Update Ӱ���ѯ���� Set ʹ��״̬ = ʹ��״̬_In Where Id = ����id_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Ӱ���ѯ_���÷���;
/

Create Or Replace Procedure Zl_Ӱ���ѯ_���Ի�����
(
  �û�id_In     Ӱ���ѯ����.�û�id%Type,
  ��ѯ����id_In Ӱ���ѯ����.��ѯ����id%Type,
  ��������_In   Ӱ���ѯ����.��������%Type,
  �б�����_In   Ӱ���ѯ����.�б�����%Type
) Is
Begin
  Update Ӱ���ѯ����
  Set �������� = ��������_In, �б����� = �б�����_In
  Where �û�id = �û�id_In And ��ѯ����id = ��ѯ����id_In;

  If Sql%RowCount <= 0 Then
    Insert Into Ӱ���ѯ����
      (ID, �û�id, ��ѯ����id, ��������, �б�����)
    Values
      (Ӱ���ѯ����_Id.Nextval, �û�id_In, ��ѯ����id_In, ��������_In, �б�����_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ӱ���ѯ_���Ի�����;
/


