----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Alter Table �ҺŰ��żƻ� Add �ϴμƻ�ID number(18);

Create Index �ҺŰ��żƻ�_IX_�ϴμƻ�ID on �ҺŰ��żƻ�(�ϴμƻ�ID) Tablespace zl9Indexhis;

Alter table �ҺŰ��żƻ�
  add constraint �ҺŰ��żƻ�_FK_�ϴμƻ�ID foreign key (�ϴμƻ�ID)
  references �ҺŰ��żƻ� (ID);

------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Declare
  Cursor c_���ڼƻ� Is
          Select Id, ����id, ��Чʱ��, ʧЧʱ��, ���ʱ��
          From �ҺŰ��żƻ�
          Where ʧЧʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') And ���ʱ�� Is Not Null
          Order By ����id, ��Чʱ��, Id;

  n_�ƻ�ID �ҺŰ��żƻ�.ID%Type;
  n_����ID �ҺŰ��żƻ�.����ID%Type;
Begin
  For r_���ڼƻ� In c_���ڼƻ� Loop
    IF Nvl(n_����ID, 0) <> r_���ڼƻ�.����ID Then
      n_����ID := r_���ڼƻ�.����ID;
      n_�ƻ�ID := 0;
    End if;
    IF Nvl(n_�ƻ�ID, 0) <> 0 Then
      Update �ҺŰ��żƻ� Set ʧЧʱ�� = r_���ڼƻ�.��Чʱ�� Where ID = n_�ƻ�ID;
      Update �ҺŰ��żƻ� Set �ϴμƻ�ID = n_�ƻ�ID Where ID = r_���ڼƻ�.ID;
    End IF;
    n_�ƻ�ID := r_���ڼƻ�.ID;
  End Loop;
End;
/

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--134254:������,2018-11-16,������������ӡ
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1214, '����', User, 'Zl_Lob_Read', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1214 And ���� = '����' And Upper(����) = Upper('Zl_Lob_Read'));


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_�ҺŰ��żƻ�_Verify
(
  Id_In         In �ҺŰ��żƻ�.Id%Type,
  ������Ч_In   Number := 0,
  �ϴμƻ�ID_In In �ҺŰ��żƻ�.�ϴμƻ�Id%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg   Varchar2(100);
  v_User_Name ��Ա��.����%Type;
  n_Valied    Number(1);
  d_��Чʱ��  �ҺŰ��żƻ�.��Чʱ��%Type;
Begin
  
  Select Nvl(Max(p.����),'') Into v_User_Name From �ϻ���Ա�� o, ��Ա�� p Where o.��Աid = p.Id And �û��� = User;
  If v_User_Name Is Null Then
    v_Err_Msg := '[ZLSOFT]��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����' || Chr(10) || Chr(13) ||
                 'ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Valied From �ҺŰ��żƻ� a Where Nvl(��Чʱ��, Sysdate) < Sysdate And a.Id = Id_In;
  If n_Valied > 0 And Nvl(������Ч_In, 0) = 0 Then
    v_Err_Msg := '[ZLSOFT]�üƻ����ŵ���Чʱ���Ѿ����ڣ����ܽ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
  
  Update �ҺŰ��żƻ�
  Set ����� = v_User_Name, ���ʱ�� = Sysdate, �ϴμƻ�ID = �ϴμƻ�ID_In,
      ��Чʱ�� = Case Nvl(������Ч_In, 0) When 0 Then ��Чʱ�� Else Sysdate - 1 / 24 / 60 / 60 End
  Where Id = Id_In And ���ʱ�� Is Null
  Return ��Чʱ�� Into d_��Чʱ��;
  If Sql%Notfound Then
    v_Err_Msg := '[ZLSOFT]�üƻ������Ѿ���������˻�ɾ��,���������![ZLSOFT]';
    Raise Err_Item;
  End If;
  IF Nvl(�ϴμƻ�ID_In, 0) <> 0 Then
    Update �ҺŰ��żƻ� Set ʧЧʱ�� = d_��Чʱ�� Where ID = �ϴμƻ�ID_In;
  End IF;
  If Nvl(������Ч_In, 0) = 1 Then
    Begin
      Zl_�ҺŰ���_Autoupdate();
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ҺŰ��żƻ�_Verify;
/

--133584:���ϴ�,2018-11-15,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_�ҺŰ��żƻ�_Cancel(Id_In In �ҺŰ��żƻ�.ID%Type) Is
  Err_Item        Exception;
  v_Err_Msg       Varchar2(100);
  v_User_Name     ��Ա��.����%Type;
  n_�ϴμƻ�ID    �ҺŰ��żƻ�.�ϴμƻ�Id%Type;
Begin
  Begin
    Select P.���� Into v_User_Name From �ϻ���Ա�� O, ��Ա�� P Where O.��Աid = P.ID And �û��� = User;
  Exception
    When Others Then
      v_User_Name := Null;
  End;
  If v_User_Name Is Null Then
    v_Err_Msg := '[ZLSOFT]��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����' || Chr(10) || Chr(13) ||
                 'ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�[ZLSOFT]';
    Raise Err_Item;
  End If;
  Begin
    Select 'Yes'
    Into v_Err_Msg
    From �ҺŰ��żƻ� L
    Where ID = Id_In And Nvl(ʵ����Ч, To_Date('3000-01-01', 'yyyy-mm-dd')) >= To_Date('3000-01-01', 'yyyy-mm-dd');
  Exception
    When Others Then
      v_Err_Msg := 'No';
  End;
  If v_Err_Msg = 'No' Then
    v_Err_Msg := '[ZLSOFT]�üƻ������Ѿ�����Ч,����ȡ�����![ZLSOFT]';
    Raise Err_Item;
  End If;

  Update �ҺŰ��żƻ� Set ����� = Null, ���ʱ�� = Null Where ID = Id_In And ���ʱ�� Is Not Null
  Return �ϴμƻ�ID Into n_�ϴμƻ�ID;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]�üƻ������Ѿ�������ȡ����˻�ɾ��,������ȡ�����![ZLSOFT]';
    Raise Err_Item;
  End If;
  If Nvl(n_�ϴμƻ�ID, 0) <> 0 Then
    Update �ҺŰ��żƻ� Set ʧЧʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Where ID = n_�ϴμƻ�ID;
    Update �ҺŰ��żƻ� Set �ϴμƻ�ID = NULL Where ID = Id_In;
  End IF;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);

  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ҺŰ��żƻ�_Cancel;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_�ҺŰ���_Autoupdate Is
  Err_Item Exception;
  v_Date Date;
  -- v_Err_Msg Varchar2(100); 
  v_Unitscount Number;
Begin
  --n_����ִ���� ���Ƿ���²��˹Һż�¼ ��������ü�¼�е�ִ���� 
  --               ����ƻ��и����� �Һ���Ŀ ��������� ���˹Һż�¼��������ü�¼�е����� 
  Select Sysdate Into v_Date From Dual;
  Select Count(0) Into v_Unitscount From ������λ���ſ��� Where Rownum = 1;

  For v_��Ч In (Select ID, ����id, ����, ��Чʱ��, ʧЧʱ��, ����, ��һ, �ܶ�, ����, ����, ����, ����, ���﷽ʽ, ��ſ���, ִ��ʱ�� As �ϴ���Чʱ��, ��Ŀid, ҽ������, ҽ��id,
                      ���, ����id
               From (Select a.Id, a.����id, a.����, a.��Чʱ��, a.ʧЧʱ��, a.����, a.��һ, a.�ܶ�, a.����, a.����, a.����, a.����, a.���﷽ʽ, a.��ſ���,
                             b.ִ��ʱ��, a.��Ŀid, a.ҽ������, a.ҽ��id, Nvl(b.ִ�мƻ�id, 0) As ִ�мƻ�id,
                             Row_Number() Over(Partition By a.����id Order By a.��Чʱ�� Desc) As ˳���, b.���, b.����id
                      From �ҺŰ��żƻ� A, �ҺŰ��� B
                      Where Sysdate Between a.��Чʱ�� + 0 And a.ʧЧʱ�� And a.����id = b.Id And
                            a.ʵ����Ч >= To_Date('3000-01-01', 'yyyy-mm-dd') And ���ʱ�� Is Not Null And
                            b.ͣ������ Is Null)
               Where ˳��� = 1 And ID <> Nvl(ִ�мƻ�id, 0)) Loop
    Update �ҺŰ��żƻ� Set ʵ����Ч = v_��Ч.�ϴ���Чʱ�� Where ID = v_��Ч.����id And ʧЧʱ�� < v_��Ч.ʧЧʱ��;
  
    Update �ҺŰ���
    Set ���� = v_��Ч.����, ��һ = v_��Ч.��һ, �ܶ� = v_��Ч.�ܶ�, ���� = v_��Ч.����, ���� = v_��Ч.����, ���� = v_��Ч.����, ���� = v_��Ч.����,
        ���﷽ʽ = v_��Ч.���﷽ʽ, ��ſ��� = v_��Ч.��ſ���, ��ʼʱ�� = Sysdate, ��ֹʱ�� = v_��Ч.ʧЧʱ��, ��Ŀid = Nvl(v_��Ч.��Ŀid, ��Ŀid), ִ��ʱ�� = v_Date,
        ִ�мƻ�id = v_��Ч.Id, ��� = 9999999, ҽ������ = v_��Ч.ҽ������, ҽ��id = v_��Ч.ҽ��id
    Where ID = v_��Ч.����id;
  
    --���µ������ 
    Update �ҺŰ��� A
    Set ��� = -1 * ���
    Where ��Ŀid = v_��Ч.��Ŀid And a.����id = v_��Ч.����id And Nvl(a.ҽ������, '-') = Nvl(v_��Ч.ҽ������, '-') And
          Nvl(a.ҽ��id, 0) = Nvl(v_��Ч.ҽ��id, 0);
    For v_��� In (Select a.Id, Rownum As ���
                 From �ҺŰ��� A
                 Where a.��Ŀid = v_��Ч.��Ŀid And a.����id = v_��Ч.����id And Nvl(a.ҽ������, '-') = Nvl(v_��Ч.ҽ������, '-') And
                       Nvl(a.ҽ��id, 0) = Nvl(v_��Ч.ҽ��id, 0)
                 Order By a.Id) Loop
      Update �ҺŰ��� A Set ��� = v_���.��� Where ID = v_���.Id;
    End Loop;
    Delete �ҺŰ������� Where �ű�id = v_��Ч.����id;
    Insert Into �ҺŰ�������
      (�ű�id, ��������)
      Select v_��Ч.����id, �������� From �Һżƻ����� Where �ƻ�id = v_��Ч.Id;
    Delete �ҺŰ������� Where ����id = v_��Ч.����id;
    Insert Into �ҺŰ�������
      (����id, ������Ŀ, �޺���, ��Լ��)
      Select v_��Ч.����id, ������Ŀ, �޺���, ��Լ�� From �Һżƻ����� Where �ƻ�id = v_��Ч.Id;
    Delete �ҺŰ���ʱ�� Where ����id = v_��Ч.����id;
    Insert Into �ҺŰ���ʱ��
      (����id, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ, ����)
      Select v_��Ч.����id, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ, ����
      From �Һżƻ�ʱ��
      Where �ƻ�id = v_��Ч.Id;
    If Nvl(v_Unitscount, 0) > 0 Then
      Delete ������λ���ſ��� Where ����id = v_��Ч.����id;
      Insert Into ������λ���ſ���
        (����id, ������λ, ������Ŀ, ���, ����)
        Select v_��Ч.����id, ������λ, ������Ŀ, ���, ���� From ������λ�ƻ����� Where �ƻ�id = v_��Ч.Id;
    End If;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ҺŰ���_Autoupdate;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_����ԤԼ����_ȡ��
(
  No_In         ������ü�¼.No%Type,
  ����_In       ������ü�¼.��ҩ����%Type,
  ����id_In     ������ü�¼.����id%Type,
  ҽ������_In   ������ü�¼.ִ���� %Type := Null,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  ժҪ_In       ���˹Һż�¼.ժҪ%Type := Null,
  ����_In       ���˹Һż�¼.����%Type := Null
) As

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count    Number(18);
  n_�Ŷ�     Number;
  n_�����Ŷ� Number;
  n_���ɶ��� Number(3);
  d_Date     Date;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�������� �ŶӽкŶ���.��������%Type;
  v_�Ŷ���� �ŶӽкŶ���.�Ŷ����%Type;
  d_�Ŷ�ʱ�� �ŶӽкŶ���.�Ŷ�ʱ��%Type;

  d_ԤԼʱ��   ������ü�¼.����ʱ��%Type;
  d_����ʱ��   ������ü�¼.����ʱ��%Type;
  n_�����¼id ���˹Һż�¼.�����¼id%Type;
  v_����Ա���� ���˹Һż�¼.������%Type;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ %Type;
  n_�Һ�id     ���˹Һż�¼.Id%Type;
  n_��id       ����ɿ����.Id%Type;

  n_��¼״̬   ���˹Һż�¼.��¼״̬%Type;
  n_����       ���˹Һż�¼.����%Type;
  v_����       ���˹Һż�¼.�ű�%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_��Դid     �ٴ������Դ.Id%Type;
  n_����id     ���˽��ʼ�¼.Id%Type;
  n_����id     �ҺŰ���.Id%Type;
  n_�ƻ�id     �ҺŰ��żƻ�.Id%Type;
  n_�Ƿ��ʱ�� �ٴ������¼.�Ƿ��ʱ��%Type := 0;

  n_����ģʽ Number := 0;
  v_Paratemp Varchar2(100);
  d_����ʱ�� Date;
  n_�Һ�ģʽ Number(3);

Begin
  n_��id := Zl_Get��id(����Ա����_In);

  n_���ɶ��� := To_Number(Nvl(Zl_Getsysparameter('�Ŷӽк�ģʽ', 1113), 0));

  --0-ԤԼ������������ģʽ 1-ԤԼ���ղ�����ģʽ 
  n_����ģʽ := Nvl(Zl_Getsysparameter('ԤԼ����ģʽ', 1111), 0);
  v_Paratemp := Nvl(Zl_Getsysparameter('�Һ��Ű�ģʽ'), 0);
  n_����ģʽ := Nvl(Zl_Getsysparameter('ԤԼ����ģʽ', 1111), 0);

  If n_�Һ�ģʽ = 1 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;

  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ, ��¼״̬, ��¼����, ������, �����¼id
    Into v_����, n_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ, n_��¼״̬, n_Count, v_����Ա����, n_�����¼id
    From ���˹Һż�¼ A
    Where ��¼״̬ In (1, 3) And NO = No_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := -1;
  End;

  If n_Count = -1 Then
    v_Err_Msg := 'ԤԼ�Һŵ�:' || No_In || '������';
    Raise Err_Item;
  End If;

  If n_Count = 1 Then
    If n_��¼״̬ = 3 Then
      v_Err_Msg := 'ԤԼ�Һŵ�:' || No_In || '�Ѿ����˺�';
      Raise Err_Item;
    End If;
    If v_����Ա���� <> ����Ա����_In Then
      v_Err_Msg := 'ԤԼ�Һŵ�:' || No_In || '�ѱ�����';
      Raise Err_Item;
    Else
      v_Err_Msg := 'ԤԼ�Һŵ�:' || No_In || '�ѱ����˽���';
      Raise Err_Item;
    End If;
  End If;

  If d_����ʱ�� Is Not Null Then
    If d_����ʱ�� < d_����ʱ�� Then
      v_Err_Msg := '��ǰԤԼ�Һŵ����ڳ�����Ű�ģʽ���ţ�������' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '֮ǰ����!';
      Raise Err_Item;
    End If;
  End If;
  
  --��ʾǩ����
  UPDATE ���˹Һż�¼ SET ��¼��־=1 WHERE no=no_In;

  --�ж��Ƿ��ʱ��
  n_�����¼id := Nvl(n_�����¼id, 0);
  If n_�����¼id = 0 Then
    n_Count := 0;
    Select Max(ID) Into n_����id From �ҺŰ��� Where ���� = v_����;
  
    Select Max(ID)
    Into n_�ƻ�id
    From �ҺŰ��żƻ� A
    Where a.����id = n_����id And ���ʱ�� Is Not Null And
          a.��Чʱ�� = (Select Max(��Чʱ��)
                    From �ҺŰ��żƻ� A
                    Where ����id = n_����id And Sysdate Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                          ʧЧʱ�� And ���ʱ�� Is Not Null);
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Max(1) Into n_�Ƿ��ʱ�� From �ҺŰ���ʱ�� A Where a.����id = n_����id And Rownum < 2;
    Else
      Select Max(1) Into n_�Ƿ��ʱ�� From �Һżƻ�ʱ�� A Where a.�ƻ�id = n_�ƻ�id And Rownum < 2;
    End If;
  Else
    --�ж��Ƿ��ʱ��
    Select Nvl(�Ƿ��ʱ��, 0), ��Դid Into n_�Ƿ��ʱ��, n_��Դid From �ٴ������¼ Where ID = n_�����¼id;
  End If;

  --��ʱ�εĺű�ֻ�ܵ������
  If n_�Ƿ��ʱ�� = 1 Then
    If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ��ȡ����ţ�';
      Raise Err_Item;
    End If;
  End If;

  If n_�Ƿ��ʱ�� = 0 Then
    --0-ԤԼ������������ģʽ 1-ԤԼ���ղ�����ģʽ
    If n_����ģʽ = 0 Then
      If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
        d_����ʱ�� := Sysdate;
      End If;
    End If;
  End If;

  If Not n_���� Is Null Then
  
    If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
      If Nvl(n_�����¼id, 0) = 0 Then
      
        --��ǰ���ջ��ӳٽ���:��ɾ�������ԤԼʱ�����
        Delete �Һ����״̬ Where ���� = v_���� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = n_����;
      
        --����ǰʱ��ĺ�
        Zl_�ҺŰ���_��ͳ_Lockno(2, v_����, d_����ʱ��, Null, n_����, Null, ����Ա����_In, n_����id, n_�ƻ�id, 0, ����Ա����_In || '����', Null, Null);
        Update �Һ����״̬
        Set ״̬ = 1, �Ǽ�ʱ�� = Sysdate
        Where Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = n_���� And ���� = v_���� And ״̬ In (2, 5);
      Else
        --��ǰ���ջ��ӳٽ���:���޸ĵ����ԤԼʱ�����
        Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where ��� = n_���� And ��¼id = n_�����¼id;
        For c_�ɰ��� In (Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
                      
                      From �ٴ������¼
                      Where ID = n_�����¼id) Loop
        
          Begin
            n_Count := 1;
            Select ID
            Into n_�����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = c_�ɰ���.�Ƿ��ʱ�� And �Ƿ���ſ��� = c_�ɰ���.�Ƿ���ſ��� And ����id = c_�ɰ���.����id And
                  Nvl(ҽ��id, 0) = Nvl(c_�ɰ���.ҽ��id, 0) And �ϰ�ʱ�� = c_�ɰ���.�ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And
                  �������� = Trunc(Sysdate) And Rownum < 2;
          Exception
            When Others Then
              n_Count := 0;
          End;
          If n_Count = 0 Then
            v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
            Raise Err_Item;
          End If;
        
          Zl_�ҺŰ���_�ٴ�����_Lockno(2, n_�����¼id, d_����ʱ��, Null, n_����, 0, ����Ա����_In || '����', Null, ����Ա����_In, Null, Null, Null,
                              v_����);
        
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
          Where ��¼id = n_�����¼id And ��� = n_���� And Nvl(�Һ�״̬, 0) In (0, 5);

          Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where NO = No_In;
        
        End Loop;
      
      End If;
    Else
      If Nvl(n_�����¼id, 0) = 0 Then
      
        Update �Һ����״̬
        Set ��� = n_����, ״̬ = 1, �Ǽ�ʱ�� = Sysdate
        Where ���� = v_���� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = n_����;
        If Sql%Rowcount = 0 Then
          Begin
            n_Count := 1;
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_����, Trunc(d_����ʱ��), n_����, 1, ����Ա����_In, Sysdate);
          Exception
            When Others Then
              n_Count := 0;
          End;
          If n_Count = 0 Then
            v_Err_Msg := '���' || n_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
          
          End If;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = n_���� Or ��ע = n_����) And ��¼id = n_�����¼id;
        If Sql%Rowcount = 0 Then
          v_Err_Msg := '���' || n_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  
  --���²�����Ϣ����������Ϣ
  Update ������Ϣ Set ����ʱ�� = d_Date, ����״̬ = 2, �������� = ����_In Where ����id = ����id_In;
  If Sql%NotFound Then
    v_Err_Msg := '����Ĳ�����Ϣ(����ID)��Ч������.';
    Raise Err_Item;
  End If;
  For c_���� In (Select a.����id, a.�����, a.�ѱ�, a.����, a.�Ա�, a.����, a.ҽ�Ƹ��ʽ, b.���� As ҽ�Ƹ��ʽ����
               From ������Ϣ A, ҽ�Ƹ��ʽ B
               Where a.ҽ�Ƹ��ʽ = b.����(+) And a.����id = ����id_In) Loop

    --����������ü�¼
    Update ������ü�¼
    Set ��¼״̬ = 1, ʵ��Ʊ�� = Null, ����id = n_����id, ���ʽ�� = 0,ʵ�ս��=0, ����id = c_����.����id, ��ʶ�� = c_����.�����, ���� = c_����.����, ���� = c_����.����,
        �Ա� = c_����.�Ա�, ���ʽ = c_����.ҽ�Ƹ��ʽ����, �ѱ� = c_����.�ѱ�, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In,
        ����Ա���� = ����Ա����_In, �ɿ���id = n_��id, ��ҩ���� = ����_In, ִ���� = ҽ������_In
    Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;
    --���˹Һż�¼
    Update ���˹Һż�¼
    Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = c_����.����id, ����� = c_����.�����, ����ʱ�� = d_����ʱ��, ���� = c_����.����,
        �Ա� = c_����.�Ա�, ���� = c_����.����, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In),
        ���� = n_����, ���� = ����_In, ִ���� = ҽ������_In, ժҪ = Nvl(ժҪ_In, ժҪ),��¼��־=1,ȡ�ű�־=1
    Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || c_����.���� || '�Ѿ�������';
      Raise Err_Item;
    End If;
  
    --����Ԥ����¼
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������,
       ��������)
    Values
      (n_Ԥ��id, 4, 1, No_In, c_����.����id, '�ֽ�', 0, d_Date, ����Ա���_In, ����Ա����_In, n_����id, '�Һ��շ�', n_��id, Null, Null, Null,
       Null, Null, n_����id, 4);
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(n_���ɶ���, 0) <> 0 Then
    
      For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
      
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, c_����.����, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
        --ԤԼ����ʱ���ı��¼��־
        Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
      
      End Loop;
    End If;
  
  End Loop;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%Rowcount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) >= d_Date;
    End If;
  End If;

  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ԤԼ����_ȡ��;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_�ҺŰ���_��ͳ_Lockno
(
  ��������_In   Integer,
  ����_In       �ҺŰ���.����%Type,
  ����_In       �Һ����״̬.����%Type,
  ����_In       �Һ����״̬.���%Type,
  ���_Out      Out �Һ����״̬.���%Type,
  ������_In     �Һ����״̬.������%Type := Null,
  ����Ա����_In �Һ����״̬.����Ա����%Type := Null,
  ����id_In     �ҺŰ���.Id%Type := Null,
  �ƻ�id_In     �ҺŰ��żƻ�.Id%Type := Null,
  �Ƿ�ԤԼ_In   Number := 0,
  ��ע_In       �Һ����״̬.��ע%Type := Null,
  ������λ_In   Varchar2 := Null,
  ʱ���_In     Varchar2 := Null
  
) Is
  --����:��ͳģʽ�����Ų���
  --��������_In�� 0-����,1-���������ݴ����������������û�ҵ�ȡһ����;2-����(ֱ��ȡһ����Ч�Ž��м���)
  --����ID_In:�����Ϊ�գ���ֱ�ӴӰ�����ȡ��
  --�ƻ�ID_In:�����Ϊ�գ���ֱ�ӴӼƻ���ȡ��
  --ʱ���_In:���Բ����룬������ʱ����ֱ��ȡ��һ����Ч�� ��ʽ:HH24:mi:ss-HH24:mi:ss
  n_����id       �ҺŰ���.Id%Type;
  n_�ƻ�id       �ҺŰ��żƻ�.Id%Type;
  v_����         �ҺŰ���.����%Type;
  n_״̬         �Һ����״̬.״̬%Type;
  v_����         �ҺŰ�������.������Ŀ%Type;
  n_����         �Һ����״̬.���%Type;
  v_��֤����     �Һ����״̬.����Ա����%Type;
  v_����Ա����   �Һ����״̬.����Ա����%Type;
  v_��֤������   �Һ����״̬.������%Type;
  v_������       �Һ����״̬.������%Type;
  n_�޺���       �ҺŰ�������.�޺���%Type;
  n_��Լ��       �ҺŰ�������.��Լ��%Type;
  n_��Լģʽ     Number(3);
  n_��ſ���     Number(3);
  n_��ʱ��       Number(3);
  n_����         Number(18);
  n_���ú�����λ Number(3);
  n_�Ƿ�Һ�     Number(3); --1-�Һ�;0-ԤԼ
  n_������       Number(3);
  d_ʱ�ο�ʼ     Date;
  d_���ʱ��     Date;
  d_ʱ�ν���     Date;
  n_Rowid        Rowid;
  v_Temp         Varchar2(32767); --��ʱXML
  Err_Item Exception;

  Function Check_Nums_Valied
  (
    ����id1_In  In �ҺŰ���.Id%Type,
    �ƻ�id1_In  In �ҺŰ��żƻ�.Id%Type,
    ����1_In    In �ҺŰ�������.������Ŀ%Type,
    �Ƿ�Һ�_In Number
  ) Return Number Is
    --���ܣ�����Ƿ񳬳����޺Ż���Լ
    --���:�Ƿ�Һ�_IN-1:�Һ�;0-ԤԼ
    --����:1-��ʾ���ݺϷ�;0-��ʾ���ݲ��Ϸ�:�������޺Ż���Լ��
    n_Count Number(18);
    n_Temp  Number(18);
  Begin
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Select Max(�޺���), Max(��Լ��)
      Into n_�޺���, n_��Լ��
      From �Һżƻ�����
      Where �ƻ�id = �ƻ�id1_In And ������Ŀ = ����1_In;
    Else
      Select Max(�޺���), Max(��Լ��)
      Into n_�޺���, n_��Լ��
      From �ҺŰ�������
      Where ����id = ����id1_In And ������Ŀ = ����1_In;
    End If;
  
    Select Count(*)
    Into n_Count
    From (Select ���
           From �Һ����״̬
           Where ���� = ����_In And Trunc(����) = Trunc(����_In)
           Union
           Select ���
           From ������λ�ƻ�����
           Where �ƻ�id = Decode(�Ƿ�Һ�_In, 1, 0, �ƻ�id1_In) And Decode(�Ƿ�Һ�_In, 1, 0, 0) = 0 And ������Ŀ = ����1_In And ���� <> 0
           Union
           Select ���
           From ������λ���ſ���
           Where ����id = Decode(�Ƿ�Һ�_In, 1, 0, ����id1_In) And Decode(�Ƿ�Һ�_In, 1, 0, 0) = 0 And ������Ŀ = ����1_In And ���� <> 0);
  
    If �Ƿ�Һ�_In = 1 And Nvl(n_�޺���, 0) <> 0 And Nvl(n_�޺���, 0) < n_Count Then
      Return 0;
    Elsif �Ƿ�Һ�_In = 0 Then
      n_Temp := Nvl(n_��Լ��, 0);
      If n_Temp = 0 Then
        n_Temp := Nvl(n_�޺���, 0);
      End If;
      If n_Temp <> 0 And n_Temp < n_Count Then
        Return 0;
      End If;
    End If;
    Return 1;
  End;

  Function Get_Next_Plannum
  (
    ����1_In       In �ҺŰ���.����%Type,
    ����1_In       In Date,
    ����id1_In     In �ҺŰ���.Id%Type,
    �ƻ�id1_In     In �ҺŰ��żƻ�.Id%Type,
    ����1_In       In �ҺŰ�������.������Ŀ%Type,
    ����Ա����1_In ��Ա��.����%Type,
    ������1_In     �Һ����״̬.������%Type,
    ��ע1_In       In �Һ����״̬.��ע%Type
  ) Return Number Is
    n_Temp_��� Number(18);
    n_Find      Number(2);
    n_������    Number(2);
    d_���ʱ��  Date;
    n_Rowid     Rowid;
  Begin
    If Nvl(�ƻ�id_In, 0) <> 0 Then
      Select Max(���) + 1
      Into n_Temp_���
      From (Select Distinct ���
             From �Һżƻ�ʱ��
             Where �ƻ�id = �ƻ�id1_In And ���� = ����1_In
             Union All
             Select Distinct ���
             From �Һ����״̬
             Where ���� = ����1_In And Trunc(����) = Trunc(����1_In));
    Else
      Select Max(���) + 1
      Into n_Temp_���
      From (Select Distinct ���
             From �ҺŰ���ʱ��
             Where ����id = ����id1_In And ���� = ����1_In
             Union
             Select Distinct ���
             From �Һ����״̬
             Where ���� = ����1_In And Trunc(����) = Trunc(����1_In));
    End If;
  
    n_Find := 0;
    While n_Find = 0 Loop
      Begin
        Select Rowid, 1,
               Case
                 When ������ = ������1_In And ����Ա���� = ����Ա����1_In And ״̬ = 5 Then
                  1
                 Else
                  0
               End
        Into n_Rowid, n_����, n_������
        From �Һ����״̬
        Where ���� = ����1_In And ���� Between Trunc(����1_In) And Trunc(����1_In) + 1 - 1 / 24 / 60 / 60 And ��� = n_Temp_���;
      Exception
        When Others Then
          n_����   := 0;
          n_������ := 0;
      End;
      If Nvl(n_����, 0) = 1 And Nvl(n_������, 0) = 1 Then
        --�Լ����ĺţ���վ��:
        Update �Һ����״̬ Set ״̬ = 5 Where Rowid = n_Rowid;
        n_Find := 1;
        Return n_Temp_���;
      End If;
    
      If Nvl(n_����, 0) = 0 Then
        --δ���ָ���ű�վ�ã������¼
        d_���ʱ�� := ����1_In;
        If ʱ���_In Is Not Null Then
          Begin
            If Nvl(�ƻ�id1_In, 0) <> 0 Then
              Select To_Date(To_Char(����1_In, 'yyyy-mm-dd') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
              Into d_���ʱ��
              From �Һżƻ�ʱ��
              Where �ƻ�id = �ƻ�id1_In And ���� = v_���� And ��� = n_����;
            Else
            
              Select To_Date(To_Char(����1_In, 'yyyy-mm-dd') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
              Into d_���ʱ��
              From �ҺŰ���ʱ��
              Where ����id = ����id1_In And ���� = v_���� And ��� = n_����;
            End If;
          Exception
            When Others Then
              d_���ʱ�� := ����1_In;
          End;
        End If;
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
        Values
          (����1_In, d_���ʱ��, n_Temp_���, 5, ����Ա����1_In, ��ע1_In, Sysdate, ������1_In);
      
        n_Find := 1;
        Return n_Temp_���;
      End If;
      n_Temp_��� := n_Temp_��� + 1;
    End Loop;
  End;

Begin

  v_������ := ������_In;
  If v_������ Is Null Then
    Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  End If;
  v_����Ա���� := ����Ա����_In;
  If v_����Ա���� Is Null Then
    v_����Ա���� := Zl_Username;
  End If;

  n_����     := ����_In;
  v_����     := ����_In;
  n_�Ƿ�Һ� := Case
              When Nvl(�Ƿ�ԤԼ_In, 0) = 0 Then
               1
              Else
               0
            End;

  If ��������_In = 0 Then
    --����
    Delete �Һ����״̬
    Where ������ = v_������ And ����Ա���� = v_����Ա���� And ״̬ = 5 And ��� = n_���� And Trunc(����) = Trunc(����_In) And ���� = ����_In;
    If Sql%NotFound Then
      v_Temp := 'û�з�����Ҫ���������';
      Raise Err_Item;
    End If;
    ���_Out := n_����;
    Return;
  End If;

  --����
  If ʱ���_In Is Not Null Then
    Begin
      d_ʱ�ο�ʼ := To_Date(To_Char(����_In, 'yyyy-mm-dd') || Substr(ʱ���_In, 1, Instr(ʱ���_In, '-') - 1),
                        'yyyy-mm-dd hh24:mi:ss');
      If Substr(ʱ���_In, Instr(ʱ���_In, '-') + 1) Is Null Then
        d_ʱ�ν��� := Null;
      Else
        d_ʱ�ν��� := To_Date(To_Char(����_In, 'yyyy-mm-dd') || Substr(ʱ���_In, Instr(ʱ���_In, '-') + 1),
                          'yyyy-mm-dd hh24:mi:ss');
      End If;
    Exception
      When Others Then
        v_Temp := '�޷����������ʱ��θ�ʽ�����飡';
        Raise Err_Item;
    End;
  End If;

  Select Decode(To_Char(����_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
  Into v_����
  From Dual;

  n_�ƻ�id := Nvl(�ƻ�id_In, 0);
  n_����id := Nvl(����id_In, 0);
  If Nvl(n_�ƻ�id, 0) <> 0 Then
    Select Max(a.��ſ���), Max(b.����)
    Into n_��ſ���, v_����
    From �ҺŰ��żƻ� A, �ҺŰ��� B
    Where a.Id = n_�ƻ�id And a.����id = b.Id;
  End If;
  If Nvl(n_����id, 0) <> 0 Then
    Select Max(��ſ���), Max(����) Into n_��ſ���, v_���� From �ҺŰ��� Where ID = n_����id;
  End If;

  If Nvl(n_�ƻ�id, 0) = 0 And Nvl(n_����id, 0) = 0 Then
    Begin
      Select ��ſ���, ID
      Into n_��ſ���, n_�ƻ�id
      From (Select ��ſ���, ID
             From �ҺŰ��żƻ�
             Where ���� = v_���� And ����_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   ʧЧʱ�� And ���ʱ�� Is Not Null
             Order By ��Чʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select ��ſ���, ID Into n_��ſ���, n_����id From �ҺŰ��� Where ���� = v_����;
    End;
  End If;

  If Nvl(n_��ſ���, 0) = 0 Then
    --δ������ţ���������
    Return;
  End If;

  If Nvl(n_�ƻ�id, 0) <> 0 Then
    Select Nvl(Max(1), 0) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum < 2;
  
    Select Nvl(Max(1), 0)
    Into n_���ú�����λ
    From ������λ�ƻ�����
    Where ������Ŀ = v_���� And �ƻ�id = n_�ƻ�id And ������λ = ������λ_In And Rownum < 2;
  Else
  
    Select Nvl(Max(1), 0) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum < 2;
    Select Nvl(Max(1), 0)
    Into n_���ú�����λ
    From ������λ���ſ���
    Where ������Ŀ = v_���� And ����id = n_����id And ������λ = ������λ_In And Rownum < 2;
  End If;

  If ��������_In = 2 Then
    --ֱ��ȡһ�º����������Ų���
    v_Temp := Zl_Fun_�ҺŰ���_��ͳ_Nextsn(����_In, n_����id, n_�ƻ�id, v_����Ա����, v_����, ��ע_In, v_������, ������λ_In, 0, Nvl(�Ƿ�ԤԼ_In, 0));
    If v_Temp Is Not Null Then
      ���_Out := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    End If;
    Return;
  End If;

  n_���� := 0;
  If ʱ���_In Is Null And ��������_In = 1 Then
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Begin
        Select 1, a.״̬, a.����Ա����, a.������, a.���
        Into n_����, n_״̬, v_��֤����, v_��֤������, n_����
        From �Һ����״̬ A, �Һżƻ�ʱ�� B
        Where a.���� = v_���� And Trunc(a.����) = Trunc(����_In) And a.��� = b.��� And b.�ƻ�id = n_�ƻ�id And b.���� = v_���� And
              To_Char(b.��ʼʱ��, 'hh24:mi') = To_Char(����_In, 'hh24:mi') And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    Else
      Begin
        Select 1, a.״̬, a.����Ա����, a.������, a.���
        Into n_����, n_״̬, v_��֤����, v_��֤������, n_����
        From �Һ����״̬ A, �ҺŰ���ʱ�� B
        Where a.���� = v_���� And Trunc(a.����) = Trunc(����_In) And a.��� = b.��� And b.����id = n_����id And b.���� = v_���� And
              To_Char(b.��ʼʱ��, 'hh24:mi') = To_Char(����_In, 'hh24:mi') And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
  End If;

  If n_���� = 1 Then
    If Not (n_״̬ = 5 And v_��֤���� = v_����Ա���� And v_������ = v_��֤������) Then
      --����ʱ�������Ѿ���ʹ��
      v_Temp := '����ʱ��' || ����_In || '������ѱ�ʹ��';
      Raise Err_Item;
    End If;
    ���_Out := n_����;
    Return;
  End If;

  If n_��ʱ�� = 1 And ��������_In = 1 Then
    If ʱ���_In Is Null Then
      --��ȷ��λ���
      Begin
        n_���� := 1;
        If Nvl(n_�ƻ�id, 0) <> 0 Then
          Select ���
          Into n_����
          From �Һżƻ�ʱ��
          Where �ƻ�id = n_�ƻ�id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi') = To_Char(����_In, 'hh24:mi') And Rownum < 2;
        Else
          Select ���
          Into n_����
          From �ҺŰ���ʱ��
          Where ����id = n_����id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi') = To_Char(����_In, 'hh24:mi') And Rownum < 2;
        End If;
      Exception
        When Others Then
          n_���� := 0;
      End;
    
      If n_���� = 1 Then
        --���ڣ������Ƿ�������վ�á�
        Begin
          Select Rowid, 1,
                 Case
                   When ������ = v_������ And ����Ա���� = v_����Ա���� And ״̬ = 5 Then
                    1
                   Else
                    0
                 End
          Into n_Rowid, n_����, n_������
          From �Һ����״̬
          Where ���� = v_���� And ���� = ����_In And ��� = n_����;
        Exception
          When Others Then
            n_����   := 0;
            n_������ := 0;
        End;
      
        If Nvl(n_����, 0) = 1 And Nvl(n_������, 0) = 1 Then
          --�Լ����ĺţ���վ��:
          Update �Һ����״̬ Set ״̬ = 5 Where Rowid = n_Rowid;
          ���_Out := n_����;
          Return;
        End If;
        If Nvl(n_����, 0) = 1 And Nvl(n_������, 0) = 0 Then
          v_Temp := '����ʱ��' || ����_In || '������ѱ�ʹ��';
          Raise Err_Item;
        End If;
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
        Values
          (v_����, ����_In, n_����, 5, v_����Ա����, ��ע_In, Sysdate, v_������);
        ���_Out := n_����;
        Return;
      End If;
      --������ʱ��ȡ��һ����,ͬʱ����޺����Ƿ���ȷ
      n_���� := Get_Next_Plannum(v_����, ����_In, n_����id, n_�ƻ�id, v_����, v_����Ա����, v_������, ��ע_In);
    
      If Check_Nums_Valied(n_����id, n_�ƻ�id, v_����, n_�Ƿ�Һ�) = 0 Then
        v_Temp := '����ű�' || ����_In || '��ǰ�������';
        Raise Err_Item;
      End If;
      ���_Out := n_����;
      Return;
    Else
      If Nvl(n_�ƻ�id, 0) <> 0 Then
        --ԤԼ�ж�: ���ȫ����ſ���ԤԼ�����Ƿ�ԤԼȫ��Ϊ0��������Ҫ���������������
      
        Select Nvl(Max(1), 0)
        Into n_����
        From �Һżƻ�ʱ��
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And �Ƿ�ԤԼ = 1 And Rownum < 2;
      
        Select Min(a.���), Min(b.״̬)
        Into n_����, n_״̬
        From �Һżƻ�ʱ�� A, �Һ����״̬ B
        Where a.��� = b.���(+) And (Nvl(b.״̬, 0) = 0 Or Nvl(b.״̬, 0) = 5) And b.����(+) = v_���� And
              Trunc(b.����(+)) = Trunc(����_In) And a.�ƻ�id = n_�ƻ�id And a.���� = v_���� And
              To_Char(a.��ʼʱ��, 'hh24:mi') >= To_Char(d_ʱ�ο�ʼ, 'hh24:mi') And
              To_Char(a.��ʼʱ��, 'hh24:mi') < To_Char(d_ʱ�ν���, 'hh24:mi') And Case
                When n_�Ƿ�Һ� = 1 Or n_���� = 0 Then
                 1
                Else
                 a.�Ƿ�ԤԼ
              End = 1; --  Decode(n_�Ƿ�Һ�, 1, 1, Decode(n_����, 1, a.�Ƿ�ԤԼ, 1))) = 1;
      Else
        Select Nvl(Max(1), 0)
        Into n_����
        From �ҺŰ���ʱ��
        Where ����id = n_����id And ���� = v_���� And �Ƿ�ԤԼ = 1 And Rownum < 2;
      
        Select Min(a.���), Min(b.״̬)
        Into n_����, n_״̬
        From �ҺŰ���ʱ�� A, �Һ����״̬ B
        Where a.��� = b.���(+) And (Nvl(b.״̬, 0) = 0 Or Nvl(b.״̬, 0) = 5) And b.����(+) = v_���� And
              Trunc(b.����(+)) = Trunc(����_In) And a.����id = n_����id And a.���� = v_���� And
              To_Char(a.��ʼʱ��, 'hh24:mi') >= To_Char(d_ʱ�ο�ʼ, 'hh24:mi') And
              To_Char(a.��ʼʱ��, 'hh24:mi') < To_Char(d_ʱ�ν���, 'hh24:mi') And Case
                When n_�Ƿ�Һ� = 1 Or n_���� = 0 Then
                 1
                Else
                 a.�Ƿ�ԤԼ
              End = 1; --  Decode(n_�Ƿ�Һ�, 1, 1, Decode(n_����, 1, a.
      
      End If;
    
      If Nvl(n_����, 0) = 0 Then
        If n_���� = 1 Then
          v_Temp := '����ʱ���' || ʱ���_In || '������ѱ�ʹ�û�δ����ԤԼ��';
          Raise Err_Item;
        End If;
      
        If d_ʱ�ν��� Is Not Null Then
          v_Temp := '����ʱ���' || ʱ���_In || '������ѱ�ʹ��';
          Raise Err_Item;
        End If;
        --������ʱ��ȡ��һ����,ͬʱ����޺����Ƿ���ȷ
        n_���� := Get_Next_Plannum(v_����, ����_In, n_����id, n_�ƻ�id, v_����, v_����Ա����, v_������, ��ע_In);
      
        If Check_Nums_Valied(n_����id, n_�ƻ�id, v_����, n_�Ƿ�Һ�) = 0 Then
          v_Temp := '����ű�' || v_���� || '��ǰ�������';
          Raise Err_Item;
        End If;
        ���_Out := n_����;
        Return;
      End If;
      --�������
      If Nvl(n_״̬, 0) = 0 Then
        --�Ϸ�ʱ��Σ������¼
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
        Values
          (v_����, ����_In, n_����, 5, v_����Ա����, ��ע_In, Sysdate, v_������);
        ���_Out := n_����;
        Return;
      
      End If;
    
      If Nvl(n_״̬, 0) = 5 Then
        Select Nvl(Max(1), 0)
        Into n_����
        From �Һ����״̬
        Where ������ = v_������ And ����Ա���� = v_����Ա���� And ��� = n_���� And ���� = v_����;
      
        If n_���� = 0 Then
          v_Temp := '����ʱ���' || ʱ���_In || '������ѱ�ʹ��';
          Raise Err_Item;
        End If;
        ���_Out := n_����;
        Return;
      End If;
    
      If d_ʱ�ν��� Is Not Null Then
        v_Temp := '����ʱ���' || ʱ���_In || '������ѱ�ʹ��';
        Raise Err_Item;
      End If;
      --������ʱ��ȡ��һ����,ͬʱ����޺����Ƿ���ȷ
      n_���� := Get_Next_Plannum(v_����, ����_In, n_����id, n_�ƻ�id, v_����, v_����Ա����, v_������, ��ע_In);
      If Check_Nums_Valied(n_����id, n_�ƻ�id, v_����, n_�Ƿ�Һ�) = 0 Then
        v_Temp := '����ű�' || ����_In || '��ǰ�������';
        Raise Err_Item;
      End If;
    End If;
    ���_Out := n_����;
    Return;
  End If;

  --����ʱ��,����������ŵ�
  If Nvl(n_�ƻ�id, 0) <> 0 Then
    Select Decode(n_�Ƿ�Һ�, 1, Max(�޺���), Max(��Լ��))
    Into n_�޺���
    From �Һżƻ�����
    Where �ƻ�id = n_�ƻ�id And ������Ŀ = v_����;
  Else
    Select Decode(n_�Ƿ�Һ�, 1, Max(�޺���), Max(��Լ��))
    Into n_�޺���
    From �ҺŰ�������
    Where ����id = n_����id And ������Ŀ = v_����;
  End If;

  n_���� := 1;
  If ������λ_In Is Null Or n_���ú�����λ = 0 Then
    For r_��� In (Select ���, ״̬, ����Ա����, ������
                 From �Һ����״̬
                 Where ���� = ����_In And Trunc(����) = Trunc(����_In)
                 Union
                 Select ���, Null, Null, Null
                 From ������λ�ƻ�����
                 Where �ƻ�id = Decode(n_�Ƿ�Һ�, 1, 0, n_�ƻ�id) And Decode(n_�Ƿ�Һ�, 1, 1, 0) = 0 And ������Ŀ = v_���� And ���� <> 0
                 Union
                 Select ���, Null, Null, Null
                 From ������λ���ſ���
                 Where ����id = Decode(n_�Ƿ�Һ�, 1, 0, n_����id) And Decode(n_�Ƿ�Һ�, 1, 1, 0) = 0 And ������Ŀ = v_���� And ���� <> 0
                 Order By ���) Loop
      --�������ŵģ����˳�
      Exit When r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������;
      If r_���.��� = n_���� Then
        n_���� := n_���� + 1;
      End If;
    End Loop;
  
    If n_���� > n_�޺��� Then
      v_Temp := '����ű�' || ����_In || '��ǰ�������';
      Raise Err_Item;
    End If;
  
    Begin
      Select Rowid, 1,
             Case
               When ������ = v_������ And ����Ա���� = v_����Ա���� And ״̬ = 5 Then
                1
               Else
                0
             End
      Into n_Rowid, n_����, n_������
      From �Һ����״̬
      Where ���� = ����_In And Trunc(����) = Trunc(����_In) And ��� = n_����;
    Exception
      When Others Then
        n_����   := 0;
        n_������ := 0;
    End;
    If n_���� = 1 And n_������ = 1 Then
      ���_Out := n_����;
      Return;
    End If;
    If n_���� = 1 And n_������ = 0 Then
      --�Ѿ�վ����
      v_Temp := '���' || n_���� || '�ѱ�ʹ��';
      Raise Err_Item;
    End If;
    Insert Into �Һ����״̬
      (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
    Values
      (����_In, Trunc(����_In), n_����, 5, v_����Ա����, ��ע_In, Sysdate, v_������);
    ���_Out := n_����;
    Return;
  End If;

  --�����˺�����λ���Ƶ�:��Լ��λ����
  If Nvl(n_�ƻ�id, 0) <> 0 Then
    Select Count(1)
    Into n_��Լģʽ
    From ������λ�ƻ�����
    Where ��� = 0 And �ƻ�id = n_�ƻ�id And ������λ = ������λ_In And ������Ŀ = v_���� And ���� <> 0;
  Else
    Select Count(1)
    Into n_��Լģʽ
    From ������λ���ſ���
    Where ��� = 0 And ����id = n_����id And ������λ = ������λ_In And ������Ŀ = v_���� And ���� <> 0;
  End If;

  If n_��Լģʽ = 0 Then
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Select Nvl(Max(���), 0)
      Into n_����
      From (Select ���
             From ������λ�ƻ����� A
             Where �ƻ�id = n_�ƻ�id And ������λ = ������λ_In And ������Ŀ = v_���� And ���� <> 0 And
                   (Not Exists
                    (Select 1
                     From �Һ����״̬
                     Where ���� = ����_In And ��� = a.��� And Trunc(����) = Trunc(����_In) And ״̬ <> 0) Or Exists
                    (Select 1
                     From �Һ����״̬
                     Where ���� = ����_In And ��� = a.��� And Trunc(����) = Trunc(����_In) And ״̬ = 5 And ����Ա���� = v_����Ա���� And
                           ������ = v_������))
             Order By ���)
      Where Rownum < 2;
    Else
    
      Select Nvl(Max(���), 0)
      Into n_����
      From (Select ���
             From ������λ���ſ��� A
             Where ����id = n_����id And ������λ = ������λ_In And ������Ŀ = v_���� And ���� <> 0 And
                   (Not Exists
                    (Select 1
                     From �Һ����״̬
                     Where ���� = ����_In And ��� = a.��� And Trunc(����) = Trunc(����_In) And ״̬ <> 0) Or Exists
                    (Select 1
                     From �Һ����״̬
                     Where ���� = ����_In And ��� = a.��� And Trunc(����) = Trunc(����_In) And ״̬ = 5 And ����Ա���� = v_����Ա���� And
                           ������ = v_������))
             Order By ���)
      Where Rownum < 2;
    End If;
  
    If Nvl(n_����, 0) = 0 Then
      v_Temp := '����ű�' || ����_In || '��ǰ�������';
      Raise Err_Item;
    End If;
    Begin
      Select Rowid, 1,
             Case
               When ������ = v_������ And ����Ա���� = v_����Ա���� And ״̬ = 5 Then
                1
               Else
                0
             End
      Into n_Rowid, n_����, n_������
      From �Һ����״̬
      Where ���� = ����_In And ���� Between Trunc(����_In) And Trunc(����_In) + 1 - 1 / 24 / 60 / 60 And ��� = n_����;
    Exception
      When Others Then
        n_����   := 0;
        n_������ := 0;
    End;
    If n_���� = 1 And n_������ = 0 Then
      v_Temp := '���Ϊ' || n_���� || '�ѱ�ʹ��';
      Raise Err_Item;
    End If;
    If Nvl(n_����, 0) = 0 Then
      Insert Into �Һ����״̬
        (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
      Values
        (����_In, Trunc(����_In), n_����, 5, v_����Ա����, ��ע_In, Sysdate, v_������);
    End If;
    ���_Out := n_����;
    Return;
  
  Else
    n_���� := 1;
    Select �޺��� Into n_�޺��� From �Һżƻ����� Where �ƻ�id = n_�ƻ�id And ������Ŀ = v_����;
    For r_��� In (Select ���, ״̬, ����Ա����, ������
                 From �Һ����״̬
                 Where ���� = ����_In And Trunc(����) = Trunc(����_In)
                 Union All
                 Select ���, Null, Null, Null
                 From ������λ�ƻ�����
                 Where �ƻ�id = n_�ƻ�id And Decode(Nvl(n_�ƻ�id, 0), 0, 0, 1) = 1 And ������Ŀ = v_���� And ���� <> 0
                 Union All
                 Select ���, Null, Null, Null
                 From ������λ���ſ���
                 Where ����id = n_����id And Decode(Nvl(n_�ƻ�id, 0), 0, 1, 0) = 1 And ������Ŀ = v_���� And ���� <> 0
                 Order By ���) Loop
      If r_���.״̬ = 5 And r_���.����Ա���� = v_����Ա���� And r_���.������ = v_������ Then
        n_���� := r_���.���;
        Exit;
      End If;
      If r_���.��� = n_���� Then
        n_���� := n_���� + 1;
      End If;
    End Loop;
  
    If n_���� > n_�޺��� Then
      v_Temp := '����ű�' || ����_In || '��ǰ�������';
      Raise Err_Item;
    End If;
  
    Begin
      Select Rowid, 1,
             Case
               When ������ = v_������ And ����Ա���� = v_����Ա���� And ״̬ = 5 Then
                1
               Else
                0
             End
      Into n_Rowid, n_����, n_������
      From �Һ����״̬
      Where ���� = ����_In And ���� Between Trunc(����_In) And Trunc(����_In) + 1 - 1 / 24 / 60 / 60 And ��� = n_����;
    Exception
      When Others Then
        n_����   := 0;
        n_������ := 0;
    End;
    If n_���� = 1 And n_������ = 0 Then
      v_Temp := '���Ϊ' || n_���� || '�ѱ�ʹ��';
      Raise Err_Item;
    End If;
  
    If Nvl(n_����, 0) = 0 Then
      Insert Into �Һ����״̬
        (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
      Values
        (����_In, Trunc(����_In), n_����, 5, v_����Ա����, ��ע_In, Sysdate, v_������);
    End If;
    ���_Out := n_����;
    Return;
  End If;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Temp || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�ҺŰ���_��ͳ_Lockno;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_�����Һ�_Turn(��������_In In Date) Is
  ------------------------------------------------ 
  --���ܣ���������Ű�ģʽ����ʱ��֮��ļƻ��Ű�Һż�¼ת��Ϊ������Ű�ģʽ�Һż�¼
  --���أ�ת����¼����
  ------------------------------------------------ 
  v_Error Varchar2(255);
  Err_Custom Exception;
  n_��������   Number(10);
  v_Para       Varchar2(500);
  n_�Һ�ģʽ   Number(3);
  d_����ʱ��   Date;
  n_����id     �ҺŰ���.Id%Type;
  n_�ƻ�id     �ҺŰ��żƻ�.Id%Type;
  v_ʱ���     ʱ���.ʱ���%Type;
  n_��ʱ��     Number(3);
  n_��ʱ����� �ٴ�������ſ���.���%Type;
  n_��ſ���   �ҺŰ���.��ſ���%Type;
  n_�����¼id �ٴ������¼.Id%Type;
  n_δ�������� Number(10);
Begin
  Begin
    d_����ʱ�� := ��������_In;
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  For r_�Һ� In (Select ID, NO, �ű�, ִ�в���id, ִ����, ��¼����, ԤԼ, ����, ����id, ����ʱ��, ����Ա����
               From ���˹Һż�¼
               Where ��¼״̬ = 1 And ����ʱ�� >= Trunc(d_����ʱ��) And �����¼id Is Null) Loop
    v_ʱ��� := Null;
    Begin
      Select ID, ��ſ���
      Into n_�ƻ�id, n_��ſ���
      From (Select a.Id, a.��ſ���
             From �ҺŰ��żƻ� A, �ҺŰ��� B
             Where a.����id = b.Id And a.���ʱ�� Is Not Null And b.���� = r_�Һ�.�ű� And r_�Һ�.����ʱ�� Between ��Чʱ�� + 0 And ʧЧʱ��
             Order By ��Чʱ�� Desc)
      Where Rownum < 2;
      Select Decode(To_Char(r_�Һ�.����ʱ��, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7', a.����,
                     Null)
      Into v_ʱ���
      From �ҺŰ��żƻ� A
      Where a.Id = n_�ƻ�id;
    Exception
      When Others Then
        n_�ƻ�id := Null;
        Select ID, ��ſ��� Into n_����id, n_��ſ��� From �ҺŰ��� Where ���� = r_�Һ�.�ű�;
        Select Decode(To_Char(r_�Һ�.����ʱ��, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7',
                       a.����, Null)
        Into v_ʱ���
        From �ҺŰ��� A
        Where a.Id = n_����id;
    End;
    If v_ʱ��� Is Not Null Then
      Begin
        If Nvl(n_�ƻ�id, 0) = 0 Then
          Select 1 Into n_��ʱ�� From �ҺŰ���ʱ�� Where ����id = n_����id And Rownum < 2;
        Else
          Select 1 Into n_��ʱ�� From �Һżƻ�ʱ�� Where �ƻ�id = n_�ƻ�id And Rownum < 2;
        End If;
      Exception
        When Others Then
          n_��ʱ�� := 0;
      End;
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = r_�Һ�.�ű� And r_�Һ�.����ʱ�� Between a.��ʼʱ�� And a.��ֹʱ�� And �ϰ�ʱ�� = v_ʱ���;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
      If n_�����¼id Is Null Then
        v_Error := '�Ѿ��Һŵļ�¼�����޷���Ӧ�ĳ����¼,��������ʧ��!';
        Raise Err_Custom;
      End If;
      If Nvl(n_��ſ���, 0) = 0 Then
        If n_��ʱ�� = 0 Then
          If r_�Һ�.��¼���� = 1 Then
            If Nvl(r_�Һ�.ԤԼ, 0) = 1 Then
              Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
              Update �ٴ������¼
              Set �ѹ��� = �ѹ��� + 1, ��Լ�� = ��Լ�� + 1, �����ѽ��� = �����ѽ��� + 1
              Where ID = n_�����¼id;
            Else
              Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
              Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
            End If;
          Else
            Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
            Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
          End If;
          n_�������� := n_�������� + 1;
        Else
          --����ſ��Ʒ�ʱ��,���⴦��
          Select ��� Into n_��ʱ����� From �ٴ�������ſ��� Where ԤԼ˳��� Is Null And ��ʼʱ�� = r_�Һ�.����ʱ��;
          If r_�Һ�.��¼���� = 1 Then
            If Nvl(r_�Һ�.ԤԼ, 0) = 1 Then
              Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
              Update �ٴ������¼
              Set �ѹ��� = �ѹ��� + 1, ��Լ�� = ��Լ�� + 1, �����ѽ��� = �����ѽ��� + 1
              Where ID = n_�����¼id;
            Else
              Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
              Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
            End If;
            Insert Into �ٴ�������ſ���
              (��¼id, ���, ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����Ա����, ��ע)
              Select ��¼id, n_��ʱ�����, r_�Һ�.����, ��ʼʱ��, ��ֹʱ��, 1, �Ƿ�ԤԼ, 1, r_�Һ�.����Ա����, r_�Һ�.����
              From �ٴ�������ſ���
              Where ��¼id = n_�����¼id And ��� = n_��ʱ����� And ԤԼ˳��� Is Null;
          Else
            Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
            Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
            Insert Into �ٴ�������ſ���
              (��¼id, ���, ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����Ա����, ��ע)
              Select ��¼id, n_��ʱ�����, r_�Һ�.����, ��ʼʱ��, ��ֹʱ��, 1, �Ƿ�ԤԼ, 2, r_�Һ�.����Ա����, r_�Һ�.����
              From �ٴ�������ſ���
              Where ��¼id = n_�����¼id And ��� = n_��ʱ����� And ԤԼ˳��� Is Null;
          End If;
          n_�������� := n_�������� + 1;
        End If;
      Else
        If r_�Һ�.��¼���� = 1 Then
          If Nvl(r_�Һ�.ԤԼ, 0) = 1 Then
            Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
            Update �ٴ������¼
            Set �ѹ��� = �ѹ��� + 1, ��Լ�� = ��Լ�� + 1, �����ѽ��� = �����ѽ��� + 1
            Where ID = n_�����¼id;
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = r_�Һ�.����Ա����
            Where ��¼id = n_�����¼id And ��� = r_�Һ�.����;
            If Sql%RowCount = 0 Then
              Insert Into �ٴ�������ſ���
                (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Һ�״̬, ����Ա����, ��ע)
              Values
                (n_�����¼id, r_�Һ�.����, r_�Һ�.����ʱ��, r_�Һ�.����ʱ��, 1, 1, r_�Һ�.����Ա����, '�Զ�ת���������');
            End If;
          Else
            Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
            Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = r_�Һ�.����Ա����
            Where ��¼id = n_�����¼id And ��� = r_�Һ�.����;
            If Sql%RowCount = 0 Then
              Insert Into �ٴ�������ſ���
                (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Һ�״̬, ����Ա����, ��ע)
              Values
                (n_�����¼id, r_�Һ�.����, r_�Һ�.����ʱ��, r_�Һ�.����ʱ��, 1, 1, r_�Һ�.����Ա����, '�Զ�ת���������');
            End If;
          End If;
        Else
          Update ���˹Һż�¼ Set �����¼id = n_�����¼id Where ID = r_�Һ�.Id;
          Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 2, ����Ա���� = r_�Һ�.����Ա����
          Where ��¼id = n_�����¼id And ��� = r_�Һ�.����;
          If Sql%RowCount = 0 Then
            Insert Into �ٴ�������ſ���
              (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Һ�״̬, ����Ա����, ��ע)
            Values
              (n_�����¼id, r_�Һ�.����, r_�Һ�.����ʱ��, r_�Һ�.����ʱ��, 1, 2, r_�Һ�.����Ա����, '�Զ�ת���������');
          End If;
        End If;
        n_�������� := n_�������� + 1;
      End If;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����Һ�_Turn;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_�ٴ������_����
(
  ����_In �ҺŰ���.����%Type := Null,
  ��ʼ_In Number := 1
) As
  -------------------------------------------------------------------------
  --����˵���������ٴ������,��Ҫ�Ǹ��ݹҺŰ��ţ��Һżƻ����ŵȱ�������ݵ���
  --��Σ�
  --    ��ʼ_In:�������ʱ��Ч����ʾ��һ��
  -------------------------------------------------------------------------
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  l_����id t_Numlist := t_Numlist();
  n_Count  Number(18);

  v_ʱ���           Varchar2(4000);
  n_����id           �ٴ������.Id%Type;
  v_ȫԺ��Դ����վ�� ���ű�.վ��%Type;

  Procedure Zl_Register_Import
  (
    ����_In             �ҺŰ���.����%Type,
    ����id_In           �ٴ������.Id%Type,
    ȫԺ��Դ����վ��_In ���ű�.վ��%Type
  ) As
    n_��Դid   �ٴ������Դ.Id%Type;
    d_����ʱ�� �ٴ������Դ.����ʱ��%Type;
  
    n_����id �ٴ������.Id%Type;
    n_����id �ٴ����ﰲ��.Id%Type;
  
    n_����id �ٴ������Դ����.Id%Type;
    n_����id ��������.Id%Type;
  
    n_�Ƿ���     Number(2);
    n_�Ƿ���ʱ���� �ٴ����ﰲ��.�Ƿ���ʱ����%Type;
  
    n_Count  Number(18);
    l_����id t_Numlist := t_Numlist();
  Begin
    For c_��Դ In (Select a.Id, a.����, a.����, a.����id, a.��Ŀid, a.ҽ������, Decode(a.ҽ��id, 0, Null, a.ҽ��id) As ҽ��id, a.���, a.����,
                        a.��һ, a.�ܶ�, a.����, a.����, a.����, a.����, a.��������, a.���﷽ʽ, a.��ſ���, a.��ʼʱ��, a.��ֹʱ��, a.ִ��ʱ��, a.ִ�мƻ�id,
                        a.Ĭ��ʱ�μ��, a.ԤԼ����, Nvl(a.�Ƿ�ɾ��, 0) As �Ƿ�ɾ��,
                        Nvl(a.ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) As ͣ������, Nvl(b.վ��, ȫԺ��Դ����վ��_In) As վ��
                 From �ҺŰ��� A, ���ű� B
                 Where a.����id = b.Id And a.���� = ����_In
                      --���ң���Ŀ��ҽ����ͬ���Ѿ�������һ���ű�Ͳ��ٵ���
                       And Not Exists (Select 1
                        From �ٴ������Դ
                        Where ����id = a.����id And ��Ŀid = a.��Ŀid And Nvl(ҽ������, '-') = Nvl(a.ҽ������, '-') And
                              Nvl(ҽ��id, 0) = Nvl(a.ҽ��id, 0))) Loop
    
      n_�Ƿ��� := 1;
      --���ڿ��ң���Ŀ��ҽ�����߶���ͬ�Ķ���ű����ȿ��ǵ�����Ч�ű��еĵ�һ�������û�У�����ʧЧ�ű��еĵ�һ��
      Select Count(1)
      Into n_Count
      From �ҺŰ���
      Where ����id = c_��Դ.����id And ��Ŀid = c_��Դ.��Ŀid And Nvl(ҽ������, '-') = Nvl(c_��Դ.ҽ������, '-') And
            Nvl(ҽ��id, 0) = Nvl(c_��Դ.ҽ��id, 0);
      If Nvl(n_Count, 0) = 1 Then
        --���ң���Ŀ��ҽ����Ψһ��
        n_�Ƿ��� := 1;
      Else
        --�Ƿ����δͣ����δɾ���ĺű�
        Select Count(1)
        Into n_Count
        From �ҺŰ���
        Where ����id = c_��Դ.����id And ��Ŀid = c_��Դ.��Ŀid And Nvl(ҽ������, '-') = Nvl(c_��Դ.ҽ������, '-') And
              Nvl(ҽ��id, 0) = Nvl(c_��Դ.ҽ��id, 0) And c_��Դ.�Ƿ�ɾ�� = 0 And
              (ͣ������ Is Null Or ͣ������ = To_Date('3000-01-01', 'yyyy-mm-dd'));
        If Nvl(n_Count, 0) = 0 Then
          --������δͣ����δɾ���ĺű�ֱ�ӵ��뵱ǰ�ű𣬼�ʧЧ�ű��еĵ�һ��
          n_�Ƿ��� := 1;
        Elsif Nvl(n_Count, 0) = 1 Then
          --ֻ����һ��δͣ����δɾ���ĺű𣬼���ǲ��ǵ�ǰ�ű�
          If c_��Դ.�Ƿ�ɾ�� = 0 And c_��Դ.ͣ������ = To_Date('3000-01-01', 'yyyy-mm-dd') Then
            n_�Ƿ��� := 1;
          Else
            n_�Ƿ��� := 0;
          End If;
        Else
          --��鵱ǰ�ű��Ƿ���ͣ�û���ɾ��
          If Not (c_��Դ.�Ƿ�ɾ�� = 0 And c_��Դ.ͣ������ = To_Date('3000-01-01', 'yyyy-mm-dd')) Then
            --��ͣ�û���ɾ���򲻵���
            n_�Ƿ��� := 0;
          Else
            --��ǰ�ű���/�ƻ��Ƿ���Ч
            Select Count(1)
            Into n_Count
            From �ҺŰ��żƻ�
            Where ����id = c_��Դ.Id And ���ʱ�� Is Not Null And ʧЧʱ�� > Sysdate And
                  Rownum < 2;
            If Nvl(n_Count, 0) = 0 Then
              --�޼ƻ�
              Select Count(1)
              Into n_Count
              From �ҺŰ��� A
              Where a.Id = c_��Դ.Id And Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And
                    Not (a.���� Is Null And a.��һ Is Null And a.�ܶ� Is Null And a.���� Is Null And a.���� Is Null And
                     a.���� Is Null And a.���� Is Null);
            Else
              --ֻҪ��Чʱ����ڵ�ǰʱ����߲����ڴ�������Чʱ��С�ڵ�ǰʱ��Ķ�����Ч�ģ�
              --��Ϊ1.��Чʱ��ͺ�����Ψһ�ģ�2.������Чʱ���������ȷ������Ч�ģ�
              --     �������ǰ�ƻ�����Чʱ��С�ڵ��ڵ�ǰʱ������������Чʱ��С�ڵ��ڵ�ǰʱ��ļƻ�����Чʱ�����ģ���ǰ�ƻ�����Ч��



              Select Count(1)
              Into n_Count
              From �ҺŰ��żƻ� A
              Where a.���ʱ�� Is Not Null And a.ʧЧʱ�� > Sysdate And
                    a.����id = c_��Դ.Id And
                    (Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) >= Sysdate Or
                    Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And Not Exists
                     (Select 1
                      From �ҺŰ��żƻ�
                      Where ����id = a.����id And ���ʱ�� Is Not Null And
                            ʧЧʱ�� > Sysdate And
                            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And
                            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) >
                            Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')))) And
                    Not (a.���� Is Null And a.��һ Is Null And a.�ܶ� Is Null And a.���� Is Null And a.���� Is Null And
                     a.���� Is Null And a.���� Is Null);
            End If;
          
            If Nvl(n_Count, 0) <> 0 Then
              --��ǰ�ű�����Ч
              n_�Ƿ��� := 1;
            Else
              --��ǰ�ű�����Ч
              n_�Ƿ��� := 0;
            End If;
          End If;
        End If;
      End If;
    
      If Nvl(n_�Ƿ���, 0) = 1 Then
        Select �ٴ������Դ_Id.Nextval Into n_��Դid From Dual;
      
        Select Nvl(Min(��ʼʱ��), Sysdate)
        Into d_����ʱ��
        From (Select Min(��ʼʱ��) As ��ʼʱ��
               From �ҺŰ���
               Where ID = c_��Դ.Id
               Union All
               Select Min(��Чʱ��) As ��ʼʱ��
               From �ҺŰ��żƻ�
               Where ���ʱ�� Is Not Null And ʧЧʱ�� > Sysdate And ����id = c_��Դ.Id);
      
        --1.�����ٴ������Դ
        Insert Into �ٴ������Դ
          (ID, ����, ����, ����id, ��Ŀid, ҽ��id, ҽ������, �Ƿ񽨲���, ԤԼ����, ����Ƶ��, ���տ���״̬, �Ƿ��ٴ��Ű�, �Ű෽ʽ, �Ƿ�ɾ��, ����ʱ��, ����ʱ��)
        Values
          (n_��Դid, c_��Դ.����, c_��Դ.����, c_��Դ.����id, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, c_��Դ.��������, c_��Դ.ԤԼ����, c_��Դ.Ĭ��ʱ�μ��, 2,
           0, 0, c_��Դ.�Ƿ�ɾ��, d_����ʱ��, c_��Դ.ͣ������);
      
        --2.�����ٴ�����ͣ���¼
        --һ��ҽ��һ��ͣ��ƻ�ֻ����һ�������ܴ���һ��ҽ������ű����������ǵ�ͣ��ƻ�һ��
        Insert Into �ٴ�����ͣ���¼
          (ID, ��¼id, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ������, ����ʱ��, ������, ����ʱ��, �Ǽ���)
          Select �ٴ�����ͣ���¼_Id.Nextval, Null, a.��ʼֹͣʱ��, a.����ֹͣʱ��, a.��ע, b.ҽ������, a.�ƶ�����, a.�ƶ���, a.�ƶ�����, a.�ƶ���
          From �ҺŰ���ͣ��״̬ A, �ҺŰ��� B
          Where a.����id = b.Id And b.Id = c_��Դ.Id And b.ҽ��id Is Not Null And Not Exists
           (Select 1
                 From �ٴ�����ͣ���¼
                 Where ��¼id Is Null And ������ = b.ҽ������ And ��ʼʱ�� = a.��ʼֹͣʱ�� And ��ֹʱ�� = a.����ֹͣʱ��);
      
        --3.������صĳ��������
        --3.1 �̶������
        If c_��Դ.վ�� Is Null Then
          n_����id := ����id_In;
        Else
          Begin
            Select ID Into n_����id From �ٴ������ Where �Ű෽ʽ = 0 And Nvl(վ��, '-') = c_��Դ.վ��;
          Exception
            When Others Then
              n_����id := 0;
          End;
          If n_����id = 0 Then
            Update �ٴ������
            Set վ�� = c_��Դ.վ��
            Where �Ű෽ʽ = 0 And Nvl(վ��, '-') = '-'
            Returning ID Into n_����id;
            If Sql%NotFound Then
              Select �ٴ������_Id.Nextval Into n_����id From Dual;
              Insert Into �ٴ������
                (ID, �Ű෽ʽ, �������, ���, վ��)
              Values
                (n_����id, 0, '�̶������', To_Number(To_Char(Sysdate, 'yyyy')), c_��Դ.վ��);
            End If;
          End If;
        End If;
      
        --3.2�����ٴ����ﰲ��
        --ʧЧ�İ��źͼƻ�������
        --ֻҪ��Чʱ����ڵ�ǰʱ����߲����ڴ�������Чʱ��С�ڵ�ǰʱ��Ķ�����Ч�ģ�
        --��Ϊ1.��Чʱ��ͺ�����Ψһ�ģ�2.������Чʱ���������ȷ������Ч�ģ�
        --     �������ǰ�ƻ�����Чʱ��С�ڵ��ڵ�ǰʱ������������Чʱ��С�ڵ��ڵ�ǰʱ��ļƻ�����Чʱ�����ģ���ǰ�ƻ�����Ч��



        For c_���� In (
                     --1.�޼ƻ��İ���
                     Select a.Id As ����id, -1 * Null As �ƻ�id, a.����id, a.��Ŀid, a.ҽ������,
                             Decode(a.ҽ��id, 0, Null, a.ҽ��id) As ҽ��id, a.����, a.��һ, a.�ܶ�, a.����, a.����, a.����, a.����, a.���﷽ʽ,
                             a.��ſ���, Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) As ��ʼʱ��,
                             Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��
                     From �ҺŰ��� A
                     Where a.Id = c_��Դ.Id And Not Exists (Select 1
                            From �ҺŰ��żƻ�
                            Where ����id = a.Id And ���ʱ�� Is Not Null And
                                  ʧЧʱ�� > Sysdate) And
                           Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate
                     Union All
                     --�мƻ��İ���,ֻ������Ч��
                     Select a.����id, a.Id As �ƻ�id, b.����id, a.��Ŀid, a.ҽ������, Decode(a.ҽ��id, 0, Null, a.ҽ��id) As ҽ��id, a.����,
                            a.��һ, a.�ܶ�, a.����, a.����, a.����, a.����, a.���﷽ʽ, a.��ſ���,
                            Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) As ��ʼʱ��,
                            Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��
                     From �ҺŰ��żƻ� A, �ҺŰ��� B
                     Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                           a.ʧЧʱ�� > Sysdate And b.Id = c_��Դ.Id And
                           (Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) >= Sysdate Or
                           Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And Not Exists
                            (Select 1
                             From �ҺŰ��żƻ�
                             Where ����id = a.����id And ���ʱ�� Is Not Null And
                                   ʧЧʱ�� > Sysdate And
                                   Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) < Sysdate And
                                   Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) >
                                   Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd'))))) Loop
        
          Select �ٴ����ﰲ��_Id.Nextval Into n_����id From Dual;
        
          n_����id := Null;
          If Nvl(c_����.���﷽ʽ, 0) = 1 Then
            Begin
              If Nvl(c_����.�ƻ�id, 0) <> 0 Then
                Select a.Id
                Into n_����id
                From �������� A, �Һżƻ����� B
                Where a.���� = b.�������� And b.�ƻ�id = c_����.�ƻ�id And Rownum < 2;
              Else
                Select a.Id
                Into n_����id
                From �������� A, �ҺŰ������� B
                Where a.���� = b.�������� And b.�ű�id = c_����.����id And Rownum < 2;
              End If;
            Exception
              When Others Then
                n_����id := Null;
            End;
          End If;
        
          --a.�ٴ����ﰲ��
          Select Count(1)
          Into n_�Ƿ���ʱ����
          From �ٴ����ﰲ��
          Where ����id = n_����id And ��Դid = n_��Դid And Rownum < 2;
          Insert Into �ٴ����ﰲ��
            (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��, �Ƿ���ʱ����)
          Values
            (n_����id, n_����id, n_��Դid, c_����.��Ŀid, c_����.ҽ��id, c_����.ҽ������, c_����.��ʼʱ��, c_����.��ֹʱ��, Zl_Username, c_����.��ʼʱ��,
             n_�Ƿ���ʱ����);
        
          --b.�ٴ���������
          --˵������Լ������0��ʾ��ֹԤԼ����Լ��Ϊ�ձ�ʾ������ԤԼ
          If Nvl(c_����.�ƻ�id, 0) <> 0 Then
            Select Count(1) Into n_Count From �Һżƻ����� Where �ƻ�id = c_����.�ƻ�id And Rownum < 2;
            If n_Count = 0 Then
              Insert Into �ٴ���������
                (ID, ����id, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ������Ŀ, �ϰ�ʱ��, ���﷽ʽ, ����id)
                Select �ٴ���������_Id.Nextval, n_����id, Null, Null, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��,
                       Nvl((Select 1
                            From �Һżƻ�����
                            Where �ƻ�id = c_����.�ƻ�id And ������Ŀ = a.������Ŀ And ��Լ�� = 0 And Rownum < 2), 0), a.������Ŀ, a.�ϰ�ʱ��,
                       c_����.���﷽ʽ, n_����id
                From (Select '����' As ������Ŀ, c_����.���� As �ϰ�ʱ��
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '��һ', c_����.��һ
                       From Dual
                       Where c_����.��һ Is Not Null
                       Union All
                       Select '�ܶ�', c_����.�ܶ�
                       From Dual
                       Where c_����.�ܶ� Is Not Null
                       Union All
                       Select '����', c_����.����
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '����', c_����.����
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '����', c_����.����
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '����', c_����.���� From Dual Where c_����.���� Is Not Null) A,
                     (Select Distinct ���� From �Һżƻ�ʱ�� Where �ƻ�id = c_����.�ƻ�id) B
                Where a.������Ŀ = b.����(+);
            Else
              Insert Into �ٴ���������
                (ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id)
                Select �ٴ���������_Id.Nextval, n_����id, ������Ŀ,
                       Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                               c_����.����, '����', c_����.����, Null), �޺���, ��Լ��, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��,
                       Nvl((Select 1
                            From �Һżƻ�����
                            Where �ƻ�id = c_����.�ƻ�id And ������Ŀ = a.������Ŀ And ��Լ�� = 0 And Rownum < 2), 0), c_����.���﷽ʽ, n_����id
                From �Һżƻ����� A, (Select Distinct ���� From �Һżƻ�ʱ�� Where �ƻ�id = c_����.�ƻ�id) B
                Where a.������Ŀ = b.����(+) And �ƻ�id = c_����.�ƻ�id;
            End If;
          Else
            Select Count(1) Into n_Count From �ҺŰ������� Where ����id = c_����.����id And Rownum < 2;
            If n_Count = 0 Then
              Insert Into �ٴ���������
                (ID, ����id, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ������Ŀ, �ϰ�ʱ��, ���﷽ʽ, ����id)
                Select �ٴ���������_Id.Nextval, n_����id, Null, Null, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��,
                       Nvl((Select 1
                            From �ҺŰ�������
                            Where ����id = c_����.����id And ������Ŀ = a.������Ŀ And ��Լ�� = 0 And Rownum < 2), 0), a.������Ŀ, a.�ϰ�ʱ��,
                       c_����.���﷽ʽ, n_����id
                From (Select '����' As ������Ŀ, c_����.���� As �ϰ�ʱ��
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '��һ', c_����.��һ
                       From Dual
                       Where c_����.��һ Is Not Null
                       Union All
                       Select '�ܶ�', c_����.�ܶ�
                       From Dual
                       Where c_����.�ܶ� Is Not Null
                       Union All
                       Select '����', c_����.����
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '����', c_����.����
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '����', c_����.����
                       From Dual
                       Where c_����.���� Is Not Null
                       Union All
                       Select '����', c_����.���� From Dual Where c_����.���� Is Not Null) A,
                     (Select Distinct ���� From �ҺŰ���ʱ�� Where ����id = c_����.����id) B
                Where a.������Ŀ = b.����(+);
            Else
              Insert Into �ٴ���������
                (ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id)
                Select �ٴ���������_Id.Nextval, n_����id, ������Ŀ,
                       Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                               c_����.����, '����', c_����.����, Null), �޺���, ��Լ��, c_����.��ſ���, Decode(b.����, Null, 0, 1) As �Ƿ��ʱ��,
                       Nvl((Select 1
                            From �ҺŰ�������
                            Where ����id = c_����.����id And ������Ŀ = a.������Ŀ And ��Լ�� = 0 And Rownum < 2), 0), c_����.���﷽ʽ, n_����id
                From �ҺŰ������� A, (Select Distinct ���� From �ҺŰ���ʱ�� Where ����id = c_����.����id) B
                Where a.������Ŀ = b.����(+) And ����id = c_����.����id;
            End If;
          End If;
        
          --c.�ٴ���������
          If Nvl(c_����.���﷽ʽ, 0) > 0 Then
            If Nvl(c_����.�ƻ�id, 0) <> 0 Then
              Insert Into �ٴ���������
                (����id, ����id)
                Select a.Id, b.����id
                From �ٴ��������� A,
                     (Select Distinct a.Id As ����id
                       From �������� A, �Һżƻ����� B
                       Where a.���� = b.�������� And b.�ƻ�id = c_����.�ƻ�id) B
                Where a.����id = n_����id;
            Else
              Insert Into �ٴ���������
                (����id, ����id)
                Select a.Id, b.����id
                From �ٴ��������� A,
                     (Select Distinct a.Id As ����id
                       From �������� A, �ҺŰ������� B
                       Where a.���� = b.�������� And b.�ű�id = c_����.����id) B
                Where a.����id = n_����id;
            End If;
          End If;
        
          --D.�ٴ�����ʱ��
          If Nvl(c_����.�ƻ�id, 0) <> 0 Then
            Insert Into �ٴ�����ʱ��
              (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
              Select a.Id, b.���, b.��ʼʱ��, b.����ʱ��, b.��������, b.�Ƿ�ԤԼ
              From �ٴ��������� A,
                   (Select n_����id As ����id, ����,
                            Decode(����, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                    c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ



                     
                     From �Һżƻ�ʱ��
                     Where �ƻ�id = c_����.�ƻ�id) B
              Where a.����id = b.����id And a.������Ŀ = b.���� And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
          
          Else
            Insert Into �ٴ�����ʱ��
              (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
              Select a.Id, b.���, b.��ʼʱ��, b.����ʱ��, b.��������, b.�Ƿ�ԤԼ
              From �ٴ��������� A,
                   (Select n_����id As ����id, ����,
                            Decode(����, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                    c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ



                     
                     From �ҺŰ���ʱ��
                     Where ����id = c_����.����id) B
              Where a.����id = b.����id And a.������Ŀ = b.���� And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
          End If;
        
          --����ʱ�ε���ſ��ƺ����������
          --��ʼʱ�䡢��ֹʱ����дʱ��εĿ�ʼʱ��ͽ���ʱ��
          For c_������Ŀ In (Select ID, �޺���, �ϰ�ʱ��
                         From �ٴ���������
                         Where ����id = n_����id And Nvl(�޺���, 0) <> 0 And Nvl(�Ƿ���ſ���, 0) = 1 And Nvl(�Ƿ��ʱ��, 0) = 0) Loop
            For I In 1 .. c_������Ŀ.�޺��� Loop
              Insert Into �ٴ�����ʱ��
                (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
                Select c_������Ŀ.Id, I, ��ʼʱ��, ��ֹʱ��, 1, 1
                From ʱ���
                Where վ�� Is Null And ���� Is Null And ʱ��� = c_������Ŀ.�ϰ�ʱ��;
            End Loop;
          End Loop;
        
          --�κ�һ����������ԤԼʱ��ʾȫ������ԤԼ
          Update �ٴ�����ʱ�� A
          Set a.�Ƿ�ԤԼ = 1
          Where ����id In (Select ID From �ٴ��������� Where ����id = n_����id) And Not Exists
           (Select 1 From �ٴ�����ʱ�� B Where a.����id = b.����id And Nvl(b.�Ƿ�ԤԼ, 0) = 1);
        
          --E.������λ�Һſ���
          If Nvl(c_����.�ƻ�id, 0) <> 0 Then
            Insert Into �ٴ�����Һſ���
              (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
              Select a.Id, b.����, b.����, b.������λ, b.���, b.���Ʒ�ʽ, b.����
              From �ٴ��������� A,
                   (Select 1 As ����, 1 As ����, ������λ, n_����id As ����id, ������Ŀ,
                            Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                    c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���,
                            Case
                               When Nvl(���, 0) = 0 And Nvl(����, 0) = 0 Then
                                0
                               When ��� = 0 And Nvl(����, 0) <> 0 Then
                                2
                               When Nvl(���, 0) <> 0 And Nvl(����, 0) <> 0 Then
                                3
                               Else
                                4
                             End As ���Ʒ�ʽ, ����
                     From ������λ�ƻ�����
                     Where �ƻ�id = c_����.�ƻ�id And
                           Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                  c_����.����, '����', c_����.����, Null) Is Not Null) B
              Where a.����id = b.����id And a.������Ŀ = b.������Ŀ And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
          Else
            Insert Into �ٴ�����Һſ���
              (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
              Select a.Id, b.����, b.����, b.������λ, b.���, b.���Ʒ�ʽ, b.����
              From �ٴ��������� A,
                   (Select 1 As ����, 1 As ����, ������λ, n_����id As ����id, ������Ŀ,
                            Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                    c_����.����, '����', c_����.����, Null) As �ϰ�ʱ��, ���,
                            Case
                              When Nvl(���, 0) = 0 And Nvl(����, 0) = 0 Then
                               0
                              When ��� = 0 And Nvl(����, 0) <> 0 Then
                               2
                              When Nvl(���, 0) <> 0 And Nvl(����, 0) <> 0 Then
                               3
                              Else
                               4
                            End As ���Ʒ�ʽ, ����
                     From ������λ���ſ���
                     Where ����id = c_����.����id And
                           Decode(������Ŀ, '����', c_����.����, '��һ', c_����.��һ, '�ܶ�', c_����.�ܶ�, '����', c_����.����, '����', c_����.����, '����',
                                  c_����.����, '����', c_����.����, Null) Is Not Null) B
              Where a.����id = b.����id And a.������Ŀ = b.������Ŀ And a.�ϰ�ʱ�� = b.�ϰ�ʱ��;
          End If;
        End Loop;
      
        --4.ͣ��û����Ч���ŵĺ�Դ����ɾ����Ч�İ���
        --��Ҫ�Ǵ������мƻ���ʧЧ������Ч�ƻ�ֻ��һ��������ƻ���һ�����ն�û���ϰ�ʱ�ε�
        Select Count(1)
        Into n_Count
        From �ٴ��������� A, �ٴ����ﰲ�� B, �ٴ������Դ C
        Where a.����id = b.Id And b.��Դid = c.Id And c.Id = n_��Դid And a.�ϰ�ʱ�� Is Not Null And Nvl(c.�Ƿ�ɾ��, 0) = 0 And
              Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And Rownum < 2;
        If n_Count = 0 Then
          --û����Ч�İ��ţ�ͣ�ú�Դ��ɾ������
          Select a.Id Bulk Collect
          Into l_����id
          From �ٴ��������� A, �ٴ����ﰲ�� B
          Where a.����id = b.Id And b.��Դid = n_��Դid;
        
          Forall I In 1 .. l_����id.Count
            Delete From �ٴ��������� Where ����id = l_����id(I);
        
          Forall I In 1 .. l_����id.Count
            Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
        
          Forall I In 1 .. l_����id.Count
            Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
        
          Forall I In 1 .. l_����id.Count
            Delete From �ٴ��������� Where ID = l_����id(I);
        
          Delete From �ٴ����ﰲ�� Where ��Դid = n_��Դid;
        
          Update �ٴ������Դ Set ����ʱ�� = Sysdate Where ID = n_��Դid;
        End If;
      
        --5.����һ�ݳ�����Ϣ��Ϊ��Դ������Ϣ
        --˵�����ϰ�ʱ�ΰ����ŵĵǼ�ʱ�䵹��ȡ��һ��
        For c_���� In (Select ID, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id
                     From (Select a.Id, a.�ϰ�ʱ��, a.�޺���, a.��Լ��, a.�Ƿ���ſ���, a.�Ƿ��ʱ��, a.ԤԼ����, a.�Ƿ��ռ, a.���﷽ʽ, a.����id,
                                   Row_Number() Over(Partition By a.�ϰ�ʱ�� Order By b.�Ǽ�ʱ�� Desc) As ���
                            From �ٴ��������� A, �ٴ����ﰲ�� B
                            Where a.����id = b.Id And b.��Դid = n_��Դid)
                     Where ��� = 1) Loop
          --a.�ٴ������Դ����
          Select �ٴ������Դ����_Id.Nextval Into n_����id From Dual;
          Insert Into �ٴ������Դ����
            (ID, ��Դid, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id)
          Values
            (n_����id, n_��Դid, c_����.�ϰ�ʱ��, c_����.�޺���, c_����.��Լ��, c_����.�Ƿ���ſ���, c_����.�Ƿ��ʱ��, c_����.ԤԼ����, c_����.�Ƿ��ռ, c_����.���﷽ʽ,
             c_����.����id);
          --b.�ٴ������Դ����
          Insert Into �ٴ������Դ����
            (����id, ����id)
            Select n_����id, ����id From �ٴ��������� Where ����id = c_����.Id;
          --c.�ٴ������Դʱ��
          Insert Into �ٴ������Դʱ��
            (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
            Select n_����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ From �ٴ�����ʱ�� Where ����id = c_����.Id;
          --d.�ٴ������Դ����
          Insert Into �ٴ������Դ����
            (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
            Select n_����id, ����, ����, ����, ���, ���Ʒ�ʽ, ���� From �ٴ�����Һſ��� Where ����id = c_����.Id;
        End Loop;
      End If;
    End Loop;
  End;
Begin
  If Nvl(��ʼ_In, 0) = 1 Then
    Select Count(1) Into n_Count From �ٴ������ A, �ٴ����ﰲ�� B Where a.Id = b.����id And Rownum < 2;
    If n_Count <> 0 Then
      v_Err_Msg := '��ǰ�Ѿ������ٴ����ﰲ���ˣ�����ɾ�������������룡';
      Raise Err_Item;
    End If;
  
    Begin
      Select f_List2str(Cast(Collect(s.ʱ���) As t_Strlist))
      Into v_ʱ���
      From (Select ʱ���, Row_Number() Over(Partition By ʱ��� Order By ʱ���) As ���
             From (Select Decode(b.�к�, 1, a.��һ, 2, a.�ܶ�, 3, a.����, 4, a.����, 5, a.����, 6, a.����, a.����) As ʱ���
                    From (Select ��һ, �ܶ�, ����, ����, ����, ����, ����
                           From �ҺŰ���
                           Union All
                           Select ��һ, �ܶ�, ����, ����, ����, ����, ����
                           From �ҺŰ��żƻ�
                           Where ���ʱ�� Is Not Null And ʧЧʱ�� > Sysdate) A,
                         (Select Level As �к� From Dual Connect By Level <= 7) B)
             Where ʱ��� Is Not Null) S, ʱ��� T
      Where s.ʱ��� = t.ʱ���(+) And t.ʱ��� Is Null And s.��� = 1;
    Exception
      When Others Then
        v_ʱ��� := Null;
    End;
  
    If v_ʱ��� Is Not Null Then
      v_Err_Msg := 'ԭ�ҺŰ����е��ϰ�ʱ��Ρ�' || v_ʱ��� || '�������ڣ������ڡ���������>�ϰ�ʱ���������ӣ�';
      Raise Err_Item;
    End If;
  
    --ɾ���������к�Դ���ڵ���֮ǰ�ѽ�������ʾ
    Select a.Id Bulk Collect Into l_����id From �ٴ������Դ���� A, �ٴ������Դ B Where a.��Դid = b.Id;
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ������Դ���� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ������Դʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ������Դ���� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ������Դ���� Where ID = l_����id(I);
  
    Delete From �ٴ������Դ;
  
    --ɾ������ͣ���¼
    Delete From �ٴ�����ͣ���¼;
  End If;

  --����վ���Դ���û��ָ��վ�㣬������һ���������
  v_ȫԺ��Դ����վ�� := zl_GetSysParameter('δ����վ��ĺ�Դ��ά��վ��', 1114);
  Begin
    Select Min(ID) Into n_����id From �ٴ������ Where �Ű෽ʽ = 0;
  Exception
    When Others Then
      n_����id := 0;
  End;
  If Nvl(n_����id, 0) = 0 Then
    Select �ٴ������_Id.Nextval Into n_����id From Dual;
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, ���, վ��)
    Values
      (n_����id, 0, '�̶������', To_Number(To_Char(Sysdate, 'yyyy')), Null);
  End If;

  If Not ����_In Is Null Then
    Zl_Register_Import(����_In, n_����id, v_ȫԺ��Դ����վ��);
    Return;
  End If;

  For c_��Դ In (Select ���� From �ҺŰ��� Order By ID Desc) Loop
    --ɾ���Լ�ͣ�õĺ�ԴҲȫ������
    Zl_Register_Import(c_��Դ.����, n_����id, v_ȫԺ��Դ����վ��);
  End Loop;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_����;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_Third_Getregstatus
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:��ȡĳ�����ҵ����Ű�ҽ���ĹҺ�����
  --���:Xml_In: 
  --  <IN> 
  --      <KSID></KSID>   --����ID
  --  </IN> 
  --����:Xml_Out 
  --  <OUTPUT> 
  --    <YS>
  --      <YSXM></YSXM>    --ҽ������
  --      <SYGHS></SYGHS>  --ʣ��Һ���
  --      <DDJZS></DDJZS>  --�ȴ�������
  --      <SWGHS></SWGHS>  --����Һ���
  --      <XWGHS></XWGHS>  --����Һ���
  --      <QTGHS></QTGHS>  --ȫ��Һ���
  --      <YSZJ><YSZJ>     --ҽ��ְ��
  --    </YS>
  --    <YS>
  --      ...
  --    </YS>  
  --  </OUTPUT> 
  -------------------------------------------------------------------------------------------------- 

  n_����id      �ҺŰ���.����id%Type;
  d_����        Date;
  v_Para        Varchar2(5000);
  n_�Һ�ģʽ    Number(3);
  d_����ʱ��    Date;
  v_Temp        Varchar2(32767); --��ʱXML 
  x_Templet     Xmltype; --ģ��XML 
  v_Err_Msg     Varchar2(200);
  n_�ѹ���      ���˹ҺŻ���.�ѹ���%Type;
  n_�޺���      �ҺŰ�������.�޺���%Type;
  n_�������    Number;
  n_�������    Number;
  n_�ȴ�����    Number;
  v_�����¼ids Varchar2(5000);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/KSID') Into n_����id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  d_���� := Sysdate;

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  If n_�Һ�ģʽ = 0 Or (n_�Һ�ģʽ = 1 And d_���� < d_����ʱ��) Then
    --�ƻ��Ű�ģʽ
    For r_ҽ�� In (Select Distinct ҽ��id, ҽ������, y.רҵ����ְ��
                 From (Select a.ҽ��id, a.ҽ������
                        From �ҺŰ��żƻ� A, �ҺŰ��� B
                        Where a.����id = b.Id And b.ͣ������ Is Null And a.���ʱ�� Is Not Null And b.����id = n_����id And
                              a.��Чʱ�� = (Select Max(��Чʱ��)
                                        From �ҺŰ��żƻ�
                                        Where ����id = b.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��) And
                              Not Exists (Select 1
                               From �ҺŰ���ͣ��״̬
                               Where ����id = b.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                              d_���� Between Nvl(b.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(b.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))
                        Union All
                        Select a.ҽ��id, a.ҽ������
                        From �ҺŰ��� A
                        Where a.ͣ������ Is Null And a.����id = n_����id And
                              d_���� Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                         (Select 1
                               From �ҺŰ���ͣ��״̬
                               Where ����id = a.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And Not Exists
                         (Select 1
                               From �ҺŰ��żƻ�
                               Where ����id = a.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��)) X, ��Ա�� Y
                 Where x.ҽ��id = y.Id(+)) Loop
      If r_ҽ��.ҽ������ Is Not Null Then
        v_Temp := '<YS>';
        v_Temp := v_Temp || '<YSXM>' || r_ҽ��.ҽ������ || '</YSXM>';
        Select Nvl(Sum(a.�ѹ���), 0)
        Into n_�ѹ���
        From ���˹ҺŻ��� A,
             (Select Distinct ����
               From (Select b.����
                      From �ҺŰ��żƻ� A, �ҺŰ��� B
                      Where a.����id = b.Id And b.ͣ������ Is Null And a.���ʱ�� Is Not Null And b.����id = n_����id And
                            a.��Чʱ�� = (Select Max(��Чʱ��)
                                      From �ҺŰ��żƻ�
                                      Where ����id = b.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��) And Not Exists
                       (Select 1
                             From �ҺŰ���ͣ��״̬
                             Where ����id = b.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                            d_���� Between Nvl(b.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                            Nvl(b.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))
                      Union All
                      Select a.����
                      From �ҺŰ��� A
                      Where a.ͣ������ Is Null And a.����id = n_����id And
                            d_���� Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                            Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                       (Select 1
                             From �ҺŰ���ͣ��״̬
                             Where ����id = a.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And Not Exists
                       (Select 1
                             From �ҺŰ��żƻ�
                             Where ����id = a.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��))) B
        Where a.ҽ������ = r_ҽ��.ҽ������ And a.����id = n_����id And a.���� = b.���� And ���� = Trunc(d_����);
      
        Select Nvl(Sum(�޺���), 0)
        Into n_�޺���
        From (Select c.�޺���
               From �ҺŰ��żƻ� A, �ҺŰ��� B, �Һżƻ����� C
               Where a.Id = c.�ƻ�id And c.������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                       '����', '6', '����', '7', '����', Null) And a.����id = b.Id And
                     b.ͣ������ Is Null And a.���ʱ�� Is Not Null And b.����id = n_����id And
                     a.��Чʱ�� = (Select Max(��Чʱ��)
                               From �ҺŰ��żƻ�
                               Where ����id = b.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��) And Not Exists
                (Select 1 From �ҺŰ���ͣ��״̬ Where ����id = b.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                     d_���� Between Nvl(b.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                     Nvl(b.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))
               Union All
               Select b.�޺���
               From �ҺŰ��� A, �ҺŰ������� B
               Where a.Id = b.����id And b.������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                       '����', '6', '����', '7', '����', Null) And a.ͣ������ Is Null And
                     a.����id = n_����id And d_���� Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                     Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                (Select 1 From �ҺŰ���ͣ��״̬ Where ����id = a.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And Not Exists
                (Select 1
                      From �ҺŰ��żƻ�
                      Where ����id = a.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��));
        If n_�޺��� - n_�ѹ��� < 0 Then
          v_Temp := v_Temp || '<SYGHS>' || 0 || '</SYGHS>';
        Else
          v_Temp := v_Temp || '<SYGHS>' || To_Char(n_�޺��� - n_�ѹ���) || '</SYGHS>';
        End If;
        
        Select Count(1)
        Into n_�ȴ�����
        From ���˹Һż�¼
        Where ִ���� = r_ҽ��.ҽ������ And ִ�в���id = n_����id And Nvl(ִ��״̬, 0) = 0 And ��¼���� = 1 And ��¼״̬ = 1 And
              ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 24 / 60 / 60;
        v_Temp := v_Temp || '<DDJZS>' || n_�ȴ����� || '</DDJZS>';
        
        Select Count(1)
        Into n_�������
        From ���˹Һż�¼
        Where ִ���� = r_ҽ��.ҽ������ And ִ�в���id = n_����id And ��¼���� = 1 And ��¼״̬ = 1 And
              ����ʱ�� Between Trunc(d_����) And Trunc(d_����) + 0.5 - 1 / 24 / 60 / 60;
        v_Temp := v_Temp || '<SWGHS>' || n_������� || '</SWGHS>';
        
        Select Count(1)
        Into n_�������
        From ���˹Һż�¼
        Where ִ���� = r_ҽ��.ҽ������ And ִ�в���id = n_����id And ��¼���� = 1 And ��¼״̬ = 1 And
              ����ʱ�� Between Trunc(d_����) + 0.5 And Trunc(d_���� + 1) - 1 / 24 / 60 / 60;
        v_Temp := v_Temp || '<XWGHS>' || n_������� || '</XWGHS>';
        
        v_Temp := v_Temp || '<QTGHS>' || To_Char(n_������� + n_�������) || '</QTGHS>';
        v_Temp := v_Temp || '<YSZJ>' || r_ҽ��.רҵ����ְ�� || '</YSZJ>';
        v_Temp := v_Temp || '</YS>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      End If;
    End Loop;
  Else
    --������Ű�ģʽ
    For r_ҽ�� In (Select Distinct a.ҽ��id, a.ҽ������, b.רҵ����ְ��
                 From �ٴ������¼ A, ��Ա�� B
                 Where a.ҽ��id = b.Id(+) And a.����id = n_����id And �Ƿ񷢲� = 1 And Nvl(�Ƿ�����, 0) = 0 And
                       (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��))) Loop
      For r_�����¼ In (Select Distinct ID
                     From �ٴ������¼
                     Where ҽ������ = r_ҽ��.ҽ������ And ����id = n_����id And �������� = Trunc(d_����) And �Ƿ񷢲� = 1 And Nvl(�Ƿ�����, 0) = 0 And
                           (��ʼʱ�� < Nvl(ͣ�￪ʼʱ��, ��ֹʱ��) Or ��ֹʱ�� > Nvl(ͣ����ֹʱ��, ��ʼʱ��))) Loop
        v_�����¼ids := v_�����¼ids || ',' || r_�����¼.Id;
      End Loop;
      If v_�����¼ids Is Not Null Then
        v_�����¼ids := Substr(v_�����¼ids, 2);
      End If;
      Select Sum(�޺���), Sum(�ѹ���)
      Into n_�޺���, n_�ѹ���
      From �ٴ������¼ A, Table(f_Str2list(v_�����¼ids)) B
      Where a.Id = b.Column_Value;
      v_Temp := '<YS>';
      v_Temp := v_Temp || '<YSXM>' || r_ҽ��.ҽ������ || '</YSXM>';
      If n_�޺��� - n_�ѹ��� < 0 Then
        v_Temp := v_Temp || '<SYGHS>' || 0 || '</SYGHS>';
      Else
        v_Temp := v_Temp || '<SYGHS>' || To_Char(n_�޺��� - n_�ѹ���) || '</SYGHS>';
      End If;
      Select Count(1)
      Into n_�ȴ�����
      From ���˹Һż�¼ A, Table(f_Str2list(v_�����¼ids)) B
      Where ִ���� = r_ҽ��.ҽ������ And ִ�в���id = n_����id And Nvl(ִ��״̬, 0) = 0 And ��¼���� = 1 And ��¼״̬ = 1 And
            a.�����¼id = b.Column_Value;
      v_Temp := v_Temp || '<DDJZS>' || n_�ȴ����� || '</DDJZS>';
      
      Select Count(1)
      Into n_�������
      From ���˹Һż�¼ A, Table(f_Str2list(v_�����¼ids)) B
      Where ִ���� = r_ҽ��.ҽ������ And ִ�в���id = n_����id And ��¼���� = 1 And ��¼״̬ = 1 And
            ����ʱ�� Between Trunc(d_����) And Trunc(d_����) + 0.5 - 1 / 24 / 60 / 60 And a.�����¼id = b.Column_Value;
      v_Temp := v_Temp || '<SWGHS>' || n_������� || '</SWGHS>';
      
      Select Count(1)
      Into n_�������
      From ���˹Һż�¼ A, Table(f_Str2list(v_�����¼ids)) B
      Where ִ���� = r_ҽ��.ҽ������ And ִ�в���id = n_����id And ��¼���� = 1 And ��¼״̬ = 1 And
            ����ʱ�� Between Trunc(d_����) + 0.5 And Trunc(d_���� + 1) - 1 / 24 / 60 / 60 And a.�����¼id = b.Column_Value;
      v_Temp := v_Temp || '<XWGHS>' || n_������� || '</XWGHS>';
      
      v_Temp := v_Temp || '<QTGHS>' || To_Char(n_������� + n_�������) || '</QTGHS>';
      v_Temp := v_Temp || '<YSZJ>' || r_ҽ��.רҵ����ְ�� || '</YSZJ>';
      v_Temp := v_Temp || '</YS>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End Loop;
  End If;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregstatus;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_Third_Getdoctor
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:��ȡĳ�������µ��Ű��ҽ�������ڻ�ʿ����ʱѡ��ҽ��
  --���:Xml_In: 
  --  <IN> 
  --      <KSID></KSID>   --����ID
  --      <RQ></RQ>   --����,Ĭ��Ϊ����
  --  </IN> 
  --����:Xml_Out 
  --  <OUTPUT> 
  --    <YS>
  --      <YSID></YSID>    --ҽ��ID
  --      <YSXM></YSXM>     --ҽ������
  --    </YS>
  --    <YS>
  --      ...
  --    </YS>  
  --  </OUTPUT> 
  -------------------------------------------------------------------------------------------------- 

  n_����id   �ҺŰ���.����id%Type;
  d_����     Date;
  v_Para     Varchar2(5000);
  n_�Һ�ģʽ Number(3);
  d_����ʱ�� Date;
  n_����ҽ�� Number(3);
  v_Temp     Varchar2(32767); --��ʱXML 
  x_Templet  Xmltype; --ģ��XML 
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/KSID'), To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd hh24:mi:ss')
  Into n_����id, d_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If d_���� Is Null Then
    d_���� := Sysdate;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  If n_�Һ�ģʽ = 0 Or (n_�Һ�ģʽ = 1 And d_���� < d_����ʱ��) Then
    --�ƻ��Ű�ģʽ
    For r_ҽ�� In (Select Distinct ҽ��id, ҽ������
                 From (Select a.ҽ��id, a.ҽ������
                        From �ҺŰ��żƻ� A, �ҺŰ��� B
                        Where a.����id = b.Id And b.ͣ������ Is Null And a.���ʱ�� Is Not Null And b.����id = n_����id And
                              a.��Чʱ�� = (Select Max(��Чʱ��)
                                        From �ҺŰ��żƻ�
                                        Where ����id = b.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��) And
                              Not Exists (Select 1
                               From �ҺŰ���ͣ��״̬
                               Where ����id = b.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                              d_���� Between Nvl(b.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(b.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))
                        Union All
                        Select a.ҽ��id, a.ҽ������
                        From �ҺŰ��� A
                        Where a.ͣ������ Is Null And a.����id = n_����id And
                              d_���� Between Nvl(a.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                              Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                         (Select 1
                               From �ҺŰ���ͣ��״̬
                               Where ����id = a.Id And d_���� Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And Not Exists
                         (Select 1
                               From �ҺŰ��żƻ�
                               Where ����id = a.Id And ���ʱ�� Is Not Null And d_���� Between ��Чʱ�� + 0 And ʧЧʱ��))) Loop
      If Nvl(r_ҽ��.ҽ��id, 0) <> 0 Or Nvl(r_ҽ��.ҽ������, '-') <> '-' Then
        v_Temp := '<YS><YSID>' || r_ҽ��.ҽ��id || '</YSID><YSXM>' || r_ҽ��.ҽ������ || '</YSXM></YS>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        n_����ҽ�� := 1;
        Exit;
      End If;
    End Loop;
  Else
    --������Ű�ģʽ
    For r_ҽ�� In (Select Distinct ҽ��id, ҽ������
                 From �ٴ������¼
                 Where ����id = n_����id And �������� = Trunc(d_����) And �Ƿ񷢲� = 1 And Nvl(�Ƿ�����, 0) = 0 And
                       (��ʼʱ�� < Nvl(ͣ�￪ʼʱ��, ��ֹʱ��) Or ��ֹʱ�� > Nvl(ͣ����ֹʱ��, ��ʼʱ��))) Loop
      If Nvl(r_ҽ��.ҽ��id, 0) <> 0 Or Nvl(r_ҽ��.ҽ������, '-') <> '-' Then
        v_Temp := '<YS><YSID>' || r_ҽ��.ҽ��id || '</YSID><YSXM>' || r_ҽ��.ҽ������ || '</YSXM></YS>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Else
        n_����ҽ�� := 1;
        Exit;
      End If;
    End Loop;
  End If;

  If Nvl(n_����ҽ��, 0) = 1 Then
    x_Templet := Xmltype('<OUTPUT></OUTPUT>');
    For r_ҽ�� In (Select Distinct a.Id, a.����
                 From ��Ա�� A, ������Ա B, ��Ա����˵�� C
                 Where a.Id = b.��Աid And b.����id = n_����id And a.Id = c.��Աid And c.��Ա���� = 'ҽ��' And
                       (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)) Loop
      v_Temp := '<YS><YSID>' || r_ҽ��.Id || '</YSID><YSXM>' || r_ҽ��.���� || '</YSXM></YS>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End Loop;
  End If;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdoctor;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
CREATE OR REPLACE Procedure Zl_Third_Docarrange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:ҽ���Ű�ƻ�
  --���:Xml_In:
  --<IN>
  --   <YSID>870</YSID>    //ҽ��ID
  --   <KDID>870</KSID>    //����ID
  --   <KSSJ>2014-10-29 </KSSJ>    //��ʼʱ��
  --   <CXTS>14</CXTS>    //��ѯ����
  --   <HZDW>֧����</HZDW> //������λ
  --   <HL>����</HL>      //���࣬�ɴ�������ö��ŷָ�����ʽ:��ͨ,ר��,...
  --   <ZD></ZD>        //վ��
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  --   <PBLIST>       //δ���ظýڵ��ʾû������
  --    <PB>
  --     <RQ>2014-10-29</RQ>     //����
  --     <SYHS>5</SYHS>    //ʣ�����
  --     <SBSJ>ȫ��</SBSJ>             //�ϰ�ʱ��
  --     <YGS>5</YGS>    //�ѹҺ���
  --    </PB>
  --   <PBLIST>
  --   <ERROR><MSG></MSG></ERROR> //�����������
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_�޺���       �ҺŰ�������.�޺���%Type;
  n_�ѹ���       �ҺŰ�������.�޺���%Type;
  n_���ѹ���     �ҺŰ�������.�޺���%Type;
  n_��Լ��       �ҺŰ�������.�޺���%Type;
  n_��Լ��       �ҺŰ�������.�޺���%Type;
  n_ʣ����       �ҺŰ�������.�޺���%Type;
  v_�ϰ�ʱ��     varchar2(300);
  n_ҽ��id       ��Ա��.Id%Type;
  n_����id       ���ű�.Id%Type;
  n_��ѯ����     Number(4);
  n_������λ���� Number(5);
  n_��Լ�ѹ���   Number(4);
  n_��Լ����     Number(3);
  n_���Ŵ���     Number(3);
  v_����         �ҺŰ���.����%Type;
  n_����id       �ҺŰ��żƻ�.����id%Type;
  n_�ƻ�id       �ҺŰ��żƻ�.Id%Type;
  v_������λ     �Һź�����λ.����%Type;
  n_Daycount     Number(4);
  d_��ʼʱ��     Date;
  d_ԭʼʱ��     Date;
  n_����         Number(3);
  v_Temp         varchar2(32767); --��ʱXML
  x_Templet      Xmltype; --ģ��XML
  v_Err_Msg      varchar2(200);
  v_����         varchar2(200);
  n_Exists       Number(2);
  n_ԤԼ����     Number(5);
  n_��������     Number(5);
  n_�Һ�ģʽ     Number(3);
  n_��Լģʽ     �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type;
  v_����ʱ��     varchar2(500);
  v_��ͨ�ȼ�     varchar2(100);
  v_Pricegrade   varchar2(500);
  v_վ��         ���ű�.վ��%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/KSID'), Extractvalue(Value(A), 'IN/CXTS'),
         To_Date(Extractvalue(Value(A), 'IN/KSSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/HL'), Extractvalue(Value(A), 'IN/ZD')
  Into n_ҽ��id, n_����id, n_��ѯ����, d_��ʼʱ��, v_������λ, v_����, v_վ��
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_��ѯ���� := Nvl(n_��ѯ����, 14);
  n_ԤԼ���� := Nvl(zl_GetSysParameter(66), 7);
  d_ԭʼʱ�� := Trunc(d_��ʼʱ��);
  d_��ʼʱ�� := Trunc(d_��ʼʱ��);
  n_Daycount := 0;

  v_Pricegrade := Zl_Get_Pricegrade(v_վ��);
  v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);

  n_�Һ�ģʽ := To_Number(Substr(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), '0'), 1, 1));
  v_����ʱ�� := Substr(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), '0'), 3);
  If n_�Һ�ģʽ = 0 Then
    If Nvl(n_����id, 0) = 0 Then
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
          d_��ʼʱ�� := Sysdate - n_Daycount;
        Else
          d_��ʼʱ�� := d_ԭʼʱ��;
        End If;
        n_���Ŵ��� := 0;
        v_�ϰ�ʱ�� := Null;
        n_���ѹ��� := 0;
        n_�ѹ���   := 0;
        n_ʣ����   := 0;
        n_�޺���   := 0;
        n_��Լ��   := 0;
        n_��Լ��   := 0;
        For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                            a.����id, a.�ƻ�id, a.����, a.����
                     
                     From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����,
                                   Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��



                            
                            From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                          Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                  Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                   Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = n_ҽ��id And
                                         Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And Ap.ͣ������ Is Null And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         Not Exists
                                    (Select Rownum
                                          From �ҺŰ��żƻ� Jh
                                          Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Jh.ʧЧʱ��)
                                   Union All
                                   Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                          Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                  Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                   Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And
                                         Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And Ap.ͣ������ Is Null And
                                         Jh.ҽ��id = n_ҽ��id And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Jh.ʧЧʱ�� And Xz.�ƻ�id(+) = Jh.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         (Jh.��Чʱ��, Jh.����id) =
                                         (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                          From �ҺŰ��żƻ� Sxjh
                                          Where Sxjh.���ʱ�� Is Not Null And d_��ʼʱ�� + n_Daycount Between
                                                Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                          Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                            Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                          ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                     Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                           b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount And
                           (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                           (b.�۸�ȼ� Is Null And Not Exists
                            (Select 1
                              From �շѼ�Ŀ
                              Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And Sysdate Between ִ������ And
                                    Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
          If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
            v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
            n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
            n_�ѹ���   := r_�Ű�.�ѹ���;
            n_�޺���   := r_�Ű�.�޺���;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_����id   := Nvl(r_�Ű�.����id, 0);
            n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
            v_����     := r_�Ű�.����;
            n_���Ŵ��� := 1;
            If v_�ϰ�ʱ�� Is Not Null Then
              If v_������λ Is Not Null Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ���ſ���
                    Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                End If;
              End If;
            
              If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                Else
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_��Լ�ѹ���
                  From ���˹Һż�¼
                  Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                        Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_��Լ�ѹ��� := 0;
                End;
                If n_������λ���� = 0 Then
                  n_������λ���� := Null;
                End If;
                If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                Else
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                End If;
              Else
                --��Լ��λ
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                End If;
                If Nvl(n_����, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                
                End If;
              End If;
            End If;
            n_������λ���� := 0;
            n_��Լ����     := 0;
            n_����         := 0;
          End If;
        End Loop;
        v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
        If n_���Ŵ��� = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
          d_��ʼʱ�� := Sysdate - n_Daycount;
        Else
          d_��ʼʱ�� := d_ԭʼʱ��;
        End If;
        v_�ϰ�ʱ�� := Null;
        n_���ѹ��� := 0;
        n_�ѹ���   := 0;
        n_ʣ����   := 0;
        n_�޺���   := 0;
        n_��Լ��   := 0;
        n_��Լ��   := 0;
        n_���Ŵ��� := 0;
        For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                            a.����id, a.�ƻ�id, a.����, a.����
                     
                     From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����,
                                   Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��



                            
                            From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                          Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                  Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                   Where Ap.����id = Bm.Id(+) And Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And
                                         Ap.ҽ��id = n_ҽ��id And Ap.����id = n_����id And Ap.ͣ������ Is Null And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         Not Exists
                                    (Select Rownum
                                          From �ҺŰ��żƻ� Jh
                                          Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Jh.ʧЧʱ��)
                                   Union All
                                   Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                          Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                          Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                  Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                          Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                   From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                   Where Jh.����id = Ap.Id And Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And
                                         Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And Ap.����id = n_����id And
                                         d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                         Jh.ʧЧʱ�� And Xz.�ƻ�id(+) = Jh.Id And
                                         Xz.������Ŀ(+) =
                                         Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                    (Select Rownum
                                          From �ҺŰ���ͣ��״̬ Ty
                                          Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                         (Jh.��Чʱ��, Jh.����id) =
                                         (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                          From �ҺŰ��żƻ� Sxjh
                                          Where Sxjh.���ʱ�� Is Not Null And d_��ʼʱ�� + n_Daycount Between
                                                Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                          Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                            Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                          ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                     Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                           b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount And
                           (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                           (b.�۸�ȼ� Is Null And Not Exists
                            (Select 1
                              From �շѼ�Ŀ
                              Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And Sysdate Between ִ������ And
                                    Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                     
                     ) Loop
          If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
            v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
            n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
            n_�ѹ���   := r_�Ű�.�ѹ���;
            n_�޺���   := r_�Ű�.�޺���;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_��Լ��   := r_�Ű�.��Լ��;
            n_����id   := Nvl(r_�Ű�.����id, 0);
            n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
            v_����     := r_�Ű�.����;
            n_���Ŵ��� := 1;
          
            If v_�ϰ�ʱ�� Is Not Null Then
              If v_������λ Is Not Null Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_��Լ����
                    From ������λ���ſ���
                    Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_��Լ���� := 0;
                  End;
                End If;
              End If;
            
              If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                Else
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null);
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                End If;
                Begin
                  Select Count(1)
                  Into n_��Լ�ѹ���
                  From ���˹Һż�¼
                  Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                        Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_��Լ�ѹ��� := 0;
                End;
                If n_������λ���� = 0 Then
                  n_������λ���� := Null;
                End If;
                If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                Else
                  n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                End If;
              Else
                --��Լ��λ
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select 1
                    Into n_����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                Else
                  Begin
                    Select 1
                    Into n_����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                          Rownum < 2;
                  Exception
                    When Others Then
                      n_���� := 0;
                  End;
                End If;
                If Nvl(n_����, 0) = 0 Then
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                
                End If;
              End If;
            End If;
            n_������λ���� := 0;
            n_��Լ����     := 0;
            n_����         := 0;
          End If;
        End Loop;
        v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
        If n_���Ŵ��� = 1 Then
          v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                    '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� || '</YGS>' ||
                    '</PB>';
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --������Ű�ģʽ
    n_�������� := Zl_Fun_Getappointmentdays;
    If Nvl(n_����id, 0) = 0 Then
      --ͨ��ҽ������
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) < To_Date(Substr(v_����ʱ��, 1, Instr(v_����ʱ��, ' ') - 1), 'yyyy-mm-dd') Then
          n_���Ŵ��� := 0;
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                              a.����id, a.�ƻ�id, a.����, a.����
                       From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��,
                                     Ap.����, Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��
                              
                              From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                            Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                    Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                     Where Ap.����id = Bm.Id(+) And Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And
                                           Ap.ҽ��id = n_ҽ��id And Ap.ͣ������ Is Null And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           Not Exists (Select Rownum
                                            From �ҺŰ��żƻ� Jh
                                            Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Jh.ʧЧʱ��)
                                     Union All
                                     Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                            Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                    Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                     Where Jh.����id = Ap.Id And Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And
                                           Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Jh.ʧЧʱ�� And Xz.�ƻ�id(+) = Jh.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           (Jh.��Чʱ��, Jh.����id) =
                                           (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                            From �ҺŰ��żƻ� Sxjh
                                            Where Sxjh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                            Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                              Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                            ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                       Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                             b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount And
                             (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                             (b.�۸�ȼ� Is Null And Not Exists
                              (Select 1
                                From �շѼ�Ŀ
                                Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And Sysdate Between ִ������ And
                                      Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
            If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
              v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
              n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
              n_�ѹ���   := r_�Ű�.�ѹ���;
              n_�޺���   := r_�Ű�.�޺���;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_����id   := Nvl(r_�Ű�.����id, 0);
              n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
              v_����     := r_�Ű�.����;
              n_���Ŵ��� := 1;
              If v_�ϰ�ʱ�� Is Not Null Then
                If v_������λ Is Not Null Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ���ſ���
                      Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  End If;
                End If;
              
                If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  Else
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  End If;
                Else
                  --��Լ��λ
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  End If;
                  If Nvl(n_����, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_��Լ�ѹ���
                      From ���˹Һż�¼
                      Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                            Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_��Լ�ѹ��� := 0;
                    End;
                    If n_�ƻ�id <> 0 Then
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ�ƻ�����
                        Where �ƻ�id = n_�ƻ�id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    Else
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ���ſ���
                        Where ����id = n_����id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    End If;
                    If n_������λ���� = 0 Then
                      n_������λ���� := Null;
                    End If;
                    n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                  
                  End If;
                End If;
              End If;
              n_������λ���� := 0;
              n_��Լ����     := 0;
              n_����         := 0;
            End If;
          End Loop;
          v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
          If n_���Ŵ��� = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_���Ŵ��� := 0;
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          If v_������λ Is Null Then
            --�Ǻ�����λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              --����Һ�
              For r_���� In (Select a.�ѹ���, a.�޺���, a.�ϰ�ʱ��
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.�ѹ���, 0);
                n_�޺���   := n_�޺��� + r_����.�޺���;
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --ԤԼ�Һ�
              For r_���� In (Select a.��Լ��, a.�޺���, a.��Լ��, a.�ϰ�ʱ��
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.��Լ��, 0);
                n_�޺���   := n_�޺��� + Nvl(r_����.��Լ��, r_����.�޺���);
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --������λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              For r_���� In (Select a.Id, a.�ѹ���, a.�޺���, a.��Լ��, a.�ϰ�ʱ��, a.�Ƿ���ſ���
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If r_����.�޺��� - r_����.�ѹ��� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + r_����.�޺��� - r_����.�ѹ���;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                    n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���Ŵ��� := 1;
                  Else
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
            
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --�ǵ���
              For r_���� In (Select a.Id, a.��Լ��, a.�ѹ���, a.�޺���, a.��Լ��, a.�ϰ�ʱ��, a.�Ƿ���ſ���
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ�� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ��;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    --��ʱ�η���ſ���
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From �ٴ�������ſ���
                      Where ԤԼ˳��� Is Not Null And ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) <> 0;
                      If r_����.���� - n_Exists > 0 Then
                        n_ʣ����   := n_ʣ���� + r_����.���� - n_Exists;
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  Else
                    For r_���� In (Select ���
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.��Լ��, 0);
                  n_ʣ����   := n_ʣ���� + r_����.��Լ�� - Nvl(r_����.��Լ��, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    Else
      --ͨ������+ҽ������
      While (n_Daycount < n_��ѯ����) Loop
        If Trunc(d_��ʼʱ�� + n_Daycount) < To_Date(Substr(v_����ʱ��, 1, Instr(v_����ʱ��, ' ') - 1), 'yyyy-mm-dd') Then
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          n_���Ŵ��� := 0;
          For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                              a.����id, a.�ƻ�id, a.����, a.����
                       
                       From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��,
                                     Ap.����, Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��
                              
                              From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                            Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                    Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                     Where Ap.����id = Bm.Id(+) And Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And
                                           Ap.ҽ��id = n_ҽ��id And Ap.����id = n_����id And Ap.ͣ������ Is Null And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           Not Exists (Select Rownum
                                            From �ҺŰ��żƻ� Jh
                                            Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Jh.ʧЧʱ��)
                                     Union All
                                     Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id,
                                            Jh.Id As �ƻ�id, Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                            Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                    Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                            Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                     From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                     Where Jh.����id = Ap.Id And Sysdate + Nvl(Ap.ԤԼ����, n_ԤԼ����) >= d_��ʼʱ�� + n_Daycount And
                                           Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And Ap.����id = n_����id And
                                           d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                           Jh.ʧЧʱ�� And Xz.�ƻ�id(+) = Jh.Id And
                                           Xz.������Ŀ(+) =
                                           Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                                  '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                      (Select Rownum
                                            From �ҺŰ���ͣ��״̬ Ty
                                            Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                           (Jh.��Чʱ��, Jh.����id) =
                                           (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                            From �ҺŰ��żƻ� Sxjh
                                            Where Sxjh.���ʱ�� Is Not Null And
                                                  d_��ʼʱ�� + n_Daycount Between
                                                  Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                                  Sxjh.ʧЧʱ�� And Sxjh.����id = Jh.����id
                                            Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                              Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                            ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                       Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                             b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount And
                             (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                             (b.�۸�ȼ� Is Null And Not Exists
                              (Select 1
                                From �շѼ�Ŀ
                                Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And Sysdate Between ִ������ And
                                      Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
            If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
              v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
              n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
              n_�ѹ���   := r_�Ű�.�ѹ���;
              n_�޺���   := r_�Ű�.�޺���;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_��Լ��   := r_�Ű�.��Լ��;
              n_����id   := Nvl(r_�Ű�.����id, 0);
              n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
              v_����     := r_�Ű�.����;
              n_���Ŵ��� := 1;
            
              If v_�ϰ�ʱ�� Is Not Null Then
                If v_������λ Is Not Null Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_��Լ����
                      From ������λ���ſ���
                      Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                    Exception
                      When Others Then
                        n_��Լ���� := 0;
                    End;
                  End If;
                End If;
              
                If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  Else
                    Begin
                      Select Sum(����)
                      Into n_������λ����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null);
                    Exception
                      When Others Then
                        n_������λ���� := 0;
                    End;
                  End If;
                  Begin
                    Select Count(1)
                    Into n_��Լ�ѹ���
                    From ���˹Һż�¼
                    Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                          Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                  Exception
                    When Others Then
                      n_��Լ�ѹ��� := 0;
                  End;
                  If n_������λ���� = 0 Then
                    n_������λ���� := Null;
                  End If;
                  If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  Else
                    n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
                  End If;
                Else
                  --��Լ��λ
                  If n_�ƻ�id <> 0 Then
                    Begin
                      Select 1
                      Into n_����
                      From ������λ�ƻ�����
                      Where �ƻ�id = n_�ƻ�id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  Else
                    Begin
                      Select 1
                      Into n_����
                      From ������λ���ſ���
                      Where ����id = n_����id And
                            ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                          '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And
                            Rownum < 2;
                    Exception
                      When Others Then
                        n_���� := 0;
                    End;
                  End If;
                  If Nvl(n_����, 0) = 0 Then
                    Begin
                      Select Count(1)
                      Into n_��Լ�ѹ���
                      From ���˹Һż�¼
                      Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                            Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                    Exception
                      When Others Then
                        n_��Լ�ѹ��� := 0;
                    End;
                    If n_�ƻ�id <> 0 Then
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ�ƻ�����
                        Where �ƻ�id = n_�ƻ�id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    Else
                      Begin
                        Select Sum(����)
                        Into n_������λ����
                        From ������λ���ſ���
                        Where ����id = n_����id And
                              ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4',
                                            '����', '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                      Exception
                        When Others Then
                          n_������λ���� := 0;
                      End;
                    End If;
                    If n_������λ���� = 0 Then
                      n_������λ���� := Null;
                    End If;
                    n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
                  
                  End If;
                End If;
              End If;
              n_������λ���� := 0;
              n_��Լ����     := 0;
              n_����         := 0;
            End If;
          End Loop;
          v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
          If n_���Ŵ��� = 1 Then
            v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                      '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                      '</YGS>' || '</PB>';
          End If;
        Else
          n_���Ŵ��� := 0;
          v_�ϰ�ʱ�� := Null;
          n_���ѹ��� := 0;
          n_�ѹ���   := 0;
          n_ʣ����   := 0;
          n_�޺���   := 0;
          n_��Լ��   := 0;
          n_��Լ��   := 0;
          If v_������λ Is Null Then
            --�Ǻ�����λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              --����Һ�
              For r_���� In (Select a.�ѹ���, a.�޺���, a.�ϰ�ʱ��
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 a.����id = n_����id And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.�ѹ���, 0);
                n_�޺���   := n_�޺��� + r_����.�޺���;
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --ԤԼ�Һ�
              For r_���� In (Select a.��Լ��, a.�޺���, a.��Լ��, a.�ϰ�ʱ��
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 a.����id = n_����id And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                n_�ѹ���   := n_�ѹ��� + Nvl(r_����.��Լ��, 0);
                n_�޺���   := n_�޺��� + Nvl(r_����.��Լ��, r_����.�޺���);
                v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                n_���Ŵ��� := 1;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_�޺��� - n_�ѹ��� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_�ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          Else
            --������λ
            If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
              For r_���� In (Select a.Id, a.�ѹ���, a.�޺���, a.��Լ��, a.�ϰ�ʱ��, a.�Ƿ���ſ���
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 a.����id = n_����id And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If r_����.�޺��� - r_����.�ѹ��� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + r_����.�޺��� - r_����.�ѹ���;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                    n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���Ŵ��� := 1;
                  Else
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                    n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.�ѹ���, 0);
                  n_ʣ����   := n_ʣ���� + r_����.�޺��� - Nvl(r_����.�ѹ���, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
            
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            Else
              --�ǵ���
              For r_���� In (Select a.Id, a.��Լ��, a.�ѹ���, a.�޺���, a.��Լ��, a.�ϰ�ʱ��, a.�Ƿ���ſ���
                           From �ٴ������¼ A, �ٴ������Դ B
                           Where a.�������� = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Դid = b.Id And
                                 Sysdate + Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� >= d_��ʼʱ�� + n_Daycount And a.ҽ��id = n_ҽ��id And
                                 a.����id = n_����id And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) And
                                 Nvl(a.�Ƿ�����, 0) = 0 And Exists
                            (Select 1
                                  From �ٴ����ﰲ�� M, �ٴ������ N
                                  Where m.Id = a.����id And m.����id = n.Id And n.����ʱ�� Is Not Null)) Loop
                Begin
                  Select ���Ʒ�ʽ
                  Into n_��Լģʽ
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1 And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լģʽ := 4;
                End;
                If n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                  Select ����
                  Into n_�޺���
                  From �ٴ�����Һſ��Ƽ�¼
                  Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1;
                  If n_��Լģʽ = 1 Then
                    n_�޺��� := n_�޺��� * Nvl(r_����.��Լ��, r_����.�޺���) / 100;
                  End If;
                  Select Count(1)
                  Into n_�ѹ���
                  From ���˹Һż�¼
                  Where �����¼id = r_����.Id And ��¼״̬ = 1 And ������λ = v_������λ;
                  If Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ�� < n_�޺��� - n_�ѹ��� Then
                    n_ʣ���� := n_ʣ���� + Nvl(r_����.��Լ��, r_����.�޺���) - r_����.��Լ��;
                  Else
                    n_ʣ���� := n_ʣ���� + n_�޺��� - n_�ѹ���;
                  End If;
                  n_���ѹ��� := n_���ѹ��� + n_�ѹ���;
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
                If n_��Լģʽ = 3 Then
                  If r_����.�Ƿ���ſ��� = 0 Then
                    --��ʱ�η���ſ���
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Select Count(1)
                      Into n_Exists
                      From �ٴ�������ſ���
                      Where ԤԼ˳��� Is Not Null And ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) <> 0;
                      If r_����.���� - n_Exists > 0 Then
                        n_ʣ����   := n_ʣ���� + r_����.���� - n_Exists;
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + n_Exists;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  Else
                    For r_���� In (Select ���
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_����.Id And ���� = 1 And ���� = v_������λ And ���� = 1) Loop
                      Begin
                        Select 1
                        Into n_Exists
                        From �ٴ�������ſ���
                        Where ��¼id = r_����.Id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0;
                      Exception
                        When Others Then
                          n_Exists := 0;
                      End;
                      If n_Exists = 1 Then
                        n_ʣ����   := n_ʣ���� + 1;
                        n_���Ŵ��� := 1;
                      Else
                        n_���ѹ��� := n_���ѹ��� + 1;
                        n_���Ŵ��� := 1;
                      End If;
                    End Loop;
                    v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  End If;
                End If;
                If n_��Լģʽ = 4 Then
                  n_���ѹ��� := n_���ѹ��� + Nvl(r_����.��Լ��, 0);
                  n_ʣ����   := n_ʣ���� + r_����.��Լ�� - Nvl(r_����.��Լ��, 0);
                  v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_����.�ϰ�ʱ��;
                  n_���Ŵ��� := 1;
                End If;
              End Loop;
              If v_�ϰ�ʱ�� Is Not Null Then
                v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
              End If;
              If n_���Ŵ��� = 1 Then
                v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                          '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� ||
                          '</YGS>' || '</PB>';
              End If;
            End If;
          End If;
        End If;
        n_Daycount := n_Daycount + 1;
      End Loop;
    End If;
    v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Docarrange;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_Third_Getnolist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ��Դ�б�(����ģʽ)
  --���:Xml_In:
  --<IN>
  --  <RQ>����</RQ>
  --  <KSID>����ID</KSID>
  --  <YSID>ҽ��ID</YSID>
  --  <YSXM>ҽ������</YSXM>
  --  <HZDW>֧����</HZDW>    //������λ�������˵�ʱ��ֻȡ������λ�ĺ�;Ϊ��ʱ��ֻȡ�Ǻ�����λ�ĺ�
  --  <FKFS>���ʽ</FKFS>  
  --  <SJJG>60</SJJG>     //ʱ����,�����򷵻����ʱ��
  --  <ZD></ZD>           //վ��
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  --  <GROUP>
  --    <RQ>����</RQ>
  --    <HBLIST>
  --     <HB>
  --        <CZJLID>1</CZJLID>     //�����¼ID
  --        <APID></APID>          //�ҺŰ���ID
  --        <JHID></JHID>          //�Һżƻ�ID
  --        <HM>235</HM>       //����
  --        <YSID>549</YSID>      //ҽ��ID
  --        <YS>����</YS>       //ҽ������
  --        <KSID>123</KSID>   //����ID
  --        <KSMC>�ڿ�</KSMC>   //��������
  --        <ZC>����ҽʦ</ZC> //ְ��
  --        <XMID>10086<XMID> //�Һ���Ŀ��ID
  --        <XMMC>�Һŷ�</XMMC> //�Һ���Ŀ������
  --        <YGHS>0</YGHS>      //�ѹҺ���
  --        <SYHS>99</SYHS>   //ʣ�����
  --        <PRICE>15</PRICE>      //�۸�
  --        <HL>��ͨ</HL>       //�Һ�����
  --        <HCXH>1</HCXH>    //�Ƿ���ڻ������ʱ��Σ�1-���� 0���߿�-������
  --        <FSD>0</FSD>      //�Ƿ��ʱ��
  --        <FWMC>����</FWMC>     //�ű�ʱ��
  --        <HBTIME>(08:00-17:59)</HBTIME> //�ɹ�ʱ��
  --     <SPANLIST>
  --            <SPAN>
  --                  <SJD></SJD>          //ʱ���,��ʽ:hh24:mi-hh24:mi
  --                  <GHZS></GHZS>      //ʱ�ιҺ�����
  --                  <SL></SL>      //ʣ������
  --            </SPAN>
  --            ����
  --          </SPANLIST>
  --      </HB>
  --      <HB>
  --      ����
  --      </HB>
  --    </HBLIST>
  --  </GROUP>
  --</OUTPUT>
  --
  --�����ʱ�Ρ�<SPANLIST>�͡�ʣ�������<SYHS>�ڵ�˵����
  --  1.������Ű�ģʽ��
  --    1.1���÷�ʱ�Σ�ͬʱ������ſ���
  --      1.1.1���������λ
  --          1�����գ�
  --                  ���ʱ�Σ�������������ʱ����Դʣ��ĿɹҺ�ʱ�Σ������ʱ��������ú�����λʣ��Ŀ�ԤԼʱ��
  --                  ʣ�������������ú�����λʣ��Ŀ�ԤԼ����
  --          2�������Ժ�
  --                  ���ʱ�Σ�������������ʱ����Դʣ��Ŀ�ԤԼʱ�Σ������ʱ��������ú�����λʣ��Ŀ�ԤԼʱ��
  --                  ʣ�������������ú�����λʣ��Ŀ�ԤԼ����
  --      1.1.2�����������λ
  --          1�����գ�
  --                  ���ʱ�Σ���Դʣ��ĿɹҺ�ʱ��
  --                  ʣ���������Դʣ��ĿɹҺ�����
  --          2�������Ժ�
  --                  ���ʱ�Σ���Դʣ��Ŀ�ԤԼʱ��
  --                  ʣ���������Դʣ��Ŀ�ԤԼ����
  --------------------------------------------------------------------------------------------------
  x_Templet Xmltype; --ģ��XML

  d_����     �ٴ������¼.��������%Type;
  n_����id   �ҺŰ���.����id%Type;
  n_ҽ��id   �ҺŰ���.ҽ��id%Type;
  v_ҽ������ �ҺŰ���.ҽ������%Type;
  v_������λ �Һź�����λ.����%Type;
  n_ʱ���� �ҺŰ���.Ĭ��ʱ�μ��%Type;

  v_�Һ�ģʽ Varchar2(500);
  n_�Һ�ģʽ Number(3);
  d_����ʱ�� Date;
  n_ԤԼ���� �ҺŰ���.ԤԼ����%Type;
  n_�������� �ҺŰ���.ԤԼ����%Type;

  v_ʣ������ �ҺŰ���ʱ��.��������%Type;
  n_����     Number(3);
  v_Temp     Varchar2(32767); --��ʱXML
  c_Xmlmain  Clob; --��ʱXML 
  v_Xmlmain  Clob; --��ʱXML 

  d_ʱ�ο�ʼ �ҺŰ���ʱ��.��ʼʱ��%Type;
  d_ʱ�ν��� �ҺŰ���ʱ��.����ʱ��%Type;
  n_ʱ������ �ҺŰ���ʱ��.��������%Type;
  n_ʱ��ʣ�� �ҺŰ���ʱ��.��������%Type;
  n_ʱ���ѹ� �ҺŰ���ʱ��.��������%Type;

  v_����         �ҺŰ�������.������Ŀ%Type;
  v_ʱ���       Varchar2(100);
  n_��ʱ��       Number(3);
  n_����ʣ��     Number(5);
  n_�ѹ���       Number(5);
  n_��Լ�ѹ���   Number(5);
  n_�ϼƽ��     �շѼ�Ŀ.�ּ�%Type;
  n_��Լ������   Number(5);
  n_��Լʣ������ Number(5);
  n_���������� Number(5);
  n_��Լģʽ     Number(3); --��Լģʽ:1-��Լ��λ������ģʽ 0-��Լ��λָ�����ģʽ
  n_�Ǻ�Լ       Number(3);
  n_�Ƿ�Ԥ��     Number(3);

  d_��ʼʱ�� �ٴ������¼.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ������¼.��ֹʱ��%Type;
  d_�Ӻ�ʱ�� �ٴ������¼.��ʼʱ��%Type;

  n_������� Number(3);
  n_ʱ������ Number(5);
  n_Ԥ������ Number(5);
  n_����ԤԼ Number(3);
  v_Timetemp Varchar2(100);
  n_Exists   Number(5);

  v_��ͨ�ȼ�   Varchar2(100);
  v_Pricegrade Varchar2(500);
  v_վ��       ���ű�.վ��%Type;
  v_���ʽ   ҽ�Ƹ��ʽ.����%Type;
  v_��ʽ       Varchar2(20);

  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  --��ȡ���ʱ��XML
  Function Gettimexml
  (
    ʱ�ο�ʼ_In �ٴ�������ſ���.��ʼʱ��%Type,
    ʱ�ν���_In �ٴ�������ſ���.��ֹʱ��%Type,
    ʱ������_In �ٴ�������ſ���.����%Type,
    ʱ��ʣ��_In �ٴ�������ſ���.����%Type
  ) Return Varchar2 Is
    v_Temp Varchar2(4000);
  Begin
    v_Temp := '';
    v_Temp := v_Temp || '<SPAN>';
    v_Temp := v_Temp || '<SJD>' || To_Char(ʱ�ο�ʼ_In, 'hh24:mi:ss') || '-' || To_Char(ʱ�ν���_In, 'hh24:mi:ss') || '</SJD>';
    v_Temp := v_Temp || '<GHZS>' || ʱ������_In || '</GHZS>';
    v_Temp := v_Temp || '<SL>' || ʱ��ʣ��_In || '</SL>';
    v_Temp := v_Temp || '</SPAN>';
    Return v_Temp;
  End;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/KSID'),
         Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/YSXM'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/SJJG'), Extractvalue(Value(A), 'IN/ZD'), Extractvalue(Value(A), 'IN/FKFS')
  Into d_����, n_����id, n_ҽ��id, v_ҽ������, v_������λ, n_ʱ����, v_վ��, v_��ʽ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --���ڽڵ�Ϊ�յ����
  d_���� := Nvl(d_����, Trunc(Sysdate));

  If v_��ʽ Is Null Then
    Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
  Else
    Select Nvl(Max(����), v_��ʽ) Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = v_��ʽ;
  End If;
  v_Pricegrade := Zl_Get_Pricegrade(v_վ��, Null, Null, v_���ʽ);
  v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);

  v_�Һ�ģʽ := zl_GetSysParameter('�Һ��Ű�ģʽ') || '|||';
  n_�Һ�ģʽ := To_Number(Substr(v_�Һ�ģʽ, 1, Instr(v_�Һ�ģʽ, '|') - 1));
  If n_�Һ�ģʽ = 1 Then
    v_�Һ�ģʽ := Substr(v_�Һ�ģʽ, Instr(v_�Һ�ģʽ, '|') + 1);
    v_Temp     := Substr(v_�Һ�ģʽ, 1, Instr(v_�Һ�ģʽ, '|') - 1);
    d_����ʱ�� := To_Date(Nvl(v_Temp, '1900-01-01'), 'yyyy-mm-dd hh24:mi:ss');
    If d_���� < d_����ʱ�� Then
      n_�Һ�ģʽ := 0;
    End If;
  End If;

  n_ԤԼ���� := Nvl(zl_GetSysParameter(66), 7);
  c_Xmlmain  := '';

  --===========================================================================================
  --�ƻ��Ű�ģʽ 
  --===========================================================================================
  If n_�Һ�ģʽ = 0 Then
    Select Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
    Into v_����
    From Dual;
    n_��Լʣ������ := 0;
  
    For r_No In (Select a.����id, a.����, a.��������, a.ҽ������, a.ҽ��id, a.ְ��, a.����, a.����id, a.�ƻ�id, a.�Ű�, a.��Ŀid, a.��Ŀ����, a.��ſ���,
                        a.�޺���, a.��Լ��, a.ԤԼ����, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��, Nvl(Hz.�����ѽ���, 0) As �ѽ���,
                        Sum(b.�ּ�) As �۸�
                 From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����,
                               Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��, Ap.ԤԼ����
                        From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                      Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                      Decode(To_Char(d_����, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4', Ap.����, '5', Ap.����,
                                              '6', Ap.����, '7', Ap.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺���,
                                      Nvl(Ap.ԤԼ����, n_ԤԼ����) As ԤԼ����
                               From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                               Where Ap.����id = Bm.Id(+) And Decode(Nvl(n_����id, 0), 0, 0, Ap.����id) = Nvl(n_����id, 0) And
                                     Decode(Nvl(n_ҽ��id, 0), 0, 0, Ap.ҽ��id) = Nvl(n_ҽ��id, 0) And
                                     Decode(Nvl(v_ҽ������, '-'), '-', '-', Ap.ҽ������) = Nvl(v_ҽ������, '-') And Ap.ͣ������ Is Null And
                                     d_���� Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                     Nvl(Ap.��ֹʱ��, To_Date('3000 - 01 - 01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                     Xz.������Ŀ(+) = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                         '����', '6', '����', '7', '����', Null) And Not Exists
                                (Select Rownum
                                      From �ҺŰ���ͣ��״̬ Ty
                                      Where Ty.����id = Ap.Id And d_���� Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And Not Exists
                                (Select Rownum
                                      From �ҺŰ��żƻ� Jh
                                      Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                            d_���� Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                            Jh.ʧЧʱ��)
                               Union All
                               Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id, Jh.Id As �ƻ�id,
                                      Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                      Decode(To_Char(d_����, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4', Jh.����, '5', Jh.����,
                                              '6', Jh.����, '7', Jh.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺���,
                                      Nvl(Ap.ԤԼ����, n_ԤԼ����) As ԤԼ����
                               From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                               Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And
                                     Decode(Nvl(n_����id, 0), 0, 0, Ap.����id) = Nvl(n_����id, 0) And
                                     Decode(Nvl(n_ҽ��id, 0), 0, 0, Jh.ҽ��id) = Nvl(n_ҽ��id, 0) And
                                     Decode(Nvl(v_ҽ������, '-'), '-', '-', Jh.ҽ������) = Nvl(v_ҽ������, '-') And
                                     d_���� Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                     Jh.ʧЧʱ�� And Xz.�ƻ�id(+) = Jh.Id And
                                     Xz.������Ŀ(+) = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                         '����', '6', '����', '7', '����', Null) And Not Exists
                                (Select Rownum
                                      From �ҺŰ���ͣ��״̬ Ty
                                      Where Ty.����id = Ap.Id And d_���� Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                     (Jh.��Чʱ��, Jh.����id) =
                                     (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                      From �ҺŰ��żƻ� Sxjh
                                      Where Sxjh.���ʱ�� Is Not Null And
                                            d_���� Between Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                            Sxjh.ʧЧʱ�� And Sxjh.����id = Jh.����id
                                      Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                        Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                      ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                 Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_����) And a.��Ŀid = b.�շ�ϸĿid And
                       Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-Mm-DD')) > d_���� And b.ִ������ <= d_���� And
                       (b.�۸�ȼ� = v_��ͨ�ȼ� Or (b.�۸�ȼ� Is Null And Not Exists
                        (Select 1
                                             From �շѼ�Ŀ
                                             Where �շ�ϸĿid = b.�շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_���� Between ִ������ And
                                                   Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                 Group By a.����id, a.����, a.��������, a.ҽ������, a.ҽ��id, a.ְ��, a.����, a.����id, a.�ƻ�id, a.�Ű�, a.��Ŀid, a.��Ŀ����, a.��ſ���,
                          a.�޺���, a.��Լ��, a.ԤԼ����, Nvl(Hz.�ѹ���, 0), Nvl(Hz.��Լ��, 0), Nvl(Hz.�����ѽ���, 0)) Loop
      Zl_�Һ����״̬_Delete(1, r_No.����);
      If Sysdate + Nvl(r_No.ԤԼ����, n_ԤԼ����) >= d_���� Then
        If r_No.�ƻ�id <> 0 Then
          Select Sign(Count(Rownum))
          Into n_��ʱ��
          From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
          Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_No.�ƻ�id And
                Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                               '����', Null) And Rownum < 2;
        Else
          Select Sign(Count(Rownum))
          Into n_��ʱ��
          From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
          Where Ap.Id = Sd.����id And Ap.Id = r_No.����id And
                Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                               '����', Null) And Rownum < 2;
        End If;
        --��ʱ�β���ſ��Ƶ����Ϊ��ͨ��
        If Trunc(Sysdate) = Trunc(d_����) And n_��ʱ�� = 1 And r_No.��ſ��� = 0 Then
          n_��ʱ�� := 0;
        End If;
        If n_��ʱ�� = 0 Then
          v_Temp := '';
          If v_������λ Is Not Null And r_No.��ſ��� = 1 Then
            If r_No.�ƻ�id <> 0 Then
              Select Nvl(Sum(����), 0)
              Into n_��Լ������
              From ������λ�ƻ�����
              Where �ƻ�id = r_No.�ƻ�id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null);
              Select Count(1)
              Into n_��Լģʽ
              From ������λ�ƻ�����
              Where �ƻ�id = r_No.�ƻ�id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And ��� = 0;
            Else
              Select Nvl(Sum(����), 0)
              Into n_��Լ������
              From ������λ���ſ���
              Where ����id = r_No.����id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null);
              Select Count(1)
              Into n_��Լģʽ
              From ������λ���ſ���
              Where ����id = r_No.����id And ������λ = v_������λ And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And ��� = 0;
            End If;
            If n_��Լģʽ = 0 Then
              If r_No.�ƻ�id <> 0 Then
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼ A
                Where �ű� = r_No.���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From ������λ�ƻ�����
                       Where �ƻ�id = r_No.�ƻ�id And ������λ = v_������λ And
                             ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                           '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
              Else
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼ A
                Where �ű� = r_No.���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And
                      Exists
                 (Select 1
                       From ������λ���ſ���
                       Where ����id = r_No.����id And ������λ = v_������λ And
                             ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                           '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
              End If;
            Else
              Select Count(1)
              Into n_��Լ�ѹ���
              From ���˹Һż�¼
              Where �ű� = r_No.���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                    Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
            End If;
            If n_��Լ������ = 0 Then
              n_��Լʣ������ := 0;
            Else
              n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
              If n_��Լʣ������ > (Nvl(r_No.�޺���, 0) - r_No.�ѹ���) Then
                n_��Լʣ������ := Nvl(r_No.�޺���, 0) - r_No.�ѹ���;
              End If;
            End If;
          End If;
        Else
          v_Temp := '<SPANLIST>';
          If r_No.�ƻ�id <> 0 Then
            Select To_Date(To_Char(d_����, 'yyyy-mm-dd') || To_Char(Max(����ʱ��), 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
            Into d_�Ӻ�ʱ��
            From �Һżƻ�ʱ��
            Where �ƻ�id = r_No.�ƻ�id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                   '����', '6', '����', '7', '����', Null);
            If r_No.��ſ��� = 1 Then
              If Trunc(d_����) = Trunc(Sysdate) Then
                n_����ԤԼ := 0;
              Else
                Select Nvl(Max(Jh.�Ƿ�ԤԼ), 0)
                Into n_����ԤԼ
                From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                       From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                       Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_No.�ƻ�id And
                             Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                            '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And
                      Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1;
              End If;
            
              d_ʱ�ο�ʼ := Null;
              d_ʱ�ν��� := Null;
              n_ʱ������ := 0;
              n_ʱ��ʣ�� := 0;
              For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��,
                                    Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����,
                                    Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                    Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_No.�ƻ�id And Not Exists
                                     (Select 1
                                           From �ҺŰ���ͣ��״̬
                                           Where ����id = Jh.����id And
                                                 To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                             Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.���
                             Order By ���) Loop
                If Nvl(n_ʱ����, 0) <> 0 Then
                  If d_ʱ�ο�ʼ Is Null Then
                    d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                    d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                    n_ʱ������ := n_ʱ������ + r_Time.��������;
                    If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                      d_ʱ�ν��� := d_�Ӻ�ʱ��;
                    End If;
                  Else
                    If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' ||
                                    '<SL>' || n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                      n_ʱ������ := r_Time.��������;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                      d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                      If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                        d_ʱ�ν��� := d_�Ӻ�ʱ��;
                      End If;
                    Else
                      n_ʱ������ := n_ʱ������ + r_Time.��������;
                    End If;
                  End If;
                End If;
                If v_������λ Is Not Null Then
                  Select Nvl(Max(1), 0)
                  Into n_��Լģʽ
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_No.�ƻ�id And ��� = 0 And ������λ = v_������λ;
                Else
                  n_��Լģʽ := 0;
                End If;
                If r_Time.ʣ���� = 0 Then
                  n_����ʣ�� := 0;
                Else
                  n_����ʣ�� := r_Time.��������;
                End If;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_No.�ƻ�id And ��� = r_Time.��� And Rownum < 2;
                  If n_Exists = 0 And r_Time.ʧЧʱ�� <> 1 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_No.���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      If n_�Ƿ�Ԥ�� = 0 Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                        End If;
                        n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                      End If;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_No.�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ�ƻ�����
                    Where �ƻ�id = r_No.�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If (n_Exists = 1 Or n_�Ǻ�Լ = 1) And r_Time.ʧЧʱ�� <> 1 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_No.���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      If n_�Ƿ�Ԥ�� = 0 Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                        End If;
                        n_��Լʣ������ := n_��Լʣ������ + 1;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                              To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' || '<SL>' ||
                              n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                n_ʱ������ := 0;
                n_ʱ��ʣ�� := 0;
                d_ʱ�ο�ʼ := Null;
                d_ʱ�ν��� := Null;
              End If;
            Else
              n_���������� := Nvl(r_No.��Լ��, Nvl(r_No.�޺���, 0)) - Nvl(r_No.��Լ��, 0);
              n_ʱ������     := 0;
              n_ʱ��ʣ��     := 0;
              d_ʱ�ο�ʼ     := Null;
              d_ʱ�ν���     := Null;
              For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ,
                                    Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                    Jh.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                    Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                    Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_No.�ƻ�id And Not Exists
                                     (Select 1
                                           From �ҺŰ���ͣ��״̬
                                           Where ����id = Jh.����id And
                                                 To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                             Where Jh.���� = Zt.����(+) And Jh.��ʼʱ�� = Zt.����(+)
                             Group By Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ
                             Order By Jh.���) Loop
                If Nvl(n_ʱ����, 0) <> 0 Then
                  If d_ʱ�ο�ʼ Is Null Then
                    d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                    d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                    n_ʱ������ := n_ʱ������ + r_Time.��������;
                    If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                      d_ʱ�ν��� := d_�Ӻ�ʱ��;
                    End If;
                  Else
                    If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' ||
                                    '<SL>' || n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                      n_ʱ������ := r_Time.��������;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                      d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                      If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                        d_ʱ�ν��� := d_�Ӻ�ʱ��;
                      End If;
                    Else
                      n_ʱ������ := n_ʱ������ + r_Time.��������;
                    End If;
                  End If;
                End If;
                If v_������λ Is Not Null Then
                  Select Nvl(Max(1), 0)
                  Into n_��Լģʽ
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_No.�ƻ�id And ��� = 0 And ������λ = v_������λ;
                Else
                  n_��Լģʽ := 0;
                End If;
                n_����ʣ�� := r_Time.ʣ����;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_No.�ƻ�id And ��� = r_Time.��� And Rownum < 2;
                  If n_Exists = 0 And r_Time.ʧЧʱ�� <> 1 Then
                    If n_���������� < n_����ʣ�� Then
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����������;
                      End If;
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                    Else
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                      End If;
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ�ƻ�����
                  Where ������Ŀ = r_Time.���� And �ƻ�id = r_No.�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ�ƻ�����
                    Where �ƻ�id = r_No.�ƻ�id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If (n_Exists = 1 Or n_�Ǻ�Լ = 1) And r_Time.ʧЧʱ�� <> 1 Then
                    If n_���������� < n_����ʣ�� Then
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����������;
                      End If;
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����������, 0);
                    Else
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                      End If;
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                              To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' || '<SL>' ||
                              n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                n_ʱ������ := 0;
                n_ʱ��ʣ�� := 0;
                d_ʱ�ο�ʼ := Null;
                d_ʱ�ν��� := Null;
              End If;
            End If;
          Else
            Select To_Date(To_Char(d_����, 'yyyy-mm-dd') || To_Char(Max(����ʱ��), 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
            Into d_�Ӻ�ʱ��
            From �ҺŰ���ʱ��
            Where ����id = r_No.����id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                   '����', '6', '����', '7', '����', Null);
            If r_No.��ſ��� = 1 Then
              If Trunc(d_����) = Trunc(Sysdate) Then
                n_����ԤԼ := 0;
              Else
                Select Nvl(Max(Ap.�Ƿ�ԤԼ), 0)
                Into n_����ԤԼ
                From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                       From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                       Where Ap.Id = Sd.����id And Ap.Id = r_No.����id And
                             Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                            '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And
                      Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1;
              End If;
              n_ʱ������ := 0;
              n_ʱ��ʣ�� := 0;
              d_ʱ�ο�ʼ := Null;
              d_ʱ�ν��� := Null;
              For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��,
                                    Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����,
                                    Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             
                             From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                    Where Ap.Id = Sd.����id And Ap.Id = r_No.����id And Not Exists
                                     (Select 1
                                           From �ҺŰ���ͣ��״̬
                                           Where ����id = Ap.Id And
                                                 To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                             Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.���
                             Order By ���) Loop
                If Nvl(n_ʱ����, 0) <> 0 Then
                  If d_ʱ�ο�ʼ Is Null Then
                    d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                    n_ʱ������ := n_ʱ������ + r_Time.��������;
                    d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                    If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                      d_ʱ�ν��� := d_�Ӻ�ʱ��;
                    End If;
                  Else
                    If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' ||
                                    '<SL>' || n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                      n_ʱ������ := r_Time.��������;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                      d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                      If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                        d_ʱ�ν��� := d_�Ӻ�ʱ��;
                      End If;
                    Else
                      n_ʱ������ := n_ʱ������ + r_Time.��������;
                    End If;
                  End If;
                End If;
                If v_������λ Is Not Null Then
                  Begin
                    Select 1
                    Into n_��Լģʽ
                    From ������λ���ſ���
                    Where ������Ŀ = r_Time.���� And ����id = r_No.����id And ��� = 0 And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_��Լģʽ := 0;
                  End;
                Else
                  n_��Լģʽ := 0;
                End If;
                If r_Time.ʣ���� = 0 Then
                  n_����ʣ�� := 0;
                Else
                  n_����ʣ�� := r_Time.��������;
                End If;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_No.����id And ��� = r_Time.��� And Rownum < 2;
                  If n_Exists = 0 And r_Time.ʧЧʱ�� <> 1 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_No.���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      If n_�Ƿ�Ԥ�� = 0 Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                        End If;
                        n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                      End If;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_No.����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ���ſ���
                    Where ����id = r_No.����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If (n_Exists = 1 Or n_�Ǻ�Լ = 1) And r_Time.ʧЧʱ�� <> 1 Then
                    If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                      Null;
                    Else
                      Select Nvl(Max(1), 0)
                      Into n_�Ƿ�Ԥ��
                      From �Һ����״̬
                      Where ״̬ In (3, 4) And ���� = r_No.���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                      If n_�Ƿ�Ԥ�� = 0 Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                        Else
                          n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                        End If;
                        n_��Լʣ������ := n_��Լʣ������ + n_����ʣ��;
                      End If;
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                              To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' || '<SL>' ||
                              n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                n_ʱ������ := 0;
                n_ʱ��ʣ�� := 0;
                d_ʱ�ο�ʼ := Null;
                d_ʱ�ν��� := Null;
              End If;
            Else
              n_���������� := Nvl(r_No.��Լ��, Nvl(r_No.�޺���, 0)) - Nvl(r_No.��Լ��, 0);
              n_ʱ������     := 0;
              n_ʱ��ʣ��     := 0;
              d_ʱ�ο�ʼ     := Null;
              d_ʱ�ν���     := Null;
              For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ,
                                    Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                    Ap.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                    Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                             From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                           To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                    From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                    Where Ap.Id = Sd.����id And Ap.Id = r_No.����id And Not Exists
                                     (Select 1
                                           From �ҺŰ���ͣ��״̬
                                           Where ����id = Ap.Id And
                                                 To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' ||
                                                         To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Between
                                                 ��ʼֹͣʱ�� And ����ֹͣʱ��) And
                                          Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                                         '5', '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                             Where Ap.���� = Zt.����(+) And Ap.��ʼʱ�� = Zt.����(+)
                             Group By Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ
                             Order By Ap.���) Loop
                If Nvl(n_ʱ����, 0) <> 0 Then
                  If d_ʱ�ο�ʼ Is Null Then
                    d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                    d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                    n_ʱ������ := n_ʱ������ + r_Time.��������;
                    If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                      d_ʱ�ν��� := d_�Ӻ�ʱ��;
                    End If;
                  Else
                    If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                      v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                                    To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' ||
                                    '<SL>' || n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                      n_ʱ������ := r_Time.��������;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                      d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                      If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                        d_ʱ�ν��� := d_�Ӻ�ʱ��;
                      End If;
                    Else
                      n_ʱ������ := n_ʱ������ + r_Time.��������;
                    End If;
                  End If;
                End If;
                If v_������λ Is Not Null Then
                  Select Nvl(Max(1), 0)
                  Into n_��Լģʽ
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_No.����id And ��� = 0 And ������λ = v_������λ;
                Else
                  n_��Լģʽ := 0;
                End If;
                n_����ʣ�� := r_Time.ʣ����;
                If v_������λ Is Null Or n_��Լģʽ = 1 Then
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_No.����id And ��� = r_Time.��� And Rownum < 2;
                  If n_Exists = 0 And r_Time.ʧЧʱ�� <> 1 Then
                    If n_���������� < n_����ʣ�� Then
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����������;
                      End If;
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                    Else
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                      End If;
                      n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                    End If;
                  End If;
                Else
                  Select Nvl(Max(1), 0)
                  Into n_Exists
                  From ������λ���ſ���
                  Where ������Ŀ = r_Time.���� And ����id = r_No.����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
                  Begin
                    Select 0
                    Into n_�Ǻ�Լ
                    From ������λ���ſ���
                    Where ����id = r_No.����id And ������λ = v_������λ And Rownum < 2;
                  Exception
                    When Others Then
                      n_�Ǻ�Լ := 1;
                  End;
                  If (n_Exists = 1 Or n_�Ǻ�Լ = 1) And r_Time.ʧЧʱ�� <> 1 Then
                    If n_���������� < n_����ʣ�� Then
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����������;
                      End If;
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����������, 0);
                    Else
                      If Nvl(n_ʱ����, 0) = 0 Then
                        v_Temp := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                      Else
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + n_����ʣ��;
                      End If;
                      n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                    End If;
                  End If;
                End If;
              End Loop;
              If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_ʱ�ο�ʼ, 'hh24:mi:ss') || '-' ||
                              To_Char(d_ʱ�ν���, 'hh24:mi:ss') || '</SJD>' || '<GHZS>' || n_ʱ������ || '</GHZS>' || '<SL>' ||
                              n_ʱ��ʣ�� || '</SL>' || '</SPAN>';
                n_ʱ������ := 0;
                n_ʱ��ʣ�� := 0;
                d_ʱ�ο�ʼ := Null;
                d_ʱ�ν��� := Null;
              End If;
            End If;
          End If;
        End If;
        If v_������λ Is Not Null Then
          If Nvl(r_No.�ƻ�id, 0) <> 0 Then
            Begin
              Select 0
              Into n_�Ǻ�Լ
              From ������λ�ƻ�����
              Where �ƻ�id = r_No.�ƻ�id And ������λ = v_������λ And Rownum < 2;
            Exception
              When Others Then
                n_�Ǻ�Լ := 1;
            End;
          Else
            Begin
              Select 0
              Into n_�Ǻ�Լ
              From ������λ���ſ���
              Where ����id = r_No.����id And ������λ = v_������λ And Rownum < 2;
            Exception
              When Others Then
                n_�Ǻ�Լ := 1;
            End;
          End If;
        End If;
        If v_������λ Is Null Or n_�Ǻ�Լ = 1 Then
          If r_No.�޺��� = 0 Then
            v_ʣ������ := '';
          Else
            If Nvl(r_No.�ƻ�id, 0) <> 0 Then
              Select Sum(����)
              Into n_��Լ������
              From ������λ�ƻ�����
              Where �ƻ�id = r_No.�ƻ�id And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null);
            Else
              Select Sum(����)
              Into n_��Լ������
              From ������λ���ſ���
              Where ����id = r_No.����id And
                    ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null);
            End If;
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼
            Where �ű� = r_No.���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_����) And
                  Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
            Select Count(1)
            Into n_Ԥ������
            From �Һ����״̬
            Where ״̬ = 3 And ���� = r_No.���� And Trunc(����) = Trunc(d_����);
            If Trunc(d_����) = Trunc(Sysdate) Then
              If Nvl(n_��Լ������, 0) = 0 Then
                v_ʣ������ := r_No.�޺��� - r_No.�ѹ��� - r_No.��Լ�� + r_No.�ѽ��� - n_Ԥ������;
              Else
                v_ʣ������ := r_No.�޺��� - r_No.�ѹ��� - r_No.��Լ�� + r_No.�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
              End If;
              n_�ѹ��� := r_No.�ѹ���;
              If Nvl(n_ʱ������, 0) < v_ʣ������ And n_��ʱ�� <> 0 Then
                n_������� := 1;
                v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_�Ӻ�ʱ��, 'hh24:mi:ss') || '-' || '</SJD>' ||
                              '<SL>' || To_Number(v_ʣ������ - Nvl(n_ʱ������, 0) - Nvl(n_��Լʣ������, 0)) || '</SL>' || '</SPAN>';
              Else
                n_������� := 0;
              End If;
            Else
              If Nvl(n_��Լ������, 0) = 0 Then
                v_ʣ������ := r_No.��Լ�� - r_No.��Լ�� - n_Ԥ������;
                If v_ʣ������ Is Null Then
                  v_ʣ������ := r_No.�޺��� - r_No.�ѹ��� - r_No.��Լ�� + r_No.�ѽ��� - n_Ԥ������;
                End If;
              Else
                v_ʣ������ := r_No.��Լ�� - r_No.��Լ�� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
                If v_ʣ������ Is Null Then
                  v_ʣ������ := r_No.�޺��� - r_No.�ѹ��� - r_No.��Լ�� + r_No.�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
                End If;
              End If;
              n_�ѹ��� := r_No.�ѹ���;
            End If;
          End If;
        Else
          If Nvl(r_No.�ƻ�id, 0) <> 0 Then
            If v_������λ Is Not Null Then
              Select Nvl(Max(1), 0)
              Into n_��Լģʽ
              From ������λ�ƻ�����
              Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And �ƻ�id = r_No.�ƻ�id And ��� = 0 And ������λ = v_������λ;
            Else
              n_��Լģʽ := 0;
            End If;
            Select Sum(����)
            Into n_��Լ������
            From ������λ�ƻ�����
            Where �ƻ�id = r_No.�ƻ�id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                     '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
          Else
            If v_������λ Is Not Null Then
              Select Nvl(Max(1), 0)
              Into n_��Լģʽ
              From ������λ���ſ���
              Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) And ����id = r_No.����id And ��� = 0 And ������λ = v_������λ;
            Else
              n_��Լģʽ := 0;
            End If;
            Select Sum(����)
            Into n_��Լ������
            From ������λ���ſ���
            Where ����id = r_No.����id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                     '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
          End If;
          If n_��Լģʽ = 0 Then
            v_ʣ������   := n_��Լʣ������;
            n_�ѹ���     := r_No.�ѹ���;
            n_��Լ�ѹ��� := Nvl(n_��Լ������, 0) - n_��Լʣ������;
          Else
            n_�ѹ��� := r_No.�ѹ���;
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼
            Where �ű� = r_No.���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                  Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
            If Nvl(n_��Լ������, 0) = 0 Then
              v_ʣ������ := '0';
            Else
              v_ʣ������ := n_��Լ������ - n_��Լ�ѹ���;
            End If;
          End If;
        End If;
        Select To_Char(��ʼʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_No.�Ű�;
        v_ʱ��� := v_Timetemp || '-';
        Select To_Char(��ֹʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_No.�Ű�;
        v_ʱ��� := v_ʱ��� || v_Timetemp;
        If v_Temp Is Not Null Then
          v_Temp := v_Temp || '</SPANLIST>';
        End If;
        If v_������λ Is Not Null Then
          If Nvl(r_No.�ƻ�id, 0) <> 0 Then
            Select Nvl(Max(1), 0)
            Into n_����
            From ������λ�ƻ�����
            Where �ƻ�id = r_No.�ƻ�id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
          Else
            Select Nvl(Max(1), 0)
            Into n_����
            From ������λ���ſ���
            Where ����id = r_No.����id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
          End If;
        End If;
        --��Լ��=0��ԤԼ��ֹ
        If Trunc(d_����) <> Trunc(Sysdate) Then
          If r_No.��Լ�� = 0 Then
            n_���� := 1;
          End If;
        End If;
        If Nvl(n_����, 0) = 0 Then
          --���������
          n_�ϼƽ�� := r_No.�۸�;
          For r_Subfee In (Select �ּ�, ��������
                           From �շѴ�����Ŀ A, �շѼ�Ŀ B
                           Where a.����id = r_No.��Ŀid And a.����id = b.�շ�ϸĿid And d_���� Between b.ִ������ And
                                 Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                                 (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                                 (b.�۸�ȼ� Is Null And Not Exists
                                  (Select 1
                                    From �շѼ�Ŀ
                                    Where �շ�ϸĿid = b.�շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_���� Between ִ������ And
                                          Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
            n_�ϼƽ�� := n_�ϼƽ�� + r_Subfee.�ּ� * r_Subfee.��������;
          End Loop;
          If Trunc(Sysdate) = Trunc(d_����) Then
            Select Nvl(Max(1), 0)
            Into n_Exists
            From (Select ʱ���
                   From ʱ���
                   Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') < '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')) Or
                         ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                         Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                 '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'))))
            Where ʱ��� = r_No.�Ű�;
          Else
            n_Exists := 1;
          End If;
          If n_Exists = 1 Then
            If v_ʣ������ > 0 Then
              c_Xmlmain := '<HB>' || '<APID>' || r_No.����id || '</APID>' || '<JHID>' || r_No.�ƻ�id || '</JHID>' || '<HM>' ||
                           r_No.���� || '</HM>' || '<YSID>' || r_No.ҽ��id || '</YSID>' || '<YS>' || r_No.ҽ������ || '</YS>' ||
                           '<KSID>' || r_No.����id || '</KSID>' || '<KSMC>' || r_No.�������� || '</KSMC>' || '<ZC>' ||
                           r_No.ְ�� || '</ZC>' || '<XMID>' || r_No.��Ŀid || '</XMID>' || '<XMMC>' || r_No.��Ŀ���� ||
                           '</XMMC>' || '<YGHS>' || n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' ||
                           n_�ϼƽ�� || '</PRICE>' || '<HCXH>' || n_������� || '</HCXH>' || '<HL>' || r_No.���� || '</HL>' ||
                           '<FSD>' || n_��ʱ�� || '</FSD>' || '<HBTIME>' || v_ʱ��� || '</HBTIME>' || '<FWMC>' || r_No.�Ű� ||
                           '</FWMC>' || v_Temp || '</HB>';
            Else
              c_Xmlmain := '<HB>' || '<APID>' || r_No.����id || '</APID>' || '<JHID>' || r_No.�ƻ�id || '</JHID>' || '<HM>' ||
                           r_No.���� || '</HM>' || '<YSID>' || r_No.ҽ��id || '</YSID>' || '<YS>' || r_No.ҽ������ || '</YS>' ||
                           '<KSID>' || r_No.����id || '</KSID>' || '<KSMC>' || r_No.�������� || '</KSMC>' || '<ZC>' ||
                           r_No.ְ�� || '</ZC>' || '<XMID>' || r_No.��Ŀid || '</XMID>' || '<XMMC>' || r_No.��Ŀ���� ||
                           '</XMMC>' || '<YGHS>' || n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' ||
                           n_�ϼƽ�� || '</PRICE>' || '<HL>' || r_No.���� || '</HL>' || '<FSD>' || n_��ʱ�� || '</FSD>' ||
                           '<HBTIME>' || v_ʱ��� || '</HBTIME>' || '<FWMC>' || r_No.�Ű� || '</FWMC>' || '</HB>';
            End If;
            v_Xmlmain := v_Xmlmain || c_Xmlmain;
          End If;
        End If;
      End If;
      n_��Լʣ������ := 0;
      n_��Լ������   := 0;
      n_ʱ������     := 0;
      n_����         := 0;
      n_�Ǻ�Լ       := 0;
    End Loop;
  
    v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_����, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain ||
                 '</HBLIST>' || '</GROUP>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  --===========================================================================================
  --������Ű�ģʽ 
  --===========================================================================================
  n_��Լʣ������ := 0;
  n_��������     := Zl_Fun_Getappointmentdays;
  --ע������¼ͣ���ˣ������������˲���ʱ��
  --�ٴ�������ſ��� �У���ʼʱ������ֹʱ����ȵ��ǼӺŵ����
  --��¼���ʣ�1-���������¼,2-��������¼
  For r_No In (Select a.��¼����, a.��¼id, a.��Դid, b.����, b.����, a.����id, c.���� As ��������, a.��Ŀid, e.���� As ��Ŀ����, a.ҽ��id, a.ҽ������,
                      d.רҵ����ְ�� As ְ��, a.�Ű�, a.��ʼʱ��, a.��ֹʱ��, a.��ſ���, a.��ʱ��, a.ԤԼ����, a.�޺���, a.��Լ��, a.�ѹ���, a.��Լ��, a.�ѽ���,
                      a.���￪ʼʱ��, a.������ֹʱ��, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��, Nvl(b.ԤԼ����, n_ԤԼ����) + n_�������� As ԤԼ����
               From (Select 1 As ��¼����, a.Id As ��¼id, a.��Դid, a.����id, a.��Ŀid, a.ҽ��id, a.ҽ������, a.�ϰ�ʱ�� As �Ű�, a.��ʼʱ��, a.��ֹʱ��,
                             Nvl(a.�Ƿ���ſ���, 0) As ��ſ���, Nvl(a.�Ƿ��ʱ��, 0) As ��ʱ��, a.ԤԼ����, a.�޺���, Nvl(a.��Լ��, a.�޺���) As ��Լ��,
                             Nvl(a.�ѹ���, 0) As �ѹ���, Nvl(a.��Լ��, 0) As ��Լ��, Nvl(a.�����ѽ���, 0) As �ѽ���, a.���￪ʼʱ��, a.������ֹʱ��,
                             a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                      From �ٴ������¼ A
                      Where Nvl(a.�Ƿ񷢲�, 0) = 1 And Nvl(a.�Ƿ�����, 0) = 0 And
                            (a.��ʼʱ�� < Nvl(a.���￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.������ֹʱ��, a.��ʼʱ��)) And a.��ʼʱ�� > Trunc(d_����ʱ��) And
                            a.��ֹʱ�� > Sysdate And
                            (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) > Sysdate Or
                             a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ�� > Sysdate Or Exists
                             (Select 1
                              From �ٴ�������ſ���
                              Where ��¼id = a.Id And Nvl(�Ƿ�ͣ��, 0) = 0 And Nvl(a.�Ƿ���ſ���, 0) = 1 And Nvl(a.�Ƿ��ʱ��, 0) = 1 And
                                    ��ʼʱ�� <> ��ֹʱ�� And ��ʼʱ�� >= Sysdate)) And
                            Decode(Nvl(n_ҽ��id, 0), 0, 0, a.ҽ��id) = Nvl(n_ҽ��id, 0) And
                            Decode(Nvl(v_ҽ������, '-'), '-', '-', a.ҽ������) = Nvl(v_ҽ������, '-') And
                            Decode(Nvl(n_����id, 0), 0, 0, a.����id) = Nvl(n_����id, 0) And a.�������� = Trunc(d_����)
                      Union All
                      Select 2 As ��¼����, a.Id As ��¼id, a.��Դid, a.����id, a.��Ŀid, a.����ҽ��id As ҽ��id, a.����ҽ������ As ҽ������,
                             a.�ϰ�ʱ�� As �Ű�, a.��ʼʱ��, a.��ֹʱ��, Nvl(a.�Ƿ���ſ���, 0) As ��ſ���, Nvl(a.�Ƿ��ʱ��, 0) As ��ʱ��, a.ԤԼ����, a.�޺���,
                             Nvl(a.��Լ��, a.�޺���) As ��Լ��, Nvl(a.�ѹ���, 0) As �ѹ���, Nvl(a.��Լ��, 0) As ��Լ��, Nvl(a.�����ѽ���, 0) As �ѽ���,
                             a.���￪ʼʱ��, a.������ֹʱ��, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                      From �ٴ������¼ A
                      Where Nvl(a.�Ƿ񷢲�, 0) = 1 And Nvl(a.�Ƿ�����, 0) = 0 And a.��ʼʱ�� > Trunc(d_����ʱ��) And a.��ֹʱ�� > Sysdate And
                            (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) And Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) > Sysdate Or
                             a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��) And a.��ֹʱ�� > Sysdate Or Exists
                             (Select 1
                              From �ٴ�������ſ���
                              Where ��¼id = a.Id And Nvl(�Ƿ�ͣ��, 0) = 0 And Nvl(a.�Ƿ���ſ���, 0) = 1 And Nvl(a.�Ƿ��ʱ��, 0) = 1 And
                                    ��ʼʱ�� <> ��ֹʱ�� And ��ʼʱ�� >= Sysdate)) And
                            Decode(Nvl(n_ҽ��id, 0), 0, 0, a.����ҽ��id) = Nvl(n_ҽ��id, 0) And
                            Decode(Nvl(v_ҽ������, '-'), '-', '-', a.����ҽ������) = Nvl(v_ҽ������, '-') And
                            Decode(Nvl(n_����id, 0), 0, 0, a.����id) = Nvl(n_����id, 0) And a.����ҽ������ Is Not Null And
                            a.�������� = Trunc(d_����)) A, �ٴ������Դ B, ���ű� C, ��Ա�� D, �շ���ĿĿ¼ E
               Where a.��Դid = b.Id And a.����id = c.Id And a.��Ŀid = e.Id And a.ҽ��id = d.Id(+)) Loop
  
    Zl_�Һ����״̬_����_Delete(r_No.��¼id);
    v_Temp := '';
    n_���� := 0;
    If Sysdate + Nvl(r_No.ԤԼ����, n_ԤԼ����) + n_�������� >= d_���� Then
      If Trunc(d_����) = Trunc(Sysdate) Then
        --����
        If v_������λ Is Null Then
          --δ���������λ
          n_�ѹ���   := r_No.�ѹ���;
          v_ʣ������ := r_No.�޺��� - Nvl(r_No.�ѹ���, 0) - (Nvl(r_No.��Լ��, 0) - Nvl(r_No.�ѽ���, 0));
          If r_No.��ʱ�� = 1 And r_No.��ſ��� = 1 Then
            v_Temp     := '<SPANLIST>';
            n_Exists   := 0;
            n_ʱ������ := 0;
            n_ʱ��ʣ�� := 0;
            d_ʱ�ο�ʼ := Null;
            d_ʱ�ν��� := Null;
            Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_No.��¼id;
            For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                           From �ٴ�������ſ���
                           Where ��¼id = r_No.��¼id And (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                 Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                 r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                 Nvl(r_No.������ֹʱ��, Sysdate - 1)) And Nvl(�Ƿ�ͣ��, 0) = 0) Loop
              If r_Time.��ʼʱ�� > Sysdate Then
                If Nvl(n_ʱ����, 0) = 0 Then
                  If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                    n_ʱ��ʣ�� := 1;
                    n_Exists   := n_Exists + 1;
                  Else
                    n_ʱ��ʣ�� := 0;
                  End If;
                  v_Temp := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, 1, n_ʱ��ʣ��);
                Else
                  If d_ʱ�ο�ʼ Is Null Then
                    n_ʱ������ := 1;
                    n_ʱ��ʣ�� := 0;
                    d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                    d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                    If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                      d_ʱ�ν��� := d_�Ӻ�ʱ��;
                    End If;
                  Else
                    If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                      v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                    
                      n_ʱ������ := 1;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                      d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                      If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                        d_ʱ�ν��� := d_�Ӻ�ʱ��;
                      End If;
                    Else
                      n_ʱ������ := n_ʱ������ + 1;
                    End If;
                  End If;
                
                  If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                    n_ʱ��ʣ�� := n_ʱ��ʣ�� + 1;
                    n_Exists   := n_Exists + 1;
                  End If;
                End If;
              End If;
            End Loop;
          
            If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
              v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
            End If;
          
            If r_No.��¼���� = 1 And n_Exists < To_Number(v_ʣ������) Then
              v_Temp := v_Temp || Gettimexml(d_�Ӻ�ʱ��, '', v_ʣ������, To_Number(v_ʣ������) - n_Exists);
            End If;
            v_Temp := v_Temp || '</SPANLIST>';
          End If;
        Else
          --���������λ
          n_�ѹ��� := r_No.�ѹ���;
          Begin
            Select ���Ʒ�ʽ
            Into n_��Լģʽ
            From �ٴ�����Һſ��Ƽ�¼
            Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ And Rownum < 2;
          Exception
            When Others Then
              n_��Լģʽ := 4;
          End;
        
          If n_��Լģʽ = 0 Then
            n_���� := 1;
          Elsif n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
            Select ����
            Into n_��Լ������
            From �ٴ�����Һſ��Ƽ�¼
            Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ;
            If n_��Լģʽ = 1 Then
              n_��Լ������ := Floor(r_No.��Լ�� * n_��Լ������ / 100);
            End If;
          
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼
            Where �ű� = r_No.���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                  Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
            n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
          
            If r_No.�޺��� - Nvl(r_No.�ѹ���, 0) < n_��Լʣ������ Then
              v_ʣ������ := r_No.�޺��� - Nvl(r_No.�ѹ���, 0);
            Else
              v_ʣ������ := n_��Լʣ������;
            End If;
          
            If r_No.��ʱ�� = 1 And r_No.��ſ��� = 1 Then
              v_Temp     := '<SPANLIST>';
              n_Exists   := 0;
              n_ʱ������ := 0;
              n_ʱ��ʣ�� := 0;
              d_ʱ�ο�ʼ := Null;
              d_ʱ�ν��� := Null;
              Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_No.��¼id;
              For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                             From �ٴ�������ſ���
                             Where ��¼id = r_No.��¼id And (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                   Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                   r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                   Nvl(r_No.������ֹʱ��, Sysdate - 1)) And Nvl(�Ƿ�ͣ��, 0) = 0) Loop
              
                If r_Time.��ʼʱ�� > Sysdate Then
                  If Nvl(n_ʱ����, 0) = 0 Then
                    If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                      n_ʱ��ʣ�� := 1;
                      n_Exists   := n_Exists + 1;
                    Else
                      n_ʱ��ʣ�� := 0;
                    End If;
                    v_Temp := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, 1, n_ʱ��ʣ��);
                  Else
                    If d_ʱ�ο�ʼ Is Null Then
                      n_ʱ������ := 1;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                      d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                      If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                        d_ʱ�ν��� := d_�Ӻ�ʱ��;
                      End If;
                    Else
                      If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                        v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                      
                        n_ʱ������ := 1;
                        n_ʱ��ʣ�� := 0;
                        d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                        d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                        If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                          d_ʱ�ν��� := d_�Ӻ�ʱ��;
                        End If;
                      Else
                        n_ʱ������ := n_ʱ������ + 1;
                      End If;
                    End If;
                  
                    If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                      n_ʱ��ʣ�� := n_ʱ��ʣ�� + 1;
                      n_Exists   := n_Exists + 1;
                    End If;
                  End If;
                End If;
              End Loop;
            
              If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
              End If;
            
              If r_No.��¼���� = 1 And n_Exists < To_Number(v_ʣ������) Then
                v_Temp := v_Temp || Gettimexml(d_�Ӻ�ʱ��, '', v_ʣ������, To_Number(v_ʣ������) - n_Exists);
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Elsif n_��Լģʽ = 3 Then
            If r_No.��ſ��� = 0 Then
              n_�ѹ���   := r_No.�ѹ���;
              v_ʣ������ := r_No.�޺��� - Nvl(r_No.�ѹ���, 0) - (Nvl(r_No.��Լ��, 0) - Nvl(r_No.�ѽ���, 0));
            Else
              v_Temp     := '<SPANLIST>';
              n_�ѹ���   := 0;
              v_ʣ������ := 0;
              n_ʱ������ := 0;
              n_ʱ��ʣ�� := 0;
              d_ʱ�ο�ʼ := Null;
              d_ʱ�ν��� := Null;
              For r_���� In (Select ���
                           From �ٴ�����Һſ��Ƽ�¼
                           Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
              
                Begin
                  Select 1, ��ʼʱ��, ��ֹʱ��
                  Into n_Exists, d_��ʼʱ��, d_��ֹʱ��
                  From �ٴ�������ſ���
                  Where ��¼id = r_No.��¼id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0 And
                        (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                        r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And Nvl(r_No.������ֹʱ��, Sysdate - 1)) And
                        Nvl(�Ƿ�ͣ��, 0) = 0;
                Exception
                  When Others Then
                    n_Exists := 0;
                End;
              
                If n_Exists = 1 Then
                  v_ʣ������ := v_ʣ������ + 1;
                Else
                  n_�ѹ��� := n_�ѹ��� + 1;
                End If;
              
                If d_��ʼʱ�� > Sysdate Then
                  If Nvl(n_ʱ����, 0) = 0 Then
                    If n_Exists = 1 Then
                      n_ʱ��ʣ�� := 1;
                    Else
                      n_ʱ��ʣ�� := 0;
                    End If;
                    v_Temp := v_Temp || Gettimexml(d_��ʼʱ��, d_��ֹʱ��, 1, n_ʱ��ʣ��);
                  Else
                    If d_ʱ�ο�ʼ Is Null Then
                      n_ʱ������ := 1;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := d_��ʼʱ��;
                      d_ʱ�ν��� := d_��ʼʱ�� + n_ʱ���� / 24 / 60;
                    Else
                      If d_��ʼʱ�� >= d_ʱ�ν��� Then
                        v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                      
                        n_ʱ������ := 1;
                        n_ʱ��ʣ�� := 0;
                        d_ʱ�ο�ʼ := d_��ʼʱ��;
                        d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                      Else
                        n_ʱ������ := n_ʱ������ + 1;
                      End If;
                    End If;
                  
                    If n_Exists = 1 Then
                      n_ʱ��ʣ�� := n_ʱ��ʣ�� + 1;
                    End If;
                  End If;
                End If;
              End Loop;
            
              If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Elsif n_��Լģʽ = 4 Then
            n_�ѹ���   := r_No.�ѹ���;
            v_ʣ������ := r_No.�޺��� - Nvl(r_No.�ѹ���, 0) - (Nvl(r_No.��Լ��, 0) - Nvl(r_No.�ѽ���, 0));
            If r_No.��ʱ�� = 1 And r_No.��ſ��� = 1 Then
              v_Temp     := '<SPANLIST>';
              n_Exists   := 0;
              n_ʱ������ := 0;
              n_ʱ��ʣ�� := 0;
              d_ʱ�ο�ʼ := Null;
              d_ʱ�ν��� := Null;
              Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_No.��¼id;
              For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                             From �ٴ�������ſ���
                             Where ��¼id = r_No.��¼id And (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                   Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                   r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                   Nvl(r_No.������ֹʱ��, Sysdate - 1)) And Nvl(�Ƿ�ͣ��, 0) = 0) Loop
              
                If r_Time.��ʼʱ�� > Sysdate Then
                  If Nvl(n_ʱ����, 0) = 0 Then
                    If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                      n_ʱ��ʣ�� := 1;
                      n_Exists   := n_Exists + 1;
                    Else
                      n_ʱ��ʣ�� := 0;
                    End If;
                    v_Temp := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, 1, n_ʱ��ʣ��);
                  Else
                    If d_ʱ�ο�ʼ Is Null Then
                      n_ʱ������ := 1;
                      n_ʱ��ʣ�� := 0;
                      d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                      d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                      If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                        d_ʱ�ν��� := d_�Ӻ�ʱ��;
                      End If;
                    Else
                      If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                        v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                      
                        n_ʱ������ := 1;
                        n_ʱ��ʣ�� := 0;
                        d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                        d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                        If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                          d_ʱ�ν��� := d_�Ӻ�ʱ��;
                        End If;
                      Else
                        n_ʱ������ := n_ʱ������ + 1;
                      End If;
                    End If;
                  
                    If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                      n_ʱ��ʣ�� := n_ʱ��ʣ�� + 1;
                      n_Exists   := n_Exists + 1;
                    End If;
                  End If;
                End If;
              End Loop;
            
              If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
              End If;
            
              If r_No.��¼���� = 1 And n_Exists < To_Number(v_ʣ������) Then
                v_Temp := v_Temp || Gettimexml(d_�Ӻ�ʱ��, '', v_ʣ������, To_Number(v_ʣ������) - n_Exists);
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          End If;
        End If;
      Else
        --ԤԼ�Һ�
        If r_No.ԤԼ���� = 1 Then
          n_���� := 1;
        Else
          --������ԤԼ
          If v_������λ Is Null Then
            If r_No.��ʱ�� = 0 Then
              n_�ѹ���   := r_No.��Լ��;
              v_ʣ������ := Nvl(r_No.��Լ��, r_No.�޺���) - Nvl(r_No.��Լ��, 0);
            Else
              --��ʱ��
              v_Temp     := '<SPANLIST>';
              n_�ѹ���   := 0;
              v_ʣ������ := 0;
              n_ʱ������ := 0;
              n_ʱ��ʣ�� := 0;
              d_ʱ�ο�ʼ := Null;
              d_ʱ�ν��� := Null;
              Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_No.��¼id;
              If r_No.��ſ��� = 0 Then
                --����ſ��Ʒ�ʱ��ԤԼ
                For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬, ����
                               From �ٴ�������ſ���
                               Where ��¼id = r_No.��¼id And ԤԼ˳��� Is Null And �Ƿ�ԤԼ = 1 And
                                     (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                     Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                     r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.ͣ�￪ʼʱ��, Sysdate) And
                                     Nvl(r_No.ͣ����ֹʱ��, Sysdate - 1))) Loop
                
                  Select Count(1)
                  Into n_ʱ���ѹ�
                  From �ٴ�������ſ���
                  Where ��¼id = r_No.��¼id And ��� = r_Time.��� And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
                
                  n_�ѹ���   := n_�ѹ��� + n_ʱ���ѹ�;
                  v_ʣ������ := v_ʣ������ + (r_Time.���� - n_ʱ���ѹ�);
                
                  If r_Time.��ʼʱ�� > Sysdate Then
                    If Nvl(n_ʱ����, 0) = 0 Then
                      n_ʱ��ʣ�� := r_Time.���� - n_ʱ���ѹ�;
                      v_Temp     := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, r_Time.����, n_ʱ��ʣ��);
                    Else
                      If d_ʱ�ο�ʼ Is Null Then
                        n_ʱ������ := r_Time.����;
                        n_ʱ��ʣ�� := 0;
                        d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                        d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                        If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                          d_ʱ�ν��� := d_�Ӻ�ʱ��;
                        End If;
                      Else
                        If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                          v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                        
                          n_ʱ������ := r_Time.����;
                          n_ʱ��ʣ�� := 0;
                          d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                          d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                          If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                            d_ʱ�ν��� := d_�Ӻ�ʱ��;
                          End If;
                        Else
                          n_ʱ������ := n_ʱ������ + r_Time.����;
                        End If;
                      End If;
                    
                      n_ʱ��ʣ�� := n_ʱ��ʣ�� + r_Time.���� - n_ʱ���ѹ�;
                    End If;
                  End If;
                End Loop;
              
                If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                  v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                End If;
              Else
                --��ſ��Ʒ�ʱ��ԤԼ
                For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                               From �ٴ�������ſ���
                               Where ��¼id = r_No.��¼id And �Ƿ�ԤԼ = 1 And
                                     (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                     Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                     r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                     Nvl(r_No.������ֹʱ��, Sysdate - 1)) And Nvl(�Ƿ�ͣ��, 0) = 0) Loop
                
                  If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                    v_ʣ������ := v_ʣ������ + 1;
                  Else
                    n_�ѹ��� := n_�ѹ��� + 1;
                  End If;
                
                  If r_Time.��ʼʱ�� > Sysdate Then
                    If Nvl(n_ʱ����, 0) = 0 Then
                      If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                        n_ʱ��ʣ�� := 1;
                      Else
                        n_ʱ��ʣ�� := 0;
                      End If;
                      v_Temp := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, 1, n_ʱ��ʣ��);
                    Else
                      If d_ʱ�ο�ʼ Is Null Then
                        n_ʱ������ := 1;
                        n_ʱ��ʣ�� := 0;
                        d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                        d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                        If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                          d_ʱ�ν��� := d_�Ӻ�ʱ��;
                        End If;
                      Else
                        If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                          v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                        
                          n_ʱ������ := 1;
                          n_ʱ��ʣ�� := 0;
                          d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                          d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                          If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                            d_ʱ�ν��� := d_�Ӻ�ʱ��;
                          End If;
                        Else
                          n_ʱ������ := n_ʱ������ + 1;
                        End If;
                      End If;
                    
                      If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                        n_ʱ��ʣ�� := n_ʱ��ʣ�� + 1;
                      End If;
                    End If;
                  End If;
                End Loop;
              
                If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                  v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                End If;
              End If;
              v_Temp := v_Temp || '</SPANLIST>';
            End If;
          Else
            --������λԤԼ�Һ�
            If r_No.ԤԼ���� = 2 Then
              n_���� := 1;
            Else
              Begin
                Select ���Ʒ�ʽ
                Into n_��Լģʽ
                From �ٴ�����Һſ��Ƽ�¼
                Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_��Լģʽ := 4;
              End;
            
              If n_��Լģʽ = 0 Then
                n_���� := 1;
              Elsif n_��Լģʽ = 1 Or n_��Լģʽ = 2 Then
                Select ����
                Into n_��Լ������
                From �ٴ�����Һſ��Ƽ�¼
                Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ;
                If n_��Լģʽ = 1 Then
                  n_��Լ������ := Floor(r_No.��Լ�� * n_��Լ������ / 100);
                End If;
              
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼
                Where �ű� = r_No.���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                      Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
              
                n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
              
                If Nvl(r_No.��Լ��, r_No.�޺���) - Nvl(r_No.��Լ��, 0) < n_��Լʣ������ Then
                  v_ʣ������ := Nvl(r_No.��Լ��, r_No.�޺���) - Nvl(r_No.��Լ��, 0);
                Else
                  v_ʣ������ := n_��Լʣ������;
                End If;
                n_�ѹ��� := r_No.��Լ��;
              
                If r_No.��ʱ�� = 1 Then
                  v_Temp     := '<SPANLIST>';
                  n_Exists   := 0;
                  n_ʱ������ := 0;
                  n_ʱ��ʣ�� := 0;
                  d_ʱ�ο�ʼ := Null;
                  d_ʱ�ν��� := Null;
                  Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_No.��¼id;
                  If r_No.��ſ��� = 1 Then
                    --��ʱ��,��ſ���
                    For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                                   From �ٴ�������ſ���
                                   Where ��¼id = r_No.��¼id And �Ƿ�ԤԼ = 1 And
                                         (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                         r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.������ֹʱ��, Sysdate - 1)) And Nvl(�Ƿ�ͣ��, 0) = 0) Loop
                    
                      If r_Time.��ʼʱ�� > Sysdate Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                            n_ʱ��ʣ�� := 1;
                            n_Exists   := n_Exists + 1;
                          Else
                            n_ʱ��ʣ�� := 0;
                          End If;
                          v_Temp := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, 1, n_ʱ��ʣ��);
                        Else
                          If d_ʱ�ο�ʼ Is Null Then
                            n_ʱ������ := 1;
                            n_ʱ��ʣ�� := 0;
                            d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                            d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                            If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                              d_ʱ�ν��� := d_�Ӻ�ʱ��;
                            End If;
                          Else
                            If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                              v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                            
                              n_ʱ������ := 1;
                              n_ʱ��ʣ�� := 0;
                              d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                              d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                              If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                                d_ʱ�ν��� := d_�Ӻ�ʱ��;
                              End If;
                            Else
                              n_ʱ������ := n_ʱ������ + 1;
                            End If;
                          End If;
                        
                          If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                            n_ʱ��ʣ�� := 1;
                            n_Exists   := n_Exists + 1;
                          End If;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                    End If;
                  
                    If n_Exists < To_Number(v_ʣ������) Then
                      v_ʣ������ := n_Exists;
                    End If;
                  Else
                    --��ʱ��,����ſ���
                    For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬, ����
                                   From �ٴ�������ſ���
                                   Where ��¼id = r_No.��¼id And ԤԼ˳��� Is Null And �Ƿ�ԤԼ = 1 And
                                         (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                         r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.ͣ�￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.ͣ����ֹʱ��, Sysdate - 1))) Loop
                    
                      Select Count(1)
                      Into n_ʱ���ѹ�
                      From �ٴ�������ſ���
                      Where ��¼id = r_No.��¼id And ��� = r_Time.��� And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
                    
                      If r_Time.��ʼʱ�� > Sysdate Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          n_ʱ��ʣ�� := r_Time.���� - n_ʱ���ѹ�;
                          n_Exists   := n_Exists + n_ʱ��ʣ��;
                          v_Temp     := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, r_Time.����, n_ʱ��ʣ��);
                        Else
                          If d_ʱ�ο�ʼ Is Null Then
                            n_ʱ������ := r_Time.����;
                            n_ʱ��ʣ�� := 0;
                            d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                            d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                            If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                              d_ʱ�ν��� := d_�Ӻ�ʱ��;
                            End If;
                          Else
                            If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                              v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                            
                              n_ʱ������ := r_Time.����;
                              n_ʱ��ʣ�� := 0;
                              d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                              d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                              If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                                d_ʱ�ν��� := d_�Ӻ�ʱ��;
                              End If;
                            Else
                              n_ʱ������ := n_ʱ������ + r_Time.����;
                            End If;
                          End If;
                        
                          n_ʱ��ʣ�� := n_ʱ��ʣ�� + (r_Time.���� - n_ʱ���ѹ�);
                          n_Exists   := n_Exists + (r_Time.���� - n_ʱ���ѹ�);
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                    End If;
                  
                    If n_Exists < To_Number(v_ʣ������) Then
                      v_ʣ������ := n_Exists;
                    End If;
                  End If;
                  v_Temp := v_Temp || '</SPANLIST>';
                End If;
              Elsif n_��Լģʽ = 3 Then
                If r_No.��ʱ�� = 0 Then
                  If r_No.��ſ��� = 0 Then
                    n_���� := 1;
                  Else
                    --��ſ��Ʋ���ʱ��
                    n_�ѹ���   := 0;
                    v_ʣ������ := 0;
                    For r_���� In (Select ���
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
                    
                      Select Count(1)
                      Into n_Exists
                      From �ٴ�������ſ���
                      Where ��¼id = r_No.��¼id And ��� = r_����.��� And �Ƿ�ԤԼ = 1 And Nvl(�Һ�״̬, 0) = 0 And
                            (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                            Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                            r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And Nvl(r_No.������ֹʱ��, Sysdate - 1)) And
                            Nvl(�Ƿ�ͣ��, 0) = 0;
                    
                      If n_Exists = 1 Then
                        v_ʣ������ := v_ʣ������ + 1;
                      Else
                        n_�ѹ��� := n_�ѹ��� + 1;
                      End If;
                    End Loop;
                  
                    If Nvl(r_No.��Լ��, r_No.�޺���) - Nvl(r_No.��Լ��, 0) < v_ʣ������ Then
                      v_ʣ������ := Nvl(r_No.��Լ��, r_No.�޺���) - Nvl(r_No.��Լ��, 0);
                    End If;
                  End If;
                Else
                  v_Temp     := '<SPANLIST>';
                  n_�ѹ���   := 0;
                  v_ʣ������ := 0;
                  n_ʱ������ := 0;
                  n_ʱ��ʣ�� := 0;
                  d_ʱ�ο�ʼ := Null;
                  d_ʱ�ν��� := Null;
                  Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_No.��¼id;
                  If r_No.��ſ��� = 0 Then
                    --��ʱ��,����ſ���
                    For r_���� In (Select ���, ����
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
                    
                      Select Count(1), Max(��ʼʱ��), Max(��ֹʱ��)
                      Into n_ʱ���ѹ�, d_��ʼʱ��, d_��ֹʱ��
                      From �ٴ�������ſ���
                      Where ��¼id = r_No.��¼id And ��� = r_����.��� And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0 And
                            (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                            Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                            r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And Nvl(r_No.������ֹʱ��, Sysdate - 1));
                    
                      n_�ѹ���   := n_�ѹ��� + n_ʱ���ѹ�;
                      v_ʣ������ := v_ʣ������ + r_����.���� - n_ʱ���ѹ�;
                    
                      If d_��ʼʱ�� > Sysdate Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          n_ʱ��ʣ�� := r_����.���� - n_ʱ���ѹ�;
                          v_Temp     := v_Temp || Gettimexml(d_��ʼʱ��, d_��ֹʱ��, r_����.����, n_ʱ��ʣ��);
                        Else
                          If d_ʱ�ο�ʼ Is Null Then
                            n_ʱ������ := r_����.����;
                            n_ʱ��ʣ�� := 0;
                            d_ʱ�ο�ʼ := d_��ʼʱ��;
                            d_ʱ�ν��� := d_��ʼʱ�� + n_ʱ���� / 24 / 60;
                            If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                              d_ʱ�ν��� := d_�Ӻ�ʱ��;
                            End If;
                          Else
                            If d_��ʼʱ�� >= d_ʱ�ν��� Then
                              v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                            
                              n_ʱ������ := r_����.����;
                              n_ʱ��ʣ�� := 0;
                              d_ʱ�ο�ʼ := d_��ʼʱ��;
                              d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                              If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                                d_ʱ�ν��� := d_�Ӻ�ʱ��;
                              End If;
                            Else
                              n_ʱ������ := n_ʱ������ + r_����.����;
                            End If;
                          End If;
                        
                          n_ʱ��ʣ�� := n_ʱ��ʣ�� + r_����.���� - n_ʱ���ѹ�;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                    End If;
                  Else
                    --��ʱ��,��ſ���
                    For r_���� In (Select ���
                                 From �ٴ�����Һſ��Ƽ�¼
                                 Where ��¼id = r_No.��¼id And ���� = 1 And ���� = 1 And ���� = v_������λ) Loop
                    
                      Select Max(1), Max(��ʼʱ��), Max(��ֹʱ��)
                      Into n_Exists, d_��ʼʱ��, d_��ֹʱ��
                      From �ٴ�������ſ���
                      Where ��¼id = r_No.��¼id And ��� = r_����.��� And Nvl(�Һ�״̬, 0) = 0 And
                            (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                            Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                            r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And Nvl(r_No.������ֹʱ��, Sysdate - 1)) And
                            Nvl(�Ƿ�ͣ��, 0) = 0;
                    
                      If n_Exists = 1 Then
                        v_ʣ������ := v_ʣ������ + 1;
                      Else
                        n_�ѹ��� := n_�ѹ��� + 1;
                      End If;
                    
                      If d_��ʼʱ�� > Sysdate Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          If n_Exists = 1 Then
                            n_ʱ��ʣ�� := 1;
                          Else
                            n_ʱ��ʣ�� := 0;
                          End If;
                          v_Temp := v_Temp || Gettimexml(d_��ʼʱ��, d_��ֹʱ��, 1, n_ʱ��ʣ��);
                        Else
                          If d_ʱ�ο�ʼ Is Null Then
                            n_ʱ������ := 1;
                            n_ʱ��ʣ�� := 0;
                            d_ʱ�ο�ʼ := d_��ʼʱ��;
                            d_ʱ�ν��� := d_��ʼʱ�� + n_ʱ���� / 24 / 60;
                            If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                              d_ʱ�ν��� := d_�Ӻ�ʱ��;
                            End If;
                          Else
                            If d_��ʼʱ�� >= d_ʱ�ν��� Then
                              v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                            
                              n_ʱ������ := 1;
                              n_ʱ��ʣ�� := 0;
                              d_ʱ�ο�ʼ := d_��ʼʱ��;
                              d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                              If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                                d_ʱ�ν��� := d_�Ӻ�ʱ��;
                              End If;
                            Else
                              n_ʱ������ := n_ʱ������ + 1;
                            End If;
                          End If;
                        
                          If n_Exists = 1 Then
                            n_ʱ��ʣ�� := n_ʱ��ʣ�� + 1;
                          End If;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                    End If;
                  End If;
                  v_Temp := v_Temp || '</SPANLIST>';
                End If;
              Elsif n_��Լģʽ = 4 Then
                If r_No.��ʱ�� = 0 Then
                  n_�ѹ���   := r_No.��Լ��;
                  v_ʣ������ := Nvl(r_No.��Լ��, r_No.�޺���) - Nvl(r_No.��Լ��, 0);
                Else
                  --��ʱ��
                  v_Temp     := '<SPANLIST>';
                  n_�ѹ���   := 0;
                  v_ʣ������ := 0;
                  n_ʱ������ := 0;
                  n_ʱ��ʣ�� := 0;
                  d_ʱ�ο�ʼ := Null;
                  d_ʱ�ν��� := Null;
                  Select Max(��ֹʱ��) Into d_�Ӻ�ʱ�� From �ٴ�������ſ��� Where ��¼id = r_No.��¼id;
                  If r_No.��ſ��� = 0 Then
                    --����ſ��Ʒ�ʱ��ԤԼ
                    For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬, ����
                                   From �ٴ�������ſ���
                                   Where ��¼id = r_No.��¼id And ԤԼ˳��� Is Null And �Ƿ�ԤԼ = 1 And
                                         (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                         r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.ͣ�￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.ͣ����ֹʱ��, Sysdate - 1))) Loop
                    
                      Select Count(1)
                      Into n_ʱ���ѹ�
                      From �ٴ�������ſ���
                      Where ��¼id = r_No.��¼id And ��� = r_Time.��� And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
                    
                      n_�ѹ���   := n_�ѹ��� + n_ʱ���ѹ�;
                      v_ʣ������ := v_ʣ������ + n_ʱ��ʣ��;
                    
                      If r_Time.��ʼʱ�� > Sysdate Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          n_ʱ��ʣ�� := r_Time.���� - n_ʱ���ѹ�;
                          v_Temp     := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, r_Time.����, n_ʱ��ʣ��);
                        Else
                          If d_ʱ�ο�ʼ Is Null Then
                            n_ʱ������ := r_Time.����;
                            n_ʱ��ʣ�� := 0;
                            d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                            d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                            If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                              d_ʱ�ν��� := d_�Ӻ�ʱ��;
                            End If;
                          Else
                            If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                              v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                            
                              n_ʱ������ := r_Time.����;
                              n_ʱ��ʣ�� := 0;
                              d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                              d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                              If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                                d_ʱ�ν��� := d_�Ӻ�ʱ��;
                              End If;
                            Else
                              n_ʱ������ := n_ʱ������ + r_Time.����;
                            End If;
                          End If;
                        
                          n_ʱ��ʣ�� := n_ʱ��ʣ�� + (r_Time.���� - n_ʱ���ѹ�);
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                    End If;
                  Else
                    For r_Time In (Select ���, ��ʼʱ��, ��ֹʱ��, �Һ�״̬
                                   From �ٴ�������ſ���
                                   Where ��¼id = r_No.��¼id And �Ƿ�ԤԼ = 1 And
                                         (r_No.��¼���� = 1 And ��ʼʱ�� Not Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.������ֹʱ��, Sysdate - 1) Or
                                         r_No.��¼���� = 2 And ��ʼʱ�� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And
                                         Nvl(r_No.������ֹʱ��, Sysdate - 1)) And Nvl(�Ƿ�ͣ��, 0) = 0) Loop
                    
                      If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                        v_ʣ������ := v_ʣ������ + 1;
                      Else
                        n_�ѹ��� := n_�ѹ��� + 1;
                      End If;
                    
                      If r_Time.��ʼʱ�� > Sysdate Then
                        If Nvl(n_ʱ����, 0) = 0 Then
                          If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                            n_ʱ��ʣ�� := 1;
                          Else
                            n_ʱ��ʣ�� := 0;
                          End If;
                          v_Temp := v_Temp || Gettimexml(r_Time.��ʼʱ��, r_Time.��ֹʱ��, 1, n_ʱ��ʣ��);
                        Else
                          If d_ʱ�ο�ʼ Is Null Then
                            n_ʱ������ := 1;
                            n_ʱ��ʣ�� := 0;
                            d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                            d_ʱ�ν��� := r_Time.��ʼʱ�� + n_ʱ���� / 24 / 60;
                            If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                              d_ʱ�ν��� := d_�Ӻ�ʱ��;
                            End If;
                          Else
                            If r_Time.��ʼʱ�� >= d_ʱ�ν��� Then
                              v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                            
                              n_ʱ������ := 1;
                              n_ʱ��ʣ�� := 0;
                              d_ʱ�ο�ʼ := r_Time.��ʼʱ��;
                              d_ʱ�ν��� := d_ʱ�ο�ʼ + n_ʱ���� / 24 / 60;
                              If d_�Ӻ�ʱ�� < d_ʱ�ν��� Then
                                d_ʱ�ν��� := d_�Ӻ�ʱ��;
                              End If;
                            Else
                              n_ʱ������ := n_ʱ������ + 1;
                            End If;
                          End If;
                        
                          If Nvl(r_Time.�Һ�״̬, 0) = 0 Then
                            n_ʱ��ʣ�� := n_ʱ��ʣ�� + 1;
                          End If;
                        End If;
                      End If;
                    End Loop;
                  
                    If Nvl(n_ʱ����, 0) <> 0 And n_ʱ������ <> 0 Then
                      v_Temp := v_Temp || Gettimexml(d_ʱ�ο�ʼ, d_ʱ�ν���, n_ʱ������, n_ʱ��ʣ��);
                    End If;
                  End If;
                  v_Temp := v_Temp || '</SPANLIST>';
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    
      If Not (r_No.��ʱ�� = 1 And r_No.��ſ��� = 1) Then
        If d_���� Between Nvl(r_No.���￪ʼʱ��, Sysdate) And Nvl(r_No.������ֹʱ��, Sysdate - 1) Then
          n_���� := 1;
        End If;
        If d_���� Between Nvl(r_No.ͣ�￪ʼʱ��, Sysdate) And Nvl(r_No.ͣ����ֹʱ��, Sysdate - 1) Then
          n_���� := 1;
        End If;
      End If;
    
      If Nvl(n_����, 0) = 0 Then
        n_�ϼƽ�� := 0;
        For r_Fee In (Select b.�ּ�, a.��������
                      From �շѴ�����Ŀ A, �շѼ�Ŀ B
                      Where a.����id = b.�շ�ϸĿid And a.����id = r_No.��Ŀid And d_���� Between b.ִ������ And
                            Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                            (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                            (b.�۸�ȼ� Is Null And Not Exists
                             (Select 1
                               From �շѼ�Ŀ
                               Where �շ�ϸĿid = b.�շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_���� Between ִ������ And
                                     Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                      Union All
                      Select b.�ּ�, 1 As ��������
                      From �շ���ĿĿ¼ A, �շѼ�Ŀ B
                      Where a.Id = b.�շ�ϸĿid And a.Id = r_No.��Ŀid And d_���� Between b.ִ������ And
                            Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                            (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                            (b.�۸�ȼ� Is Null And Not Exists
                             (Select 1
                               From �շѼ�Ŀ
                               Where �շ�ϸĿid = b.�շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_���� Between ִ������ And
                                     Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))) Loop
          n_�ϼƽ�� := n_�ϼƽ�� + r_Fee.�ּ� * r_Fee.��������;
        End Loop;
      
        v_ʱ���  := To_Char(r_No.��ʼʱ��, 'HH24:MI') || '-' || To_Char(r_No.��ֹʱ��, 'HH24:MI');
        c_Xmlmain := '<HB>';
        c_Xmlmain := c_Xmlmain || '<CZJLID>' || r_No.��¼id || '</CZJLID>';
        c_Xmlmain := c_Xmlmain || '<HM>' || r_No.���� || '</HM>';
        c_Xmlmain := c_Xmlmain || '<YSID>' || r_No.ҽ��id || '</YSID>';
        c_Xmlmain := c_Xmlmain || '<YS>' || r_No.ҽ������ || '</YS>';
        c_Xmlmain := c_Xmlmain || '<KSID>' || r_No.����id || '</KSID>';
        c_Xmlmain := c_Xmlmain || '<KSMC>' || r_No.�������� || '</KSMC>';
        c_Xmlmain := c_Xmlmain || '<ZC>' || r_No.ְ�� || '</ZC>';
        c_Xmlmain := c_Xmlmain || '<XMID>' || r_No.��Ŀid || '</XMID>';
        c_Xmlmain := c_Xmlmain || '<XMMC>' || r_No.��Ŀ���� || '</XMMC>';
        c_Xmlmain := c_Xmlmain || '<PRICE>' || n_�ϼƽ�� || '</PRICE>';
        c_Xmlmain := c_Xmlmain || '<HL>' || r_No.���� || '</HL>';
        c_Xmlmain := c_Xmlmain || '<FSD>' || r_No.��ʱ�� || '</FSD>';
        c_Xmlmain := c_Xmlmain || '<HBTIME>' || v_ʱ��� || '</HBTIME>';
        c_Xmlmain := c_Xmlmain || '<FWMC>' || r_No.�Ű� || '</FWMC>';
        If Trunc(Sysdate) = Trunc(d_����) Or r_No.��Լ�� < r_No.��Լ�� Then
          c_Xmlmain := c_Xmlmain || '<YGHS>' || n_�ѹ��� || '</YGHS>';
          c_Xmlmain := c_Xmlmain || '<SYHS>' || v_ʣ������ || '</SYHS>';
          c_Xmlmain := c_Xmlmain || v_Temp;
        Else
          c_Xmlmain := c_Xmlmain || '<YGHS>' || r_No.��Լ�� || '</YGHS>';
          c_Xmlmain := c_Xmlmain || '<SYHS>' || 0 || '</SYHS>';
        End If;
        c_Xmlmain := c_Xmlmain || '</HB>';
        v_Xmlmain := v_Xmlmain || c_Xmlmain;
      End If;
    End If;
  End Loop;
  v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_����, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain || '</HBLIST>' ||
               '</GROUP>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getnolist;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_���������Һ�_Insert
(
  ������ʽ_In      Integer,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          �ҺŰ���.����%Type,
  ����_In          �Һ����״̬.���%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���㷽ʽ_In      Varchar2,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ������λ_In      �Һź�����λ.����%Type,
  �ҺŽ��ϼ�_In  ������ü�¼.ʵ�ս��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type,
  ����˵��_In      ����Ԥ����¼.����˵��%Type,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  �������״̬_In  Number := 0,
  �Ƿ������豸_In  Number := 0,
  ����id_In        ������ü�¼.����id%Type := Null,
  ��������_In      Number := 0,
  ���ս���_In      Varchar2 := Null,
  ��Ԥ��_In        Number := Null,
  ֧������_In      ����Ԥ����¼.����%Type := Null,
  �˺�����_In      Number := 1,
  �ѱ�_In          ������ü�¼.�ѱ�%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  ������_In        �Һ����״̬.������%Type := Null,
  ��������_In      Number := 0,
  ������_In      Number := 0,
  �����¼id_In    �ٴ������¼.Id%Type := Null,
  ���ʷ���_In      Number := 0,
  ���ʽ_In      ҽ�Ƹ��ʽ.����%Type := Null
) As
  --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�)
  --���:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
  --      ���㷽ʽ_IN:֧�ֶ��ֽ��㷽ʽ,���ֽ��㷽ʽʱ�������ʽ����:���㷽ʽ����1,���,�������,��������־|���㷽ʽ����2,���,�������,��������־|...
  --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
  --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
  --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
  --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  --      ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  Err_Item Exception;
  Err_Special Exception;
  v_Err_Msg            Varchar2(255);
  n_��ӡid             Ʊ�ݴ�ӡ����.Id%Type;
  n_����ֵ             ����Ԥ����¼.���%Type;
  v_�ŶӺ���           Varchar2(20);
  v_��������           �ŶӽкŶ���.��������%Type;
  n_Ԥ��id             ����Ԥ����¼.Id%Type;
  n_�Һ�id             ���˹Һż�¼.Id%Type;
  v_��������           Varchar2(3000);
  v_��ǰ����           Varchar2(150);
  d_����ʱ��           Date;
  v_���㷽ʽ           ����Ԥ����¼.���㷽ʽ%Type;
  n_������           ����Ԥ����¼.��Ԥ��%Type;
  n_����ϼ�           Number(16, 5);
  n_Ԥ�����           ����Ԥ����¼.��Ԥ��%Type;
  n_��id               ����ɿ����.Id%Type;
  d_�Ŷ�ʱ��           Date;
  n_����               Number;
  n_����ԤԼ������     Number(18);
  n_��Լ����           Number(18);
  n_������λ����       Number(18);
  n_�Ƿ񿪷�           Number(1);
  n_Count              Number(18);
  n_�к�               Number(18);
  n_���               ���˹Һż�¼.����%Type;
  n_����id             ������ü�¼.Id%Type;
  n_�۸񸸺�           Number(18);
  n_ԭ��Ŀid           �շ���ĿĿ¼.Id%Type;
  n_ԭ������Ŀid       �շ���ĿĿ¼.Id%Type;
  v_����               ���˹Һż�¼.����%Type;
  n_����id             �ҺŰ���.Id%Type;
  n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
  n_��������id         ������ü�¼.��������id%Type;
  n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_����id             ���˽��ʼ�¼.Id%Type;
  v_Temp               Varchar2(500);
  n_ԤԼʱ�����       Number;
  n_ԤԼ����           Number;
  n_Exists             Number;
  n_��ʱ����ʾ         Number;
  d_ʱ�ο�ʼʱ��       Date;
  v_��Ԥ������ids      Varchar2(4000);
  v_�շ���Ŀids        Varchar2(300);
  n_ԤԼ����           ������λ�ҺŻ���.��Լ��%Type;
  n_����               ���˹Һż�¼.����%Type;
  d_�Ǽ�ʱ��           Date;
  v_����Ա���         ��Ա��.���%Type;
  v_����Ա����         ��Ա��.����%Type;
  n_����               ���˹Һż�¼.����%Type;
  n_ԤԼ               Integer;
  v_����               �ҺŰ���ʱ��.����%Type;
  n_���÷�ʱ��         Integer;
  n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
  n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
  n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ���ɶ���       Number;
  d_Date               Date;
  n_�Һ����           Number;
  v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
  v_������             �Һ����״̬.������%Type;
  v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
  v_��Ż�����         �Һ����״̬.������%Type;
  n_�������           Number := 0;
  n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
  v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_�ѱ�               ������ü�¼.�ѱ�%Type;
  n_���ηѱ�           Number(3) := 0;
  n_Tmp����id          �ҺŰ���.Id%Type;
  n_�ƻ�id             �ҺŰ��żƻ�.Id%Type;
  v_����               ������Ϣ.����%Type;
  n_������λ������ģʽ Number;
  n_�����¼id         �ٴ������¼.Id%Type;
  n_�Һ�ģʽ           Number(3);
  n_ͬ���޺���         Number;
  n_ͬ����Լ��         Number;
  n_ͬԴ�޺���         Number;
  n_���˹Һſ�����     Number;
  d_����ʱ��           Date;
  v_Para               Varchar2(2000);
  n_ר�ҺŹҺ�����     Number;
  n_ר�Һ�ԤԼ����     Number;
  v_վ��               ���ű�.վ��%Type;
  v_��ͨ�ȼ�           Varchar2(100);
  v_Pricegrade         Varchar2(500);
  v_ʱ���             ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��       ʱ���.��ʼʱ��%Type;
  d_������ʱ��       ʱ���.��ֹʱ��%Type;
  v_����               Varchar2(100);
  n_������Ŀid         �ҺŰ���.��Ŀid%Type;
  n_��Ŀid             �ҺŰ���.��Ŀid%Type;

  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ, a.��������, a.���֤��
    From ������Ϣ A, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);

  r_Pati c_Pati%RowType;

  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;

  Cursor c_����
  (
    v_����        �ҺŰ���.����%Type,
    d_����ʱ��_In Date
  ) Is
    Select *
    From (With ����ʱ��� As (Select ʱ���
                         From (Select ʱ���,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��,
                                       To_Date('3000-01-10 ' || To_Char(d_����ʱ��_In, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ǰʱ��,
                                       To_Date('3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��1,
                                       To_Date('3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��1
                                From ʱ���)
                         Where ��ǰʱ�� Between ��ʼʱ�� And ��ֹʱ��1 Or ��ǰʱ�� Between ��ʼʱ��1 And ��ֹʱ��)
           Select Distinct p.Id, p.����, p.����, p.����id, b.���� As ���ұ���, b.���� As ��������, p.��Ŀid, c.���� As ��Ŀ����, c.���� As ��Ŀ����,
                           p.ҽ��id, d.��� As ҽ�����, p.ҽ������, p.�޺���, p.��Լ��, p.���� As ��, p.��һ As һ, p.�ܶ� As ��, p.���� As ��,
                           p.���� As ��, p.���� As ��, p.���� As ��, p.��ſ���, p.�ƻ�id
           From (Select p.Id, p.����, p.����, p.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(p.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�, Null As �ƻ�id
                  From �ҺŰ��� P, �ҺŰ������� B
                  Where p.ͣ������ Is Null And p.Id = b.����id(+) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And
                        d_����ʱ��_In Between Nvl(p.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                        Nvl(p.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                   (Select 1
                         From �ҺŰ��żƻ�
                         Where ����id = p.Id And (d_����ʱ��_In Between ��Чʱ�� + 0 And ʧЧʱ��) And ���ʱ�� Is Not Null) And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = p.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����
                  Union All
                  Select c.Id, c.����, c.����, c.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(c.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�, p.Id As �ƻ�id
                  From �ҺŰ��żƻ� P, �ҺŰ��� C, �Һżƻ����� B,
                       (Select Max(a.��Чʱ��) As ��Ч, ����id
                         From �ҺŰ��żƻ� A, �ҺŰ��� B
                         Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                               ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                               a.ʧЧʱ�� And b.���� = ����_In
                         Group By ����id) E
                  Where p.����id = c.Id And p.Id = b.�ƻ�id(+) And p.��Чʱ�� = e.��Ч And p.����id = e.����id And
                        Nvl(p.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And (d_����ʱ��_In Between p.��Чʱ�� + 0 And p.ʧЧʱ��) And
                        p.���ʱ�� Is Not Null And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = c.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����) P, ���ű� B, �շ���ĿĿ¼ C,
                ��Ա�� D
           Where p.����id = b.Id And p.ҽ��id = d.Id(+) And p.��Ŀid = c.Id And
                 (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                 (Nvl(p.ҽ��id, 0) = 0 Or Exists
                  (Select 1
                   From ��Ա�� Q
                   Where p.ҽ��id = q.Id And (q.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or q.����ʱ�� Is Null))) And Exists
            (Select 1 From ����ʱ��� Where ʱ��� = p.�Ű�))
           Order By ����;


  r_���� c_����%RowType;

  Function Zl_����(����_In �ҺŰ���.����%Type) Return Varchar2 As
    n_���﷽ʽ �ҺŰ���.���﷽ʽ%Type;
    n_����id   �ҺŰ���.Id%Type;
    v_����     ���˹Һż�¼.����%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    If ��������_In = 2 Then
      --�Ե��ݽ��н���,���ȼ���Ƿ��������
      Select Count(Rowid) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      If n_���� = 0 Then
        v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
        Raise Err_Item;
      End If;
      Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
    End If;
  
    Begin
      Select ID, Nvl(���﷽ʽ, 0) Into n_����id, n_���﷽ʽ From �ҺŰ��� Where ���� = ����_In;
    Exception
      When Others Then
        n_����id := -1;
    End;
  
    If n_����id = -1 Then
      v_Err_Msg := '����(' || ����_In || ')δ�ҵ�!';
      Raise Err_Item;
    End If;
    --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
    v_���� := Null;
    If n_���﷽ʽ = 1 Then
      --1-ָ������
      Begin
        Select �������� Into v_���� From �ҺŰ������� Where �ű�id = n_����id;
      Exception
        When Others Then
          v_���� := Null;
      End;
    End If;
    If n_���﷽ʽ = 2 Then
      --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
      For c_���� In (Select ��������, Sum(Num) As Num
                   From (Select ��������, 0 As Num
                          From �ҺŰ�������
                          Where �ű�id = n_����id
                          Union All
                          Select ����, Count(����) As Num
                          From ���˹Һż�¼
                          Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
                                ���� In (Select �������� From �ҺŰ������� Where �ű�id = n_����id)
                          Group By ����)
                   Group By ��������
                   Order By Num) Loop
        v_���� := c_����.��������;
        Exit;
      End Loop;
    End If;
    If n_���﷽ʽ = 3 Then
    
      --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
      n_Next  := 0;
      n_First := 1;
      For c_���� In (Select Rowid As Rid, �ű�id, ��������, ��ǰ���� From �ҺŰ������� Where �ű�id = n_����id) Loop
        If n_First = 1 Then
          v_Rowid := c_����.Rid;
        End If;
        If n_Next = 1 Then
          v_���� := c_����.��������;
          Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
          Exit;
        End If;
        If Nvl(c_����.��ǰ����, 0) = 1 Then
          Update �ҺŰ������� Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_���� Is Null Then
        Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning �������� Into v_����;
      End If;
    End If;
  
    Return v_����;
  End;

  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_���������Һ�_����_Insert
  (
    ��¼id_In        �ٴ������¼.Id%Type,
    ������ʽ_In      Integer,
    ����id_In        ������ü�¼.����id%Type,
    ����_In          �ҺŰ���.����%Type,
    ����_In          �Һ����״̬.���%Type,
    ���ݺ�_In        ������ü�¼.No%Type,
    Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
    ���㷽ʽ_In      Varchar2,
    ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
    ����ʱ��_In      ������ü�¼.����ʱ��%Type,
    �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
    ������λ_In      �Һź�����λ.����%Type,
    �ҺŽ��ϼ�_In  ������ü�¼.ʵ�ս��%Type,
    ����id_In        Ʊ��ʹ����ϸ.����id%Type,
    �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
    ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In      ����Ԥ����¼.����˵��%Type,
    ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
    Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
    �����id_In      ����Ԥ����¼.�����id%Type := Null,
    �������״̬_In  Number := 0,
    �Ƿ������豸_In  Number := 0,
    ����id_In        ������ü�¼.����id%Type := Null,
    ��������_In      Number := 0,
    ���ս���_In      Varchar2 := Null,
    ��Ԥ��_In        Number := Null,
    ֧������_In      ����Ԥ����¼.����%Type := Null,
    �ѱ�_In          ������ü�¼.�ѱ�%Type := Null,
    ��Ԥ������ids_In Varchar2 := Null,
    ������_In        �Һ����״̬.������%Type := Null,
    ��������_In      Number := 0,
    ������_In      Number := 0,
    ���ʷ���_In      Number := 0,
    ���ʽ_In      ҽ�Ƹ��ʽ.����%Type := Null
  ) As
    --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�),������Ű�ģʽ��ʹ��
    --���: ������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
    --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
    --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
    --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
    --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
    --      ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
    Err_Item Exception;
    Err_Special Exception;
    v_Err_Msg  Varchar2(255);
    n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
    n_����ֵ   ����Ԥ����¼.���%Type;
    v_�ŶӺ��� Varchar2(20);
    v_�������� �ŶӽкŶ���.��������%Type;
    n_Ԥ��id   ����Ԥ����¼.Id%Type;
    n_�Һ�id   ���˹Һż�¼.Id%Type;
    v_�������� Varchar2(3000);
    v_��ǰ���� Varchar2(150);
  
    v_���㷽ʽ           ����Ԥ����¼.���㷽ʽ%Type;
    n_������           ����Ԥ����¼.��Ԥ��%Type;
    n_����ϼ�           Number(16, 5);
    n_Ԥ�����           ����Ԥ����¼.��Ԥ��%Type;
    n_��id               ����ɿ����.Id%Type;
    d_�Ŷ�ʱ��           Date;
    n_����               Number;
    n_����ԤԼ������     Number(18);
    n_��Լ����           Number(18);
    d_����ʱ��           Date;
    n_������λ����       Number(18);
    n_�Ƿ񿪷�           Number(1);
    n_Count              Number(18);
    n_�к�               Number(18);
    n_����id             ������ü�¼.Id%Type;
    n_�۸񸸺�           Number(18);
    n_ԭ��Ŀid           �շ���ĿĿ¼.Id%Type;
    n_ԭ������Ŀid       �շ���ĿĿ¼.Id%Type;
    v_����               ���˹Һż�¼.����%Type;
    n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
    n_��������id         ������ü�¼.��������id%Type;
    n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_����               ���˹Һż�¼.����%Type;
    n_����id             ���˽��ʼ�¼.Id%Type;
    v_Temp               Varchar2(500);
    v_���㷽ʽ��¼       Varchar2(1000);
    n_ԤԼʱ�����       Number;
    n_��ſ���           �ٴ������¼.�Ƿ���ſ���%Type;
    n_��Լ��             �ٴ������¼.��Լ��%Type;
    n_��Ŀid             �ٴ������¼.��Ŀid%Type;
    n_����id             �ٴ������¼.����id%Type;
    d_��ֹʱ��           �ٴ������¼.��ֹʱ��%Type;
    v_ҽ������           �ٴ������¼.ҽ������%Type;
    n_ҽ��id             �ٴ������¼.ҽ��id%Type;
    n_ԤԼ˳���         �ٴ�������ſ���.ԤԼ˳���%Type;
    n_ԤԼ����           Number;
    d_ʱ�ο�ʼʱ��       Date;
    d_ʱ����ֹʱ��       Date;
    v_�շ���Ŀids        Varchar2(300);
    n_��������־         Number;
    n_����               ���˹Һż�¼.����%Type;
    d_�Ǽ�ʱ��           Date;
    n_���ʽ��           ����Ԥ����¼.��Ԥ��%Type;
    v_�������           ����Ԥ����¼.�������%Type;
    v_����Ա���         ��Ա��.���%Type;
    v_����Ա����         ��Ա��.����%Type;
    n_ԤԼ               Integer;
    n_��ʱ����ʾ         Number;
    v_�ֽ�               ����Ԥ����¼.���㷽ʽ%Type;
    n_���÷�ʱ��         Integer;
    n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
    n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
    n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
    n_ԤԼ���ɶ���       Number;
    n_�޺���             �ٴ������¼.�޺���%Type;
    d_Date               Date;
    n_�Һ����           Number;
    v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
    v_������             �Һ����״̬.������%Type;
    v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
    v_��Ż�����         �Һ����״̬.������%Type;
    n_�������           Number := 0;
    n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
    v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
    v_�ѱ�               ������ü�¼.�ѱ�%Type;
    n_���ηѱ�           Number(3) := 0;
    v_����               ������Ϣ.����%Type;
    n_������λ������ģʽ Number;
    n_ͬ���޺���         Number;
    n_ͬ����Լ��         Number;
    n_ͬԴ�޺���         Number;
    n_���˹Һſ�����     Number;
    n_Exists             Number(5);
    v_Exists             Varchar2(4000);
    v_��Ԥ������ids      Varchar2(4000);
    n_����ҽ��id         �ٴ������¼.����ҽ��id%Type;
    v_����ҽ������       �ٴ������¼.����ҽ������%Type;
    d_���￪ʼʱ��       �ٴ������¼.���￪ʼʱ��%Type;
    d_������ֹʱ��       �ٴ������¼.������ֹʱ��%Type;
    n_ר�ҺŹҺ�����     Number;
    n_ר�Һ�ԤԼ����     Number;
    v_վ��               ���ű�.վ��%Type;
    v_��ͨ�ȼ�           Varchar2(100);
    v_Pricegrade         Varchar2(500);
    v_����               Varchar2(100);
    n_������Ŀid         �ҺŰ���.��Ŀid%Type;
  
    Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
      Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ, a.��������, a.���֤��
      From ������Ϣ A, ҽ�Ƹ��ʽ C
      Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);
  
    r_Pati c_Pati%RowType;
  
    --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
    --��ID�������ȳ��ϴ�δ����ġ�
    Cursor c_Deposit
    (
      v_����id        ������Ϣ.����id%Type,
      v_��Ԥ������ids Varchar2
    ) Is
      Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
             Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
       Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
      Group By NO, ����id
      Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;
  
    Function Zl_����(��¼id_In �ٴ������¼.Id%Type) Return Varchar2 As
      n_���﷽ʽ �ٴ������¼.���﷽ʽ%Type;
      v_����     ���˹Һż�¼.����%Type;
      v_Rowid    Varchar2(500);
      n_Next     Integer;
      n_First    Integer;
    Begin
    
      If ��������_In = 2 Then
        --�Ե��ݽ��н���,���ȼ���Ƿ��������
        Select Count(Rowid)
        Into n_����
        From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
        If n_���� = 0 Then
          v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
          Raise Err_Item;
        End If;
        Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      End If;
    
      Begin
        Select Nvl(���﷽ʽ, 0) Into n_���﷽ʽ From �ٴ������¼ Where ID = ��¼id_In;
      Exception
        When Others Then
          v_Err_Msg := '�����¼(' || ��¼id_In || ')δ�ҵ�!';
          Raise Err_Item;
      End;
    
      --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
      v_���� := Null;
      If n_���﷽ʽ = 1 Then
        --1-ָ������
        Begin
          Select b.���� Into v_���� From �ٴ��������Ҽ�¼ A, �������� B Where a.����id = b.Id And a.��¼id = ��¼id_In;
        Exception
          When Others Then
            v_���� := Null;
        End;
      End If;
      If n_���﷽ʽ = 2 Then
        --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
        For c_���� In (Select ��������, Sum(Num) As Num
                     From (Select b.���� As ��������, 0 As Num
                            From �ٴ��������Ҽ�¼ A, �������� B
                            Where a.����id = b.Id And a.��¼id = ��¼id_In
                            Union All
                            Select ����, Count(����) As Num
                            From ���˹Һż�¼
                            Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
                                  ���� In (Select d.����
                                         From �ٴ��������Ҽ�¼ C, �������� D
                                         Where c.����id = d.Id And c.��¼id = ��¼id_In)
                            Group By ����)
                     Group By ��������
                     Order By Num) Loop
          v_���� := c_����.��������;
          Exit;
        End Loop;
      End If;
      If n_���﷽ʽ = 3 Then
        --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        n_Next  := 0;
        n_First := 1;
        For c_���� In (Select a.Rowid As Rid, b.���� As ��������, a.��ǰ����
                     From �ٴ��������Ҽ�¼ A, �������� B
                     Where a.����id = b.Id And a.��¼id = ��¼id_In) Loop
          If n_First = 1 Then
            v_Rowid := c_����.Rid;
          End If;
          If n_Next = 1 Then
            v_���� := c_����.��������;
            Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
            Exit;
          End If;
          If Nvl(c_����.��ǰ����, 0) = 1 Then
            Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
            n_Next := 1;
          End If;
        End Loop;
        If v_���� Is Null Then
          Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning ����id Into v_����;
          Select ���� Into v_���� From �������� Where ID = v_����;
        End If;
      End If;
      Return v_����;
    End;
  
    Function Zl_����Ա
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
      -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
    Begin
      If Type_In = 0 Then
        --ȱʡ����
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  
  Begin
    d_����ʱ�� := ����ʱ��_In;
  
    If d_����ʱ�� Is Null Then
      d_����ʱ�� := Sysdate;
    End If;
  
    If ���ʽ_In Is Null Then
      Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
    Else
      Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
      If v_���ʽ Is Null Then
        v_���ʽ := ���ʽ_In;
      End If;
    End If;
    v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
  
    If �ѱ�_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, ����id_In) Into v_�ѱ� From Dual;
    Else
      v_�ѱ� := �ѱ�_In;
    End If;
    If v_�ѱ� Is Null Then
      n_���ηѱ� := 1;
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
    End If;
    Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
  
    If ��������_In = 1 Then
      Select Zl_Age_Calc(����id_In) Into v_���� From Dual;
      If v_���� Is Not Null Then
        Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;
      End If;
    End If;
    --��ȡ��ǰ��������
    If ������_In Is Not Null Then
      v_������ := ������_In;
    Else
      Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_ʵ�ս��ϼ� := 0;
  
    Select Count(*) + 1
    Into n_�Һ����
    From ���˹Һż�¼
    Where �����¼id = ��¼id_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
  
    If �Ǽ�ʱ��_In Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    Else
      d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
    End If;
    If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
      v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
      Raise Err_Item;
    End If;
  
    v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
    n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
    n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
    n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
    n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
    n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
    n_ͬԴ�޺���     := To_Number(Nvl(zl_GetSysParameter('����ͬһ��Դ�޹�N����', 1111), '0'));
  
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Item;
    End If;
    n_��������id := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
    n_��id       := Zl_Get��id(v_����Ա����);
  
    --֧���������ύ���
    Select Nvl(Max(1), 0)
    Into n_Exists
    From ���˹Һż�¼
    Where ����id = ����id_In And �ű� = ����_In And ���� = ����_In And ����Ա���� = v_����Ա���� And Nvl(��¼id_In, 0) = Nvl(�����¼id, 0) And
          �Ǽ�ʱ�� > Sysdate - 0.01 And ��¼״̬ = 1 And ����ʱ�� = ����ʱ��_In;
    If n_Exists = 1 Then
      v_Err_Msg := '�����Ѿ��Һ�,�����ظ�����ͬ�ĺţ�';
      Raise Err_Special;
    End If;
  
    If ������ʽ_In <> 1 Then
      --ԤԼ����Ƿ���Ӻ�����λ����
      --��������˺�����λ���� ��
      Begin
        Select 1
        Into n_������λ����
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = 1 And ���� = 1 And ���Ʒ�ʽ <> 4 And Rownum < 2;
      Exception
        When Others Then
          n_������λ���� := 0;
      End;
    End If;
  
    If ������ʽ_In <> 2 Then
      v_���� := Zl_����(��¼id_In);
    End If;
  
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*) Into n_Count From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ In (1, 3) And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    Open c_Pati(����id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '����δ�ҵ������ܼ�����';
      Raise Err_Item;
    End If;
  
    Begin
      Select Nvl(�Ƿ��ʱ��, 0), �޺���, �ѹ���, �����ѽ���, ��Լ��, �Ƿ���ſ���, ��Լ��, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���￪ʼʱ��, ������ֹʱ��
      Into n_���÷�ʱ��, n_�޺���, n_�ѹ���, n_�����ѽ���, n_��Լ��, n_��ſ���, n_��Լ��, n_��Ŀid, n_����id, n_ҽ��id, v_ҽ������, n_����ҽ��id, v_����ҽ������,
           d_���￪ʼʱ��, d_������ֹʱ��
      From �ٴ������¼
      Where ID = ��¼id_In And Nvl(�Ƿ�����, 0) = 0;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
  
    Select Min(վ��) Into v_վ�� From ���ű� Where ID = n_����id;
    v_Pricegrade := Zl_Get_Pricegrade(v_վ��, ����id_In, Null, v_���ʽ);
    v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    If ����ʱ��_In Between Nvl(d_���￪ʼʱ��, Sysdate) And Nvl(d_������ֹʱ��, Sysdate - 1) And v_����ҽ������ Is Not Null Then
      n_ҽ��id   := n_����ҽ��id;
      v_ҽ������ := v_����ҽ������;
    End If;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And �����¼id = ��¼id_In;
        If n_Count >= Nvl(n_ר�Һ�ԤԼ����, 0) And Nvl(n_ר�Һ�ԤԼ����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ���������ԤԼ����,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And �����¼id = ��¼id_In;
        If n_Count >= Nvl(n_ר�ҺŹҺ�����, 0) And Nvl(n_ר�ҺŹҺ�����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��������ŹҺ�����,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_ͬԴ�޺���, 0) <> 0 Then
      If �����¼id_In Is Null Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = ����_In;
      Else
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And �����¼id = �����¼id_In;
      End If;
      If n_Count >= Nvl(n_ͬԴ�޺���, 0) And Nvl(n_ͬԴ�޺���, 0) > 0 Then
        v_Err_Msg := 'ͬһ���������ͬʱ��(ԤԼ)[' || Nvl(n_ͬԴ�޺���, 0) || ']����ͬ�ű�ĺ�,�����ٹҺ�(ԤԼ)��';
        Raise Err_Item;
      End If;
    End If;
  
    d_Date         := Null;
    d_ʱ�ο�ʼʱ�� := Null;
  
    If Nvl(n_�޺���, 0) >= 0 Or n_�޺��� Is Null Then
      If n_���÷�ʱ�� = 1 Then
        If Nvl(n_��ſ���, 0) = 1 Then
          If Nvl(�Ƿ������豸_In, 0) = 0 Then
            Select Count(*), Max(��ʼʱ��)
            Into n_Count, d_ʱ�ο�ʼʱ��
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0);
          
            v_Temp := '�Һ�';
            If ������ʽ_In > 1 Then
              v_Temp := 'ԤԼ�Һ�';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ����в��������Ϊ' || Nvl(����_In, 0) || '�İ���,������' || v_Temp || '��';
              Raise Err_Item;
            End If;
          End If;
        
          --�����,����ѡ��Һ�
          If Trunc(Sysdate) = Trunc(����ʱ��_In) Then
            --�ҵ���ĺ�
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                         To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��, ����, �Ƿ�ԤԼ
                         From �ٴ�������ſ���
                         Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0)) Loop
              If Sysdate > v_ʱ��.��ֹʱ�� Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                Raise Err_Item;
              End If;
            End Loop;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          For v_ʱ�� In (Select ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ
                       From �ٴ�������ſ���
                       Where ��¼id = ��¼id_In And
                             (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                             '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                             ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-11 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
            n_ԤԼʱ����� := v_ʱ��.���;
            d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            d_ʱ����ֹʱ�� := v_ʱ��.��ֹʱ��;
          
            Select Count(*), Max(���), Max(ԤԼ˳���) + 1
            Into n_Count, n_ԤԼ����, n_ԤԼ˳���
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) Not In (0, 4, 5);
          
            If Nvl(n_Count, 0) > Nvl(v_ʱ��.����, 0) And ��������_In <> 2 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                           To_Char(v_ʱ��.��ֹʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.����, 0) || '��,�����ٽ���ԤԼ�Һţ�';
              Raise Err_Item;
            End If;
            n_Count := 1;
          End Loop;
        
          If n_Count = 0 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�����û����صİ��żƻ�(' || To_Char(����ʱ��_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),���ܽ���ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If ������ʽ_In = 1 And ��������_In <> 2 Then
      --�ҺŹ���:
      --  �ѹ������ܴ����޺���
      If n_�ѹ��� >= Nvl(n_�޺���, 0) And n_�޺��� Is Not Null Then
        v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(n_�޺���, 0) || '�����ٹҺţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If ������ʽ_In > 1 Then
      --ԤԼ����ؼ��
      --����:
      --   1.����Լ���ܳ�����Լ��
      --   2.����Ƿ�����ʱ�ε�
      If n_��Լ�� >= Nvl(n_��Լ��, 0) And Nvl(n_��Լ��, 0) <> 0 And n_��Լ�� Is Not Null And ��������_In <> 2 Then
        v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(n_��Լ��, 0) || '������ԤԼ�Һţ�';
        Raise Err_Item;
      End If;
      If ԤԼ��ʽ_In Is Not Null Then
        Select Zl_Fun_Get�ٴ�����ԤԼ״̬(��¼id_In, ����ʱ��_In, ����_In, ԤԼ��ʽ_In, Null, 0, v_����Ա����, v_������)
        Into v_Exists
        From Dual;
        If To_Number(Substr(v_Exists, 1, 1)) <> 0 Then
          v_Err_Msg := '�����ԤԼ��ʽ' || ԤԼ��ʽ_In || '������,ԭ��:' || Substr(v_Exists, 3);
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
      If Nvl(n_��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
        v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
        Raise Err_Item;
      End If; --Nvl(r_����.��ſ���, 0) =0
    
      --������λ����ģʽ
      Begin
        Select Nvl(���Ʒ�ʽ, 0)
        Into n_������λ������ģʽ
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And Rownum < 2;
      Exception
        When Others Then
          n_������λ������ģʽ := 4;
      End;
    
      If n_������λ������ģʽ = 0 Then
        v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || 'δ����' || ������λ_In || '��ԤԼ,���ܼ�����';
        Raise Err_Item;
      End If;
      If n_������λ������ģʽ = 1 Or n_������λ������ģʽ = 2 Then
        Select ����
        Into n_Count
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1;
        If n_������λ������ģʽ = 1 Then
          n_Count := Round(Nvl(n_��Լ��, n_�޺���) * n_Count / 100);
        End If;
        Select Count(1)
        Into n_Exists
        From ���˹Һż�¼
        Where ��¼״̬ = 1 And �����¼id = ��¼id_In And ������λ = ������λ_In;
        If n_Exists >= n_Count Then
          v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
      --������ż��
      If n_������λ������ģʽ = 3 Then
        For c_������λ In (Select ���, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And ��� = ����_In) Loop
          If n_��ſ��� = 1 Then
            Begin
              Select 1
              Into n_Count
              From �ٴ�������ſ���
              Where ��¼id = ��¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 0;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 1 Then
              n_�Ƿ񿪷� := 1;
            Else
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            End If;
          Else
            Select Count(1)
            Into n_Count
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = ����_In And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
            If n_Count >= c_������λ.���� Then
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            Else
              n_�Ƿ񿪷� := 1;
            End If;
          End If;
        End Loop;
      
        If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
          v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    --����޺�������Լ��
    n_�к�         := 1;
    n_ԭ��Ŀid     := 0;
    n_ԭ������Ŀid := 0;
    n_ʵ�ս��ϼ� := 0;
    If ��������_In <> 1 Then
      If ������ʽ_In <> 2 Then
        If Nvl(����id_In, 0) = 0 Then
          --����Ӧ�ó�����
          Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
        Else
          n_����id := ����id_In;
        End If;
      Else
        n_����id := Null;
      End If;
    End If;
  
    If Nvl(��¼id_In, 0) <> 0 Then
      v_���� := '2|' || ��¼id_In;
    End If;
    If v_���� Is Null Then
      v_���� := '3|' || ����_In;
    End If;
  
    n_������Ŀid := Zl_Custom_Getregeventitem(r_Pati.����id, r_Pati.����, r_Pati.���֤��, r_Pati.��������, r_Pati.�Ա�, r_Pati.����, v_����);
    If Nvl(n_������Ŀid, 0) <> 0 Then
      n_��Ŀid := n_������Ŀid;
    End If;
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := n_��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := n_��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������,
                          Nvl(a.��Ŀ����, 0) As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_��Ŀid And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Order By ����, ��Ŀ����, �������) Loop
      If c_Item.���� = 1 Then
        n_���� := Nvl(c_Item.����, 0);
      End If;
      n_�۸񸸺� := Null;
      If n_ԭ��Ŀid = c_Item.��Ŀid Then
        If n_ԭ������Ŀid <> c_Item.������Ŀid Then
          n_�۸񸸺� := n_�к�;
        End If;
        n_ԭ������Ŀid := c_Item.������Ŀid;
      End If;
      n_ԭ��Ŀid := c_Item.��Ŀid;
      n_Ӧ�ս�� := Round(c_Item.���� * c_Item.����, 5);
      n_ʵ�ս�� := n_Ӧ�ս��;
      If Nvl(c_Item.���ηѱ�, 0) <> 1 And n_���ηѱ� = 0 Then
        --����:
        v_Temp     := Zl_Actualmoney(r_Pati.�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_ʵ�ս�� := Zl_To_Number(v_Temp);
      End If;
      n_ʵ�ս��ϼ� := Nvl(n_ʵ�ս��ϼ�, 0) + n_ʵ�ս��;
    
      --�������ݲ���������
      If ��������_In <> 1 Then
        --�������˹Һŷ���(���ܵ����ǻ������������)
        Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
        --:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
        Values
          (n_����id, 4, Decode(������ʽ_In, 2, 0, 1), n_�к�, n_�۸񸸺�, c_Item.��������, ���ݺ�_In, Ʊ�ݺ�_In, 1, n_����, Null,
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, n_����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, n_ʵ�ս��)),
           Decode(Nvl(���ʷ���_In, 0), 1, Null, n_����id), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), n_����id, v_ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null, Null,
           ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(����ʱ��_In);
    End If;
  
    --���¹Һ����״̬
    If ��������_In <> 2 Then
      n_���� := ����_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And ��� = n_���� And Nvl(�Һ�״̬, 0) Not In (0, 5);
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 Then
        n_���� := Null;
      End If;
      If n_���÷�ʱ�� = 1 And Nvl(n_��ſ���, 0) = 1 Then
        v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
      Select Nvl(Min(���), 0)
      Into n_����
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
      If n_���� = 0 Then
        Select Nvl(Min(���), 0) Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 0;
        If n_���� = 0 Then
          Select Nvl(Max(���), 0) + 1 Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
        End If;
      End If;
    End If;
  
    If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
      If ������ʽ_In > 1 And Nvl(n_��ſ���, 0) = 0 Then
        --����:ԤԼʱ�����||ԤԼ��
        If Nvl(n_ԤԼ����, 0) = 0 Then
          v_Temp := Nvl(n_��Լ��, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_ԤԼʱ����� || v_Temp;
          n_���� := To_Number(v_Temp);
        Else
          n_���� := n_ԤԼ���� + 1;
        End If;
      End If;
    End If;
  
    If Nvl(n_��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
      --������ŵĴ���
      Begin
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where �Һ�״̬ = 5 And ��¼id = ��¼id_In And ��� = n_����;
        n_������� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_�������   := 0;
      End;
      If n_������� = 0 Then
        If n_���÷�ʱ�� = 1 And n_��ſ��� = 0 Then
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����, ��ע)
            Select ��¼id_In, n_ԤԼʱ�����, n_ԤԼ˳���, d_ʱ�ο�ʼʱ��, d_ʱ����ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1),
                   1, ������λ_In, v_����Ա����, n_����
            From Dual;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����Ա���� = v_����Ա����
          Where ��¼id = ��¼id_In And ��� = n_����;
        End If;
        If Sql%RowCount = 0 Then
          Begin
            If n_���÷�ʱ�� = 1 Then
              --��ʱ��
              If n_��ſ��� = 1 Then
                --��ſ���
                Select Max(��ֹʱ��) Into d_��ֹʱ�� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
                If Sysdate > d_��ֹʱ�� Then
                  d_��ֹʱ�� := Sysdate;
                End If;
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                  Select ��¼id_In, n_����, d_��ֹʱ��, d_��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1,
                         ������λ_In, v_����Ա����
                  From Dual;
              Else
                --��ʱ��,����ſ���
                Null;
              End If;
            Else
              --����ʱ��
              Insert Into �ٴ�������ſ���
                (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                Select ��¼id_In, n_����, ��ʼʱ��, ��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1, ������λ_In,
                       v_����Ա����
                From �ٴ�������ſ���
                Where ��¼id = ��¼id_In And ��� = 1;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_���� || '�ѱ�����' || v_������ || '����,������ѡ��һ�����.';
          Raise Err_Item;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����ʱ�� = Null
          Where ��¼id = ��¼id_In And ��� = n_���� And �Һ�״̬ = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
        End If;
      End If;
    End If;
  
    --�������ݲ������κ� ����
    If ������ʽ_In <> 2 And ��������_In <> 1 And Nvl(���ʷ���_In, 0) = 0 Then
      --�Һ�,ԤԼ�Һ��Ѿ��ۿ��
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      n_����ϼ� := 0;
      If ���ս���_In Is Not Null Then
        --�������ս���
        v_�������� := ���ս���_In || '||';
        n_����ϼ� := 0;
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
          n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
          n_����ϼ� := Nvl(n_����ϼ�, 0) + Nvl(n_������, 0);
          v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(��Ԥ��_In, 0) <> 0 Then
        --������Ԥ��
        n_����ϼ� := n_����ϼ� + Nvl(��Ԥ��_In, 0);
        n_Ԥ����� := ��Ԥ��_In;
        For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
          n_������ := Case
                      When r_Deposit.��� - n_Ԥ����� < 0 Then
                       r_Deposit.���
                      Else
                       n_Ԥ�����
                    End;
          If r_Deposit.����id = 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (r_Deposit.����id, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          --����Ƿ��Ѿ�������
          If r_Deposit.��� <= n_������ Then
            n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
          Else
            n_Ԥ����� := 0;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
        If n_Ԥ����� > 0 Then
          v_Err_Msg := 'Ԥ������֧������֧�����,���ܼ���������';
          Raise Err_Item;
        End If;
      End If;
      --ʣ�����,��ָ�����㷽֧��
      n_������ := Nvl(n_ʵ�ս��ϼ�, 0) - Nvl(n_����ϼ�, 0);
      If Nvl(n_������, 0) < 0 Then
        v_Err_Msg := '�Һŵ���ؽ�������˵�ǰʵ����,���ܼ���������';
        Raise Err_Item;
      End If;
    
      If Nvl(n_������, 0) <> 0 Or (Nvl(n_������, 0) = 0 And Nvl(��Ԥ��_In, 0) = 0) Then
        If ���㷽ʽ_In Is Null Then
          v_Err_Msg := 'δ����ָ���Ľ��㷽ʽ,���ܼ���������';
          Raise Err_Item;
        End If;
      
        If Nvl(Ԥ��id_In, 0) <> 0 Then
          --�����Ԥ��ID_In��Ҫ��Ϊ�˽����������,���ҽ������վ���˸�ID,��Ҫ���µ�ID���и���,����������ת���ID
          Update ����Ԥ����¼ Set ID = n_Ԥ��id Where ID = Nvl(Ԥ��id_In, 0);
          n_Ԥ��id := Nvl(Ԥ��id_In, 0);
        End If;
        If Instr(���㷽ʽ_In, ',') = 0 Then
          --ֻ����һ�ֽ��㷽ʽ��
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
        Else
          v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
          n_Exists       := 0;
          v_���㷽ʽ��¼ := '';
          While v_�������� Is Not Null Loop
            v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
            v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            n_���ʽ�� := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            n_��������־ := To_Number(v_��ǰ����);
          
            If Instr('|' || v_���㷽ʽ��¼ || '|', '|' || Nvl(v_���㷽ʽ, v_�ֽ�) || '|') <> 0 Then
              v_Err_Msg := 'ʹ�����ظ��Ľ��㷽ʽ,����!';
              Raise Err_Item;
            Else
              v_���㷽ʽ��¼ := v_���㷽ʽ��¼ || '|' || Nvl(v_���㷽ʽ, v_�ֽ�);
            End If;
          
            If n_��������־ = 0 Then
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
            Else
              If n_Exists = 1 Then
                v_Err_Msg := 'Ŀǰ�ҺŽ�֧��һ���������㷽ʽ,���ܼ���������';
                Raise Err_Item;
              End If;
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
              n_Exists := 1;
            End If;
          
            v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
          End Loop;
        End If;
      End If;
    
      --������Ա�ɿ�����
    
      For v_�ɿ� In (Select ���㷽ʽ, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ A
                   Where a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And Nvl(����id, 0) = Nvl(����id_In, 0)
                   Group By ���㷽ʽ) Loop
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(v_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_�ɿ�.���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, v_�ɿ�.���㷽ʽ, 1, Nvl(v_�ɿ�.��Ԥ��, 0));
          n_����ֵ := Nvl(v_�ɿ�.��Ԥ��, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = v_�ɿ�.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End Loop;
    End If;
  
    --����Һż�¼
    If ��������_In = 2 Then
      Begin
        Select ID Into n_�Һ�id From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    End If;
  
    Update ���˹Һż�¼
    Set ��¼���� = Decode(������ʽ_In, 2, 2, 1), ��¼״̬ = Decode(��������_In, 1, 0, 1), ����� = r_Pati.�����, ����Ա���� = v_����Ա����,
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
        ����ʱ�� = Case ��������_In
                  When 1 Then
                   Null
                  Else
                   Case ������ʽ_In
                     When 2 Then
                      Null
                     Else
                      d_�Ǽ�ʱ��
                   End
                End, ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), ������λ = Nvl(������λ_In, ������λ),
        ԤԼ����Ա = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա, Null), Nvl(ԤԼ����Ա, v_����Ա����)),
        ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���)), �����¼id = ��¼id_In
    Where ID = n_�Һ�id;
    If Sql%NotFound Then
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = r_Pati.���ʽ And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
         ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���, �����¼id)
      Values
        (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
         r_Pati.�Ա�, r_Pati.����, ����_In, n_����, v_����, Null, n_����id, v_ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
         Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���), ��¼id_In);
    End If;
    --�������ݲ��ܲ�������
    If ��������_In <> 1 Then
      n_ԤԼ���ɶ��� := 0;
      If ������ʽ_In > 1 Then
        n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
      End If;
      --�Һź��շѵ�ԤԼ��ֱ�ӽ������(�շ�ԤԼȱ�ٽ��չ���,����ֱ�Ӻ͹Һ�һ��ֱ�ӽ������)
      If ������ʽ_In <> 2 Or n_ԤԼ���ɶ��� = 1 Then
        If Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113)) <> 0 Then
          --�Ŷӽк�ģʽ:-0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
          If Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(n_����id, 0))) = 0 Or n_ԤԼ���ɶ��� = 1 Then
            n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(������ʽ_In, 0) > 1 And n_��ʱ����ʾ = 1 And n_���÷�ʱ�� = 1 Then
              n_��ʱ����ʾ := 1;
            Else
              n_��ʱ����ʾ := Null;
            End If;
            --��������
            --.����ִ�в��š� �ķ�ʽ���ɶ���
            v_�������� := n_����id;
            v_�ŶӺ��� := Zlgetnextqueue(n_����id, n_�Һ�id, ����_In || '|' || ����_In);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
            --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
            Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, n_����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, v_ҽ������, d_�Ŷ�ʱ��,
                             ԤԼ��ʽ_In, n_��ʱ����ʾ, v_�Ŷ����);
          End If;
        End If;
      End If;
    
      If Nvl(������ʽ_In, 0) = 1 And Nvl(���ʷ���_In, 0) = 0 Then
        --����Ʊ��ʹ�����
        If Ʊ�ݺ�_In Is Not Null Then
          Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
          --����Ʊ��
          Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Values
            (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�Ǽ�ʱ��, v_����Ա����, �ҺŽ��ϼ�_In);
          --״̬�Ķ�
          Update Ʊ�����ü�¼
          Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
          Where ID = Nvl(����id_In, 0);
        End If;
        --���˱��ξ���(�Է���ʱ��Ϊ׼)
        If Nvl(r_Pati.����id, 0) <> 0 Then
          Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = v_���� Where ����id = r_Pati.����id;
        End If;
      End If;
    End If;
    --���˹ҺŻ���
    --��������ʱ�����ٶԻ��ܵ��ݽ���ͳ���� ����������ʱ�Ѿ������˻���
    If ��������_In <> 2 Then
      --������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
      --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
      --����zl_third_lockno�������ţ�������ʹ�ñ���������
      n_ԤԼ := Case
                When Nvl(������ʽ_In, 0) = 1 Then
                 0
                When Nvl(������ʽ_In, 0) = 2 Then
                 1
                When Nvl(������ʽ_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_��Ŀid, n_����id, ����ʱ��_In, n_ԤԼ, ����_In, 0, ��¼id_In);
    End If;
  
    If ��������_In <> 1 Then
      --��Ϣ����,����ʱ��������Ϣ
      Begin
        Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
          Using 1, n_�Һ�id;
      Exception
        When Others Then
          Null;
      End;
      b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Err_Special Then
      Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_�����¼id := �����¼id_In;
  v_Para       := zl_GetSysParameter(256);
  n_�Һ�ģʽ   := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  d_����ʱ�� := ����ʱ��_In;
  If d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;

  If ���ʽ_In Is Null Then
    Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
  Else
    Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
    If v_���ʽ Is Null Then
      v_���ʽ := ���ʽ_In;
    End If;
  End If;

  If Sysdate - 10 > Nvl(d_����ʱ��, Sysdate - 30) Then
    If n_�Һ�ģʽ = 1 And Nvl(����ʱ��_In, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
      v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
      Raise Err_Item;
    End If;
  Else
    If n_�Һ�ģʽ = 1 And Nvl(����ʱ��_In, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = ����_In And Nvl(����ʱ��_In, Sysdate) Between a.��ʼʱ�� And a.��ֹʱ��;
      Exception
        When Others Then
          v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If n_�����¼id Is Not Null Then
    --������Ű�ģʽ
    Zl_���������Һ�_����_Insert(n_�����¼id, ������ʽ_In, ����id_In, ����_In, ����_In, ���ݺ�_In, Ʊ�ݺ�_In, ���㷽ʽ_In, ժҪ_In, ����ʱ��_In, �Ǽ�ʱ��_In,
                        ������λ_In, �ҺŽ��ϼ�_In, ����id_In, �շ�Ʊ��_In, ������ˮ��_In, ����˵��_In, ԤԼ��ʽ_In, Ԥ��id_In, �����id_In, �������״̬_In,
                        �Ƿ������豸_In, ����id_In, ��������_In, ���ս���_In, ��Ԥ��_In, ֧������_In, �ѱ�_In, ��Ԥ������ids_In, ������_In, ��������_In,
                        ������_In, ���ʷ���_In, ���ʽ_In);
  Else
    v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
    v_Temp          := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          Null;
      End;
      If ����ʱ��_In > d_����ʱ�� Then
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
        Raise Err_Item;
      End If;
    End If;
    If �ѱ�_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, ����id_In) Into v_�ѱ� From Dual;
    Else
      v_�ѱ� := �ѱ�_In;
    End If;
    If v_�ѱ� Is Null Then
      n_���ηѱ� := 1;
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
    End If;
    Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
  
    If ��������_In = 1 Then
      Select Zl_Age_Calc(����id_In) Into v_���� From Dual;
      If v_���� Is Not Null Then
        Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;
      End If;
    End If;
    --��ȡ��ǰ��������
    If ������_In Is Not Null Then
      v_������ := ������_In;
    Else
      Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_ʵ�ս��ϼ� := 0;
    Select Count(*) + 1
    Into n_�Һ����
    From ���˹Һż�¼
    Where �ű� = ����_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
    --Begin
    v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
    n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
    n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
    n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
    n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
    n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
    n_ͬԴ�޺���     := To_Number(Nvl(zl_GetSysParameter('����ͬһ��Դ�޹�N����', 1111), '0'));
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If �Ǽ�ʱ��_In Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    Else
      d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
    End If;
    If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
      v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
      Raise Err_Item;
    End If;
    n_��������id := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
    n_��id       := Zl_Get��id(v_����Ա����);
  
    --֧���������ύ���
    Select Nvl(Max(1), 0)
    Into n_Exists
    From ���˹Һż�¼
    Where ����id = ����id_In And �ű� = ����_In And ���� = ����_In And ����Ա���� = v_����Ա���� And Nvl(n_�����¼id, 0) = Nvl(�����¼id, 0) And
          �Ǽ�ʱ�� > Sysdate - 0.01 And ��¼״̬ = 1 And ����ʱ�� = ����ʱ��_In;
    If n_Exists = 1 Then
      v_Err_Msg := '�����Ѿ��Һ�,�����ظ�����ͬ�ĺţ�';
      Raise Err_Special;
    End If;
  
    If ������ʽ_In <> 1 Then
      --ԤԼ����Ƿ���Ӻ�����λ����
      --��������˺�����λ���� ��
      Begin
        Select ID
        Into n_�ƻ�id
        From �ҺŰ��żƻ�
        Where ���� = ����_In And ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
              ʧЧʱ�� And Rownum < 2
        Order By ��Чʱ�� Desc;
      Exception
        When Others Then
          Select ID Into n_Tmp����id From �ҺŰ��� Where ���� = ����_In;
      End;
      If Nvl(n_�ƻ�id, 0) <> 0 Then
        Select Count(0)
        Into n_������λ����
        From ������λ�ƻ�����
        Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And Rownum < 2;
      Else
        Select Count(0)
        Into n_������λ����
        From ������λ���ſ���
        Where ������λ = ������λ_In And ����id = n_Tmp����id And Rownum < 2;
      End If;
    End If;
  
    If ������ʽ_In <> 2 Then
      v_���� := Zl_����(����_In);
    End If;
    If ������ʽ_In <> 2 And ���㷽ʽ_In Is Not Null Then
      --�����㷽ʽ�Ƿ��걸
      Select Count(*) Into n_Count From ���㷽ʽ Where ���� = Nvl(���㷽ʽ_In, 'Lxh') And ���� In (2, 7, 8);
      If Nvl(�����id_In, 0) <> 0 And n_Count = 0 Then
        Select Count(1)
        Into n_Count
        From ҽ�ƿ����
        Where ID = Nvl(�����id_In, 0) And ���㷽ʽ = Nvl(���㷽ʽ_In, 'lxh');
      End If;
      If n_Count = 0 Then
        v_Err_Msg := '���㷽ʽ(' || ���㷽ʽ_In || ')δ����,���ڽ��㷽ʽ���������á�';
        Raise Err_Item;
      End If;
    End If;
  
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*) Into n_Count From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ In (1, 3) And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    Open c_Pati(����id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '����δ�ҵ������ܼ�����';
      Raise Err_Item;
    End If;
  
    Open c_����(����_In, ����ʱ��_In);
    Begin
      Fetch c_����
        Into r_����;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
  
    Select Min(վ��) Into v_վ�� From ���ű� Where ID = r_����.����id;
    v_Pricegrade := Zl_Get_Pricegrade(v_վ��, ����id_In, Null, v_���ʽ);
    v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                   '����')
    Into v_����
    From Dual;
    Begin
      If r_����.�ƻ�id Is Null Then
        Select Max(1) Into n_���÷�ʱ�� From �ҺŰ���ʱ�� Where ����id = r_����.Id And ���� = v_���� And Rownum < 2;
        Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
        Into v_ʱ���
        From �ҺŰ���
        Where ID = r_����.Id;
      Else
        Select Max(1)
        Into n_���÷�ʱ��
        From �Һżƻ�ʱ��
        Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And Rownum < 2;
        Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
        Into v_ʱ���
        From �ҺŰ��żƻ�
        Where ID = r_����.�ƻ�id;
      End If;
    Exception
      When Others Then
        n_���÷�ʱ�� := 0;
    End;
  
    If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
      --����Ƿ��ģʽ�ҺŰ���
      Select To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_��鿪ʼʱ��, d_������ʱ��
      From ʱ���
      Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
      If d_��鿪ʼʱ�� > d_������ʱ�� Then
        d_������ʱ�� := d_������ʱ�� + 1;
      End If;
      If d_������ʱ�� > d_����ʱ�� Then
        --��ȡ�����¼id
        Begin
          Select a.Id
          Into n_�����¼id
          From �ٴ������¼ A, �ٴ������Դ B
          Where a.��Դid = b.Id And b.���� = ����_In And �ϰ�ʱ�� = v_ʱ��� And ����ʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
        Exception
          When Others Then
            n_�����¼id := Null;
        End;
      End If;
    End If;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> r_����.����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = r_����.����id;
        If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = r_����.����;
        If n_Count >= Nvl(n_ר�Һ�ԤԼ����, 0) And Nvl(n_ר�Һ�ԤԼ����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ���������ԤԼ����,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> r_����.����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = r_����.����id;
        If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = r_����.����;
        If n_Count >= Nvl(n_ר�ҺŹҺ�����, 0) And Nvl(n_ר�ҺŹҺ�����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��������ŹҺ�����,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_ͬԴ�޺���, 0) <> 0 Then
      If �����¼id_In Is Null Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = ����_In;
      Else
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And �����¼id = �����¼id_In;
      End If;
      If n_Count >= Nvl(n_ͬԴ�޺���, 0) And Nvl(n_ͬԴ�޺���, 0) > 0 Then
        v_Err_Msg := 'ͬһ���������ͬʱ��(ԤԼ)[' || Nvl(n_ͬԴ�޺���, 0) || ']����ͬ�ű�ĺ�,�����ٹҺ�(ԤԼ)��';
        Raise Err_Item;
      End If;
    End If;
  
    d_Date         := Null;
    d_ʱ�ο�ʼʱ�� := Null;
  
    If Nvl(r_����.�޺���, 0) >= 0 Or r_����.�޺��� Is Null Then
    
      Select Nvl(Sum(Nvl(b.�ѹ���, 0)), 0), Nvl(Sum(Nvl(b.�����ѽ���, 0)), 0), Nvl(Sum(Nvl(b.��Լ��, 0)), 0)
      Into n_�ѹ���, n_�����ѽ���, n_��Լ��
      From �ҺŰ��� A, ���˹ҺŻ��� B
      Where a.����id = b.����id And a.��Ŀid = b.��Ŀid And a.���� = ����_In And b.���� Between Trunc(����ʱ��_In) And
            Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And (a.���� = b.���� Or b.���� Is Null) And Nvl(a.ҽ��id, 0) = Nvl(b.ҽ��id, 0) And
            Nvl(a.ҽ������, 'ҽ��') = Nvl(b.ҽ������, 'ҽ��');
    
      If n_���÷�ʱ�� = 1 Then
        If Nvl(r_����.��ſ���, 0) = 1 Then
          If Nvl(�Ƿ������豸_In, 0) = 0 Then
            If r_����.�ƻ�id Is Null Then
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �ҺŰ���ʱ��
              Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0);
            Else
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �Һżƻ�ʱ��
              Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0);
            End If;
            v_Temp := '�Һ�';
            If ������ʽ_In > 1 Then
              v_Temp := 'ԤԼ�Һ�';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ����в��������Ϊ' || Nvl(����_In, 0) || '�İ���,������' || v_Temp || '��';
              Raise Err_Item;
            End If;
          End If;
          --�����,����ѡ��Һ�
          If Trunc(Sysdate) = Trunc(����ʱ��_In) Then
            --�ҵ���ĺ�
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            If r_����.�ƻ�id Is Null Then
              For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                           From �ҺŰ���ʱ��
                           Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
                If Sysdate > v_ʱ��.����ʱ�� Then
                  v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                  Raise Err_Item;
                End If;
              End Loop;
            Else
              For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                           From �Һżƻ�ʱ��
                           Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
                If Sysdate > v_ʱ��.����ʱ�� Then
                  v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                  Raise Err_Item;
                End If;
              End Loop;
            End If;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          If r_����.�ƻ�id Is Null Then
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �ҺŰ���ʱ��
                         Where ����id = r_����.Id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_ԤԼʱ����� := v_ʱ��.���;
              d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            
              Select Count(*), Max(���)
              Into n_Count, n_ԤԼ����
              From �Һ����״̬
              Where ���� = ����_In And ���� = ����ʱ��_In And ״̬ Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_ʱ��.��������, 0) And ��������_In <> 2 Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                             To_Char(v_ʱ��.����ʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.��������, 0) || '��,�����ٽ���ԤԼ�Һţ�';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          Else
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �Һżƻ�ʱ��
                         Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_ԤԼʱ����� := v_ʱ��.���;
              d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            
              Select Count(*), Max(���)
              Into n_Count, n_ԤԼ����
              From �Һ����״̬
              Where ���� = ����_In And ���� = ����ʱ��_In And ״̬ Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_ʱ��.��������, 0) And ��������_In <> 2 Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                             To_Char(v_ʱ��.����ʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.��������, 0) || '��,�����ٽ���ԤԼ�Һţ�';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          End If;
        
          If n_Count = 0 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�����û����صİ��żƻ�(' || To_Char(����ʱ��_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),���ܽ���ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If ������ʽ_In = 1 And ��������_In <> 2 Then
      --�ҺŹ���:
      --  �ѹ������ܴ����޺���
      If n_�ѹ��� >= Nvl(r_����.�޺���, 0) And r_����.�޺��� Is Not Null Then
        v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(r_����.�޺���, 0) || '�����ٹҺţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If ������ʽ_In > 1 Then
      --ԤԼ����ؼ��
      --����:
      --   1.����Լ���ܳ�����Լ��
      --   2.����Ƿ�����ʱ�ε�
      If n_��Լ�� >= Nvl(r_����.��Լ��, 0) And Nvl(r_����.��Լ��, 0) <> 0 And r_����.��Լ�� Is Not Null And ��������_In <> 2 Then
        v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(r_����.��Լ��, 0) || '������ԤԼ�Һţ�';
        Raise Err_Item;
      End If;
    End If;
    If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
    
      If Nvl(r_����.��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
        v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
        Raise Err_Item;
      End If; --Nvl(r_����.��ſ���, 0) =0
    
      n_��� := Case
                When Nvl(r_����.��ſ���, 0) = 1 Or n_���÷�ʱ�� = 1 And ������ʽ_In > 1 Then
                 Nvl(����_In, 0)
                Else
                 0
              End;
    
      --������λ������ģʽ
      Begin
        If Nvl(n_�ƻ�id, 0) <> 0 Then
          Select 0
          Into n_���
          From ������λ�ƻ�����
          Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
        Else
          Select 0
          Into n_���
          From ������λ���ſ���
          Where ������λ = ������λ_In And ����id = n_Tmp����id And
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
        End If;
        n_������λ������ģʽ := 1;
      Exception
        When Others Then
          n_������λ������ģʽ := 0;
      End;
      --������ż��
      For c_������λ In (Select c.���, ����
                     From �ҺŰ��� A, ������λ���ſ��� C
                     Where a.���� = ����_In And Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                   '����', '6', '����', '7', '����', Null) = c.������Ŀ(+) And a.Id = c.����id And
                           c.������λ = ������λ_In And c.��� = n_��� And Not Exists
                      (Select 1
                            From �ҺŰ��żƻ� D
                            Where d.����id = a.Id And d.���ʱ�� Is Not Null And
                                  ����ʱ��_In Between Nvl(d.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  d.ʧЧʱ��)
                     Union All
                     Select c.���, ����
                     From �ҺŰ��żƻ� A, �ҺŰ��� D, ������λ�ƻ����� C,
                          (Select Max(a.��Чʱ��) As ��Ч, ����id
                            From �ҺŰ��żƻ� A, �ҺŰ��� B
                            Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                                  ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  a.ʧЧʱ�� And b.���� = ����_In
                            Group By ����id) E
                     Where a.����id = d.Id And a.���ʱ�� Is Not Null And d.���� = ����_In And a.����id = e.����id And
                           Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                           Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) = c.������Ŀ(+) And a.Id = c.�ƻ�id And c.������λ = ������λ_In And c.��� = n_��� And
                           ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           a.ʧЧʱ��) Loop
      
        If Nvl(r_����.��ſ���, 0) = 1 And c_������λ.��� = n_��� And n_������λ������ģʽ = 0 Then
          n_�Ƿ񿪷� := 1;
          Exit;
        Elsif (Nvl(r_����.��ſ���, 0) = 0 And c_������λ.��� = n_���) Or n_������λ������ģʽ = 1 Then
          Begin
            Select Nvl(��Լ��, 0)
            Into n_ԤԼ����
            From ������λ�ҺŻ���
            Where ������λ = ������λ_In And ���� = Trunc(����ʱ��_In) And ���� = ����_In;
          Exception
            When Others Then
              n_ԤԼ���� := 0;
          End;
          If c_������λ.���� <= n_ԤԼ���� And Nvl(c_������λ.����, 0) > 0 And ��������_In <> 2 Then
            v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(c_������λ.����, 0) || '������ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
          n_�Ƿ񿪷� := 1;
          Exit;
        End If;
      
      End Loop;
    
      If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
        v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
        Raise Err_Item;
      End If;
    End If;
  
    --����޺�������Լ��
    n_�к�         := 1;
    n_ԭ��Ŀid     := 0;
    n_ԭ������Ŀid := 0;
    n_ʵ�ս��ϼ� := 0;
    If ��������_In <> 1 Then
      If ������ʽ_In <> 2 Then
        If Nvl(����id_In, 0) = 0 Then
          --����Ӧ�ó�����
          Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
        Else
          n_����id := ����id_In;
        End If;
      Else
        n_����id := Null;
      End If;
    End If;
    n_��Ŀid := r_����.��Ŀid;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      v_���� := '1|' || n_�ƻ�id;
    Else
      If Nvl(r_����.Id, 0) <> 0 Then
        v_���� := '0|' || r_����.Id;
      End If;
    End If;
    If v_���� Is Null Then
      v_���� := '3|' || ����_In;
    End If;
  
    n_������Ŀid := Zl_Custom_Getregeventitem(r_Pati.����id, r_Pati.����, r_Pati.���֤��, r_Pati.��������, r_Pati.�Ա�, r_Pati.����, v_����);
    If Nvl(n_������Ŀid, 0) <> 0 Then
      n_��Ŀid := n_������Ŀid;
    End If;
  
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := n_��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := n_��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������,
                          Nvl(a.��Ŀ����, 0) As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = r_����.��Ŀid And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Order By ����, ��Ŀ����, �������) Loop
      If c_Item.���� = 1 Then
        n_���� := Nvl(c_Item.����, 0);
      End If;
      n_�۸񸸺� := Null;
      If n_ԭ��Ŀid = c_Item.��Ŀid Then
        If n_ԭ������Ŀid <> c_Item.������Ŀid Then
          n_�۸񸸺� := n_�к�;
        End If;
        n_ԭ������Ŀid := c_Item.������Ŀid;
      End If;
      n_ԭ��Ŀid := c_Item.��Ŀid;
      n_Ӧ�ս�� := Round(c_Item.���� * c_Item.����, 5);
      n_ʵ�ս�� := n_Ӧ�ս��;
      If Nvl(c_Item.���ηѱ�, 0) <> 1 And n_���ηѱ� = 0 Then
        --����:
        v_Temp     := Zl_Actualmoney(r_Pati.�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_ʵ�ս�� := Zl_To_Number(v_Temp);
      End If;
      n_ʵ�ս��ϼ� := Nvl(n_ʵ�ս��ϼ�, 0) + n_ʵ�ս��;
    
      --�������ݲ���������
      If ��������_In <> 1 Then
        --�������˹Һŷ���(���ܵ����ǻ������������)
        Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
        --:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
        Values
          (n_����id, 4, Decode(������ʽ_In, 2, 0, 1), n_�к�, n_�۸񸸺�, c_Item.��������, ���ݺ�_In, Ʊ�ݺ�_In, 1, n_����, Null,
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, r_����.����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, n_ʵ�ս��)),
           Decode(Nvl(���ʷ���_In, 0), 1, Null, n_����id), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), r_����.����id, r_����.ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null,
           Null, ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(����ʱ��_In);
    End If;
  
    --���¹Һ����״̬
    If ��������_In <> 2 Then
      n_���� := ����_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From �Һ����״̬
      Where Trunc(����) = Trunc(����ʱ��_In) And ���� = ����_In And ��� = n_���� And ״̬ <> 5;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 Then
        n_���� := Null;
      End If;
      If n_���÷�ʱ�� = 1 And Nvl(r_����.��ſ���, 0) = 1 Then
        v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
        Raise Err_Item;
      End If;
    End If;
    n_Count := 0;
    If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
      If �˺�����_In = 1 Then
        Select Nvl(Max(���), 0) + 1
        Into n_����
        From �Һ����״̬
        Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ Not In (4, 5);
      Else
        Select Nvl(Max(���), 0) + 1
        Into n_����
        From �Һ����״̬
        Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ <> 5;
      End If;
    End If;
    If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
    
      If ������ʽ_In > 1 And Nvl(r_����.��ſ���, 0) = 0 Then
        --����:ԤԼʱ�����||ԤԼ��
        If Nvl(n_ԤԼ����, 0) = 0 Then
          v_Temp := Nvl(r_����.��Լ��, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_ԤԼʱ����� || v_Temp;
          n_���� := To_Number(v_Temp);
        Else
          n_���� := n_ԤԼ���� + 1;
        End If;
      End If;
    End If;
  
    If Nvl(r_����.��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
      --������ŵĴ���
      Begin
        Select ����Ա����, ������
        Into v_��Ų���Ա, v_��Ż�����
        From �Һ����״̬
        Where ״̬ = 5 And ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_����;
        n_������� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_�������   := 0;
      End;
      If n_������� = 0 Then
        Update �Һ����״̬
        Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
        Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ����Ա���� = v_����Ա����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
            Values
              (����_In, d_Date, n_����, Decode(������ʽ_In, 2, 2, 1), v_����Ա����, Decode(������ʽ_In, 1, 0, 1), Sysdate);
          
            If n_������λ���� > 0 And ������ʽ_In > 1 And Nvl(n_�Ƿ񿪷�, 0) = 1 Then
              Update ������λ�ҺŻ���
              Set ��Լ�� = ��Լ�� + Decode(������ʽ_In, 2, 1, 0), �ѽ��� = �ѽ��� + Decode(������ʽ_In, 3, 1, 0)
              Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ������λ = ������λ_In;
              If Sql%NotFound Then
                Insert Into ������λ�ҺŻ���
                  (����, ����, ���, ������λ, ��Լ��, �ѽ���)
                Values
                  (����_In, d_Date, n_����, ������λ_In, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 3, 1, 0));
              End If;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_���� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
          Raise Err_Item;
        Else
          Update �Һ����״̬
          Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
          Where ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_���� And ״̬ = 5 And ����Ա���� = v_����Ա���� And ������ = v_������;
        End If;
      End If;
    End If;
  
    If n_�����¼id Is Not Null Then
      Update �ٴ�������ſ���
      Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����Ա���� = v_����Ա����
      Where ��¼id = n_�����¼id And ��� = n_���;
      If ������ʽ_In = 2 Then
        Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
      Else
        If ������ʽ_In <> 1 Then
          Update �ٴ������¼
          Set ��Լ�� = ��Լ�� + 1, �ѹ��� = �ѹ��� + 1, �����ѽ��� = �����ѽ��� + 1
          Where ID = n_�����¼id;
        Else
          Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
        End If;
      End If;
    End If;
  
    --�������ݲ������κ� ����
    If ������ʽ_In <> 2 And ��������_In <> 1 And Nvl(���ʷ���_In, 0) = 0 Then
      --�Һ�,ԤԼ�Һ��Ѿ��ۿ��
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      n_����ϼ� := 0;
      If ���ս���_In Is Not Null Then
        --�������ս���
        v_�������� := ���ս���_In || '||';
        n_����ϼ� := 0;
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
          n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
          n_����ϼ� := Nvl(n_����ϼ�, 0) + Nvl(n_������, 0);
          v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(��Ԥ��_In, 0) <> 0 Then
        --������Ԥ��
        n_����ϼ� := n_����ϼ� + Nvl(��Ԥ��_In, 0);
        n_Ԥ����� := ��Ԥ��_In;
        For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
          n_������ := Case
                      When r_Deposit.��� - n_Ԥ����� < 0 Then
                       r_Deposit.���
                      Else
                       n_Ԥ�����
                    End;
          If r_Deposit.����id = 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (r_Deposit.����id, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          --����Ƿ��Ѿ�������
          If r_Deposit.��� <= n_������ Then
            n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
          Else
            n_Ԥ����� := 0;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
        If n_Ԥ����� > 0 Then
          v_Err_Msg := 'Ԥ������֧������֧�����,���ܼ���������';
          Raise Err_Item;
        End If;
      End If;
      --ʣ�����,��ָ�����㷽֧��
      n_������ := Nvl(n_ʵ�ս��ϼ�, 0) - Nvl(n_����ϼ�, 0);
      If Nvl(n_������, 0) < 0 Then
        v_Err_Msg := '�Һŵ���ؽ�������˵�ǰʵ����,���ܼ���������';
        Raise Err_Item;
      End If;
      If Nvl(n_������, 0) <> 0 Or (Nvl(n_������, 0) = 0 And Nvl(��Ԥ��_In, 0) = 0) Then
        If ���㷽ʽ_In Is Null Then
          v_Err_Msg := 'δ����ָ���Ľ��㷽ʽ,���ܼ���������';
          Raise Err_Item;
        End If;
      
        If Nvl(Ԥ��id_In, 0) <> 0 Then
          --�����Ԥ��ID_In��Ҫ��Ϊ�˽����������,���ҽ������վ���˸�ID,��Ҫ���µ�ID���и���,����������ת���ID
          Update ����Ԥ����¼ Set ID = n_Ԥ��id Where ID = Nvl(Ԥ��id_In, 0);
          n_Ԥ��id := Nvl(Ԥ��id_In, 0);
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ������ˮ��, ����˵��, �������, ������λ, �����id, ����,
           ��������)
        Values
          (n_Ԥ��id, 4, 1, ���ݺ�_In, r_Pati.����id, ���㷽ʽ_In, Nvl(n_������, 0), d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_����id,
           ������λ_In || '�ɿ�', n_��id, ������ˮ��_In, ����˵��_In, n_����id, ������λ_In, �����id_In, ֧������_In, 4);
      End If;
    
      --������Ա�ɿ�����
    
      For v_�ɿ� In (Select ���㷽ʽ, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ A
                   Where a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And Nvl(����id, 0) = Nvl(����id_In, 0)
                   Group By ���㷽ʽ) Loop
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(v_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_�ɿ�.���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, v_�ɿ�.���㷽ʽ, 1, Nvl(v_�ɿ�.��Ԥ��, 0));
          n_����ֵ := Nvl(v_�ɿ�.��Ԥ��, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = ���㷽ʽ_In And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      
      End Loop;
    
    End If;
  
    --����Һż�¼
    If ��������_In = 2 Then
      Begin
        Select ID Into n_�Һ�id From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    End If;
  
    Update ���˹Һż�¼
    Set ��¼���� = Decode(������ʽ_In, 2, 2, 1), ��¼״̬ = Decode(��������_In, 1, 0, 1), ����� = r_Pati.�����, ����Ա���� = v_����Ա����,
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
        ����ʱ�� = Case ��������_In
                  When 1 Then
                   Null
                  Else
                   Case ������ʽ_In
                     When 2 Then
                      Null
                     Else
                      d_�Ǽ�ʱ��
                   End
                End, ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), ������λ = Nvl(������λ_In, ������λ),
        ԤԼ����Ա = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա, Null), Nvl(ԤԼ����Ա, v_����Ա����)),
        ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���))
    Where ID = n_�Һ�id;
    If Sql%NotFound Then
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = r_Pati.���ʽ And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
         ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���)
      Values
        (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
         r_Pati.�Ա�, r_Pati.����, ����_In, n_����, v_����, Null, r_����.����id, r_����.ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
         Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���));
    End If;
    --�������ݲ��ܲ�������
    If ��������_In <> 1 Then
      n_ԤԼ���ɶ��� := 0;
      If ������ʽ_In > 1 Then
        n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
      End If;
      --�Һź��շѵ�ԤԼ��ֱ�ӽ������(�շ�ԤԼȱ�ٽ��չ���,����ֱ�Ӻ͹Һ�һ��ֱ�ӽ������)
      If ������ʽ_In <> 2 Or n_ԤԼ���ɶ��� = 1 Then
        If Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113)) <> 0 Then
          --�Ŷӽк�ģʽ:-0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
          If Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(r_����.����id, 0))) = 0 Or n_ԤԼ���ɶ��� = 1 Then
            n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(������ʽ_In, 0) > 1 And n_��ʱ����ʾ = 1 And n_���÷�ʱ�� = 1 Then
              n_��ʱ����ʾ := 1;
            Else
              n_��ʱ����ʾ := Null;
            End If;
            --��������
            --.����ִ�в��š� �ķ�ʽ���ɶ���
            v_�������� := r_����.����id;
            v_�ŶӺ��� := Zlgetnextqueue(r_����.����id, n_�Һ�id, ����_In || '|' || ����_In);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
            --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
            Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, r_����.����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, r_����.ҽ������,
                             d_�Ŷ�ʱ��, ԤԼ��ʽ_In, n_��ʱ����ʾ, v_�Ŷ����);
          End If;
        End If;
      End If;
    
      If Nvl(������ʽ_In, 0) = 1 Then
        --����Ʊ��ʹ�����
        If Ʊ�ݺ�_In Is Not Null Then
          Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
          --����Ʊ��
          Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Values
            (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�Ǽ�ʱ��, v_����Ա����, �ҺŽ��ϼ�_In);
          --״̬�Ķ�
          Update Ʊ�����ü�¼
          Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
          Where ID = Nvl(����id_In, 0);
        End If;
        --���˱��ξ���(�Է���ʱ��Ϊ׼)
        If Nvl(r_Pati.����id, 0) <> 0 Then
          Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = v_���� Where ����id = r_Pati.����id;
        End If;
      End If;
    End If;
    --���˹ҺŻ���
    --��������ʱ�����ٶԻ��ܵ��ݽ���ͳ���� ����������ʱ�Ѿ������˻���
    If ��������_In <> 2 Then
      --������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
      --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
      --����zl_third_lockno�������ţ�������ʹ�ñ���������
      n_ԤԼ := Case
                When Nvl(������ʽ_In, 0) = 1 Then
                 0
                When Nvl(������ʽ_In, 0) = 2 Then
                 1
                When Nvl(������ʽ_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_���˹ҺŻ���_Update(r_����.ҽ������, r_����.ҽ��id, r_����.��Ŀid, r_����.����id, ����ʱ��_In, n_ԤԼ, ����_In);
    End If;
  
    If ��������_In <> 1 Then
      --��Ϣ����,����ʱ��������Ϣ
      Begin
        Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
          Using 1, n_�Һ�id;
      Exception
        When Others Then
          Null;
      End;
      b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Insert;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_���������Һ�_Delete
(
  ���ݺ�_In     ������ü�¼.No%Type,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
  ����˵��_In   ����Ԥ����¼.����˵��%Type,
  �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
) As
  v_Error Varchar2(255);
  Err_Custom Exception;

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo
  (
    v_״̬     ���˹Һż�¼.��¼״̬%Type,
    v_����     ���˹Һż�¼.��¼����%Type,
    v_��Ч���� Number := 0
  ) Is
  
    Select a.����ʱ��, a.�Ǽ�ʱ��, b.�շ�ϸĿid As ��Ŀid, a.ִ�в���id As ����id, a.ִ���� As ҽ������, c.Id As ҽ��id, a.�ű� As ����
    From ���˹Һż�¼ A, ������ü�¼ B, ��Ա�� C
    Where a.��¼���� = Decode(v_��Ч����, 0, v_����, a.��¼����) And b.��¼���� = 4 And a.��¼״̬ = v_״̬ And a.No = ���ݺ�_In And a.No = b.No And
          a.ִ���� = c.����(+) And b.��� = 1 And Rownum = 1;

  r_Registrow c_Registinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney Is
    Select Distinct b.���㷽ʽ, b.��Ԥ��
    From ������ü�¼ A, ����Ԥ����¼ B
    Where a.����id = b.����id And a.No = ���ݺ�_In And a.��¼���� = 4 And a.��¼״̬ = 3 And b.��¼���� = 4 And b.��¼״̬ = 3 And
          Nvl(b.��Ԥ��, 0) <> 0;

  n_ִ��״̬       ���˹Һż�¼.ִ��״̬%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������ü�¼.����id%Type;
  n_ԭ����id       ����Ԥ����¼.����id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_����ֵ         �������.Ԥ�����%Type;
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ�Һ�       Number;
  n_��Ч����       Number; --��Ч����û�в������õ���
  n_�Һ����ɶ���   Number;
  n_Count          Number;
  n_��id           ����ɿ����.Id%Type;
  d_�˺�ʱ��       Date;
  v_����Ա���     ��Ա��.���%Type;
  v_����Ա����     ��Ա��.����%Type;
  v_������λ       ������λ�ҺŻ���.������λ%Type;
  n_ԤԼ״̬       ���˹Һż�¼.ԤԼ%Type;
  v_Temp           Varchar2(100);
  d_�Ǽ�ʱ��       ���˹Һż�¼.�Ǽ�ʱ��%Type;
  v_�ű�           ���˹Һż�¼.�ű�%Type;
  n_����           ���˹Һż�¼.����%Type;
  d_ԤԼʱ��       ���˹Һż�¼.ԤԼʱ��%Type;
  n_������λ����   Number(18);
  n_ԤԼ���ɶ���   Number;
  n_��¼����       Number;
  n_״̬           Number;
  n_�˺�����       Number(3);
  n_�Һ��Ű�ģʽ   Number;
  n_����           ������ü�¼.���ʷ���%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_�ѽ���         Number;
  n_���ض�         �������.�������%Type;
  n_Ԥ��֧��       Number(3);
  n_����֧��       Number(3);
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   Date;
  d_������ʱ��   Date;
  d_����ʱ��       Date;
  n_�����¼id     Number(18);
  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_���������Һ�_����_Delete
  (
    ���ݺ�_In     ������ü�¼.No%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
    Ԥ��id_In     ����Ԥ����¼.Id%Type := Null
  ) As
    v_Error Varchar2(255);
    Err_Custom Exception;
  
    --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
    Cursor c_Registinfo
    (
      v_״̬     ���˹Һż�¼.��¼״̬%Type,
      v_����     ���˹Һż�¼.��¼����%Type,
      v_��Ч���� Number := 0
    ) Is
      Select a.����ʱ��, a.�Ǽ�ʱ��, b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id, b.Id As ��¼id, a.�ű� As ����
      From ���˹Һż�¼ A, �ٴ������¼ B
      Where a.��¼���� = Decode(v_��Ч����, 0, v_����, a.��¼����) And a.��¼״̬ = v_״̬ And a.No = ���ݺ�_In And a.�����¼id = b.Id And
            Rownum < 2;
  
    r_Registrow c_Registinfo%RowType;
  
    --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
    Cursor c_Opermoney Is
      Select Distinct b.���㷽ʽ, b.��Ԥ��
      From ������ü�¼ A, ����Ԥ����¼ B
      Where a.����id = b.����id And a.No = ���ݺ�_In And a.��¼���� = 4 And a.��¼״̬ = 3 And b.��¼���� = 4 And b.��¼״̬ = 3 And
            Nvl(b.��Ԥ��, 0) <> 0;
  
    n_ִ��״̬       ���˹Һż�¼.ִ��״̬%Type;
    n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
    n_����id         ������ü�¼.����id%Type;
    n_ԭ����id       ����Ԥ����¼.����id%Type;
    n_����id         ������Ϣ.����id%Type;
    n_����ֵ         �������.Ԥ�����%Type;
    n_����̨ǩ���Ŷ� Number;
    n_ԤԼ�Һ�       Number;
    n_��Ч����       Number; --��Ч����û�в������õ���
    n_�Һ����ɶ���   Number;
    n_Count          Number;
    n_��id           ����ɿ����.Id%Type;
    d_�˺�ʱ��       Date;
    v_����Ա���     ��Ա��.���%Type;
    v_����Ա����     ��Ա��.����%Type;
    v_������λ       ������λ�ҺŻ���.������λ%Type;
    n_ԤԼ״̬       ���˹Һż�¼.ԤԼ%Type;
    v_Temp           Varchar2(100);
    d_�Ǽ�ʱ��       ���˹Һż�¼.�Ǽ�ʱ��%Type;
    v_�ű�           ���˹Һż�¼.�ű�%Type;
    n_����           ���˹Һż�¼.����%Type;
    d_ԤԼʱ��       ���˹Һż�¼.ԤԼʱ��%Type;
    n_������λ����   Number(18);
    n_ԤԼ���ɶ���   Number;
    n_��¼����       Number;
    n_״̬           Number;
    n_�˺�����       Number(3);
    n_�Һ�id         ���˹Һż�¼.Id%Type;
    n_����           ������ü�¼.���ʷ���%Type;
    n_��¼id         �ٴ������¼.Id%Type;
  
    n_�ѽ���   Number;
    n_���ض�   �������.�������%Type;
    n_Ԥ��֧�� Number(3);
    n_����֧�� Number(3);
    Function Zl_����Ա
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
      -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
    Begin
      If Type_In = 0 Then
        --ȱʡ����
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  Begin
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Custom;
    End If;
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
  
    n_��id := Zl_Get��id(v_����Ա����);
  
    d_�˺�ʱ�� := �˺�ʱ��_In;
    If d_�˺�ʱ�� Is Null Then
      d_�˺�ʱ�� := Sysdate;
    End If;
  
    --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
    Begin
      Select Decode(��¼����, 2, 1, 0), ��¼����, �Ǽ�ʱ��, �ű�, ����, Nvl(ԤԼʱ��, ����ʱ��), ������λ, Nvl(ԤԼ, 0), Decode(��¼״̬, 0, 1, 0), �����¼id
      Into n_ԤԼ�Һ�, n_��¼����, d_�Ǽ�ʱ��, v_�ű�, n_����, d_ԤԼʱ��, v_������λ, n_ԤԼ״̬, n_��Ч����, n_��¼id
      From ���˹Һż�¼
      Where NO = ���ݺ�_In And ��¼״̬ In (0, 1) And Rownum < 2;
    Exception
      When Others Then
        n_ԤԼ�Һ� := -1;
    End;
  
    If n_ԤԼ�Һ� = -1 Then
      v_Error := '���ݿ����Ѿ����˺Ż򵥾��������!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(���ʷ���, 0), Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
      Into n_����, n_�ѽ���
      From ������ü�¼
      Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_����   := 0;
        n_�ѽ��� := 0;
    End;
  
    --ԤԼ����Ƿ���Ӻ�����λ����
    --��������˺�����λ���� ��
    Select Count(0) Into n_������λ���� From �ٴ�����Һſ��Ƽ�¼ Where ���� = 1 And ���� = 1 And Rownum < 2;
    --���¹Һ����״̬
    n_�˺����� := Zl_To_Number(zl_GetSysParameter('�����������Һ�', 1111));
    If n_�˺����� = 0 Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 4 Where ��¼id = n_��¼id And (��� = n_���� Or ��ע = To_Char(n_����));
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ���� = Null, ���� = Null, ����Ա���� = Null, ����վ���� = Null
      Where ��¼id = n_��¼id And (��� = n_���� Or ��ע = To_Char(n_����));
    End If;
    If Nvl(n_ԤԼ�Һ�, 0) = 1 Or Nvl(n_��Ч����, 0) = 1 Then
      If Nvl(n_��Ч����, 0) = 0 Then
        --N���ڲ���ȡ��ԤԼ��
        n_Count := Zl_To_Number(zl_GetSysParameter('N���ڲ���ȡ��ԤԼ��', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_�Ǽ�ʱ�� Then
            v_Error := '�����˵�ԤԼ��' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '��ǰ��ԤԼ��!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_״̬ := Case n_��Ч����
                When 1 Then
                 0
                Else
                 1
              End;
      --������Լ��
      Open c_Registinfo(n_״̬, 2, n_��Ч����);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -n_ԤԼ״̬, Decode(n_ԤԼ״̬, 0, 1, 0));
      End If;
    
      Update �ٴ������¼
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ID = n_��¼id;
      Close c_Registinfo;
    
      If Nvl(n_��Ч����, 0) = 0 Then
        --ɾ��������ü�¼
        Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
        --���ԤԼ���ɶ���ʱ��Ҫ�������
        n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If Nvl(n_�Һ����ɶ���, 0) = 1 Then
          n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
          If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
            --Ҫɾ������
            For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
              Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    
      --���¹Һ����״̬
    
      --���˾���״̬
      Select ����id
      Into n_����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ��� = 1;
    
      If n_����id Is Not Null Then
        Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
        --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      End If;
    
      --������ü�¼
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
         ����, ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��,
         �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
        Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
               v_����Ա���, v_����Ա����, ����ʱ��, d_�˺�ʱ��, n_����id,
               Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����, ժҪ,
               Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
    
      --ԭʼ��¼
      If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
        Update ������ü�¼
        Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      Else
        Update ������ü�¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      End If;
    
      n_ԭ����id := 0;
      If n_���� = 0 Then
        --��ȡ����ID
        Select Nvl(����id, 0)
        Into n_ԭ����id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
      End If;
    
      If n_���� = 1 Then
        --����
        For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                     From ������ü�¼
                     Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Nvl(���ʷ���, 0) = 1) Loop
          --�������
          Update �������
          Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
          Returning ������� Into n_���ض�;
        
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, ����, �������, Ԥ�����)
            Values
              (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
            n_���ض� := Nvl(c_����.ʵ�ս��, 0);
          End If;
          If Nvl(n_���ض�, 0) = 0 Then
            Delete �������
            Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
          --����δ�����
          Update ����δ�����
          Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
                Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And
                ������Ŀid + 0 = c_����.������Ŀid And ��Դ;�� + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into ����δ�����
              (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
            Values
              (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
          End If;
        End Loop;
        Delete ����δ�����
        Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
      End If;
    
      If n_���� = 0 Then
        Begin
          Select 1
          Into n_Ԥ��֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_Ԥ��֧�� := 0;
        End;
        Begin
          Select 1
          Into n_����֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_����֧�� := 0;
        End;
        If n_Ԥ��֧�� = 1 And n_����֧�� = 1 Then
          v_Error := '���ܴ�����ֽ��㷽ʽ,���鴫����˺ŵ����Ƿ���ȷ!';
          Raise Err_Custom;
        End If;
        If n_Ԥ��֧�� = 1 Then
          --ԭ���˻�Ԥ��
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                     d_�˺�ʱ��, v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_�˺�ʱ��,
                     v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          End If;
          --������Ԥ�����
          For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                       From ����Ԥ����¼
                       Where ��¼���� In (1, 11) And ����id = n_����id
                       Group By ����id, Ԥ�����) Loop
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
            Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, Ԥ�����, ����, ����)
              Values
                (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
              n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��,
                     n_����id, n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��, n_����id,
                     n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          End If;
          Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
        End If;
        --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
        --�����һ�εĴ�ӡ������ȡ
        Select Max(ID)
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
        If n_��ӡid Is Not Null Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_�˺�ʱ��, v_����Ա����, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1;
        End If;
      End If;
    
      --��ػ��ܱ�Ĵ���
    
      --���˹ҺŻ���
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --ֻ�ղ�����ʱ�޺ű�,������
        Close c_Registinfo;
      Else
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1), ִ��״̬
          Into n_ԤԼ�Һ�, n_ִ��״̬
          From ���˹Һż�¼
          Where NO = ���ݺ�_In And ��¼״̬ = 1 And Rownum = 1;
        Exception
          When Others Then
            n_ԤԼ�Һ� := 0;
        End;
        --0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
        If n_ִ��״̬ > 0 Then
          If n_ִ��״̬ = 1 Then
            v_Error := '�ò����Ѿ���ɾ���,�������˺�!';
          Else
            v_Error := '�ò������ھ���, �����˺�!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
              Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
              (���� = r_Registrow.���� Or ���� Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
             Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�);
        End If;
      
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ID = n_��¼id;
        Close c_Registinfo;
      End If;
    
      --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
      If n_���� = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
          Returning ��� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
            n_����ֵ := r_Opermoney.��Ԥ��;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
      If n_�Һ����ɶ��� <> 0 Then
        --Ҫɾ������
        For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
          n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, Nvl(v_�Һ�.ִ�в���id, 0)));
          If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End If;
        End Loop;
      End If;
    
      --ҽ�������ľ���ǼǼ�¼
      Delete From ����ǼǼ�¼
      Where (����id, ��ҳid, ����ʱ��) In (Select ����id, ��ҳid, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In);
    End If;
  
    If Nvl(n_��Ч����, 0) = 0 Then
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1;
      If Sql%NotFound Then
        v_Error := 'δ�ҵ��Һŵ���,����!';
        Raise Err_Custom;
      End If;
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, �����¼id)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_�˺�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��_In, ����˵��_In, ������λ, �����¼id
        From ���˹Һż�¼
        Where NO = ���ݺ�_In And ��¼״̬ = 3;
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 2, ���ݺ�_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_�Һ��Ű�ģʽ := To_Number(Substr(Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0), 1, 1));
  If n_�Һ��Ű�ģʽ = 1 Then
    --������Ű�ģʽ
    Zl_���������Һ�_����_Delete(���ݺ�_In, ������ˮ��_In, ����˵��_In, �˺�ʱ��_In, Ԥ��id_In);
  Else
    v_Temp := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          d_����ʱ�� := Null;
      End;
    End If;
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Error := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Custom;
    End If;
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
  
    n_��id := Zl_Get��id(v_����Ա����);
  
    d_�˺�ʱ�� := �˺�ʱ��_In;
    If d_�˺�ʱ�� Is Null Then
      d_�˺�ʱ�� := Sysdate;
    End If;
  
    --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
    Begin
      Select Decode(��¼����, 2, 1, 0), ��¼����, �Ǽ�ʱ��, �ű�, ����, Nvl(ԤԼʱ��, ����ʱ��), ������λ, Nvl(ԤԼ, 0), Decode(��¼״̬, 0, 1, 0)
      Into n_ԤԼ�Һ�, n_��¼����, d_�Ǽ�ʱ��, v_�ű�, n_����, d_ԤԼʱ��, v_������λ, n_ԤԼ״̬, n_��Ч����
      From ���˹Һż�¼
      Where NO = ���ݺ�_In And ��¼״̬ In (0, 1) And Rownum <= 1;
    Exception
      When Others Then
        n_ԤԼ�Һ� := -1;
    End;
  
    If n_ԤԼ�Һ� = -1 Then
      v_Error := '���ݿ����Ѿ����˺Ż򵥾��������!';
      Raise Err_Custom;
    End If;
  
    Begin
      Select Nvl(���ʷ���, 0), Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
      Into n_����, n_�ѽ���
      From ������ü�¼
      Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;
    Exception
      When Others Then
        n_����   := 0;
        n_�ѽ��� := 0;
    End;
  
    Begin
      Select a.Id Into n_����id From �ҺŰ��� A Where a.���� = v_�ű�;
    Exception
      When Others Then
        n_����id := -1;
    End;
  
    Begin
      Select ID
      Into n_�ƻ�id
      From �ҺŰ��żƻ�
      Where ����id = n_����id And ���ʱ�� Is Not Null And
            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.��Чʱ��) As ��Ч
             From �ҺŰ��żƻ� A
             Where a.���ʱ�� Is Not Null And d_ԤԼʱ�� Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   a.ʧЧʱ�� And a.����id = n_����id) And
            d_ԤԼʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And ʧЧʱ��;
    Exception
      When Others Then
        n_�ƻ�id := 0;
    End;
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Decode(To_Char(d_ԤԼʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Decode(To_Char(d_ԤԼʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  
    If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
      --����Ƿ��ģʽ�ҺŰ���
      Select To_Date(To_Char(d_ԤԼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(d_ԤԼʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_��鿪ʼʱ��, d_������ʱ��
      From ʱ���
      Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
      If d_��鿪ʼʱ�� > d_������ʱ�� Then
        d_������ʱ�� := d_������ʱ�� + 1;
      End If;
      If d_������ʱ�� > d_����ʱ�� Then
        --��ȡ�����¼id
        Select Max(a.Id)
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_�ű� And �ϰ�ʱ�� = v_ʱ��� And d_ԤԼʱ�� Between ��ʼʱ�� And ��ֹʱ��;
      End If;
    End If;
  
    --ԤԼ����Ƿ���Ӻ�����λ����
    --��������˺�����λ���� ��
    Select Count(0) Into n_������λ���� From ������λ���ſ��� Where Rownum = 1;
    --���¹Һ����״̬
    n_�˺����� := Zl_To_Number(zl_GetSysParameter('�����������Һ�', 1111));
    If n_�˺����� = 0 Then
      Update �Һ����״̬
      Set ״̬ = 4
      Where ���� = v_�ű� And ��� = n_���� And ���� Between Trunc(d_ԤԼʱ��) And Trunc(d_ԤԼʱ�� + 1) - 1 / 24 / 60 / 60;
    Else
      Delete �Һ����״̬
      Where ���� = v_�ű� And ��� = n_���� And ���� Between Trunc(d_ԤԼʱ��) And Trunc(d_ԤԼʱ�� + 1) - 1 / 24 / 60 / 60;
    End If;
    If Nvl(n_ԤԼ�Һ�, 0) = 1 Or Nvl(n_��Ч����, 0) = 1 Then
      If Nvl(n_��Ч����, 0) = 0 Then
        --N���ڲ���ȡ��ԤԼ��
        n_Count := Zl_To_Number(zl_GetSysParameter('N���ڲ���ȡ��ԤԼ��', 1111));
        If n_Count <> 0 Then
          If Trunc(Sysdate - n_Count) < d_�Ǽ�ʱ�� Then
            v_Error := '�����˵�ԤԼ��' || To_Char(Trunc(Sysdate - n_Count), 'yyyy-mm-dd') || '��ǰ��ԤԼ��!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      n_״̬ := Case n_��Ч����
                When 1 Then
                 0
                Else
                 1
              End;
      --������Լ��
      Open c_Registinfo(n_״̬, 2, n_��Ч����);
      Fetch c_Registinfo
        Into r_Registrow;
    
      Update ���˹ҺŻ���
      Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��, �ѹ���)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -n_ԤԼ״̬, Decode(n_ԤԼ״̬, 0, 1, 0));
      End If;
    
      If n_�����¼id Is Not Null Then
        Update �ٴ������¼
        Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѹ��� = Nvl(�ѹ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
        Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
        Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_����;
      End If;
      If Nvl(n_������λ����, 0) <> 0 And v_������λ Is Not Null And Nvl(n_ԤԼ״̬, 0) <> 0 Then
        Update ������λ�ҺŻ���
        Set ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬, �ѽ��� = Nvl(�ѽ���, 0) - Decode(n_ԤԼ״̬, 0, 1, 0)
        Where ���� = Trunc(r_Registrow.����ʱ��) And (���� = r_Registrow.���� Or ���� Is Null) And ������λ = v_������λ And
              ��� = Nvl(n_����, 0);
        If Sql%RowCount = 0 Then
          Insert Into ������λ�ҺŻ���
            (����, ����, ��Լ��, ������λ, ���, �ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����, -n_ԤԼ״̬, v_������λ, Nvl(n_����, 0), -decode(n_ԤԼ״̬, 0, 1, 0));
        End If;
      End If;
      Close c_Registinfo;
    
      If Nvl(n_��Ч����, 0) = 0 Then
        --ɾ��������ü�¼
        Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
        --���ԤԼ���ɶ���ʱ��Ҫ�������
        n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        If Nvl(n_�Һ����ɶ���, 0) = 1 Then
          n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
          If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
            --Ҫɾ������
            For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
              Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
            End Loop;
          End If;
        End If;
      End If;
    Else
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    
      --���¹Һ����״̬
    
      --���˾���״̬
      Select ����id
      Into n_����id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ��� = 1;
    
      If n_����id Is Not Null Then
        Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
        --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      End If;
    
      --������ü�¼
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
         ����, ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��,
         �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
        Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
               v_����Ա���, v_����Ա����, ����ʱ��, d_�˺�ʱ��, n_����id,
               Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����, ժҪ,
               Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
    
      --ԭʼ��¼
      If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
        Update ������ü�¼
        Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
        Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      Else
        Update ������ü�¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In;
      End If;
    
      n_ԭ����id := 0;
      If n_���� = 0 Then
        --��ȡ����ID
        Select Nvl(����id, 0)
        Into n_ԭ����id
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
      End If;
    
      If n_���� = 1 Then
        --����
        For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                     From ������ü�¼
                     Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Nvl(���ʷ���, 0) = 1) Loop
          --�������
          Update �������
          Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
          Returning ������� Into n_���ض�;
        
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, ����, �������, Ԥ�����)
            Values
              (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
            n_���ض� := Nvl(c_����.ʵ�ս��, 0);
          End If;
          If Nvl(n_���ض�, 0) = 0 Then
            Delete �������
            Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
          --����δ�����
          Update ����δ�����
          Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
          Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
                Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And
                ������Ŀid + 0 = c_����.������Ŀid And ��Դ;�� + 0 = 1;
        
          If Sql%RowCount = 0 Then
            Insert Into ����δ�����
              (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
            Values
              (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
          End If;
        End Loop;
        Delete ����δ�����
        Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
      End If;
    
      If n_���� = 0 Then
        Begin
          Select 1
          Into n_Ԥ��֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_Ԥ��֧�� := 0;
        End;
      
        Begin
          Select 1
          Into n_����֧��
          From ����Ԥ����¼
          Where Mod(��¼����, 10) = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum < 2;
        Exception
          When Others Then
            n_����֧�� := 0;
        End;
      
        If n_Ԥ��֧�� = 1 And n_����֧�� = 1 Then
          v_Error := '���ܴ�����ֽ��㷽ʽ,���鴫����˺ŵ����Ƿ���ȷ!';
          Raise Err_Custom;
        End If;
      
        If n_Ԥ��֧�� = 1 Then
          --ԭ���˻�Ԥ��
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                     d_�˺�ʱ��, v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_�˺�ʱ��,
                     v_����Ա����, v_����Ա���, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
          End If;
          --������Ԥ�����
          For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                       From ����Ԥ����¼
                       Where ��¼���� In (1, 11) And ����id = n_����id
                       Group By ����id, Ԥ�����) Loop
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
            Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, Ԥ�����, ����, ����)
              Values
                (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
              n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End Loop;
        Else
          If Nvl(Ԥ��id_In, 0) = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��,
                     n_����id, n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, ������ˮ��, ����˵��,
               ������λ, �������, �����id, ��������)
              Select Ԥ��id_In, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_�˺�ʱ��, v_����Ա���, v_����Ա����, -��Ԥ��, n_����id,
                     n_��id, ������ˮ��_In, ����˵��_In, ������λ, n_����id, �����id, ��������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
          End If;
        
          Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_ԭ����id;
        End If;
        --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
        --�����һ�εĴ�ӡ������ȡ
        Select Max(ID)
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
        If n_��ӡid Is Not Null Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_�˺�ʱ��, v_����Ա����, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1;
        End If;
      End If;
    
      --��ػ��ܱ�Ĵ���
    
      --���˹ҺŻ���
      Open c_Registinfo(1, 1);
      Fetch c_Registinfo
        Into r_Registrow;
    
      If c_Registinfo%RowCount = 0 Then
        --ֻ�ղ�����ʱ�޺ű�,������
        Close c_Registinfo;
      Else
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1), ִ��״̬
          Into n_ԤԼ�Һ�, n_ִ��״̬
          From ���˹Һż�¼
          Where NO = ���ݺ�_In And ��¼״̬ = 1 And Rownum = 1;
        Exception
          When Others Then
            n_ԤԼ�Һ� := 0;
        End;
        --0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
        If n_ִ��״̬ > 0 Then
          If n_ִ��״̬ = 1 Then
            v_Error := '�ò����Ѿ���ɾ���,�������˺�!';
          Else
            v_Error := '�ò������ھ���, �����˺�!';
          End If;
          Raise Err_Custom;
        End If;
      
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
        Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
              Nvl(ҽ������, 'ҽ��') = Nvl(r_Registrow.ҽ������, 'ҽ��') And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
              (���� = r_Registrow.���� Or ���� Is Null);
      
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���)
          Values
            (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
             Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�);
        End If;
        If n_�����¼id Is Not Null Then
          Update �ٴ������¼
          Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ״̬, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ״̬
          Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
          Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_����;
        End If;
        If Nvl(n_������λ����, 0) <> 0 And v_������λ Is Not Null And Nvl(n_ԤԼ״̬, 0) <> 0 Then
          Update ������λ�ҺŻ���
          Set �ѽ��� = Nvl(�ѽ���, 0) - 1, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
          Where ���� = Trunc(r_Registrow.����ʱ��) And (���� = r_Registrow.���� Or ���� Is Null) And ������λ = v_������λ And
                ��� = Nvl(n_����, 0);
          If Sql%RowCount = 0 Then
            Insert Into ������λ�ҺŻ���
              (����, ����, ��Լ��, ������λ, �ѽ���, ���)
            Values
              (Trunc(r_Registrow.����ʱ��), r_Registrow.����, -1, v_������λ, -1 * n_ԤԼ�Һ�, Nvl(n_����, 0));
          End If;
        End If;
        Close c_Registinfo;
      End If;
    
      --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
      If n_���� = 0 Then
        For r_Opermoney In c_Opermoney Loop
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
          Returning ��� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
            n_����ֵ := r_Opermoney.��Ԥ��;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = v_����Ա���� And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End Loop;
      End If;
    
      n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
      If n_�Һ����ɶ��� <> 0 Then
        --Ҫɾ������
        For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
          n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
          If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
            Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
          End If;
        End Loop;
      End If;
    
      --ҽ�������ľ���ǼǼ�¼
      Delete From ����ǼǼ�¼
      Where (����id, ��ҳid, ����ʱ��) In (Select ����id, ��ҳid, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In);
    End If;
  
    If Nvl(n_��Ч����, 0) = 0 Then
      Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1;
      If Sql%NotFound Then
        v_Error := 'δ�ҵ��Һŵ���,����!';
        Raise Err_Custom;
      End If;
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ)
        Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_�˺�ʱ��, ����ʱ��,
               v_����Ա���, v_����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��_In, ����˵��_In, ������λ
        From ���˹Һż�¼
        Where NO = ���ݺ�_In And ��¼״̬ = 3;
    End If;
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 2, ���ݺ�_In;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Error || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Delete;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:HIS�Һ� 
  --���:Xml_In: 
  --<IN> 
  --   <CZFS>3</CZFS>    //������ʽ 
  --   <CZJLID>1</CZJLID>    //�����¼ID 
  --   <HM>����</HM>    //���� 
  --   <HX>����</HX>     //���� 
  --   <JKFS>0</JKFS>  //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ� 
  --   <YYSJ>2014-10-21 </YYSJ>    //ԤԼ���� YYYY-MM-DD,��ʱ�η���ſ�����Ҫ����ʱ�� 
  --   <JE>���</JE>     //��� 
  --   <JSLIST> 
  --     <JS>            //������Ϣ���Һŷ�ҽ������Ŀǰ��֧��һ�����ṹ���շ�һ�� 
  --       <JSKLB>���㿨���</JSKLB>    //���㿨��� 
  --       <JSKH>֧�����ʺ�</JSKH>           //���㿨��(֧�����ʺ�) 
  --       <JYSM>����˵��</JYSM>            //˵�����̶���֧���� 
  --       <JYLSH>��ˮ��</JYLSH>           //��ˮ�ţ��������� 
  --       <JSFS>���㷽ʽ</JSFS>            //���㷽ʽ:�ֽ�֧Ʊ�������������,���Դ��� 
  --       <JSJE>������</JSJE>            //������ 
  --       <ZY>ժҪ</ZY>                  //ժҪ 
  --       <SFCYJ></SFCYJ>              //�Ƿ��Ԥ�����Һ�Ŀǰ���� 
  --       <SFXFK></SFXFK>              //�Ƿ����ѿ�,�Һ�Ŀǰ���� 
  --       <EXPENDLIST>                 //��չ��Ϣ 
  --         <EXPEND> 
  --           <JYMC>��������</JYMC>        //�������� 
  --           <JYLR>��������<JYLR>         //�������� 
  --         </EXPEND> 
  --         <EXPEND> 
  --           ... 
  --         </EXPEND> 
  --       </EXPENDLIST> 
  --     </JS> 
  --   </JSLIST> 
  --   <HZDW>������λ</HZDW>        //������λ���� 
  --   <YYFS>֧����<YYFS>    //ԤԼ��ʽ,����������֧���� 
  --   <BRID>����ID</BRID>     //����ID 
  --   <SFZH>���֤��</SFZH>     //���֤�� 
  --   <XM>����</XM>            //���� 
  --   <BRLX></BRLX>             //ҽ���������� 
  --   <FB>��ͨ</FB>               //���˷ѱ𣬿��Բ��� 
  --   <JQM>������</JQM>            //������ 
  --</IN> 

  --����:Xml_Out 
  --<OUTPUT> 
  -- <GHDH>�Һŵ���</GHDH>          //�Һŵ��� 
  -- <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ�� 
  -- <JZID>����ID</JZID>          //���ν���ID 
  -- <ERROR><MSG>������Ϣ</MSG></ERROR>  //����ʱ���� 
  --</ OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_����     �ҺŰ���.����%Type;
  d_����ʱ�� Date;
  d_ԭʼʱ�� Date;
  d_�Ǽ�ʱ�� Date;

  n_Ӧ�ս��   ������ü�¼.Ӧ�ս��%Type;
  v_��ˮ��     ����Ԥ����¼.������ˮ��%Type;
  v_˵��       ������ü�¼.ժҪ%Type;
  n_����id     ������Ϣ.����id%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  v_ԤԼ��ʽ   ԤԼ��ʽ.����%Type;
  v_��������� ҽ�ƿ����.����%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  n_�����     ������ü�¼.��ʶ��%Type;
  v_����       ������ü�¼.����%Type;
  v_�Ա�       ������ü�¼.�Ա�%Type;
  v_����       ������ü�¼.����%Type;
  v_���ʽ   ������ü�¼.���ʽ%Type;
  v_�ѱ�       ������ü�¼.�ѱ�%Type;
  v_No         ���˹Һż�¼.No%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_�շ�ϸĿid ������ü�¼.�շ�ϸĿid%Type;
  n_���˿���id ������ü�¼.���˿���id%Type;
  n_��������id ������ü�¼.��������id%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_ҽ������   �ҺŰ���.ҽ������%Type;
  n_ҽ��id     �ҺŰ���.ҽ��id%Type;
  n_����id     ������ü�¼.����id%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_�Ű�       �ҺŰ���.����%Type;
  n_����id     �ҺŰ���.Id%Type;
  n_�ƻ�id     �ҺŰ��żƻ�.Id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_��ſ���   �ҺŰ���.��ſ���%Type;
  n_����       �Һ����״̬.���%Type;
  v_����       �ҺŰ�������.������Ŀ%Type;
  v_��������   ������Ϣ.��������%Type;
  n_����       Number(3);
  v_�ֽ�       ���㷽ʽ.����%Type;
  n_��ʱ��     Number(3);
  v_��������   Varchar2(3000);
  v_������λ   ���˹Һż�¼.������λ%Type;
  v_������     �Һ����״̬.������%Type;
  n_�ɿʽ   Number(3);
  n_�Һ�ģʽ   Number(3);
  n_Exists     Number(3);
  v_���ս���   Varchar2(1000);
  n_��¼id     �ٴ������¼.Id%Type;
  v_Temp       Varchar2(32767); --��ʱXML 
  x_Templet    Xmltype; --ģ��XML 
  v_Err_Msg    Varchar2(200);
  d_����ʱ��   Date;
  n_Count      Number(3);
  v_�����     �������׼�¼.���%Type;
  n_��Ԥ��     ����Ԥ����¼.��Ԥ��%Type;
  v_Para       Varchar2(2000);
  Err_Item Exception;
  Err_Special Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into v_����, n_����, d_ԭʼʱ��, n_Ӧ�ս��, v_ԤԼ��ʽ, v_������λ, n_����id, v_��������, v_�ѱ�, v_������, n_�ɿʽ, n_��¼id, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����!';
    Raise Err_Item;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  If Sysdate - 10 > Nvl(d_����ʱ��, Sysdate - 30) Then
    If n_�Һ�ģʽ = 1 And Nvl(d_ԭʼʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_��¼id Is Null Then
      v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
      Raise Err_Item;
    End If;
  Else
    If n_�Һ�ģʽ = 1 And Nvl(d_ԭʼʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_��¼id Is Null Then
      Begin
        Select a.Id
        Into n_��¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_���� And Nvl(d_ԭʼʱ��, Sysdate) Between a.��ʼʱ�� And a.��ֹʱ��;
      Exception
        When Others Then
          v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  d_�Ǽ�ʱ�� := Sysdate;
  d_����ʱ�� := Trunc(d_ԭʼʱ��);

  For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    --��Ԥ������Ҫ���������� 
    If Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 0 Then
      If c_���׼�¼.���㿨��� Is Null Then
        v_����� := c_���׼�¼.���㷽ʽ;
      Else
        Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
      
        If Nvl(n_Count, 0) = 1 Then
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
        Else
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
        End If;
      End If;
    
      If v_����� Is Null Then
        v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 4) = 0 Then
        v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  If v_�������� Is Not Null Then
    Begin
      Select 1 Into n_���� From �������� Where ���� = v_��������;
    Exception
      When Others Then
        v_Err_Msg := 'û�з���Ϊ(' || v_�������� || ')�Ĳ�������';
        Raise Err_Item;
    End;
    Update ������Ϣ Set �������� = Nvl(��������, v_��������) Where ����id = n_����id;
  End If;

  Select a.�����, a.����, a.�Ա�, a.����, Nvl(b.����, c.����)
  Into n_�����, v_����, v_�Ա�, v_����, v_���ʽ
  From ������Ϣ A, ҽ�Ƹ��ʽ B, (Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = '1' And Rownum < 2) C
  Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = b.����(+);

  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_��������id From Dual;

  v_No := Nextno(12);
  Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;

  If n_��¼id Is Null Then
    For r_���� In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_����.�Ƿ��Ԥ��, 0) = 0 Then
        If r_����.���㷽ʽ Is Null Then
          Begin
            Select b.���㷽ʽ, b.Id
            Into v_���㷽ʽ, n_�����id
            From ҽ�ƿ���� B
            Where b.���� = r_����.���㿨��� And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
              Raise Err_Item;
          End;
        Else
          Select Nvl(Max(1), 0) Into n_Exists From ���㷽ʽ Where ���� = r_����.���㷽ʽ And ���� In (3, 4);
          If n_Exists = 1 Then
            v_���ս��� := v_���ս��� || '||' || r_����.���㷽ʽ || '|' || r_����.������;
          Else
            If v_���㷽ʽ Is Null Then
              v_���㷽ʽ := r_����.���㷽ʽ;
            Else
              v_Err_Msg := 'Ŀǰ�ƻ��Ű�ҺŲ�֧�ַ�ҽ����Ķ��ֽ��㷽ʽ,����!';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        If r_����.���㿨��� Is Not Null Then
          v_��������� := r_����.���㿨���;
          v_���㿨��   := r_����.���㿨��;
          v_��ˮ��     := r_����.������ˮ��;
          v_˵��       := r_����.����˵��;
        
          If n_�����id Is Null Then
            Begin
              Select b.Id Into n_�����id From ҽ�ƿ���� B Where b.���� = r_����.���㿨��� And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_����.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(r_����.���㿨���);
          Else
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = r_����.���㿨���;
          End If;
        Else
          v_����� := r_����.���㷽ʽ;
        End If;
      
        Update �������׼�¼
        Set ҵ�����id = n_����id
        Where ��ˮ�� = r_����.������ˮ�� And ��� = v_����� And ҵ������ = 4;
      Else
        n_��Ԥ�� := r_����.������;
      End If;
    End Loop;
  
    If v_���ս��� Is Not Null Then
      v_���ս��� := Substr(v_���ս���, 3);
    End If;
  
    Select Decode(To_Char(d_ԭʼʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                   Null)
    Into v_����
    From Dual;
  
    Begin
      Select ID
      Into n_�ƻ�id
      From (Select ID
             From �ҺŰ��żƻ�
             Where ���� = v_���� And d_ԭʼʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   ʧЧʱ�� And ���ʱ�� Is Not Null
             Order By ��Чʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select ID Into n_����id From �ҺŰ��� Where ���� = v_����;
    End;
  
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      --�Ӽƻ���ȡ��Ϣ 
      Select a.��Ŀid, b.����id, a.ҽ������, a.ҽ��id,
             Decode(To_Char(d_����ʱ��, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7', a.����,
                     Null), Nvl(a.��ſ���, 0)
      Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
      From �ҺŰ��żƻ� A, �ҺŰ��� B
      Where a.Id = n_�ƻ�id And b.Id = a.����id;
      Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum < 2;
    
      --������λ��� 
      If v_������λ Is Not Null Then
        Begin
          Select 1 Into n_���� From ������λ�ƻ����� Where �ƻ�id = n_�ƻ�id And ���� = 0 And ������λ = v_������λ;
        Exception
          When Others Then
            n_���� := 0;
        End;
      End If;
    
      If n_���� = 1 Then
        v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
        Raise Err_Item;
      End If;
    
      If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
        d_����ʱ�� := d_ԭʼʱ��;
        Select ���
        Into n_����
        From �Һżƻ�ʱ��
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_����ʱ��
          From �Һżƻ�ʱ��
          Where �ƻ�id = n_�ƻ�id And ���� = v_���� And ��� = Nvl(n_����, 0);
        Exception
          When Others Then
            If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From �Һżƻ�ʱ��
              Where �ƻ�id = n_�ƻ�id And ���� = v_����;
            Else
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From ʱ���
              Where ʱ��� = v_�Ű�;
            End If;
            If d_����ʱ�� < d_�Ǽ�ʱ�� Then
              d_����ʱ�� := d_�Ǽ�ʱ��;
            End If;
        End;
      End If;
    Else
      --�Ӱ��Ŷ�ȡ��Ϣ 
      Select b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id,
             Decode(To_Char(d_����ʱ��, 'D'), '1', b.����, '2', b.��һ, '3', b.�ܶ�, '4', b.����, '5', b.����, '6', b.����, '7', b.����,
                     Null), Nvl(b.��ſ���, 0)
      Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
      From �ҺŰ��� B
      Where b.Id = n_����id;
      Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum < 2;
    
      --������λ��� 
      If v_������λ Is Not Null Then
        Begin
          Select 1 Into n_���� From ������λ���ſ��� Where ����id = n_����id And ���� = 0 And ������λ = v_������λ;
        Exception
          When Others Then
            n_���� := 0;
        End;
      End If;
    
      If n_���� = 1 Then
        v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
        Raise Err_Item;
      End If;
    
      If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
        d_����ʱ�� := d_ԭʼʱ��;
        Select ���
        Into n_����
        From �ҺŰ���ʱ��
        Where ����id = n_����id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_����ʱ��
          From �ҺŰ���ʱ��
          Where ����id = n_����id And ���� = v_���� And ��� = Nvl(n_����, 0);
        Exception
          When Others Then
            If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From �ҺŰ���ʱ��
              Where ����id = n_����id And ���� = v_����;
            Else
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From ʱ���
              Where ʱ��� = v_�Ű�;
            End If;
            If d_����ʱ�� < d_�Ǽ�ʱ�� Then
              d_����ʱ�� := d_�Ǽ�ʱ��;
            End If;
        End;
      End If;
    End If;
  
    If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(3, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      End If;
    Else
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(1, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      End If;
    End If;
  
    For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --������Ű�ģʽ 
    For r_���� In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_����.�Ƿ��Ԥ��, 0) = 0 Then
        If r_����.���㷽ʽ Is Null Then
          Begin
            Select b.���㷽ʽ, b.Id
            Into v_���㷽ʽ, n_�����id
            From ҽ�ƿ���� B
            Where b.���� = r_����.���㿨��� And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
              Raise Err_Item;
          End;
          v_�������� := v_�������� || '|' || v_���㷽ʽ || ',' || r_����.������ || ',,';
        Else
          v_�������� := v_�������� || '|' || r_����.���㷽ʽ || ',' || r_����.������ || ',,';
        End If;
      
        If r_����.���㿨��� Is Not Null Then
          v_��������   := v_�������� || '1';
          v_��������� := r_����.���㿨���;
          v_���㿨��   := r_����.���㿨��;
          v_��ˮ��     := r_����.������ˮ��;
          v_˵��       := r_����.����˵��;
          If n_�����id Is Null Then
            Begin
              Select b.Id Into n_�����id From ҽ�ƿ���� B Where b.���� = r_����.���㿨��� And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_����.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(r_����.���㿨���);
          Else
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = r_����.���㿨���;
          End If;
        Else
          v_�������� := v_�������� || '0';
          v_�����   := r_����.���㷽ʽ;
        End If;
      
        Update �������׼�¼
        Set ҵ�����id = n_����id
        Where ��ˮ�� = r_����.������ˮ�� And ��� = v_����� And ҵ������ = 4;
      Else
        n_��Ԥ�� := r_����.������;
      End If;
    End Loop;
  
    If v_�������� Is Not Null Then
      v_�������� := Substr(v_��������, 2);
    Else
      Begin
        Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
      Exception
        When Others Then
          v_�ֽ� := '�ֽ�';
      End;
      v_�������� := v_�ֽ� || ',' || 0 || ',,0';
    End If;
  
    Select ��Ŀid, ����id, ҽ������, ҽ��id, �Ƿ���ſ���, �Ƿ��ʱ��
    Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, n_��ſ���, n_��ʱ��
    From �ٴ������¼
    Where ID = n_��¼id;
  
    Begin
      Select ��ʼʱ�� Into d_����ʱ�� From �ٴ�������ſ��� Where ��¼id = n_��¼id And ��� = n_����;
    Exception
      When Others Then
        d_����ʱ�� := d_ԭʼʱ��;
    End;
  
    If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(3, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      End If;
    Else
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(1, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      End If;
    End If;
  
    For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_���˹Һż�¼_Insert
(
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0,
  ������������_In  Number := 0,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null,
  ���½������_In  Number := 1
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  --     ���½������_In:0-��zl_��Ա�ɿ����_Update �и��� 1-�ڱ������и���
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, Null)) As �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), �տ�ʱ��;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ����
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;

  n_����id        ������ü�¼.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_��ǰ���      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_��id           ����ɿ����.Id%Type;
  n_�����         ������Ϣ.�����%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  d_�Ŷ�ʱ��       Date;
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type := 0;
  v_����           �ҺŰ�������.������Ŀ%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;

  n_�ҳ��������� Number(4) := 0;
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_Temp           Varchar2(3000);
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   ʱ���.��ʼʱ��%Type;
  d_������ʱ��   ʱ���.��ֹʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  n_��ʱ����ʾ     Number(3);
  d_����ʱ��       Date;
Begin
  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);

  If �ѱ�_In Is Null Then
    Begin
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
        Raise Err_Item;
    End;
  Else
    v_�ѱ� := �ѱ�_In;
    If Nvl(�������˷ѱ�_In, 0) = 1 Then
      Begin
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(������������_In, 0) = 1 Then
    Begin
      Update ������Ϣ Set ���� = ����_In Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
        Raise Err_Item;
    End;
  End If;

  If �����_In Is Not Null Then
    Begin
      Select Nvl(�����, 0) Into n_����� From ������Ϣ Where ����id = ����id_In;
    Exception
      When Others Then
        n_����� := 0;
    End;
    If n_����� = 0 Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In;
    End If;
  End If;

  Begin
    Delete From �Һ����״̬
    Where ���� = �ű�_In And ���� = ����ʱ��_In And ��� = ����_In And ״̬ = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;
  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
    If d_����ʱ�� Is Not Null Then
      If ����ʱ��_In > d_����ʱ�� Then
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    --��ȡ���㷽ʽ����
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
    Begin
      Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
    Exception
      When Others Then
        v_�����ʻ� := '�����ʻ�';
    End;
  End If;

  n_��� := ����_In;
  Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
  Into v_����
  From Dual;

  --�ҺŻ�ȡ����
  Begin
    Select a.Id, a.��ſ���, Nvl(b.�޺���, 0), Nvl(b.��Լ��, 0)
    Into n_����id, n_��ſ���, n_�޺���, n_��Լ��
    From �ҺŰ��� A, �ҺŰ������� B
    Where a.Id = b.����id(+) And b.������Ŀ(+) = v_���� And a.���� = �ű�_In;
  
  Exception
    When Others Then
      n_����id := -1;
  End;

  --����ǲ����ѻ��ߺű�Ϊ��ʱ�����
  If Nvl(������_In, 0) = 0 Or �ű�_In Is Not Null Then
    If n_����id = -1 Then
      v_Err_Msg := '������Ӧ�ĹҺŰ�������,����';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
    --���Ȼ�ȡ�ƻ�
    Begin
      Select ID
      Into n_�ƻ�id
      From �ҺŰ��żƻ�
      Where ����id = n_����id And ���ʱ�� Is Not Null And
            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.��Чʱ��) As ��Ч
             From �ҺŰ��żƻ� A
             Where a.���ʱ�� Is Not Null And ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   a.ʧЧʱ�� And a.����id = n_����id) And
            ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And ʧЧʱ��;
    
    Exception
      When Others Then
        n_�ƻ�id := 0;
    End;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Begin
        --��ȡ�ƻ�������
        Select a.Id, a.��ſ���, Nvl(b.�޺���, 0) As �޺���, Nvl(b.��Լ��, 0) As ��Լ��
        Into n_�ƻ�id, n_��ſ���, n_�޺���, n_��Լ��
        From �ҺŰ��żƻ� A, �Һżƻ����� B
        Where a.���� = �ű�_In And a.Id = n_�ƻ�id And a.���ʱ�� Is Not Null And a.Id = b.�ƻ�id(+) And b.������Ŀ(+) = v_����;
      Exception
        When Others Then
          v_Err_Msg := '������Ӧ�ĹҺŰ��Ż�ƻ�����,����';
          Raise Err_Item;
      End;
    End If;
  End If;

  --��ȡ�Ƿ��ʱ��
  Begin
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum <= 1;
      Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum <= 1;
      Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null And ���_In = 1 Then
    --����Ƿ��ģʽ�ҺŰ���
    Select To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_��鿪ʼʱ��, d_������ʱ��
    From ʱ���
    Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
    If d_��鿪ʼʱ�� > d_������ʱ�� Then
      d_������ʱ�� := d_������ʱ�� + 1;
    End If;
    If d_��鿪ʼʱ�� < d_����ʱ�� And d_������ʱ�� > d_����ʱ�� Then
      --��ȡ�����¼id
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = �ű�_In And �ϰ�ʱ�� = v_ʱ��� And ����ʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
    End If;
  End If;

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    --����ʱ��_in>Sysdate ����ʱ��>����ʱ��ʱ��--����_in is null
    Begin
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And Nvl(��������, 0) <> 0;
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_׷�Ӻ� := 0;
    End;
  End If;
  d_ʱ��ʱ�� := ����ʱ��_In;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 Then
    --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
    Begin
      Select Nvl(���, 0),
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
      Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And
            (���, ����id, ����) In (Select Nvl(Max(���), -1), ����id, ����
                               From �ҺŰ���ʱ��
                               Where ����id = n_����id And ���� = v_���� And
                                     Decode(��������_In + n_׷�Ӻ�, 0, To_Char(����ʱ��_In, 'hh24:mi'),
                                            To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By ����id, ����);
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ�� > 0 Then
    --ԤԼ��,ȡ�ƻ�
    Begin
      If Nvl(n_�ƻ�id, 0) = 0 Then
        --û�ƻ���Ч,ȡ���ŵ�����
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ҺŰ���ʱ�� C
        Where ����id = n_����id And ���� = v_���� And
              (���, ����id, ����) In
              (Select Nvl(Max(c.���), -1), ����id, ����
               From �ҺŰ���ʱ�� C
               Where ����id = n_����id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By ����id, ����);
      Else
        --�мƻ���Чȡ�ƻ�
        --û��Ч�������ǴӹҺżƻ�ʱ�β�ѯ
        Select Nvl(���, -1),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �Һżƻ�ʱ�� C
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And
              (���, �ƻ�id, ����) In
              (Select Nvl(Max(c.���), -1), �ƻ�id, ����
               From �Һżƻ�ʱ�� C
               Where �ƻ�id = n_�ƻ�id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By �ƻ�id, ����);
      End If;
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 Then
  
    --��ȡ��ǰδʹ�õ����
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>
      Begin
        --������
        If �˺�����_In = 1 Then
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
      End;
      If n_��� Is Null Then
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end>
    
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 Then
        --<����ԤԼ�Һ�-->
        Begin
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.���� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
          Into n_�ѹ���, n_��Լ��
          From �Һ����״̬ A
          Where a.���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And
                ״̬ Not In (4, 5);
        Exception
          When Others Then
            n_�ѹ��� := 0;
            n_��Լ�� := 0;
        End;
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
        If n_�޺��� <= n_�ѹ��� Then
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
          Raise Err_Item;
        End If;
      End If;
    
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷��
    
      --��ȡ����ҳ���������
      If �˺�����_In = 1 Then
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ Not In (4, 5);
      Else
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ <> 5;
      End If;
    
      --�������
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_��Լ�� + 1;
      If n_��� <= Nvl(n_�ҳ���������, 0) Then
        n_��� := Nvl(n_�ҳ���������, 0) + 1;
      End If;
    
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        If �˺�����_In = 1 Then
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
          n_�ѹ���   := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111));
        n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111));
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ����), 1, 1, 0))
            Into n_ʧЧ��
            From �Һ����״̬
            Where ���� = �ű�_In And ���� Between Trunc(Sysdate) And Sysdate And Nvl(ԤԼ, 0) = 1 And ״̬ = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      Begin
        Select �ѹ���, ��Լ��
        Into n_��������, n_��Լ��
        From ���˹ҺŻ���
        Where ���� = Trunc(����ʱ��_In) And ���� = �ű�_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  End If;

  --���¹Һ����״̬
  If ���_In = 1 And Not n_��� Is Null Then
    If n_��ʱ�� = 1 Then
      d_���ʱ�� := ����ʱ��_In;
    Else
      d_���ʱ�� := Trunc(����ʱ��_In);
    End If;
    --������ŵĴ���
    Begin
      Select ����Ա����, ������
      Into v_��Ų���Ա, v_��Ż�����
      From �Һ����״̬
      Where ״̬ = 5 And ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
      n_���� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_����       := 0;
    End;
    If n_���� = 0 Then
      Update �Һ����״̬
      Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
      Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 3 And ����Ա���� = ����Ա����_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_��ʱ��, 0) = 0 Or Nvl(ԤԼ�Һ�_In, 0) = 1 Or (Nvl(n_��ſ���, 0) = 0 And Nvl(����_In, 0) = 0) Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
            Values
              (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          Elsif Nvl(n_��ʱ��, 0) > 0 Then
            --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
            Update �Һ����״̬
            Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In, ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
            Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 2;
            If Sql%NotFound Then
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
              Values
                (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        Update �Һ����״̬
        Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
        Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 5 And ����Ա���� = ����Ա����_In And ������ = v_������;
        If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
          Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
        End If;
      End If;
    End If;
  End If;

  If n_�����¼id Is Not Null Then
    Update �ٴ�������ſ���
    Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
    Where ��¼id = n_�����¼id And ��� = n_���;
    If ԤԼ�Һ�_In = 1 Then
      Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
    Else
      If ԤԼ����_In = 1 Then
        Update �ٴ������¼
        Set ��Լ�� = ��Լ�� + 1, �ѹ��� = �ѹ��� + 1, �����ѽ��� = �����ѽ��� + 1
        Where ID = n_�����¼id;
      Else
        Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
      End If;
    End If;
  End If;

  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, Decode(�շѵ�_In, Null, ժҪ_In, '����:' || �շѵ�_In), ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
  
    If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), �Ǽ�ʱ��_In,
         ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    
      If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
        Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, �ֽ�֧��_In, n_Ԥ��id, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In);
      End If;
    End If;
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
      n_Ԥ����� := Ԥ��֧��_In;
      For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
        n_��ǰ��� := Case
                    When r_Deposit.��� - n_Ԥ����� < 0 Then
                     r_Deposit.���
                    Else
                     n_Ԥ�����
                  End;
      
        If r_Deposit.����id = 0 Then
          --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
        
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2);
        --����Ƿ��Ѿ�������
        If r_Deposit.��� < n_Ԥ����� Then
          n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
        Else
          n_Ԥ����� := 0;
        End If;
      
        If n_Ԥ����� = 0 Then
          Exit;
        End If;
      End Loop;
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 And Nvl(���½������_In, 1) = 1 Then
      If Nvl(�ֽ�֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + �ֽ�֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�)
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, Nvl(���㷽ʽ_In, v_�ֽ�), 1, �ֽ�֧��_In);
          n_����ֵ := �ֽ�֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End If;
    
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �շѵ�)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �շѵ�_In);
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(ִ�в���id_In, 0)));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
        n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then
          n_��ʱ����ʾ := 1;
        Else
          n_��ʱ����ʾ := Null;
        End If;
      
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         n_��ʱ����ʾ, v_�Ŷ����);
      
        --�Һ������Ŷ�
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
        End If;
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) >= Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_Insert;
/

--133584:���ϴ�,2018-11-12,����ʧЧʱ���ѯ�ҺŰ��żƻ�
Create Or Replace Procedure Zl_���˹Һż�¼_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲����� 3-�˸��ӷ� 4-�˹Һ��벡�� 5-�˹Һ��븽��
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  �ջ�Ʊ�ݺ�_In   Varchar2 := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷�
  --    2.�����ӿ�,��ʱ��֧��
  -- �ҺŷѲ����ѷֿ���,����
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ���
  --    Ԥ����:Ԥ����,�˲���
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������
  --    ���ѿ�:ԭ�������ò����������ѿ�
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö�����
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, �ҺŰ��� B, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.��� = 1 And a.��¼״̬ = v_״̬ And c.No = a.No And c.ִ���� = d.����(+) And a.No = ���ݺ�_In And
          Nvl(a.���㵥λ, '�ű�') = c.�ű� And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ���
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney(n_Id ����Ԥ����¼.����id%Type) Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ����Ԥ����¼ B
    Where b.����id = n_Id And b.��¼���� = 4 And b.��¼״̬ = 2 And Nvl(b.��Ԥ��, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ���
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;
  d_����ʱ��       Date;
  d_����ʱ��       ���˹Һż�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_����           �ҺŰ���.����%Type;
  n_���           ���˹Һż�¼.����%Type;
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   Date;
  d_������ʱ��   Date;
  v_Temp           Varchar2(500);
  v_����ids        Varchar2(500);
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;

  Select �ű�, ����, ����ʱ�� Into v_����, n_���, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  Begin
    Select a.Id Into n_����id From �ҺŰ��� A Where a.���� = v_����;
  Exception
    When Others Then
      n_����id := -1;
  End;

  Begin
    Select ID
    Into n_�ƻ�id
    From �ҺŰ��żƻ�
    Where ����id = n_����id And ���ʱ�� Is Not Null And
          Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
          (Select Max(a.��Чʱ��) As ��Ч
           From �ҺŰ��żƻ� A
           Where a.���ʱ�� Is Not Null And d_����ʱ�� Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                 a.ʧЧʱ�� And a.����id = n_����id) And
          d_����ʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And ʧЧʱ��;
  Exception
    When Others Then
      n_�ƻ�id := 0;
  End;

  Begin
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
    --����Ƿ��ģʽ�ҺŰ���
    Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_��鿪ʼʱ��, d_������ʱ��
    From ʱ���
    Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
    If d_��鿪ʼʱ�� > d_������ʱ�� Then
      d_������ʱ�� := d_������ʱ�� + 1;
    End If;
    If d_������ʱ�� > d_����ʱ�� Then
      --��ȡ�����¼id
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_���� And �ϰ�ʱ�� = v_ʱ��� And d_����ʱ�� Between ��ʼʱ�� And ��ֹʱ��;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
    End If;
  End If;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_����ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_����ids := Null;
  End;

  --1.ԤԼ����
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ��
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
    Close c_Registinfo;
  
    --���¹Һ����״̬
    Delete �Һ����״̬
    Where ״̬ = 2 And
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, Trunc(����ʱ��)
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In) Or
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, ����ʱ��
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In);
  
    --��Ӳ��˹Һż�¼�� ������¼
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    If n_�����¼id Is Not Null Then
      Update �ٴ������¼ Set ��Լ�� = ��Լ�� - 1 Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
      Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
    End If;
  
    --Update ���˹Һż�¼ set ժҪ=nvl(ժҪ_IN,ժҪ) where NO=���ݺ�_IN;
    --ɾ��������ü�¼
    Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    --���ԤԼ���ɶ���ʱ��Ҫ�������
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ���
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    --���ǹ��˲�����ʱ����
    --���¹Һ����״̬
    If �˺�����_In = 1 Then
      Delete �Һ����״̬
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    Else
      Update �Һ����״̬
      Set ״̬ = 4
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    End If;
  
    --���˾���״̬
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) Not In (2, 3) Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹�����
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  If Nvl(�˷�����_In, 0) = 0 Or Nvl(�˷�����_In, 0) = 2 Then
    --ȫ��,�˲�����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 1 Or Nvl(�˷�����_In, 0) = 4 Then
    --�˹Һŷ�,�˹Һ��벡����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 3 Then
    --�˸��ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 5 Then
    --�˹Һ��븽�ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    End If;
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
  End If;

  If n_���� = 1 Then
    --����
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
      Returning ������� Into n_���ض�;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
        n_���ض� := Nvl(c_����.ʵ�ս��, 0);
      End If;
      If Nvl(n_���ض�, 0) = 0 Then
        Delete �������
        Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
    Delete ����δ�����
    Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
  End If;

  If n_���� = 0 Then
    --1.�˷�
    --���˹ҺŽ���:�ֽ�͸����ʻ�����
    If ��ԭ���˽���_In Is Not Null Then
      --�˿����ȡ
      If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
        --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
        Begin
          --��ȡ�����˿���
          Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
          Into n_�˿���
          From ������ü�¼
          Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
        
        Exception
          When Others Then
            v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                           When 1 Then
                            '�Һŷ���'
                           When 2 Then
                            '������'
                         End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
            Raise Err_Item;
        End;
        Begin
          Select ��Ԥ��
          Into n_�˷ѽ��
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        Exception
          When Others Then
            n_�˷ѽ�� := 0;
        End;
      
        --a.����Ľ��㷽ʽ
      
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
      
        If n_�˷ѽ�� = 0 Then
          --b.����������ֽ�
          If n_�˿��� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ�
              Begin
                Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
              Exception
                When Others Then
                  v_��ָ�����㷽ʽ := '�ֽ�';
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���), ����˵�� = Nvl(����˵��_In, ����˵��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                 �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                       Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --a.����Ľ��㷽ʽԭ����
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
      
        --b.����������ֽ�
        Begin
          Select Sum(��Ԥ��)
          Into n_�˷ѽ��
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
        Exception
          When Others Then
            n_�˷ѽ�� := 0;
        End;
        If n_�˷ѽ�� <> 0 Then
          If v_��ָ�����㷽ʽ Is Null Then
            --�˸��ֽ�
            Begin
              Select ���㷽ʽ
              Into v_��ָ�����㷽ʽ
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
            
            Exception
              When Others Then
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
            End;
          End If;
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��), ����˵�� = Nvl(����˵��_In, ����˵��)
          Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                     ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                     Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
              From ����Ԥ����¼ A
              Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
          End If;
        End If;
      End If;
    Else
      --�˿����ȡ
      If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
        --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
        Begin
          --��ȡ�����˿���
          Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
          Into n_�˿���
          From ������ü�¼
          Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
        Exception
          When Others Then
            v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                           When 1 Then
                            '�Һŷ���'
                           When 2 Then
                            '������'
                         End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
            Raise Err_Item;
        End;
      End If;
      If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
        --�״�ȫ��
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -1 * ��Ԥ��,
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
      Else
        --�����˷�,���߱��ε���һ����
        --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
           ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                 -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And ��Ԥ�� = n_�˿��� And
                Rownum < 2;
        If Sql%RowCount = 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And Rownum < 2;
        End If;
        If Sql%RowCount = 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
          If Sql%RowCount = 0 Then
            --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������
            n_Ԥ����� := n_�˿���;
          End If;
        End If;
      
      End If;
    End If;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ�
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ�����
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
    End If;
  
    --������Ԥ�����
    For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_����id
                 Group By ����id, Ԥ�����) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
      Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, Ԥ�����, ����, ����)
        Values
          (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
        n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End Loop;
  
    If �ջ�Ʊ�ݺ�_In Is Not Null Then
      --���˹Һŷ�,������Ʊ��
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�δ�ӡ��������ȡ
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_��ӡid := Null;
      End;
    
      --���ջ�ԭƱ��
      If n_��ӡid Is Not Null Then
        Begin
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        Exception
          When Others Then
            Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
            Insert Into Ʊ��ʹ����ϸ
              (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
              Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
              From Ʊ��ʹ����ϸ
              Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼
  --��ػ��ܱ�Ĵ���

  --���˹ҺŻ���
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    Open c_Registinfo(3);
    Fetch c_Registinfo
      Into r_Registrow;
  
    If c_Registinfo%RowCount = 0 Then
      --ֻ�ղ�����ʱ�޺ű�,������
      Close c_Registinfo;
    Else
    
      --��Ҫȷ���Ƿ�ԤԼ�Һ�
      --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
      --2.����������Һ�,��ֻ���ѹ���
    
      Begin
        Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
      Exception
        When Others Then
          n_ԤԼ�Һ� := 0;
      End;
    
      Update ���˹ҺŻ���
      Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
      End If;
    
      If n_�����¼id Is Not Null Then
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
        Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
        Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_Opermoney In c_Opermoney(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If n_�Һ����ɶ��� <> 0 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
        End If;
      End Loop;
    End If;
  
    --ҽ�������ľ���ǼǼ�¼
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --���˹Һż�¼
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 2, ���ݺ�_In;
  Exception
    When Others Then
      Null;
  End;
  b_Message.Zlhis_Regist_003(n_�Һ�id, ���ݺ�_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_Delete;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0037' Where ���=&n_System;
Commit;