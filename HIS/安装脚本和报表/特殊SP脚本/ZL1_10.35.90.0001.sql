----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------



------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--114815:������,2018-03-09,��Ⱦ������ָ����ӡ����
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1285,'���鴫Ⱦ�������ӡ','�鿴�ʹ�ӡ���鴫Ⱦ������',&n_System,'zl9LisInsideComm');

Insert Into zlMenus
  (���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ϵͳ, ģ��)
  Select 'ȱʡ', Zlmenus_Id.Nextval, ID, '���鴫Ⱦ�������ӡ', '���鴫Ⱦ��', Null, 105, '�鿴�ʹ�ӡ���鴫Ⱦ������', &n_System, 1285
  From zlMenus
  Where ϵͳ = &n_System And ��� = 'ȱʡ' And ���� = '��Ⱦ������ϵͳ' And ģ�� Is Null;

--122609:������,2018-03-08,����ƽ̨��Ϣ���
Insert Into Zlmsg_Lists(Bz_Type, Code, Name, Key_Define, Note)
Select '�ٴ�', 'ZLHIS_CIS_004', '��������ҽ��', '<root><����ID></����ID><��ҳID></��ҳID><ҽ��״̬></ҽ��״̬><ID></ID></root>', 'סԺ��ʿ����վ:У������ҽ��ʱ;סԺҽ������վ:��������ҽ������ʱ'  From Dual Union All 
Select '�ٴ�', 'ZLHIS_CIS_005', '������������ҽ��', '<root><����ID></����ID><��ҳID></��ҳID><ҽ��״̬></ҽ��״̬><ID></ID></root>', 'סԺҽ������վ/סԺ��ʿ����վ:����סԺ��������ҽ��ʱ'  From Dual Union All 
Select '�ٴ�', 'ZLHIS_CIS_006', '���߻�����ҽ��', '<root><����ID></����ID><��ҳID></��ҳID><���ͺ�></���ͺ�><ID></ID></root>', 'סԺ��ʿ����վ:���͵�ҽ��Ϊ������ҽ��ʱ'  From Dual Union All 
Select '�ٴ�', 'ZLHIS_CIS_007', '�������߻�����ҽ��', '<root><����ID></����ID><��ҳID></��ҳID><���ͺ�></���ͺ�><ID></ID><NO></NO></root>', 'סԺ��ʿ����վ:���˻�����ҽ������ʱ'  From Dual;




-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--114815:������,2018-03-09,��Ⱦ������ָ����ӡ����
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1285,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
  Select '����',1,'����Ȩ�ޡ�',1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;





-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--122592:����,2018-03-09,�޷�ͬ������ѹ
CREATE OR REPLACE Procedure Zl_���˻�������_Update
(
  �ļ�id_In   In ���˻�������.�ļ�id%Type,
  ����ʱ��_In In ���˻�������.����ʱ��%Type,
  ��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1��ǩ����¼=5����ǩ��¼=15
  ��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
  ��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
  ���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null,
  ���˼�¼_In In Number := 1,
  ������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
  ��ǩ_In     In Number := 0,
  ����Ա_In   In ���˻�������.������%Type := Null,
  ��¼���_In In ���˻�����ϸ.��¼���%Type := Null, --���÷������(һ�����ݶ�Ӧ������ͬ��Ŀ����ϸ)
  ������_In In ���˻�����ϸ.������%Type := Null, --���÷������(��¼������Ŀ������������Ŀ���)
  δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null --��������洢ҽ��ID:���ͺ�
) Is
  Intins      Number(18);
  Int����     Number(1);
  n_Newid     ���˻�������.Id%Type;
  n_Oldid     ���˻�������.Id%Type;
  n_����      ���˻����ӡ.����%Type;
  n_Mutilbill Number(1);
  n_Syntend   Number(1);
  n_Synchro   Number(1);
  n_δ��˵��  Number(1);
  n_����      Number(1);
  n_Num       Number(18);
  v_Name      ���¼�¼��Ŀ.��¼��%Type;

  n_�������     ���˻�������.�������%Type;
  v_����id       ���ű�.Id%Type;
  v_������       ��Ա��.����%Type;
  v_��¼��       ��Ա��.����%Type;
  n_�ļ�id       ���˻�������.�ļ�id%Type;
  n_��¼id       ���˻�������.Id%Type;
  n_��ϸid       ���˻�����ϸ.Id%Type;
  n_��Դid       ���˻�����ϸ.��Դid%Type;
  v_������Դ     ���˻�����ϸ.������Դ%Type;
  n_��߰汾     ���˻�����ϸ.��ʼ�汾%Type;
  n_��Ŀ����     �����¼��Ŀ.��Ŀ����%Type;
  n_����id       ���˻����ļ�.����id%Type;
  n_��ҳid       ���˻����ļ�.��ҳid%Type;
  n_Ӥ��         ���˻����ļ�.Ӥ��%Type;
  d_Ӥ����Ժʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  d_�ļ���ʼʱ�� ���˻����ļ�.��ʼʱ��%Type;
  --��ȡ�ò��˵�ǰ��������δ�����Ļ����ļ������ļ���ʼʱ��С�ڵ��ڼ�¼����ʱ����ļ��б�ͬ������ʹ��
  Cursor Cur_Fileformats Is
    Select a.Id As ��ʽid, b.Id As �ļ�id, a.����, a.����, b.Ӥ��
    From �����ļ��б� A, ���˻����ļ� B, ���˻����ļ� C, ���˻������� D
    Where a.���� = 3 And a.���� <> 1 And a.Id = b.��ʽid And b.Id <> c.Id And b.����ʱ�� Is Null And b.��ʼʱ�� <= d.����ʱ�� And
          (a.ͨ�� = 1 Or (a.ͨ�� = 2 And b.����id = c.����id)) And c.����id = b.����id And c.��ҳid = b.��ҳid And c.Ӥ�� = b.Ӥ�� And
          c.Id = d.�ļ�id And d.Id = n_��¼id And c.Id = �ļ�id_In
    Order By a.���;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --ȡ��¼ID
  Int����     := 0;
  n_��¼id    := 0;
  n_Mutilbill := 0;
  n_Syntend   := 0;
  n_δ��˵��  := 0;
  n_����      := 0;

  If ����Ա_In Is Null Then
    v_������ := Zl_Username;
  Else
    v_������ := ����Ա_In;
  End If;

  --����Ƕ�Ӧ��ݻ����ļ�ֵΪ1����ʾ��ͬ�����������ļ������򲻴����ļ�ͬ��
  n_Mutilbill := Zl_To_Number(zl_GetSysParameter('��Ӧ��ݻ����ļ�', 1255));
  --��������ݻ����ļ�֮������ͬ��,���Զ�ͬ��,����ͬ��
  n_Syntend := Zl_To_Number(zl_GetSysParameter('��������ͬ��', 1255));

  Begin
    Select ��¼�� Into v_Name From ���¼�¼��Ŀ Where ��Ŀ��� = ��Ŀ���_In;
  Exception
    When Others Then
      v_Name := '';
  End;

  Begin
    Select ID, �������
    Into n_��¼id, n_�������
    From ���˻�������
    Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  Exception
    When Others Then
      n_��¼id := 0;
  End;

  --����ǲ��Ǳ��˵ļ�¼
  ---------------------------------------------------------------------------------------------------------------------
  If ���˼�¼_In = 0 And n_��¼id > 0 And ��ǩ_In = 0 Then
    v_��¼�� := '';
    Begin
      Select ��¼��
      Into v_��¼��
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
    Exception
      When Others Then
        v_��¼�� := '';
    End;
    If v_��¼�� Is Not Null And v_��¼�� <> v_������ Then
      v_Error := '����Ȩ�޸����˵ǼǵĻ������ݣ�';
      Raise Err_Custom;
    End If;
  End If;

  --����Ƿ����
  Select ����id, ��ҳid, Nvl(Ӥ��, 0), ��ʼʱ��
  Into n_����id, n_��ҳid, n_Ӥ��, d_�ļ���ʼʱ��
  From ���˻����ļ�
  Where ID = �ļ�id_In;
  d_Ӥ����Ժʱ�� := Null;
  If n_Ӥ�� <> 0 Then
    Begin
      Select ��ʼִ��ʱ��
      Into d_Ӥ����Ժʱ��
      From ����ҽ����¼ B, ������ĿĿ¼ C
      Where b.������Ŀid + 0 = c.Id And b.ҽ��״̬ = 8 And Nvl(b.Ӥ��, 0) <> 0 And c.��� = 'Z' And
            Instr(',3,5,11,', ',' || c.�������� || ',', 1) > 0 And b.����id = n_����id And b.��ҳid = n_��ҳid And b.Ӥ�� = n_Ӥ��;
    Exception
      When Others Then
        d_Ӥ����Ժʱ�� := Null;
    End;
  End If;
  If d_Ӥ����Ժʱ�� Is Null Then
    v_����id := 0;
    Begin
      Select a.����id
      Into v_����id
      From ���˱䶯��¼ A, ���˻����ļ� B
      Where a.����id Is Not Null And a.����id = b.����id And a.��ҳid = b.��ҳid And b.Id = �ļ�id_In And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '59', 'YYYY-MM-DD HH24:MI:SS') >= a.��ʼʱ�� And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '00', 'YYYY-MM-DD HH24:MI:SS') < = Nvl(a.��ֹʱ��, Sysdate) Or
            a.��ֹʱ�� Is Null)) And Rownum < 2;
    Exception
      When Others Then
        v_����id := 0;
    End;
    If v_����id = 0 Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  Else
    If ����ʱ��_In < d_�ļ���ʼʱ�� Or ����ʱ��_In > d_Ӥ����Ժʱ�� Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  End If;

  --���������Դ<>0���˳�
  n_��Դid := 0;
  If n_��¼id > 0 Then
    Begin
      Select ������Դ, Nvl(��Դid, 0)
      Into v_������Դ, n_��Դid
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Nvl(��Ŀ���, 0) = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0);
    Exception
      When Others Then
        v_������Դ := 0;
    End;
    If v_������Դ > 0 And n_��Դid > 0 Then
      Return;
    End If;
  End If;

  --ȡ��߰汾
  Select Nvl(Max(Nvl(a.��ʼ�汾, 1)), 0) + 1, Count(b.Id)
  Into n_��߰汾, Intins
  From ���˻�����ϸ A, ���˻������� B
  Where b.Id = n_��¼id And a.��¼id = b.Id And Mod(a.��¼����, 10) = 5;

  --Ŀǰ�Ѿ�ǩ�������ݲ����޸ģ�ֻ������ǩģʽ�½����޸ģ�����ǩ_In=1
  If ��ǩ_In <> 1 And Intins > 0 Then
    v_Error := '����ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ����Ӧ�������Ѿ�ǩ������ǩ�����ܼ���������' || Chr(13) || Chr(10) ||
               '��������������粢����������ģ���ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;
  Intins := 0;

  --������ʱ,Ҫ������ݣ���ǩ����ʱ���Զ������ǩ�������޸ĵ����ݣ����Դ˴�ֻ�迼����ǩ���ɣ�
  If ��¼����_In Is Null Then
    Begin
      Select ID
      Into n_��ϸid
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Nvl(��Ŀ���, 0) = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
    Exception
      --�������˳�
      When Others Then
        Return;
    End;

    --���ҳ��˱���Ҫɾ�������ݣ��Ƿ񻹴�������Ч�����ݣ��������ֻɾ���������ݣ�����ɾ���˷���ʱ���Ӧ���������ݡ�
    Select Count(ID)
    Into Intins
    From ���˻�����ϸ
    Where ��¼id = n_��¼id And Mod(��¼����, 10) <> 5 And ��ֹ�汾 Is Null And ID <> n_��ϸid;
    If Intins = 0 Then
      Delete From ���˻�����ϸ Where ��¼id = n_��¼id;
    Else
      Delete From ���˻�����ϸ Where ID = n_��ϸid;
    End If;

    Delete From ���˻������� A
    Where a.Id = n_��¼id And Not Exists (Select 1 From ���˻�����ϸ B Where b.��¼id = a.Id);

    --�����ɾ��ǩ�����޸Ĳ��������һ������,��Ӧ��ǩ����¼����ֹ�汾��Ϊ��
    Begin
      Select 1
      Into Intins
      From ���˻�����ϸ
      Where ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null And ��¼���� = 1 And ��¼id = n_��¼id;
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Update ���˻�����ϸ Set ��ֹ�汾 = Null Where ��¼���� = 5 And ��ʼ�汾 = n_��߰汾 - 1 And ��¼id = n_��¼id;
    End If;
    If Nvl(n_�������, 0) <> 0 Then
      Return;
    End If;

    --############
    --�����������
    --############
    For Rsdel In (Select Distinct ��¼id From ���˻�����ϸ Where ��Դid = n_��ϸid) Loop

      Delete ���˻�����ϸ Where ��Դid = n_��ϸid And ��¼id = Rsdel.��¼id;
      --ɾ����Ӧ�Ĵ�ӡ����
      Begin
        Select Count(*) Into Intins From ���˻�����ϸ Where ��¼id = Rsdel.��¼id;
      Exception
        When Others Then
          Intins := 0;
      End;
      If Intins = 0 Then
        --��ȡ������ݶ�Ӧ���ļ�ID
        Begin
          Select b.Id, a.����
          Into n_�ļ�id, Intins
          From �����ļ��б� A, ���˻����ļ� B, ���˻������� C
          Where a.Id = b.��ʽid And b.Id = c.�ļ�id And c.Id = Rsdel.��¼id;
        Exception
          When Others Then
            n_�ļ�id := 0;
        End;
        Delete ���˻������� Where ID = Rsdel.��¼id;
        If Intins <> -1 Then
          Zl_���˻����ӡ_Update(n_�ļ�id, ����ʱ��_In, 1, 1);
        End If;
      End If;
    End Loop;
  Else
    --���¼�����Ŀ�Ƿ����ڸü�¼��
    Begin
      Select 1
      Into Intins
      From (Select b.��Ŀ���
             From �����ļ��ṹ A, �����¼��Ŀ B
             Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��� = ��Ŀ���_In And
                   ��id = (Select b.Id
                          From ���˻����ļ� A, �����ļ��ṹ B
                          Where a.Id = �ļ�id_In And a.��ʽid = b.�ļ�id And b.��id Is Null And b.������� = 4)
             Union
             Select ��Ŀ���
             From �����¼��Ŀ
             Where ��Ŀ���� = 2 And ��Ŀ��� = ��Ŀ���_In);
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Return;
    End If;
    If n_��¼id = 0 Then
      Select ���˻�������_Id.Nextval Into n_��¼id From Dual;

      Insert Into ���˻�������
        (ID, �ļ�id, ����ʱ��, ���汾, ������, ����ʱ��)
      Values
        (n_��¼id, �ļ�id_In, ����ʱ��_In, n_��߰汾, v_������, Sysdate);
    End If;

    --���뱾�εǼǵĲ��˻�����ϸ
    Update ���˻�����ϸ
    Set ��¼���� = ��¼����_In, ������Դ = ������Դ_In, δ��˵�� = δ��˵��_In, ��¼�� = v_������, ��¼ʱ�� = Sysdate
    Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null;
    If Sql%RowCount = 0 Then
      Select ���˻�����ϸ_Id.Nextval Into n_��ϸid From Dual;
      Insert Into ���˻�����ϸ
        (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ������, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼���, ���²�λ, ������Դ, ����, δ��˵��, ��ʼ�汾, ��ֹ�汾,
         ��¼��, ��¼ʱ��)
        Select n_��ϸid, n_��¼id, ��¼����_In, a.������, a.��Ŀid, ������_In, a.��Ŀ���, Upper(a.��Ŀ����), a.��Ŀ����, ��¼����_In, a.��Ŀ��λ, 0,
               ��¼���_In, ���²�λ_In, ������Դ_In, Nvl(b.����, 0), δ��˵��_In, n_��߰汾, Null, v_������, Sysdate
        From �����¼��Ŀ A, ���˻�����ϸ B
        Where a.��Ŀ��� = b.��Ŀ���(+) And b.��ֹ�汾(+) Is Null And b.��¼id(+) = n_��¼id And a.��Ŀ��� = ��Ŀ���_In And Rownum < 2;
    End If;
    Select ID
    Into n_��ϸid
    From ���˻�����ϸ
    Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null;
    --��д��ʷ���ݼ�ǩ����¼����ֹ�汾
    Update ���˻�����ϸ
    Set ��ֹ�汾 = n_��߰汾
    Where ��¼id = n_��¼id And ((Mod(��¼����, 10) <> 5 And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0)) Or ��¼���� = Decode(��ǩ_In, 1, 15, 5)) And ��ʼ�汾 <= n_��߰汾 - 1 And ��ֹ�汾 Is Null;

    --�����δǩ�����ݣ�����޸Ĳ���Ա��Ϊ�ü�¼�ı����˸���
    If n_��߰汾 = 1 Then
      Update ���˻������� Set ������ = v_������, ����ʱ�� = Sysdate Where ID = n_��¼id;
    End If;

    If Nvl(n_�������, 0) <> 0 Then
      Return;
    End If;

    --############
    --ͬ����������
    --############
    --1\�ȴ������µ���һ������ʼ��ֻ����һ����Ч�����µ��ļ���
    --������±������ͬ����ʱ������ݣ�ʹ������ID
    --CL,2015-12-30,��¼��ͬ��������Ŀ�����µ�
    For Row_Format In Cur_Fileformats Loop
      If Row_Format.���� = -1 Then
        If Row_Format.���� = '1' Then
          If ��Ŀ���_In = 4 Or ��Ŀ���_In = 5 Then
            Select Max(�����ı�)
            Into n_Num
            From ���˻����ļ� A, �����ļ��ṹ B
            Where a.��ʽid = b.�ļ�id And a.Id = Row_Format.�ļ�id And Ҫ������ = 'Ӥ�����µ�';
            If Not (n_Num = 1) Then
              v_Name := 'Ѫѹ';
            End If;
          End If;
          Begin
            Select 1, h.��Ŀ����
            Into Intins, n_��Ŀ����
            From (With Q2 As (Select g.��Ŀ���� As ��Ŀ����, g.��Ŀ����
                              From (Select ���
                                     From ���������Ŀ
                                     Start With ��� = (Select Max(���)
                                                      From ���������Ŀ
                                                      Where ����� Is Null
                                                      Start With ��� = ��Ŀ���_In
                                                      Connect By Prior ����� = ���)
                                     Connect By Prior ��� = �����) A, �����¼��Ŀ G
                              Where a.��� = g.��Ŀ���), Q1 As (Select To_Char(f.��¼��) As ��Ŀ����, g.��Ŀ����
                                                           From ���¼�¼��Ŀ F, �����¼��Ŀ G
                                                           Where f.��Ŀ��� = g.��Ŀ��� And g.��Ŀ���� = 2 And
                                                                 (g.���ÿ��� = 1 Or
                                                                 (g.���ÿ��� = 2 And Exists
                                                                  (Select 1
                                                                    From �������ÿ��� D
                                                                    Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id))) And
                                                                 Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And
                                                                 (Nvl(g.���ò���, 0) = 0 Or
                                                                 Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2))
                                                           Union All
                                                           Select b.Ҫ������ As ��Ŀ����, 1 As ��Ŀ����
                                                           From �����ļ��ṹ A, �����ļ��ṹ B
                                                           Where a.�ļ�id = Row_Format.��ʽid And a.��id Is Null And
                                                                 a.������� In (2, 3) And b.��id = a.Id)
                   Select *
                   From Q1
                   Union
                   Select *
                   From Q2
                   Where Exists (Select 1 From Q1, Q2 Where Q1.��Ŀ���� = Q2.��Ŀ����)) H
                   Where Instr(',' || h.��Ŀ���� || ',', ',' || v_Name || ',', 1) > 0;


          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1, h.��Ŀ����
            Into Intins, n_��Ŀ����
            From (With Q2 As (Select g.��Ŀ���, g.���ò���, g.���ÿ���, g.����ȼ�, g.��Ŀ����, g.Ӧ�÷�ʽ
                              From (Select ���
                                     From ���������Ŀ
                                     Start With ��� = (Select Max(���)
                                                      From ���������Ŀ
                                                      Where ����� Is Null
                                                      Start With ��� = ��Ŀ���_In
                                                      Connect By Prior ����� = ���)
                                     Connect By Prior ��� = �����) A, �����¼��Ŀ G
                              Where a.��� = g.��Ŀ���), Q1 As (Select g.��Ŀ���, g.���ò���, g.���ÿ���, g.����ȼ�, g.��Ŀ����, g.Ӧ�÷�ʽ
                                                           From ���¼�¼��Ŀ F, �����¼��Ŀ G
                                                           Where f.��Ŀ��� = g.��Ŀ���)
                   Select *
                   From Q1
                   Union
                   Select *
                   From Q2
                   Where Exists (Select 1 From Q1, Q2 Where Q1.��Ŀ��� = Q2.��Ŀ���)) H
                   Where Nvl(h.Ӧ�÷�ʽ, 0) <> 0 And h.����ȼ� >= 0 And
                         (Nvl(h.���ò���, 0) = 0 Or Nvl(h.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2)) And
                         h.��Ŀ��� = ��Ŀ���_In And
                         (h.���ÿ��� = 1 Or
                          (h.���ÿ��� = 2 And Exists
                           (Select 1 From �������ÿ��� D Where h.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id)));


          Exception
            When Others Then
              Intins := 0;
          End;
        End If;

        If Intins > 0 Then
          --LPF,2013-01-23,������Ŀ�Ƿ���Ҫ����ͬ��(������ǰ�Ѿ�ͬ���������ݣ�Ϊ�˱�֤��¼�������µ�����һֱ�������ݴ˺����жϡ�)
          n_Synchro := Zl_Temperatureprogram(�ļ�id_In, v_����id, ��Ŀ���_In, ����ʱ��_In);
          Begin
            Select b.Id
            Into n_Newid
            From ���˻����ļ� A, ���˻������� B
            Where a.Id = Row_Format.�ļ�id And b.�ļ�id = a.Id And b.����ʱ�� = ����ʱ��_In;
          Exception
            When Others Then
              n_Newid := 0;
          End;
          n_Oldid := n_Newid;
          If n_Newid = 0 And n_Synchro = 1 Then
            Select ���˻�������_Id.Nextval Into n_Newid From Dual;
            --�������µ�����¼
            Insert Into ���˻�������
              (ID, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
            Values
              (n_Newid, Row_Format.�ļ�id, v_������, Sysdate, ����ʱ��_In, 1);
          End If;

          Begin
            Select To_Number(��¼����_In) Into n_Num From Dual;
          Exception
            When Invalid_Number Then
              Begin
                Select 1 Into n_���� From ���¼�¼��Ŀ Where ��Ŀ��� = ��Ŀ���_In And ��¼�� = 1;
              Exception
                When Others Then
                  n_���� := 0;
              End;
              Begin
                Select 1 Into n_δ��˵�� From ��������˵�� Where ���� = ��¼����_In;
              Exception
                When Others Then
                  n_δ��˵�� := 0;
              End;
          End;

          If n_Newid > 0 Then
            --����δͬ�������µ�����(��ȻҪ���Ӷ���ѯ)
            Select Count(*)
            Into v_������Դ
            From ���˻�����ϸ
            Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                  Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��');
            If v_������Դ = 0 Then
              --˵����ͬ����ʼ�Ѿ����й����
              If n_Synchro = 1 Then
                --û�м�����Ŀ�Ƿ���Ҫͬ��
                If n_���� = 1 And n_δ��˵�� = 1 Then
                  Insert Into ���˻�����ϸ
                    (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, δ��˵��, ��ʼ�汾, ��ֹ�汾,
                     ��¼��, ��¼ʱ��, ��¼���)
                    Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, Null, b.��Ŀ��λ,
                           b.��¼���, b.���²�λ, 1, b.Id, b.��¼����, 1, Null, b.��¼��, Sysdate, 1
                    From (Select ��Ŀ���_In As ��Ŀ���, Nvl(���²�λ_In, '��') As ���²�λ
                           From Dual
                           Minus
                           Select f.��Ŀ���, Decode(Nvl(f.��Ŀ����, 1), 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��'))
                           From ���˻�����ϸ E, �����¼��Ŀ F
                           Where e.��¼id = n_Newid And e.��Ŀ��� = f.��Ŀ���) A, ���˻�����ϸ B
                    Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                  If Sql%RowCount > 0 Then
                    Int���� := 1;
                  End If;
                Else
                  Insert Into ���˻�����ϸ
                    (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, ��ʼ�汾, ��ֹ�汾, ��¼��,
                     ��¼ʱ��, ��¼���)
                    Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                           b.��¼���, b.���²�λ, 1, b.Id, 1, Null, b.��¼��, Sysdate, 1
                    From (Select ��Ŀ���_In As ��Ŀ���, Nvl(���²�λ_In, '��') As ���²�λ
                           From Dual
                           Minus
                           Select f.��Ŀ���, Decode(Nvl(f.��Ŀ����, 1), 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��'))
                           From ���˻�����ϸ E, �����¼��Ŀ F
                           Where e.��¼id = n_Newid And e.��Ŀ��� = f.��Ŀ���) A, ���˻�����ϸ B
                    Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                  If Sql%RowCount > 0 Then
                    Int���� := 1;
                  End If;
                End If;
              End If;
            Else
              If n_���� = 1 And n_δ��˵�� = 1 Then
                Update ���˻�����ϸ
                Set δ��˵�� = ��¼����_In, ��Դid = n_��ϸid, ��¼���� = Null
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                      Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ������Դ > 0;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              Else
                Update ���˻�����ϸ
                Set ��¼���� = ��¼����_In, ��Դid = n_��ϸid
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                      Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ������Դ > 0;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
        --2\��ѭ�������¼��
      Else
        If n_Mutilbill = 1 And n_Syntend = 1 Then
          --��ȡ��¼���뵱ǰ��¼�������ص����������ݵĹ̶���Ŀ
          Select Count(*)
          Into Intins
          From (Select b.��Ŀ���
                 From �����ļ��ṹ A, �����¼��Ŀ B
                 Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                       ��id =
                       (Select ID From �����ļ��ṹ Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                 Intersect
                 Select b.��Ŀ���
                 From �����ļ��ṹ A, �����¼��Ŀ B, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ G
                 Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                       b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                       a.��id = (Select ID From �����ļ��ṹ E Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4));

          If Intins > 0 Then
            n_Newid := 0;
            --����ָ���ļ��Ѿ�������ͬ����ʱ������ݣ�ֱ��������ID����
            Begin
              Select c.Id
              Into n_Newid
              From ���˻������� C
              Where c.�ļ�id = Row_Format.�ļ�id And c.����ʱ�� = ����ʱ��_In;
            Exception
              When Others Then
                n_Newid := 0;
            End;

            If n_Newid = 0 Then
              --������¼������¼
              Select ���˻�������_Id.Nextval Into n_Newid From Dual;

              Insert Into ���˻�������
                (ID, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
                Select n_Newid, Row_Format.�ļ�id, c.������, c.����ʱ��, c.����ʱ��, 1
                From ���˻������� C
                Where c.Id = n_��¼id;
            End If;

            If n_Newid > 0 Then
              --����δͬ���ļ�¼������
              Select Count(*) Into v_������Դ From ���˻�����ϸ Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In;
              If v_������Դ = 0 Then
                Insert Into ���˻�����ϸ
                  (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, δ��˵��, ��ʼ�汾, ��ֹ�汾,
                   ��¼��, ��¼ʱ��)
                  Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                         b.��¼���, b.���²�λ, 1, b.Id, b.δ��˵��, 1, Null, b.��¼��, Sysdate
                  From (Select b.��Ŀ���
                         From �����ļ��ṹ A, �����¼��Ŀ B
                         Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                               ��id = (Select ID
                                      From �����ļ��ṹ
                                      Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                         Intersect
                         Select b.��Ŀ���
                         From �����ļ��ṹ A, �����¼��Ŀ B, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ G
                         Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                               b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                               a.��id =
                               (Select ID From �����ļ��ṹ E Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4)) A, ���˻�����ϸ B
                  Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                  --ԭ������Ҫ��
                  Begin
                    Select ���� Into n_���� From ���˻����ӡ Where �ļ�id = Row_Format.�ļ�id And ��¼id = n_Newid;
                  Exception
                    When Others Then
                      n_���� := 1;
                  End;
                  Zl_���˻����ӡ_Update(Row_Format.�ļ�id, ����ʱ��_In, n_����, 0);
                End If;
              Else
                Update ���˻�����ϸ
                Set ��¼���� = ��¼����_In, δ��˵�� = δ��˵��_In, ��Դid = n_��ϸid
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And ������Դ > 0;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End Loop;

    If Int���� = 1 Then
      Update ���˻�����ϸ Set ���� = 1 Where ID = n_��ϸid;
      --����ʷ���ݵĹ��ñ�־����ΪNULL
      Update ���˻�����ϸ Set ���� = Null Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And ID <> n_��ϸid;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˻�������_Update;
/

--122609:������,2018-03-08,����ƽ̨��Ϣ���
Create Or Replace Package b_Message Is
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  );
  --����ƽ̨��������
  Procedure Set_Platform_Call(Platform_Call Number);
  --��������
  Procedure Zlhis_Dict_001(Id_In ���ű�.Id%Type);
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In ���ű�.Id%Type);
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In ���ű�.Id%Type);
  --���ò���
  Procedure Zlhis_Dict_004(����id_In ���ű�.Id%Type);
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In ��Ա��.Id%Type);
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In ��Ա��.Id%Type);
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In ��Ա��.Id%Type);
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In ��Ա��.Id%Type);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In �շ���ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In ������ĿĿ¼.Id%Type);
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In ������ĿĿ¼.Id%Type);
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In ������ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In ������ĿĿ¼.Id%Type);
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In ������ĿĿ¼.Id%Type);
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In ������ĿĿ¼.Id%Type);
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In ������ĿĿ¼.Id%Type,
    ����_In   ����������Ŀ.����%Type,
    ������_In ����������Ŀ.������%Type,
    Ӣ����_In ����������Ŀ.Ӣ����%Type
  );

  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In ��������Ŀ¼.Id%Type);
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In ��������Ŀ¼.Id%Type);
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In ��������Ŀ¼.Id%Type);
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In ��������Ŀ¼.Id%Type);
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  );
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type,
    ����_In ������ĿĿ¼.����%Type,
    ����_In ������ĿĿ¼.����%Type
  );
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ���_In   �շ���ĿĿ¼.���%Type,
    ����_In   �շ���ĿĿ¼.����%Type
  );
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  );
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In ������ĿĿ¼.Id%Type);
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In �շ���ĿĿ¼.Id%Type);
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In �շ���ĿĿ¼.Id%Type);
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
    Id_In   �շ���ĿĿ¼.Id%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ���_In �շ���ĿĿ¼.���%Type,
    ����_In �շ���ĿĿ¼.����%Type
  );
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In �շ���ĿĿ¼.Id%Type);
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In �շ���ĿĿ¼.Id%Type);
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type
  );
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type
  );
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  );
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  );
  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In ҩƷ�շ���¼.No%Type);
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In ҩƷ�շ���¼.No%Type);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In ҩƷ�շ���¼.No%Type);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004(No_In ҩƷ�շ���¼.No%Type);
  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In ҩƷ�շ���¼.�ⷿid%Type,
    �շ�id_In ҩƷ�շ���¼.Id%Type
  );
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    ����_In       ҩƷ�շ���¼.ʵ������%Type,
    ����id_In     ������ü�¼.Id%Type
  );
  --ҩƷ����
  Procedure Zlhis_Drug_007(�۸�id_In ҩƷ�۸��¼.Id%Type);
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2);
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In ҩƷ�۸��¼.Id%Type,
    ʱ��_In   Number
  );
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In �ɱ��۵�����Ϣ.Id%Type);
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In �շѼ�Ŀ.Id%Type,
    ʱ��_In   Number
  );
  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In ����ҽ����¼.����id%Type,
    ��ҳid_In  In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In  In ����ҽ����¼.Id%Type,
    ҽ��ids_In In Varchar2
  );
  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --7.�������߻�����ҽ����סԺ  
  Procedure Zlhis_Cis_007
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In ������ϼ�¼.Id%Type
  );
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In ���˹Һż�¼.����id%Type,
    ����id_In   In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In ������ϼ�¼.Id%Type,
    ����id_In   In ������ϼ�¼.����id%Type,
    ���id_In   In ������ϼ�¼.���id%Type,
    �������_In In ������ϼ�¼.�������%Type
  );

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  );

  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  );

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type
  );

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  );

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ��������_In In ����ҽ������.��������%Type,
    �״�ʱ��_In In ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_In In ����ҽ������.ĩ��ʱ��%Type,
    ��������_In In ����ҽ������.��������%Type
  );
  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type
  );

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type,
    ��������_In In ����ҽ��ִ��.��������%Type,
    ִ�н��_In In ����ҽ��ִ��.ִ�н��%Type,
    ִ��ժҪ_In In ����ҽ��ִ��.ִ��ժҪ%Type,
    ִ�п���_In In ����ҽ��ִ��.ִ�п���id%Type,
    ִ����_In   In ����ҽ��ִ��.ִ����%Type,
    �˶���_In   In ����ҽ��ִ��.�˶���%Type,
    ��¼��Դ_In In ����ҽ��ִ��.��¼��Դ%Type
  );
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  );

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  );
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  );
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Ӱ�����¼.ҽ��id%Type);
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In In Ris���ԤԼ.ԤԼid%Type
  );
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In       In Ris���ԤԼ.ԤԼid%Type,
    ԤԼ����_In     In Ris���ԤԼ.ԤԼ����%Type,
    ԤԼ���_In     In Ris���ԤԼ.���%Type,
    ����豸����_In In Ris���ԤԼ.����豸����%Type
  );

  --36.���߷���
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  );

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  );

  --38.�����˿�
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ԭ����_In   In ����ҽ�ƿ���Ϣ.����%Type,
    �¿���_In   In ����ҽ�ƿ���Ϣ.����%Type
  );

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  );

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type,
    ����_In   In ���˹Һż�¼.����%Type
  );

  --41.�����˺�
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  );

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In �ٴ������¼.Id%Type,
    �䶯id_In   In �ٴ�����䶯��¼.Id%Type
  );

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In ���˹Һż�¼.No%Type,
    �䶯ԭ��_In   Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In ����䶯��¼.Id%Type
  );

  --���������շѼ��������
  --��������_In:1-�շѽ��㣬2-�������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  );

  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  );

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In In ����Ԥ����¼.No%Type
  );

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In   In ����Ԥ����¼.No%Type
  );

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type
  );

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  );

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    ������ʽ_In In Varchar2
  );
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  );
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  );
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In ������ҳ.����id%Type);
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In ������ҳ.����id%Type);

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017
  (
    ����id_In   In ������ҳ.����id%Type,
    ԭ����id_In In ������ҳ.����id%Type
  );

  --69.����ת����ת��
  Procedure Zlhis_Patient_026
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  );

  Procedure Zlhis_Patient_028(����id_In In ������ҳ.����id%Type);

  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In ����ҽ����¼.Id%Type);
  --Ѫ��:������Ѫ�ܾ�
  Procedure Zlhis_Blood_002(ҽ��id_In In ����ҽ����¼.Id%Type);

  --70.����걾���
  Procedure Zlhis_Lis_001(�걾id_In In ����걾��¼.Id%Type);
  --71.����걾��˳���
  Procedure Zlhis_Lis_002(�걾id_In In ����걾��¼.Id%Type);
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  );
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  );
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In ����걾��¼.Id%Type);
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In ����걾��¼.Id%Type);
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In ����걾��¼.Id%Type);
End b_Message;
/

--122609:������,2018-03-08,����ƽ̨��Ϣ���
Create Or Replace Package Body b_Message Is
  --�Ƿ���ƽ̨����
  Is_Platform_Call Number(1) := 0;
  --��Ϣ��������
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --������Ϣ��ѯ���
  Type Tmap_Msg_Using Is Table Of Number(1) Index By Varchar2(30);
  Zlmsg_Map Tmap_Msg_Using;
  --��Ϣ�Ƿ�����
  Function p_Msg_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    Begin
      n_Using := Zlmsg_Map(v_Code);
      Return n_Using;
    Exception
      When No_Data_Found Then
        --����ȡMax�ݴ��������൱�����,�û�����û�в�ȡͬ���޸Ļ��Լ���������Ϣ���͵���δע�ᵽZlmsg_Lists���������������ִ���


      
        Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
        Zlmsg_Map(v_Code) := n_Using;
        --��ѯ������Ϣ����Ա�������������ִ�д���
        If Message_Creator Is Null Then
          Message_Creator := Zl_Username;
        End If;
        Return n_Using;
    End;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || 'δ��Zlmsg_Lists���ҵ���Ϣ"' || v_Code || '"������ϵ����Ա���д���' || '[ZLSOFT]');
      Return 0;
  End;
  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If p_Msg_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --���õ�ǰ�ỰΪƽ̨����
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --��ϢZlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In ���ű�.Id%Type) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_DICT_001') = 0 Then
      Return;
    End If;
    Begin
      Select Xmltype(Key_Define) Into v_Define From Zlmsg_Lists Where Code = 'ZLHIS_DICT_001';
    Exception
      When Others Then
        v_Define := Xmltype('<root><ID>NULL</ID></root>');
    End;
    Select Updatexml(v_Define, '/root/ID/text()', Id_In).Getstringval() Into v_Value From Dual;
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --���ò���
  Procedure Zlhis_Dict_004(����id_In ���ű�.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In ��Ա��.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In ������ĿĿ¼.Id%Type,
    ����_In   ����������Ŀ.����%Type,
    ������_In ����������Ŀ.������%Type,
    Ӣ����_In ����������Ŀ.Ӣ����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID>' || '<����>' || ����_In || '</����>' || '<������>' || ������_In || '</������>' ||
               '<Ӣ����>' || Ӣ����_In || '</Ӣ����>' || '<ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In ��������Ŀ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type,
    ����_In ������ĿĿ¼.����%Type,
    ����_In ������ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In ������ĿĿ¼.���%Type,
    Id_In   ������ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ����_In   �շ���ĿĿ¼.����%Type,
    ���_In   �շ���ĿĿ¼.���%Type,
    ����_In   �շ���ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><���>' || ���_In || '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   �շ���ĿĿ¼.���%Type,
    ҩƷid_In �շ���ĿĿ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In ������ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
    Id_In   �շ���ĿĿ¼.Id%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ����_In �շ���ĿĿ¼.����%Type,
    ���_In �շ���ĿĿ¼.���%Type,
    ����_In �շ���ĿĿ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID><����>' || ����_In || '</����><����>' || ����_In || '</����><���>' || ���_In ||
               '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In �շ���ĿĿ¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In ����֧����Ŀ.����%Type,
    �շ�ϸĿid_In In ����֧����Ŀ.�շ�ϸĿid%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ��Ŀ����_In   In �շ���ĿĿ¼.����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type,
    ҽ������_In   In ����֧����Ŀ.��Ŀ����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID><��Ŀ����>' || ��Ŀ����_In || '</��Ŀ����><��Ŀ����>' ||
               ��Ŀ����_In || '</��Ŀ����><ҽ������>' || ҽ������_In || '</ҽ������><ҽ������>' || ҽ������_In || '</ҽ������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End Zlhis_Dict_051;
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In ���Ʒ���Ŀ¼.����%Type,
    Id_In   ���Ʒ���Ŀ¼.Id%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type,
    ����_In ���Ʒ���Ŀ¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End Zlhis_Dict_052;
  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004(No_In ҩƷ�շ���¼.No%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In ҩƷ�շ���¼.�ⷿid%Type,
    �շ�id_In ҩƷ�շ���¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�ⷿID>' || �ⷿid_In || '</�ⷿID><�շ�ID>' || �շ�id_In || '</�շ�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    �����շ�id_In ҩƷ�շ���¼.Id%Type,
    ����_In       ҩƷ�շ���¼.ʵ������%Type,
    ����id_In     ������ü�¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������¼ID>' || �����շ�id_In || '</������¼ID><������¼ID>' || �����շ�id_In || '</������¼ID><����>' || ����_In ||
               '</����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --ҩƷ����
  Procedure Zlhis_Drug_007(�۸�id_In ҩƷ�۸��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End Zlhis_Drug_007;
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2) Is
    v_Value  Zlmsg_Todo.Key_Value%Type;
    n_��¼id ��Һ��ҩ��¼.Id%Type;
    v_Tmp    Varchar2(4000);
  Begin
    If ��¼ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := ��¼ids_In || ',';
    End If;
  
    v_Value := '<root><��¼IDS>';
  
    While v_Tmp Is Not Null Loop
      --�ֽⵥ��ID��
      n_��¼id := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_��¼id || ',');
    
      v_Value := v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>';
    End Loop;
  
    v_Value := v_Value || '</��¼IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End Zlhis_Drug_008;
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In ҩƷ�۸��¼.Id%Type,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End Zlhis_Drug_009;
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In �ɱ��۵�����Ϣ.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End Zlhis_Drug_010;
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In �շѼ�Ŀ.Id%Type,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End Zlhis_Drug_011;

  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In ����ҽ����¼.����id%Type,
    ��ҳid_In  In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In  In ����ҽ����¼.Id%Type,
    ҽ��ids_In In Varchar2
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                   '</ID></root>');
    Else
      For R In (Select '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ID || '</ID></root>' As Xml_Value
                From ����ҽ����¼
                Where ID In (Select Column_Value From Table(f_Num2list(ҽ��ids_In))) And ���id Is Null) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Cis_002;
  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.�������߻�����ҽ����סԺ  
  Procedure Zlhis_Cis_007
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In ������ϼ�¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || ���id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In ���˹Һż�¼.����id%Type,
    ����id_In   In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In ������ϼ�¼.Id%Type,
    ����id_In   In ������ϼ�¼.����id%Type,
    ���id_In   In ������ϼ�¼.���id%Type,
    �������_In In ������ϼ�¼.�������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || Id_In || '</ID><����ID>' ||
               ����id_In || '</����ID><���ID>' || ���id_In || '</���ID><�������>' || �������_In || '</�������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    If v_�������� = '����' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In ���˹Һż�¼.����id%Type,
    ����id_In In ���˹Һż�¼.Id%Type, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In ����ҽ����¼.Id%Type,
    ��Ϣid_In In ҵ����Ϣ�嵥.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ������Դ_In In ����ҽ����¼.������Դ%Type --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� ������ĿĿ¼.��������%Type;
  Begin
    Select Max(a.��������)
    Into v_��������
    From ������ĿĿ¼ A, ����ҽ����¼ B
    Where b.������Ŀid = a.Id And b.Id = ҽ��id_In;
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    If v_�������� = '����' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type,
    No_In     In ����ҽ������.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    No_In       In ����ҽ������.No%Type,
    ��������_In In ����ҽ������.��������%Type,
    �״�ʱ��_In In ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_In In ����ҽ������.ĩ��ʱ��%Type,
    ��������_In In ����ҽ������.��������%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO><��������>' || ��������_In || '</��������><�״�ʱ��>' ||
               To_Char(�״�ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</�״�ʱ��><ĩ��ʱ��>' ||
               To_Char(ĩ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ĩ��ʱ��><��������>' || ��������_In || '</��������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In ����ҽ����¼.����id%Type,
    ��ҳid_In   In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In   In ���˹Һż�¼.No%Type,
    ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
    ҽ��id_In   In ����ҽ����¼.Id%Type,
    Ҫ��ʱ��_In In ����ҽ��ִ��.Ҫ��ʱ��%Type,
    ִ��ʱ��_In In ����ҽ��ִ��.ִ��ʱ��%Type,
    ��������_In In ����ҽ��ִ��.��������%Type,
    ִ�н��_In In ����ҽ��ִ��.ִ�н��%Type,
    ִ��ժҪ_In In ����ҽ��ִ��.ִ��ժҪ%Type,
    ִ�п���_In In ����ҽ��ִ��.ִ�п���id%Type,
    ִ����_In   In ����ҽ��ִ��.ִ����%Type,
    �˶���_In   In ����ҽ��ִ��.�˶���%Type,
    ��¼��Դ_In In ����ҽ��ִ��.��¼��Դ%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��><��������>' || ��������_In ||
               '</��������><ִ�н��>' || ִ�н��_In || '</ִ�н��><ִ��ժҪ>' || ִ��ժҪ_In || '</ִ��ժҪ><ִ�п���ID>' || ִ�п���_In ||
               '</ִ�п���ID><ִ����>' || ִ����_In || '</ִ����><�˶���>' || �˶���_In || '</�˶���><��¼��Դ>' || ��¼��Դ_In || '</��¼��Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In ����ҽ����¼.����id%Type,
    ��ҳid_In In ����ҽ����¼.��ҳid%Type,
    �Һŵ�_In In ���˹Һż�¼.No%Type,
    ���ͺ�_In In ����ҽ������.���ͺ�%Type,
    ҽ��id_In In ����ҽ����¼.Id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '<��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԭ״̬_In In ����ҽ������.ִ�й���%Type,
    ��״̬_In In ����ҽ������.ִ�й���%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Ӱ�����¼.ҽ��id%Type,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If p_Msg_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '<��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Ӱ�����¼.ҽ��id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In In Ris���ԤԼ.ԤԼid%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Ӱ�����¼.ҽ��id%Type,
    ԤԼid_In       In Ris���ԤԼ.ԤԼid%Type,
    ԤԼ����_In     In Ris���ԤԼ.ԤԼ����%Type,
    ԤԼ���_In     In Ris���ԤԼ.���%Type,
    ����豸����_In In Ris���ԤԼ.����豸����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID><ԤԼ����>' || ԤԼ����_In || '</ԤԼ����><ԤԼ���>' ||
               ԤԼ���_In || '</ԤԼ���><����豸����>' || ����豸����_In || '</����豸����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;

  --36.���߷�����󶨿�
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ����_In     In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.���߲���/����
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In ���˱䶯��¼.Id%Type,
    ����id_In   In ������Ϣ.����id%Type,
    �����id_In In ҽ�ƿ����.Id%Type,
    ԭ����_In   In ����ҽ�ƿ���Ϣ.����%Type,
    �¿���_In   In ����ҽ�ƿ���Ϣ.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><ԭ����>' || ԭ����_In || '</ԭ����><�¿���>' || �¿���_In || '</�¿���></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type,
    ����_In   In ���˹Һż�¼.����%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO><����>' || Nvl(����_In, '') || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.�����˺ţ���ȡ��ԤԼ)
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In ���˹Һż�¼.Id%Type,
    No_In     In ���˹Һż�¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In �ٴ������¼.Id%Type,
    �䶯id_In   In �ٴ�����䶯��¼.Id%Type
    
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><��¼ID>' || ��¼id_In || '</��¼ID><�䶯ID>' || �䶯id_In ||
               '</�䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In ���˹Һż�¼.No%Type,
    �䶯ԭ��_In   Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In ����䶯��¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><����䶯ID>' || ����䶯id_In ||
               '</����䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;

  --���������շѼ��������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --��������_In:1-�շѽ��㣬2-�������
    v_Value := '<root><��������>' || ��������_In || '</��������><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;

  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In ������ü�¼.����id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --�˷�����_In:1-�շѽ��㣬2-�������
    v_Value := '<root><�˷�����>' || �˷�����_In || '</�˷�����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In In ����Ԥ����¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><Ԥ��ID>' || Ԥ��id_In || '</Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In ����Ԥ����¼.Id%Type,
    ���ݺ�_In   In ����Ԥ����¼.No%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��Ԥ��ID>' || ��Ԥ��id_In || '</��Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In סԺ���ü�¼.�շ����%Type,
    ����id_In   In סԺ���ü�¼.Id%Type,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_�շ�id  Varchar2(50);
    v_����    Varchar2(20);
  Begin
    If p_Msg_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID>';
  
    If �շ�ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<�շ�IDS>' || '<�շ�ID>' || '</�շ�ID>' || '<����>' || '</����>' || '</�շ�IDS>';
    Else
      v_Infotmp := �շ�ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --�ֽ��շ�ID��
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_�շ�id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_����    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        v_Tmp := v_Tmp || '<�շ�IDS>' || '<�շ�ID>' || v_�շ�id || '</�շ�ID>' || '<����>' || v_���� || '</����>' || '</�շ�IDS>';
      End Loop;
    End If;
  
    v_Value := v_Value || v_Tmp || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_001;
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_002;
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭ����   Varchar2(255);
    v_�´���   Varchar2(255);
    n_�䶯id   Number(18);
    n_��ʼԭ�� Number(3);
    d_��ʼʱ�� Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
    Select ID, ����, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_�´���, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����)
    Into v_ԭ����
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����>' ||
                                 v_ԭ���� || '</ԭ����>' || '<�´���>' || v_�´��� || '</�´���>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_005;
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    ������ʽ_In In Varchar2
  ) Is
    n_����id     ���˱䶯��¼.����id%Type;
    n_����id     ���˱䶯��¼.����id%Type;
    n_����ȼ�id ���˱䶯��¼.����ȼ�id%Type;
    n_ҽ��С��id ���˱䶯��¼.ҽ��С��id%Type;
    v_����       ���˱䶯��¼.����%Type;
    v_���λ�ʿ   ���˱䶯��¼.���λ�ʿ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����ҽʦ   ���˱䶯��¼.����ҽʦ%Type;
    v_����       ���˱䶯��¼.����%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
    Select Max(����id), Max(����id), Max(����ȼ�id), Max(ҽ��С��id), Max(����), Max(���λ�ʿ), Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(����)
    Into n_����id, n_����id, n_����ȼ�id, n_ҽ��С��id, v_����, v_���λ�ʿ, v_����ҽʦ, v_����ҽʦ, v_����ҽʦ, v_����
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><������ʽ>' || ������ʽ_In ||
                                 '</������ʽ><����ID>' || n_����id || '</����ID>' || '<����ID>' || n_����id || '</����ID>' || '<����ȼ�ID>' ||
                                 n_����ȼ�id || '</����ȼ�ID>' || '<ҽ��С��ID>' || n_ҽ��С��id || '</ҽ��С��ID>' || '<����>' || v_���� ||
                                 '</����>' || '<���λ�ʿ>' || v_���λ�ʿ || '</���λ�ʿ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' ||
                                 '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����>' || v_���� ||
                                 '</����>' || '</root>');
  End Zlhis_Patient_006;
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭסԺҽ�� Varchar2(100);
    v_��סԺҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ���λ�ʿ Varchar2(100);
    v_�����λ�ʿ Varchar2(100);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
    Select ID, ����ҽʦ, ����ҽʦ, ����ҽʦ, ���λ�ʿ, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_��סԺҽ��, v_������ҽ��, v_������ҽ��, v_�����λ�ʿ, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(���λ�ʿ)
    Into v_ԭסԺҽ��, v_ԭ����ҽ��, v_ԭ����ҽ��, v_ԭ���λ�ʿ
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭסԺҽ��>' ||
                                 v_ԭסԺҽ�� || '</ԭסԺҽ��>' || '<��סԺҽ��>' || v_��סԺҽ�� || '</��סԺҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� ||
                                 '</ԭ����ҽ��>' || '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� || '</ԭ����ҽ��>' ||
                                 '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ���λ�ʿ>' || v_ԭ���λ�ʿ || '</ԭ���λ�ʿ>' || '<�����λ�ʿ>' ||
                                 v_�����λ�ʿ || '</�����λ�ʿ>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_007;
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ԭ����ȼ�id Number(18);
    v_�»���ȼ�id Number(18);
    n_�䶯id       Number(18);
    n_��ʼԭ��     Number(3);
    d_��ʼʱ��     Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
    Select ID, ����ȼ�id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_�»���ȼ�id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����ȼ�id)
    Into v_ԭ����ȼ�id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����ȼ�ID>' ||
                                 v_ԭ����ȼ�id || '</ԭ����ȼ�ID>' || '<�»���ȼ�ID>' || v_�»���ȼ�id || '</�»���ȼ�ID>' || '<�䶯ID>' ||
                                 n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    n_�䶯id ���˱䶯��¼.Id%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
    Select Max(ID)
    Into n_�䶯id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_009;
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID></root>');
  End Zlhis_Patient_010;
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_011;
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ת������id Number(18);
    v_ת�����id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
    Select ID, ����id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_ת�����id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����id)
    Into v_ת������id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�����ID>' || v_ת�����id || '</ת�����ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In ������ҳ.����id%Type,
    ��ҳid_In   In ������ҳ.��ҳid%Type,
    Ӥ�����_In ����ҽ����¼.Ӥ��%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_013;
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In ������ҳ.����id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_015;
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In ������ҳ.����id%Type) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_016;

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017
  (
    ����id_In   In ������ҳ.����id%Type,
    ԭ����id_In In ������ҳ.����id%Type
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017',
                                '<root><����ID>' || ����id_In || '</����ID><ԭ����ID>' || ԭ����id_In || '</ԭ����ID></root>');
  End Zlhis_Patient_017;

  --69.סԺ����ת�벡��
  Procedure Zlhis_Patient_026
  (
    ����id_In In ������ҳ.����id%Type,
    ��ҳid_In In ������ҳ.��ҳid%Type
  ) Is
    v_ת������id Number(18);
    v_ת�벡��id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
    Select ID, ����id, ��ʼʱ��, ��ʼԭ��
    Into n_�䶯id, v_ת�벡��id, d_��ʼʱ��, n_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Select Max(����id)
    Into v_ת������id
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = d_��ʼʱ�� And ��ֹԭ�� = n_��ʼԭ�� And Nvl(���Ӵ�λ, 0) = 0;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�벡��ID>' || v_ת�벡��id || '</ת�벡��ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(����id_In In ������ҳ.����id%Type) Is
    v_����     ������Ϣ.����%Type;
    v_�Ա�     ������Ϣ.�Ա�%Type;
    v_����     ������Ϣ.����%Type;
    v_�������� ������Ϣ.��������%Type;
    v_�����   ������Ϣ.�����%Type;
    v_���֤�� ������Ϣ.���֤��%Type;
  Begin
    If p_Msg_Using('ZLHIS_PATIENT_028') = 0 Then
      Return;
    End If;
    Select ����, �Ա�, ����, ��������, �����, ���֤��
    Into v_����, v_�Ա�, v_����, v_��������, v_�����, v_���֤��
    From ������Ϣ
    Where ����id = ����id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028',
                                '<root><����ID>' || ����id_In || '</����ID><����>' || v_���� || '</����>' || '<�Ա�>' || v_�Ա� ||
                                 '</�Ա�>' || '<����>' || v_���� || '</����>' || '<��������>' || v_�������� || '</��������>' || '<�����>' ||
                                 v_����� || '</�����>' || '<���֤��>' || v_���֤�� || '</���֤��>' || '</root>');
  End Zlhis_Patient_028;

  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In ����ҽ����¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --Ѫ��:���Ҿܾ���Ѫ
  Procedure Zlhis_Blood_002(ҽ��id_In In ����ҽ����¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.���鱨�����
  Procedure Zlhis_Lis_001(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.���鱨����˳���
  Procedure Zlhis_Lis_002(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      For R In (Select '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>' As Xml_Value
                From ����ҽ������
                Where ҽ��id In (Select Column_Value From Table(f_Num2list(ҽ��ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In ����ҽ������.��������%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If p_Msg_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      For R In (Select '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>' As Xml_Value
                From ����ҽ������
                Where ҽ��id In (Select Column_Value From Table(f_Num2list(ҽ��ids_In)))) Loop
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', r.Xml_Value);
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In ����걾��¼.Id%Type) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

End b_Message;
/

--122609:������,2018-03-08,����ƽ̨��Ϣ���
Create Or Replace Procedure Zl_����ҽ����¼_����
(
  ҽ��id_In     In ����ҽ����¼.Id%Type,
  Flag_In       In Number := 0,
  ҽ������_In   In ����ҽ����¼.ҽ������%Type := Null,
  ��������_In   In ����ҽ��״̬.��������%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null
  --���ܣ�����סԺҽ����״̬�������Ͳ���(������������ͨ������Zl_����ҽ����¼_��������������)
  --������ҽ��ID_IN=һ��ҽ��ID
  --      FLAG_IN=�������ݡ�����ֹͣ��0=���ִ����ֹʱ��,1=�������е�ִ����ֹʱ�䡣
  --      ҽ������_IN=�ù��̱��������˵���ʱ���ã����ڴ�����ʾ��
  --      ��������_IN=�ù��̱��������˵���ʱ���ã����ں˶Ի������ݡ�0-���˷���,n=���˾���ҽ������
) Is
  --����ָ��ҽ���Ĳ�����¼,��һ��ΪҪ���˵�����(״̬��������)
  --���������˷��ͺ���Զ�ֹͣ,�ڻ��˷���ʱ�Զ�����ֹͣ����
  Cursor c_Rolladvice Is
    Select b.������Ա, b.����ʱ��, 0 As ���ͺ�, Null As NO, b.��������, 0 As ִ��״̬, Sysdate + Null As �״�ʱ��, Sysdate + Null As ĩ��ʱ��,
           a.�ϴ�ִ��ʱ��, a.ҽ����Ч, a.������� As ���, a.������Ŀid, Null As ����, a.����id, a.��ҳid, a.Ӥ��, 0 As ��¼����, 0 As �������, 0 As ��������id,
           a.��˱��, a.����ҽ��, a.ִ�п���id, Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, -null As ��������, Null As ��������
    From ����ҽ����¼ A, ����ҽ��״̬ B
    Where a.Id = b.ҽ��id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
          (Nvl(a.ҽ����Ч, 0) = 0 And b.�������� Not In (1, 2, 3) Or Nvl(a.ҽ����Ч, 0) = 1 And b.�������� Not In (1, 2, 3, 8))
    Union
    Select b.������ As ������Ա, b.����ʱ�� As ����ʱ��, b.���ͺ�, b.No, -null As ��������, b.ִ��״̬, b.�״�ʱ��, b.ĩ��ʱ��, a.�ϴ�ִ��ʱ��, a.ҽ����Ч, c.���,
           a.������Ŀid, c.�������� As ����, a.����id, a.��ҳid, a.Ӥ��, b.��¼����, b.�������, a.��������id, a.��˱��, a.����ҽ��, a.ִ�п���id,
           Nvl(a.���id, a.Id) As ��id, a.���id, a.Id As ҽ��id, b.��������, b.��������
    From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
    Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By ����ʱ�� Desc, ���ͺ�;
  r_Rolladvice c_Rolladvice%RowType;

  --��ʽͬc_Rolladvice��ֻȡ���Ͳ��������Զ����˴���
  Cursor c_Rollsend(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Distinct b.ҽ��id, b.����ʱ�� As ����ʱ��, b.���ͺ�, b.ִ��״̬, a.������� As ���, c.��ǰ����id As ���˲���id, a.���˿���id,
                    b.ִ�в���id As ִ�п���id
    From ����ҽ����¼ A, ����ҽ������ B, ������ҳ C
    Where a.Id = b.ҽ��id And b.���ͺ� = v_���ͺ� And a.����id = c.����id And a.��ҳid = c.��ҳid And
          (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By b.����ʱ�� Desc, b.���ͺ�;

  --����ҽ��������NO������λ���Ҫ���ʵķ��ü�¼
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,�ҿ���NO��ͬ(ҩƷ��,�÷��巨��һ����)
  --���ܷ��ͼ�¼�ļƷ�״̬(��������Ʒ�),�з��ü�¼��Ȼ��������
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ�������
  --ֻ�ܼ�¼״̬Ϊ1�ķ���,���������ʻ򲿷����ʵļ�¼,���ٴ�������"��¼״̬=3"�Ķ�ȡ�����������жϣ�������
  Cursor c_Rollmoneyout
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From ������ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  Cursor c_Rollmoneyin
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From סԺ���ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  --ȡ����סԺ����ʱ�Զ����ŵ�����(��û�����ϵ�)
  Cursor c_Stuff_Drug(v_����id ҩƷ�շ���¼.����id%Type) Is
    Select ID
    From ҩƷ�շ���¼
    Where ����id = v_����id And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0) And ����� Is Not Null
    Order By ҩƷid;

  --���ڴ�������ҽ���Ļ���
  Cursor c_Patilog
  (
    v_����id ���˱䶯��¼.����id%Type,
    v_��ҳid ���˱䶯��¼.��ҳid%Type
  ) Is
    Select *
    From ���˱䶯��¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null
    Order By ��ʼʱ�� Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
  t_Adviceids t_Numlist;

  v_ҽ��״̬     ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ����Ч     ����ҽ����¼.ҽ����Ч%Type;
  v_����no       ����ҽ������.No%Type;
  v_�������     Varchar2(255);
  v_ĩ��ʱ��     ����ҽ������.ĩ��ʱ��%Type;
  v_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_��������     ������ĿĿ¼.��������%Type;
  v_ִ��Ƶ��     ������ĿĿ¼.ִ��Ƶ��%Type;
  v_�ϴ�ʱ��     ����ҽ����¼.�ϴ�ִ��ʱ��%Type;
  v_ִ��ʱ��     ����ҽ����¼.ִ��ʱ�䷽��%Type;
  v_��ʼִ��ʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_�ϴδ�ӡʱ�� ����ҽ����¼.�ϴδ�ӡʱ��%Type;
  v_Ƶ�ʼ��     ����ҽ����¼.Ƶ�ʼ��%Type;
  v_�����λ     ����ҽ����¼.�����λ%Type;
  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  n_����ȼ�id   ���˱䶯��¼.����ȼ�id%Type;
  d_��ʼʱ��     ���˱䶯��¼.��ʼʱ��%Type;
  d_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_Tmp���ͺ�    ����ҽ������.���ͺ�%Type;
  n_ִ��         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       סԺ���ü�¼.No%Type,
    v_��¼���� סԺ���ü�¼.��¼����%Type,
    v_���     סԺ���ü�¼.���%Type,
    n_����     Number := 0 --0סԺ��1����
  ) Return Number Is
    n_Num      Number;
    n_ִ��״̬ Number;
  Begin
    n_Num := 0;
    If n_���� = 0 Then
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    Else
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --�������¼����ִ�У�����ִ�еģ����Զ��ˡ�
    If n_ִ��״̬ <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp���ͺ� := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰû�п��Ի��˵����ݡ�';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_����ʱ�� <> r_Rolladvice.����ʱ�� And d_����ʱ�� Is Not Null;
    d_����ʱ�� := r_Rolladvice.����ʱ��;
  
    --�������˵���ʱ�ж�
    If ҽ������_In Is Not Null Then
      If Nvl(r_Rolladvice.��������, 0) <> Nvl(��������_In, 0) Then
        v_Error := Nvl(ҽ������_In, '��ҽ��') || '�����뵱ǰҽ��һ����ˣ����ܸ�ҽ���Ѿ�ִ��������������';
        Raise Err_Custom;
      End If;
    End If;
  
    --һ�鷢�ͺ�ִֻ��һ��
    If v_Tmp���ͺ� <> r_Rolladvice.���ͺ� Then
      v_Tmp���ͺ� := r_Rolladvice.���ͺ�;
      n_ִ��      := 1;
    Else
      n_ִ�� := 0;
    End If;
  
    If n_ִ�� = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.���ͺ� = 0 Then
        --����ҽ��״̬����(��ʱ��ؼ���)
        --4-���ϣ�5-������6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ��;13-ͣ������
        ------------------------------------------------------------------
        --���ֻ���˻ص�У��״̬
        If r_Rolladvice.�������� = 3 Then
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰ����ͨ��У��״̬�������ٻ��ˡ�';
          Raise Err_Custom;
        Elsif r_Rolladvice.�������� = 4 And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          If r_Rolladvice.��� = 'H' Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              v_Error := '����ȼ����Ϻ����ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --����Ƿ�������������֮ǰ�Ĳ���
        If r_Rolladvice.�������� <> 5 Then
          --ȡ�������ʱ��
          Select Nvl(ҽ������ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_����ʱ��
          From ������ҳ
          Where ����id = r_Rolladvice.����id And ��ҳid = r_Rolladvice.��ҳid;
        
          If r_Rolladvice.����ʱ�� < v_����ʱ�� Then
            v_Error := '�ò������������֮ǰ�Ĳ��������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        --ɾ��(����ҽ��)�����״̬������¼
        Delete /*+ Rule*/
        From ����ҽ��״̬
        Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ����ʱ�� = r_Rolladvice.����ʱ��;
      
        --ȡɾ����Ӧ�ָ���ҽ��״̬
        Select ��������
        Into v_ҽ��״̬
        From ����ҽ��״̬
        Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
      
        --�ָ�(����ҽ��)���˺��״̬
        Update ����ҽ����¼ Set ҽ��״̬ = v_ҽ��״̬ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --��������Ĵ���
        If r_Rolladvice.�������� = 8 Then
          --�����ڷ����ջع���ҽ�� ���������������ģʽ�����ж϶�Ӧ�ġ����˷������ʡ������Ƿ�ȡ��������������ˣ���������
          --                       ����ǲ�����������ģʽ���������ٻ��ˡ�
          --���ܳ��ڷ����ջ�ʱ��ȫ���ջ�(���ϴ�ִ��ʱ��)
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From ����ҽ����¼ A, ����ҽ������ B
          Where b.ҽ��id = a.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
                b.���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                a.ִ����ֹʱ�� Is Not Null And ((a.�ϴ�ִ��ʱ�� < b.ĩ��ʱ��) Or (a.�ϴ�ִ��ʱ�� Is Null And b.ĩ��ʱ�� Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('�����ջز�����������', 1254) = '1' Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
              Raise Err_Custom;
            Else
              --����Ѿ�ȡ���������룬���������.
              Select Count(1)
              Into v_Count
              From ���˷������� A, סԺ���ü�¼ B, ����ҽ����¼ C
              Where a.����id = b.Id And c.Id = b.ҽ����� And (c.Id = ҽ��id_In Or c.���id = ҽ��id_In);
              If v_Count > 0 Then
                v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
                Raise Err_Custom;
              Else
                --�õ��ϴ�ִ��ʱ�����Ϣ
                Select �ϴ�ִ��ʱ��, ִ��ʱ�䷽��, ��ʼִ��ʱ��, �ϴδ�ӡʱ��, Ƶ�ʼ��, �����λ
                Into v_�ϴ�ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ
                From ����ҽ����¼
                Where ID = ҽ��id_In;
                v_�ϴ�ʱ�� := To_Date(To_Char(v_�ϴ�ʱ�� + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --�޸��ϴ�ִ��ʱ��Ϊ�ջغ��ĩ��ִ��ʱ�䡣
                v_ĩ��ʱ�� := Null;
                Begin
                  --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С��
                  --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ��
                  --����ҩ;������ҩ�÷�����δ��д���ͼ�¼
                  Select /*+ Rule*/
                   ĩ��ʱ��, ���ͺ�
                  Into v_ĩ��ʱ��, v_���ͺ�
                  From ����ҽ������
                  Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                        ���ͺ� = (Select Max(���ͺ�)
                               From ����ҽ������
                               Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
              
                Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In And ���ͺ� = v_���ͺ�;
                If v_Count > 0 Then
                  --��ԭҽ��ִ��ʱ��
                  Select Zl_Adviceexetimes(ҽ��id_In, v_�ϴ�ʱ��, v_ĩ��ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ, 0)
                  Into v_Time
                  From Dual;
                  Insert Into ҽ��ִ��ʱ��
                    (Ҫ��ʱ��, ҽ��id, ���ͺ�)
                    Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), ҽ��id_In, v_���ͺ�
                    From Table(f_Str2list(v_Time));
                End If;
              End If;
            End If;
          End If;
        
          --����ȼ��䶯�������������䶯ʱ�����������
          If r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              Select Count(*), Max(a.����ȼ�id), Max(a.��ʼʱ��)
              Into v_Count, n_����ȼ�id, d_��ʼʱ��
              From ���˱䶯��¼ A
              Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6 And a.��ֹʱ�� Is Null And
                    a.���Ӵ�λ = 0;
              --���û���ҵ����һ���ǻ���ȼ��䶯���ֹ
              If v_Count = 0 Then
                --ҽ������ȼ�����סʱ��Ļ���ȼ�һ��ʱҪ�����ж�
                Select Count(*)
                Into v_Count
                From ���˱䶯��¼ A
                Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6;
                If v_Count > 0 Then
                  v_Error := '���ڻ���ȼ�ҽ��ֹͣ��ò����Ѿ������������䶯��¼,���ܻ��˸�ҽ����ֹͣ������';
                  Raise Err_Custom;
                End If;
              Else
                --���n_����ȼ�IDΪNull�������Ƿ��ǵ�ǰ���˵�ҽ����Ӧ�ı䶯��¼,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˡ�
                --���n_����ȼ�ID��ΪNull�����п�����У����һ������ȼ�ʱ���Զ�ֹͣ�ģ�δ�����䶯��¼��
                --     ����Ҫ��鵱ǰ���һ���䶯�Ļ���ȼ�ID�Ƿ��ǵ�ǰҽ���Ļ���ȼ�ID,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˣ����������Ҫ�ٳ������һ�α䶯��ֱ�ӻ���ҽ�����ɡ�
                If n_����ȼ�id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From ���˱䶯��¼ B, ����ҽ���Ƽ� C
                  Where b.����id = r_Rolladvice.����id And b.��ҳid = r_Rolladvice.��ҳid And c.ҽ��id = ҽ��id_In And
                        c.�շ�ϸĿid = b.����ȼ�id And b.��ֹʱ�� = d_��ʼʱ�� And b.��ֹԭ�� = 6 And b.���Ӵ�λ = 0;
                Else
                  --��ʼʱ��ֻȡ���ӶԱȣ�У�Ե�ʱ����ȼ��Ŀ�ʼʱ����ҽ����ʼʱ��+��ǰʱ�������
                  Select Count(*)
                  Into v_Count
                  From ����ҽ���Ƽ� C, ����ҽ����¼ A
                  Where a.Id = c.ҽ��id And a.Id = ҽ��id_In And c.�շ�ϸĿid = n_����ȼ�id And
                        a.��ʼִ��ʱ�� = To_Date(To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '�����˵�ҽ���������һ������ȼ�ҽ�����뽫����Ļ���ȼ�ҽ�����Ϻ��ٻ��˱���ҽ����';
                  Raise Err_Custom;
                End If;
              
                If n_����ȼ�id Is Null Then
                  --��ǰ������Ա
                  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
                    v_��Ա��� := ����Ա���_In;
                    v_��Ա���� := ����Ա����_In;
                  Else
                    v_Temp     := Zl_Identity;
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
                    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
                    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
                  End If;
                
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
                End If;
              End If;
            End If;
          End If;
        
          If r_Rolladvice.��� = 'Z' And Instr(',9,10,', ',' || r_Rolladvice.���� || ',') > 0 And
             Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            --���˲���ҽ��ʱ�����ñ䶯��¼����
            Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, '�����䶯');
          End If;
        
          --����ҽ��ֹͣʱ,���ͣ��ҽ����ʱ��,�����ʵϰҽʦ�������˵ģ���ָ������״̬
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Flag_In, 1, ִ����ֹʱ��, Null), ͣ��ҽ�� = Null, ͣ��ʱ�� = Null,
              ��˱�� = Decode(r_Rolladvice.��˱��, 3, 2, r_Rolladvice.��˱��)
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 9 Then
          --����ҽ��ȷ��ֹͣʱ,����Ƿ��Ѵ�ӡͣ��ʱ��
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From ����ҽ����ӡ
          Where ��ӡ��� = 1 And ҽ��id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ͣ��ʱ���Ѿ���ӡ�������ٳ���ȷ��ֹͣ������';
            Raise Err_Custom;
          End If;
        
          --����ҽ��ȷ��ֹͣʱ,���ͣ��ҽ����ʱ��
          Update ����ҽ����¼ Set ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 10 Then
          --���˱�עƤ�Խ��,ͬʱɾ�������Ǽ�(+)��(-),���ݼ�¼ʱ��
          --�����ļ�¼��ҽ�������޹ۣ�����Ҫ����
          Delete From ���˹�����¼
          Where ����id = r_Rolladvice.����id And Nvl(��ҳid, 0) = Nvl(r_Rolladvice.��ҳid, 0) And ��¼ʱ�� = r_Rolladvice.����ʱ�� And
                Nvl(���, 0) = 0;
        
          Update ����ҽ����¼ Set Ƥ�Խ�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 13 Then
          If Instr(r_Rolladvice.����ҽ��, '/') > 0 Then
            Update ����ҽ����¼ Set ��˱�� = 1 Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          Else
            Update ����ҽ����¼ Set ��˱�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          End If;
        End If;
        --��������ҽ�����ϲ���
        --����ҽ��  
        If r_Rolladvice.�������� = 4 And r_Rolladvice.��� = 'Z' Then
          Select Count(1) Into v_Count From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid And �������� = '4';
          If v_Count = 1 Then          
            b_Message.Zlhis_Cis_004(r_Rolladvice.����id, r_Rolladvice.��ҳid, ҽ��id_In);
          End If;
        End If;
      Else
        --����ҽ������(�Է��ͺŹؼ���)
        ------------------------------------------------------------------
        --��ǰ������Ա
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ������������ѯ������˵������Һ��¼
        Begin
          Select Decode(Max(�Ƿ�����), 1, 1, 0)
          Into v_Count
          From ��Һ��ҩ��¼
          Where ҽ��id = ҽ��id_In And ���ͺ� = r_Rolladvice.���ͺ�;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����ҺҩƷ���Ѿ�����Һ�����������������ܻ��˷��͡�';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_��Һ��ҩ��¼_ҽ������(ҽ��id_In, r_Rolladvice.���ͺ�, v_��Ա����, Sysdate);
        End If;
      
        --����Ƿ����δ��˵���������
        Select Count(*)
        Into v_Count
        From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C, ���˷������� D
        Where (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ����� And c.Id = d.����id And
              c.��¼״̬ In (0, 1, 3) And d.״̬ = 0;
      
        If v_Count > 0 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����δ��˵��������룬��ȡ�����������������ٻ��˷��͡�';
          Raise Err_Custom;
        End If;
      
        --���ҽ���Ƿ������Ч��ҽ������
        Select Count(*)
        Into v_Count
        From ����ҽ������ A, סԺ���ü�¼ B
        Where a.ҽ��id = b.ҽ����� And a.No = b.No And b.��¼״̬ = 1 And b.ʵ�ս�� <> 0 And a.���ͺ� = r_Rolladvice.���ͺ� And
              a.ҽ��id In (Select Column_Value From Table(t_Adviceids));
        If v_Count > 0 Then
          v_Error := '��ҽ���»����ڸ�����Ŀ�����ȳ�����';
          Raise Err_Custom;
        End If;
      
        --���Ʒ����Զ�ִ��ʱ������Ҳ�Զ�����ִ��(����ʿվ�д˹���)
        --�Ǹ������õ�����ҽ����ͬ��ͨҽ��ִ�д���
        Select ҽ����Ч Into v_ҽ����Ч From ����ҽ����¼ Where ID = ҽ��id_In;
        If Substr(zl_GetSysParameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.���ͺ�) Loop
            If Nvl(r_Rollsend.ִ��״̬, 0) = 1 And
               (Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˲���id, 0) Or
                Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˿���id, 0)) Then
            
              --ҽ����ִ��״̬
              Update ����ҽ������ Set ִ��״̬ = 0 Where ���ͺ� = r_Rollsend.���ͺ� And ҽ��id = r_Rollsend.ҽ��id;
              v_Update := 1;
            
              If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
                --���õ�ִ��״̬
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ�����
                      Update סԺ���ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ�����
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --סԺ���ҷ�ҩ��ҩƷ�Զ���ҩ
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --סԺ���˷��÷��͵�����������������Դ����סԺ��
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ�����
                      Update ������ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ����Զ�����
                      For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                        Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                      End Loop;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --�����ҷ�ҩ��ҩƷ�Զ���ҩ
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------
        --�������ջصĳ���ҩƷҽ�����������(���˷��þͶ�����)
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 0 Then
          If r_Rolladvice.�ϴ�ִ��ʱ�� Is Not Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            If r_Rolladvice.�ϴ�ִ��ʱ�� < r_Rolladvice.ĩ��ʱ�� Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '������ڷ��͵������ѱ��ջأ������ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.�ϴ�ִ��ʱ�� Is Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            --�������ܱ�ȫ�������ջ�
            v_Error := Nvl(ҽ������_In, '��ҽ��') || 'δ�����ͣ����͵������ѱ�ȫ�������ջأ������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.ִ��״̬, 0) In (1, 3) And v_Update <> 1 Then
          --1-��ȫִ��;3-����ִ��
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
          Raise Err_Custom;
        Else
          --������ҽ����ִ�У���ҲҪ���ƻ��ˣ����磺����Ĳɼ���ʽ��
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ִ��״̬ In (1, 3) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------
        --������ҽ���ķ�������(��һ��ҽ�������в�ͬNO����)
        --���ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж�
        v_����no   := Null;
        v_������� := Null;
        If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ��
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                Zl_סԺ���ʼ�¼_Delete(v_����no, Substr(v_�������, 2), v_��Ա���, v_��Ա����, 2, 0, 0);
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ��
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 And Not (Nvl(r_Rollmoney.ִ��״̬, 0) = -1 And Nvl(r_Rollmoney.��¼״̬, 0) = 0) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            --�շѵ������շ�
            If r_Rollmoney.��¼״̬ = 1 And Not (r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵����ﵥ��"' || r_Rollmoney.No || '"���շѣ����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
                  --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���)
                  Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
                Else
                  Zl_���ﻮ�ۼ�¼_Delete(v_����no, Substr(v_�������, 2));
                End If;
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        End If;
        If v_������� Is Not Null And v_����no Is Not Null Then
          v_������� := Substr(v_�������, 2);
          If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
            Zl_סԺ���ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����, 2, 0, 0);
          Elsif r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
            --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���)
            Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
          Else
            Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
          End If;
        End If;
        --��Ѫҽ����ɾ������ҽ������
        Delete From ����ҽ������ Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ��ҽ��ִ��ʱ�� (����ҽ��ID�Ų����˼�¼)
        Delete From ҽ��ִ��ʱ�� Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ�����ͼ�¼(����ҽ����)
        Delete /*+ Rule*/
        From ����ҽ������
        Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id In (Select Column_Value From Table(t_Adviceids));
      
        --���(����ҽ��)�ϴ�ִ��ʱ��(���ϴη��͵�ĩ��ִ��ʱ��)
        --���г���(���������Գ���)����ʱ����д��ĩ��ʱ��
        --��������û�У���ֻ���ܷ�����һ�Ρ�
        v_ĩ��ʱ�� := Null;
        Begin
          --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С��
          --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ��
          --����ҩ;������ҩ�÷�����δ��д���ͼ�¼
          Select /*+ Rule*/
           ĩ��ʱ��
          Into v_ĩ��ʱ��
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --������������ʱ��ͬʱ�Զ�����ֹͣ
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 1 Then
          --ɾ��(��������)�����ֹͣ״̬������¼
          Delete /*+ Rule*/
          From ����ҽ��״̬
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And �������� = 8;
          --r_RollAdvice.����ʱ��:����ʱ����ܲ����Զ�ֹͣʱ����ͬ��
        
          --ȡɾ����Ӧ�ָ���ҽ��״̬
          Select ��������
          Into v_ҽ��״̬
          From ����ҽ��״̬
          Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
        
          --�ָ�(����ҽ��)���˺��״̬
          Update ����ҽ����¼
          Set ҽ��״̬ = v_ҽ��״̬, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        End If;
      
        --סԺ����ҽ�����ͺ�Ļ���(3-ת��;5-��Ժ;6-תԺ,11-����)
        If r_Rolladvice.��� = 'Z' And Instr(',3,5,6,11,', ',' || r_Rolladvice.���� || ',') > 0 And
           Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.����id, r_Rolladvice.��ҳid);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.���� = '3' And r_Patilog.��ʼԭ�� = 3 Then
              --ȡ������ת��״̬
              If r_Patilog.��ʼʱ�� Is Null Then
                --ת��ҽ�������⴦����һ������������ת��ҽ��ʱ��ֻ�ܻ��������һ��,70443
                Select Count(1)
                Into v_Count
                From ����ҽ����¼ A, ������ĿĿ¼ B
                Where a.������Ŀid = b.Id And a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.������� = 'Z' And
                      b.�������� = '3' And a.ҽ��״̬ = 8 And
                      a.��ʼִ��ʱ�� > (Select ��ʼִ��ʱ�� From ����ҽ����¼ Where ID = ҽ��id_In);
                If v_Count = 0 Then
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'ת��');
                Else
                  v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.���� In ('5', '6', '11') And r_Patilog.��ʼԭ�� = 10 Then
              --ȡ������Ԥ��Ժ״̬
              Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'Ԥ��Ժ');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --���˲���ʱ��
        --1.�����¼�(ֻ��һ��ҽ����¼)��������7-����,8-����,11-����
        If r_Rolladvice.��� = 'F' Or r_Rolladvice.��� = 'Z' And Instr(',7,8,11,', ',' || r_Rolladvice.���� || ',') > 0 Then
          Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, ҽ��id_In);
        End If;
      
        --2.���⴦��֪��ͬ����(������ص�֪��ͬ�����ٴε��ã���Ϊ����������������Ŀ�����й�����֪��ͬ����)
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.���) > 0 Then
          For R In (Select a.Id, a.������� From ����ҽ����¼ A Where a.Id = ҽ��id_In Or a.���id = ҽ��id_In) Loop
            --���id��һ��ҽ����һ����������ģ�����Ҫ���ж�һ�����
            If Instr('C,D,E,F,G,K,L', r.�������) > 0 Then
              Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, r.Id);
            End If;
          End Loop;
        End If;
      
        --�˴������Ϣ����
        If r_Rolladvice.��� = 'E' And r_Rolladvice.�������� = '6' Then
          --����
          b_Message.Zlhis_Cis_036(r_Rolladvice.����id, r_Rolladvice.��ҳid, Null, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No, 2);
        Elsif r_Rolladvice.��� = 'D' And r_Rolladvice.���id Is Null Then
          --���
          b_Message.Zlhis_Cis_037(r_Rolladvice.����id, r_Rolladvice.��ҳid, Null, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No, 2);
        Elsif r_Rolladvice.��� = 'F' And r_Rolladvice.���id Is Null Then
          --����
          b_Message.Zlhis_Cis_038(r_Rolladvice.����id, r_Rolladvice.��ҳid, Null, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'K' And r_Rolladvice.���id Is Null Then
          --��Ѫ
          b_Message.Zlhis_Cis_039(r_Rolladvice.����id, r_Rolladvice.��ҳid, Null, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '6' Then
          --����
          b_Message.Zlhis_Cis_040(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '8' Then
          --����
          b_Message.Zlhis_Cis_041(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'Z' And r_Rolladvice.�������� = '11' Then
          --����
          b_Message.Zlhis_Cis_042(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'E' And r_Rolladvice.�������� = '5' Then
          --��������
          b_Message.Zlhis_Cis_043(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        Elsif r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.��������, '0') = '0' Then
          --������
          b_Message.Zlhis_Cis_007(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.��id,
                                  r_Rolladvice.No);
        End If;
      
        Select Count(1)
        Into v_Count
        From ��������˵�� A
        Where a.����id = r_Rolladvice.ִ�п���id And a.�������� = '����';
        If v_Count > 0 Then
          --����ִ��ҽ�����˷���
          b_Message.Zlhis_Cis_044(r_Rolladvice.����id, r_Rolladvice.��ҳid, r_Rolladvice.���ͺ�, r_Rolladvice.ҽ��id,
                                  r_Rolladvice.No, r_Rolladvice.��������, r_Rolladvice.�״�ʱ��, r_Rolladvice.ĩ��ʱ��,
                                  r_Rolladvice.��������);
        End If;
      
      End If;
    End If;
    Exit When r_Rolladvice.���ͺ� = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_����;
/

--122609:������,2018-03-08,����ƽ̨��Ϣ���
Create Or Replace Procedure Zl_����ҽ����¼_����
(
  Id_In         In ����ҽ����¼.Id%Type,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null,
  ����ҽ��id_In In ����ҽ����¼.Id%Type := Null,
  ����ʱ��_In   In ����ҽ��״̬.����ʱ��%Type := Null
) Is
  --���ܣ�����ָ����ҽ��(δ���͵ĳ���������)
  --˵����һ����ҩ��ֻ�ܵ���һ��(������ʾ�ж���)
  --������ID_IN=��ҽ��ID
  --      ����ҽ��id_In ȡ�����������ϵĻ���ȼ�ҽ�����������Զ�ֹͣ�Ļ���ȼ�ҽ��id
  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  v_����no       ������ü�¼.No%Type;
  v_��¼����     ������ü�¼.��¼����%Type;
  v_�������     Varchar2(255);
  n_�Զ�ȡ��ִ�� Number(1) := 0;
  n_�����Ϻ���ҩ Number(1) := 0;

  v_Date     Date;
  v_Count    Number;
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_No       ����ҽ������.No%Type;

  --����ҽ�������Ϣ
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.����id, a.�Һŵ�, a.��ҳid, a.Ӥ��, a.ҽ��״̬, a.�ϴ�ִ��ʱ��, a.ҽ������, a.�������, b.��������, a.������Դ,
           a.ִ�п���id, b.ִ��Ƶ��, a.������Ŀid, a.��ʼִ��ʱ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.Id = Id_In;
  r_Advice c_Advice%RowType;

  --����ҽ������ʱ��ȡ��Ӧ�ķ������ʻ�����(�շѻ��۵�)��
  --����ҽ��������NO������λ���Ҫ���ʻ��˷ѵļ�¼
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,Ҳ��һ�����Ʒ���,�ҿ���NO��ͬ
  --ֻ�ܼ�¼״̬Ϊ1�ļ�¼,����Ѿ����ʻ򲿷����ʵļ�¼,���ٴ���
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ�������
  --���������ҩ�������Ϻ���ҩ��,�򲻶���Ӧ����(������ҩ;����)���м��ʹ���,�����ǻ�û��ִ�еļ��ʵ�,��δִ�С��շѵĻ��۵���������ɾ�ˡ�


  Cursor c_Rollmoney(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Decode(a.��¼����, 11, 1, a.��¼����) As ��¼����, a.��¼״̬, a.No, a.���, a.ִ��״̬ As ����ִ��, c.ִ��״̬ As ҽ��ִ��, c.ִ�в���id, b.���˿���id,
           b.�������, i.��������
    From ������ü�¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ I
    Where c.ҽ��id = b.Id And c.���ͺ� = v_���ͺ� And (b.Id = Id_In Or b.���id = Id_In) And a.ҽ����� = b.Id And a.��¼״̬ In (0, 1) And
          a.No = c.No And (a.��¼���� = c.��¼���� Or a.��¼���� = 11 And c.��¼���� = 1) And b.������Ŀid = i.Id And a.�۸񸸺� Is Null And
          (n_�����Ϻ���ҩ = 0 Or
          n_�����Ϻ���ҩ = 1 And
          Not (Exists (Select 1
                        From ������ü�¼ D
                        Where d.ҽ����� = b.Id And d.��¼״̬ In (0, 1) And d.No = c.No And
                              (d.��¼���� = c.��¼���� Or d.��¼���� = 11 And c.��¼���� = 1) And d.�շ���� In ('5', '6', '7'))) Or
          Nvl(a.ִ��״̬, 0) = 0 And Not (a.��¼���� = 1 And a.��¼״̬ <> 0))
    Order By a.��¼����, a.No, a.���, a.�շ�ϸĿid;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --���ҽ��״̬�Ƿ���ȷ:��������
  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;

  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա��� := ����Ա���_In;
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --����Ƿ��Ѿ����˱��浥���Ѿ������浥��ҽ�����ܹ�����
  Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = Id_In;
  If v_Count > 0 Then
    If Not (r_Advice.�������� = '7' And r_Advice.������� = 'Z') Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�ѳ����棬�������ϡ�';
      Raise Err_Custom;
    End If;
  End If;

  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
  Select Count(1) Into v_Count From ��Һ��ҩ��¼ Where �Ƿ����� = 1 And ҽ��id = Id_In;
  If v_Count > 0 Then
    v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"����ҺҩƷ���Ѿ�����Һ���������������������ϡ�';
    Raise Err_Custom;
  End If;

  If r_Advice.�Һŵ� Is Null And r_Advice.������Դ <> 3 Then
    If r_Advice.ҽ��״̬ In (4, 8, 9) Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ������ϻ�ֹͣ�����������ϡ�';
      Raise Err_Custom;
    Elsif r_Advice.�ϴ�ִ��ʱ�� Is Not Null Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ����ͣ����ܱ����ϡ�';
      Raise Err_Custom;
    End If;
  
    --�����Ի���ȼ����뷢�ͣ�У�Ժ�Ϳ������Զ��Ʒѣ����ϼ��������϶�Ӧ��ֹͣ���̴���
    If r_Advice.������� = 'H' And r_Advice.�������� = '1' And r_Advice.ִ��Ƶ�� = '2' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      --(��ȡ�������ڴ����޷���Ժ�����������ţ�45977)a.��ʼʱ���ǵ���֮ǰ�ģ�˵������Ч���Զ����ü��㣩�����������ϡ�
      --ҽ����ʱ��ֻ��ȷ���˷��ӣ����Ա䶯��¼�Ŀ�ʼʱ��Ҫȥ�������Ƚϡ�
      v_Count := 0;
      Begin
        Select b.��ֹʱ��
        Into v_Date
        From ���˱䶯��¼ B, ����ҽ���Ƽ� C
        Where b.����id = r_Advice.����id And b.��ҳid = r_Advice.��ҳid And c.ҽ��id = Id_In And c.�շ�ϸĿid = b.����ȼ�id And
              b.��ʼԭ�� = 6 And b.���Ӵ�λ = 0 And
              To_Char(b.��ʼʱ��, 'yyyy-mm-dd hh24:mi') = To_Char(r_Advice.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi');
      Exception
        When Others Then
          v_Count := 1;
      End;
      If v_Count = 0 Then
        --d.�����������䶯����
        If v_Date Is Not Null Then
          v_Error := '���ڻ���ȼ�ҽ����Ч���Ѿ������������䶯��¼,�������ϸ�ҽ����';
          Raise Err_Custom;
        Else
          --������Ҫ�Զ����õĻ���ȼ��������ԭ������ȼ���ͬ���ó�������䶯��¼
          If Nvl(����ҽ��id_In, 0) <> 0 Then
            Delete ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In And �������� In (8, 9);
            Select ��������
            Into v_Count
            From (Select �������� From ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In Order By ����ʱ�� Desc)
            Where Rownum < 2;
            Update ����ҽ����¼
            Set ҽ��״̬ = v_Count, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null, ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null
            Where ID = ����ҽ��id_In;
            --�ų�����Ƶ���Ĳ���
            Select Count(a.Id)
            Into v_Count
            From ����ҽ����¼ A, �����շѹ�ϵ B, ������ҳ C
            Where a.������Ŀid = b.������Ŀid And c.����ȼ�id = b.�շ���Ŀid And c.����id = a.����id And c.��ҳid = a.��ҳid And
                  a.Id = ����ҽ��id_In;
          End If;
          If v_Count = 0 Then
            --c.����ȼ������һ���䶯
            Zl_���˱䶯��¼_Undo(r_Advice.����id, r_Advice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
          End If;
        End If;
      Else
        --�ָ����һ�α��Զ�ֹͣ�Ļ���ȼ�
        If Nvl(����ҽ��id_In, 0) <> 0 Then
          Delete ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In And �������� In (8, 9);
          Select ��������
          Into v_Count
          From (Select �������� From ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In Order By ����ʱ�� Desc)
          Where Rownum < 2;
          Update ����ҽ����¼
          Set ҽ��״̬ = v_Count, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null, ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null
          Where ID = ����ҽ��id_In;
        Else
          --������Ժʱָ���Ļ��������ı䶯��¼��ҽ���¿������ı䶯��¼��ͬ������Ҫ���ж�
          Select Count(a.Id)
          Into v_Count
          From ���˱䶯��¼ A
          Where a.����id = r_Advice.����id And a.��ҳid = r_Advice.��ҳid And a.��ʼԭ�� = 6;
          If v_Count <> 0 Then
            --b.�������ǰ�Ļ���ȼ���ͬ����У��ʱû�в�������ȼ��䶯,��������ȼ�ֹͣ�䶯
            Zl_���˱䶯��¼_Nurse(r_Advice.����id, r_Advice.��ҳid, Null, Sysdate, v_��Ա���, v_��Ա����);
          End If;
        End If;
      End If;
    End If;
  Else
    If r_Advice.ҽ��״̬ <> 8 Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"��δ���ͻ��Ѿ����ϡ�';
      Raise Err_Custom;
    End If;
    --ҽ�������ж�
    Select Count(1)
    Into v_Count
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.ҽ��id = b.Id And (b.Id = Id_In Or b.���id = Id_In);
    If v_Count <> 0 Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"���ڸ��ӷ��ã��������ϡ�';
      Raise Err_Custom;
    End If;
  
    Begin
      --ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
      Select Distinct ���ͺ�
      Into v_���ͺ�
      From ����ҽ������
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
    Exception
      When Others Then
        v_���ͺ� := Null;
    End;
  
    Select Zl_To_Number(Nvl(zl_GetSysParameter(68), 0)) Into n_�����Ϻ���ҩ From Dual;
    Select Zl_To_Number(Nvl(zl_GetSysParameter('���ﱾ���Զ�ִ��', '1252'), 0)) Into n_�Զ�ȡ��ִ�� From Dual;
    If n_�Զ�ȡ��ִ�� = 1 And v_���ͺ� Is Not Null Then
      --�ȸ���ҽ���ͷ��õ�ִ��״̬����Ϊ�������жϣ��Լ�����Zl_������ʼ�¼_Delete���м��
      For Rc In (Select a.ҽ��id, a.ִ�в���id
                 From ����ҽ������ A, ����ҽ����¼ B
                 Where a.ҽ��id = b.Id And (b.Id = Id_In Or b.���id = Id_In) And a.ִ�в���id = b.���˿���id) Loop
        Zl_����ҽ��ִ��_Cancel(Rc.ҽ��id, v_���ͺ�, Null, 1, Rc.ִ�в���id);
      End Loop;
    End If;
  
    --����ҽ��ֻ���ܷ���һ��
    --�����˷�ʱ���м�飬��Ϊ����ҽ��û�з��ã�����Ҫ���һ��ִ��״̬
    Select Count(*)
    Into v_Count
    From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ I
    Where a.ҽ��id = b.Id And b.������Ŀid = i.Id And a.ִ��״̬ In (1, 3) And (b.Id = Id_In Or b.���id = Id_In) And
          (n_�����Ϻ���ҩ = 0 Or
          n_�����Ϻ���ҩ = 1 And Not (b.������� In ('5', '6', '7') Or b.������� = 'E' And i.�������� In ('2', '3', '4')));
    If v_Count > 0 Then
      v_Error := '��ҽ���Ѿ�ִ�л�����ִ�У��������ϡ�';
      Raise Err_Custom;
    End If;
  End If;

  If ����ʱ��_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := ����ʱ��_In;
  End If;

  Update ����ҽ����¼ Set ҽ��״̬ = 4 Where ID = Id_In Or ���id = Id_In;

  Insert Into ����ҽ��״̬
    (ҽ��id, ��������, ������Ա, ����ʱ��)
    Select ID, 4, v_��Ա����, v_Date From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In;

  --סԺҽ������ʱ,δ��ӡ�������,ȱʡ����Ϊ���δ�ӡ
  If r_Advice.�Һŵ� Is Null Then
    Select Count(*)
    Into v_Count
    From ����ҽ����ӡ
    Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
    If Nvl(v_Count, 0) = 0 Then
      Zl_����ҽ����¼_���δ�ӡ(Id_In, 1);
    End If;
    If Nvl(r_Advice.Ӥ��, 0) > 0 And r_Advice.�������� = '11' Then
      Update ������������¼
      Set ����ʱ�� = Null
      Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And ��� = Nvl(r_Advice.Ӥ��, 0);
    End If;
  Else
    --����ҽ��(����)����ʱ����Ҫ�����������:ֻ��һ�η���
    --���˻��ۻ���ʷ���
    If v_���ͺ� Is Not Null Then
      --������ҽ���ķ���ɾ��������(��һ��ҽ�������в�ͬNO����)
      --������ʣ����ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж�
      --���ﻮ�ۣ�������շѣ�������ɾ��
      v_����no   := Null;
      v_������� := Null;
      For r_Rollmoney In c_Rollmoney(v_���ͺ�) Loop
        If Nvl(r_Rollmoney.ҽ��ִ��, 0) In (1, 3) Then
          --1-��ȫִ��;3-����ִ��
          v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ�ִ�л�����ִ�У��������ϡ�';
          Raise Err_Custom;
        End If;
        If Nvl(r_Rollmoney.����ִ��, 0) In (1, 2) Then
          --1-��ȫִ��;2-����ִ��
          v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�е������Ѿ�ȫ���򲿷�ִ�У��������ϡ�';
          Raise Err_Custom;
        End If;
        If r_Rollmoney.����ִ�� = 9 Then
          v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�е��շѽ�������쳣���������ϡ�';
          Raise Err_Custom;
        End If;
        v_Count := 1;
        If r_Rollmoney.��¼���� = 1 And r_Rollmoney.��¼״̬ <> 0 Then
          If 1 = n_�����Ϻ���ҩ And r_Rollmoney.������� = 'E' And r_Rollmoney.�������� In ('2', '3', '4') Then
            v_Count := 0;
          Else
            v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�Ѿ��շѣ��������ϡ�';
            Raise Err_Custom;
          End If;
        End If;
        If 1 = v_Count Then
          If Nvl(v_����no, '��') <> r_Rollmoney.No Then
            If v_������� Is Not Null And v_����no Is Not Null Then
              v_������� := Substr(v_�������, 2);
              If v_��¼���� = 1 Then
                Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
              Elsif v_��¼���� = 2 Then
                Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
              End If;
            End If;
            v_������� := Null;
          End If;
          v_��¼���� := r_Rollmoney.��¼����;
          v_����no   := r_Rollmoney.No;
          v_������� := v_������� || ',' || r_Rollmoney.���;
        End If;
      End Loop;
      If v_������� Is Not Null And v_����no Is Not Null Then
        v_������� := Substr(v_�������, 2);
        If v_��¼���� = 1 Then
          Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
        Elsif v_��¼���� = 2 Then
          Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
        End If;
      End If;
    
      --���"����ҩ�������Ϻ���ҩ"�����Ӧ�ĸ�ҩ;����������Ϊδִ�У��Ա��˷�
      If n_�����Ϻ���ҩ = 1 Then
        Update ������ü�¼
        Set ִ��״̬ = 0
        Where ִ��״̬ = 1 And ҽ����� = Id_In And Exists
         (Select 1
               From ����ҽ����¼ A, ������ĿĿ¼ B
               Where a.������Ŀid = b.Id And b.��� = 'E' And b.�������� In ('2', '3', '4') And a.Id = Id_In);
      End If;
    
      --����ҽ�����ͼ�¼(��ִ�м�¼)
      Delete From ����ҽ��ִ�� Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
      Delete From ����ҽ������ Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
    
      --��������ҽ���Ĵ���
      If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      
        If r_Advice.�������� = '1' And r_Advice.ִ�п���id Is Not Null Then
          --����ҽ��
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0 And ��Ժ����id = r_Advice.ִ�п���id And �������� In (1, 2);
          If v_Count = 1 Then
            Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0);
          End If;
        Elsif r_Advice.�������� = '2' And r_Advice.ִ�п���id Is Not Null Then
          --סԺҽ��
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0 And ��Ժ����id = r_Advice.ִ�п���id And Nvl(��������, 0) = 0;
          If v_Count = 1 Then
            Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0);
          End If;
        End If;
      End If;
    End If;
  End If;

  --ɾ�������ǼǼ�¼
  If r_Advice.������� = 'E' And r_Advice.�������� = '1' Then
    --Update ����ҽ����¼ Set Ƥ�Խ��=Null Where ID=ID_IN; --��������Ƥ�Խ��
    --ɾ���������ļ�¼��������¼��������Ϊ����ҽ���Ƿ����ϣ����˶Ը�ҩ����
    For r_Test In (Select ����ʱ�� From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 10) Loop
      Delete From ���˹�����¼
      Where ����id = r_Advice.����id And ��¼��Դ = 2 And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And ��¼ʱ�� = r_Test.����ʱ�� And
            Nvl(���, 0) = 0;
    End Loop;
  End If;
  If r_Advice.������Ŀid Is Not Null Then
    --������¼��ҽ����סԺ����ִ��ҽ�����ϣ�Zlhis_Cis_003
    If r_Advice.��ҳid Is Not Null Then
      For R In (Select a.Id
                From ����ҽ����¼ A
                Where (a.Id = Id_In Or a.���id = Id_In) And Exists
                 (Select 1 From ��������˵�� B Where b.����id = a.ִ�п���id And b.�������� = '����')) Loop
        b_Message.Zlhis_Cis_003(r_Advice.����id, r_Advice.��ҳid, Null, r.Id);
      End Loop;
    
      If r_Advice.������� = 'Z' And r_Advice.�������� = '4' Then
        --����ҽ������
        b_Message.Zlhis_Cis_005(r_Advice.����id, r_Advice.��ҳid, r_Advice.��id);
      End If;
    End If;
  
    If r_Advice.�Һŵ� Is Not Null Then
      If r_Advice.������� = 'E' And r_Advice.�������� = '6' Or Instr(',D,F,K,', r_Advice.�������) > 0 Then
        Select Max(a.No) Into v_No From ����ҽ������ A Where a.ҽ��id = r_Advice.��id;
        If r_Advice.������� = 'E' And r_Advice.�������� = '6' Then
          --����
          b_Message.Zlhis_Cis_036(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No, 1);
        Elsif r_Advice.������� = 'D' Then
          --���
          b_Message.Zlhis_Cis_037(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No, 1);
        Elsif r_Advice.������� = 'F' Then
          --����
          b_Message.Zlhis_Cis_038(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No);
        Elsif r_Advice.������� = 'K' Then
          --��Ѫ
          b_Message.Zlhis_Cis_039(r_Advice.����id, Null, r_Advice.�Һŵ�, v_���ͺ�, r_Advice.��id, v_No);
        End If;
      End If;
    End If;
  End If;
  Close c_Advice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_����;
/

--122609:������,2018-03-08,����ƽ̨��Ϣ���
Create Or Replace Procedure Zl_����ҽ����¼_У��
(
  --���ܣ�У��ָ����ҽ��
  --������ҽ��ID_IN=Nvl(���ID,ID)
  --      ״̬_IN=У��ͨ��3��У������2
  --      �Զ�У��_IN=����֮������Զ�У��,�Զ���д�Ƽ�����
  --˵����һ��ҽ��ֻ�ܵ���һ��,����ͬʱ��ɴ���һ��ҽ����У��
  ҽ��id_In     In ����ҽ����¼.Id%Type,
  ״̬_In       In ����ҽ����¼.ҽ��״̬%Type,
  У��ʱ��_In   In ����ҽ��״̬.����ʱ��%Type,
  У��˵��_In   In ����ҽ��״̬.����˵��%Type := Null,
  �Զ�У��_In   In Number := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null
) Is
  --����ҽ�����
  v_״̬       ����ҽ����¼.ҽ��״̬%Type;
  v_��Ч       ����ҽ����¼.ҽ����Ч%Type;
  v_����id     ����ҽ����¼.����id%Type;
  v_��ҳid     ����ҽ����¼.��ҳid%Type;
  v_Ӥ��       ����ҽ����¼.Ӥ��%Type;
  v_ҽ������   ����ҽ����¼.ҽ������%Type;
  v_����ʱ��   ����ҽ����¼.����ʱ��%Type;
  v_��ʼʱ��   ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_����ҽ��   ����ҽ����¼.����ҽ��%Type;
  v_ǰ��id     ����ҽ����¼.ǰ��id%Type;
  v_ִ�б��   ����ҽ����¼.ִ�б��%Type;
  v_ִ�п���id ����ҽ����¼.ִ�п���id%Type;
  v_�걾��λ   ����ҽ����¼.�걾��λ%Type;
  v_ֹͣʱ��   ����ҽ����¼.����ʱ��%Type;
  v_��������id ����ҽ����¼.��������id%Type;
  n_���˿���id ����ҽ����¼.���˿���id%Type;

  --���ڱ������ȼ�
  v_�������   ����ҽ����¼.�������%Type;
  v_������Ŀid ����ҽ����¼.������Ŀid%Type;
  v_��������   ������ĿĿ¼.��������%Type;
  v_����ȼ�id ������ҳ.����ȼ�id%Type;
  v_������־   ����ҽ����¼.������־%Type;
  v_��Ժ��ʽ   ��Ժ��ʽ.����%Type;

  v_ҩƷ�ȼ�   �շѼ۸�ȼ�.����%Type;
  v_���ĵȼ�   �շѼ۸�ȼ�.����%Type;
  v_��ͨ�ȼ�   �շѼ۸�ȼ�.����%Type;
  v_Pricegrade Varchar2(1000);
  v_վ��       ���ű�.վ��%Type;

  v_Stopadviceids ����ҽ����¼.ҽ������%Type;
  n_Adviceid      ����ҽ����¼.����id%Type;
  n_���          Number(18);
  --�����Ŀͬһ�Զ�ֹͣ���������Ŀ:����Ӧ�ö��ǳ���(������ǰҽ��),����Ӧ�Ѽ�顣
  --ע��Ӧ��Ӥ������,ͬʱҲӦֹͣ����ǰҽ�����������ͬ������Ŀ��ҽ����
  Cursor c_Exclude Is
    Select Distinct b.Id As ҽ��id, b.��ʼִ��ʱ��, b.ִ����ֹʱ��, b.�ϴ�ִ��ʱ��, b.����ҽ��, b.ִ��ʱ�䷽��, b.Ƶ�ʼ��, b.Ƶ�ʴ���, b.�����λ
    From ���ƻ�����Ŀ A, ����ҽ����¼ B
    Where a.���� = 3 And a.��Ŀid = b.������Ŀid And b.Id <> ҽ��id_In And Nvl(b.ҽ����Ч, 0) = 0 And b.ҽ��״̬ In (3, 5, 6, 7) And
          b.����id = v_����id And Nvl(b.��ҳid, 0) = Nvl(v_��ҳid, 0) And Nvl(b.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And
          a.���� In (Select Distinct ���� From ���ƻ�����Ŀ Where ���� = 3 And ��Ŀid = v_������Ŀid)
    Order By b.Id;
  v_��ֹʱ�� ����ҽ����¼.ִ����ֹʱ��%Type;

  --����ȼ�����
  Cursor c_Nurse Is
    Select a.Id As ҽ��id, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.�ϴ�ִ��ʱ��, a.����ҽ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.������� = 'H' And b.�������� = '1' And a.����id = v_����id And Nvl(a.��ҳid, 0) = Nvl(v_��ҳid, 0) And
          Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ In (3, 5, 6, 7) And a.Id <> ҽ��id_In;

  --��¼���������
  Cursor c_Patiio Is
    Select a.Id As ҽ��id, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.�ϴ�ִ��ʱ��, a.����ҽ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.������� = 'Z' And b.�������� = '12' And a.����id = v_����id And Nvl(a.��ҳid, 0) = Nvl(v_��ҳid, 0) And
          Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ In (3, 5, 6, 7) And a.Id <> ҽ��id_In;

  --��¼���黥��
  Cursor c_Patistate Is
    Select a.Id As ҽ��id, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.�ϴ�ִ��ʱ��, a.����ҽ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.������� = 'Z' And b.�������� In ('9', '10') And a.����id = v_����id And
          Nvl(a.��ҳid, 0) = Nvl(v_��ҳid, 0) And Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And Nvl(a.ҽ����Ч, 0) = 0 And
          a.ҽ��״̬ In (3, 5, 6, 7) And a.Id <> ҽ��id_In;
  --�䶯��Ч��¼
  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From ���˱䶯��¼ C
           Where c.����id = v_����id And c.��ҳid = v_��ҳid And
                 c.��ʼʱ�� = (Select Min(y.��ʼʱ��)
                           From ���˱䶯��¼ Y
                           Where y.����id = v_����id And y.��ҳid = v_��ҳid And y.��ʼʱ�� > v_��ʼʱ��) And
                 Nvl(c.��ֹʱ�� || '', '��') =
                 (Select Nvl(Min(x.��ֹʱ��) || '', '��')
                  From ���˱䶯��¼ X
                  Where x.����id = v_����id And x.��ҳid = v_��ҳid And x.��ʼʱ�� > v_��ʼʱ��)) A, ���˱䶯��¼ B
    Where b.����id = v_����id And b.��ҳid = v_��ҳid And a.��ʼʱ�� = b.��ֹʱ�� And a.��ʼԭ�� = b.��ֹԭ�� And a.���Ӵ�λ = b.���Ӵ�λ
    Union
    Select a.*
    From ���˱䶯��¼ A
    Where a.����id = v_����id And a.��ҳid = v_��ҳid And a.��ֹʱ�� Is Null And a.��ʼʱ�� <= v_��ʼʱ��;

  Cursor c_Endinfo Is
    Select * From ���˱䶯��¼ Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null;
  r_Oldinfo      c_Oldinfo%RowType;
  r_Endinfo      c_Endinfo%RowType;
  v_�䶯��ֹԭ�� ���˱䶯��¼.��ֹԭ��%Type;
  v_�䶯��ֹʱ�� ���˱䶯��¼.��ֹʱ��%Type;
  v_�䶯��ֹ��Ա ���˱䶯��¼.��ֹ��Ա%Type;

  --��������(Ӥ��)������δͣ����(���䷽����)
  Cursor c_Needstop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select a.Id, a.�������, b.��������, b.ִ��Ƶ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id(+) And a.����id = v_����id And a.��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� < v_Stoptime
    Order By a.���;
  --��������(Ӥ��)����ͣ��δȷ�ϵĳ���,��ִֹ��ʱ����ָ��ʱ��֮��
  Cursor c_Havestop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select ID
    From ����ҽ����¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And Nvl(ҽ����Ч, 0) = 0 And
          ҽ��״̬ = 8 And ִ����ֹʱ�� > v_Stoptime And ��ʼִ��ʱ�� < v_Stoptime
    Order By ���;

  --ȡһ��ҽ���ļƼ�����
  Cursor c_Price Is
    Select a.Id, b.�շ���Ŀid, b.�շ�����, b.������Ŀ, b.��������, b.�շѷ�ʽ, c.��� As �շ����, a.�������, e.��������, e.�Թܱ���,
           Sum(Decode(Nvl(c.�Ƿ���, 0), 1, Nvl(d.ȱʡ�۸�, d.ԭ��), Null)) As ����
    From ����ҽ����¼ A, �����շѹ�ϵ B, �շ���ĿĿ¼ C, �շѼ�Ŀ D, ������ĿĿ¼ E
    Where a.������Ŀid = b.������Ŀid And b.�շ���Ŀid = c.Id And c.Id = d.�շ�ϸĿid And
          ((Instr(';5;6;7;', ';' || c.��� || ';') > 0 And d.�۸�ȼ� = v_ҩƷ�ȼ�) Or
          (Instr(';4;', ';' || c.��� || ';') > 0 And d.�۸�ȼ� = v_���ĵȼ�) Or
          (Instr(';4;5;6;7;', ';' || c.��� || ';') = 0 And d.�۸�ȼ� = v_��ͨ�ȼ�) Or
          (d.�۸�ȼ� Is Null And Not Exists
           (Select 1
             From �շѼ�Ŀ
             Where c.Id = �շ�ϸĿid And ((Instr(';5;6;7;', ';' || c.��� || ';') > 0 And �۸�ȼ� = v_ҩƷ�ȼ�) Or
                   (Instr(';4;', ';' || c.��� || ';') > 0 And �۸�ȼ� = v_���ĵȼ�) Or
                   (Instr(';4;5;6;7;', ';' || c.��� || ';') = 0 And �۸�ȼ� = v_��ͨ�ȼ�))))) And
          (a.���id Is Null And a.ִ�б�� In (1, 2) And b.�������� = 1 Or
          a.�걾��λ = b.��鲿λ And a.��鷽�� = b.��鷽�� And Nvl(b.��������, 0) = 0 Or
          a.��鷽�� Is Null And Nvl(b.��������, 0) = 0 And b.��鲿λ Is Null And b.��鷽�� Is Null) And
          a.������� Not In ('5', '6', '7') And Nvl(a.�Ƽ�����, 0) = 0 And Nvl(a.ִ������, 0) Not In (0, 5) And c.������� In (2, 3) And
          (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And Sysdate Between d.ִ������ And
          Nvl(d.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(b.�շ�����, 0) <> 0 And
          Not (Nvl(c.�Ƿ���, 0) = 1 And Nvl(Nvl(d.ȱʡ�۸�, d.ԭ��), 0) = 0) And a.������Ŀid = e.Id And
          (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Group By a.Id, b.�շ���Ŀid, b.�շ�����, b.������Ŀ, b.��������, b.�շѷ�ʽ, c.���, a.�������, e.��������, e.�Թܱ���;

  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select * From ������Ϣ Where ����id = v_����id;
  r_Pati c_Pati%RowType;

  v_����id ��Ѫ������.����id%Type;

  --������ʱ����
  v_Count    Number;
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_Parͣ��  Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_����ִ�� Varchar2(5);

  v_Error Varchar2(255);
  Err_Custom Exception;

  Function Getadvicetext(v_ҽ��id ����ҽ����¼.Id%Type) Return Varchar2 Is
    v_Text ����ҽ����¼.ҽ������%Type;
    v_��� ����ҽ����¼.�������%Type;
    v_�䷽ Number;
  Begin
    Select �������, ҽ������ Into v_���, v_Text From ����ҽ����¼ Where ID = v_ҽ��id;
    If v_��� = 'E' Then
      --��ҩ���г�ҩ��ҽ������
      Begin
        Select �������, Decode(�������, '7', v_Text, ҽ������)
        Into v_���, v_Text
        From ����ҽ����¼
        Where ���id = v_ҽ��id And ������� In ('5', '6', '7') And Rownum = 1;
      Exception
        When Others Then
          Null;
      End;
      If v_��� = '7' Then
        v_�䷽ := 1;
      End If;
    End If;
    If Length(v_Text) > 30 Then
      v_Text := Substr(v_Text, 1, 30) || '...';
    End If;
    If Length(v_Text) > 20 Then
      v_Text := '"' || v_Text || '"' || Chr(13) || Chr(10);
    Else
      v_Text := '"' || v_Text || '"';
    End If;
    If v_�䷽ = 1 Then
      v_Text := '��ҩ�䷽' || v_Text;
    End If;
    Return(v_Text);
  End;
Begin
  --���ҽ��״̬�Ƿ���ȷ:��������
  Begin
    Select a.ҽ����Ч, a.ҽ��״̬, a.����ʱ��, a.����ҽ��, a.��ʼִ��ʱ��, a.����id, a.��ҳid, a.Ӥ��, a.ҽ������, a.�������, a.������Ŀid, a.ǰ��id,
           Nvl(b.��������, '0'), Nvl(a.ִ�б��, 0), a.ִ�п���id, a.�걾��λ, a.��������id, Nvl(a.������־, 0) As ������־, a.���˿���id
    Into v_��Ч, v_״̬, v_����ʱ��, v_����ҽ��, v_��ʼʱ��, v_����id, v_��ҳid, v_Ӥ��, v_ҽ������, v_�������, v_������Ŀid, v_ǰ��id, v_��������, v_ִ�б��,
         v_ִ�п���id, v_�걾��λ, v_��������id, v_������־, n_���˿���id
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id(+) And a.Id = ҽ��id_In;
  Exception
    When Others Then
      Begin
        v_Error := 'ҽ���ѱ�ɾ�������ܽ���У�ԡ�' || Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
        Raise Err_Custom;
      End;
  End;
  If v_״̬ <> 1 Then
    v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"�����¿���ҽ��������ͨ��У�ԡ�' || Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
    Raise Err_Custom;
  End If;
  --�ٴμ��У��ʱ�����Ч��:��������
  If To_Char(v_����ʱ��, 'YYYY-MM-DD HH24:MI') <= To_Char(v_��ʼʱ��, 'YYYY-MM-DD HH24:MI') Then
    If To_Char(У��ʱ��_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_����ʱ��, 'YYYY-MM-DD HH24:MI') Then
      v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"��У��ʱ�䲻��С�ڿ���ʱ�� ' || To_Char(v_����ʱ��, 'YYYY-MM-DD HH24:MI') || '��' ||
                 Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
      Raise Err_Custom;
    End If;
  Else
    If To_Char(У��ʱ��_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_��ʼʱ��, 'YYYY-MM-DD HH24:MI') Then
      v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"��У��ʱ�䲻��С�ڿ�ʼִ��ʱ�� ' || To_Char(v_��ʼʱ��, 'YYYY-MM-DD HH24:MI') || '��' ||
                 Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
      Raise Err_Custom;
    End If;
  End If;

  --���Ҫ��ǩ�������У��ʱ�Ƿ���ǩ��(����ȡ��ǩ��)
  If ״̬_In = 3 Then
    Select Zl_Fun_Getsignpar(Decode(v_ǰ��id, Null, 1, 3), v_��������id) Into v_Count From Dual;
    If v_Count = 1 Then
      --֤��ͣ�û�δע��֤�鲻����ǩ������ֻ�ж�һ�����ݼ���
      For C In (Select a.�Ƿ�ͣ��
                From ��Ա֤���¼ A, ��Ա�� B
                Where a.��Աid = b.Id And b.���� = v_����ҽ��
                Order By a.ע��ʱ�� Desc) Loop
        If Nvl(c.�Ƿ�ͣ��, 0) = 0 Then
          Select Count(*)
          Into v_Count
          From ����ҽ��״̬ A
          Where �������� = 1 And ҽ��id = ҽ��id_In And
                (ǩ��id Is Null And Exists
                 (Select 1
                  From ��Ա�� R, ��Ա����˵�� X
                  Where r.Id = x.��Աid And r.���� = a.������Ա And x.��Ա���� = '��ʿ') And Not Exists
                 (Select 1
                  From ��Ա�� R, ��Ա����˵�� Y
                  Where r.Id = y.��Աid And r.���� = a.������Ա And y.��Ա���� = 'ҽ��') Or ǩ��id Is Not Null Or a.������Ա <> v_����ҽ��);
          If Nvl(v_Count, 0) = 0 Then
            v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"��û�е���ǩ��������ͨ��У�ԡ�';
            Raise Err_Custom;
          End If;
        End If;
        Exit;
      End Loop;
    End If;
  End If;

  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա��� := ����Ա���_In;
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --��Ϊ����ͬʱ���¿�->�Զ�У��->�����Զ�ֹͣ,��˷ֱ�-2,-1��
  Select Sysdate - 1 / 60 / 60 / 24 Into v_Date From Dual;

  Update ����ҽ����¼
  Set ҽ��״̬ = ״̬_In, У�Ի�ʿ = v_��Ա����, У��ʱ�� = У��ʱ��_In
  Where ID = ҽ��id_In Or ���id = ҽ��id_In;

  Insert Into ����ҽ��״̬
    (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
    Select ID, ״̬_In, v_��Ա����, v_Date, У��˵��_In From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;

  --У��ͨ��ʱ����������
  If ״̬_In = 3 Then
    --�Զ�У��ʱ���Զ���дȱʡ�ļƼ�����
    If Nvl(�Զ�У��_In, 0) = 1 Then
      --1.��۵ļƼ���Ŀ,�������޼۲�Ϊ0,��ȱʡΪ����޼�,���򲻼���;�����ֹ��Ƽ�.
      --2.���ڷ�ҩ��ҩƷ����������δ��ִ�п���,����ʱ��ȡȱʡ��,�����ֹ����á�
      Select Min(վ��) Into v_վ�� From ���ű� Where ID = n_���˿���id;
    
      v_Pricegrade := Zl_Get_Pricegrade(v_վ��, v_����id, v_��ҳid);
      v_ҩƷ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
      v_Pricegrade := Substr(v_Pricegrade, Instr(v_Pricegrade, '|') + 1);
      v_���ĵȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
      v_Pricegrade := Substr(v_Pricegrade, Instr(v_Pricegrade, '|') + 1);
      v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
      For r_Price In c_Price Loop
        --ȡ(����)ҽ���Ĺ���Ͳ���,�ɼ���ʽ�Լ�����Ŀ��Ϊ׼
        v_����id := Null;
        If r_Price.������� = 'E' And r_Price.�������� = '6' Then
          Begin
            Select c.����id
            Into v_����id
            From ����ҽ����¼ A, ������ĿĿ¼ B, ��Ѫ������ C
            Where a.������Ŀid = b.Id And b.�Թܱ��� = c.���� And a.���id = r_Price.Id And Rownum = 1;
          Exception
            When Others Then
              Null;
          End;
        Elsif r_Price.������� = 'C' And r_Price.�Թܱ��� Is Not Null Then
          Begin
            Select ����id Into v_����id From ��Ѫ������ Where ���� = r_Price.�Թܱ���;
          Exception
            When Others Then
              Null;
          End;
        End If;
      
        --�жϴ�������Թܷ��õ���ȡ
        If (Nvl(r_Price.�շѷ�ʽ, 0) = 1 And r_Price.�շ���� = '4' And r_Price.�շ���Ŀid = Nvl(v_����id, 0) Or
           Not (Nvl(r_Price.�շѷ�ʽ, 0) = 1 And r_Price.�շ���� = '4' And Nvl(v_����id, 0) <> 0)) Then
          Insert Into ����ҽ���Ƽ�
            (ҽ��id, �շ�ϸĿid, ����, ����, ����, ִ�п���id, ��������, �շѷ�ʽ)
          Values
            (r_Price.Id, r_Price.�շ���Ŀid, r_Price.�շ�����, r_Price.����, r_Price.������Ŀ, Null, r_Price.��������, r_Price.�շѷ�ʽ);
        End If;
      End Loop;
    End If;
  
    --����¼�������ҽ�����Ϊֹͣ
    If Nvl(v_��Ч, 0) = 1 And v_������Ŀid Is Null Then
      Update ����ҽ����¼
      Set ҽ��״̬ = 8, ͣ��ʱ�� = У��ʱ��_In, ͣ��ҽ�� = v_����ҽ��
      Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    
      Insert Into ����ҽ��״̬
        (ҽ��id, ��������, ������Ա, ����ʱ��)
        Select ID, 8, v_��Ա����, Sysdate From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    End If;
  
    --�ж��Ƿ���������Ҫִ��
    v_����ִ�� := zl_GetSysParameter(288);
    If v_����ִ�� = 1 And v_������Ŀid Is Null Then
      Insert Into ����ҽ������
        (ҽ��id, ���ͺ�, ��¼����, NO, ��¼���, ��������, ������, ����ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬, �״�ʱ��, ĩ��ʱ��)
      Values
        (ҽ��id_In, Nextno('10', '0', '', '1'), '2', Nextno('14', '0', '', '1'), '1', '1', v_��Ա����, Sysdate, '0', v_ִ�п���id,
         '0', Sysdate, Sysdate);
    End If;
  
    v_Parͣ�� := zl_GetSysParameter(271);
  
    --��ͬһ�Զ�ֹͣ�������еĲ�������ҽ��ֹͣ(�����δֹͣ)
    For r_Exclude In c_Exclude Loop
      Select Decode(Sign(r_Exclude.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Exclude.��ʼִ��ʱ��, v_��ʼʱ��)
      Into v_��ֹʱ��
      From Dual;
      Select Decode(Sign(r_Exclude.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Exclude.ִ����ֹʱ��, v_��ʼʱ��)
      Into v_��ֹʱ��
      From Dual;
      If v_Parͣ�� = '1' Then
        v_Temp := '�Զ�ֹͣ��ҽ�����⡣';
      Else
        v_Temp := Null;
      End If;
      Zl_����ҽ����¼_ֹͣ(r_Exclude.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1, 0, 0, Null, Null, v_Temp);
      v_Stopadviceids := v_Stopadviceids || ',' || r_Exclude.ҽ��id;
    End Loop;
  
    --��һЩ����ҽ���Ĵ���
    If v_������� = 'H' And v_�������� = '1' And Nvl(v_��Ч, 0) = 0 Then
      --У�Ի���ȼ�ʱ,ͬ�����Ĳ��˻���ȼ�
      If Nvl(v_Ӥ��, 0) = 0 Then
        --���˵�ǰӦ��������סԺ״̬
        v_Temp := Null;
        Begin
          Select Decode(״̬, 1, '�ȴ����', 2, '����ת��', 3, '��Ԥ��Ժ', Null)
          Into v_Temp
          From ������ҳ
          Where ����id = v_����id And ��ҳid = v_��ҳid;
        Exception
          When Others Then
            Null;
        End;
        If v_Temp Is Not Null Then
          v_Error := '���˵�ǰ����' || v_Temp || '״̬,ҽ��"' || v_ҽ������ || '"����ͨ��У�ԡ�';
          Raise Err_Custom;
        End If;
      
        Begin
          --�����շѶ��մ�����ǰҽ���Ƽ۱�û����д
          --δ����ʱ,��������ͬʱ,�������ж��ʱ,ֻȡһ����
          Select a.�շ���Ŀid
          Into v_����ȼ�id
          From �����շѹ�ϵ A, �շ���ĿĿ¼ B
          Where a.�շ���Ŀid = b.Id And b.��� = 'H' And Nvl(b.��Ŀ����, 0) <> 0 And a.������Ŀid = v_������Ŀid And Rownum = 1 And
                Not Exists
           (Select 1 From ������ҳ Where ����id = v_����id And ��ҳid = v_��ҳid And ����ȼ�id = a.�շ���Ŀid);
        Exception
          When Others Then
            Null;
        End;
      End If;
    
      --�䶯��¼��ʱ������룬�Ա���˲���ʱ����ͬһ���ֵ�У�ԡ�ֹͣ�Ȳ���
      v_��ʼʱ�� := To_Date(To_Char(v_��ʼʱ��, 'yyyy-mm-dd hh24:mi') || To_Char(Sysdate, 'ss'), 'yyyy-mm-dd hh24:mi:ss');
      If v_����ȼ�id Is Not Null Then
        Zl_���˱䶯��¼_Nurse(v_����id, v_��ҳid, v_����ȼ�id, v_��ʼʱ��, v_��Ա���, v_��Ա����);
      End If;
    
      --��ֹͣ��������ȼ�ҽ��(����ȼ�Ӧ�ö�Ϊ"������"����,��ֻ��һ��δͣ)
      For r_Nurse In c_Nurse Loop
        Select Decode(Sign(r_Nurse.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Nurse.��ʼִ��ʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Select Decode(Sign(r_Nurse.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Nurse.ִ����ֹʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        If v_Parͣ�� = '1' Then
          v_Temp := '�Զ�ֹͣ������ȼ���';
        Else
          v_Temp := Null;
        End If;
        Zl_����ҽ����¼_ֹͣ(r_Nurse.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1, 0, 0, Null, Null, v_Temp);
        Zl_����ҽ����¼_ȷ��ֹͣ(r_Nurse.ҽ��id, v_��ֹʱ��, v_��Ա����, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Nurse.ҽ��id;
      End Loop;
    Elsif v_������� = 'Z' And v_�������� In ('9', '10') And Nvl(v_��Ч, 0) = 0 And Nvl(v_Ӥ��, 0) = 0 Then
      --���ز�Σҽ����9-����;10-��Σ
      --ֹͣ��ͬҽ��
      For r_Patistate In c_Patistate Loop
        Select Decode(Sign(r_Patistate.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Patistate.��ʼִ��ʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Select Decode(Sign(r_Patistate.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Patistate.ִ����ֹʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        If v_Parͣ�� = '1' Then
          If v_�������� = '9' Then
            v_Temp := '�Զ�ֹͣ������ҽ����';
          Else
            v_Temp := '�Զ�ֹͣ����Σҽ����';
          End If;
        Else
          v_Temp := Null;
        End If;
        Zl_����ҽ����¼_ֹͣ(r_Patistate.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1, 0, 0, Null, Null, v_Temp);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patistate.ҽ��id;
      End Loop;
    
      b_Message.Zlhis_Patient_005(v_����id, v_��ҳid);
    
      --��������䶯
      Open c_Oldinfo; --�����ڴ���֮ǰ�ȴ�
      Fetch c_Oldinfo
        Into r_Oldinfo;
      Open c_Endinfo;
      Fetch c_Endinfo
        Into r_Endinfo;
      If c_Endinfo%RowCount = 0 Then
        Close c_Endinfo;
        v_Error := 'δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
        Raise Err_Custom;
      End If;
      Select Count(*)
      Into v_Count
      From ���˱䶯��¼
      Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� Is Null And ��ֹʱ�� Is Null;
      If v_Count > 0 Then
        v_Error := '���˵�ǰ����ת��״̬�����Ȱ���ת��ȷ�ϻ���ȡ��ת��״̬��';
        Raise Err_Custom;
      End If;
    
      Update ������ҳ
      Set ��ǰ���� = Decode(v_��������, '9', '��', '10', 'Σ')
      Where ����id = v_����id And ��ҳid = v_��ҳid;
    
      --ȡ���ϴα䶯
      If r_Oldinfo.��ֹʱ�� Is Not Null Then
        v_�䶯��ֹʱ�� := r_Oldinfo.��ֹʱ��;
        v_�䶯��ֹԭ�� := r_Oldinfo.��ֹԭ��;
        v_�䶯��ֹ��Ա := r_Oldinfo.��ֹ��Ա;
        --ȡ���ϴα䶯
        Update ���˱䶯��¼
        Set ��ֹʱ�� = v_��ʼʱ��, ��ֹԭ�� = 13, ��ֹ��Ա = v_��Ա����, �ϴμ���ʱ�� = Null
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� = v_�䶯��ֹʱ�� And ��ֹԭ�� = v_�䶯��ֹԭ��;
        --���½����ļ�¼�����ֹͣ����������ɾ���ϴμ���ʱ��
        Update ���˱䶯��¼
        Set ���� = Decode(v_��������, '9', '��', '10', 'Σ'), �ϴμ���ʱ�� = Null
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� > v_��ʼʱ��;
      Else
        Update ���˱䶯��¼
        Set ��ֹʱ�� = v_��ʼʱ��, ��ֹԭ�� = 13, ��ֹ��Ա = v_��Ա����,
            �ϴμ���ʱ�� = Decode(Sign(Nvl(�ϴμ���ʱ��, v_��ʼʱ��) - v_��ʼʱ��), 1, Null, �ϴμ���ʱ��)
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null;
      End If;
    
      While c_Oldinfo%Found Loop
        Insert Into ���˱䶯��¼
          (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ���Ӵ�λ, ����id, ����id, ����ȼ�id, ��λ�ȼ�id, ����, ���λ�ʿ, ����ҽʦ, ����ҽʦ, ����ҽʦ, ����, ����Ա���, ����Ա����,
           ��ֹʱ��, ��ֹԭ��, ��ֹ��Ա)
        Values
          (���˱䶯��¼_Id.Nextval, v_����id, v_��ҳid, v_��ʼʱ��, 13, r_Oldinfo.���Ӵ�λ, r_Oldinfo.����id, r_Oldinfo.����id,
           r_Oldinfo.����ȼ�id, r_Oldinfo.��λ�ȼ�id, r_Oldinfo.����, r_Oldinfo.���λ�ʿ, r_Oldinfo.����ҽʦ, r_Oldinfo.����ҽʦ,
           r_Oldinfo.����ҽʦ, Decode(v_��������, '9', '��', '10', 'Σ'), v_��Ա���, v_��Ա����, v_�䶯��ֹʱ��, v_�䶯��ֹԭ��, v_�䶯��ֹ��Ա);
      
        Fetch c_Oldinfo
          Into r_Oldinfo;
      End Loop;
    
      Close c_Oldinfo;
      Close c_Endinfo;
    Elsif v_������� = 'Z' And v_�������� = '12' And Nvl(v_��Ч, 0) = 0 And Nvl(v_Ӥ��, 0) = 0 Then
      --��¼�������ҽ��������
      For r_Patiio In c_Patiio Loop
        Select Decode(Sign(r_Patiio.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Patiio.��ʼִ��ʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Select Decode(Sign(r_Patiio.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Patiio.ִ����ֹʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Zl_����ҽ����¼_ֹͣ(r_Patiio.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patiio.ҽ��id;
      End Loop;
    Elsif (v_������� = 'Z' And v_�������� In ('3', '4', '5', '6', '11', '14') And
          (v_�������� <> '14' Or v_�������� = '14' And v_ִ�б�� = 1)) Or (v_������� = 'F' And v_ִ�б�� = 1) Then
      v_Count := 0;
      If v_�������� = '4' Or v_�������� = '14' Or v_������� = 'F' Then
        --��������ǰУ��ʱ��ͬ�Ĵ���
        If Nvl(v_Ӥ��, 0) = 0 Then
          v_Count := 1;
        End If;
      Else
        --�⼸������ҽ����У����ֹͣҽ�����¼ӵ����ݣ������뷢������ͬ�Ĵ���
        v_Count := 1;
        If Nvl(v_Ӥ��, 0) = 0 Then
          v_Ӥ�� := -1;
        Else
          v_Ӥ�� := Nvl(v_Ӥ��, 0);
        End If;
      End If;
      If v_Count = 1 Then
        If v_������� = 'F' And v_ִ�б�� = 1 Then
          --����������(ȡ��)ֹͣ
          v_��ʼʱ�� := Trunc(To_Date(v_�걾��λ, 'yyyy-mm-dd hh24:mi:ss'));
        End If;
      
        --��������ҽ��У��ʱֹͣǰ��ĳ���,��ҽ����ʼʱ��ֹ��3-ת��;4-����;5-��Ժ;6-תԺ,11-����,14-��ǰ
        For r_Needstop In c_Needstop(v_����id, v_��ҳid, v_Ӥ��, v_��ʼʱ��) Loop
          Select Decode(Sign(��ʼִ��ʱ�� - v_��ʼʱ��), 1, ��ʼִ��ʱ��, v_��ʼʱ��)
          Into v_ֹͣʱ��
          From ����ҽ����¼
          Where ID = r_Needstop.Id;
          Update ����ҽ����¼
          Set ҽ��״̬ = 8, ִ����ֹʱ�� = v_ֹͣʱ��, ͣ��ʱ�� = У��ʱ��_In, ͣ��ҽ�� = v_����ҽ��
          Where ID = r_Needstop.Id;
        
          Insert Into ����ҽ��״̬
            (ҽ��id, ��������, ������Ա, ����ʱ��)
            Select ID, 8, v_��Ա����, У��ʱ��_In From ����ҽ����¼ Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --��ֹͣδȷ�ϵĳ���,��ֹʱ����ҽ����ʼ���,��ǰ����ֹʱ��(ͬʱ�������ҽ�������)
        For r_Havestop In c_Havestop(v_����id, v_��ҳid, v_Ӥ��, v_��ʼʱ��) Loop
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Sign(��ʼִ��ʱ�� - v_��ʼʱ��), 1, ��ʼִ��ʱ��, v_��ʼʱ��), ͣ��ʱ�� = У��ʱ��_In, ͣ��ҽ�� = v_����ҽ��
          Where ID = r_Havestop.Id;
        
          --���޸�ֹͣҽ���Ĳ�����Ա����Ϊֹͣʱ��ҽ�������ѽ��е���ǩ��
          Update ����ҽ��״̬ Set ����ʱ�� = У��ʱ��_In Where ҽ��id = r_Havestop.Id And �������� = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --�����ڱ���ҽ��(û��ִ�У����ͣ����ı��δ�ã�
        Update ����ҽ����¼
        Set ִ�б�� = -1
        Where ����id = v_����id And ��ҳid = v_��ҳid And ҽ����Ч = 0 And ִ��Ƶ�� = '��Ҫʱ' And �ϴ�ִ��ʱ�� Is Null And ҽ��״̬ In (3, 5, 6, 7) And
              ִ�б�� <> -1;
        --�����תԺת��������Ժҽ��ͬʱ������ʱ����ҽ����
        If v_�������� In ('3', '5', '6', '11') Then
          Update ����ҽ����¼
          Set ִ�б�� = -1
          Where ����id = v_����id And ��ҳid = v_��ҳid And ҽ����Ч = 1 And ִ��Ƶ�� = '��Ҫʱ' And ҽ��״̬ = 3 And ִ�б�� <> -1;
        End If;
      End If;
    Elsif v_������� = 'Z' And v_�������� = '2' Then
      --�����۲����´���Ժ֪ͨ;
      --ԤԼ�Ǽǵ�������1.��ǰ��ԤԼ,2.��ǰ���������۲��ˣ���ԺʱҲ������Ϊ��Ҫ��ԤԼ,��Ժ����ʱ����˱����Ժ����ܽ��գ�
      Select Count(*) Into v_Count From ������ҳ Where ����id = v_����id And Nvl(��ҳid, 0) = 0;
      If v_Count = 0 Then
        Select Count(*) Into v_Count From ������ҳ Where ����id = v_����id And ��ҳid = v_��ҳid And �������� <> 1;
      End If;
      If v_Count = 0 Then
        Open c_Pati(v_����id);
        Fetch c_Pati
          Into r_Pati;
        Close c_Pati;
      
        v_��Ժ��ʽ := Null;
        If v_������־ = 1 Then
          v_��Ժ��ʽ := '����';
        End If;
      
        Zl_��Ժ������ҳ_Insert(1, 0, r_Pati.����id, r_Pati.סԺ��, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�, r_Pati.��������,
                         r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���, r_Pati.���֤��, r_Pati.�����ص�,
                         r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ, r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ,
                         r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ, r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������,
                         r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������, r_Pati.��������, v_ִ�п���id, Null, Null, v_��Ժ��ʽ, Null, Null,
                         v_����ҽ��, r_Pati.����, r_Pati.����, v_��ʼʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null,
                         Null, r_Pati.����, v_��Ա���, v_��Ա����, 0, Null, Null, 0);
      End If;
    End If;
    --ҽ��ֹͣ��Ϣ�Ĵ���
    If v_Stopadviceids Is Not Null Then
      v_Stopadviceids := Substr(v_Stopadviceids, 2);
      b_Message.Zlhis_Cis_002(v_����id, v_��ҳid, Null, v_Stopadviceids);
      Select Max(a.Id)
      Into n_���
      From ����ҽ����¼ A
      Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.ҽ����Ч = 0 And a.ҽ��״̬ = 8 And
            Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
      If n_��� Is Not Null Then
        Select Max(a.Id)
        Into n_Adviceid
        From ����ҽ����¼ A
        Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.������־ = 1 And a.ҽ����Ч = 0 And
              a.ҽ��״̬ = 8 And Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
        If n_Adviceid Is Not Null Then
          n_Adviceid := n_���;
          Select Nvl(Max(0), 2)
          Into n_���
          From ҵ����Ϣ�嵥 A
          Where a.����id = v_����id And a.����id = v_��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.���ȳ̶� = 2 And a.�Ƿ����� = 0;
        Else
          Select Nvl(Max(0), 1)
          Into n_���
          From ҵ����Ϣ�嵥 A
          Where a.����id = v_����id And a.����id = v_��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.�Ƿ����� = 0;
        End If;
        If n_��� > 0 Then
          For R In (Select a.�������� As ����, a.��Ժ����id As ����id, a.��ǰ����id As ����id
                    From ������ҳ A
                    Where a.����id = v_����id And a.��ҳid = v_��ҳid) Loop
            Zl_ҵ����Ϣ�嵥_Insert(v_����id, v_��ҳid, r.����id, r.����id, r.����, '����ֹͣҽ����', '0010', 'ZLHIS_CIS_002', n_Adviceid, n_���,
                             0, Null, r.����id);
          End Loop;
        End If;
      End If;
    End If;
  End If;

  --����ִ��ҽ��У����Ϣ
  For R In (Select a.Id, a.����id, a.��ҳid
            From ����ҽ����¼ A
            Where (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And Exists
             (Select 1 From ��������˵�� B Where b.����id = a.ִ�п���id And b.�������� = '����')) Loop
    b_Message.Zlhis_Cis_012(r.����id, r.��ҳid, r.Id);
  End Loop;
  --У������ҽ��
  If v_������� = 'Z' And v_�������� = '4' Then
    b_Message.Zlhis_Cis_004(v_����id, v_��ҳid, ҽ��id_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_У��;
/

--122609:������,2018-03-08,����ƽ̨��Ϣ���
Create Or Replace Procedure Zl_����ҽ������_Insert
(
  ҽ��id_In     In ����ҽ������.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ������.���ͺ�%Type,
  ��¼����_In   In ����ҽ������.��¼����%Type,
  No_In         In ����ҽ������.No%Type,
  ��¼���_In   In ����ҽ������.��¼���%Type,
  ��������_In   In ����ҽ������.��������%Type,
  �״�ʱ��_In   In ����ҽ������.�״�ʱ��%Type,
  ĩ��ʱ��_In   In ����ҽ������.ĩ��ʱ��%Type,
  ����ʱ��_In   In ����ҽ������.����ʱ��%Type,
  ִ��״̬_In   In ����ҽ������.ִ��״̬%Type,
  ִ�в���id_In In ����ҽ������.ִ�в���id%Type,
  �Ʒ�״̬_In   In ����ҽ������.�Ʒ�״̬%Type,
  First_In      In Number := 0,
  ��������_In   In ����ҽ������.��������%Type := Null,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null,
  ��ҩ��_In     In δ��ҩƷ��¼.��ҩ��%Type := Null,
  �������_In   In ����ҽ������.�������%Type := Null,
  �ֽ�ʱ��_In   In Varchar2 := Null,
  ԭҺƤ��_In   In Varchar2 := Null
  --���ܣ���д����ҽ�����ͼ�¼
  --������
  --      ҽ��id_In=Ҫ���͵�ÿ��ҽ��ID
  --      First_IN=��ʾ�Ƿ�һ��ҽ���ĵ�һҽ����,�Ա㴦��ҽ���������(���ҩ,�䷽�ĵ�һ��,��Ϊ��ҩ;��,�䷽�巨,�÷�����Ϊ����������)
  --      ��������_IN,�״�ʱ��_IN,ĩ��ʱ��_IN:��"������"����,����д��������,����д��ĩ��ʱ��(���ڻ���)��
  --      �������_In,סԺ�������͵��������ʱ����дΪ1����Ϊ��¼������2����������סԺ���ʣ��������������ա�
  --      ԴҺƤ��_In ԭҺƤ��ҽ��ID�������7107/bug115972���ڹ���ҩƷҽ���к�Ƥ��ҽ���С������ֶ�Ϊ ����ҽ������.�걾�������� ����ҩƷ�е�ҽ��IDֵ
  --      ��ʽ��1ҽ��ID,2ҽ��ID ǰ��һ��ΪƤ��ҽ����ҽ��ID���ڶ���ΪҩƷ��ҽ����ҽ��ID
) Is
  --�������˼�ҽ��(һ��ҽ���е�һ��)�����Ϣ���α�
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.���, a.����id, a.��ҳid, a.Ӥ��, a.����, a.���˿���id, c.��������, a.�������, a.ҽ����Ч, a.ҽ��״̬, a.ҽ������,
           a.����ҽ��, a.����ʱ��, a.��ʼִ��ʱ��, a.�ϴ�ִ��ʱ��, a.ִ����ֹʱ��, a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.��������id, a.�걾��λ, a.ִ�п���id,
           a.���id, a.������Ŀid, a.�Һŵ�
    From ����ҽ����¼ A, ������ĿĿ¼ C
    Where a.������Ŀid = c.Id And a.Id = ҽ��id_In;
  r_Advice c_Advice%RowType;

  --��������(Ӥ��)������δͣ����(���䷽����),Ӥ������-1��ʾ������
  Cursor c_Needstop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select a.Id, a.�������, b.��������, b.ִ��Ƶ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id(+) And a.����id = v_����id And a.��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� < v_Stoptime
    Order By a.���;
  --��������(Ӥ��)����ͣ��δȷ�ϵĳ���,��ִֹ��ʱ����ָ��ʱ��֮��,Ӥ������-1��ʾ������
  Cursor c_Havestop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select ID
    From ����ҽ����¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And Nvl(ҽ����Ч, 0) = 0 And
          ҽ��״̬ = 8 And ִ����ֹʱ�� > v_Stoptime And ��ʼִ��ʱ�� < v_Stoptime
    Order By ���;

  --������ʱ����
  v_Ӥ��       ����ҽ����¼.Ӥ��%Type;
  v_������     Number(1); --�Ƿ�����Գ���
  v_Autostop   Number(1);
  v_Date       Date;
  v_Temp       Varchar2(255);
  v_��Ա���   ��Ա��.���%Type;
  v_��Ա����   ��Ա��.����%Type;
  v_ֹͣʱ��   ����ҽ����¼.����ʱ��%Type;
  n_ִ��״̬   ����ҽ������.ִ��״̬%Type;
  d_��ʼʱ��   ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_Count      Number;
  n_Ƥ�Ա��   ����ҽ������.ҽ��id%Type;
  n_Ƥ��ҽ��id ����ҽ������.ҽ��id%Type;

  v_Stopadviceids ����ҽ����¼.ҽ������%Type;
  n_Adviceid      ����ҽ����¼.����id%Type;
  n_���          Number(18);
  v_Error         Varchar2(255);
  Err_Custom Exception;
Begin
  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա��� := ����Ա���_In;
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;
  --����״�ʱ��Ϊ�������뿪ʼִ��ʱ��
  If �״�ʱ��_In Is Null Or �ֽ�ʱ��_In Is Null Or ĩ��ʱ��_In Is Null Then
    Select ��ʼִ��ʱ�� Into d_��ʼʱ�� From ����ҽ����¼ Where ID = ҽ��id_In;
  End If;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  Close c_Advice;

  --��һ��ҽ���ĵ�һ��ʱ����ҽ������
  If Nvl(First_In, 0) = 1 Then
    --�����������
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.ҽ��״̬, 0) = 4 Then
      --���Ҫ���͵�ҽ���Ƿ�����
      v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ������������ϡ�' || Chr(13) || Chr(10) ||
                 '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
      Raise Err_Custom;
    End If;
  
    If Nvl(r_Advice.ҽ����Ч, 0) = 0 Then
      --����������ҩ����,�䷽����,��ҩ"��ѡƵ��"����,��ҩ"������"����
    
      --��鳤���Ƿ��ѱ�����
      If r_Advice.�ϴ�ִ��ʱ�� Is Not Null Then
        If r_Advice.�ϴ�ִ��ʱ�� >= �״�ʱ��_In Then
          v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ��������˷��͡�' || Chr(13) || Chr(10) ||
                     '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
          Raise Err_Custom;
        End If;
      End If;
    
      --��鳤������ǰ�Ƿ��ѱ��Զ�ֹͣ(������)
      If r_Advice.ִ����ֹʱ�� Is Not Null Then
        If �״�ʱ��_In > r_Advice.ִ����ֹʱ�� Then
          v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ���ֹͣ��' || Chr(13) || Chr(10) ||
                     '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
          Raise Err_Custom;
        End If;
      End If;
    Elsif Nvl(r_Advice.ҽ��״̬, 0) In (8, 9) Then
      --���������䷽����
    
      --����Ƿ��ѱ�����(��������ԭ���Զ�ֹͣ)
      v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ��������˷��͡�' || Chr(13) || Chr(10) ||
                 '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
      Raise Err_Custom;
    End If;
  
    --���ͺ��ҽ������
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.ҽ����Ч, 0) = 0 Then
      --����ҽ��:�����ϴ�ִ��ʱ��
      Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = ĩ��ʱ��_In Where ID = r_Advice.��id Or ���id = r_Advice.��id;
    
      --�ж��Ƿ�����Գ���
      v_������ := 0;
      If r_Advice.ִ��ʱ�䷽�� Is Null And (Nvl(r_Advice.Ƶ�ʴ���, 0) = 0 Or Nvl(r_Advice.Ƶ�ʼ��, 0) = 0 Or r_Advice.�����λ Is Null) Then
        v_������ := 1;
      End If;
    
      --Ԥ������ֹʱ����δֹͣ���Զ�ֹͣ
      If r_Advice.ִ����ֹʱ�� Is Not Null And Nvl(r_Advice.ҽ��״̬, 0) Not In (8, 9) Then
        v_Autostop := 0;
        If v_������ = 1 Then
          --��ҩ"������"����
          If Trunc(ĩ��ʱ��_In) = Trunc(r_Advice.ִ����ֹʱ�� - 1) Then
            v_Autostop := 1; --��ֹ���첻ִ��
          End If;
        Elsif Zl_Advicenexttime(ҽ��id_In) > r_Advice.ִ����ֹʱ�� Then
          --��ҩ�������ҩ"��ѡƵ��"����
          v_Autostop := 1; --����ǵ���,������ִ��һ��
        End If;
      
        If v_Autostop = 1 Then
          Update ����ҽ����¼
          Set ҽ��״̬ = 8, ͣ��ʱ�� = ĩ��ʱ��_In, ͣ��ҽ�� = r_Advice.����ҽ��
          Where ID = r_Advice.��id Or ���id = r_Advice.��id;
          v_Temp := zl_GetSysParameter(271);
          If v_Temp = '1' Then
            v_Temp := '�Զ�ֹͣ��Ԥ��ֹͣʱ�䡣';
          Else
            v_Temp := Null;
          End If;
          Insert Into ����ҽ��״̬
            (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
            Select ID, 8, r_Advice.����ҽ��, ����ʱ��_In, v_Temp
            From ����ҽ����¼
            Where ID = r_Advice.��id Or ���id = r_Advice.��id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Advice.��id;
        End If;
      End If;
    Else
      --����ֹͣ��
      --סԺҽ������ʱ�Զ�У�ԡ�ֹͣ��У������Sysdateȡ��,Ϊ�����ظ�,ֹͣʱ��ҲȡSysdate
      Select Sysdate Into v_Date From Dual;
      Update ����ҽ����¼
      Set ҽ��״̬ = 8, ִ����ֹʱ�� = ĩ��ʱ��_In,
          --Ϊһ��������ʱû��
          �ϴ�ִ��ʱ�� = ĩ��ʱ��_In,
          --Ϊһ��������ʱû��
          ͣ��ʱ�� = v_Date,
          --����ʱ��_IN,
          ͣ��ҽ�� = r_Advice.����ҽ��
      Where ID = r_Advice.��id Or ���id = r_Advice.��id;
    
      Insert Into ����ҽ��״̬
        (ҽ��id, ��������, ������Ա, ����ʱ��)
        Select ID, 8, v_��Ա����, v_Date --����ʱ��_IN
        From ����ҽ����¼
        Where ID = r_Advice.��id Or ���id = r_Advice.��id;
    End If;
  
    --����ҽ���Ĵ���
    ---------------------------------------------------------------------------------------
    If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' Then
      --(1-����;2-סԺ;)3-ת��;4-����(������);5-��Ժ;6-תԺ,7-����,11-����
    
      --��������ҽ��Ҫ�Զ�ֹͣ���˸�ҽ��֮ǰ(��ʱ����)����δͣ�ĳ���
      If r_Advice.�������� In ('3', '5', '6', '11') Then
        If Nvl(r_Advice.Ӥ��, 0) = 0 Then
          v_Ӥ�� := -1;
        Else
          v_Ӥ�� := Nvl(r_Advice.Ӥ��, 0);
        End If;
        For r_Needstop In c_Needstop(r_Advice.����id, r_Advice.��ҳid, v_Ӥ��, r_Advice.��ʼִ��ʱ��) Loop
          Select Decode(Sign(��ʼִ��ʱ�� - r_Advice.��ʼִ��ʱ��), 1, ��ʼִ��ʱ��, r_Advice.��ʼִ��ʱ��)
          Into v_ֹͣʱ��
          From ����ҽ����¼
          Where ID = r_Needstop.Id;
          Update ����ҽ����¼
          Set ҽ��״̬ = 8, ִ����ֹʱ�� = v_ֹͣʱ��, ͣ��ʱ�� = ����ʱ��_In, ͣ��ҽ�� = r_Advice.����ҽ��
          Where ID = r_Needstop.Id;
        
          Insert Into ����ҽ��״̬
            (ҽ��id, ��������, ������Ա, ����ʱ��)
            Select ID, 8, v_��Ա����, ����ʱ��_In From ����ҽ����¼ Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --��ֹͣδȷ�ϵĳ���,��ֹʱ����ҽ����ʼ���,��ǰ����ֹʱ��(ͬʱ�������ҽ�������)
        For r_Havestop In c_Havestop(r_Advice.����id, r_Advice.��ҳid, v_Ӥ��, r_Advice.��ʼִ��ʱ��) Loop
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Sign(��ʼִ��ʱ�� - r_Advice.��ʼִ��ʱ��), 1, ��ʼִ��ʱ��, r_Advice.��ʼִ��ʱ��), ͣ��ʱ�� = ����ʱ��_In,
              ͣ��ҽ�� = r_Advice.����ҽ��
          Where ID = r_Havestop.Id;
        
          --���޸�ֹͣҽ���Ĳ�����Ա����Ϊֹͣʱ��ҽ�������ѽ��е���ǩ��
          Update ����ҽ��״̬ Set ����ʱ�� = ����ʱ��_In Where ҽ��id = r_Havestop.Id And �������� = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --�����ڱ���ҽ��(û��ִ�У����ͣ����ı��δ�ã�,ͬʱ��������
        Update ����ҽ����¼
        Set ִ�б�� = -1
        Where ����id = r_Advice.����id And ��ҳid = r_Advice.��ҳid And
              (ҽ����Ч = 0 And ִ��Ƶ�� = '��Ҫʱ' And �ϴ�ִ��ʱ�� Is Null And ҽ��״̬ In (3, 5, 6, 7) Or
              ҽ����Ч = 1 And ִ��Ƶ�� = '��Ҫʱ' And ҽ��״̬ = 3) And ִ�б�� <> -1;
      End If;
    
      --��������⴦��
      If Nvl(r_Advice.Ӥ��, 0) = 0 Then
        If r_Advice.�������� = '3' And ִ�в���id_In Is Not Null And r_Advice.���˿���id Is Not Null And
           Nvl(r_Advice.���˿���id, 0) <> Nvl(ִ�в���id_In, 0) Then
          --ת��ҽ��,�����˵Ǽ�ת�Ƶ�"ִ�п���ID"(��Ժ�����ҵ�ǰ������ת����Ҳ�ͬ�Ŵ���)
          Select Count(1)
          Into v_Temp
          From ������ҳ
          Where ����id = r_Advice.����id And ��ҳid = r_Advice.��ҳid And ��Ժ����id <> ִ�в���id_In;
          If v_Temp = '1' Then
            Zl_���˱䶯��¼_Change(r_Advice.����id, r_Advice.��ҳid, ִ�в���id_In, v_��Ա���, v_��Ա����);
          End If;
        Elsif r_Advice.�������� In ('5', '6', '11') Then
          --��Ժ��תԺ������ҽ��,�����˱��ΪԤ��Ժ
          Begin
            Select ��ʼʱ��
            Into v_Date
            From ���˱䶯��¼
            Where ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null And ����id = r_Advice.����id And ��ҳid = r_Advice.��ҳid;
          Exception
            When Others Then
              v_Date := To_Date('1900-01-01', 'YYYY-MM-DD');
          End;
          If r_Advice.��ʼִ��ʱ�� <= v_Date Then
            v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ŀ�ʼʱ��Ӧ���ڸò����ϴα䶯ʱ�� ' || To_Char(v_Date, 'YYYY-MM-DD HH24:Mi') || ' ��';
            Raise Err_Custom;
          End If;
          Zl_���˱䶯��¼_Preout(r_Advice.����id, r_Advice.��ҳid, r_Advice.��ʼִ��ʱ��);
        End If;
      Else
        If r_Advice.�������� = '11' Then
          Update ������������¼
          Set ����ʱ�� = r_Advice.��ʼִ��ʱ��
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And Nvl(���, 0) = Nvl(r_Advice.Ӥ��, 0);
        End If;
      End If;
    End If;
    --12Сʱδִ�еı�����������Ϊ���δ��
    If r_Advice.ҽ����Ч = 1 Then
      Update ����ҽ����¼
      Set ִ�б�� = -1
      Where ����id = r_Advice.����id And ��ҳid = r_Advice.��ҳid And ִ�б�� <> -1 And ҽ����Ч = 1 And ִ��Ƶ�� = '��Ҫʱ' And
            Sysdate - ��ʼִ��ʱ�� > 0.5 And ҽ��״̬ = 3;
    End If;
  End If;

  --��д���ͼ�¼
  ---------------------------------------------------------------------------------------
  n_ִ��״̬ := ִ��״̬_In;
  If ִ��״̬_In = 1 Then
    v_Temp := zl_GetSysParameter(186);
    If v_Temp = '11' Then
      If r_Advice.������� = 'E' And r_Advice.�������� In ('1', '8') Or r_Advice.������� = 'K' Then
        n_ִ��״̬ := 0;
      End If;
    Elsif v_Temp = '01' Then
      If r_Advice.������� = 'E' And r_Advice.�������� = '1' Then
        n_ִ��״̬ := 0;
      End If;
    Elsif v_Temp = '10' Then
      If r_Advice.������� = 'E' And r_Advice.�������� = '8' Or r_Advice.������� = 'K' Then
        n_ִ��״̬ := 0;
      End If;
    End If;
  End If;

  If ԭҺƤ��_In Is Not Null Then
    v_Count      := Instr(ԭҺƤ��_In, ',');
    n_Ƥ��ҽ��id := Substr(ԭҺƤ��_In, 1, v_Count - 1);
    n_Ƥ�Ա��   := Substr(ԭҺƤ��_In, v_Count + 1);
    Update ����ҽ������ Set �걾�������� = n_Ƥ�Ա�� Where ҽ��id = n_Ƥ��ҽ��id;
  End If;

  Insert Into ����ҽ������
    (ҽ��id, ���ͺ�, ��¼����, NO, ��¼���, ��������, ������, ����ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬, �״�ʱ��, ĩ��ʱ��, ��������, �������, �걾��������)
  Values
    (ҽ��id_In, ���ͺ�_In, ��¼����_In, No_In, ��¼���_In, ��������_In, v_��Ա����, ����ʱ��_In, n_ִ��״̬, ִ�в���id_In, �Ʒ�״̬_In,
     Nvl(�״�ʱ��_In, d_��ʼʱ��), Nvl(ĩ��ʱ��_In, d_��ʼʱ��), ��������_In, �������_In, n_Ƥ�Ա��);

  --�����ͼ��ҽ��ͬ��������ҽ���ļƷ�״̬
  If �Ʒ�״̬_In = 1 And r_Advice.��id <> ҽ��id_In And (r_Advice.������� = 'D' Or r_Advice.������� = 'F') Then
    Update ����ҽ������ Set �Ʒ�״̬ = 1 Where ҽ��id = r_Advice.��id And ���ͺ� = ���ͺ�_In;
  End If;

  --��ҩ�ŵ���д
  If ��ҩ��_In Is Not Null Then
    Update δ��ҩƷ��¼ Set ��ҩ�� = ��ҩ��_In Where NO = No_In And ���� = 9 And ��ҩ�� Is Null;
    Update ҩƷ�շ���¼ Set ��Ʒ�ϸ�֤ = ��ҩ��_In Where NO = No_In And ���� = 9 And ��Ʒ�ϸ�֤ Is Null;
  End If;

  --�Զ���Ϊ��ִ��ʱ����Ҫͬ���������ִ��״̬����˻���״̬
  If ִ��״̬_In = 1 Then
    Zl_����ҽ��ִ��_Finish(ҽ��id_In, ���ͺ�_In, Null, Null, v_��Ա���, v_��Ա����, ִ�в���id_In);
  End If;

  --����ҽ��ִ��ʱ���¼(ֻ��������¼��)
  If Nvl(�ֽ�ʱ��_In, To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi:ss')) Is Not Null Then
    If r_Advice.���id Is Null Then
      Insert Into ҽ��ִ��ʱ��
        (Ҫ��ʱ��, ҽ��id, ���ͺ�)
        Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), ҽ��id_In, ���ͺ�_In
        From Table(f_Str2list(Nvl(�ֽ�ʱ��_In, To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi:ss'))));
    End If;
  End If;

  --������дʱ������д
  If r_Advice.������� = 'F' Then
    --һ������ֻ��һ��
    If r_Advice.���id Is Null Then
      If Not r_Advice.�걾��λ Is Null Then
        v_Date := To_Date(r_Advice.�걾��λ, 'yyyy-mm-dd hh24:mi:ss');
      Else
        v_Date := r_Advice.��ʼִ��ʱ��;
      End If;
      Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, v_Date, v_Date,
                       r_Advice.ִ�п���id);
    End If;
  Elsif r_Advice.������� = 'Z' And r_Advice.�������� = '7' Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id);
  Elsif r_Advice.������� = 'Z' And r_Advice.�������� = '8' Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id);
  Elsif r_Advice.������� = 'Z' And r_Advice.�������� = '11' Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id);
  End If;
  --�������(֪���ļ�������������ŵ���)
  If Instr('C,D,E,F,G,K,L', r_Advice.�������) > 0 Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '֪������', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id, r_Advice.������Ŀid, r_Advice.ҽ������);
  End If;
  --ҽ��ֹͣ��Ϣ�Ĵ���
  If v_Stopadviceids Is Not Null Then
    v_Stopadviceids := Substr(v_Stopadviceids, 2);
    b_Message.Zlhis_Cis_002(r_Advice.����id, r_Advice.��ҳid, Null, v_Stopadviceids);
    Select Max(a.Id)
    Into n_���
    From ����ҽ����¼ A
    Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.ҽ����Ч = 0 And a.ҽ��״̬ = 8 And
          Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
    If n_��� Is Not Null Then
      Select Max(a.Id)
      Into n_Adviceid
      From ����ҽ����¼ A
      Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.������־ = 1 And a.ҽ����Ч = 0 And
            a.ҽ��״̬ = 8 And Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
      If n_Adviceid Is Not Null Then
        Select Nvl(Max(0), 2)
        Into n_���
        From ҵ����Ϣ�嵥 A
        Where a.����id = r_Advice.����id And a.����id = r_Advice.��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.���ȳ̶� = 2 And
              a.�Ƿ����� = 0;
      Else
        n_Adviceid := n_���;
        Select Nvl(Max(0), 1)
        Into n_���
        From ҵ����Ϣ�嵥 A
        Where a.����id = r_Advice.����id And a.����id = r_Advice.��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.�Ƿ����� = 0;
      End If;
      If n_��� > 0 Then
        For R In (Select a.�������� As ����, a.��Ժ����id As ����id, a.��ǰ����id As ����id
                  From ������ҳ A
                  Where a.����id = r_Advice.����id And a.��ҳid = r_Advice.��ҳid) Loop
          Zl_ҵ����Ϣ�嵥_Insert(r_Advice.����id, r_Advice.��ҳid, r.����id, r.����id, r.����, '����ֹͣҽ����', '0010', 'ZLHIS_CIS_002',
                           n_Adviceid, n_���, 0, Null, r.����id);
        End Loop;
      End If;
    End If;
  End If;

  If r_Advice.������� = 'E' And r_Advice.�������� = '6' Then
    --������Ŀ
    b_Message.Zlhis_Cis_016(r_Advice.����id, r_Advice.��ҳid, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id, 2);
  Elsif r_Advice.������� = 'D' And r_Advice.���id Is Null Then
    b_Message.Zlhis_Cis_017(r_Advice.����id, r_Advice.��ҳid, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id, 2);
  Elsif r_Advice.������� = 'F' And r_Advice.���id Is Null Then
    b_Message.Zlhis_Cis_018(r_Advice.����id, r_Advice.��ҳid, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id);
  Elsif r_Advice.������� = 'K' Then
    b_Message.Zlhis_Cis_019(r_Advice.����id, r_Advice.��ҳid, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id);
  Elsif r_Advice.������� = 'Z' Then
    If r_Advice.�������� = '7' Then
      b_Message.Zlhis_Cis_020(r_Advice.����id, r_Advice.��ҳid, ���ͺ�_In, r_Advice.��id);
    Elsif r_Advice.�������� = '8' Then
      b_Message.Zlhis_Cis_021(r_Advice.����id, r_Advice.��ҳid, ���ͺ�_In, r_Advice.��id);
    Elsif r_Advice.�������� = '11' Then
      b_Message.Zlhis_Cis_022(r_Advice.����id, r_Advice.��ҳid, ���ͺ�_In, r_Advice.��id);
    End If;
  Elsif r_Advice.������� = 'E' And r_Advice.�������� = '5' Then
    b_Message.Zlhis_Cis_023(r_Advice.����id, r_Advice.��ҳid, ���ͺ�_In, r_Advice.��id);
  Elsif r_Advice.������� = 'H' And Nvl(r_Advice.��������, '0') = '0' Then
    b_Message.Zlhis_Cis_006(r_Advice.����id, r_Advice.��ҳid, ���ͺ�_In, r_Advice.��id);
  End If;

  --����ִ��ҽ������
  Select Count(1) Into n_��� From ��������˵�� B Where b.����id = r_Advice.ִ�п���id And b.�������� = '����';
  If n_��� > 0 Then
    b_Message.Zlhis_Cis_026(r_Advice.����id, r_Advice.��ҳid, ���ͺ�_In, ҽ��id_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ������_Insert;
/






------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0001' Where ���=&n_System;
Commit;
