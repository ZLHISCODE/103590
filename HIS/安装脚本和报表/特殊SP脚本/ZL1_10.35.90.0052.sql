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


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--132271:���ϴ�,2019-03-04,�Ӻ�ʱ�Զ�ȡ���
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
    �Ƿ�Һ�_In Number,
    ������λCheck_In   Varchar2 := Null
  ) Return Number Is
    --���ܣ�����Ƿ񳬳����޺Ż���Լ
    --���:�Ƿ�Һ�_IN-1:�Һ�;0-ԤԼ
    --����:1-��ʾ���ݺϷ�;0-��ʾ���ݲ��Ϸ�:�������޺Ż���Լ��
    n_Count Number(18);
    n_Temp  Number(18);
  Begin
    If Nvl(�ƻ�id1_In, 0) <> 0 Then
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
           Where ���� = ����_In And ���� between Trunc(����_In) And Trunc(����_In) + (1-1/24/60/60) And Nvl(״̬, 0) <> 4
           Union
           Select ���
           From ������λ�ƻ�����
           Where �ƻ�id = Decode(�Ƿ�Һ�_In, 1, 0, �ƻ�id1_In) And ������λ <> Nvl(������λCheck_In,'-') And ��� <> 0 And ������Ŀ = ����1_In And ���� <> 0
           Union
           Select ���
           From ������λ���ſ���
           Where ����id = Decode(�Ƿ�Һ�_In, 1, 0, ����id1_In) And ������λ <> Nvl(������λCheck_In,'-') And ��� <> 0 And ������Ŀ = ����1_In And ���� <> 0);
  
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
    If Nvl(�ƻ�id1_In, 0) <> 0 Then
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
    
      If Check_Nums_Valied(n_����id, n_�ƻ�id, v_����, n_�Ƿ�Һ�, ������λ_In) = 0 Then
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
      
        If Check_Nums_Valied(n_����id, n_�ƻ�id, v_����, n_�Ƿ�Һ�, ������λ_In) = 0 Then
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
      If Check_Nums_Valied(n_����id, n_�ƻ�id, v_����, n_�Ƿ�Һ�, ������λ_In) = 0 Then
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





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0052' Where ���=&n_System;
Commit;
