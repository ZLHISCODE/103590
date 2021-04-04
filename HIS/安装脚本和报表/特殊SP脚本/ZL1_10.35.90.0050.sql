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
--134869:��С��,2019-02-19,�ɼ�վ���������ӡ���Ȩ�ޡ���ɲɼ���
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1211,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0
Union All Select '��ɲɼ�',7,'�и�Ȩ��ʱ�����ڲ��������ӡ����ɲɼ�',0 From Dual) A;




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--137272:���ϴ�,2019-02-22,ԤԼ��ǰ�����������
CREATE OR REPLACE Procedure Zl_�Һ����״̬_Lock
(
  ����_In       Number, --1-����,2-�������
  ����Ա����_In �Һ����״̬.����Ա����%Type,
  ����_In       �Һ����״̬.����%Type := Null,
  ����_In       �Һ����״̬.����%Type := Null,
  ���_In       �Һ����״̬.���%Type := Null,
  �����¼ID_In �ٴ������¼.ID%type := Null,
  ��ע_In       �Һ����״̬.��ע%Type := Null
) As

  v_����       �Һ����״̬.����Ա����%Type;
  v_״̬       �Һ����״̬.״̬%Type;
  v_������     �Һ����״̬.������%Type;
  v_��֤������ �Һ����״̬.������%Type;
  v_��ע       �Һ����״̬.��ע%Type;
  v_����վIP   �ٴ�������ſ���.����վIP%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;
Begin
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  Select SYS_CONTEXT('USERENV','IP_ADDRESS') Into v_����վIP from dual;
  If ����_In = 1 Then
    If ��ע_In Is Null Then
      v_��ע := '����������';
    Else
      v_��ע := ��ע_In;
    End If;
    --�����Һ����״̬
    If �����¼ID_In is Null then
      Begin
        Select ����Ա����, ״̬, ������
        Into v_����, v_״̬, v_��֤������
        From �Һ����״̬
        Where ���� = ����_In And ���� = ����_In And ��� = ���_In;
      Exception
        When Others Then
          Null;
      End;
      If v_���� Is Null Then
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, ��ע, �Ǽ�ʱ��, ������)
        Values
          (����_In, ����_In, ���_In, 5, ����Ա����_In, v_��ע, Sysdate, v_������);
      Else
        v_Error := '���' || ���_In || '�ѱ�����Ա' || v_����;
        If v_״̬ = 1 Then
          v_Error := v_Error || 'ʹ��';
        Elsif v_״̬ = 2 Then
          v_Error := v_Error || 'ԤԼ';
        Elsif v_״̬ = 3 Then
          v_Error := v_Error || 'Ԥ��';
        Elsif v_״̬ = 4 Then
          v_Error := v_Error || '�˺�';
        Elsif v_״̬ = 5 Then
          v_Error := v_Error || '(' || v_��֤������ || ')����';
        End If;
        Raise Err_Custom;
      End If;
    Else
      Begin
        Select ����Ա����, �Һ�״̬, ����վ����
        Into v_����, v_״̬, v_��֤������
        From �ٴ�������ſ���
        Where ��¼ID = �����¼ID_In And ��� = ���_In;
      Exception
        When Others Then
          Null;
      End;
      
      If Nvl(v_״̬,0) = 0 Then
        Update �ٴ�������ſ��� set �Һ�״̬=5,����ʱ��=Sysdate,����Ա����=����Ա����_In,����վIP=v_����վIP,����վ����=v_������,��ע=v_��ע
        Where ��¼ID=�����¼ID_In  And ���=���_In;
      Else
        v_Error := '���' || ���_In || '�ѱ�����Ա' || v_����;
        If v_״̬ = 1 Then
          v_Error := v_Error || 'ʹ��';
        Elsif v_״̬ = 2 Then
          v_Error := v_Error || 'ԤԼ';
        Elsif v_״̬ = 3 Then
          v_Error := v_Error || 'Ԥ��';
        Elsif v_״̬ = 4 Then
          v_Error := v_Error || '�˺�';
        Elsif v_״̬ = 5 Then
          v_Error := v_Error || '(' || v_��֤������ || ')����';
        End If;
        Raise Err_Custom;
      End If;
    End If;
  Elsif ����_In = 2 Then
    If �����¼ID_In is Null then
       Delete �Һ����״̬ Where ������ = v_������ And ����Ա���� = ����Ա����_In And ״̬ = 5;
       
       Update �ٴ�������ſ��� A set A.�Һ�״̬=0,A.����ʱ��=NULL,A.����Ա����=NULL,A.����վIP=NULL,A.����վ����=NULL,A.����=NULL,A.����=NULL,A.��ע=NULL
       Where A.����վ���� =v_������ And A.����վIP=v_����վIP And A.����Ա���� = ����Ա����_In And A.�Һ�״̬ = 5 And A.����ʱ�� > Trunc(Sysdate);
    Else
      Update �ٴ�������ſ��� A set A.�Һ�״̬=0,A.����ʱ��=NULL,A.����Ա����=NULL,A.����վIP=NULL,A.����վ����=NULL,A.����=NULL,A.����=NULL,A.��ע=NULL
      Where A.����վ���� =v_������ And A.����վIP=v_����վIP And A.����Ա���� = ����Ա����_In And A.�Һ�״̬ = 5 And A.����ʱ�� > Trunc(Sysdate)
        And Exists (Select 1 From �ٴ������¼ B Where A.��¼ID=B.ID And B.�Ƿ���ſ��� = 1);
      
      Update �ٴ�������ſ��� A set A.�Һ�״̬=4,A.����ʱ��=NULL,A.����Ա����=NULL,A.����վIP=NULL,A.����վ����=NULL,A.����=NULL,A.����=NULL,A.��ע=NULL
      Where A.����վ���� =v_������ And A.����վIP=v_����վIP And A.����Ա���� = ����Ա����_In And A.�Һ�״̬ = 5 And A.����ʱ�� > Trunc(Sysdate)
        And Exists (Select 1 From �ٴ������¼ B Where A.��¼ID=B.ID And B.�Ƿ���ſ��� = 0);
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�Һ����״̬_Lock;
/

--137272:���ϴ�,2019-02-22,ԤԼ��ǰ�����������
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 1,
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  --���½������_In:0-��zl_��Ա�ɿ����_Update �и��� 1-�ڱ������и���
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, Id, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(Ԥ�����, 2) = 1 Having
     Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By No, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, No;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;

  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;
  v_�ű�     ������ü�¼.���㵥λ%Type;
  v_����     ������ü�¼.��ҩ����%Type;
  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_�������      ����Ԥ����¼.���%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;

  d_Date     Date;
  d_ԤԼʱ�� ������ü�¼.����ʱ��%Type;
  d_����ʱ�� Date;
  d_�Ŷ�ʱ�� Date;
  n_ʱ��     Number := 0;
  n_����     Number := 0;
  v_�Ŷ���� �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ ������Ϣ.����ģʽ%Type;

  v_���ʽ   ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_����Ա���� ���˹Һż�¼.������%Type;
  n_����ģʽ   Number := 0;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);

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
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Begin
    Select 1
    Into n_ʱ��
    From Dual
    Where Exists (Select 1
           From �ҺŰ���ʱ�� A, �ҺŰ��� B
           Where a.����id = b.Id And b.���� = v_�ű� And Rownum < 2
           Union All
           Select 1
           From �Һżƻ�ʱ�� C, �ҺŰ��żƻ� D ��
           Where c.�ƻ�id = d.Id And d.���� = v_�ű� And d.��Чʱ�� > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_ʱ�� := 0;
  End;
  --��ʱ�εĺű�ֻ�ܵ������
  If n_ʱ�� = 1 And ��������_In = 0 And n_����ģʽ = 0 Then
    If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
      Raise Err_Item;
    End If;
  End If;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;
  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
      
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
          --����Ƿ�����,n_���� ��ʾ�Ƿ�����ʹ��
          Select Count(1)
          Into n_����
          From �Һ����״̬
          Where ���� = v_�ű� And Trunc(����) = Trunc(Sysdate) And ��� = ����_In And (״̬ <> 5 Or ״̬ = 5 And ����Ա���� <> ����Ա����_In);
          If n_���� = 0 Then
            Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(Sysdate) And ��� = ����_In;
            v_���� := ����_In;
          Else
            Begin
              Select 1 Into n_���� From �Һ����״̬ Where ���� = v_�ű� And ���� = Trunc(Sysdate) And ��� = v_����;
            Exception
              When Others Then
                n_���� := 0;
            End;
          End IF;
          If n_���� = 0 Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
          Else
            --�����ѱ�ʹ�õ����
            Begin
              v_���� := 1;
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                Select Min(��� + 1)
                Into v_����
                From �Һ����״̬ A
                Where ���� = v_�ű� And ���� = Trunc(Sysdate) And Not Exists
                 (Select 1 From �Һ����״̬ Where ���� = a.���� And ���� = a.���� And ��� = a.��� + 1);
                Insert Into �Һ����״̬
                  (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
                Values
                  (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            End;
          End If;
        Else
          Update �Һ����״̬
          Set ״̬ = 1, �Ǽ�ʱ�� = Sysdate
          Where Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_���� And ���� = v_�ű� And ״̬ = 2;
          If Sql% NotFound Then
            Begin
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update �Һ����״̬
        Set ��� = ����_In, ״̬ = 1, �Ǽ�ʱ�� = Sysdate
        Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(d_����ʱ��), v_����, 1, ����Ա����_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      Delete �Һ����״̬
      Where ���� = v_�ű� And Trunc(����) = Trunc(Sysdate) And ��� = ����_In And ״̬ = 5 And ����Ա���� = ����Ա����_In;
      Begin
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
        Values
          (v_�ű�, Trunc(Sysdate), ����_In, 1, ����Ա����_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '���' || ����_In || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
      End;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Decode(�շѵ�_In, Null, Nvl(ժҪ_In, ժҪ), '����:' || �շѵ�_In)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      ժҪ = Nvl(ժҪ_In, ժҪ), �շѵ� = �շѵ�_In
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ, �շѵ�)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �շѵ�_In
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
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
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
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
      End If;
    End Loop;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And
     Nvl(���ʷ���_In, 0) = 0 Then
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������,
       ��������)
    Values
      (n_Ԥ��id, 4, 1, No_In, ����id_In, Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), d_Date, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
       n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ����id_In, 4);
  
    If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
      Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, �ֽ�֧��_In, n_Ԥ��id, ����Ա���_In, ����Ա����_In, d_Date);
    End If;
  
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Select Nvl(Sum(Nvl(Ԥ�����, 0) - Nvl(�������, 0)), 0)
    Into n_�������
    From �������
    Where ����id In (Select /*+cardinality(d,10)*/
                    d.Column_Value
                   From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(����, 0) = 1 And Nvl(����, 0) = 1;
    if n_������� < Ԥ��֧��_In Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || Ltrim(To_Char(n_�������, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                   Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
      Raise Err_Item;
    End if;
    
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
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
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
    IF n_Ԥ����� > 0 Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����С�ڱ���֧����� ' || Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '�����ܼ���������';
      Raise Err_Item;
    End IF;
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(�ֽ�֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 1) = 1 Then
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
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 1) = 1 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
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
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
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
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_Insert;
/

--137272:���ϴ�,2019-02-22,ԤԼ��ǰ�����������
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_����_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,

  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      Varchar2, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 1, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, Id, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(Ԥ�����, 2) = 1 Having
     Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By No, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, No;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;
  v_����Ա���� ���˹Һż�¼.������%Type;
  v_�ֽ�       ���㷽ʽ.����%Type;
  v_�����ʻ�   ���㷽ʽ.����%Type;
  v_��������   �ŶӽкŶ���.��������%Type;
  v_�ű�       ������ü�¼.���㵥λ%Type;
  v_����       ������ü�¼.��ҩ����%Type;
  v_�ŶӺ���   �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_�������      ����Ԥ����¼.���%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;

  d_Date         Date;
  d_ԤԼʱ��     ������ü�¼.����ʱ��%Type;
  d_����ʱ��     Date;
  d_�Ŷ�ʱ��     Date;
  n_ʱ��         Number := 0;
  n_����         Number := 0;
  v_��������     Varchar2(2000);
  v_��ǰ����     Varchar2(500);
  n_������     ����Ԥ����¼.��Ԥ��%Type;
  v_�������     ����Ԥ����¼.�������%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־   Number(3);
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ     ������Ϣ.����ģʽ%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_����ģʽ     Number := 0;
  n_�����¼id   ���˹Һż�¼.�����¼id%Type;
  n_�³����¼id ���˹Һż�¼.�����¼id%Type;
  n_��Դid       �ٴ������¼.��Դid%Type;
  n_ԤԼ˳���   �ٴ�������ſ���.ԤԼ˳���%Type;
  n_�ɷ�ʱ��     �ٴ������¼.�Ƿ��ʱ��%Type;
  n_����ſ���   �ٴ������¼.�Ƿ���ſ���%Type;
  n_�ɿ���id     �ٴ������¼.����id%Type;
  n_����Ŀid     �ٴ������¼.��Ŀid%Type;
  n_��ҽ��id     �ٴ������¼.ҽ��id%Type;
  n_�Һ�ģʽ     Number(3);
  d_����ʱ��     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_���         Number(3);
  n_��ſ���     �ٴ������¼.�Ƿ���ſ���%Type;
  v_���ϰ�ʱ��   �ٴ������¼.�ϰ�ʱ��%Type;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);
  n_�Һ�ģʽ      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_�Һ�ģʽ = 1 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
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
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ, �����¼id
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ, n_�����¼id
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Select Nvl(�Ƿ��ʱ��, 0), ��Դid, Nvl(�Ƿ���ſ���, 0)
  Into n_ʱ��, n_��Դid, n_��ſ���
  From �ٴ������¼
  Where ID = n_�����¼id;

  If n_ʱ�� = 1 And ��������_In = 0 And n_����ģʽ = 0 Then
    If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
      Raise Err_Item;
    End If;
  End If;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;

  If d_����ʱ�� Is Not Null Then
    If d_����ʱ�� < d_����ʱ�� Then
      v_Err_Msg := '��ǰԤԼ�Һŵ����ڳ�����Ű�ģʽ���ţ�������' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '֮ǰ����!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where ��� = v_���� And ��¼id = n_�����¼id;
        
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
        
          --����Ƿ�������n_���� ��ʾ��ſ���
          Select Count(1)
          Into n_����
          From �ٴ�������ſ���
          Where ��¼id = n_�³����¼id And ��� = ����_In And (Nvl(�Һ�״̬, 0) = 0 Or Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In) And
                Rownum < 2;
          If n_���� = 1 Then
            Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where ��¼id = n_�³����¼id And ��� = ����_In;
            v_���� := ����_In;
          ElsIF v_���� <> ����_In Then
            Begin
              Select 1
              Into n_����
              From �ٴ�������ſ���
              Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
            Exception
              When Others Then
                n_���� := 0;
            End;
          End if;
        
          If n_���� = 1 Then
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Else
            --�����ѱ�ʹ�õ����
            Select Min(���) Into v_���� From �ٴ�������ſ��� Where ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
            If v_���� Is Null Then
              v_Err_Msg := '���յ���û�п������,�޷�����!';
              Raise Err_Item;
            End If;
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          End If;
        Else
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
          Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
          Returning ԤԼ˳��� Into n_ԤԼ˳���;
        
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
          Where ��� = v_���� And ��¼id = n_�³����¼id And (Nvl(�Һ�״̬, 0) = 0 Or Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In);
          If Sql% RowCount = 0 Then
            v_Err_Msg := '���յ������' || v_���� || '�ѱ�������ʹ��,�޷�����.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
        Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
        From �ٴ������¼
        Where ID = n_�����¼id;
        Begin
          Select ID
          Into n_�³����¼id
          From �ٴ������¼
          Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
            Raise Err_Item;
        End;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
        Returning ԤԼ˳��� Into n_ԤԼ˳���;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
        Where ��� = ����_In And ��¼id = n_�³����¼id And (Nvl(�Һ�״̬, 0) = 0 Or Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In);
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���յ������' || ����_In || '�ѱ�������ʹ��,�޷�����.';
          Raise Err_Item;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id;
      End If;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;
  
  --���Һ���Ŀ,��Ϊzl_Custom_GetRegeventItem��Ŀ���ܲ�һ�£�ֻ������
  Select Count(1)
  Into n_Count
  From �ٴ������¼ a, ������ü�¼ b
  Where a.Id = Nvl(n_�³����¼id, n_�����¼id) And b.No = No_In And b.��� = 1 And a.����id = b.ִ�в���id;
  If n_Count = 0 Then
    v_Err_Msg := '�Һſ��Ҳ�һ�£��޷����գ�';
    Raise Err_Item;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Decode(�շѵ�_In, Null, Nvl(ժҪ_In, ժҪ), '����:' || �շѵ�_In)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('�Һ��Ű�ģʽ');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_����ʱ�� Then
        v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || 'δ���ó�����Ű�ģʽ,Ŀǰ�޷�����!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_���
      From �ٴ������¼
      Where ID = Nvl(n_�³����¼id, n_�����¼id) And d_����ʱ�� Between ͣ�￪ʼʱ�� And ͣ����ֹʱ��;
    Exception
      When Others Then
        n_��� := 0;
    End;
    If n_��� = 1 And Not (n_ʱ�� = 1 And n_��ſ��� = 1) Then
      v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '�İ����Ѿ���ͣ��,�޷�����!';
      Raise Err_Item;
    End If;
  End If;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      �����¼id = Nvl(n_�³����¼id, n_�����¼id), ժҪ = Nvl(ժҪ_In, ժҪ), �շѵ� = �շѵ�_In
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ, �����¼id, �շѵ�)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, Nvl(n_�³����¼id, n_�����¼id),
               �շѵ�_In
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(v_�Һ�.ִ�в���id, 0)));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
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
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
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
      End If;
    End Loop;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 Then
      v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, Null, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4, v_�������);
        
          If Nvl(���㿨���_In, 0) <> 0 And Nvl(n_������, 0) <> 0 Then
            Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, n_������, n_Ԥ��id, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In);
          End If;
        End If;
      
        If Nvl(���½������_In, 1) = 1 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + n_������
          Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�)
          Returning ��� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, Nvl(v_���㷽ʽ, v_�ֽ�), 1, n_������);
            n_����ֵ := n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Select Nvl(Sum(Nvl(Ԥ�����, 0) - Nvl(�������, 0)), 0)
    Into n_�������
    From �������
    Where ����id In (Select /*+cardinality(d,10)*/
                    d.Column_Value
                   From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(����, 0) = 1 And Nvl(����, 0) = 1;
    if n_������� < Ԥ��֧��_In Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || Ltrim(To_Char(n_�������, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                   Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
      Raise Err_Item;
    End if;
    
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
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
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
    IF n_Ԥ����� > 0 Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����С�ڱ���֧����� ' || Ltrim(To_Char(Ԥ��֧��_In, '9999999990.00')) || '�����ܼ���������';
      Raise Err_Item;
    End IF;
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 1) = 1 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
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
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
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
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_����_Insert;
/

--124583:��ҵ��,2019-02-21,���ŷ�ҩ,��ҩ��д��ҩ����
Create Or Replace Procedure Zl_ҩƷ�շ���¼_������ҩ
(
  Billinfo_In   In Varchar2, --��ʽ:"id1,����1|id2,����2|....."
  Partid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  People_In     In ҩƷ�շ���¼.�����%Type,
  Date_In       In ҩƷ�շ���¼.�������%Type,
  ��ҩ��ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3,
  ��ҩ��_In     In ҩƷ�շ���¼.������%Type := Null,
  ���ܷ�ҩ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
  Intdigit_In   In Number := 2,
  ��ҩ��_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
  �˲���_In     In ҩƷ�շ���¼.�˲���%Type := Null
) Is
  --ֻ������
  v_Infotmp     Varchar2(4000);
  v_Fields      Varchar2(4000);
  n_Billid      ҩƷ�շ���¼.Id%Type;
  n_����        ҩƷ�շ���¼.����%Type;
  Lng������id Number(18);
  Int���ϵ��   Number;
  Intִ��״̬   Number;
  Int����       ҩƷ�շ���¼.����%Type;
  Strno         ҩƷ�շ���¼.No%Type;
  Lng�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  LngҩƷid     ҩƷ�շ���¼.ҩƷid%Type;
  Lng����id     ҩƷ�շ���¼.����id%Type;
  Dbl�����     Number;
  v_���ۼ�      ҩƷ�շ���¼.���ۼ�%Type;
  Intδ����     δ��ҩƷ��¼.δ����%Type;
  v_�˲�����    ҩƷ�շ���¼.�˲�����%Type;
  --��д����
  Dblʵ������ ҩƷ�շ���¼.ʵ������%Type;
  Dblʵ�ʽ�� ҩƷ�շ���¼.���۽��%Type;
  Dbl�ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  Dblʵ�ʲ�� ҩƷ�շ���¼.���%Type;
  --2002-07-31����
  --LNGLAST���� ��ҩǰȷ��������(�Ѽ���������)
  Strҩ��           Varchar2(200);
  Dbl��������       ҩƷ�շ���¼.��д����%Type;
  Lnglast����       ҩƷ�շ���¼.����%Type;
  Lngcur����        ҩƷ�շ���¼.����%Type;
  Str����           ҩƷ�շ���¼.����%Type;
  StrЧ��           ҩƷ�շ���¼.Ч��%Type;
  n_�ϴι�Ӧ��id    ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ϴβɹ���      ҩƷ���.�ϴβɹ���%Type;
  v_�ϴβ���        ҩƷ���.�ϴβ���%Type;
  d_�ϴ���������    ҩƷ���.�ϴ���������%Type;
  v_��׼�ĺ�        ҩƷ���.��׼�ĺ�%Type;
  n_��¼״̬        ҩƷ�շ���¼.��¼״̬%Type;
  n_ƽ���ɱ���      ҩƷ���.ƽ���ɱ���%Type;
  n_��ҩ��ʽ        ҩƷ�շ���¼.��ҩ��ʽ%Type;
  v_ժҪ            ҩƷ�շ���¼.ժҪ%Type;
  Bln�շ��뷢ҩ���� Number(1);
  v_Error           Varchar2(255);
  Err_Custom Exception;
  n_��ͨ�ۼ�С��   Number;
  n_��ͨ���С��   Number;
  n_���۹���ģʽ Number(1);
  n_ҩƷ���۹��� Number(1);
  n_��������       δ��ҩƷ��¼.��������%Type;
Begin
  Select Sysdate Into v_�˲����� From Dual;
  If Billinfo_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := Billinfo_In || '|';
  End If;
  While v_Infotmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
    n_Billid  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_����    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
  
    --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID,���۽�ʵ��������������ID
    Begin
      Select a.����, a.No, a.ҩƷid, a.�ⷿid, a.����id, Nvl(a.���ۼ�, 0), Nvl(a.���۽��, 0), Nvl(a.ʵ������, 0) * Nvl(a.����, 1), a.������id,
             a.���ϵ��, Nvl(a.����, 0), a.����, a.Ч��, a.��ҩ��λid, a.����, a.��������, a.��׼�ĺ�, Nvl(a.��ҩ��ʽ, 0), a.ժҪ, ��¼״̬
      Into Int����, Strno, LngҩƷid, Lng�ⷿid, Lng����id, v_���ۼ�, Dblʵ�ʽ��, Dblʵ������, Lng������id, Int���ϵ��, Lnglast����, Str����, StrЧ��,
           n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_��ҩ��ʽ, v_ժҪ, n_��¼״̬
      From ҩƷ�շ���¼ A
      Where a.Id = n_Billid And a.������� Is Null
      For Update Nowait;
    
      Select '[' || c.���� || ']' || c.���� Into Strҩ�� From �շ���ĿĿ¼ C Where c.Id = LngҩƷid;
    Exception
      When Others Then
        Int���� := 0;
        v_Error := '���������û���ִ�з�ҩ�������ظ�������';
        Raise Err_Custom;
    End;
  
    --ȡ��ͨҵ�񾫶�λ��
    --���:1-ҩƷ 2-����
    --���ݣ�2-���ۼ� 4-���
    --��λ��ҩƷ:1-�ۼ� 5-��λ
    Begin
      Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
    Exception
      When Others Then
        n_��ͨ���С�� := 2;
    End;
  
    Begin
      Select ���� Into n_��ͨ�ۼ�С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 2 And ��λ = 1;
    Exception
      When Others Then
        n_��ͨ�ۼ�С�� := 2;
    End;
  
    Select Zl_To_Number(Nvl(zl_GetSysParameter(275), '0')) Into n_���۹���ģʽ From Dual;
  
    If n_��ҩ��ʽ = -1 Or v_ժҪ = '�ܷ�' Then
      Int���� := 0;
    End If;
  
    If Int���� > 0 Then
      If Nvl(n_����, 0) = 0 Then
        Lngcur���� := Lnglast����;
      Else
        Lngcur���� := Nvl(n_����, 0);
      End If;
    
      --����Ƿ��Ѿ���д�ⷿ
      Bln�շ��뷢ҩ���� := 0;
      If Lng�ⷿid Is Null Then
        Bln�շ��뷢ҩ���� := 1;
      End If;
      Lng�ⷿid := Partid_In;
    
      --ȡ����ҩƷ������
      Begin
        Select �ϴ�����, Ч��, Nvl(��������, 0), �ϴι�Ӧ��id, �ϴβ���, �ϴ���������, ��׼�ĺ�, �ϴβɹ���
        Into Str����, StrЧ��, Dbl��������, n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���
        From ҩƷ���
        Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lngcur����;
      Exception
        When Others Then
          n_�ϴβɹ��� := 0;
          Dbl��������  := 0;
      End;
    
      --���������������˳�
      If Lngcur���� <> Nvl(Lnglast����, 0) Then
        If Dbl�������� < Dblʵ������ And Lngcur���� <> 0 Then
          v_Error := Strҩ�� || '�Ŀ����������㣬������ֹ��';
          Raise Err_Custom;
        End If;
      End If;
    
      If n_���۹���ģʽ <> 0 Then
        Select Nvl(�Ƿ����۹���, 0) Into n_ҩƷ���۹��� From ҩƷ��� Where ҩƷid = LngҩƷid;
      End If;
    
      If n_��¼״̬ = 1 Then
        --ԭʼ��ҩ��¼��ȡ���¼۸�
        n_ƽ���ɱ��� := Zl_Fun_Getoutcost(LngҩƷid, Lngcur����, Lng�ⷿid);
      
        If n_���۹���ģʽ <> 0 And n_ҩƷ���۹��� = 1 And (v_���ۼ� = n_ƽ���ɱ��� Or Round(v_���ۼ�, n_��ͨ�ۼ�С��) = Round(n_ƽ���ɱ���, n_��ͨ�ۼ�С��)) Then
          Dbl�ɱ���� := Dblʵ�ʽ��;
        Else
          Dbl�ɱ���� := Round(n_ƽ���ɱ��� * Nvl(Dblʵ������, 0), n_��ͨ���С��);
        End If;
      Else
        --��ҩ�ٷ���¼��ȡԭʼ���ݼ۸�
        Select a.�ɱ���
        Into n_ƽ���ɱ���
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = n_Billid And a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Nvl(a.����, 0) = Nvl(b.����, 0) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);
      
        Dbl�ɱ���� := Round(n_ƽ���ɱ��� * Nvl(Dblʵ������, 0), n_��ͨ���С��);
      End If;
    
      Dblʵ�ʲ�� := Round(Dblʵ�ʽ�� - Dbl�ɱ����, n_��ͨ���С��);
    
      --����ҩƷ�շ���¼�����۽��ɱ������
      Update ҩƷ�շ���¼
      Set �ⷿid = Lng�ⷿid, �ɱ��� = n_ƽ���ɱ���, �ɱ���� = Dbl�ɱ����, ��� = Dblʵ�ʲ��, ���� = Lngcur����, ���� = Str����, Ч�� = StrЧ��,
          ��ҩ�� = ��ҩ��_In, �˲��� = �˲���_In, �˲����� = v_�˲�����, ����� = People_In, ������� = Date_In, ��ҩ��ʽ = ��ҩ��ʽ_In, ������ = ��ҩ��_In,
          ���ܷ�ҩ�� = ���ܷ�ҩ��_In, ��ҩ��λid = n_�ϴι�Ӧ��id, ���� = v_�ϴβ���, �������� = d_�ϴ���������, ��׼�ĺ� = v_��׼�ĺ�
      Where ID = n_Billid;
      --�����������
      If Sql%RowCount = 0 Then
        v_Error := 'Ҫ��ҩ��ҩƷ��¼"' || Strҩ�� || '"�����ڣ�������ֹ��';
        Raise Err_Custom;
      End If;
    
      --����סԺ���ü�¼��ִ��״̬(��ִ��)
      Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 1, 0, 1, 2)
      Into Intִ��״̬
      From ҩƷ�շ���¼
      Where ���� = Int���� And NO = Strno And ����id = Lng����id And ����� Is Null;
      Update סԺ���ü�¼
      Set ִ��״̬ = Intִ��״̬, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ��ʱ�� = Date_In, ִ�в���id = Partid_In
      Where ID = Lng����id;
    
      --����δ��ҩƷ��¼(���δ����Ϊ����ɾ��)
      Select Count(*)
      Into Intδ����
      From ҩƷ�շ���¼
      Where ���� = Int���� And (�ⷿid + 0 = Lng�ⷿid Or �ⷿid Is Null) And NO = Strno And ����� Is Null And
            Nvl(LTrim(RTrim(ժҪ)), 'С��') <> '�ܷ�';
    
      If Intδ���� = 0 Then
        Delete δ��ҩƷ��¼
        Where NO = Strno And ���� = Int���� And (�ⷿid + 0 = Lng�ⷿid Or �ⷿid Is Null)
        Returning �������� Into n_��������;
      
        --���´������ͣ������Ŵ�������
        Update ҩƷ�շ���¼ Set ע��֤�� = n_�������� Where ���� = Int���� And NO = Strno And �ⷿid = Lng�ⷿid;
      End If;
    
      --���¿��
      Zl_ҩƷ���_Update(n_Billid, 2, 1);
    
      Zl_δ��ҩƷ��¼_Delete(n_Billid);
    
      --�����������
      Zl_ҩƷ�շ���¼_��������(n_Billid);
    
      b_Message.Zlhis_Drug_005(Lng�ⷿid, n_Billid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_������ҩ;
/

--124583:��ҵ��,2019-02-21,���ŷ�ҩ,��ҩ��д��ҩ����
Create Or Replace Procedure Zl_ҩƷ�շ���¼_������ҩ
(
  Billid_In     In ҩƷ�շ���¼.Id%Type,
  People_In     In ҩƷ�շ���¼.�����%Type,
  Date_In       In ҩƷ�շ���¼.�������%Type,
  ����_In       In ҩƷ���.�ϴ�����%Type := Null,
  Ч��_In       In ҩƷ���.Ч��%Type := Null,
  ����_In       In ҩƷ���.�ϴβ���%Type := Null,
  ��ҩ����_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
  ��ҩ�ⷿ_In   In ҩƷ�շ���¼.�ⷿid%Type := Null,
  ��ҩ��_In     In ҩƷ�շ���¼.������%Type := Null,
  Intdigit_In   In Number := 2,
  ����_In       In Number := 2,
  ���ܷ�ҩ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null
) Is
  --ֻ������
  Int��¼״̬   ҩƷ�շ���¼.��¼״̬%Type;
  Intִ��״̬   סԺ���ü�¼.ִ��״̬%Type;
  Bln������ҩ   Number;
  Lng������id Number(18);
  Strno         ҩƷ�շ���¼.No%Type;
  Int����       ҩƷ�շ���¼.����%Type;
  Lng�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  LngҩƷid     ҩƷ�շ���¼.ҩƷid%Type;
  Dblʵ������   ҩƷ�շ���¼.ʵ������%Type;
  Dblʵ�ʽ��   ҩƷ�շ���¼.���۽��%Type;
  Dblʵ�ʳɱ�   ҩƷ�շ���¼.�ɱ����%Type;
  Dblʵ�ʲ��   ҩƷ�շ���¼.���%Type;
  Lng����id     ҩƷ�շ���¼.����id%Type;
  n_���ۼ�      ҩƷ�շ���¼.���ۼ�%Type;
  n_�Ƿ���    Number;
  n_ʱ�۷���    Number;

  --20020731 Modified by zyb
  --������ҩʱ�������������ʸı��Ĵ���
  Lng������ ҩƷ�շ���¼.����%Type;
  Lng����   ҩƷ���.ҩ������%Type;
  Lng����   ҩƷ�շ���¼.����%Type; --ԭ����

  Str����        ҩƷ�շ���¼.����%Type; --ԭ����
  DateЧ��       ҩƷ�շ���¼.Ч��%Type; --ԭЧ��
  n_�ϴι�Ӧ��id ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ϴβɹ���   ҩƷ���.�ϴβɹ���%Type;
  v_�ϴβ���     ҩƷ���.�ϴβ���%Type;
  v_ԭ����       ҩƷ���.ԭ����%Type;
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
  v_��׼�ĺ�     ҩƷ���.��׼�ĺ�%Type;

  n_��¼����   סԺ���ü�¼.��¼����%Type;
  v_�շ����   סԺ���ü�¼.�շ����%Type;
  n_����       ҩƷ�շ���¼.����%Type;
  n_ԭʼ����   ҩƷ�շ���¼.ʵ������%Type;
  v_������¼id ҩƷ�շ���¼.Id%Type;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_��ҩȷ�� ҩ����ҩ����.��ҩȷ��%Type;
  v_��ҩ     ҩ����ҩ����.��ҩ%Type;
  v_�Ŷ�״̬ Number(1);
  v_ִ��ʱ�� ҩƷ�շ���¼.�������%Type;

Begin
  If ��ҩ����_In Is Not Null Then
    If ��ҩ����_In = 0 Then
      Return;
    End If;
  End If;

  --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID
  Select a.����, a.No, a.�ⷿid, a.ҩƷid, a.����id, a.������id, a.��¼״̬, Nvl(a.����, 0), a.����, a.Ч��, a.��ҩ��λid, a.����, a.ԭ����, a.��������,
         a.��׼�ĺ�, a.�ɱ���, a.����, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) As ʵ������, a.���ۼ�, Nvl(b.�Ƿ���, 0) �Ƿ���
  Into Int����, Strno, Lng�ⷿid, LngҩƷid, Lng����id, Lng������id, Int��¼״̬, Lng����, Str����, DateЧ��, n_�ϴι�Ӧ��id, v_�ϴβ���, v_ԭ����,
       d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���, n_����, n_ԭʼ����, n_���ۼ�, n_�Ƿ���
  From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
  Where a.ҩƷid = b.Id And a.Id = Billid_In;

  Begin
    Select Nvl(��ҩȷ��, 0), Nvl(��ҩ, 0)
    Into v_��ҩȷ��, v_��ҩ
    From ҩ����ҩ����
    Where ҩ��id = Lng�ⷿid And Rownum = 1;
  
  Exception
    When Others Then
      v_��ҩȷ�� := 0;
      v_��ҩ     := 0;
      Null;
  End;

  If v_��ҩȷ�� = 0 And v_��ҩ = 0 Then
    v_�Ŷ�״̬ := 2;
  Elsif v_��ҩȷ�� = 1 Then
    v_�Ŷ�״̬ := 0;
  Elsif v_��ҩ = 1 Then
    v_�Ŷ�״̬ := 1;
  End If;

  --��ȡ�ñʼ�¼ʣ��δ�������������
  --������������δ���������
  Select Sum(Nvl(ʵ������, 0) * Nvl(����, 1)), Sum(Nvl(���۽��, 0)), Sum(Nvl(�ɱ����, 0)), Sum(Nvl(���, 0))
  Into Dblʵ������, Dblʵ�ʽ��, Dblʵ�ʳɱ�, Dblʵ�ʲ��
  From ҩƷ�շ���¼
  Where ����� Is Not Null And NO = Strno And ���� = Int���� And ��� = (Select ��� From ҩƷ�շ���¼ Where ID = Billid_In);

  --���������ҩ��Ϊ�㣬��ʾ����ҩ
  If Dblʵ������ = 0 Then
    v_Error := '�õ����ѱ���������Ա��ҩ����ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;
  If Nvl(��ҩ����_In, 0) > Dblʵ������ Then
    v_Error := '�õ����ѱ���������Ա������ҩ����ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;

  --��ȡ��ҩƷ��ǰ�Ƿ��������Ϣ
  Select Nvl(ҩ������, 0) Into Lng���� From ҩƷ��� Where ҩƷid = LngҩƷid;
  --����ǲ�����ҩ�������¼������۽����
  Bln������ҩ := 0;
  If Not (��ҩ����_In Is Null Or Nvl(��ҩ����_In, 0) = Dblʵ������) Then
    Bln������ҩ := 1;
  End If;
  If Bln������ҩ = 1 Then
    Dblʵ�ʽ�� := Round(Dblʵ�ʽ�� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ�ʳɱ� := Round(Dblʵ�ʳɱ� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ�ʲ�� := Round(Dblʵ�ʲ�� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ������ := ��ҩ����_In;
  End If;

  If n_ԭʼ���� = ��ҩ����_In Then
    Dblʵ������ := ��ҩ����_In / n_����;
  Else
    n_���� := 1;
  End If;

  --lng����:0-������;1-����;2-ԭ�������ֲ�������������������;3-ԭ���������ַ���������������
  If Lng���� = 0 And Lng���� <> 0 Then
    --ԭ�������ֲ�������������������
    Lng���� := 2;
  Elsif Lng���� <> 0 And Lng���� = 0 Then
    --ԭ������,�ַ���,�����µ����Σ������²����ķ�ҩ��¼��ʹ��
    Lng���� := 3;
  Else
    If Lng���� = 0 Then
      Lng���� := 0;
    Else
      Lng���� := 1;
    End If;
  End If;
  --�ж��Ƿ�ʱ�۷���
  If (Lng���� = 1 Or Lng���� = 3) And n_�Ƿ��� = 1 Then
    n_ʱ�۷��� := 1;
  Else
    n_ʱ�۷��� := 0;
  End If;

  --��¼״̬�ĺ��������仯
  --�����ļ�¼״̬        :iif(int��¼״̬=1,0,1)+1
  --�������ļ�¼״̬        :iif(int��¼״̬=1,0,1)+2
  --�ȴ���ҩ�ļ�¼״̬    :iif(int��¼״̬=1,0,1)+3

  --����������¼
  Select ҩƷ�շ���¼_Id.Nextval Into v_������¼id From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ���, ������, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��, ��ҩ��ʽ, ע��֤��, ԭ����)
    Select v_������¼id, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 1, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����,
           ����, Ч��, n_����, -dblʵ������, -dblʵ������, �ɱ���, -dblʵ�ʳɱ�, ����, ���ۼ�, -dblʵ�ʽ��, -dblʵ�ʲ��, ժҪ, People_In, Date_In, ��ҩ��,
           People_In, Date_In, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ�ⷿ_In, ��ҩ��_In, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��_In, ��ҩ��ʽ, ע��֤��, ԭ����
    From ҩƷ�շ���¼
    Where ID = Billid_In;

  --����ǲ��ֳ�����������Ϊ1��ʵ������Ϊ������ʵ�������Ļ�
  --����������¼�Թ�������ҩ
  Select ҩƷ�շ���¼_Id.Nextval Into Lng������ From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, ԭ����)
    Select Lng������, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 3, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid,
           Decode(Lng����, 1, ����, 3, Lng������, 0), Decode(Lng����, 3, ����_In, 1, ����, ����), Decode(Lng����, 3, ����_In, ����),
           Decode(Lng����, 3, Ч��_In, Ч��), n_����, Dblʵ������, Dblʵ������, �ɱ���, Dblʵ�ʳɱ�, ����, ���ۼ�, Dblʵ�ʽ��, Dblʵ�ʲ��, ժҪ, ������, ��������,
           Null, Null, Null, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, ԭ����
    
    From ҩƷ�շ���¼
    Where ID = Billid_In;

  Zl_δ��ҩƷ��¼_Insert(Lng������);

  --���·��ü�¼��ִ��״̬(0-δִ��;1-��ȫִ��;2-����ִ��)
  Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 0, 0, 0, 2)
  Into Intִ��״̬
  From ҩƷ�շ���¼
  Where ���� = Int���� And NO = Strno And ����id = Lng����id And ����� Is Not Null;

  If ����_In = 1 Then
    Select ��¼����, �շ���� Into n_��¼����, v_�շ���� From ������ü�¼ Where ID = Lng����id;
  Else
    Select ��¼����, �շ���� Into n_��¼����, v_�շ���� From סԺ���ü�¼ Where ID = Lng����id;
  End If;

  If Intִ��״̬ = 0 Then
    If ����_In = 1 Then
      Update ������ü�¼
      Set ִ��״̬ = Intִ��״̬, ִ���� = Null, ִ��ʱ�� = Null
      Where NO = Strno And
            ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = Billid_In)) And
            Mod(��¼����, 10) = n_��¼���� And ��¼״̬ <> 2 And ִ�в���id = Lng�ⷿid;
    Else
      Update סԺ���ü�¼ Set ִ��״̬ = Intִ��״̬, ִ���� = Null, ִ��ʱ�� = Null Where ID = Lng����id;
    End If;
  Else
    If ����_In = 1 Then
      Update ������ü�¼
      Set ִ��״̬ = Intִ��״̬
      Where NO = Strno And
            ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = Billid_In)) And
            Mod(��¼����, 10) = n_��¼���� And ��¼״̬ <> 2 And ִ�в���id = Lng�ⷿid;
    Else
      Update סԺ���ü�¼ Set ִ��״̬ = Intִ��״̬ Where ID = Lng����id;
    End If;
  End If;

  --����δ��ҩƷ��¼
  Begin
    If ����_In = 1 Then
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��ҩ����, ��������, ���շ�, ��ҩ��, ��ӡ״̬, δ����, ��ҩ��, �Ŷ�״̬)
        Select a.����, a.No, a.����id, Null, a.����, Nvl(b.���ȼ�, 0) ���ȼ�, a.�Է�����id, a.�ⷿid, a.��ҩ����, a.��������, a.���շ�, Null, 1, 1,
               a.��Ʒ�ϸ�֤, v_�Ŷ�״̬
        From (Select b.����, b.No, a.����id, a.����, Decode(a.��¼״̬, 0, 0, 1) ���շ�, b.�Է�����id, b.�ⷿid, b.��ҩ����, b.��������, c.���,
                      b.��Ʒ�ϸ�֤
               From ������ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
               Where b.Id = Billid_In And a.Id = b.����id + 0 And a.����id = c.����id(+)) A, ��� B
        Where b.����(+) = a.���;
    Else
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��ҩ����, ��������, ���շ�, ��ҩ��, ��ӡ״̬, δ����, ��ҩ��, �Ŷ�״̬)
        Select a.����, a.No, a.����id, a.��ҳid, a.����, Nvl(b.���ȼ�, 0) ���ȼ�, a.�Է�����id, a.�ⷿid, a.��ҩ����, a.��������, a.���շ�, Null, 1, 1,
               a.��Ʒ�ϸ�֤, v_�Ŷ�״̬
        From (Select b.����, b.No, a.����id, a.��ҳid, a.����, Decode(a.��¼״̬, 0, 0, 1) ���շ�, b.�Է�����id, b.�ⷿid, b.��ҩ����, b.��������,
                      c.���, b.��Ʒ�ϸ�֤
               From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
               Where b.Id = Billid_In And a.Id = b.����id + 0 And a.����id = c.����id(+)) A, ��� B
        Where b.����(+) = a.���;
    End If;
  
    --�޸Ĵ�������
    Zl_Prescription_Type_Update(Strno, n_��¼����, LngҩƷid, v_�շ����);
  Exception
    When Others Then
      Null;
  End;

  --�޸�ԭ��¼Ϊ��������¼
  Update ҩƷ�շ���¼ Set ��¼״̬ = Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 2 Where ID = Billid_In;

  --�޸�ҩƷ���(������)
  If Lng���� <> 3 Then
    --����������Ҫ������ʵ�������ͽ���ۻ���ȥ���������û�����ڿ����������
    Zl_ҩƷ���_Update(v_������¼id, 3, 0);
  Else
    --ԭ�����������ڷ�����ֱ���ڿ�������µ���
    Insert Into ҩƷ���
      (�ⷿid, ҩƷid, ����, Ч��, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�, �ϴ�����, �ϴβ���, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
    Values
      (Lng�ⷿid, LngҩƷid, Lng������, Ч��_In, 1, Dblʵ������ * n_����, Dblʵ�ʽ��, Dblʵ�ʲ��, Decode(n_ʱ�۷���, 1, n_���ۼ�, Null), ����_In,
       ����_In, n_�ϴι�Ӧ��id, n_�ϴβɹ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���);
  End If;

  Delete ҩƷ���
  Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
        Nvl(ʵ�ʲ��, 0) = 0;

  --�����������
  Zl_ҩƷ�շ���¼_��������(v_������¼id);

  Begin
    --�ƶ�֧������Ŀ�ڷ�ҩ��̬��������������Ϣ�Ĺ���
    Execute Immediate 'Begin zl_������Ϣ_����(:1,:2); End;'
      Using 7, Billid_In || ',' || ��ҩ����_In || ',' || ����_In;
  Exception
    When Others Then
      Null;
  End;

  --��Ϣ����ʣ��ȫ����������0
  If Bln������ҩ = 1 Then
    b_Message.Zlhis_Drug_006(v_������¼id, Lng������, Dblʵ������ * n_����, Lng����id);
  Else
    b_Message.Zlhis_Drug_006(v_������¼id, Lng������, 0, Lng����id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_������ҩ;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0050' Where ���=&n_System;
Commit;
