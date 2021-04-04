create or replace package b_PacsInterface is
  Type t_Refcur Is Ref Cursor;


  -----------------------------------------------------------------------------
  --��ȡ��������Ϣ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  Procedure GetDeptItems
  (
  Cursor_Out  Out	t_Refcur,
  ��������_In  In  Varchar2:=Null
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ�ѱ�
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  Procedure GetChargeTypes
  (
    Cursor_Out  Out  t_Refcur,
    ��������_In  In  Varchar2:=Null
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡpacs�����Ŀ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  procedure GetPacsItems
  (
    Cursor_Out  Out  t_Refcur,
    ��������_In  In  Varchar2:=Null
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ�����Ŀ��ϸ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  procedure GetAdviceItems
  (
    Cursor_Out  Out  t_Refcur,
    ҽ��id_In  In  ����ҽ����¼.ID%Type
  );

  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ��������ϸ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  procedure GetAdviceFees
  (
    Cursor_Out  Out  t_Refcur,
    ҽ��id_In  In  ����ҽ����¼.ID%Type
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ����ؿ���ҽ����Ϣ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  procedure GetPacsDeptDoctor
  (
    Cursor_Out  Out  t_Refcur,
    ��������_In  In  Varchar2:=Null
  );

  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ������Ϣ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  Procedure GetPatient
  (
    Cursor_Out  Out  t_Refcur,
    ���ҷ�ʽ_In  In  Number,
    ��������_In  In  Varchar2
  );


  ---------------------------------------------------------------------------------------------------------------
  -- ��    �ܣ���ȡҽ������״̬
  -- �����б�
  --
  --
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestStatus
  (
    Cursor_Out  Out  t_Refcur,
    ҽ��id_In  In  ����ҽ����¼.ID%Type
  );


  ---------------------------------------------------------------------------------------------------------------
  -- ��    �ܣ���ȡ���������Ϣ
  -- �����б�
  --
  --
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestInfo
  (
    Cursor_Out  Out  t_Refcur,
    ���ҷ�ʽ_In  In  Number,
    ��������_In  In  Varchar2,
    ��������_In  In  Varchar2:=null
  );



  ---------------------------------------------------------------------------------------------------------------
	-- ��    �ܣ���ȡ�ĵ����������Ϣ
  -- �����б�
  --
  --
	---------------------------------------------------------------------------------------------------------------
	Procedure GetRequestInfo1
	(
		Cursor_Out	Out	t_Refcur,
		��ʼ����_In	In	Varchar2,
		��������_In	In	Varchar2,
    ������_In In  Varchar2
	);


  -----------------------------------------------------------------------------
  -- ��    �ܣ�ȡ������/�������
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  Procedure CancelRequest
  (
    ҽ��id_In  In  ����ҽ������.ҽ��ID%Type,
    ����ִ��_In   Number := 0,
    ִ�в���ID_IN ���ű�.id%Type := 0
  );


    -----------------------------------------------------------------------------
  -- ��    �ܣ�ɾ��������Ϣ
  -----------------------------------------------------------------------------
  PROCEDURE DeleteReport
  (
    ҽ��id_IN ����ҽ������.ҽ��ID%TYPE
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ�ɾ���ĵ籨����Ϣ
  -----------------------------------------------------------------------------
  PROCEDURE DeleteElectrocardioReport
  (
    ҽ��id_IN ����ҽ������.ҽ��ID%TYPE
  );



  -----------------------------------------------------------------------------
  -- ��    �ܣ������������
  -----------------------------------------------------------------------------
  PROCEDURE ClearPacsReport
  (
    ҽ��id_IN ����ҽ������.ҽ��ID%TYPE
  );

  -----------------------------------------------------------------------------
  -- ��    �ܣ����ռ���/�������
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  Procedure RecevieRequest
  (
    ҽ��id_IN     ����ҽ������.ҽ��ID%TYPE,

    ִ�м�_IN     ����ҽ������.ִ�м�%TYPE:=Null,
    ����_IN     Ӱ�����¼.����%TYPE:=NULL,
    ����豸_IN   Ӱ�����¼.����豸%TYPE:=Null,
    ���_IN       Ӱ�����¼.���%TYPE:=Null,
    ����_IN       Ӱ�����¼.����%TYPE:=Null,
    ��鼼ʦ_IN   Ӱ�����¼.��鼼ʦ%TYPE:=Null,
    ִ��ʱ��_IN   ����ҽ������.����ʱ��%TYPE:=Null,
    ִ��˵��_IN   ����ҽ������.ִ��˵��%TYPE:=NULL,
    ����ִ��_In   Number := 0,
    ִ�в���ID_IN ���ű�.id%Type := 0
  );

  -----------------------------------------------------------------------------
  -- ��    �ܣ����ͱ����ı���Ϣ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  procedure SendReport
  (
    ҽ��id_IN       ����ҽ������.ҽ��ID%TYPE,
    ��������_IN     ���Ӳ�������.�����ı�%TYPE,
    ���潨��_IN     ���Ӳ�������.�����ı�%TYPE,
    ����ҽ��_IN     ���Ӳ�����¼.������%TYPE,
    ���ҽ��_IN     Ӱ�����¼.������%TYPE := Null,
    ִ�в���ID_IN ���ű�.id%Type := 0
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ������ĵ籨����Ϣ
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  procedure SendElectrocardioReport
  (
    ҽ��id_IN       ����ҽ������.ҽ��ID%TYPE,
    �������_IN     ���Ӳ�������.�����ı�%TYPE,
    ��Ͻ��_IN     ���Ӳ�������.�����ı�%TYPE,
    ��Ͻ���_IN     ���Ӳ�������.�����ı�%TYPE,
    ����ҽ��_IN     ���Ӳ�����¼.������%TYPE,
    ���ҽ��_IN     Ӱ�����¼.������%TYPE := Null
  );

  -----------------------------------------------------------------------------
  -- ��    �ܣ�������渽��
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  procedure ClearReportAffix
  (
    ����id_In       ���Ӳ�������.����ID%TYPE,
    �������_IN     ���Ӳ�������.������%TYPE
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ���ӱ��渽��
  -- �����б�
  --
  --
  -----------------------------------------------------------------------------
  Procedure AddReportAffix
  (
    ����id_In In ���Ӳ�������.����id%Type,
    �ļ���_In In ���Ӳ�������.�ļ���%Type,
    ��С_In   In ���Ӳ�������.��С%Type,
    �������_IN in  ���Ӳ�������.������%TYPE
  );


  -----------------------------------------------------------------------------
  -- ��    �ܣ�����ĵ籨��ͼ��
  -- �����б�
  --����ͼ���м�¼ID
  --
  -----------------------------------------------------------------------------
  function AddElectrocardioReportImage
  (
    ҽ��id_IN       ����ҽ������.ҽ��ID%TYPE
  )return number;

end b_PacsInterface;
/
create or replace package body b_PacsInterface is

  -----------------------------------------------------------------------------
  --��ȡ��������Ϣ
  -----------------------------------------------------------------------------
  Procedure GetDeptItems
  (
  Cursor_Out  Out  t_Refcur,
  ��������_In  In  Varchar2:=Null
  ) is
  begin

    If ��������_In Is Null Then
      Open Cursor_Out For
        Select p.ID, P.����, p.����, p.����, p.λ�� from ���ű� P, ��������˵�� C where P.id = c.����id and c.�������� = '���';
    Else
      Open Cursor_Out For
        Select p.ID, P.����, p.����, p.����, p.λ�� from ���ű� P, ��������˵�� C where P.id = c.����id and c.�������� = '���'
        And (p.���� = ��������_In Or p.���� Like '%'||��������_In||'%');
    End If;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetDeptItems;


  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ�ѱ�
  -----------------------------------------------------------------------------
  Procedure GetChargeTypes
  (
    Cursor_Out  Out  t_Refcur,
    ��������_In  In  Varchar2:=Null
  ) is
  begin

    If ��������_In Is Null Then
      Open Cursor_Out For
        Select ����,����,ȱʡ��־ From �ѱ� a;

    Else
      Open Cursor_Out For
        Select ����,����,ȱʡ��־ From �ѱ� a
        Where (a.���� = ��������_In Or a.���� Like '%'||��������_In||'%');
    End If;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetChargeTypes;


  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡpacs�����Ŀ
  -----------------------------------------------------------------------------
  procedure GetPacsItems
  (
    Cursor_Out  Out  t_Refcur,
    ��������_In  In  Varchar2:=Null
  ) is
  begin

    If ��������_In Is Null Then
      Open Cursor_Out For
        Select /*+ RULE */
          Distinct a.Id as ������ĿID, a.����, a.����, Decode(a.�����Ա�, 1, '��', 2, 'Ů', 'ͨ��') �����Ա�, a.���㵥λ As ��λ,
                   Decode(a.�������, 1, '����', 2, 'סԺ', 'ͨ��') ���ó���, a.�������� �������, b.��λ ��鲿λ,
                   b.���� ��鷽��, (a.���� || '_' || b.��λ || '_' || b.���� ) as ��λ�������
          From ������ĿĿ¼ a, ������Ŀ��λ b
          Where Nvl(To_Char(a.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And a.��� = 'D' And a.Id = b.��Ŀid(+);

    Else
      Open Cursor_Out For
        Select /*+ RULE */
          Distinct a.Id as ������ĿID, a.����, a.����, Decode(a.�����Ա�, 1, '��', 2, 'Ů', 'ͨ��') �����Ա�, a.���㵥λ As ��λ,
                   Decode(a.�������, 1, '����', 2, 'סԺ', 'ͨ��') ���ó���, a.�������� �������, b.��λ ��鲿λ,
                   b.���� ��鷽��, (a.���� || '_' || b.��λ || '_' || b.���� ) as ��λ�������
          From ������ĿĿ¼ a, ������Ŀ��λ b
          Where Nvl(To_Char(a.����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And a.��� = 'D' And a.Id = b.��Ŀid(+)
          And (a.���� = ��������_In Or a.���� Like '%'||��������_In||'%');
    End If;

   Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetPacsItems;


    -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ�����Ŀ��ϸ
  -----------------------------------------------------------------------------
  procedure GetAdviceItems
  (
    Cursor_Out  Out  t_Refcur,
    ҽ��id_In  In  ����ҽ����¼.ID%Type
  ) Is
  Begin
    Open Cursor_Out For
          select /*+ RULE */ a.ID As ��λҽ��ID, a.������Ŀid, c.���� As ������Ŀ����, a.�걾��λ, a.��鷽��, Decode(c.��������,'X��','DR','MRI','MR',c.��������) as �������,
                (/*c.���� || '_' || */a.�걾��λ /*|| '_' */|| replace(replace(a.��鷽��,'(','') ,')','')) as ��λ�������
                from ����ҽ����¼ a, ������ĿĿ¼ c ,����ҽ������ b
                Where a.������Ŀid=c.id and  a.Id = b.ҽ��ID And b.ִ��״̬=0 And ���id=ҽ��id_In ;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetAdviceItems;

      -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ�����Ŀ����
  -----------------------------------------------------------------------------
  procedure GetAdviceFees
  (
    Cursor_Out  Out  t_Refcur,
    ҽ��id_In  In  ����ҽ����¼.ID%Type
  ) Is
  v_������Դ ����ҽ����¼.������Դ%Type;
  v_��¼���� ����ҽ������.��¼����%Type;
  v_������� ����ҽ������.�������%Type;
  v_���ݺ�   ����ҽ������.NO%Type;
  v_���ͺ�   ����ҽ������.���ͺ�%Type;
  strSQL varchar2(2000);
  strFeeTable Varchar2(20);

  Begin

   Select a.������Դ,b.��¼����,nvl(b.�������,0) As �������,b.NO,b.���ͺ� Into v_������Դ,v_��¼����,v_�������,v_���ݺ�,v_���ͺ�
           From ����ҽ����¼ a,����ҽ������ b
           Where a.id=b.ҽ��ID And a.id = ҽ��id_In;


    If v_������Դ = 2 And v_��¼���� = 2 And v_������� = 0 Then
        --�� "סԺ���ü�¼"
        strFeeTable :='סԺ���ü�¼';
    Else
        --�� "������ü�¼"
        strFeeTable :='������ü�¼';
    End If;

     strSQL := 'Select  ''������'' As ��������,decode(A.��¼����,1,''�շѵ���'',''���ʵ���'') As ��������,
                 A.NO As ���ݺ�,A.Ӧ�ս��,A.ʵ�ս��,A.���� || '' '' || A.���㵥λ as ����,
                      Decode(A.��¼����,1,Decode(A.��¼״̬,0,''�շѻ���'',1,''���շ�'',3,''���˷�''),2,
                          Decode(A.��¼״̬,0,''���ʻ���'',1,''�Ѽ���'',3,''������''),''δ�Ʒ�'') as �Ʒ�״̬,e.����|| '' ''|| e.��� as ��Ŀ
                 From ' || strFeeTable || ' A,����ҽ����¼ B ,����ҽ������ C,�շ���ĿĿ¼ E
                 Where A.NO= ''' || v_���ݺ� || ''' And A.��¼״̬ IN(0,1,3) And A.ҽ�����+0=B.ID And A.��¼����= ' || v_��¼����
                       || ' And  c.ҽ��ID=b.Id  And A.�շ�ϸĿID=E.Id
          Union ALL
          Select  ''���ӷ���'' As ��������,decode(B.��¼����,1,''�շѵ���'',''���ʵ���'') As ��������,
                 B.NO As ���ݺ�,B.Ӧ�ս��,B.ʵ�ս��,B.���� || '' '' || B.���㵥λ as ����,
                      Decode(B.��¼����,1,Decode(B.��¼״̬,0,''�շѻ���'',1,''���շ�'',3,''���˷�''),2,
                          Decode(B.��¼״̬,0,''���ʻ���'',1,''�Ѽ���'',3,''������''),''δ�Ʒ�'') as �Ʒ�״̬,e.����|| '' ''|| e.��� as ��Ŀ
                 From ����ҽ����¼ C,' || strFeeTable || ' B,����ҽ������ A ,����ҽ������ D ,�շ���ĿĿ¼ E
                 Where A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ��ID=B.ҽ�����+0
                       And A.ҽ��ID IN (Select ID From ����ҽ����¼ Where (ID= ' || ҽ��id_In || ' Or ���ID= ' ||
                       ҽ��id_In || ') )
                       And A.���ͺ�= ' || v_���ͺ� || ' And B.��¼״̬ IN(0,1,3) And A.ҽ��ID=C.ID And A.��¼����= ' ||
                       v_��¼���� || ' And d.ҽ��ID =c.Id  And B.�շ�ϸĿID=E.Id ';

     Open Cursor_Out For strSQL;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetAdviceFees;

  -----------------------------------------------------------------------------
  -- ��    �ܣ����ؿ���ҽ����Ϣ
  -----------------------------------------------------------------------------
  procedure GetPacsDeptDoctor
  (
    Cursor_Out  Out  t_Refcur,
    ��������_In  In  Varchar2:=Null
  ) is
  Begin
    --...
    null;
  end GetPacsDeptDoctor;


  -----------------------------------------------------------------------------
  -- ��    �ܣ���ȡ������Ϣ
  -----------------------------------------------------------------------------
  Procedure GetPatient
  (
    Cursor_Out  Out  t_Refcur,
    ���ҷ�ʽ_In  In  Number,
    ��������_In  In  Varchar2
  ) is
  begin
    If ���ҷ�ʽ_In=1 Then
      Open Cursor_Out For
        Select ����id,����,�Ա�,����,��������,replace(���֤��,'δ��','') As  ���֤��,����״��,����,����,/*ְҵ,����,*/ѧ��,��ϵ������,nvl(��ϵ�˵绰,��ͥ�绰) As ��ϵ�˵绰,nvl(��ͥ��ַ,��ϵ�˵�ַ) As ��ϵ�˵�ַ,������λ,���￨��,������,�����,סԺ��,�ѱ�,��ǰ���� From ������Ϣ Where ����id = zl_to_number(��������_In);
    ElsIf ���ҷ�ʽ_In=2 Then
      Open Cursor_Out For
        Select ����id,����,�Ա�,����,��������,replace(���֤��,'δ��','') As  ���֤��,����״��,����,����,/*ְҵ,����,*/ѧ��,��ϵ������,nvl(��ϵ�˵绰,��ͥ�绰) As ��ϵ�˵绰,nvl(��ͥ��ַ,��ϵ�˵�ַ) As ��ϵ�˵�ַ,������λ,���￨��,������,�����,סԺ��,�ѱ�,��ǰ���� From ������Ϣ Where סԺ�� = zl_to_number(��������_In);
    ElsIf ���ҷ�ʽ_In=3 Then
      Open Cursor_Out For
        Select ����id,����,�Ա�,����,��������,replace(���֤��,'δ��','') As  ���֤��,����״��,����,����,/*ְҵ,����,*/ѧ��,��ϵ������,nvl(��ϵ�˵绰,��ͥ�绰) As ��ϵ�˵绰,nvl(��ͥ��ַ,��ϵ�˵�ַ) As ��ϵ�˵�ַ,������λ,���￨��,������,�����,סԺ��,�ѱ�,��ǰ���� From ������Ϣ Where ����� = zl_to_number(��������_In);
    ElsIf ���ҷ�ʽ_In=4 Then
      Open Cursor_Out For
        Select ����id,����,�Ա�,����,��������,replace(���֤��,'δ��','') As  ���֤��,����״��,����,����,/*ְҵ,����,*/ѧ��,��ϵ������,nvl(��ϵ�˵绰,��ͥ�绰) As ��ϵ�˵绰,nvl(��ͥ��ַ,��ϵ�˵�ַ) As ��ϵ�˵�ַ,������λ,���￨��,������,�����,סԺ��,�ѱ�,��ǰ���� From ������Ϣ Where ���￨�� = ��������_In;
    ElsIf ���ҷ�ʽ_In=5 Then
      Open Cursor_Out For
        Select ����id,����,�Ա�,����,��������,replace(���֤��,'δ��','') As  ���֤��,����״��,����,����,/*ְҵ,����,*/ѧ��,��ϵ������,nvl(��ϵ�˵绰,��ͥ�绰) As ��ϵ�˵绰,nvl(��ͥ��ַ,��ϵ�˵�ַ) As ��ϵ�˵�ַ,������λ,���￨��,������,�����,סԺ��,�ѱ�,��ǰ���� From ������Ϣ Where ���֤�� = ��������_In;
    ElsIf ���ҷ�ʽ_In=6 Then
      Open Cursor_Out For
        Select ����id,����,�Ա�,����,��������,replace(���֤��,'δ��','') As  ���֤��,����״��,����,����,/*ְҵ,����,*/ѧ��,��ϵ������,nvl(��ϵ�˵绰,��ͥ�绰) As ��ϵ�˵绰,nvl(��ͥ��ַ,��ϵ�˵�ַ) As ��ϵ�˵�ַ,������λ,���￨��,������,�����,סԺ��,�ѱ�,��ǰ���� From ������Ϣ Where ������ = ��������_In;
    ElsIf ���ҷ�ʽ_In=7 Then
      Open Cursor_Out For
        Select ����id,����,�Ա�,����,��������,replace(���֤��,'δ��','') As  ���֤��,����״��,����,����,/*ְҵ,����,*/ѧ��,��ϵ������,nvl(��ϵ�˵绰,��ͥ�绰) As ��ϵ�˵绰,nvl(��ͥ��ַ,��ϵ�˵�ַ) As ��ϵ�˵�ַ,������λ,���￨��,������,�����,סԺ��,�ѱ�,��ǰ���� From ������Ϣ Where ���� Like '%'||��������_In||'%';
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetPatient;


  ---------------------------------------------------------------------------------------------------------------
  -- ��    �ܣ���ȡҽ������״̬
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestStatus
  (
    Cursor_Out  Out  t_Refcur,
    ҽ��id_In  In  ����ҽ����¼.ID%Type
  )is
  begin
    Open Cursor_Out For
      Select a.ҽ��״̬,b.ִ��״̬, b.ִ�й���
      From ����ҽ����¼ a,����ҽ������ b
      Where a.ID=b.ҽ��id and a.ID=ҽ��id_In And RowNum<2;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetRequestStatus;


  ---------------------------------------------------------------------------------------------------------------
  -- ��    �ܣ���ȡ���������Ϣ
  ---------------------------------------------------------------------------------------------------------------
  Procedure GetRequestInfo
  (
    Cursor_Out  Out  t_Refcur,
    ���ҷ�ʽ_In  In  Number,
    ��������_In  In  Varchar2,
    ��������_In  In  Varchar2:=null
  ) is
    strSql varchar2(2000);
    v_����ժҪҪ��ID ����ҽ������.Ҫ��ID%Type;
    v_�ٴ����Ҫ��ID ����ҽ������.Ҫ��ID%Type;
  Begin


    select ID into v_����ժҪҪ��ID from ����������Ŀ where  ����������Ŀ.������='��������';
    select ID Into v_�ٴ����Ҫ��ID from ����������Ŀ where  ����������Ŀ.������='������';

    If ���ҷ�ʽ_In=1 Then
      strSql :=  ' c.����id =' || zl_to_number(��������_In);
    ElsIf ���ҷ�ʽ_In=2 Then
      strSql := ' c.סԺ�� =' || zl_to_number(��������_In);
    ElsIf ���ҷ�ʽ_In=3 Then
      strSql := ' c.����� =' || zl_to_number(��������_In);
    ElsIf ���ҷ�ʽ_In=4 Then
      strSql := ' c.���￨�� =''' || ��������_In || '''';
    ElsIf ���ҷ�ʽ_In=5 Then
      strSql := ' c.���֤�� =''' || ��������_In || '''';
    ElsIf ���ҷ�ʽ_In=6 Then
      strSql := ' c.������ =''' || ��������_In || '''';
    ElsIf ���ҷ�ʽ_In=7 Then
      strSql := ' c.���� like ''%' || ��������_In  || '%''';
    ElsIf ���ҷ�ʽ_In=8 Then
      strSql := ' a.Id =' || zl_to_number(��������_In);
    End If;

    if Nvl(��������_In,'')<>'' then
      strSql :='And '||��������_In;
    end if;
--����ǰ����δִ�еļ����Ŀ
    strSql := 'Select Distinct nvl(a.���ID,a.Id ) As ҽ��ID,c.����,c.�����,c.סԺ��,c.�Ա�,c.����
           From ����ҽ����¼ a, ����ҽ������ b, ������Ϣ c ,
                (Select ����ID From ��������˵�� Where �������� =''���'') d
           Where b.ִ�в���ID = d.����ID
                 And (b.ִ�й��� Is Null Or b.ִ�й���=1 Or b.ִ�й��� = 0) And b.ִ��״̬ = 0
                 And b.����ʱ�� > To_Date(To_Char(Sysdate - 3,''yyyy-mm-dd'') || ''23:59:59'',''yyyy-mm-dd hh24:mi:ss'')
                 And a.������� = ''D'' And a.Id = b.ҽ��ID
                 And a.����ID = c.����ID and ' || strSql;


    strSql :='Select /*+ RULE */
        k.ҽ��id,m.��ҳid,m.��������ID As �������ID,p.���� As �������,m.����ҽ�� As ������,
        m.����ʱ�� As ����ʱ��,replace(m.ҽ������,'','',''|'') as ҽ������,m.������Ŀid,m.ִ�п���ID As ִ�в���ID ,n.���� As ִ�в���,
        m.����id,k.����,k.�����,k.סԺ��,k.�Ա�,k.����,
        Decode(m.������Դ, 1, ''����'', 2, ''סԺ'', 3, ''����'', 4, ''���'')  as ������Դ,
        Decode(m.������־,1,1,nvl((Select ���� From ���˹Һż�¼ Where No = m.�Һŵ�),0)) As ������־,
        (select ���� from ����ҽ������ where Ҫ��ID=' || v_����ժҪҪ��ID || ' and ҽ��ID= m.id ) as ����ժҪ,
        (select ���� from ����ҽ������ where Ҫ��ID=' || v_�ٴ����Ҫ��ID || ' and ҽ��ID=  m.id) as �ٴ����
      From ( ' || strSql || ' )  k ,����ҽ����¼ m ,���ű� n,���ű� p
      Where k.ҽ��ID = m.Id And m.ִ�п���ID = n.Id And m.��������ID = p.Id';

   Open Cursor_Out For strSql;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetRequestInfo;


  ---------------------------------------------------------------------------------------------------------------
	-- ��    �ܣ���ȡ�ĵ����������Ϣ
	---------------------------------------------------------------------------------------------------------------
	Procedure GetRequestInfo1
	(
		Cursor_Out	Out	t_Refcur,
		��ʼ����_In	In	Varchar2,
		��������_In	In	Varchar2,
    ������_In In  Varchar2
	)is
    strSql varchar2(2000);
    v_����ժҪҪ��ID ����ҽ������.Ҫ��ID%Type;
    v_�ٴ����Ҫ��ID ����ҽ������.Ҫ��ID%Type;
  Begin


    select ID into v_����ժҪҪ��ID from ����������Ŀ where  ����������Ŀ.������='��������';
    select ID Into v_�ٴ����Ҫ��ID from ����������Ŀ where  ����������Ŀ.������='������';

    strSql := 'Select Distinct nvl(a.���ID,a.Id ) As ҽ��ID,c.����,c.�����,c.סԺ��,c.�Ա�,c.����
           From ����ҽ����¼ a, ����ҽ������ b, ������Ϣ c ,������ĿĿ¼ e,
                (Select ����ID From ��������˵�� Where �������� =''���'') d
           Where b.ִ�в���ID = d.����ID
                 And (b.ִ�й��� Is Null Or b.ִ�й���=1 Or b.ִ�й��� = 0) And b.ִ��״̬ = 0
                 And a.������� = ''D'' And a.Id = b.ҽ��ID and a.������ĿID=e.id and e.�������� like ''%' || ������_In || '%''
                 And a.����ID = c.����ID and b.����ʱ�� between to_date(''' || ��ʼ����_In || ''', ''yyyy-mm-dd hh24:mi:ss'')  and  to_date(''' || ��������_In || ''', ''yyyy-mm-dd hh24:mi:ss'')';


    strSql :='Select /*+ RULE */
        k.ҽ��id,m.��ҳid,m.��������ID As �������ID,p.���� As �������,m.����ҽ�� As ������,
        m.����ʱ�� As ����ʱ��,m.ҽ������,m.������Ŀid,m.ִ�п���ID As ִ�в���ID ,n.���� As ִ�в���,
        m.����id,k.����,k.�����,k.סԺ��,k.�Ա�,k.����,
        Decode(m.������Դ, 1, ''����'', 2, ''סԺ'', 3, ''����'', 4, ''���'') ������Դ,
        Decode(m.������־,1,1,nvl((Select ���� From ���˹Һż�¼ Where No = m.�Һŵ�),0)) As ������־,
        (select ���� from ����ҽ������ where Ҫ��ID=' || v_����ժҪҪ��ID || ' and ҽ��ID= m.id ) as ����ժҪ,
        (select ���� from ����ҽ������ where Ҫ��ID=' || v_�ٴ����Ҫ��ID || ' and ҽ��ID=  m.id) as �ٴ����
      From ( ' || strSql || ' )  k ,����ҽ����¼ m ,���ű� n,���ű� p
      Where k.ҽ��ID = m.Id And m.ִ�п���ID = n.Id And m.��������ID = p.Id';


   Open Cursor_Out For strSql;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end GetRequestInfo1;


  -----------------------------------------------------------------------------
  -- ��    �ܣ�ȡ������/�������
  -----------------------------------------------------------------------------
  Procedure CancelRequest
  (
    ҽ��id_In  In  ����ҽ������.ҽ��ID%Type,
    ����ִ��_In   Number := 0,
    ִ�в���ID_IN ���ű�.id%Type := 0
    --������ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  ) is
    v_���ͺ� Ӱ�����¼.���ͺ�%Type;
  Begin

    select ���ͺ� into V_���ͺ� from ����ҽ������ where ҽ��ID=ҽ��id_In;

    Zl_Ӱ����_Cancel(ҽ��id_In, v_���ͺ�,����ִ��_In,ִ�в���ID_IN);

  EXCEPTION
    WHEN OTHERS THEN
      zl_ErrorCenter (SQLCODE, SQLERRM);
  end CancelRequest;


  -----------------------------------------------------------------------------
  -- ��    �ܣ�ɾ��������Ϣ
  -----------------------------------------------------------------------------
  PROCEDURE DeleteReport
  (
    ҽ��id_IN ����ҽ������.ҽ��ID%TYPE
  )Is
     v_Count         Number;
     v_��ҽ��ID      ����ҽ������.ҽ��ID%Type;
  Begin
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��ID From ����ҽ����¼ Where ID = ҽ��id_In;

    --���������
    ClearPacsReport(v_��ҽ��ID);

    Zl_Ӱ�񱨸���_Clear(v_��ҽ��ID);

    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ�ļ�����룬ɾ������󲻸���ִ��״̬
    Select Count(*) Into v_Count From ����ҽ����¼ a, ������ҳ b
    Where  a.����ID=b.����ID And a.��ҳID = b.��ҳID And b.��Ժ���� Is Not Null And a.Id = v_��ҽ��ID;

    If v_Count =0 Then
       --ɾ�����棬��ȡ��ҽ�����״̬
       Update ����ҽ������
       Set ִ��״̬ = 0, ִ�й��� = 2
       Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = v_��ҽ��ID Or ���id = v_��ҽ��ID))
             And ִ��״̬ = 1;
    End If;

  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
  END DeleteReport;



  -----------------------------------------------------------------------------
  -- ��    �ܣ������������
  -----------------------------------------------------------------------------
  PROCEDURE ClearPacsReport
  (
    ҽ��id_IN ����ҽ������.ҽ��ID%TYPE
  )IS
  BEGIN
    --������ر��м���ɾ������,������Ӳ�����¼һ��ɾ��
    Delete ���Ӳ�����¼ Where Id In (Select ����ID From ����ҽ������ Where ҽ��ID=ҽ��id_IN);
  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
  END ClearPacsReport;


  -----------------------------------------------------------------------------
  -- ��    �ܣ����ռ���/�������
  -----------------------------------------------------------------------------
  Procedure RecevieRequest
  (
    ҽ��id_IN ����ҽ������.ҽ��ID%TYPE,
    ִ�м�_IN ����ҽ������.ִ�м�%TYPE:=Null,
    ����_IN Ӱ�����¼.����%TYPE:=NULL,
    ����豸_IN Ӱ�����¼.����豸%TYPE:=Null,
    ���_IN Ӱ�����¼.���%TYPE:=Null,
    ����_IN Ӱ�����¼.����%TYPE:=Null,
    ��鼼ʦ_IN Ӱ�����¼.��鼼ʦ%TYPE:=Null,
    ִ��ʱ��_IN ����ҽ������.����ʱ��%TYPE:=Null,
    ִ��˵��_IN ����ҽ������.ִ��˵��%TYPE:=NULL,
    ����ִ��_In   Number := 0,
    ִ�в���ID_IN ���ű�.id%Type := 0
    --������ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  ) Is

    Cursor c_AdviceInfo Is
       Select ID, ���id, Nvl(���id, ID) As ��id, �������, ������Դ,ִ�п���ID From ����ҽ����¼ Where ID = ҽ��id_In;
    r_AdviceInfo c_AdviceInfo%Rowtype;

    v_ԭ���� Ӱ�����¼.����%Type;
    v_�¼��� Ӱ�����¼.����%Type;
    v_����   Ӱ�����¼.����%Type;
    v_Ӣ���� Ӱ�����¼.Ӣ����%Type;
    v_Ӱ����� Ӱ�����¼.Ӱ�����%Type;
    v_�����ͺ� Ӱ�����¼.���ͺ�%Type;
    v_���ͺ� Ӱ�����¼.���ͺ�%Type;
    v_������Դ ����ҽ����¼.������Դ%Type;
    v_��Ա��� ��Ա��.���%Type;
    v_��Ա���� ��Ա��.����%Type;
    v_Count Number;
    v_Error Varchar2(255);
    Err_Custom Exception;

  Begin
    --��ȡҽ������ҽ��ID������ID
    Open c_AdviceInfo;
         Fetch c_AdviceInfo
               Into r_AdviceInfo;
    Close c_AdviceInfo;

    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ���߳�Ժ�ļ�����룬������ʼ������˷���
    Select Count(*) Into v_Count From ����ҽ����¼ a, ������ҳ b
    Where  a.����ID=b.����ID And a.��ҳID = b.��ҳID And (b.��Ժ���� Is Not Null Or b.״̬ = 3)
       And a.Id = r_AdviceInfo.��id;

    If v_Count >0 Then
      v_Error := 'סԺ�����Ѿ���Ժ����Ԥ��Ժ���޷���ʼ��顣';
      Raise Err_Custom;
    End If;

    --��ʼִ��ҽ��
    If Nvl(����ִ��_In, 0) = 1 Then
       -- ������λҽ������ִ��
       Update ����ҽ������
       Set �״�ʱ�� = Sysdate, ĩ��ʱ�� = Sysdate,ִ��״̬ =3,ִ�м� = ִ�м�_In, ����ʱ�� = ִ��ʱ��_IN,
           ִ��˵�� = ִ��˵��_IN
       Where ҽ��ID = ҽ��id_In;
    Else
       Update ����ҽ������
       Set �״�ʱ�� = Sysdate,ĩ��ʱ�� = Sysdate, ִ��״̬ = 3,ִ�м� = ִ�м�_In,����ʱ�� = ִ��ʱ��_IN,
           ִ��˵�� = ִ��˵��_IN
       Where ҽ��ID In (Select ID From ����ҽ����¼ Where (ID = r_AdviceInfo.��ID Or ���ID = r_AdviceInfo.��ID));
    End If;

    --�������Ա�����ͱ�ţ���� ��鼼ʦ_IN Ϊ�գ�����д user
    If ��鼼ʦ_IN Is Null Then
       v_��Ա���� := User;
       v_��Ա��� := User;
    Else
       Begin
            Select ���,���� Into v_��Ա���,v_��Ա���� From ��Ա�� a,������Ա b
            Where a.Id = b.��ԱID And b.����ID=r_AdviceInfo.ִ�п���ID And a.����=��鼼ʦ_IN And Rownum =1;
       Exception
            When Others Then
                 v_��Ա���� := User;
                 v_��Ա��� := User;
       End;
    End If;
    --�������
    Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��ID = ҽ��id_IN;
    zl_Ӱ�����ִ��(ҽ��id_IN, v_���ͺ�, 2,����ִ��_In,v_��Ա���,v_��Ա����,ִ�в���ID_IN);

    --��ȡ��ҽ�������Ϣ
    Select A.���ͺ�,C.����,zlspellcode(C.����) Ӣ����,D.��������,B.������Դ
    Into v_�����ͺ�,v_����,v_Ӣ����,v_Ӱ�����,v_������Դ
    From ����ҽ������ A,����ҽ����¼ B,������Ϣ C, ������ĿĿ¼ D
    Where A.ҽ��ID=B.id And B.id = r_AdviceInfo.��ID And B.����ID = C.����ID And B.������ĿID = D.ID;

    --�������
    If ����_IN Is Null Then --û����������HIS������������¼���
      begin
        Select /*+ rule */ ���� Into v_ԭ���� From Ӱ�����¼ Where ҽ��id = r_AdviceInfo.��ID;
      Exception
        When Others Then
          Select ������+1 Into v_�¼��� From Ӱ������� Where ����=v_Ӱ�����;
      End;
    End If;

    Update /*+ RULE */ Ӱ�����¼
    Set Ӱ����� = v_Ӱ�����, ���� = NVL(Nvl(����_In, v_ԭ����),v_�¼���), ���� = v_����, Ӣ���� = v_Ӣ����, ��� = ���_In,
        ���� = ����_In, ����豸 = ����豸_In, ��鼼ʦ = ��鼼ʦ_In
    Where ҽ��id = r_AdviceInfo.��ID;

    If Sql%Rowcount = 0 Then
      Insert Into Ӱ�����¼(ҽ��id, ���ͺ�, Ӱ�����, ����, ����, Ӣ����, ���, ����, ����豸, ��鼼ʦ)
      Values(r_AdviceInfo.��ID, v_�����ͺ�, v_Ӱ�����, NVL(Nvl(����_In, v_ԭ����),v_�¼���),
             v_����, v_Ӣ����, ���_In, ����_In, ����豸_In, ��鼼ʦ_In);
    End If;

    If v_�¼��� Is NOT Null Then
      Update Ӱ������� Set ������ = v_�¼��� Where ���� = v_Ӱ�����;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end RecevieRequest;


  -----------------------------------------------------------------------------
  -- ��    �ܣ����ͱ����ı���Ϣ
  -----------------------------------------------------------------------------
  procedure SendReport
  (
    ҽ��id_IN       ����ҽ������.ҽ��ID%TYPE,
    ��������_IN     ���Ӳ�������.�����ı�%TYPE,
    ���潨��_IN     ���Ӳ�������.�����ı�%TYPE,
    ����ҽ��_IN     ���Ӳ�����¼.������%TYPE,
    ���ҽ��_IN     Ӱ�����¼.������%TYPE := Null,
    ִ�в���ID_IN ���ű�.id%Type := 0
  )Is

    --��ȡ����ҽ��������������Ϣ
    CURSOR c_Advice(v_��ID Number) IS
        Select E.Id,E.������Դ,E.����ID,E.��ҳID,E.Ӥ��,E.���˿���ID,E.�ļ�id, E.��������,E.��������,F.����ID,E.ִ�п���ID
        From (Select C.ID,C.������Դ,C.����ID,C.��ҳID,C.Ӥ��,C.���˿���ID,C.�ļ�id, D.���� ��������, D.���� ��������,C.ִ�п���ID
          From (Select A.ID,A.������Դ,A.����ID,A.��ҳID,A.Ӥ��,A.���˿���ID, B.�����ļ�id �ļ�id,A.ִ�п���ID
                     From ����ҽ����¼ A, ��������Ӧ�� B
                     Where A.Id=v_��ID And A.������Ŀid = B.������Ŀid(+) And B.Ӧ�ó���(+) = Decode(A.������Դ, 2, 2, 4, 4, 1)) C,�����ļ��б� D
          Where C.�ļ�id = D.Id(+)) E,����ҽ������ F
        Where E.Id=F.ҽ��ID(+);

    --�����ļ������Ԫ��
    CURSOR c_File(v_File number) IS
        Select A.Id, A.�ļ�id, A.��id, A.�������, A.��������, A.������, A.��������, A.��������, A.�����д�,
               A.�����ı�, A.�Ƿ���, A.Ԥ�����id, A.�������, A.ʹ��ʱ��, A.����Ҫ��id, A.�滻��, A.Ҫ������,
               A.Ҫ������, A.Ҫ�س���, A.Ҫ��С��, A.Ҫ�ص�λ, A.Ҫ�ر�ʾ, A.������̬, A.Ҫ��ֵ��
        From �����ļ��ṹ A
        Where A.�ļ�id = v_File
        Order By A.�������;

    Cursor c_Report(v_���Ӳ�����¼ID Number) Is
        Select /*+ rule */ B.Id, A.�����ı�
               From ���Ӳ������� A, ���Ӳ������� B
               Where A.�ļ�id = v_���Ӳ�����¼ID And Nvl(A.�������id, 0) <> 0 And
                     (A.�����ı� like '%����%' Or A.�����ı� like '%����%' Or A.�����ı� like '%����ҽ��%' ) And
                     B.��id = A.Id And B.�Ƿ��� = 1;

    Cursor c_ExecutAdvice(v_��ID Number) Is
         Select ҽ��ID,���ͺ� From ����ҽ����¼ a,����ҽ������ b
         Where a.ID=b.ҽ��ID And (a.id =v_��ID Or a.���ID =v_��ID ) And b.ִ��״̬ = 3;
    r_ExecutAdvice c_ExecutAdvice%Rowtype;

    r_Advice      c_Advice%Rowtype;
    v_����id      ���Ӳ�������.�ļ�ID%Type;
    v_��������id  ���Ӳ�������.Id%Type;
    v_��������idNew  ���Ӳ�������.Id%Type;
    v_�������    ���Ӳ�������.�������%Type;
    v_��ID        ���Ӳ�������.��ID%Type;
    v_�����ı�    ���Ӳ�������.�����ı�%Type;
    v_�������ID  ���Ӳ�������.�������ID%Type;
    --v_��ʽ����    ���Ӳ�����ʽ.����%Type;
    v_Error         Varchar2(255);
    Err_Custom      Exception;
    v_Count         Number;
    v_��ҽ��ID      ����ҽ������.ҽ��ID%Type;
    v_��Ա���      ��Ա��.���%Type;
    v_��Ա����      ��Ա��.����%Type;
    v_����ʱ��     ���Ӳ�����¼.����ʱ��%Type;
  Begin

    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��ID From ����ҽ����¼ Where ID = ҽ��id_In;

    Open c_Advice(v_��ҽ��ID);
      Fetch c_Advice Into r_Advice;

    If Nvl(r_Advice.�ļ�ID,0)=0 Then
        v_Error:='���μ����Ŀû�ж�Ӧ��صļ�鱨�棬�������Ա��ϵ��';
        Raise Err_Custom;
    Else
        If Nvl(r_Advice.����id,0)>0 Then  ----����������
            --�ҳ��������д�ı�������к���'%����%','%����%,'%����%','%���%',���ô���Ĳ�������
            For r_Report In c_Report(r_Advice.����id) Loop
                If r_Report.�����ı� like '%����%' Then
                    Update ���Ӳ������� Set �����ı�=��������_IN Where ID=r_Report.Id;
                Elsif r_Report.�����ı� like '%����%' Then
                    Update ���Ӳ������� Set �����ı�=���潨��_IN Where ID=r_Report.Id;
                Elsif r_Report.�����ı� like '%����ҽ��%' Then
                    Update ���Ӳ������� Set �����ı�=����ҽ��_IN Where ID=r_Report.Id;
                --Elsif r_Report.�����ı� like '%����ʱ��%' Then
                    --Update ���Ӳ������� Set �����ı�=���潨��_IN Where ID=r_Report.Id;
                End If;
            End Loop;
            --���±���ʱ��
            Update ���Ӳ�����¼ Set ���ʱ��=Sysdate,������=����ҽ��_IN,����ʱ��=Sysdate Where ID=r_Advice.����id;
        Else
            --�������Ӳ�����¼
            Select ���Ӳ�����¼_ID.Nextval Into v_����id From Dual;
            Insert Into ���Ӳ�����¼
              (Id, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ���ʱ��,
               ������, ����ʱ��, ���汾, ǩ������)
            Values
              (v_����id, r_Advice.������Դ, r_Advice.����id, r_Advice.��ҳid, r_Advice.Ӥ��, r_Advice.���˿���id,
               r_Advice.��������, r_Advice.�ļ�id, r_Advice.��������, ����ҽ��_IN, Sysdate, Sysdate, ����ҽ��_IN, Sysdate, 1, 2);

            --����ҽ�������¼
            Insert Into ����ҽ������ (ҽ��ID,����ID) Values(v_��ҽ��ID,v_����ID);
            --���뱨��ʱ��
            Select to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi')  Into v_����ʱ�� From ���Ӳ�����¼ a,����ҽ������ b Where a.id=b.����id And b.ҽ��id=ҽ��id_In;
            --�²�����������
            For r_File In c_File(r_Advice.�ļ�ID) Loop
                Select ���Ӳ�������_ID.Nextval Into v_��������id From Dual;
                If nvl(v_�������,0)=0 Then
                   v_�������:=r_File.�������;
                Else
                   v_�������:=v_�������+1;
                End If;

                If NVL(r_File.��ID,0)<>0 And (r_File.�����ı� like '%����%' Or r_File.�����ı� like '%����%') Then--����������(�����)
                     v_�����ı�:=chr(32)||chr(32)||chr(32)||��������_IN || Chr(13) || Chr(13);
                     v_�������ID:=0;
                Elsif NVL(r_File.��ID,0)<>0 And (r_File.�����ı� like '%����%' Or r_File.�����ı� like '%���%') Then--���鶨����(�����)
                     v_�����ı�:=chr(32)||chr(32)||chr(32)||���潨��_IN || Chr(13) || Chr(13);
                     v_�������ID:=0;
                Elsif Nvl(r_File.��id, 0) <> 0 And (r_File.�����ı� Like '%����ҽ��%') Then--����ҽ��������(�����)
 				            v_�����ı�   := '����ҽ��: ' || ����ҽ��_IN || Chr(13) || Chr(13);
					          v_�������id := 0;
                Elsif Nvl(r_File.��id, 0) <> 0 And (r_File.�����ı� Like '%����ʱ��%') Then--����ʱ�䶨����(�����)
 				            v_�����ı�   := '����ʱ��: ' || v_����ʱ�� || Chr(13) || Chr(10)||'���˱�������ٴ����Ҳ鿴���,�����Դ�ӡ��ֽ�ʱ��浥Ϊ׼����';
					          v_�������id := 0;
                Elsif nvl(r_File.��������,0)=1 And NVL(r_File.��ID,0)=0 Then--��ٶ�����
                     v_��ID:=v_��������id;
                     v_�����ı�:=r_File.�����ı�;
                     v_�������ID:=r_File.id;
                Elsif nvl(r_File.��������,0)=4 And r_File.Ҫ������ Is Not Null Then  --�Զ��滻Ҫ��
                     v_�����ı�:=zl_replace_element_value(r_File.Ҫ������,r_Advice.����ID,r_Advice.��ҳID,r_Advice.������Դ,r_Advice.Id);
                     v_�������ID:=0;
                Else
                    v_�����ı�:=r_File.�����ı�;
                    v_�������ID:=0;
                End If;

                --�������ݵ���дһ��
                If NVL(r_File.��ID,0)<>0 And (r_File.�����ı� like '%����%' Or r_File.�����ı� like '%����%') Then--��д�����ʾ���ƣ���д���ݣ�ͬʱ������ŷ����仯
                   Select ���Ӳ�������_ID.Nextval Into v_��������idNew From Dual;
                   v_������� := v_������� + 1;
                    Insert Into ���Ӳ�������
                      (Id, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�,
                       �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������,
                       Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
                    Values
                      (v_��������idNew, v_����id, 0, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������,
                       r_File.������, r_File.��������, 0, Null, v_�����ı�, r_File.�Ƿ���,
                       r_File.Ԥ�����id, r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id, r_File.�滻��,
                       r_File.Ҫ������, r_File.Ҫ������, r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ,
                       r_File.Ҫ�ر�ʾ, r_File.������̬, r_File.Ҫ��ֵ��, Decode(v_�������id, 0, Null, v_�������id));
                    v_������� := v_������� - 1;
                    v_�����ı�:=r_File.�����ı�;
                End If;

                Insert Into ���Ӳ�������
                  (Id, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�,
                   �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���,
                   Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
                Values
                  (v_��������id, v_����id, 1, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������,
                   r_File.������, r_File.��������, r_File.��������, Null, v_�����ı�, r_File.�Ƿ���, r_File.Ԥ�����id,
                   r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id, r_File.�滻��, r_File.Ҫ������, r_File.Ҫ������,
                   r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ, r_File.Ҫ�ر�ʾ, r_File.������̬, r_File.Ҫ��ֵ��,
                   Decode(v_�������id, 0, Null, v_�������id));
             End Loop;

        /* ����Ӳ�����ʽ�к����������ָ�ʽ�����ַ�������֮���������ֽ����ɼ�
        Select ���� Into v_��ʽ���� From �����ļ���ʽ Where �ļ�ID=r_Advice.�ļ�ID;
        Insert Into ���Ӳ�����ʽ (�ļ�ID,����) Values (v_����id,v_��ʽ����);
        */

        End If;

        --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ�ļ�����룬��ӱ���󲻸���ִ��״̬
        Select Count(*) Into v_Count From ����ҽ����¼ a, ������ҳ b
        Where  a.����ID=b.����ID And a.��ҳID = b.��ҳID And b.��Ժ���� Is Not Null And a.Id = v_��ҽ��ID;

        If v_Count =0 Then
           --ֻ���Ѿ��������룬����ִ�е�ҽ���Ÿ��£�����Ϊ ���״̬����˹���
           Update ����ҽ������ Set ִ��״̬=1, ִ�й���=6, ���ʱ��=sysdate
           Where ҽ��id in(select id from ����ҽ����¼ where id= v_��ҽ��ID or ���id=v_��ҽ��ID);
                 --And ִ��״̬ = 3 ;--����Ҫִ�С��������롱����˲���Ҫ�ж�ִ��״̬��ֱ�Ӹ��������

           --�������Ա�����ͱ�ţ���� ��鼼ʦ_IN Ϊ�գ�����д user
           If ����ҽ��_IN Is Null Then
              v_��Ա���� := User;
              v_��Ա��� := User;
           Else
               Begin
                    Select ���,���� Into v_��Ա���,v_��Ա���� From ��Ա�� a,������Ա b
                    Where a.Id = b.��ԱID And b.����ID=r_Advice.ִ�п���ID And a.����=����ҽ��_IN And Rownum =1;
               Exception
                    When Others Then
                         v_��Ա���� := User;
                         v_��Ա��� := User;
               End;
           End If;

           --�������
           For r_ExecutAdvice In c_ExecutAdvice(v_��ҽ��ID) Loop
               zl_Ӱ�����ִ��(r_ExecutAdvice.ҽ��ID,r_ExecutAdvice.���ͺ� , 6,1,v_��Ա���,v_��Ա����,ִ�в���ID_IN);
           End Loop;
        End If;

        Update Ӱ�����¼ set ������=����ҽ��_IN, ������=���ҽ��_IN where ҽ��id=v_��ҽ��ID;
      End If;
      Close c_Advice;
    Exception
      When Err_Custom Then
        Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
      When Others Then
        Zl_ErrorCenter(Sqlcode, Sqlerrm);
  end SendReport;


  -----------------------------------------------------------------------------
  -- ��    �ܣ�ɾ���ĵ籨����Ϣ
  -----------------------------------------------------------------------------
  PROCEDURE DeleteElectrocardioReport
  (
    ҽ��id_IN  ����ҽ������.ҽ��ID%TYPE
  )Is
     v_Count         Number;
     v_��ҽ��ID      ����ҽ������.ҽ��ID%Type;
  Begin
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��ID From ����ҽ����¼ Where ID = ҽ��id_In;

    --���������
    Delete ���Ӳ�����¼ Where Id In (Select ����ID From ����ҽ������ Where ҽ��ID=v_��ҽ��ID);


    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ�ļ�����룬ɾ������󲻸���ִ��״̬
    Select Count(*) Into v_Count From ����ҽ����¼ a, ������ҳ b
    Where  a.����ID=b.����ID And a.��ҳID = b.��ҳID And b.��Ժ���� Is Not Null And a.Id = v_��ҽ��ID;

    If v_Count =0 Then
       --ɾ�����棬��ȡ��ҽ�����״̬
       Update ����ҽ������
       Set ִ��״̬ = 3, ִ�й��� = 2
       Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = v_��ҽ��ID Or ���id = v_��ҽ��ID))
             And ִ��״̬ = 1;
    End If;

  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
  END DeleteElectrocardioReport;



  -----------------------------------------------------------------------------
  -- ��    �ܣ������ĵ籨����Ϣ
  -----------------------------------------------------------------------------
  procedure SendElectrocardioReport
  (
    ҽ��id_IN       ����ҽ������.ҽ��ID%TYPE,
    �������_IN     ���Ӳ�������.�����ı�%TYPE,
    ��Ͻ��_IN     ���Ӳ�������.�����ı�%TYPE,
    ��Ͻ���_IN     ���Ӳ�������.�����ı�%TYPE,
    ����ҽ��_IN     ���Ӳ�����¼.������%TYPE,
    ���ҽ��_IN     Ӱ�����¼.������%TYPE := Null
  )is
    cursor c_AdviceInf(v_��ID Number) is
           select A.������Դ,A.����ID,A.��ҳID,A.��������ID,A.����,A.�Ա�,A.����,B.�����,B.סԺ��
           from ����ҽ����¼ A,������Ϣ B
           where A.����ID = b.����id and a.id =v_��ID;

    r_AdviceInf  c_AdviceInf%RowType;
    v_����ID     ����ҽ������.����ID%Type;
    v_StudyInf   Varchar2(2048);

    v_Count         Number;
    v_��ҽ��ID      ����ҽ������.ҽ��ID%Type;
    v_��ʽID        �����ļ��б�.id%Type;
    v_���          �����ļ��б�.���%type;
    v_��ʽ��     varchar2(255);

    v_Error      varchar2(255);
    Err_Custom   Exception;
    v_��ID   number;
  begin
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��ID From ����ҽ����¼ Where ID = ҽ��id_In;

    open c_AdviceInf(v_��ҽ��ID);
    fetch c_AdviceInf into r_AdviceInf;

    if c_AdviceInf%Rowcount = 0 then
       close c_AdviceInf;
       v_Error := '����ҽ����¼��ȡʧ�ܣ����鴫�ݵ�ҽ��ID�Ƿ���ȷ��';
       raise Err_Custom;
    end if;


    begin
      select Id into v_��ʽID from �����ļ��б� where ����='�ĵ籨���ʽ';
    exception
      When Others Then
        v_��ʽID := 0;
    end;

    --�����ĵ籨���ʽ
    if v_��ʽID = 0 then
       select �����ļ��б�_ID.NEXTVAL into v_��ʽID from dual;
       select max(���)+1 into v_��� from �����ļ��б�;

       v_��ʽ�� :=  '�ĵ籨���ʽ';

       insert into �����ļ��б�(id,����,���,����,ҳ��)
       values(v_��ʽID, 7, v_���, v_��ʽ��, v_���);


       --��ʽ˵���� ֽ�Ŵ�С;����;ֽ�Ÿ߶�;ֽ�ſ��;��߾�;�ұ߾�;�ϱ߾�;�±߾�;���ֱ���ɫ;ֽ�ű���ɫ;��ʾҳ��
       --ԭʼ��ʽ��9;1;16840;11907;849;849;1587;1417;10070188;16777215;1;
       --����Ҫ������ҳ������ʱ������plSql�е���ִ��������䣺
       --update ����ҳ���ʽ set ��ʽ='256;1;16840;16442;283;283;482;283;10070188;16777215;1' where ����='�ĵ籨���ʽ'
       insert into ����ҳ���ʽ(����,���,����,��ʽ)
       values(7, v_���,v_��ʽ��,'256;1;20840;11907;240;240;1587;1417;10070188;16777215;1');
    end if;

    --���ɲ���ID
    select ���Ӳ�����¼_ID.Nextval into v_����ID from dual;



    --������Ӳ�����¼
    insert into ���Ӳ�����¼(ID,������Դ,����ID,��ҳID,����ID,��������,�ļ�ID,��������,���ʱ��,������,����ʱ��,������,����ʱ��)
      select v_����ID,r_AdviceInf.������Դ,r_AdviceInf.����ID,r_AdviceInf.��ҳID,r_AdviceInf.��������ID,
             7,v_��ʽID,'�ĵ��鱨�浥',sysdate,����ҽ��_IN,sysdate,����ҽ��_IN,sysdate
      from ����ҽ����¼
      where ID=ҽ��ID_IN;



    --������Ӳ�������
    v_StudyInf :=lpad(' ',trunc((60 - length(�������_IN))/2), ' ') || �������_IN;
    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���)
    values(���Ӳ�������_ID.NEXTVAL,v_����ID,1,2, v_StudyInf ,1); --�������


    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���) values(���Ӳ�������_ID.NEXTVAL,v_����ID,2,2, '' ,1); --����
    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���) values(���Ӳ�������_ID.NEXTVAL,v_����ID,3,2, '' ,1); --����

----  ������ 2011/10/18 ������٣��������޷���ȡ��Ͻ��������
--   -- ��������=1,������=6,�����ı�='������',Ԥ�����=-8 �������=0
--    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,������,�����ı�,Ԥ�����ID,�������,�Ƿ���)
--    values(���Ӳ�������_ID.NEXTVAL,v_����ID,16,1,6,'������' ,-8,0,1);
--    commit;
--****--
    v_StudyInf := '  ������ ' || rpad(nvl(to_char( r_AdviceInf.����), ' '),15, ' ') || '  �Ա� ' || rpad(nvl(to_char( r_AdviceInf.�Ա�),' '),15,' ') || ' ���䣺 ' || r_AdviceInf.����;
    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���)
    values(���Ӳ�������_ID.NEXTVAL,v_����ID,4,2, v_StudyInf ,1); --�����Ϣ


    v_StudyInf := '����ţ� ' || rpad(nvl(to_char(r_AdviceInf.�����), ' '),15, ' ') || 'סԺ�ţ� ' || rpad(nvl(to_char(r_AdviceInf.סԺ��), ' '),15,' ') || ' ���ڣ� ' || to_char(sysdate, 'yyyy-mm-dd');
    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���)
    values(���Ӳ�������_ID.NEXTVAL,v_����ID,5,2, v_StudyInf ,1); --�����Ϣ


    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���) values(���Ӳ�������_ID.NEXTVAL,v_����ID,6,2, '' ,1); --����
    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���) values(���Ӳ�������_ID.NEXTVAL,v_����ID,7,2, '' ,1); --����


    if not(��Ͻ��_IN is null) then
      v_StudyInf := '��Ͻ����';
      insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���)
      values(���Ӳ�������_ID.NEXTVAL,v_����ID,8,2, v_StudyInf ,1); --�����Ϣ

      insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���)
      values(���Ӳ�������_ID.NEXTVAL,v_����ID,9,2, ��Ͻ��_IN ,1); --�����Ϣ
    end if;

    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���) values(���Ӳ�������_ID.NEXTVAL,v_����ID,10,2, '' ,1); --����


--- ������ 2011/10/18 �޸�  ������Ͻ�� ��ID ����
    begin
  Select id  into v_��ID from  ���Ӳ������� where �ļ�id=v_����ID and �����ı�='������';
    exception
      when others then
        v_��ID := null;
    end;


    if not(��Ͻ���_IN is null) then
      v_StudyInf := '��Ͻ��飺';
      insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���)
      values(���Ӳ�������_ID.NEXTVAL,v_����ID,11,2, v_StudyInf ,1); --�����Ϣ

      insert into ���Ӳ�������(ID, �ļ�ID,��ֹ��,��ID,�������,��������,��������,�����ı�,�Ƿ���)
      values(���Ӳ�������_ID.NEXTVAL,v_����ID,0,v_��ID,12,2,0, ��Ͻ���_IN ,1); --�����Ϣ
    end if;

    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���) values(���Ӳ�������_ID.NEXTVAL,v_����ID,13,2, '' ,1); --����

    v_StudyInf := '------------------------------------------------------------------------';
    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,�����ı�,�Ƿ���) values(���Ӳ�������_ID.NEXTVAL,v_����ID,14,2, v_StudyInf ,1); --ͼ�ηָ�


    --����ҽ������
    insert into ����ҽ������(ҽ��ID,����ID) values(v_��ҽ��ID,v_����ID);



    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ�ļ�����룬��ӱ���󲻸���ִ��״̬
    Select Count(*) Into v_Count From ����ҽ����¼ a, ������ҳ b
    Where  a.����ID=b.����ID And a.��ҳID = b.��ҳID And b.��Ժ���� Is Not Null And a.Id = v_��ҽ��ID;

    If v_Count =0 Then
        --ֻ���Ѿ��������룬����ִ�е�ҽ���Ÿ��£�����Ϊ ���״̬����˹���
        Update ����ҽ������ Set ִ��״̬=1, ִ�й���=6, ���ʱ��=sysdate
        Where ҽ��id in(select id from ����ҽ����¼ where id= v_��ҽ��ID or ���id=v_��ҽ��ID) ;  --And ִ��״̬ = 3 --���ڲ���Ҫִ�С��������롱�Ĺ��̣���˲���Ҫ�ж�ִ��״̬
    End If;

    --����Ӱ�����¼�ı�����Ϣ
    update Ӱ�����¼ set ������=����ҽ��_IN where ҽ��ID=ҽ��ID_IN;

    close c_AdviceInf;

    exception
      when others then
        zl_ErrorCenter(sqlCode, sqlErrm);
  end SendElectrocardioReport;



    -----------------------------------------------------------------------------
  -- ��    �ܣ�����ĵ�ͼ��
  -----------------------------------------------------------------------------
  function AddElectrocardioReportImage
  (
    ҽ��id_IN       ����ҽ������.ҽ��ID%TYPE
  )return number is
  PRAGMA AUTONOMOUS_TRANSACTION;
    v_��ҽ��ID      ����ҽ������.ҽ��ID%Type;
    v_����ID        ���Ӳ�����¼.ID%type;
    v_����ID        ���Ӳ�������.ID%Type;
    v_���          ���Ӳ�������.�������%Type;
  begin
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��ID From ����ҽ����¼ Where ID = ҽ��id_In;

    select ����ID into v_����ID from ����ҽ������ where ҽ��ID= v_��ҽ��ID;

    select ���Ӳ�������_ID.NEXTVAL into v_����ID from dual;
    select nvl(max(�������),0) + 1 into v_��� from ���Ӳ������� where �ļ�ID=v_����ID;

    insert into ���Ӳ�������(ID, �ļ�ID,�������,��������,��������,�Ƿ���) values(v_����ID,v_����ID,v_���,5, '2;0;0;0;0;12150;13020;1;1;1;0' ,1);

    commit;

    return v_����ID;

    exception
      when others then
        zl_ErrorCenter(sqlCode, sqlErrm);

  end AddElectrocardioReportImage;



  -----------------------------------------------------------------------------
  -- ��    �ܣ�������渽��
  -----------------------------------------------------------------------------
  procedure ClearReportAffix
  (
    ����id_In       ���Ӳ�������.����ID%TYPE,
    �������_IN     ���Ӳ�������.������%TYPE
  )is
  begin
      Delete From ���Ӳ������� Where ����id = ����id_In And ������ = �������_IN;
    Exception
      When Others Then
        Zl_ErrorCenter(Sqlcode, Sqlerrm);
  end ClearReportAffix;

  -----------------------------------------------------------------------------
  -- ��    �ܣ���ӱ��渽��
  -----------------------------------------------------------------------------
  Procedure AddReportAffix
  (
    ����id_In In ���Ӳ�������.����id%Type,
    �ļ���_In In ���Ӳ�������.�ļ���%Type,
    ��С_In   In ���Ӳ�������.��С%Type,
    �������_IN in  ���Ӳ�������.������%TYPE
  )is
  begin
    Insert Into ���Ӳ�������(����id, ���, �ļ���, ��С, ������, ����)
    Values(����id_In, 10000, �ļ���_In, ��С_In, �������_IN, Sysdate);

    Exception
      When Others Then
        Zl_ErrorCenter(Sqlcode, Sqlerrm);
  end;


end b_PacsInterface;
/
