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
--139591:��ҵ��,2019-04-18,ҩƷ���Ŀ��ͼ۸�����Ϊ�պ�0�����⴦��
Declare
  n_�������� ҩƷ���.��������%Type;
  n_����     ҩƷ���.ʵ������%Type;
  n_���     ҩƷ���.ʵ�ʽ��%Type;
  n_���     ҩƷ���.ʵ�ʲ��%Type;
  n_ʱ���ۼ� ҩƷ���.���ۼ�%Type;
  n_�ɱ���   ҩƷ���.ƽ���ɱ���%Type;
  n_Count    Number(18) := 0;
Begin
  --1.�ⷿ��������ֻ��һ�����Σ�������=null
  --��������¼
  --д������־
  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = User || '_ҩƷ�����������_20190312_1';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values (User || '_ҩƷ�����������_20190312_1', Null);
  End If;

  Update ҩƷ���
  Set ���� = 0
  Where ���� = 1 And ���� Is Null And
        (�ⷿid, ҩƷid) In (Select b.�ⷿid, b.ҩƷid
                         From ҩƷ��� B,
                              (Select a.ҩƷid, a.�ⷿid
                                From ҩƷ��� A
                                Where a.���� = 1 And Zl_Fun_Getbatchpro(a.�ⷿid, a.ҩƷid) = 0
                                Group By a.�ⷿid, a.ҩƷid
                                Having Count(Nvl(a.����, 0)) = 1) A
                         Where b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And b.���� Is Null);

  Commit;

  Update Zlupgradeconfig Set ���� = '�Ѵ�������Ϊnull���' Where ��Ŀ = User || '_ҩƷ�����������_20190312_1';
  Commit;

  --2.�ⷿ��������������2�����Σ����ܼ�������Ϊnull�ģ�Ҳ������=0��  
  --����ҩƷ���  
  --д������־
  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = User || '_ҩƷ�����������_20190312_2';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values (User || '_ҩƷ�����������_20190312_2', Null);
  End If;
  For r_������ε��� In (Select Distinct b.�ⷿid, b.ҩƷid, Nvl(c.�Ƿ���, 0) As �Ƿ�ʱ��, Nvl(d.�ϴ��ۼ�, e.�ּ�) As ʱ���ۼ�, d.�ɱ���
                   From ҩƷ��� B, �շ���ĿĿ¼ C, ҩƷ��� D, �շѼ�Ŀ E,
                        (Select a.ҩƷid, a.�ⷿid
                          From ҩƷ��� A, ҩƷ��� B
                          Where a.���� = 1 And a.ҩƷid = b.ҩƷid And Zl_Fun_Getbatchpro(a.�ⷿid, b.ҩƷid) = 0
                          Group By a.�ⷿid, a.ҩƷid
                          Having Count(Nvl(a.����, 0)) > 1) A
                   Where b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And c.Id = a.ҩƷid And d.ҩƷid = c.Id And
                         e.�շ�ϸĿid = c.Id And Sysdate Between e.ִ������ And Nvl(e.��ֹ����, Sysdate)
                   Union All
                   Select Distinct b.�ⷿid, b.ҩƷid, Nvl(c.�Ƿ���, 0) As �Ƿ�ʱ��, Nvl(d.�ϴ��ۼ�, e.�ּ�) As ʱ���ۼ�, d.�ɱ���
                   From ҩƷ��� B, �շ���ĿĿ¼ C, �������� D, �շѼ�Ŀ E,
                        (Select a.ҩƷid, a.�ⷿid
                          From ҩƷ��� A, �������� B
                          Where a.���� = 1 And a.ҩƷid = b.����id And Zl_Fun_Getbatchpro(a.�ⷿid, b.����id) = 0
                          Group By a.�ⷿid, a.ҩƷid
                          Having Count(Nvl(a.����, 0)) > 1) A
                   Where b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And c.Id = a.ҩƷid And d.����id = c.Id And
                         e.�շ�ϸĿid = c.Id And Sysdate Between e.ִ������ And Nvl(e.��ֹ����, Sysdate)
                   Order By �ⷿid, ҩƷid) Loop
  
    --����кϲ����ε�����������ۣ����������ۼۣ�ʱ�ۣ���ƽ���ɱ��ۣ��ϲ�������=0
    Select Sum(Nvl(��������, 0)), Sum(Nvl(ʵ������, 0)), Sum(Nvl(ʵ�ʽ��, 0)), Sum(Nvl(ʵ�ʲ��, 0))
    Into n_��������, n_����, n_���, n_���
    From ҩƷ���
    Where ���� = 1 And �ⷿid = r_������ε���.�ⷿid And ҩƷid = r_������ε���.ҩƷid And Nvl(����, 0) = 0;
  
    --����ʱ���ۼ�        
    If r_������ε���.�Ƿ�ʱ�� = 1 Then
      If n_���� <> 0 Then
        n_ʱ���ۼ� := n_��� / n_����;
      End If;
    
      If n_���� = 0 Or Nvl(n_ʱ���ۼ�, 0) <= 0 Then
        n_ʱ���ۼ� := r_������ε���.ʱ���ۼ�;
      End If;
    End If;
  
    --����ɱ���
    If n_���� <> 0 Then
      n_�ɱ��� := (n_��� - n_���) / n_����;
    End If;
  
    If n_���� = 0 Or Nvl(n_�ɱ���, 0) <= 0 Then
      n_�ɱ��� := r_������ε���.�ɱ���;
    End If;
  
    --��������=0�ļ�¼
    Update ҩƷ���
    Set �������� = n_��������, ʵ������ = n_����, ʵ�ʽ�� = n_���, ʵ�ʲ�� = n_���, ���ۼ� = Decode(r_������ε���.�Ƿ�ʱ��, 1, n_ʱ���ۼ�, Null),
        ƽ���ɱ��� = n_�ɱ���
    Where ���� = 1 And �ⷿid = r_������ε���.�ⷿid And ҩƷid = r_������ε���.ҩƷid And ���� = 0;
  End Loop;

  --ɾ������=null�ļ�¼
  Delete From ҩƷ��� A Where a.���� = 1 And a.���� Is Null And Zl_Fun_Getbatchpro(a.�ⷿid, a.ҩƷid) = 0;

  Commit;

  Update Zlupgradeconfig Set ���� = '�Ѵ�������Ϊ0��null���' Where ��Ŀ = User || '_ҩƷ�����������_20190312_2';
  Commit;

  --д������־
  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = User || '_ҩƷ�����������_20190312_3';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values (User || '_ҩƷ�����������_20190312_3', Null);
  End If;

  --3.ɾ���۸��������Ϊ�յļ�¼
  Delete From ҩƷ�۸��¼ Where ���� Is Null;
  Commit;

  --4.���ݿ���¼����=0�����ݼ���Ӧ�ļ۸��
  For r_�۸���� In (Select a.�ⷿid, a.ҩƷid, a.����, Nvl(c.�Ƿ���, 0) As ʱ��, Nvl(a.���ۼ�, 0) As ���ۼ�, a.ƽ���ɱ���
                 From ҩƷ��� A, �շ���ĿĿ¼ C
                 Where a.ҩƷid = c.Id And a.���� = 1 And a.���� = 0
                 Order By a.�ⷿid, a.ҩƷid) Loop
  
    --����ʱ���ۼ�
    If r_�۸����.ʱ�� = 1 Then
      Begin
        Select Count(ID)
        Into n_Count
        From ҩƷ�۸��¼
        Where �۸����� = 1 And ��¼״̬ = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
    
      --�������=0����Ч�ļ۸����2�������ϣ���ɾ��ֻ����1��
      If n_Count > 1 Then
        Delete From ҩƷ�۸��¼
        Where �۸����� = 1 And ��¼״̬ = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0 And
              ID < (Select Max(ID)
                    From ҩƷ�۸��¼
                    Where �۸����� = 1 And ��¼״̬ = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0);
      End If;
    
      --ֹͣ�۸���еļ۸񣬲��²�������=0�ļ۸�
      Update ҩƷ�۸��¼
      Set ��ֹ���� = Sysdate - 1 / 24 / 60 / 60, ��¼״̬ = 2
      Where �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0 And ��¼״̬ = 1 And �۸����� = 1;
    
      --����ʱ���ۼۼ۸�
      Insert Into ҩƷ�۸��¼
        (ID, ԭ��id, �۸�����, ҩƷid, �ⷿid, ����, ԭ��, �ּ�, ��ҩ��λid, ����, Ч��, ����, ���Ч��, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������, ��ֹ����, ��¼״̬,
         ��������, ����˵��, ������, ���ۻ��ܺ�, �շ�id)
        Select ҩƷ�۸��¼_Id.Nextval, Null, 1, ҩƷid, �ⷿid, ����, 0, r_�۸����.���ۼ�, �ϴι�Ӧ��id, �ϴ�����, Ч��, �ϴβ���, ���Ч��, Null, Null,
               Null, Null, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'), 1, 0, '���κϲ�', 'ZLHIS', Null, Null
        From ҩƷ���
        Where ���� = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0;
    End If;
  
    --����ɱ���
    Begin
      Select Count(ID)
      Into n_Count
      From ҩƷ�۸��¼
      Where �۸����� = 2 And ��¼״̬ = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
  
    --�������=0����Ч�ļ۸����2�������ϣ���ɾ��ֻ����1��
    If n_Count > 1 Then
      Delete From ҩƷ�۸��¼
      Where �۸����� = 2 And ��¼״̬ = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0 And
            ID < (Select Max(ID)
                  From ҩƷ�۸��¼
                  Where �۸����� = 2 And ��¼״̬ = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0);
    End If;
  
    --ֹͣ�۸���еļ۸񣬲��²�������=0�ļ۸�
    Update ҩƷ�۸��¼
    Set ��ֹ���� = Sysdate - 1 / 24 / 60 / 60, ��¼״̬ = 2
    Where �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0 And ��¼״̬ = 1 And �۸����� = 2;
  
    --����ʱ���ۼۼ۸�
    Insert Into ҩƷ�۸��¼
      (ID, ԭ��id, �۸�����, ҩƷid, �ⷿid, ����, ԭ��, �ּ�, ��ҩ��λid, ����, Ч��, ����, ���Ч��, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������, ��ֹ����, ��¼״̬, ��������,
       ����˵��, ������, ���ۻ��ܺ�, �շ�id)
      Select ҩƷ�۸��¼_Id.Nextval, Null, 2, ҩƷid, �ⷿid, ����, 0, r_�۸����.ƽ���ɱ���, �ϴι�Ӧ��id, �ϴ�����, Ч��, �ϴβ���, ���Ч��, Null, Null,
             Null, Null, Sysdate, To_Date('3000-01-01', 'yyyy-mm-dd'), 1, 0, '���κϲ�', 'ZLHIS', Null, Null
      From ҩƷ���
      Where ���� = 1 And �ⷿid = r_�۸����.�ⷿid And ҩƷid = r_�۸����.ҩƷid And ���� = 0;
  End Loop;
  Commit;

  Update Zlupgradeconfig Set ���� = '�Ѵ�������Ϊ0�۸�' Where ��Ŀ = User || '_ҩƷ�����������_20190312_3';
  Commit;

  --5.����۸���п����ж������Ч������=0��û�п���¼�ļ۸�
  --д������־
  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = User || '_ҩƷ�����������_20190312_4';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values (User || '_ҩƷ�����������_20190312_4', Null);
  End If;
  For r_�޿��۸� In (Select a.�ⷿid, a.ҩƷid, a.�۸�����
                  From ҩƷ�۸��¼ A
                  Where a.��¼״̬ = 1 And a.���� = 0 And Not Exists
                   (Select 1
                         From ҩƷ��� B
                         Where b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And b.���� = a.����)
                  Group By a.�ⷿid, a.ҩƷid, a.�۸�����
                  Having Count(a.����) > 1
                  Order By a.�۸�����, a.�ⷿid, a.ҩƷid) Loop
  
    --ɾ������ļ۸�ֻ����1��
    Delete From ҩƷ�۸��¼
    Where �۸����� = r_�޿��۸�.�۸����� And ��¼״̬ = 1 And �ⷿid = r_�޿��۸�.�ⷿid And ҩƷid = r_�޿��۸�.ҩƷid And ���� = 0 And
          ID < (Select Max(ID)
                From ҩƷ�۸��¼
                Where �۸����� = r_�޿��۸�.�۸����� And ��¼״̬ = 1 And �ⷿid = r_�޿��۸�.�ⷿid And ҩƷid = r_�޿��۸�.ҩƷid And ���� = 0);
  End Loop;
  Commit;

  Update Zlupgradeconfig Set ���� = '�Ѵ���������Ϊ0�۸�' Where ��Ŀ = User || '_ҩƷ�����������_20190312_4';
  Commit;
End;
/


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--139585:����,2019-04-18,��������Ϊ�յ����
Create Or Replace Procedure Zl_ҩƷ�շ���¼_Adjust
(
  ҩƷid_In   In Number, --ҩƷID,Ϊ0ʱ�������Ԥ����
  ���۷�ʽ_In In Number := 0 --0-����ۼۺͳɱ���Ԥ����,1-ֻ����ۼ�Ԥ����,2-ֻ���ɱ���Ԥ����
) As
  Classid          Number(18); --������
  v_Billno         ҩƷ�շ���¼.No%Type; --���۵���
  Adjustdate       Date; --����ʱ��
  n_����           Number(18);
  n_�ּ�           �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��           �շѼ�Ŀ.ԭ��%Type;
  n_���           Number(8);
  n_ԭ��id         �շѼ�Ŀ.ԭ��id%Type;
  n_���۽��       ҩƷ���.ʵ�ʽ��%Type;
  n_�շ�id         ҩƷ�շ���¼.Id%Type;
  n_��ͨ���С��   Number;
  n_Stockid        ҩƷ�շ���¼.�ⷿid%Type;
  n_������id     ҩƷ�շ���¼.������id%Type;
  n_���ϵ��       ҩƷ�շ���¼.���ϵ��%Type;
  n_�۸�id         �շѼ�Ŀ.Id%Type;
  n_�޿�����ģʽ Number(1) := 0;
  n_��������       Number(1) := 0;
  n_��Ϣ����       Number(1) := 0;
  --�����ۼۣ�ʱ���ۼ�Ԥ���ۼ�¼
  --�۸����ͣ�0-�����ۼ�,1-ʱ���ۼ�
  Cursor c_Priceadjust Is
    Select 0 As �۸�����, p.Id As �۸�id, p.ԭ��id, p.ִ������, p.ԭ��, p.�ּ�, i.ҩƷid, s.�ⷿid As �ⷿid, Nvl(s.����, 0) As ����, s.�ϴ����� As ����,
           s.Ч��, s.�ϴβ��� As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��,
           Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�, s.ƽ���ɱ���, s.Rowid As ����¼, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������,
           p.���ۻ��ܺ�
    From �շѼ�Ŀ P, ҩƷ��� I, �շ���ĿĿ¼ A, ҩƷ��� S
    Where i.ҩƷid = a.Id And p.�շ�ϸĿid = i.ҩƷid And i.ҩƷid = s.ҩƷid(+) And s.����(+) = 1 And Nvl(a.�Ƿ���, 0) = 0 And
          Sysdate Between p.ִ������ And Nvl(p.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.�䶯ԭ��, 0) = 0 And
          p.�շ�ϸĿid = Decode(ҩƷid_In, 0, p.�շ�ϸĿid, ҩƷid_In) And ���۷�ʽ_In In (0, 1)
    Union All
    Select �۸�����, p.Id As �۸�id, p.ԭ��id, p.ִ������, p.ԭ��, p.�ּ�, i.ҩƷid, p.�ⷿid As �ⷿid, Nvl(p.����, 0) As ����, p.���� As ����, p.Ч��,
           p.���� As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��,
           Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�, s.ƽ���ɱ���, s.Rowid As ����¼, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������,
           p.���ۻ��ܺ�
    From ҩƷ�۸��¼ P, ҩƷ��� I, �շ���ĿĿ¼ A, ҩƷ��� S
    Where i.ҩƷid = a.Id And p.ҩƷid = i.ҩƷid And p.�ⷿid = s.�ⷿid(+) And p.ҩƷid = s.ҩƷid(+) And
          Nvl(p.����, 0) = Nvl(s.����(+), 0) And s.����(+) = 1 And Sysdate Between p.ִ������ And
          Nvl(p.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.��¼״̬, 0) = 0 And
          p.ҩƷid = Decode(ҩƷid_In, 0, p.ҩƷid, ҩƷid_In) And �۸����� = 1 And ���۷�ʽ_In In (0, 1)
    Order By �۸�����, ҩƷid, ����, �ⷿid;

  r_Priceadjust c_Priceadjust%RowType;

  --�ɱ���Ԥ���ۼ�¼
  Cursor c_Costadjust Is
    Select �۸�����, p.Id As �۸�id, p.ԭ��id, p.ִ������, p.ԭ��, p.�ּ�, i.ҩƷid, p.�ⷿid As �ⷿid, Nvl(p.����, 0) As ����, p.���� As ����, p.Ч��,
           p.���� As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, Nvl(s.ʵ������, 0) As ʵ������, s.�ϴο��� As ����, Nvl(s.ʵ�ʽ��, 0) As ʵ�ʽ��,
           Nvl(s.ʵ�ʲ��, 0) As ʵ�ʲ��, Nvl(s.���ۼ�, 0) As ���ۼ�, s.ƽ���ɱ���, s.Rowid As ����¼, s.���Ч��, s.��׼�ĺ�, s.�ϴ��������� As ��������,
           p.���ۻ��ܺ�
    From ҩƷ�۸��¼ P, ҩƷ��� I, �շ���ĿĿ¼ A, ҩƷ��� S
    Where i.ҩƷid = a.Id And p.ҩƷid = i.ҩƷid And p.�ⷿid = s.�ⷿid(+) And p.ҩƷid = s.ҩƷid(+) And
          Nvl(p.����, 0) = Nvl(s.����(+), 0) And s.����(+) = 1 And Sysdate Between p.ִ������ And
          Nvl(p.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(p.��¼״̬, 0) = 0 And
          p.ҩƷid = Decode(ҩƷid_In, 0, p.ҩƷid, ҩƷid_In) And �۸����� = 2 And ���۷�ʽ_In In (0, 2)
    Order By ҩƷid, ����, �ⷿid;

  r_Costadjust c_Costadjust%RowType;

  --��ǰ��Ч�ļ۸������޿�����
  Cursor c_Nostockadjust
  (
    Drugid_In ҩƷ�۸��¼.ҩƷid%Type,
    Type_In   ҩƷ�۸��¼.�۸�����%Type
  ) Is
    Select a.�۸�����, a.Id As �۸�id, a.ԭ��, a.�ּ�, a.ҩƷid, a.�ⷿid, nvl(a.����,0) as ����, a.��ҩ��λid, a.����, a.Ч��, a.����
    From ҩƷ�۸��¼ A
    Where Sysdate Between a.ִ������ And Nvl(a.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.��¼״̬ = 1 And
          a.ҩƷid = Drugid_In And a.�۸����� = Type_In And a.�ⷿid Is Not Null
    Order By a.�ⷿid, a.ҩƷid, a.����;

  r_Nostockadjust c_Nostockadjust%RowType;
Begin
  --ȡ��ͨҵ�񾫶�λ��
  --���:1-ҩƷ 2-����
  --���ݣ�2-���ۼ� 4-���
  --��λ��ҩƷ:1-�ۼ� 5-��λ
  Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;

  --ȡ������ID
  Select ���id Into Classid From ҩƷ�������� Where ���� = 13;

  Adjustdate := Sysdate;

  --�ۼ۵��۴���
  If ���۷�ʽ_In = 0 Or ���۷�ʽ_In = 1 Then
  
    n_��� := 0;
  
    --ȡ����NOȡ
    Select Nextno(147) Into v_Billno From Dual;
  
    For r_Priceadjust In c_Priceadjust Loop
      If r_Priceadjust.�ⷿid Is Not Null Then
        --�пⷿid��������
      
        --ȡ��������
        n_�������� := Zl_Fun_Getbatchpro(r_Priceadjust.�ⷿid, r_Priceadjust.ҩƷid);
      
        --��������ӯ����¼��������1.Ҫ�п���¼��2.�������ԺͿ������һ��
        If r_Priceadjust.����¼ Is Not Null And ((n_�������� = 1 And r_Priceadjust.���� > 0) Or
           (n_�������� = 0 And r_Priceadjust.���� = 0)) Then
          n_��� := n_��� + 1;
        
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        
          n_ԭ�� := r_Priceadjust.ԭ��;
        
          --ʱ�۵��ۣ����ԭ�ۺ͵�ǰ��治һ�£����Ե�ǰ���Ϊ׼
          If r_Priceadjust.�۸����� = 1 And r_Priceadjust.ԭ�� <> r_Priceadjust.���ۼ� And r_Priceadjust.����¼ Is Not Null Then
            n_ԭ�� := r_Priceadjust.���ۼ�;
          End If;
        
          n_���۽�� := Round(r_Priceadjust.�ּ� * r_Priceadjust.ʵ������, n_��ͨ���С��) - Round(n_ԭ�� * r_Priceadjust.ʵ������, n_��ͨ���С��);
        
          n_�۸�id := r_Priceadjust.�۸�id;
          If r_Priceadjust.�۸����� = 1 Then
            Select ID
            Into n_�۸�id
            From �շѼ�Ŀ
            Where Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And �շ�ϸĿid = r_Priceadjust.ҩƷid;
          End If;
        
          --��������Ӱ���¼
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
             ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
          Values
            (n_�շ�id, 1, 13, v_Billno, n_���, Classid, r_Priceadjust.ҩƷid, r_Priceadjust.����, r_Priceadjust.����,
             r_Priceadjust.Ч��, r_Priceadjust.����, 1, r_Priceadjust.ʵ������, 0, n_ԭ��, 0, r_Priceadjust.�ּ�, r_Priceadjust.����,
             n_���۽��, n_���۽��, 'ҩƷ����', Zl_Username, Adjustdate, r_Priceadjust.�ⷿid, 1, n_�۸�id, Zl_Username, Adjustdate,
             r_Priceadjust.ʵ�ʽ��, r_Priceadjust.ʵ�ʲ��, r_Priceadjust.��Ӧ��id);
        
          Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
        
          --����ҩƷ��棬�޿�治ִ��
          If r_Priceadjust.����¼ Is Not Null Then
            Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
          End If;
        End If;
      
        --����ԭ�۸���Ϣ
        If r_Priceadjust.�۸����� = 1 Then
          Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Priceadjust.ԭ��id;
        End If;
      
        --ʱ�۵��۸��¼۸���е���Ϣ
        If r_Priceadjust.�۸����� = 1 Then
          --���µ�ǰ�۸���Ϣ
          If r_Priceadjust.����¼ Is Not Null Then
            Update ҩƷ�۸��¼
            Set ���� = r_Priceadjust.����, Ч�� = r_Priceadjust.Ч��, ���� = r_Priceadjust.����, ���Ч�� = r_Priceadjust.���Ч��,
                ��ҩ��λid = r_Priceadjust.��Ӧ��id, ԭ�� = n_ԭ��, �շ�id = n_�շ�id, ��¼״̬ = 1
            Where ID = r_Priceadjust.�۸�id;
          Else
            --�޿��ʱֻ���¼�¼״̬���շ�id
            Update ҩƷ�۸��¼ Set �շ�id = n_�շ�id, ��¼״̬ = 1 Where ID = r_Priceadjust.�۸�id;
          End If;
        End If;
      
        --�������Ŷ��ձ��ۼ�
        If r_Priceadjust.�۸����� = 1 Then
          --�����ʱ�ۣ�����¸�ҩƷ���ζ�Ӧ�ļ۸�
          Update ҩƷ���Ŷ���
          Set �ۼ� = r_Priceadjust.�ּ�
          Where ҩƷid = r_Priceadjust.ҩƷid And Nvl(����, 0) = r_Priceadjust.���� And �ۼ� <> r_Priceadjust.�ּ�;
        End If;
      
        --��Ϣ����
        --����ֻ����һ����Ϣ��ʱ�ۿɶ�ε���
        If (r_Priceadjust.�۸����� = 0 And n_��Ϣ���� = 0) Or r_Priceadjust.�۸����� = 1 Then
          n_��Ϣ���� := 1;
          b_Message.Zlhis_Drug_009(r_Priceadjust.�۸�id, r_Priceadjust.�۸�����);
        End If;
      Else
        --�޿�����ģʽ���۸���и�ҩƷ������Ч�ļ۸�Ҫ���޿�����ʱ�ļ۸����
      
        If r_Priceadjust.�۸����� = 1 Then
          --����ԭ�۸���Ϣ
          Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Priceadjust.ԭ��id;
        
          --�����ּ۸�״̬
          Update ҩƷ�۸��¼ Set ��¼״̬ = 1 Where ID = r_Priceadjust.�۸�id;
        End If;
      
        n_�޿�����ģʽ := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Priceadjust.ҩƷid, 1) Loop
          If r_Priceadjust.�ּ� <> r_Nostockadjust.�ּ� Then
            Zl_ҩƷ�۸��¼_Stop(1, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_ҩƷ�۸��¼_Insert(1, 1, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����, Null,
                             r_Priceadjust.�ּ�, Adjustdate - 1 / 24 / 60 / 60, 'ҩƷ����', Zl_Username, r_Priceadjust.���ۻ��ܺ�,
                             r_Nostockadjust.��ҩ��λid, r_Nostockadjust.����, r_Nostockadjust.Ч��, r_Nostockadjust.����);
            n_�޿�����ģʽ := 1;
          End If;
        End Loop;
        If n_�޿�����ģʽ = 1 Then
          Zl_ҩƷ�շ���¼_Adjust(r_Priceadjust.ҩƷid, 1);
        End If;
      End If;
    
      --���¹��۸�
      If r_Priceadjust.�ּ� <> r_Priceadjust.ԭ�� Then
        Update ҩƷ���
        Set �ϴ��ۼ� = r_Priceadjust.�ּ�
        Where ҩƷid = r_Priceadjust.ҩƷid And �ϴ��ۼ� <> r_Priceadjust.�ּ�;
      End If;
    
      If r_Priceadjust.�۸����� = 0 Then
        n_�۸�id := r_Priceadjust.�۸�id;
      Else
        Begin
          Select ID
          Into n_�۸�id
          From �շѼ�Ŀ
          Where Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Nvl(�䶯ԭ��, 0) = 0 And
                �շ�ϸĿid = r_Priceadjust.ҩƷid;
        Exception
          When Others Then
            n_�۸�id := 0;
        End;
      End If;
    
      If n_�۸�id > 0 Then
        Update �շѼ�Ŀ Set �䶯ԭ�� = 1 Where Nvl(�䶯ԭ��, 0) = 0 And ID = n_�۸�id;
      End If;
    
      --�������Ŷ��ձ��ۼ�
      If r_Priceadjust.�۸����� = 0 Then
        --����Ƕ��ۣ�����¸�ҩƷ��Ӧ���������ε��ۼ�
        Update ҩƷ���Ŷ���
        Set �ۼ� = r_Priceadjust.�ּ�
        Where ҩƷid = r_Priceadjust.ҩƷid And �ۼ� <> r_Priceadjust.�ּ�;
      End If;
    End Loop;
  End If;

  --�ɱ��۵��۴���
  If ���۷�ʽ_In = 0 Or ���۷�ʽ_In = 2 Then
  
    n_���    := 0;
    n_Stockid := 0;
  
    Select b.Id, b.ϵ��
    Into n_������id, n_���ϵ��
    From ҩƷ�������� A, ҩƷ������ B
    Where a.���id = b.Id And a.���� = 5 And Rownum < 2;
  
    v_Billno := Nextno(25, n_Stockid);
  
    For r_Costadjust In c_Costadjust Loop
      If r_Costadjust.�ⷿid Is Not Null Then
        --�пⷿid��������
      
        --ȡ��������
        n_�������� := Zl_Fun_Getbatchpro(r_Costadjust.�ⷿid, r_Costadjust.ҩƷid);
      
        --��������ӯ����¼��������1.Ҫ�п���¼��2.�������ԺͿ������һ��
        If r_Costadjust.����¼ Is Not Null And ((n_�������� = 1 And r_Costadjust.���� > 0) Or
           (n_�������� = 0 And r_Costadjust.���� = 0)) Then
          n_��� := n_��� + 1;
        
          --��������۵�����
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        
          --���ԭ�ۺ͵�ǰ��治һ�£����Ե�ǰ���Ϊ׼
          n_ԭ�� := r_Costadjust.ԭ��;
          If r_Costadjust.ԭ�� <> r_Costadjust.ƽ���ɱ��� And r_Costadjust.����¼ Is Not Null Then
            n_ԭ�� := r_Costadjust.ƽ���ɱ���;
          End If;
        
          n_���۽�� := Round(n_ԭ�� * r_Costadjust.ʵ������, n_��ͨ���С��) - Round(r_Costadjust.�ּ� * r_Costadjust.ʵ������, n_��ͨ���С��);
        
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, ���ۼ�, ���۽��, �ɱ���, �ɱ����,
             ���, ժҪ, ������, ��������, �����, �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����, ���Ч��)
          Values
            (n_�շ�id, 1, 5, v_Billno, n_���, r_Costadjust.�ⷿid, n_������id, r_Costadjust.��Ӧ��id, n_���ϵ��, r_Costadjust.ҩƷid,
             r_Costadjust.����, r_Costadjust.����, r_Costadjust.����, r_Costadjust.Ч��, r_Costadjust.ʵ������, 0, r_Costadjust.ʵ�ʽ��,
             0, r_Costadjust.ʵ�ʲ��, 0, n_���۽��, '�ɱ��۵���', Zl_Username, Adjustdate, Zl_Username, Adjustdate,
             r_Costadjust.��������, r_Costadjust.��׼�ĺ�, r_Costadjust.�ּ�, 1, n_ԭ��, r_Costadjust.���Ч��);
        
          Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
        
          Zl_ҩƷ���_Update(n_�շ�id, 2, 0);
        
          --���µ�ǰ�۸���Ϣ
          Update ҩƷ�۸��¼
          Set ���� = r_Costadjust.����, Ч�� = r_Costadjust.Ч��, ���� = r_Costadjust.����, ���Ч�� = r_Costadjust.���Ч��,
              ��ҩ��λid = r_Costadjust.��Ӧ��id, ԭ�� = n_ԭ��, �շ�id = n_�շ�id, ��¼״̬ = 1
          Where ID = r_Costadjust.�۸�id;
        Else
          --�޿��ʱֻ���¼�¼״̬
          Update ҩƷ�۸��¼ Set ��¼״̬ = 1 Where ID = r_Costadjust.�۸�id;
        End If;
      
        --����ԭ�۸���Ϣ
        Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Costadjust.ԭ��id;
      
        --�������Ŷ��ձ�ɱ���
        Update ҩƷ���Ŷ���
        Set �ɱ��� = r_Costadjust.�ּ�
        Where ҩƷid = r_Costadjust.ҩƷid And Nvl(����, 0) = r_Costadjust.���� And �ɱ��� <> r_Costadjust.�ּ�;
      Else
        --�޿�����ģʽ���۸���и�ҩƷ������Ч�ļ۸�Ҫ���޿�����ʱ�ļ۸����
      
        --����ԭ�۸���Ϣ
        Update ҩƷ�۸��¼ Set ��¼״̬ = 2 Where ID = r_Costadjust.ԭ��id;
      
        --�����ּ۸�״̬
        Update ҩƷ�۸��¼ Set ��¼״̬ = 1 Where ID = r_Costadjust.�۸�id;
      
        n_�޿�����ģʽ := 0;
        For r_Nostockadjust In c_Nostockadjust(r_Costadjust.ҩƷid, 2) Loop
          If r_Costadjust.�ּ� <> r_Nostockadjust.�ּ� Then
            Zl_ҩƷ�۸��¼_Stop(2, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����,
                           Adjustdate - 2 / 24 / 60 / 60);
            Zl_ҩƷ�۸��¼_Insert(1, 2, r_Nostockadjust.�ⷿid, r_Nostockadjust.ҩƷid, r_Nostockadjust.����, Null,
                             r_Costadjust.�ּ�, Adjustdate - 1 / 24 / 60 / 60, '�ɱ��۵���', Zl_Username, r_Costadjust.���ۻ��ܺ�,
                             r_Nostockadjust.��ҩ��λid, r_Nostockadjust.����, r_Nostockadjust.Ч��, r_Nostockadjust.����);
            n_�޿�����ģʽ := 1;
          End If;
        End Loop;
        If n_�޿�����ģʽ = 1 Then
          Zl_ҩƷ�շ���¼_Adjust(r_Costadjust.ҩƷid, 2);
        End If;
      End If;
    
      --���¹��۸�
      If r_Costadjust.ԭ�� <> r_Costadjust.�ּ� Then
        Update ҩƷ���
        Set �ɱ��� = r_Costadjust.�ּ�
        Where ҩƷid = r_Costadjust.ҩƷid And �ɱ��� <> r_Costadjust.�ּ�;
      End If;
    
      --��Ϣ����
      b_Message.Zlhis_Drug_007(r_Costadjust.�۸�id);
    End Loop;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_Adjust;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_ҩƷ���_Update
(
  Id_In       In ҩƷ�շ���¼.Id%Type,
  ҵ������_In In Number := 0,
  �������_In In Number := 0,
  ��������_In In Number := 0
) Is
  --���ܣ�
  --      ����ҵ�����ʹ����������ҵ��ҩƷ����ҵ������ķ���ҵ�������ڲ���ͨҵ�񲻴���
  --id_in  ��Ҫ�����շ���¼��
  --ҵ������_in  ҵ�����ͣ�0-������1-ɾ����2-��ˡ�3-����
  --�������_in  0-��⣬1-����
  --�������ͣ�����ҵ�����ȷ��
  ----�⹺����У���ʾ������ˣ�  0-���ǲ�����ˣ�1-������ˣ�Ŀǰֻ���⹺����в������
  ----���죬�ƿ��г������̣�0-�����������̣�1-���룬��˳�������

  n_�������� ҩƷ���.ʵ������%Type;
  n_ʵ������ ҩƷ���.ʵ������%Type;
  n_���۽�� ҩƷ���.ʵ�ʽ��%Type;
  n_���     ҩƷ���.ʵ�ʲ��%Type;
  n_ʱ��     Number(1);
  n_�ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  n_���ۼ�   ҩƷ���.���ۼ�%Type;

  n_�������     ҩƷ���.ʵ������%Type;
  n_���ƽ����   ҩƷ���.ƽ���ɱ���%Type;
  n_����ۼ�     ҩƷ���.���ۼ�%Type;
  n_������       ҩƷ�շ���¼.ʵ������%Type;
  n_�ܳɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_���ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  n_�ⷿ����     ҩƷ���.ҩ�����%Type;
  n_�������     Number(1);
  n_���¿��     Number(1) := 0;
  v_�ּ�         ҩƷ�շ���¼.���ۼ�%Type;
  v_ִ���¼۸�   Number(1) := 0;
  n_�п��       Number(1) := 0;
  v_�������     ҩƷ�շ���¼.�������%Type;
  n_�۸����     Number(1) := 0;
  n_����ʱ���ۼ� Number(1) := 0;
  n_�����ɱ���   Number(1) := 0;
  n_��������     Number(1) := 0; --0-�������Բ����ϣ�1-�������Է���
  n_�������     Number(1) := 0; --0-���¿�棬1-�������
  n_����۸�     Number(1) := 0; --0-������۸�,1-����۸�
  --ҵ����ϸ���ݣ��ѿ�����ݸ�����Ҫ�����ݶ��г���
  Cursor c_Detail Is
    Select a.Id, a.��¼״̬, a.����, a.No, a.���, a.�ⷿid, a.��ҩ��λid, a.������id, a.�Է�����id, a.���ϵ��, Nvl(a.��ҩ��ʽ, 0) As ��ҩ��ʽ, a.ҩƷid,
           Nvl(a.����, 0) ����, a.����, a.ԭ����, a.����, a.��������, a.Ч��, a.����, Nvl(a.��д����, 0) As ��д����, a.ʵ������, a.�ɱ���, a.�ɱ����, a.����,
           a.���ۼ�, Nvl(a.���۽��, 0) As ���۽��, Nvl(a.���, 0) As ���, a.��ҩ��, a.��ҩ����, a.�����, a.�������, a.�������, a.���Ч��, a.��׼�ĺ�,
           a.��Ʒ����, a.�ڲ�����, Nvl(b.�Ƿ���, 0) As �Ƿ���, a.����, a.Ƶ��, a.ժҪ, Nvl(a.����id, 0) As ����id,
           Decode(a.����, Null, 1, 0) ������
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.Id = Id_In;

  r_Detail c_Detail%RowType;
Begin
  For r_Detail In c_Detail Loop
  	If r_Detail.������ = 1 Then
      Update ҩƷ�շ���¼ Set ���� = 0 Where NO = r_Detail.No And ���� = r_Detail.���� And ��� = r_Detail.���;
    End If;

    If Zl_Fun_Getbatchpro(r_Detail.�ⷿid, r_Detail.ҩƷid) = 1 Then
      If r_Detail.���� > 0 Then
        n_�������� := 1;
      Else
        n_�������� := 0;
      End If;
    Else
      If r_Detail.���� = 0 Then
        n_�������� := 1;
      Else
        n_�������� := 0;
      End If;
    End If;
  
    n_ʵ������ := r_Detail.���ϵ�� * r_Detail.ʵ������ * Nvl(r_Detail.����, 1);
    If n_ʵ������ Is Null Then
      n_ʵ������ := 0;
    End If;
    n_�������� := 0;
    n_���ۼ�   := r_Detail.���ۼ�;
    If r_Detail.���� = 12 Then
      n_�ɱ��� := r_Detail.����;
    Else
      n_�ɱ��� := r_Detail.�ɱ���;
    End If;
    n_���۽�� := r_Detail.���ϵ�� * r_Detail.���۽��;
    n_���     := r_Detail.���ϵ�� * r_Detail.���;
  
    --��ȡ���͵��ݵ������ͳɱ���
    Begin
      Select Nvl(ʵ������, 0)
      Into n_�������
      From ҩƷ���
      Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.����;
    Exception
      When Others Then
        n_������� := 0;
    End;
  
    n_���ƽ���� := Zl_Fun_Getoutcost(r_Detail.ҩƷid, r_Detail.����, r_Detail.�ⷿid);
    n_����ۼ�   := Zl_Fun_Getoutprice(r_Detail.ҩƷid, r_Detail.����, r_Detail.�ⷿid);
  
    --ʱ��ҩƷ����Ҫ���¿������ۼ��ֶ�
    If r_Detail.�Ƿ��� = 1 Then
      n_ʱ�� := 1;
    Else
      n_ʱ�� := 0;
    End If;
  
    --����ҵ������
    --����ҵ��--5-��۵�����13-���۱䶯
    --����5��13����ҵ������_in��2-��ˡ��������_in  0-���
    If r_Detail.���� = 5 Or r_Detail.���� = 13 Then
      --�������͵ĵ����շ���¼�ɱ����ֶβ��Ǳ���������ɱ��۶��Ǵ洢����������
      If r_Detail.���� = 5 Then
        If r_Detail.��д���� <> 0 Then
          n_���ۼ� := Nvl(r_Detail.���ۼ�, 0) / r_Detail.��д����;
        Else
          n_���ۼ� := 0;
        End If;
        --���
        If r_Detail.��¼״̬ = 1 Then
          --��۵�����ҩ��ʽ=0���������ۡ��˻�����ҩ�����ĵ���������ҩ��ʽ=1
          n_�ɱ��� := r_Detail.����;
        Else
          --���� ��ԭԭʼ�ɱ���
          Begin
            --�ɱ���=(���-���)/����
            n_�ɱ��� := (Nvl(r_Detail.���ۼ�, 0) - Nvl(r_Detail.�ɱ���, 0)) / r_Detail.��д����;
          Exception
            When Others Then
              Select �ɱ��� Into n_�ɱ��� From ҩƷ��� Where ҩƷid = r_Detail.ҩƷid;
          End;
        End If;
      Else
        n_�ɱ��� := Nvl(r_Detail.����, 0) - Nvl(r_Detail.Ƶ��, 0);
      End If;
    
      If r_Detail.���� = 5 Then
        --����=5 �ĳɱ���������¼ ƽ���ɱ��۲���Ҫ���㣬��Ϊ���������¼۸��
        If r_Detail.ժҪ = '�⹺�˿�������Զ�����' Or r_Detail.ժҪ = '������˼۸�䶯����' Then
          --��һ���϶����⹺�˿⣬�⹺�˿�ֻ���³ɱ���,�ҿ϶��п��
          Update ҩƷ���
          Set ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.����;
        Else
          Update ҩƷ���
          Set ƽ���ɱ��� = n_�ɱ���, �ϴβɹ��� = n_�ɱ���, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.����;
          If Sql%NotFound Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, ʵ�ʲ��, �ϴ�����, Ч��, �ϴβ���, ԭ����, �ϴι�Ӧ��id, �ϴ���������, ��׼�ĺ�, ʵ�ʽ��, �ϴβɹ���, ƽ���ɱ���)
            Values
              (r_Detail.�ⷿid, r_Detail.ҩƷid, r_Detail.����, 1, n_���, r_Detail.����, r_Detail.Ч��, r_Detail.����, r_Detail.ԭ����,
               r_Detail.��ҩ��λid, r_Detail.��������, r_Detail.��׼�ĺ�, n_���۽��, n_�ɱ���, n_�ɱ���);
          
            Insert Into ҩƷ�����Ϣ
              (ҩƷid, �ⷿid, ����, �������)
              Select r_Detail.ҩƷid, r_Detail.�ⷿid, r_Detail.����, r_Detail.�������
              From Dual
              Where Not Exists (Select 1
                     From ҩƷ�����Ϣ
                     Where ҩƷid = r_Detail.ҩƷid And �ⷿid = r_Detail.�ⷿid And ���� = r_Detail.����);
          End If;
        End If;
      Elsif r_Detail.���� = 13 Then
        --����=13 ���ۼ�������¼ ͬ�����µĽ��Ͳ�ۣ����Բ���Ҫ����ƽ���ɱ���
        If r_Detail.����id = 0 Then
          Update ҩƷ���
          Set ���ۼ� = Decode(n_ʱ��, 1, n_���ۼ�, Null), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.����;
        Else
          Update ҩƷ���
          Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.����;
        End If;
      
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�)
          Values
            (r_Detail.�ⷿid, r_Detail.ҩƷid, r_Detail.����, 1, 0, 0, n_���۽��, n_���۽��, Decode(n_ʱ��, 1, n_���ۼ�, Null));
        
          Insert Into ҩƷ�����Ϣ
            (ҩƷid, �ⷿid, ����, �������)
            Select r_Detail.ҩƷid, r_Detail.�ⷿid, r_Detail.����, r_Detail.�������
            From Dual
            Where Not Exists (Select 1
                   From ҩƷ�����Ϣ
                   Where ҩƷid = r_Detail.ҩƷid And �ⷿid = r_Detail.�ⷿid And ���� = r_Detail.����);
        End If;
      End If;
    Else
      --һ��ҵ������
      --����ҵ��--1-�⹺��⣻2-������⣻3-Эҩ��⣻4-������⣻6-�ⷿ�Ƴ���
      --7-�������ã�8-�շѴ�����ҩ��9-���ʵ�������ҩ��10-���ʱ�����ҩ��11-�������⣻
      --12-�̵㣻14-ҩƷ�̵��¼��
      --21-�����������⣻24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ�������
      If ҵ������_In = 0 Or ҵ������_In = 1 Then
        --������ɾ��
        If r_Detail.���� = 8 Or r_Detail.���� = 9 Or r_Detail.���� = 10 Or r_Detail.���� = 21 Or r_Detail.���� = 24 Or
           r_Detail.���� = 25 Or r_Detail.���� = 26 Or r_Detail.���� = 7 Or r_Detail.���� = 11 Or
           ((r_Detail.���� = 2 Or r_Detail.���� = 3 Or r_Detail.���� = 12) And r_Detail.���ϵ�� = -1) Or
           (r_Detail.���� = 1 And r_Detail.��ҩ��ʽ = 1) Or (r_Detail.���� = 6 And r_Detail.���ϵ�� = -1 And r_Detail.��¼״̬ = 1) Or
           (r_Detail.���� = 6 And Mod(r_Detail.��¼״̬, 3) = 2 And r_Detail.���ϵ�� = 1) Then
          --��Ҫ������/ɾ������ʱ����/���ӿ��������ĵ�������
          ----1.��ҩ/���ϵ���(�շѴ��������˵������˱�)
          ----2.��ͨ���⣨���á��������⡢�ƿ��г����Ǳ�(r_Detail.���� = 6 And r_Detail.���ϵ�� = -1 And r_Detail.��¼״̬ = 1)���̵㵥���̿��Ǳʣ�
          ----3.�˿ⵥ��r_Detail.���� = 1 And r_Detail.��ҩ��ʽ = 1��
          ----4.�ƿ������������ԭ����Ǳʵĳ�����¼��r_Detail.���� = 6 And Mod(r_Detail.��¼״̬, 3) = 2 And r_Detail.���ϵ�� = 1��
        
          --������ɾ������ʱ����Ϊû���������ֻ����������������Ͳ��
        
          If ҵ������_In = 0 Then
            --����ʱ���������������
            n_�������� := n_ʵ������;
          Else
            --ɾ��ʱ���෴�������������
            n_�������� := -1 * n_ʵ������;
          End If;
        
          n_ʵ������ := 0;
          n_���۽�� := 0;
          n_���     := 0;
        
          --������
          Update ҩƷ���
          Set �������� = �������� + n_��������
          Where ҩƷid = r_Detail.ҩƷid And �ⷿid = r_Detail.�ⷿid And Nvl(����, 0) = r_Detail.���� And ���� = 1;
        
          n_���¿�� := 1;
        End If;
      Elsif ҵ������_In = 2 Then
        --���
        --10.35��ʼ�����������еĳ����൥�������ʱ�����ٴ����������
        If r_Detail.���� = 8 Or r_Detail.���� = 9 Or r_Detail.���� = 10 Or r_Detail.���� = 21 Or r_Detail.���� = 24 Or
           r_Detail.���� = 25 Or r_Detail.���� = 26 Or r_Detail.���� = 7 Or r_Detail.���� = 11 Or
           ((r_Detail.���� = 2 Or r_Detail.���� = 3 Or r_Detail.���� = 12) And r_Detail.���ϵ�� = -1) Or
           (r_Detail.���� = 1 And r_Detail.��ҩ��ʽ = 1) Or (r_Detail.���� = 6 And r_Detail.���ϵ�� = -1 And r_Detail.��¼״̬ = 1) Or
           (r_Detail.���� = 6 And Mod(r_Detail.��¼״̬, 3) = 2 And r_Detail.���ϵ�� = 1) Then
          n_�������� := 0;
        Else
          n_�������� := n_ʵ������;
        End If;
      
        --������
        If �������_In = 0 Then
          --������
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_��������, ʵ������ = Nvl(ʵ������, 0) + n_ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��,
              ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���, �ϴι�Ӧ��id = r_Detail.��ҩ��λid,
              �ϴβɹ��� = Decode(r_Detail.����, 1, Decode(r_Detail.��ҩ��ʽ, 1, �ϴβɹ���, n_�ɱ���), n_�ɱ���),
              �ϴ����� = Nvl(r_Detail.����, �ϴ�����), �ϴ��������� = Nvl(r_Detail.��������, �ϴ���������), �ϴβ��� = Nvl(r_Detail.����, �ϴβ���),
              ԭ���� = Nvl(r_Detail.ԭ����, ԭ����), ���Ч�� = Nvl(r_Detail.���Ч��, ���Ч��), Ч�� = Nvl(r_Detail.Ч��, Ч��),
              ��׼�ĺ� = Nvl(r_Detail.��׼�ĺ�, ��׼�ĺ�), �ϴο��� = Decode(r_Detail.����, 12, �ϴο���, r_Detail.����),
              ��Ʒ���� = Nvl(r_Detail.��Ʒ����, ��Ʒ����), �ڲ����� = Nvl(r_Detail.�ڲ�����, �ڲ�����)
          Where �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.���� And ���� = 1;
        Else
          --������ˣ�ֻ��Ҫ�������ͽ��
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_��������, ʵ������ = Nvl(ʵ������, 0) + n_ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��,
              ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���, ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_�ɱ���, ƽ���ɱ���),
              �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_�ɱ���, �ϴβɹ���)
          Where �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.���� And ���� = 1;
        End If;
      
        n_���¿�� := 1;
      Elsif ҵ������_In = 3 Then
        --����
        If r_Detail.���� = 8 Or r_Detail.���� = 9 Or r_Detail.���� = 10 Or r_Detail.���� = 21 Or r_Detail.���� = 24 Or
           r_Detail.���� = 25 Or r_Detail.���� = 26 Then
          --��ҩ/���ϵ���ҩ/����ʱͬʱ�ֲ�����δ�����ݣ����ԾͲ������������
          n_�������� := 0;
        Elsif r_Detail.���� = 6 And Mod(r_Detail.��¼״̬, 3) = 2 And r_Detail.���ϵ�� = 1 Then
          --ҩ�ⵥ�ĳ������ݣ�Ҫ�ж��Ƿ���Ҫ����
          n_������� := ��������_In;
          If n_������� = 0 Then
            --����Ҫ������ڳ���ʱ�����������
            n_�������� := n_ʵ������;
          Else
            --��Ҫ����ģ��Ѿ�������ʱ�����˿�������
            n_�������� := 0;
          End If;
        Else
          n_�������� := n_ʵ������;
        End If;
      
        --������
        If �������_In = 0 Then
          --���ⵥ�ݳ�����Ҫ�����ⷿ���ݶ�����
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_��������, ʵ������ = Nvl(ʵ������, 0) + n_ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��,
              ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���, �ϴι�Ӧ��id = r_Detail.��ҩ��λid,
              �ϴβɹ��� = Decode(r_Detail.����, 1, Decode(r_Detail.��ҩ��ʽ, 1, �ϴβɹ���, n_�ɱ���), �ϴβɹ���),
              �ϴ����� = Nvl(r_Detail.����, �ϴ�����), �ϴ��������� = Nvl(r_Detail.��������, �ϴ���������), �ϴβ��� = Nvl(r_Detail.����, �ϴβ���),
              ԭ���� = Nvl(r_Detail.ԭ����, ԭ����), ���Ч�� = Nvl(r_Detail.���Ч��, ���Ч��), Ч�� = Nvl(r_Detail.Ч��, Ч��),
              ��׼�ĺ� = Nvl(r_Detail.��׼�ĺ�, ��׼�ĺ�), �ϴο��� = Decode(r_Detail.����, 12, �ϴο���, r_Detail.����),
              ��Ʒ���� = Nvl(r_Detail.��Ʒ����, ��Ʒ����), �ڲ����� = Nvl(r_Detail.�ڲ�����, �ڲ�����)
          Where �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.���� And ���� = 1;
        Else
          --��ⵥ�ݳ���ֻ��Ҫ�������ͽ��
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_��������, ʵ������ = Nvl(ʵ������, 0) + n_ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��,
              ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.���� And ���� = 1;
        End If;
      
        n_���¿�� := 1;
      End If;
    
      --����/ɾ��/���/����ҵ��ʱ������δ�ҵ���������Ҫ��������������Ϣ
      If Sql%RowCount = 0 And n_���¿�� = 1 Then
        --���ҵ��ȡ����۸�֮ǰ�Ѿ�ȡ���ˣ�������ҵ�����ҵ��ȡ���¼۸�
        If ҵ������_In = 3 Or �������_In = 1 Then
          --ȡ���³ɱ���
          v_�ּ� := Zl_Fun_Getoutcost(r_Detail.ҩƷid, r_Detail.����, r_Detail.�ⷿid, n_�ɱ���);
          If v_�ּ� Is Not Null Then
            n_�ɱ��� := v_�ּ�;
          End If;
        
          --ʱ���ۼ�ȡ���¼۸�
          If r_Detail.�Ƿ��� = 1 Then
            v_�ּ� := Zl_Fun_Getoutprice(r_Detail.ҩƷid, r_Detail.����, r_Detail.�ⷿid, n_���ۼ�);
            If v_�ּ� Is Not Null Then
              n_���ۼ� := v_�ּ�;
            End If;
          End If;
        End If;
      
        --�������
        n_������� := 1;
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, Ч��, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ԭ����, ���Ч��, ��׼�ĺ�, ���ۼ�,
           �ϴο���, ��Ʒ����, �ڲ�����, ƽ���ɱ���)
        Values
          (r_Detail.�ⷿid, r_Detail.ҩƷid, r_Detail.����, r_Detail.Ч��, 1, n_��������, n_ʵ������, n_���۽��, n_���, r_Detail.��ҩ��λid,
           n_�ɱ���, r_Detail.����, r_Detail.��������, r_Detail.����, r_Detail.ԭ����, r_Detail.���Ч��, r_Detail.��׼�ĺ�,
           Decode(n_ʱ��, 1, n_���ۼ�, Null), r_Detail.����, r_Detail.��Ʒ����, r_Detail.�ڲ�����, n_�ɱ���);
      
        Insert Into ҩƷ�����Ϣ
          (ҩƷid, �ⷿid, ����, �������)
          Select r_Detail.ҩƷid, r_Detail.�ⷿid, r_Detail.����, r_Detail.�������
          From Dual
          Where Not Exists (Select 1
                 From ҩƷ�����Ϣ
                 Where ҩƷid = r_Detail.ҩƷid And �ⷿid = r_Detail.�ⷿid And ���� = r_Detail.����);
      End If;
    
      --����ƽ���ɱ��ۣ���������Ҫ����ƽ���ɱ��ۺ����ۼۣ�ע��ֻ���ڲ�����ҩƷ������ҩƷ�������㣨ȷ����֮ǰ��������һ�£�
      --ֻ�и��¿��״̬��Ҫ���¼���۸��������״̬��������
      If �������_In = 0 And ҵ������_In = 2 And r_Detail.���� = 0 And n_������� <> 1 Then
        --���ܽ��/��������ʽ����ƽ���ɱ��۶����ã����-��ۣ�/������Ϊ�����ݵ�׼ȷ��
        n_����۸� := 1;
        n_������   := (n_������� + n_ʵ������);
        If n_������ <> 0 Then
          n_�ܳɱ��� := (n_������� * n_���ƽ���� + n_ʵ������ * n_�ɱ���) / n_������;
        
          If n_�ܳɱ��� < 0 Then
            n_�ܳɱ��� := n_�ɱ���;
          End If;
        
          Update ҩƷ���
          Set ƽ���ɱ��� = n_�ܳɱ���
          Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.����;
        
          --����ʱ�����ۼ�
          If n_ʱ�� = 1 Then
            n_���ۼ� := (n_������� * n_����ۼ� + n_ʵ������ * n_���ۼ�) / n_������;
            Update ҩƷ���
            Set ���ۼ� = n_���ۼ�
            Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.���� And
                  Nvl(ʵ������, 0) <> 0;
            If Sql%NotFound Then
              n_���ۼ� := n_���ۼ�;
              Update ҩƷ���
              Set ���ۼ� = n_���ۼ�
              Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.����;
            End If;
          End If;
        End If;
      End If;
    
      --�۸���
      --����������ȷʱ�Ž��м۸���
      If n_�������� = 1 Then
        --������棬���߸��¿�沢�������˼۸������²Ŵ���۸�
        If n_������� = 1 Or (n_������� = 0 And n_����۸� = 1) Then
          --ʱ�ۼ۸�
          If r_Detail.�Ƿ��� = 1 Then
            Zl_ҩƷ�۸��¼_Addnew(n_�������, 0, 1, r_Detail.�ⷿid, r_Detail.ҩƷid, r_Detail.����, 0, Nvl(n_���ۼ�, n_���ۼ�), Sysdate,
                             '����������μ۸�', Zl_Username, Null, r_Detail.��ҩ��λid, r_Detail.����, r_Detail.Ч��, r_Detail.����, Null,
                             Null, Null, Null, Null, 1);
          End If;
        
          --�ɱ��ۼ۸�
          Zl_ҩƷ�۸��¼_Addnew(n_�������, 0, 2, r_Detail.�ⷿid, r_Detail.ҩƷid, r_Detail.����, 0, Nvl(n_�ܳɱ���, n_�ɱ���), Sysdate,
                           '����������μ۸�', Zl_Username, Null, r_Detail.��ҩ��λid, r_Detail.����, r_Detail.Ч��, r_Detail.����, Null,
                           Null, Null, Null, Null, 1);
        End If;
      End If;
    End If;
  
    --ɾ������Ŀ�����ݣ��⹺���������Ϊ��ȷ����治������������ݱ��뱣֤��ɾ�����
    If Not (r_Detail.���� = 1 And ��������_In = 1) Then
      Delete From ҩƷ���
      Where ���� = 1 And �ⷿid = r_Detail.�ⷿid And ҩƷid = r_Detail.ҩƷid And Nvl(����, 0) = r_Detail.���� And Nvl(��������, 0) = 0 And
            Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
    End If;
  
    Zl_ҩƷ���_���������쳣����(r_Detail.�ⷿid, r_Detail.ҩƷid, r_Detail.����);
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ���_Update;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_ҩƷ�������_Insert
(
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  ������id_In In ҩƷ�շ���¼.������id%Type,
  ҩƷid_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type,
  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type,
  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type,
  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
  ��׼�ĺ�_In   In ҩƷ�շ���¼.��׼�ĺ�%Type := Null,
  ���_In       In ҩƷ�շ���¼.���%Type := Null,
  ����_In     In ҩƷ�շ���¼.���۽��%Type := Null,
  ԭ����_In       In ҩƷ�շ���¼.ԭ����%Type := Null,
  �޸���_In     In ҩƷ�շ���¼.�޸���%Type,
  �޸�����_In   In ҩƷ�շ���¼.�޸�����%Type
) Is
  v_Lngid    ҩƷ�շ���¼.Id%Type; --�շ�ID 
  v_���ϵ�� ҩƷ�շ���¼.���ϵ��%Type;
  v_����     ҩƷ�շ���¼.����%Type := 0; --���� 
  v_ҩ����� Integer; --�Ƿ�ҩ�����    1:����;0�������� 
  v_ҩ������ Integer; --�Ƿ�ҩ�����    1:����;0�������� 
  v_ʱ�۷��� Number(1);

Begin

  If Not ��׼�ĺ�_In Is Null And Not ����_In Is Null Then
    Update ҩƷ�����̶��� Set ��׼�ĺ� = ��׼�ĺ�_In Where ҩƷid = ҩƷid_In And �������� = ����_In;
  End If;
  If Sql%RowCount = 0 And Not ����_In Is Null And Not ��׼�ĺ�_In Is Null Then
    Insert Into ҩƷ�����̶��� (ҩƷid, ��������, ��׼�ĺ�) Values (ҩƷid_In, ����_In, ��׼�ĺ�_In);
  End If;

  v_���ϵ�� := 1;
  Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;
  Select Nvl(ҩ�����, 0), Nvl(ҩ������, 0) Into v_ҩ�����, v_ҩ������ From ҩƷ��� Where ҩƷid = ҩƷid_In;

  If v_ҩ������ = 0 Then
    If v_ҩ����� = 1 Then
      Begin
        Select Distinct 0
        Into v_ҩ�����
        From ��������˵��
        Where ((�������� Like '%ҩ��') Or (�������� Like '�Ƽ���')) And ����id = �ⷿid_In;
      Exception
        When Others Then
          v_ҩ����� := 1;
      End;
    
      If v_ҩ����� = 1 Then
        v_���� := Zl_Fun_Getbatchnum(ҩƷid_In, ����_In, ����_In, �ɱ���_In, ���ۼ�_In, v_Lngid, Null);
      End If;
    End If;
  Else
    v_���� := Zl_Fun_Getbatchnum(ҩƷid_In, ����_In, ����_In, �ɱ���_In, ���ۼ�_In, v_Lngid, Null);
  End If;

  Select Nvl(�Ƿ���, 0) Into v_ʱ�۷��� From �շ���ĿĿ¼ Where ID = ҩƷid_In;

  If v_ʱ�۷��� = 1 And v_���� > 0 Then
    v_ʱ�۷��� := 1;
  Else
    v_ʱ�۷��� := 0;
  End If;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����,����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ, ������,
     ��������, ��������, ��׼�ĺ�, ���, �÷�, �޸���, �޸�����)
  Values
    (v_Lngid, 1, 4, No_In, ���_In, �ⷿid_In, ������id_In, v_���ϵ��, ҩƷid_In, v_����, ����_In, ԭ����_In,����_In, Ч��_In, ʵ������_In, ʵ������_In,
     �ɱ���_In, �ɱ����_In, ���ۼ�_In, ���۽��_In, ���_In, ժҪ_In, ������_In, ��������_In, ��������_In, ��׼�ĺ�_In, ���_In,
     Decode(v_ʱ�۷���, 1, ����_In, Null), �޸���_In, �޸�����_In);
  
  Zl_δ��ҩƷ��¼_Insert(v_Lngid);
  
  --���¿��
  Zl_ҩƷ���_Update(v_Lngid, 0);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�������_Insert;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_Э�����_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Isstriked Exception;
  

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, �ⷿid, ������id, ���ϵ��, ҩƷid, ��д����, ����, ʵ������, �ɱ���, ���۽��, ���, ����, ����, Ч��, ��ҩ��λid, ��������, ��׼�ĺ�
    From ҩƷ�շ���¼ A
    Where NO = No_In And ���� = 3 And ��¼״̬ = 2
	Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 3 And ��¼״̬ = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;
  
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, �����, �������, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�)
    Select ҩƷ�շ���¼_Id.Nextval, 2, ����, No_In, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, nvl(����,0), ����, ����, Ч��, -��д����, -ʵ������, �ɱ���,
           -�ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 3 And ��¼״̬ = 3;
  
  
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    If v_ҩƷ�շ���¼.���ϵ�� = 1 Then
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 1);
    Else
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 0);
    End If;
  
    --������ۺ����
    Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Э�����_Strike;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
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
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ���, ������, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��, ��ҩ��ʽ, ע��֤��, �ƻ�id,
     ԭ����)
    Select v_������¼id, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 1, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, nvl(����,0), ����,
           ����, Ч��, n_����, -dblʵ������, -dblʵ������, �ɱ���, -dblʵ�ʳɱ�, ����, ���ۼ�, -dblʵ�ʽ��, -dblʵ�ʲ��, ժҪ, People_In, Date_In, ��ҩ��,
           People_In, Date_In, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ�ⷿ_In, ��ҩ��_In, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��_In, ��ҩ��ʽ, ע��֤��, �ƻ�id,
           ԭ����
    From ҩƷ�շ���¼
    Where ID = Billid_In;

  --����ǲ��ֳ�����������Ϊ1��ʵ������Ϊ������ʵ�������Ļ�
  --����������¼�Թ�������ҩ
  Select ҩƷ�շ���¼_Id.Nextval Into Lng������ From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, �ƻ�id, ԭ����)
    Select Lng������, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 3, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid,
           Decode(Lng����, 1, nvl(����,0), 3, Lng������, 0), Decode(Lng����, 3, ����_In, 1, ����, ����), Decode(Lng����, 3, ����_In, ����),
           Decode(Lng����, 3, Ч��_In, Ч��), n_����, Dblʵ������, Dblʵ������, �ɱ���, Dblʵ�ʳɱ�, ����, ���ۼ�, Dblʵ�ʽ��, Dblʵ�ʲ��, ժҪ, ������, ��������,
           Null, Null, Null, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��, �ƻ�id, ԭ����
    
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

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_ҩƷ�̵�_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Isstriked Exception;
  Err_Isbatch Exception;
  v_Err_Msg     Varchar2(255);
  n_Batch_Count Number;
  n_ҩƷid      ҩƷ�շ���¼.ҩƷid%Type;

  Cursor c_ҩƷ�շ���¼ Is
    Select a.Id, a.ʵ������, a.���۽��, a.���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As �Ƿ���, a.�ⷿid, a.ҩƷid, a.����, a.����, a.Ч��, a.����, a.ԭ����, a.������id,
           a.���ϵ��, a.����, a.��׼�ĺ�, a.��ҩ��λid, a.��������
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And NO = No_In And ���� = 12 And ��¼״̬ = 2
    Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 12 And ��¼״̬ = 1;
  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;

  --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������ 
  Select Count(*), Max(a.ҩƷid)
  Into n_Batch_Count, n_ҩƷid
  From ҩƷ�շ���¼ A, ҩƷ��� B
  Where a.ҩƷid = b.ҩƷid And a.No = No_In And a.���� = 12 And a.��¼״̬ = 3 And Nvl(a.����, 0) = 0 And
        ((Nvl(b.ҩ������, 0) = 1 And
        a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or Nvl(b.ҩ������, 0) = 1);

  If n_Batch_Count > 0 Then
    Begin
      Select ���� || '-' || ���� Into v_Err_Msg From �շ���ĿĿ¼ Where ID = n_ҩƷid;
    Exception
      When Others Then
        Null;
    End;
    v_Err_Msg := '�õ�����Ϊ:' || v_Err_Msg || Chr(10) || Chr(13) || '��ҩƷ,ԭ��������,�����ڷ�������˲�����ˣ�';
    Raise Err_Isbatch;
  End If;
  
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ, ������,
     ��������, �����, �������, Ƶ��, ����, ��׼�ĺ�, ��ҩ��λid, ��������, �ⷿ��λ)
    Select ҩƷ�շ���¼_Id.Nextval, 2, ����, NO, ���, �ⷿid, ������id, ���ϵ��, a.ҩƷid,
           Decode(Nvl(a.����, 0), 0, 0, (Decode(Nvl(b.ҩ�����, 0), 0, 0, a.����))), a.����, a.ԭ����, ����, Ч��, ��д����, a.����, -ʵ������,
           a.�ɱ���, �ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, Ƶ��, ����, a.��׼�ĺ�, a.��ҩ��λid, a.��������, 
		   a.�ⷿ��λ
    From (Select * From ҩƷ�շ���¼ Where NO = No_In And ���� = 12 And ��¼״̬ = 3 Order By ҩƷid) A, ҩƷ��� B
    Where a.ҩƷid = b.ҩƷid;
  
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --������
    If v_ҩƷ�շ���¼.���ϵ�� = 1 Then
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 1);
    Else
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 0);
    End If;
  
    --������ۺ����
    Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
  End Loop;
Exception
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�̵�_Strike;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_ҩƷ��������_Strike
(
  �д�_In       In Integer,
  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ҩƷid_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ

  v_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_������id   ҩƷ�շ���¼.������id%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_ԭ����       ҩƷ�շ���¼.ԭ����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_Ч��         ҩƷ�շ���¼.Ч��%Type;
  v_�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  v_�ɱ����     ҩƷ�շ���¼.�ɱ����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_���ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  v_���۽��     ҩƷ�շ���¼.���۽��%Type;
  v_���         ҩƷ�շ���¼.���%Type;
  v_ժҪ         ҩƷ�շ���¼.ժҪ%Type;
  v_ʣ������     ҩƷ�շ���¼.ʵ������%Type;
  v_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  v_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
  v_���ϵ��     ҩƷ�շ���¼.���ϵ��%Type;
  v_�����       ҩƷ�շ���¼.����%Type;
  v_�����λ     ҩƷ�շ���¼.��ҩ����%Type;
  v_��׼�ĺ�     ҩƷ�շ���¼.��׼�ĺ�%Type;
  v_��ֵ˰��     ҩƷ�շ���¼.Ƶ��%Type;

  v_�շ�id ҩƷ�շ���¼.Id%Type;
  Intdigit Number;

  n_�ϴι�Ӧ��id ҩƷ���.�ϴι�Ӧ��id%Type;
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
Begin
  --��ȡ���С��λ��
  Select Nvl(����, 2) Into Intdigit From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;

  If �д�_In = 1 Then
    Update ҩƷ�շ���¼
    Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
    Where NO = No_In And ���� = 11 And ��¼״̬ = ԭ��¼״̬_In;
    If Sql%RowCount = 0 Then
      Raise Err_Isstriked;
    End If;
  End If;

  --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
  Begin
    Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
    Into v_Druginf
    From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
    Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 11 And
          a.ҩƷid + 0 = ҩƷid_In And Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And
          ((Nvl(b.ҩ�����, 0) = 1 And
          a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or Nvl(b.ҩ������, 0) = 1) And
          Rownum = 1;
  Exception
    When Others Then
      v_Druginf := '';
  End;

  If v_Druginf Is Not Null Then
    Raise Err_Isbatch;
  End If;

  Select Sum(ʵ������) As ʣ������, Sum(�ɱ����) As ʣ��ɱ����, Sum(���۽��) As ʣ�����۽��, �ⷿid, ������id, ���ϵ��, Nvl(����, 0) As ����, ����, ԭ����, ����, Ч��, �ɱ���, ����, ���ۼ�,
         ժҪ, ����, ��ҩ����, ��׼�ĺ�, ��ҩ��λid, ��������, To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000'))) As ��ֵ˰��
  Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ⷿid, v_������id, v_���ϵ��, v_����, v_����, v_ԭ����, v_����, v_Ч��, v_�ɱ���, v_����, v_���ۼ�, v_ժҪ, v_�����,
       v_�����λ, v_��׼�ĺ�, n_�ϴι�Ӧ��id, d_�ϴ���������, v_��ֵ˰��
  From ҩƷ�շ���¼
  Where NO = No_In And ���� = 11 And ҩƷid = ҩƷid_In And ��� = ���_In
  Group By �ⷿid, ������id, ���ϵ��, Nvl(����, 0), ����, ԭ����, ����, Ч��, �ɱ���, ����, ���ۼ�, ժҪ, ����, ��ҩ����, ��׼�ĺ�, ��ҩ��λid, ��������,
           To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000')));

  --������������ʣ��������������
  If Abs(v_ʣ������) < Abs(��������_In) Then
    Raise Err_Isnonum;
  End If;

  v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
  v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
  v_���     := v_���۽�� - v_�ɱ����;

  Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ����, ��ҩ����, ժҪ,
     ������, ��������, �����, �������, ��׼�ĺ�, ��ҩ��λid, ��������, ����, Ƶ��)
  Values
    (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 11, No_In, ���_In, v_�ⷿid, v_������id, v_���ϵ��, ҩƷid_In, v_����, v_����,v_ԭ����, 
     v_����, v_Ч��, -��������_In, -��������_In, v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, v_�����, v_�����λ, v_ժҪ, ������_In, ��������_In, ������_In,
     ��������_In, v_��׼�ĺ�, n_�ϴι�Ӧ��id, d_�ϴ���������, v_����, v_��ֵ˰��);
  
  --���¿�棬������������
  Zl_ҩƷ���_Update(v_�շ�id, 3, 0);

  --������ۺ����
  Zl_ҩƷ�շ���¼_��������(v_�շ�id);
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']�����ܳ�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ��������_Strike;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_ҩƷ����_Strike
(
  �д�_In       In Integer,
  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ҩƷid_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ������ʽ_In   In Integer := 0, --0������������ʽ��1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��  
  ����ԭ��_In   In ҩƷ�շ���¼.����ԭ��%Type := Null
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ 

  v_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_�Է�����id   ҩƷ�շ���¼.�Է�����id%Type;
  v_������id   ҩƷ�շ���¼.������id%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_ԭ����       ҩƷ�շ���¼.ԭ����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_Ч��         ҩƷ�շ���¼.Ч��%Type;
  v_�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  v_�ɱ����     ҩƷ�շ���¼.�ɱ����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_���ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  v_���۽��     ҩƷ�շ���¼.���۽��%Type;
  v_���         ҩƷ�շ���¼.���%Type;
  v_ժҪ         ҩƷ�շ���¼.ժҪ%Type;
  v_ʣ������     ҩƷ�շ���¼.ʵ������%Type;
  v_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  v_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
  v_���ϵ��     ҩƷ�շ���¼.���ϵ��%Type;

  v_�շ�id   ҩƷ�շ���¼.Id%Type;
  v_������   ҩƷ�շ���¼.������%Type;
  v_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%Type;
  v_��ҩ��ʽ ҩƷ�շ���¼.��ҩ��ʽ%Type;

  v_�Ƿ���     �շ���ĿĿ¼.�Ƿ���%Type;
  Intdigit       Number;
  n_�ϴι�Ӧ��id ҩƷ���.�ϴι�Ӧ��id%Type;
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
  v_������������ Varchar2(4000);
Begin
  --��ȡ���С��λ�� 
  Select Nvl(����, 2) Into Intdigit From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;
  Select Nvl(�Ƿ���, 0) Into v_�Ƿ��� From �շ���ĿĿ¼ Where Id = ҩƷid_In;
  Select Zl_Getsysparameter('������������', 1305) Into v_������������ From Dual;

  If ������ʽ_In = 1 Then
    --�����������뵥�ݣ�����д����ˡ�������ڣ������¿���¼ 
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where No = No_In And ���� = 7 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ a, ҩƷ��� b, �շ���ĿĿ¼ i, �շ���Ŀ���� n
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 7 And
            Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.ҩƷid + 0 = ҩƷid_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(ʵ������) As ʣ������, Sum(�ɱ����) As ʣ��ɱ����, Sum(���۽��) As ʣ�����۽��, �ⷿid, �Է�����id, ������id, ���ϵ��, Nvl(����, 0) As ����, ����, ԭ����, ����, Ч��, �ɱ���,
           ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, v_����, v_����, v_ԭ����, v_����, v_Ч��, v_�ɱ���, v_����, v_���ۼ�,
         v_ժҪ, v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������
    From ҩƷ�շ���¼
    Where No = No_In And ���� = 7 And ҩƷid = ҩƷid_In And ��� = ���_In
    Group By �ⷿid, �Է�����id, ������id, ���ϵ��, Nvl(����, 0), ����, ԭ����, ����, Ч��, �ɱ���, ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������;
  
    --������������ʣ�������������� 
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
  
    Insert Into ҩƷ�շ���¼
      (Id, ��¼״̬, ����, No, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
       ������, ��������, �����, �������, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������, ����, ����ԭ��)
    Values
      (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 7, No_In, ���_In, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, ҩƷid_In, v_����,
       v_����, v_ԭ����, v_����, v_Ч��, -��������_In, -��������_In, v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, v_ժҪ, ������_In, ��������_In, Null, Null,
       v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������, v_����, ����ԭ��_In);
    
    Zl_δ��ҩƷ��¼_Insert(v_�շ�id);
    
  Elsif ������ʽ_In = 2 Then
    --����Ѳ����ĳ������뵥�ݣ���д����ˡ�������ڣ����¿���¼ 
  
    --��д����ˡ�������� 
    Update ҩƷ�շ���¼
    Set ����� = ������_In, ������� = ��������_In
    Where ���� = 7 And No = No_In And ��� = ���_In And ��¼״̬ = ԭ��¼״̬_In;
  
    --��ѯ��ǰ�м�¼�Ķ�ӦID
    Select Id
    Into v_�շ�id
    From ҩƷ�շ���¼
    Where ���� = 7 And No = No_In And ��� = ���_In And ��¼״̬ = ԭ��¼״̬_In;
  
    --���¿����Ϣ ���ó����൱����� 
    Zl_ҩƷ���_Update(v_�շ�id, 3, 0);
    
    Zl_δ��ҩƷ��¼_Delete(v_�շ�id);

    --����ҩƷ���洦�� 
    If v_��ҩ��ʽ = 1 Then
      Update ҩƷ����
      Set �������� = Nvl(��������, 0) + ��������_In, ʵ������ = Nvl(ʵ������, 0) + ��������_In, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + v_���۽��
      Where �ڼ� = To_Char(Sysdate, Decode(v_������������, '1', 'yyyymm', 'yyyy')) And ����id = v_�Է�����id And �ⷿid = v_�ⷿid And
            ҩƷid = ҩƷid_In;
      --��������������0�ļ�¼ɾ���� 
      Delete From ҩƷ���� Where Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0;
    End If;
  
    --������ۺ���� 
    Zl_ҩƷ�շ���¼_��������(v_�շ�id);
  Else
    --����������ʽ������������¼����д����ˡ�������ڣ����¿���¼      
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where No = No_In And ���� = 7 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%Rowcount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ a, ҩƷ��� b, �շ���ĿĿ¼ i, �շ���Ŀ���� n
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 7 And
            Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.ҩƷid + 0 = ҩƷid_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(ʵ������) As ʣ������, Sum(�ɱ����) As ʣ��ɱ����, Sum(���۽��) As ʣ�����۽��, �ⷿid, �Է�����id, ������id, ���ϵ��, Nvl(����, 0) As ����, ����, ԭ����, ����, Ч��, �ɱ���,
           ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, v_����, v_����, v_ԭ����, v_����, v_Ч��, v_�ɱ���, v_����, v_���ۼ�,
         v_ժҪ, v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������
    From ҩƷ�շ���¼
    Where No = No_In And ���� = 7 And ҩƷid = ҩƷid_In And ��� = ���_In
    Group By �ⷿid, �Է�����id, ������id, ���ϵ��, Nvl(����, 0), ����, ԭ����, ����, Ч��, �ɱ���, ����, ���ۼ�, ժҪ, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������;
  
    --������������ʣ�������������� 
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
    Insert Into ҩƷ�շ���¼
      (Id, ��¼״̬, ����, No, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
       ������, ��������, �����, �������, ������, ��׼�ĺ�, ��ҩ��ʽ, ��ҩ��λid, ��������, ����, ����ԭ��)
    Values
      (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 7, No_In, ���_In, v_�ⷿid, v_�Է�����id, v_������id, v_���ϵ��, ҩƷid_In, v_����,
       v_����, v_ԭ����, v_����, v_Ч��, -��������_In, -��������_In, v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, v_ժҪ, ������_In, ��������_In, ������_In,
       ��������_In, v_������, v_��׼�ĺ�, v_��ҩ��ʽ, n_�ϴι�Ӧ��id, d_�ϴ���������, v_����, ����ԭ��_In);
    
    --���¿����Ϣ ���ó����൱����� 
    Zl_ҩƷ���_Update(v_�շ�id, 3, 0);
    
    --����ҩƷ���洦�� 
    If v_��ҩ��ʽ = 1 Then
      Update ҩƷ����
      Set �������� = Nvl(��������, 0) + ��������_In, ʵ������ = Nvl(ʵ������, 0) + ��������_In, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + v_���۽��
      Where �ڼ� = To_Char(Sysdate, Decode(v_������������, '1', 'yyyymm', 'yyyy')) And ����id = v_�Է�����id And �ⷿid = v_�ⷿid And
            ҩƷid = ҩƷid_In;
      --��������������0�ļ�¼ɾ���� 
      Delete From ҩƷ���� Where Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0;
    End If;
  
    --������ۺ���� 
    Zl_ҩƷ�շ���¼_��������(v_�շ�id);
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102,
                            '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']�����ܳ�����[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_ҩƷ����_Strike;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_ҩƷ�ƿ�_Strike
(
  �д�_In       In Integer,
  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ҩƷid_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
  ������ʽ_In   In Integer := 0 --0������������ʽ��1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf      Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ
  v_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  v_�ɱ����     ҩƷ�շ���¼.�ɱ����%Type;
  v_���ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  v_���۽��     ҩƷ�շ���¼.���۽��%Type;
  v_���         ҩƷ�շ���¼.���%Type;
  v_ʣ������     ҩƷ�շ���¼.ʵ������%Type;
  v_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  v_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
  v_�շ�id       ҩƷ�շ���¼.Id%Type;
  v_��׼�ĺ�     ҩƷ�շ���¼.��׼�ĺ�%Type;

  v_ҩ����� Integer;
  v_ҩ������ Integer;
  Intdigit   Number;
  n_�������� Number;

  Cursor c_ҩƷ�շ���¼ Is
    Select a.Id, a.���, a.�ⷿid, a.�Է�����id, a.������id, a.���ϵ��, a.ҩƷid, nvl(a.����,0) ����, a.����, a.ԭ����, a.����, a.Ч��, a.��ҩ��, a.��ҩ����, a.ժҪ,
           a.��ҩ��λid, a.��׼�ĺ�, a.��������, a.�ɱ���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As ʱ��, a.����, a.����, a.Ƶ��
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = No_In And a.���� = 6 And (a.��� >= ���_In And a.��� <= ���_In + 1) And
          (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid,a.����;

  Cursor c_���������¼ Is
    Select a.Id, a.���, a.�ⷿid, a.�Է�����id, a.������id, a.���ϵ��, a.ҩƷid, nvl(a.����,0) ����, a.����, a.ԭ����, a.����, a.Ч��, a.��ҩ��, a.��ҩ����, a.ժҪ,
           a.��ҩ��λid, a.��׼�ĺ�, a.��������, a.�ɱ���, a.ʵ������, a.���۽��, a.���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As ʱ��, a.����, a.����, a.Ƶ��
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = No_In And a.���� = 6 And (a.��� >= ���_In And a.��� <= ���_In + 1) And
          (a.��¼״̬ = ԭ��¼״̬_In And Mod(a.��¼״̬, 3) = 2) And a.������� Is Null
    Order By a.ҩƷid,a.����;
Begin
  --��ȡ���С��λ��
  Select Nvl(����, 2) Into Intdigit From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;

  If ������ʽ_In = 0 Then
    n_�������� := 0;
  Else
    n_�������� := 1;
  End If;

  If ������ʽ_In = 1 Then
    --�����������뵥�ݣ�����д����ˡ�������ڣ������¿���¼
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where NO = No_In And ���� = 6 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 6 And
            a.ҩƷid + 0 = ҩƷid_In And Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.��� = ���_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.ʵ������) As ʣ������, Sum(a.�ɱ����) As ʣ��ɱ����, Sum(a.���۽��) As ʣ�����۽��, a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0),
           b.ҩ�����, b.ҩ������, a.��׼�ĺ�
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ɱ���, v_���ۼ�, v_�ⷿid, v_����, v_ҩ�����, v_ҩ������, v_��׼�ĺ�
    From ҩƷ�շ���¼ A, ҩƷ��� B
    Where a.No = No_In And a.ҩƷid = b.ҩƷid And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In
    Group By a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0), b.ҩ�����, b.ҩ������, a.��׼�ĺ�;
  
    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    Select Nvl(a.����, 0)
    Into v_����
    From ҩƷ�շ���¼ A
    Where a.No = No_In And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In + 1 And Mod(a.��¼״̬, 3) = 0;
  
    --������������ʣ��������������
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    
      Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��,
         ���, ժҪ, ������, ��������, �����, �������, ��ҩ��, ��ҩ����, ��ҩ��λid, ��׼�ĺ�, ��������, ����, ����, Ƶ��)
      Values
        (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 6, No_In, v_ҩƷ�շ���¼.���, v_ҩƷ�շ���¼.�ⷿid, v_ҩƷ�շ���¼.�Է�����id,
         v_ҩƷ�շ���¼.������id, v_ҩƷ�շ���¼.���ϵ��, ҩƷid_In, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.ԭ����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ч��,
         -��������_In, -��������_In, v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, ժҪ_In, ������_In, ��������_In, Null, Null, v_ҩƷ�շ���¼.��ҩ��,
         v_ҩƷ�շ���¼.��ҩ����, v_ҩƷ�շ���¼.��ҩ��λid, v_ҩƷ�շ���¼.��׼�ĺ�, v_ҩƷ�շ���¼.��������, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ƶ��);
    
      Zl_δ��ҩƷ��¼_Insert(v_�շ�id);
    
      --�����棬ԭ����Ǳ��൱�ڳ���
      If v_ҩƷ�շ���¼.���ϵ�� = 1 Then
        Zl_ҩƷ���_Update(v_�շ�id, 0, 1);
      End If;
    
      If v_ҩƷ�շ���¼.���ϵ�� = -1 Then
        v_�ⷿid := v_ҩƷ�շ���¼.�ⷿid;
      End If;
    End Loop;
  
  Elsif ������ʽ_In = 2 Then
    --����Ѳ����ĳ������뵥�ݣ���д����ˡ�������ڣ����¿���¼
    For v_ҩƷ�շ���¼ In c_���������¼ Loop
      --��д����ˡ��������
      Update ҩƷ�շ���¼
      Set ����� = ������_In, ������� = ��������_In
      Where NO = No_In And ���� = 6 And ID = v_ҩƷ�շ���¼.Id;
    
      --����ҩƷ�������Ӧ���ݣ�ע����ʱ������������Ǹ���
      --����Ϊ1��ʾ�������ʱ�¿�������������ԭ����ⷿ�����˿��������Ͳ����ٸ��¿���������
      If v_ҩƷ�շ���¼.���ϵ�� = 1 Then
        Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 1, n_��������);
      Else
        Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 0);
      End If;
    
      Zl_δ��ҩƷ��¼_Delete(v_ҩƷ�շ���¼.Id);
    
      If v_ҩƷ�շ���¼.���ϵ�� = -1 Then
        v_�ⷿid := v_ҩƷ�շ���¼.�ⷿid;
      
      End If;
    
      --������ۺ����
      Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
    End Loop;
  
    b_Message.Zlhis_Drug_004(No_In);
  Else
    --����������ʽ������������¼����д����ˡ�������ڣ����¿���¼
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where NO = No_In And ���� = 6 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 6 And
            a.ҩƷid + 0 = ҩƷid_In And Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.��� = ���_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.ʵ������) As ʣ������, Sum(a.�ɱ����) As ʣ��ɱ����, Sum(a.���۽��) As ʣ�����۽��, a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0),
           b.ҩ�����, b.ҩ������, a.��׼�ĺ�
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ɱ���, v_���ۼ�, v_�ⷿid, v_����, v_ҩ�����, v_ҩ������, v_��׼�ĺ�
    From ҩƷ�շ���¼ A, ҩƷ��� B
    Where a.No = No_In And a.ҩƷid = b.ҩƷid And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In
    Group By a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0), b.ҩ�����, b.ҩ������, a.��׼�ĺ�;
  
    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    Select Nvl(a.����, 0)
    Into v_����
    From ҩƷ�շ���¼ A
    Where a.No = No_In And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In + 1 And Mod(a.��¼״̬, 3) = 0;
  
    --������������ʣ��������������
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
      Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ԭ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��,
         ���, ժҪ, ������, ��������, �����, �������, ��ҩ��, ��ҩ����, ��ҩ��λid, ��׼�ĺ�, ��������, ����, ����, Ƶ��)
      Values
        (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 6, No_In, v_ҩƷ�շ���¼.���, v_ҩƷ�շ���¼.�ⷿid, v_ҩƷ�շ���¼.�Է�����id,
         v_ҩƷ�շ���¼.������id, v_ҩƷ�շ���¼.���ϵ��, ҩƷid_In, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.ԭ����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ч��,
         -��������_In, -��������_In, v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, ժҪ_In, ������_In, ��������_In, ������_In, ��������_In,
         v_ҩƷ�շ���¼.��ҩ��, v_ҩƷ�շ���¼.��ҩ����, v_ҩƷ�շ���¼.��ҩ��λid, v_ҩƷ�շ���¼.��׼�ĺ�, v_ҩƷ�շ���¼.��������, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����,
         v_ҩƷ�շ���¼.Ƶ��);
    
      --����ҩƷ�������Ӧ����
      If v_ҩƷ�շ���¼.���ϵ�� = 1 Then
        Zl_ҩƷ���_Update(v_�շ�id, 3, 1, n_��������);
      Else
        Zl_ҩƷ���_Update(v_�շ�id, 3, 0);
      End If;
    
      --������ۺ����
      Zl_ҩƷ�շ���¼_��������(v_�շ�id);
    End Loop;
  
    b_Message.Zlhis_Drug_004(No_In);
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']�����ܳ�����[ZLSOFT]');
  When Err_Isnonum Then
    Raise_Application_Error(-20103, '[ZLSOFT]�õ����е�' || Ceil(���_In / 2) || '�е�ҩƷ����������������ʣ������ݣ����ܳ�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�ƿ�_Strike;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Procedure Zl_�������_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Isstriked Exception;

  v_������id ҩƷ�շ���¼.������id%Type;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, ���۽��, ���, ��ҩ��λid, ��������, ��׼�ĺ�, ����
    From ҩƷ�շ���¼ A
    Where NO = No_In And ���� = 2 And ��¼״̬ = 2
    Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 2 And ��¼״̬ = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;
  
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, �����, �������, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�, ����)
    Select ҩƷ�շ���¼_Id.Nextval, 2, 2, No_In, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, nvl(����,0), ����, ����, Ч��, -��д����, -ʵ������, �ɱ���,
           -�ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�, ����
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 2 And ��¼״̬ = 3;
  

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    If v_ҩƷ�շ���¼.���ϵ�� = 1 Then
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 1);
    Else
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 3, 0);
    End If;
  
    --������ۺ����
    Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Strike;
/

--139589:����,2019-04-18,��ֹ����Ϊ�����⴦��
Create Or Replace Function Zl_Fun_Getbatchnum
(
  ҩƷid_In   ҩƷ���Ŷ���.ҩƷid%Type,
  ��������_In ҩƷ���Ŷ���.��������%Type,
  ����_In     ҩƷ���Ŷ���.����%Type,
  �ɱ���_In   ҩƷ���Ŷ���.�ɱ���%Type,
  �ۼ�_In     ҩƷ���Ŷ���.�ۼ�%Type,
  ������_In   ҩƷ���Ŷ���.����%Type,
  ��Ӧ��ID_In   ҩƷ���Ŷ���.��Ӧ��ID%Type
) Return Number Is
  --���ܣ�ҩƷ����������¼ʱ���ݴ��ݹ����Ĳ����Ҷ�Ӧ������
  --����ֵ����ѯ�������Σ��������>0��˵���ҵ�������,�������=0��˵��û���ҵ�
  --������
  --     ��������_in����⴫�ݹ�����������
  --     ����_in�����ʱ¼�������
  --     �ɱ���_in ���ʱ�ĳɱ���
  --     �ۼ�_in  ���ʱ���ۼ�
  --     
  n_����     ҩƷ���Ŷ���.����%Type;
  n_ҩ���װ ҩƷ���.ҩ���װ%Type;
  n_�Ƿ��� �շ���ĿĿ¼.�Ƿ���%Type;
  n_Count    Number(1);
Begin
  --ֻ�����������Һ����Ų�Ϊ�յ����
  If ��������_In Is Not Null And ����_In Is Not Null Then
    Begin
      Select nvl(����,0)
      Into n_����
      From ҩƷ���Ŷ���
      Where ҩƷid = ҩƷid_In And Nvl(��������, 'a') = Nvl(��������_In, 'a') And Nvl(����, 'b') = Nvl(����_In, 'b') And �ɱ��� = �ɱ���_In And
            �ۼ� = �ۼ�_In And Nvl(��Ӧ��id, 0) = Nvl(��Ӧ��id_In, 0);
    Exception
      When Others Then
        n_���� := ������_In;
      
        If n_���� > 0 Then
          --��������ظ���¼
          Begin
            Select 1
            Into n_Count
            From ҩƷ���Ŷ���
            Where ҩƷid = ҩƷid_In And Nvl(��������, 'a') = Nvl(��������_In, 'a') And Nvl(����, 'b') = Nvl(����_In, 'b') And
                  nvl(����,0) = n_����;
          Exception
            When Others Then
              n_Count := 0;
          End;
          
          --û���ظ���¼���ܲ���
          If n_Count = 0 Then
            Insert Into ҩƷ���Ŷ���
              (ҩƷid, ��������, ����, ����, �ɱ���, �ۼ�,��Ӧ��ID)
            Values
              (ҩƷid_In, ��������_In, ����_In, ������_In, �ɱ���_In, �ۼ�_In,��Ӧ��ID_In);
          End If;
        End If;
    End;
  End If;

  Return(n_����);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Getbatchnum;
/






------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0058' Where ���=&n_System;
Commit;
