----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------

--123419:����,2018-05-10,�����ֹ����Զ����ʷ��ø��ӱ�־����
create or replace view ��Ժ�����Զ����� as
Select p.����, p.����id, p.��ҳid, Nvl(a.����, i.����) As ����, Nvl(a.�Ա�, i.�Ա�) As �Ա�, Nvl(a.����, i.����) As ����, Nvl(a.סԺ��, i.סԺ��) As סԺ��,
       a.�ѱ�, p.����id, p.����id, p.����, p.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, 1 As ��־, p.�ּ� As ��׼����, p.��ʼ����, p.��ֹ����,
       p.��ֹ���� - p.��ʼ���� As ����, p.����, p.����ҽʦ, p.���λ�ʿ, p.����Ա���, p.����Ա����,p.ҽ��С��id
From ������Ϣ I, ������ҳ A,
     (Select 2 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����,
               b.ҽ��С��id
       From �Զ��Ƽ���Ŀ A,
            (Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, ��λ�ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��,
                     a.ҽ��С��id
              From ���˱䶯��¼ A
              Where ��ʼԭ�� <> 10
              Union All
              Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����,
                     �ϴμ���ʱ��, b.ҽ��С��id
              From ���˱䶯��¼ B, �շѴ�����Ŀ I
              Where b.��λ�ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) B, �շѼ�Ŀ P
       Where a.����id = b.����id And Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�����־ = 1 And b.��λ�ȼ�id = p.�շ�ϸĿid And Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2))) And p.�۸�ȼ� Is Null
       Union All
       Select 1 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����,
              b.ҽ��С��id
       From �Զ��Ƽ���Ŀ A,
            (Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, ����ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��, ҽ��С��id
              From ���˱䶯��¼
              Where ��ʼԭ�� <> 10
              Union All
              Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, i.����id As ����ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����,
                     �ϴμ���ʱ��, b.ҽ��С��id
              From ���˱䶯��¼ B, �շѴ�����Ŀ I
              Where b.����ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) B, �շѼ�Ŀ P, �շ���ĿĿ¼ C
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And a.�����־ = 2 And
             b.����ȼ�id = p.�շ�ϸĿid And b.����ȼ�id = c.Id And Nvl(c.���㷽ʽ, 0) <> 1 And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2))) And p.�۸�ȼ� Is Null
       Union All
       Select 3 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, a.����,
              b.ҽ��С��id
       From (Select ����id, �����־, �շ�ϸĿid, 1 As ����, ��������
              From �Զ��Ƽ���Ŀ
              Union All
              Select ����id, �����־, ����id, i.�������� As ����, ��������
              From �Զ��Ƽ���Ŀ A, �շѴ�����Ŀ I
              Where a.�շ�ϸĿid = i.����id And i.���д��� > 0) A, ���˱䶯��¼ B, �շѼ�Ŀ P
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And b.��ʼԭ�� <> 10 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�շ�ϸĿid = p.�շ�ϸĿid And (a.�����־ = 6 And b.��λ�ȼ�id Is Not Null Or a.�����־ =7) And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2))) And p.�۸�ȼ� Is Null) P
Where i.����id = p.����id And a.����id = p.����id And a.��ҳid = p.��ҳid;

create or replace view ��Ժ�����Զ����� as
Select p.����,p.����id, p.��ҳid, Nvl(a.����, i.����) As ����, Nvl(a.�Ա�, i.�Ա�) As �Ա�, Nvl(a.����, i.����) As ����, Nvl(a.סԺ��, i.סԺ��) As סԺ��,
       a.�ѱ�, p.����id, p.����id, p.����, p.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, 1 As ��־, p.�ּ� As ��׼����, p.��ʼ����, p.��ֹ����,
       p.��ֹ���� - p.��ʼ���� As ����, p.����, p.����ҽʦ, p.���λ�ʿ, p.����Ա���, p.����Ա����,p.ҽ��С��id
From ������Ϣ I, ������ҳ A,
     (Select 2 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����,
              b.ҽ��С��id
       From �Զ��Ƽ���Ŀ A,
            (Select a.����id, a.��ҳid, a.��ʼʱ��, a.���Ӵ�λ, a.����id, a.����id, a.����, a.��λ�ȼ�id, 1 As ����, a.���λ�ʿ, a.����ҽʦ, a.��ֹʱ��,
                     a.����Ա���, a.����Ա����, a.�ϴμ���ʱ��, a.ҽ��С��id
              From ���˱䶯��¼ A, ������Ϣ B
              Where a.��ʼԭ�� <> 10 And a.����id = b.����id And a.��ҳid = b.��ҳid And b.��Ժ = 1
              Union All
              Select b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���,
                     ����Ա����, �ϴμ���ʱ��, b.ҽ��С��id
              From ���˱䶯��¼ B, �շѴ�����Ŀ I, ������Ϣ C
              Where b.����id = c.����id And b.��ҳid = c.��ҳid And c.��Ժ = 1 And b.��λ�ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) B,
            �շѼ�Ŀ P
       Where a.����id = b.����id And Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�����־ = 1 And b.��λ�ȼ�id = p.�շ�ϸĿid And Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2))) And p.�۸�ȼ� Is Null
       Union All
       Select 1 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����,
              b.ҽ��С��id
       From �Զ��Ƽ���Ŀ A,
            (Select a.����id, a.��ҳid, ��ʼʱ��, ���Ӵ�λ, a.����id, a.����id, ����, ����ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��,
                     a.ҽ��С��id
              From ���˱䶯��¼ A, ������Ϣ B
              Where ��ʼԭ�� <> 10 And a.����id = b.����id And a.��ҳid = b.��ҳid And b.��Ժ = 1
              Union All
              Select b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ����ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���,
                     ����Ա����, �ϴμ���ʱ��, b.ҽ��С��id
              From ���˱䶯��¼ B, �շѴ�����Ŀ I, ������Ϣ C
              Where b.����ȼ�id = i.����id And b.����id = c.����id And b.��ҳid = c.��ҳid And c.��Ժ = 1 And b.��ʼԭ�� <> 10 And i.���д��� > 0) B,
            �շѼ�Ŀ P, �շ���ĿĿ¼ C
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And a.�����־ = 2 And
             b.����ȼ�id = p.�շ�ϸĿid And b.����ȼ�id = c.Id And Nvl(c.���㷽ʽ, 0) <> 1 And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2))) And p.�۸�ȼ� Is Null
       Union All
       Select 3 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, a.����,
               b.ҽ��С��id
       From (Select ����id, �����־, �շ�ϸĿid, 1 As ����, ��������
              From �Զ��Ƽ���Ŀ
              Union All
              Select ����id, �����־, ����id, i.�������� As ����, ��������
              From �Զ��Ƽ���Ŀ A, �շѴ�����Ŀ I
              Where a.�շ�ϸĿid = i.����id And i.���д��� > 0) A, ���˱䶯��¼ B, �շѼ�Ŀ P, ������Ϣ C
       Where a.����id = b.����id And b.����id = c.����id And b.��ҳid = c.��ҳid And c.��Ժ = 1 And b.���Ӵ�λ <> 1 And b.��ʼԭ�� <> 10 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�շ�ϸĿid = p.�շ�ϸĿid And (a.�����־ = 6 And b.��λ�ȼ�id Is Not Null Or a.�����־=7) And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2))) And p.�۸�ȼ� Is Null) P
Where i.����id = p.����id And a.����id = p.����id And a.��ҳid = p.��ҳid;

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
--123419:����,2018-05-10,�����ֹ����Զ����ʷ��ø��ӱ�־����
CREATE OR REPLACE Procedure Zl1_Autocptone
( 
  ����id_In   In Number, 
  ��ҳid_In   In Number, 
  �ڼ�_In     In Varchar2, 
  ��Ժ����_In In Number := 0, 
  ǿ�Ƽ���_In In Number := 0 
) As 
 
  ------------------------------------------------------------------------- 
  --����˵�������ָ������ָ���ڼ��Զ��Ƽ���Ŀ�������Զ��������Ŀ���м��ʴ��� 
  --          1��ϵͳ���ȸ���ϵͳ����"���������Զ��Ʒ�"���޸������ò����Զ����ʼ�¼��־; 
  --          2���ۺϲ��˵Ĵ�λ�仯�����ת�������������ȶ������أ�����ڼ��ȡ����˷� 
  --             �����ɷ��õ���ȷ���㣺 
  --             ��������Ѿ����㣬���޸ı�־Ϊ����;���δ���㣬������µ��Զ����ʼ�¼; 
  --             ������ǰ�Ĵ������ļ�¼; 
  --             ͳ�Ʊ��α䶯(����������)����д����ͻ��ܱ�; 
  --��ڲ����� 
  --       ����ID_IN  number    �������ID 
  --       ��ҳID_IN  number    ������ҳID������������ͬȷ����Ҫ����Ĳ��� 
  --       �ڼ�_IN  varchar2     ��Ҫ�������С�ڼ� 
  --       ��Ժ����_IN number   Ϊ1ʱ,��������Ժ���˵ķ��� 
  --       ǿ�Ƽ���_IN number   Ϊ1ʱ,���ܲ�����ҳ.��ֹ�Զ��������Կ��� 
  --���ù�ϵ��zl1_AutoCptPati/zl1_AutoCptWard/zl1_AutoCptAll ���ñ����� 
  n_Count  Number(5);
   
  Cursor v_Autocur 
  ( 
    �ڼ�_In Varchar2, 
    Insure  ������ҳ.����%Type 
  ) Is 
    Select l.����,l.����id, l.��ҳid, l.����, l.�Ա�, l.����, l.סԺ��, l.�ѱ�, l.����id, l.����id, l.����, l.���Ӵ�λ, l.�շ�ϸĿid, l.������Ŀid, l.��־, l.��׼����, 
           Greatest(l.��ʼ����, Trunc(p.��ʼ����)) As ��ʼ����, l.��ֹ����, l.����, l.����, l.����ҽʦ, l.���λ�ʿ, l.����Ա���, l.����Ա����, i.����, i.����id, 
           k.�㷨, k.ͳ��ȶ�, l.ҽ��С��id 
    From (Select * From ��Ժ�����Զ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In) L, 
         (Select Min(��ʼ����) As ��ʼ���� From �ڼ�� Where �ڼ� >= �ڼ�_In) P, ����֧����Ŀ I, ����֧������ K 
    Where Trunc(l.��ֹ����) >= Trunc(p.��ʼ����) And l.�շ�ϸĿid = i.�շ�ϸĿid(+) And i.����(+) = Insure And i.����id = k.Id(+) 
    Order By l.��ʼ����; 
 
  Cursor v_Autocurzy 
  ( 
    �ڼ�_In Varchar2, 
    Insure  ������ҳ.����%Type 
  ) Is 
    Select l.����, l.����id, l.��ҳid, l.����, l.�Ա�, l.����, l.סԺ��, l.�ѱ�, l.����id, l.����id, l.����, l.���Ӵ�λ, l.�շ�ϸĿid, l.������Ŀid, l.��־, l.��׼����, 
           Greatest(l.��ʼ����, Trunc(p.��ʼ����)) As ��ʼ����, l.��ֹ����, l.����, l.����, l.����ҽʦ, l.���λ�ʿ, l.����Ա���, l.����Ա����, i.����, i.����id, 
           k.�㷨, k.ͳ��ȶ�,l.ҽ��С��id 
    From (Select * From ��Ժ�����Զ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In) L, 
         (Select Min(��ʼ����) As ��ʼ���� From �ڼ�� Where �ڼ� >= �ڼ�_In) P, ����֧����Ŀ I, ����֧������ K 
    Where Trunc(l.��ֹ����) >= Trunc(p.��ʼ����) And l.�շ�ϸĿid = i.�շ�ϸĿid(+) And i.����(+) = Insure And i.����id = k.Id(+) 
    Order By l.��ʼ����; 
 
  n_Insure       ������ҳ.����%Type; 
  v_Billno       Varchar2(8); --���ñ�ʵ�ʵ��Զ����ʺ��� 
  n_Datecount    Integer; --���ڼ����� 
  d_Datefrom     Date; --��ʼ�������� 
  d_Dateto       Date; --��ֹ�������� 
  d_Datelast     Date; 
  n_Billcount    Number(5) := 0; --������ż����� 
  n_Exsetax      Number(16, 2) := 0; --������ȡ���� 
  n_Exsetax_Temp Number(16, 2) := 0; --������ȡ���� 
  n_Summoney     Number(16, 2) := 0; --��� 
 
  Cursor v_Sumcur 
  ( 
    Billno    Varchar2, 
    Datestart Date 
  ) Is 
    Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Decode(���ӱ�־, 0, 1, -1) * Ӧ�ս��) As Ӧ�ս��, 
           Sum(Decode(���ӱ�־, 0, 1, -1) * ʵ�ս��) As ʵ�ս�� 
    From סԺ���ü�¼ 
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And 
          (NO = Billno Or ���ӱ�־ = 5 And ����ʱ�� >= Datestart) 
    Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid; 
 
  n_Dec            Number; --���С��λ�� 
  d_�Ǽ�ʱ��       Date; --�Ǽ�ʱ�� 
  d_����ʱ��       Date; --����ʱ�� 
  n_Dates          Number(3, 1); --��ǰ��¼��������ȫ��Ϊ1 
  n_Do             Number(1); 
  n_����ֵ         �������.Ԥ�����%Type; 
  n_Delete         Number; 
  n_ҽ��С��id     סԺ���ü�¼.ҽ��С��id%Type; 
  n_��������׼   Number(2); --����Ѽ����׼ 
  n_�շ�ϸĿid     Number(18); 
  n_Temp           Number(18); 
  l_����id         t_Numlist := t_Numlist(); 
  l_����ȼ�       t_Numlist := t_Numlist(); 
  n_������Ŀ       Number(2); --1:�ǻ�����Ŀ;0-�ǲ����� 
  n_�۸�           �շѼ�Ŀ.�ּ�%Type; 
  n_�����Ѵ���     Number(2); --1-������Ѿ�����,;0-δ���� 
  n_������Ŀid     Number(18); 
  n_������Ŀ       Number(2); 
  n_��˱�־       ������ҳ.��˱�־%Type; 
  n_סԺ״̬       ������ҳ.״̬%Type; 
  n_������˷�ʽ   Number(2); 
  n_δ��ƽ�ֹ���� Number(2); 
  n_��ֹ�Զ�����   Number(2); 
 
  n_���˲���id סԺ���ü�¼.���˲���id%Type; 
  n_��������id סԺ���ü�¼.��������id%Type; 
 
  --�Ѿ������˵Ļ������� 
  Type t_����_Rec Is Record( 
    �շ�ϸĿid �շ���ĿĿ¼.Id%Type, 
    ����       Date); 
  Type t_���� Is Table Of t_����_Rec; 
  c_���� t_���� := t_����(); 
Begin 
  Begin 
    Select ����, Nvl(��˱�־, 0), Nvl(״̬, 0), Nvl(�Ƿ��ֹ�Զ�����, 0) 
    Into n_Insure, n_��˱�־, n_סԺ״̬, n_��ֹ�Զ����� 
    From ������ҳ 
    Where ����id = ����id_In And ��ҳid = ��ҳid_In; 
  Exception 
    When Others Then 
      Return; 
  End; 
 
  If ǿ�Ƽ���_In = 0 And n_��ֹ�Զ����� = 1 Then 
    Return; 
  End If; 
 
  n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0); 
  n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0); 
  If n_������˷�ʽ = 1 And Nvl(n_��˱�־, 0) >= 1 Then 
    Return; 
  End If; 
  If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then 
    Return; 
  End If; 
 
  v_Billno := Nextno(17); 
 
  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(160), '0')) 
  Into n_Dec, n_��������׼ 
  From Dual; 
 
  --ÿ��5����ǰ������¼ʱ��Ǽ�Ϊ���죬����Ǽ�Ϊ��ʱ 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_�Ǽ�ʱ�� 
  From Dual; 
 
  --�����ò��˵ļ�¼,�����ظ����� 
  Update ������ҳ Set ״̬ = ״̬ Where ����id = ����id_In And ��ҳid = ��ҳid_In; 
 
  ----------------------------------------------------------------- 
  d_Datefrom := Sysdate + 1000; 
  d_Dateto   := Sysdate - 1000; 
  n_Do       := 0; 
  -------------------------------------------------------------------- 
  If n_��������׼ = 1 Then 
    --ͬ������߼�λ�Ļ����Ϊ׼,�Ƚ��令��ȼ���ס, 
    For v_���� In (Select Distinct ����ȼ�id 
                 From (Select ����ȼ�id 
                        From ���˱䶯��¼ 
                        Where ��ʼԭ�� <> 10 And ����id = ����id_In And ��ҳid = ��ҳid_In 
                        Union All 
                        Select i.����id As ����ȼ�id 
                        From ���˱䶯��¼ B, �շѴ�����Ŀ I 
                        Where b.����ȼ�id = i.����id And ����id = ����id_In And ��ҳid = ��ҳid_In And b.��ʼԭ�� <> 10 And i.���д��� > 0)) Loop 
      If Nvl(v_����.����ȼ�id, 0) <> 0 Then 
        l_����id.Extend; 
        l_����id(l_����id.Count) := v_����.����ȼ�id; 
      End If; 
    End Loop; 
  End If; 
  ----------------------------------------------------------------- 
  --ѭ���������������������ȷ���¼���ļ�¼ 
  ----------------------------------------------------------------- 
  If ��Ժ����_In = 1 Then 
    For v_Currrow In v_Autocurzy(�ڼ�_In, n_Insure) Loop 
      If v_Currrow.ҽ��С��id Is Null Then 
        n_ҽ��С��id := Zl_ҽ��С��_Get(v_Currrow.����id, v_Currrow.����Ա����, v_Currrow.����id, v_Currrow.��ҳid, d_����ʱ��); 
      Else 
        n_ҽ��С��id := v_Currrow.ҽ��С��id; 
      End If; 
 
      If d_Datefrom > v_Currrow.��ʼ���� Then 
        d_Datefrom := v_Currrow.��ʼ����; 
        n_Do       := 1; 
        --�����ο�ʼ����ʱ���Ժ���Ѽ����¼��־�޸� 
        Update סԺ���ü�¼ 
        Set ���ӱ�־ = 5 
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ <> 8 And Nvl(ҽ�����, 0) = 0 And 
              ����ʱ�� >= v_Currrow.��ʼ����; 
      End If; 
 
      If d_Dateto < v_Currrow.��ֹ���� Then 
        d_Dateto := v_Currrow.��ֹ����; 
      End If; 
      n_�շ�ϸĿid := v_Currrow.�շ�ϸĿid; 
      n_������Ŀ   := 0; 
      --����Ѽ����׼:0-�����һ�λ������;1-���۸���ߵĻ���ȼ����㡣 
      If n_��������׼ = 1 Then 
        --��ȷ���Ƿ�����Ŀ,�����,����Ҫ���½��м��� 
        Select Count(*) Into n_������Ŀ From Table(l_����id) Where Column_Value = n_�շ�ϸĿid; 
      End If; 
 
      --��ȡ��ǰ������Ŀ���շѱ��� 
      Begin 
        Select ʵ�ձ��� 
        Into n_Exsetax 
        From (Select ʵ�ձ��� 
               From �ѱ���ϸ 
               Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_Currrow.�շ�ϸĿid And 
                     (Abs(v_Currrow.��׼���� * v_Currrow.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) 
               Union All 
               Select ʵ�ձ��� 
               From �ѱ���ϸ 
               Where �ѱ� = v_Currrow.�ѱ� And ������Ŀid = v_Currrow.������Ŀid And 
                     (Abs(v_Currrow.��׼���� * v_Currrow.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) And Not Exists 
                (Select 1 From �ѱ���ϸ Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_Currrow.�շ�ϸĿid)); 
      Exception 
        When Others Then 
          n_Exsetax := 100.00; 
      End; 
 
      n_Exsetax := Nvl(n_Exsetax, 100); 
      For n_Datecount In 0 .. (Trunc(v_Currrow.��ֹ���� + 0.5) - Trunc(v_Currrow.��ʼ����)) - 1 Loop 
        d_����ʱ�� := Greatest(v_Currrow.��ʼ����, Trunc(v_Currrow.��ʼ���� + n_Datecount)); 
        n_Dates    := Least(Trunc(v_Currrow.��ʼ���� + n_Datecount + 1), v_Currrow.��ֹ����) - 
                      Greatest(v_Currrow.��ʼ����, Trunc(v_Currrow.��ʼ���� + n_Datecount)); 
        --�ж��Ƿ��ֹ�����
        Select Count(1)
        Into n_Count
        From סԺ���ü�¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 2 And Nvl(���ӱ�־, 0) = 1 And
              �շ���� = Decode(v_Currrow.����, 1, 'H', 2, 'J', �շ����) And ����ʱ�� = d_����ʱ�� And
              �շ�ϸĿid = Decode(v_Currrow.����, 3, v_Currrow.�շ�ϸĿid, �շ�ϸĿid); 

        n_�����Ѵ��� := 0; 
        If n_������Ŀ = 1 Then 
          --1.�ȼ�鵱���Ƿ���ڻ���䶯,ֻ�д��ڶ������䶯��,�Ż�ȥ����(������ĿΪ׼) 
          n_������Ŀ := 1; 
          If l_����ȼ�.Count > 0 Then 
            l_����ȼ�.Delete; 
          End If; 
          For v_���� In (Select Distinct ����ȼ�id 
                       From ���˱䶯��¼ 
                       Where ��ʼԭ�� <> 10 And ����id = ����id_In And ��ҳid = ��ҳid_In And 
                             (Trunc(��ʼʱ��) = Trunc(d_����ʱ��) Or Trunc(Nvl(��ֹʱ��, Sysdate)) = Trunc(d_����ʱ��))) Loop 
            If Nvl(v_����.����ȼ�id, 0) <> 0 Then 
              l_����ȼ�.Extend; 
              l_����ȼ�(l_����ȼ�.Count) := v_����.����ȼ�id; 
              If Nvl(v_����.����ȼ�id, 0) = Nvl(v_Currrow.�շ�ϸĿid, 0) Then 
                n_������Ŀ := 0; 
              End If; 
            End If; 
          End Loop; 
          If l_����ȼ�.Count > 1 Then 
            --2. �����������ϱ䶯,��ȡ��λ��ߵ� 
            n_Temp       := v_Currrow.�շ�ϸĿid; 
            n_�۸�       := Nvl(v_Currrow.��׼����, 0); 
            n_������Ŀid := v_Currrow.������Ŀid; 
            --�����Ǵ�����Ŀʱ,��������Ŀ����ʱ,�Ѿ������˵�,���ԾͲ��ټ��� 
            If Nvl(n_������Ŀ, 0) = 1 Then 
              n_�����Ѵ��� := 1; 
            End If; 
            --��Ϊ���ܴ��ڶ��������Ŀ,���շ�ϸĿ��ͬ�����,���,�����ȼ�����Ŀ�Ƿ��Ѿ����������� 
            For I In 1 .. c_����.Count Loop 
              If c_����(I).�շ�ϸĿid = v_Currrow.�շ�ϸĿid And c_����(I).���� = Trunc(d_����ʱ��) Then 
                n_�����Ѵ��� := 1; 
                Exit; 
              End If; 
            End Loop; 
            If Nvl(n_�����Ѵ���, 0) = 0 Then 
              c_����.Extend; 
              c_����(c_����.Count).�շ�ϸĿid := v_Currrow.�շ�ϸĿid; 
              c_����(c_����.Count).���� := Trunc(d_����ʱ��); 
            End If; 
            If Nvl(n_������Ŀ, 0) = 0 And Nvl(n_�����Ѵ���, 0) = 0 Then 
              --3.������߼�λ 
              For v_��λ In (Select /*+ rule */ 
                            a.Column_Value As �շ�ϸĿid, p.�ּ�, p.������Ŀid 
                           From Table(l_����ȼ�) A, �շѼ�Ŀ P, �շ���ĿĿ¼ C 
                           Where a.Column_Value = p.�շ�ϸĿid And a.Column_Value = c.Id And d_����ʱ�� Between p.ִ������ And 
                                 Nvl(p.��ֹ����, Sysdate) And Nvl(c.���㷽ʽ, 0) <> 1 And p.�۸�ȼ� Is Null) Loop 
                If Nvl(v_��λ.�ּ�, 0) > n_�۸� Then 
                  n_�۸�       := Nvl(v_��λ.�ּ�, 0); 
                  n_Temp       := v_��λ.�շ�ϸĿid; 
                  n_������Ŀid := v_��λ.������Ŀid; 
                End If; 
              End Loop; 
 
              If n_Temp <> v_Currrow.�շ�ϸĿid And Nvl(n_�����Ѵ���, 0) = 0 Then 
 
                n_��������id := v_Currrow.����id; 
                n_���˲���id := v_Currrow.����id; 
 
                For c_�䶯��¼ In (Select ����id, ����id 
                               From ���˱䶯��¼ 
                               Where ��ʼԭ�� <> 10 And ����id = ����id_In And ��ҳid = ��ҳid_In And ����ȼ�id + 0 = n_Temp And 
                                     (Trunc(��ʼʱ��) = Trunc(d_����ʱ��) Or Trunc(Nvl(��ֹʱ��, Sysdate)) = Trunc(d_����ʱ��)) 
                               Order By ��ʼʱ�� Desc) Loop 
                  n_��������id := c_�䶯��¼.����id; 
                  n_���˲���id := c_�䶯��¼.����id; 
                  Exit; 
                End Loop;               
                      
                --4. ���ȵĻ�,��Ҫ���´�����ط��� 
                For v_���� In (Select n_Temp As �շ�ϸĿid, v_Currrow.���� As ����, n_�۸� As ����, n_������Ŀid As ������Ŀid 
                             From Dual 
                             Union All 
                             Select ����id As �շ�ϸĿid, a.�������� As ����, p.�ּ� As ����, p.������Ŀid 
                             From �շѴ�����Ŀ A, �շѼ�Ŀ P, �շ���ĿĿ¼ C 
                             Where a.����id = p.�շ�ϸĿid And a.����id = c.Id And Nvl(c.���㷽ʽ, 0) <> 1 And a.����id = n_Temp And 
                                   d_����ʱ�� Between p.ִ������ And Nvl(p.��ֹ����, Sysdate) And p.�۸�ȼ� Is Null) Loop 
                  --ȷ������ 
                  Begin 
                    Select ʵ�ձ��� 
                    Into n_Exsetax_Temp 
                    From (Select ʵ�ձ��� 
                           From �ѱ���ϸ 
                           Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_����.�շ�ϸĿid And 
                                 (Abs(v_����.���� * v_����.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) 
                           Union All 
                           Select ʵ�ձ��� 
                           From �ѱ���ϸ 
                           Where �ѱ� = v_Currrow.�ѱ� And ������Ŀid = v_����.������Ŀid And 
                                 (Abs(v_����.���� * v_����.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) And Not Exists 
                            (Select 1 From �ѱ���ϸ Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_����.�շ�ϸĿid)); 
                  Exception 
                    When Others Then 
                      n_Exsetax_Temp := 100.00; 
                  End; 
                  n_Exsetax_Temp := Nvl(n_Exsetax_Temp, 100); 
                  --����Ѿ����㣬ԭ��¼������ȫ��ȷ����ֱ���޸Ľ���־���� 
                  Update סԺ���ü�¼ 
                  Set ���ӱ�־ = 0 
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And Nvl(�Ӱ��־, 0) = v_Currrow.���Ӵ�λ And 
                        ���˿���id = Nvl(n_��������id, 0) And ���˲���id = Nvl(n_���˲���id, 0) And Nvl(����, 0) = Nvl(v_Currrow.����, 0) And 
                        �շ�ϸĿid = v_����.�շ�ϸĿid And ������Ŀid = v_����.������Ŀid And ����ʱ�� = d_����ʱ�� And ���� = v_����.���� * n_Dates And 
                        ��׼���� = v_����.���� And Ӧ�ս�� = Round(v_����.���� * v_����.���� * n_Dates, n_Dec) And 
                        ʵ�ս�� = Round(v_����.���� * v_����.���� * n_Dates * n_Exsetax / 100, n_Dec); 
 
                  If Sql%RowCount = 0 And n_Count=0 Then 
                    --���δ�������������������ȷ�ļ����¼ 
                    Insert Into סԺ���ü�¼ 
                      (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, 
                       ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, 
                       �վݷ�Ŀ, ������, ������, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id) 
                      Select ���˷��ü�¼_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                             Decode(v_Currrow.��ҳid, Null, 1, 2), v_Currrow.����id, v_Currrow.��ҳid, n_���˲���id, n_��������id, 
                             n_��������id, n_���˲���id, v_Currrow.����, v_Currrow.�Ա�, v_Currrow.����, v_Currrow.סԺ��, v_Currrow.����, 
                             v_Currrow.�ѱ�, 1, v_����.�շ�ϸĿid, v_����.������Ŀid, 0, v_����.����, 1, v_����.���� * n_Dates, 
                             Round(v_����.���� * v_����.���� * n_Dates, n_Dec), 
                             Round(v_����.���� * v_����.���� * n_Dates * n_Exsetax / 100, n_Dec), i.���, i.���㵥λ, v_Currrow.���Ӵ�λ, 
                             j.�վݷ�Ŀ, v_Currrow.����ҽʦ, v_Currrow.���λ�ʿ, v_Currrow.����Ա���, v_Currrow.����Ա����, d_����ʱ��, d_�Ǽ�ʱ��, 
                             Decode(v_Currrow.����, Null, 0, 1), v_Currrow.����id, 
                             Decode(v_Currrow.�㷨, 1, 
                                     Round(v_����.���� * v_����.���� * n_Dates * n_Exsetax / 100 * v_Currrow.ͳ��ȶ� / 100, n_Dec), 2, 
                                     v_Currrow.ͳ��ȶ�, 0), n_ҽ��С��id 
                      From (Select ���, ���㵥λ 
                             From �շ�ϸĿ 
                             Where ID = v_����.�շ�ϸĿid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) I, 
                           (Select �վݷ�Ŀ 
                             From ������Ŀ 
                             Where ID = v_����.������Ŀid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) J; 
                    n_Billcount := n_Billcount + Sql%RowCount; 
                  End If; 
                  n_�����Ѵ��� := 1; 
                End Loop; 
              End If; 
            End If; 
          End If; 
        End If; 
 
        If Nvl(n_�����Ѵ���, 0) = 0 Then 
          --����Ѿ����㣬ԭ��¼������ȫ��ȷ����ֱ���޸Ľ���־���� 
          Update סԺ���ü�¼ 
          Set ���ӱ�־ = 0 
          Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And Nvl(�Ӱ��־, 0) = v_Currrow.���Ӵ�λ And 
                ���˿���id = v_Currrow.����id And ���˲���id = Nvl(v_Currrow.����id, 0) And Nvl(����, 0) = Nvl(v_Currrow.����, 0) And 
                �շ�ϸĿid = v_Currrow.�շ�ϸĿid And ������Ŀid = v_Currrow.������Ŀid And ����ʱ�� = d_����ʱ�� And 
                ���� = v_Currrow.���� * n_Dates And ��׼���� = v_Currrow.��׼���� And 
                Ӧ�ս�� = Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates, n_Dec) And 
                ʵ�ս�� = Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates * n_Exsetax / 100, n_Dec); 
 
          If Sql%RowCount = 0 And n_Count=0 Then 
            --���δ�������������������ȷ�ļ����¼\ 
            Insert Into סԺ���ü�¼ 
              (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ����, �Ա�, 
               ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, ������, 
               ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id) 
              Select ���˷��ü�¼_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(v_Currrow.��ҳid, Null, 1, 2), v_Currrow.����id, v_Currrow.��ҳid, v_Currrow.����id, v_Currrow.����id, 
                     v_Currrow.����id, v_Currrow.����id, v_Currrow.����, v_Currrow.�Ա�, v_Currrow.����, v_Currrow.סԺ��, 
                     v_Currrow.����, v_Currrow.�ѱ�, 1, v_Currrow.�շ�ϸĿid, v_Currrow.������Ŀid, 0, v_Currrow.��׼����, 1, 
                     v_Currrow.���� * n_Dates, Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates, n_Dec), 
                     Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates * n_Exsetax / 100, n_Dec), i.���, i.���㵥λ, 
                     v_Currrow.���Ӵ�λ, j.�վݷ�Ŀ, v_Currrow.����ҽʦ, v_Currrow.���λ�ʿ, v_Currrow.����Ա���, v_Currrow.����Ա����, d_����ʱ��, 
                     d_�Ǽ�ʱ��, Decode(v_Currrow.����, Null, 0, 1), v_Currrow.����id, 
                     Decode(v_Currrow.�㷨, 1, 
                             Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates * n_Exsetax / 100 * v_Currrow.ͳ��ȶ� / 100, n_Dec), 
                             2, v_Currrow.ͳ��ȶ�, 0), n_ҽ��С��id 
              From (Select ���, ���㵥λ 
                     From �շ�ϸĿ 
                     Where ID = v_Currrow.�շ�ϸĿid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) I, 
                   (Select �վݷ�Ŀ 
                     From ������Ŀ 
                     Where ID = v_Currrow.������Ŀid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) J; 
 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End If; 
      End Loop; 
    End Loop; 
  Else 
    For v_Currrow In v_Autocur(�ڼ�_In, n_Insure) Loop 
 
      If v_Currrow.ҽ��С��id Is Null Then 
        n_ҽ��С��id := Zl_ҽ��С��_Get(v_Currrow.����id, v_Currrow.����Ա����, v_Currrow.����id, v_Currrow.��ҳid, d_����ʱ��); 
      Else 
        n_ҽ��С��id := v_Currrow.ҽ��С��id; 
      End If; 
 
      If d_Datefrom > v_Currrow.��ʼ���� Then 
        d_Datefrom := v_Currrow.��ʼ����; 
        n_Do       := 1; 
        --�����ο�ʼ����ʱ���Ժ���Ѽ����¼��־�޸� 
        Update סԺ���ü�¼ 
        Set ���ӱ�־ = 5 
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ <> 8 And Nvl(ҽ�����, 0) = 0 And 
              ����ʱ�� >= v_Currrow.��ʼ����; 
      End If; 
 
      If d_Dateto < v_Currrow.��ֹ���� Then 
        d_Dateto := v_Currrow.��ֹ����; 
      End If; 
      n_�շ�ϸĿid := v_Currrow.�շ�ϸĿid; 
      n_������Ŀ   := 0; 
      --����Ѽ����׼:0-�����һ�λ������;1-���۸���ߵĻ���ȼ����㡣 
      If n_��������׼ = 1 Then 
        --��ȷ���Ƿ�����Ŀ,�����,����Ҫ���½��м��� 
        Select Count(*) Into n_������Ŀ From Table(l_����id) Where Column_Value = n_�շ�ϸĿid; 
      End If; 
 
      --��ȡ��ǰ������Ŀ���շѱ��� 
      Begin 
        Select ʵ�ձ��� 
        Into n_Exsetax 
        From (Select ʵ�ձ��� 
               From �ѱ���ϸ 
               Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_Currrow.�շ�ϸĿid And 
                     (Abs(v_Currrow.��׼���� * v_Currrow.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) 
               Union All 
               Select ʵ�ձ��� 
               From �ѱ���ϸ 
               Where �ѱ� = v_Currrow.�ѱ� And ������Ŀid = v_Currrow.������Ŀid And 
                     (Abs(v_Currrow.��׼���� * v_Currrow.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) And Not Exists 
                (Select 1 From �ѱ���ϸ Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_Currrow.�շ�ϸĿid)); 
      Exception 
        When Others Then 
          n_Exsetax := 100.00; 
      End; 
 
      n_Exsetax := Nvl(n_Exsetax, 100); 
      For n_Datecount In 0 .. (Trunc(v_Currrow.��ֹ���� + 0.5) - Trunc(v_Currrow.��ʼ����)) - 1 Loop 
        d_����ʱ�� := Greatest(v_Currrow.��ʼ����, Trunc(v_Currrow.��ʼ���� + n_Datecount)); 
        n_Dates    := Least(Trunc(v_Currrow.��ʼ���� + n_Datecount + 1), v_Currrow.��ֹ����) - 
                      Greatest(v_Currrow.��ʼ����, Trunc(v_Currrow.��ʼ���� + n_Datecount)); 
      
      --�ж��Ƿ��ֹ�����
      Select Count(1)
      Into n_Count
      From סԺ���ü�¼
      Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 2 And Nvl(���ӱ�־, 0) = 1 And
            �շ���� = Decode(v_Currrow.����, 1, 'H', 2, 'J', �շ����) And ����ʱ�� = d_����ʱ�� And
            �շ�ϸĿid = Decode(v_Currrow.����, 3, v_Currrow.�շ�ϸĿid, �շ�ϸĿid); 

        n_�����Ѵ��� := 0; 
        If n_������Ŀ = 1 Then 
          --1.�ȼ�鵱���Ƿ���ڻ���䶯,ֻ�д��ڶ������䶯��,�Ż�ȥ����(������ĿΪ׼) 
          n_������Ŀ := 1; 
          If l_����ȼ�.Count > 0 Then 
            l_����ȼ�.Delete; 
          End If; 
          For v_���� In (Select Distinct ����ȼ�id 
                       From ���˱䶯��¼ 
                       Where ��ʼԭ�� <> 10 And ����id = ����id_In And ��ҳid = ��ҳid_In And 
                             (Trunc(��ʼʱ��) = Trunc(d_����ʱ��) Or Trunc(Nvl(��ֹʱ��, Sysdate)) = Trunc(d_����ʱ��))) Loop 
            If Nvl(v_����.����ȼ�id, 0) <> 0 Then 
              l_����ȼ�.Extend; 
              l_����ȼ�(l_����ȼ�.Count) := v_����.����ȼ�id; 
              If Nvl(v_����.����ȼ�id, 0) = Nvl(v_Currrow.�շ�ϸĿid, 0) Then 
                n_������Ŀ := 0; 
              End If; 
            End If; 
          End Loop; 
          If l_����ȼ�.Count > 1 Then 
            --2. �����������ϱ䶯,��ȡ��λ��ߵ� 
            n_Temp       := v_Currrow.�շ�ϸĿid; 
            n_�۸�       := Nvl(v_Currrow.��׼����, 0); 
            n_������Ŀid := v_Currrow.������Ŀid; 
            --�����Ǵ�����Ŀʱ,��������Ŀ����ʱ,�Ѿ������˵�,���ԾͲ��ټ��� 
            If Nvl(n_������Ŀ, 0) = 1 Then 
              n_�����Ѵ��� := 1; 
            End If; 
            --��Ϊ���ܴ��ڶ��������Ŀ,���շ�ϸĿ��ͬ�����,���,�����ȼ�����Ŀ�Ƿ��Ѿ����������� 
            For I In 1 .. c_����.Count Loop 
              If c_����(I).�շ�ϸĿid = v_Currrow.�շ�ϸĿid And c_����(I).���� = Trunc(d_����ʱ��) Then 
                n_�����Ѵ��� := 1; 
                Exit; 
              End If; 
            End Loop; 
            If Nvl(n_�����Ѵ���, 0) = 0 Then 
              c_����.Extend; 
              c_����(c_����.Count).�շ�ϸĿid := v_Currrow.�շ�ϸĿid; 
              c_����(c_����.Count).���� := Trunc(d_����ʱ��); 
            End If; 
            If Nvl(n_������Ŀ, 0) = 0 And Nvl(n_�����Ѵ���, 0) = 0 Then 
              --3.������߼�λ 
              For v_��λ In (Select /*+ rule */ 
                            a.Column_Value As �շ�ϸĿid, p.�ּ�, p.������Ŀid 
                           From Table(l_����ȼ�) A, �շѼ�Ŀ P, �շ���ĿĿ¼ C 
                           Where a.Column_Value = p.�շ�ϸĿid And a.Column_Value = c.Id And d_����ʱ�� Between p.ִ������ And 
                                 Nvl(p.��ֹ����, Sysdate) And Nvl(c.���㷽ʽ, 0) <> 1 And p.�۸�ȼ� Is Null) Loop 
                If Nvl(v_��λ.�ּ�, 0) > n_�۸� Then 
                  n_�۸�       := Nvl(v_��λ.�ּ�, 0); 
                  n_Temp       := v_��λ.�շ�ϸĿid; 
                  n_������Ŀid := v_��λ.������Ŀid; 
                End If; 
              End Loop; 
 
              If n_Temp <> v_Currrow.�շ�ϸĿid And Nvl(n_�����Ѵ���, 0) = 0 Then 
 
                n_��������id := v_Currrow.����id; 
                n_���˲���id := v_Currrow.����id; 
 
                For c_�䶯��¼ In (Select ����id, ����id 
                               From ���˱䶯��¼ 
                               Where ��ʼԭ�� <> 10 And ����id = ����id_In And ��ҳid = ��ҳid_In And ����ȼ�id + 0 = n_Temp And 
                                     (Trunc(��ʼʱ��) = Trunc(d_����ʱ��) Or Trunc(Nvl(��ֹʱ��, Sysdate)) = Trunc(d_����ʱ��)) 
                               Order By ��ʼʱ�� Desc) Loop 
                  n_��������id := c_�䶯��¼.����id; 
                  n_���˲���id := c_�䶯��¼.����id; 
                  Exit; 
                End Loop; 
                      
                --4. ���ȵĻ�,��Ҫ���´�����ط��� 
                For v_���� In (Select n_Temp As �շ�ϸĿid, v_Currrow.���� As ����, n_�۸� As ����, n_������Ŀid As ������Ŀid 
                             From Dual 
                             Union All 
                             Select ����id As �շ�ϸĿid, a.�������� As ����, p.�ּ� As ����, p.������Ŀid 
                             From �շѴ�����Ŀ A, �շѼ�Ŀ P, �շ���ĿĿ¼ C 
                             Where a.����id = p.�շ�ϸĿid And a.����id = c.Id And Nvl(c.���㷽ʽ, 0) <> 1 And a.����id = n_Temp And 
                                   d_����ʱ�� Between p.ִ������ And Nvl(p.��ֹ����, Sysdate) And p.�۸�ȼ� Is Null) Loop 
                  --ȷ������ 
                  Begin 
                    Select ʵ�ձ��� 
                    Into n_Exsetax_Temp 
                    From (Select ʵ�ձ��� 
                           From �ѱ���ϸ 
                           Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_����.�շ�ϸĿid And 
                                 (Abs(v_����.���� * v_����.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) 
                           Union All 
                           Select ʵ�ձ��� 
                           From �ѱ���ϸ 
                           Where �ѱ� = v_Currrow.�ѱ� And ������Ŀid = v_����.������Ŀid And 
                                 (Abs(v_����.���� * v_����.����) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) And Not Exists 
                            (Select 1 From �ѱ���ϸ Where �ѱ� = v_Currrow.�ѱ� And �շ�ϸĿid = v_����.�շ�ϸĿid)); 
                  Exception 
                    When Others Then 
                      n_Exsetax_Temp := 100.00; 
                  End; 
                  n_Exsetax_Temp := Nvl(n_Exsetax_Temp, 100); 
                  --����Ѿ����㣬ԭ��¼������ȫ��ȷ����ֱ���޸Ľ���־���� 
                  Update סԺ���ü�¼ 
                  Set ���ӱ�־ = 0 
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And Nvl(�Ӱ��־, 0) = v_Currrow.���Ӵ�λ And 
                        ���˿���id = Nvl(n_��������id, 0) And ���˲���id = Nvl(n_���˲���id, 0) And Nvl(����, 0) = Nvl(v_Currrow.����, 0) And 
                        �շ�ϸĿid = v_����.�շ�ϸĿid And ������Ŀid = v_����.������Ŀid And ����ʱ�� = d_����ʱ�� And ���� = v_����.���� * n_Dates And 
                        ��׼���� = v_����.���� And Ӧ�ս�� = Round(v_����.���� * v_����.���� * n_Dates, n_Dec) And 
                        ʵ�ս�� = Round(v_����.���� * v_����.���� * n_Dates * n_Exsetax / 100, n_Dec); 
 
                  If Sql%RowCount = 0 And n_Count=0Then 
                    --���δ�������������������ȷ�ļ����¼ 
                    Insert Into סԺ���ü�¼ 
                      (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, 
                       ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, 
                       �վݷ�Ŀ, ������, ������, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id) 
                      Select ���˷��ü�¼_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                             Decode(v_Currrow.��ҳid, Null, 1, 2), v_Currrow.����id, v_Currrow.��ҳid, n_���˲���id, n_��������id, 
                             n_��������id, n_���˲���id, v_Currrow.����, v_Currrow.�Ա�, v_Currrow.����, v_Currrow.סԺ��, v_Currrow.����, 
                             v_Currrow.�ѱ�, 1, v_����.�շ�ϸĿid, v_����.������Ŀid, 0, v_����.����, 1, v_����.���� * n_Dates, 
                             Round(v_����.���� * v_����.���� * n_Dates, n_Dec), 
                             Round(v_����.���� * v_����.���� * n_Dates * n_Exsetax / 100, n_Dec), i.���, i.���㵥λ, v_Currrow.���Ӵ�λ, 
                             j.�վݷ�Ŀ, v_Currrow.����ҽʦ, v_Currrow.���λ�ʿ, v_Currrow.����Ա���, v_Currrow.����Ա����, d_����ʱ��, d_�Ǽ�ʱ��, 
                             Decode(v_Currrow.����, Null, 0, 1), v_Currrow.����id, 
                             Decode(v_Currrow.�㷨, 1, 
                                     Round(v_����.���� * v_����.���� * n_Dates * n_Exsetax / 100 * v_Currrow.ͳ��ȶ� / 100, n_Dec), 2, 
                                     v_Currrow.ͳ��ȶ�, 0), n_ҽ��С��id 
                      From (Select ���, ���㵥λ 
                             From �շ�ϸĿ 
                             Where ID = v_����.�շ�ϸĿid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) I, 
                           (Select �վݷ�Ŀ 
                             From ������Ŀ 
                             Where ID = v_����.������Ŀid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) J; 
                    n_Billcount := n_Billcount + Sql%RowCount; 
                  End If; 
                  n_�����Ѵ��� := 1; 
                End Loop; 
              End If; 
            End If; 
          End If; 
        End If; 
 
        If Nvl(n_�����Ѵ���, 0) = 0 Then 
          --����Ѿ����㣬ԭ��¼������ȫ��ȷ����ֱ���޸Ľ���־���� 
          Update סԺ���ü�¼ 
          Set ���ӱ�־ = 0 
          Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And Nvl(�Ӱ��־, 0) = v_Currrow.���Ӵ�λ And 
                ���˿���id = v_Currrow.����id And ���˲���id = Nvl(v_Currrow.����id, 0) And Nvl(����, 0) = Nvl(v_Currrow.����, 0) And 
                �շ�ϸĿid = v_Currrow.�շ�ϸĿid And ������Ŀid = v_Currrow.������Ŀid And ����ʱ�� = d_����ʱ�� And 
                ���� = v_Currrow.���� * n_Dates And ��׼���� = v_Currrow.��׼���� And 
                Ӧ�ս�� = Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates, n_Dec) And 
                ʵ�ս�� = Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates * n_Exsetax / 100, n_Dec); 
 
          If Sql%RowCount = 0 And n_Count=0 Then 
            --���δ�������������������ȷ�ļ����¼\ 
            Insert Into סԺ���ü�¼ 
              (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ����, �Ա�, 
               ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, ������, 
               ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id) 
              Select ���˷��ü�¼_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(v_Currrow.��ҳid, Null, 1, 2), v_Currrow.����id, v_Currrow.��ҳid, v_Currrow.����id, v_Currrow.����id, 
                     v_Currrow.����id, v_Currrow.����id, v_Currrow.����, v_Currrow.�Ա�, v_Currrow.����, v_Currrow.סԺ��, 
                     v_Currrow.����, v_Currrow.�ѱ�, 1, v_Currrow.�շ�ϸĿid, v_Currrow.������Ŀid, 0, v_Currrow.��׼����, 1, 
                     v_Currrow.���� * n_Dates, Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates, n_Dec), 
                     Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates * n_Exsetax / 100, n_Dec), i.���, i.���㵥λ, 
                     v_Currrow.���Ӵ�λ, j.�վݷ�Ŀ, v_Currrow.����ҽʦ, v_Currrow.���λ�ʿ, v_Currrow.����Ա���, v_Currrow.����Ա����, d_����ʱ��, 
                     d_�Ǽ�ʱ��, Decode(v_Currrow.����, Null, 0, 1), v_Currrow.����id, 
                     Decode(v_Currrow.�㷨, 1, 
                             Round(v_Currrow.��׼���� * v_Currrow.���� * n_Dates * n_Exsetax / 100 * v_Currrow.ͳ��ȶ� / 100, n_Dec), 
                             2, v_Currrow.ͳ��ȶ�, 0), n_ҽ��С��id 
              From (Select ���, ���㵥λ 
                     From �շ�ϸĿ 
                     Where ID = v_Currrow.�շ�ϸĿid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) I, 
                   (Select �վݷ�Ŀ 
                     From ������Ŀ 
                     Where ID = v_Currrow.������Ŀid And (����ʱ�� Is Null Or ����ʱ�� > d_����ʱ��)) J; 
 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End If; 
      End Loop; 
    End Loop; 
  End If; 
  If n_Do = 0 Then 
    --������Ժ��,����޸ĳ�Ժʱ��Ϊ��Ժ�����򲻲����·���,����ǰ�ķ���Ҫ���� 
    Begin 
      Select Nvl(Trunc(b.�ϴμ���ʱ��), Trunc(b.��ֹʱ��)) 
      Into d_Datelast 
      From ���˱䶯��¼ A, ���˱䶯��¼ B 
      Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And a.��ֹԭ�� = 1 And a.����id = b.����id And a.��ҳid = b.��ҳid And b.��ʼԭ�� = 1 And 
            Trunc(b.��ʼʱ��) = Trunc(a.��ֹʱ��) And a.���Ӵ�λ = 0 And b.���Ӵ�λ = 0; 
    Exception 
      When Others Then 
        Null; 
    End; 
    If d_Datelast Is Not Null Then 
      d_Datefrom := d_Datelast; 
      d_Dateto   := Sysdate; 
      Update סԺ���ü�¼ 
      Set ���ӱ�־ = 5 
      Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ <> 8 And Nvl(ҽ�����, 0) = 0 And 
            ����ʱ�� >= d_Datefrom; 
    End If; 
  End If; 
 
  ----------------------------------------------------------------- 
  --������ǰ����Ĵ����¼ 
  ----------------------------------------------------------------- 
  Insert Into סԺ���ü�¼ 
    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ����, �Ա�, ����, ��ʶ��, 
     ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, ������, ����Ա���, ����Ա����, ����ʱ��, 
     �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id) 
    Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, 2, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, 
           ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, 0, ��׼����, ����, -����, -Ӧ�ս��, -ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, 
           ������, ����Ա���, ����Ա����, ����ʱ��, d_�Ǽ�ʱ��, ������Ŀ��, ���մ���id, -ͳ����, ҽ��С��id 
    From סԺ���ü�¼ 
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 5 And ����ʱ�� >= d_Datefrom; 
 
  ----------------------------------------------------------------- 
  --��д������� 
  ----------------------------------------------------------------- 
  Select Sum(Decode(���ӱ�־, 0, 1, -1) * ʵ�ս��) As ʵ�ս�� 
  Into n_Summoney 
  From סԺ���ü�¼ 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And 
        (NO = v_Billno Or ���ӱ�־ = 5 And ����ʱ�� >= d_Datefrom); 
 
  Update ������� 
  Set ������� = Nvl(�������, 0) + Nvl(n_Summoney, 0) 
  Where ����id = ����id_In And ���� = 1 And ���� = 2 
  Returning ������� Into n_����ֵ; 
 
  If Sql%RowCount = 0 Then 
    Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 2, n_Summoney, 0); 
    n_����ֵ := n_Summoney; 
  End If; 
 
  If Nvl(n_����ֵ, 0) = 0 Then 
    Delete From ������� Where ���� = 1 And ����id = ����id_In And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --��д���˻��ܷ��� 
  ----------------------------------------------------------------- 
  n_Delete := 0; 
  For v_Currrow In v_Sumcur(v_Billno, d_Datefrom) Loop 
    Update ����δ����� 
    Set ��� = Nvl(���, 0) + Nvl(v_Currrow.ʵ�ս��, 0) 
    Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(v_Currrow.���˲���id, 0) And 
          Nvl(���˿���id, 0) = Nvl(v_Currrow.���˿���id, 0) And Nvl(��������id, 0) = Nvl(v_Currrow.��������id, 0) And 
          Nvl(ִ�в���id, 0) = Nvl(v_Currrow.ִ�в���id, 0) And Nvl(������Ŀid, 0) = Nvl(v_Currrow.������Ŀid, 0) And ��Դ;�� + 0 = 2 
    Returning ��� Into n_����ֵ; 
 
    If Sql%RowCount = 0 Then 
      Insert Into ����δ����� 
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���) 
      Values 
        (����id_In, ��ҳid_In, v_Currrow.���˲���id, v_Currrow.���˿���id, v_Currrow.��������id, v_Currrow.ִ�в���id, v_Currrow.������Ŀid, 2, 
         v_Currrow.ʵ�ս��); 
      n_����ֵ := v_Currrow.ʵ�ս��; 
    End If; 
    If Nvl(n_����ֵ, 0) = 0 Then 
      n_Delete := 1; 
    End If; 
  End Loop; 
 
  If Nvl(n_Delete, 0) = 1 Then 
    Delete From ����δ����� Where ����id = ����id_In And ��� = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --�������޸ĵĸ��ӱ�־��ԭΪ������־ 
  ----------------------------------------------------------------- 
  Update סԺ���ü�¼ 
  Set ���ӱ�־ = 0, ��¼״̬ = 3 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 5 And ����ʱ�� >= d_Datefrom; 
 
  ----------------------------------------------------------------- 
  --�޸ļ���ʱ���־ 
  ----------------------------------------------------------------- 
  Update ���˱䶯��¼ 
  Set �ϴμ���ʱ�� = Least(d_Dateto, Nvl(��ֹʱ��, Greatest(��ʼʱ��, Sysdate))) 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(��ֹʱ��, Sysdate) > d_Datefrom; 
  Commit; --���������ύ 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl1_Autocptone;
/

--123419:����,2018-05-10,�����ֹ����Զ����ʷ��ø��ӱ�־����
CREATE OR REPLACE Procedure Zl1_Autocalc_Pati_Charge_Nm
( 
  ����id_In       In ������ҳ.����id%Type, 
  ��ҳid_In       In ������ҳ.��ҳid%Type, 
  �ڼ�_In         In �ڼ��.�ڼ�%Type, 
  ǿ�Ƽ���_In     In Number := 0, 
  ���ü۸�ȼ�_In In Number := -1 
) As 
  ------------------------------------------------------------------------- 
  --����˵�������ָ������ָ���ڼ���Զ�����(��Ҫ���������Ƭ�����Զ�������Ŀ�ļ���) 
  --          1��ϵͳ���ȸ���ϵͳ����"���������Զ��Ʒ�"���޸������ò����Զ����ʼ�¼��־; 
  --          2���ۺϲ��˵Ĵ�λ�仯�����ת�������������ȶ������أ�����ڼ��ȡ����˷� 
  --             �����ɷ��õ���ȷ���㣺 
  --             ��������Ѿ����㣬���޸ı�־Ϊ����;���δ���㣬������µ��Զ����ʼ�¼; 
  --             ������ǰ�Ĵ������ļ�¼; 
  --             ͳ�Ʊ��α䶯(����������)����д����ͻ��ܱ�; 
  --��ڲ����� 
  --       ����ID_IN  number    �������ID 
  --       ��ҳID_IN  number    ������ҳID������������ͬȷ����Ҫ����Ĳ��� 
  --       �ڼ�_IN  varchar2     ��Ҫ�������С�ڼ� 
  --       ǿ�Ƽ���_IN number   Ϊ1ʱ,���ܲ�����ҳ.��ֹ�Զ��������Կ��� 
  --       ���ü۸�ȼ�_In number ��-1��ʾδ�жϼ۸�ȼ�,�ڲ����Զ�ȥ���;0-�����ü۸�ȼ�;1-�����˼۸�ȼ��� 
  --���ù�ϵ��zl1_AutoCptPati/zl1_AutoCptWard/zl1_AutoCptAll ���ñ����� 
  --�Զ����ʹ���˵��: 
  --   1. ��λ:  ���벻�Ƴ�, ������;������(ת�ƣ�ת�������ȼ��䶯��),12����ǰ����ת�����Ϊ׼;12���Ժ���ת������Ϊ׼ 
  --   2.������������:  ��Ժ���찴һ�����,��Ժ��������12��֮ǰ����죬12��֮����һ�� 
  ---------------------------------------------------------------------------------------------------------------------------------- 
  v_�۸�ȼ�         �շѼ۸�ȼ�.����%Type; 
  v_���ʽ�۸�ȼ� �շѼ۸�ȼ�.����%Type; 
 
  v_Temp      Varchar2(500); 
  v_Billno    Varchar2(8); --���ñ�ʵ�ʵ��Զ����ʺ��� 
  n_Billcount Number(5) := 0; --������ż����� 
 
  n_Exsetax  Number(16, 2) := 0; --������ȡ���� 
  n_Summoney Number(16, 2) := 0; --��� 
 
  n_Dec    Number; --���С��λ�� 
  n_Dates  Number(4, 1); --��ǰ��¼��������ȫ��Ϊ1 
  n_Delete Number; 
  n_Exists Number; 
  n_Count  Number(5);
  n_����ֵ �������.Ԥ�����%Type; 
 
  v_�վݷ�Ŀ   ������Ŀ.�վݷ�Ŀ%Type; 
  v_���㵥λ   �շ���ĿĿ¼.���㵥λ%Type; 
  n_סԺ״̬   ������ҳ.״̬%Type; 
  n_��׼�۸�   �շѼ�Ŀ.�ּ�%Type; 
  n_������Ŀid ������Ŀ.Id%Type; 
  v_���       �շ���ĿĿ¼.���%Type; 
  n_�㷨       ����֧������.�㷨%Type; 
  n_ͳ��ȶ�   ����֧������.ͳ��ȶ�%Type; 
  n_�������   �����Զ�����.����%Type; 
 
  n_������˷�ʽ   Number(2); 
  n_δ��ƽ�ֹ���� Number(2); 
  n_����۸�����   Number(2); 
  n_�Ƿ��ü۸�ȼ� Number(2); 
  n_�Ƿ�������   Number(2); 
  n_Finded         Number(2); 
  n_����           Number(2); --1-����;2- ��λ;3-���� 
  n_Find           Number(2); 
  n_Last           Number(2); 
  n_ǰ����id       �����Զ�����.����id%Type; 
  n_ǰ����id       �����Զ�����.����id%Type; 
  n_ǰ�շ�ϸĿid   �����Զ�����.����ȼ�id%Type; 
  v_ǰ����         �����Զ�����.����%Type; 
  n_ǰ��λ�ȼ�id   �����Զ�����.��λ�ȼ�id%Type; 
  v_����վ��       ���ű�.վ��%Type; 
 
  d_Start_Date Date; 
  d_�Ǽ�ʱ��   Date; --�Ǽ�ʱ�� 
  d_����ʱ��   Date; --����ʱ�� 
  d_Temp       Date; 
 
  d_��λʱ��_Max Date; 
  d_����ʱ��_Max Date; 
  d_����ʱ��_Max Date; 
 
  l_Mulit_ϸĿid t_Numlist := t_Numlist(); 
 
  Type t_�۸�_Rec Is Ref Cursor; 
  c_�۸�_Rec t_�۸�_Rec; 
 
  Type t_���˱䶯_Rec Is Record( 
    ID         ���˱䶯��¼.Id%Type, 
    ��ʼʱ��   ���˱䶯��¼.��ʼʱ��%Type, 
    ��ֹʱ��   ���˱䶯��¼.��ʼʱ��%Type, 
    ����id     ���˱䶯��¼.����id%Type, 
    ����id     ���˱䶯��¼.����id%Type, 
    ����ҽʦ   ���˱䶯��¼.����ҽʦ%Type, 
    ���λ�ʿ   ���˱䶯��¼.���λ�ʿ%Type, 
    ҽ��С��id ���˱䶯��¼.ҽ��С��id%Type); 
 
  Type c_���˱䶯_Rec Is Table Of t_���˱䶯_Rec; 
  r_���˱䶯 c_���˱䶯_Rec := c_���˱䶯_Rec(); 
  r_�䶯_Cur c_���˱䶯_Rec := c_���˱䶯_Rec(); 
 
  Cursor c_Sumcur_Rec 
  ( 
    Billno    Varchar2, 
    Datestart Date 
  ) Is 
    Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Decode(���ӱ�־, 0, 1, -1) * Ӧ�ս��) As Ӧ�ս��, 
           Sum(Decode(���ӱ�־, 0, 1, -1) * ʵ�ս��) As ʵ�ս�� 
    From סԺ���ü�¼ 
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And 
          (NO = Billno Or ���ӱ�־ = 5 And ����ʱ�� >= Datestart) 
    Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid; 
 
  Cursor c_Pati Is 
    Select a.����id, a.��ҳid, Nvl(a.����, i.����) As ����, Nvl(a.�Ա�, i.�Ա�) As �Ա�, Nvl(a.����, i.����) As ����, Nvl(a.סԺ��, i.סԺ��) As סԺ��, 
           a.�ѱ�, Nvl(a.����, 0) As ����, Nvl(a.��˱�־, 0) As ��˱�־, Nvl(a.״̬, 0) As סԺ״̬, Nvl(a.�Ƿ��ֹ�Զ�����, 0) As �Ƿ��ֹ�Զ�����, 
           a.ҽ�Ƹ��ʽ As ���ʽ, a.��Ժ����, a.��Ժ���� 
    From ������ҳ A, ������Ϣ I 
    Where a.����id = i.����id And a.����id = ����id_In And a.��ҳid = ��ҳid_In; 
 
  r_Pati c_Pati%RowType; 
  Cursor c_Pati_Change 
  ( 
    ����id_In In ������ҳ.����id%Type, 
    ��ҳid_In In ������ҳ.��ҳid%Type, 
    ����_In   In ������ҳ.����%Type, 
    �ڼ�_In   In Varchar2 
  ) Is 
    Select a.����, a.Id, a.����id, a.��ҳid, a.����id, a.����id, a.����, a.���Ӵ�λ, a.�շ�ϸĿid, a.����Ա���, a.����Ա����, a.��ʼʱ��, a.��ֹʱ��, a.��������, 
           a.����, Greatest(a.��ʼ����, Trunc(p.��ʼ����)) As ��ʼ����, a.��ֹ����, a.����, Nvl(Q1.վ��, Q2.վ��) As վ��, m.���㵥λ, m.���, i.����, 
           i.����id, k.�㷨, k.ͳ��ȶ�, Nvl(m.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��Ŀ����ʱ��, a.�����־ 
    From (Select 2 As ����, b.Id, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, b.��λ�ȼ�id As �շ�ϸĿid, b.����Ա���, b.����Ա����, 
                  b.��ʼʱ��, b.��ֹʱ��, a.��������, b.����, Trunc(b.��ʼʱ��) As ��ʼ����, Trunc(Nvl(b.��ֹʱ��, Sysdate)) As ��ֹ����, 
                  Trunc(Nvl(b.��ֹʱ��, Sysdate)) - Trunc(b.��ʼʱ��) As ����, 0 As �����־ 
           From �Զ��Ƽ���Ŀ A, 
                (Select a.Id, a.����id, a.��ҳid, a.��ʼʱ��, a.���Ӵ�λ, a.����id, a.����id, a.����, a.��λ�ȼ�id, 1 As ����, a.��ֹʱ��, a.����Ա���, 
                         a.����Ա����, a.�ϴμ���ʱ�� 
                  From �����Զ����� A 
                  Where a.���� = 2 And a.����id = ����id_In And a.��ҳid = ��ҳid_In And 
                        Nvl(a.�ϴμ���ʱ��, a.��ʼʱ��) <= Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) 
                  Union All 
                  Select b.Id, b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, b.��ֹʱ��, 
                         b.����Ա���, b.����Ա����, b.�ϴμ���ʱ�� 
                  From �����Զ����� B, �շѴ�����Ŀ I 
                  Where b.���� = 2 And b.����id = ����id_In And b.��ҳid = ��ҳid_In And b.��λ�ȼ�id = i.����id And i.���д��� > 0 And 
                        Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��) <= Nvl(b.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))) B 
           Where a.����id = b.����id And a.�����־ = 1 And Trunc(Nvl(b.��ֹʱ��, Sysdate)) >= a.�������� 
           Union All 
           Select 1 As ����, b.Id, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, b.����ȼ�id As �շ�ϸĿid, b.����Ա���, b.����Ա����, 
                  b.��ʼʱ��, b.��ֹʱ��, a.��������, b.����, Trunc(b.��ʼʱ��) As ��ʼ����, 
                  Decode(Trunc(b.��ֹʱ��), Trunc(b.��ʼʱ��), Trunc(Nvl(b.��ֹʱ��, Sysdate)), 
                          Zl_Date_Half(Nvl(b.��ֹʱ��, Trunc(Sysdate)), 1)) As ��ֹ����, 
                  Decode(Trunc(b.��ֹʱ��), Trunc(b.��ʼʱ��), Trunc(Nvl(b.��ֹʱ��, Sysdate)), 
                          Zl_Date_Half(Nvl(b.��ֹʱ��, Trunc(Sysdate)), 1)) - Trunc(b.��ʼʱ��) As ����, 0 As �����־ 
           From �Զ��Ƽ���Ŀ A, 
                (Select a.Id, a.����id, a.��ҳid, a.��ʼʱ��, a.���Ӵ�λ, a.����id, a.����id, a.����, a.����ȼ�id, 1 As ����, a.��ֹʱ��, a.����Ա���, 
                         a.����Ա����, a.�ϴμ���ʱ�� 
                  From �����Զ����� A 
                  Where a.���� = 1 And a.����id = ����id_In And a.��ҳid = ��ҳid_In And 
                        Nvl(a.�ϴμ���ʱ��, a.��ʼʱ��) <= Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) 
                  Union All 
                  Select b.Id, b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, b.��ֹʱ��, 
                         b.����Ա���, b.����Ա����, b.�ϴμ���ʱ�� 
                  From �����Զ����� B, �շѴ�����Ŀ I 
                  Where b.���� = 1 And b.����id = ����id_In And b.��ҳid = ��ҳid_In And b.����ȼ�id = i.����id And i.���д��� > 0 And 
                        Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��) <= Nvl(b.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))) B 
           Where a.����id = b.����id And a.�����־ = 2 And 
                 Decode(Trunc(b.��ֹʱ��), Trunc(b.��ʼʱ��), Trunc(Nvl(b.��ֹʱ��, Sysdate)), 
                        Zl_Date_Half(Nvl(b.��ֹʱ��, Trunc(Sysdate)), 1)) >= a.�������� 
           Union All 
           Select 3 As ����, b.Id, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, a.�շ�ϸĿid, b.����Ա���, b.����Ա����, b.��ʼʱ��, b.��ֹʱ��, 
                  a.��������, a.����, Trunc(b.��ʼʱ��) As ��ʼ����, Zl_Date_Half(Nvl(b.��ֹʱ��, Trunc(Sysdate)), 1) As ��ֹ����, 
                  Trunc(Nvl(b.��ֹʱ��, Sysdate)) - Trunc(b.��ʼʱ��) As ����, a.�����־ 
           From (Select ����id, �����־, �շ�ϸĿid, 1 As ����, �������� 
                  From �Զ��Ƽ���Ŀ 
                  Union All 
                  Select ����id, �����־, ����id, i.�������� As ����, �������� 
                  From �Զ��Ƽ���Ŀ A, �շѴ�����Ŀ I 
                  Where a.�շ�ϸĿid = i.����id And i.���д��� > 0) A, �����Զ����� B 
           Where a.����id = b.����id And b.����id = ����id_In And Zl_Date_Half(Nvl(b.��ֹʱ��, Trunc(Sysdate)), 1) >= a.�������� And 
                 b.��ҳid = ��ҳid_In And b.���� = 3 And Nvl(b.���Ӵ�λ, 0) = 0 And 
                 (a.�����־ = 6 And b.��λ�ȼ�id Is Not Null Or a.�����־=7) And 
                 Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��) <= Nvl(b.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))) A, 
         (Select Min(��ʼ����) As ��ʼ���� From �ڼ�� Where �ڼ� >= �ڼ�_In) P, ����֧����Ŀ I, ����֧������ K, �շ���ĿĿ¼ M, ���ű� Q1, ���ű� Q2 
    Where Trunc(Nvl(a.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD'))) >= Trunc(p.��ʼ����) And a.�շ�ϸĿid = i.�շ�ϸĿid(+) And 
          i.����(+) = Nvl(����_In, 0) And i.����id = k.Id(+) And a.�շ�ϸĿid = m.Id(+) And a.����id = Q1.Id(+) And 
          a.����id = Q2.Id(+) 
    Order By ����, ���Ӵ�λ, ��ʼʱ��; 
 
  r_Pati_Change     c_Pati_Change%RowType; 
  r_Pati_Change_Pre c_Pati_Change%RowType; 
 
  Function Get_Discount_Rate 
  ( 
    �ѱ�_In       ������Ϣ.�ѱ�%Type, 
    �շ�ϸĿid_In �ѱ���ϸ.�շ�ϸĿid%Type, 
    ������Ŀid_In �ѱ���ϸ.������Ŀid%Type, 
    ���_In       �ѱ���ϸ.Ӧ�ն���ֵ%Type 
  ) Return Number As 
    n_Discount_Rate Number(16, 5); 
  Begin 
    Begin 
      Select ʵ�ձ��� 
      Into n_Discount_Rate 
      From (Select ʵ�ձ��� 
             From �ѱ���ϸ 
             Where �ѱ� = Nvl(�ѱ�_In, '-') And �շ�ϸĿid = Nvl(�շ�ϸĿid_In, 0) And (���_In Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) 
             Union All 
             Select ʵ�ձ��� 
             From �ѱ���ϸ 
             Where �ѱ� = Nvl(�ѱ�_In, '-') And ������Ŀid = Nvl(������Ŀid_In, 0) And (���_In Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) And Not Exists 
              (Select 1 From �ѱ���ϸ Where �ѱ� = Nvl(�ѱ�_In, '-') And �շ�ϸĿid = Nvl(�շ�ϸĿid_In, 0))); 
    Exception 
      When Others Then 
        n_Discount_Rate := 100.00; 
    End; 
    n_Discount_Rate := Nvl(n_Discount_Rate, 100); 
    Return n_Discount_Rate; 
  End Get_Discount_Rate; 
 
Begin 
 
  --��ȡ������Ϣ 
  Begin 
    Open c_Pati; 
    Fetch c_Pati 
      Into r_Pati; 
  Exception 
    When Others Then 
      Return; 
  End; 
 
  If Nvl(ǿ�Ƽ���_In, 0) = 0 And Nvl(r_Pati.�Ƿ��ֹ�Զ�����, 0) = 1 Then 
    Return; 
  End If; 
 
  n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0); 
  n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0); 
 
  If n_������˷�ʽ = 1 And Nvl(r_Pati.��˱�־, 0) >= 1 Then 
    Return; 
  End If; 
 
  If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then 
    Return; 
  End If; 
 
  -------------------------------------------------------------------------------- 
  --1.��ʼ����صĲ��� 
  n_�Ƿ��ü۸�ȼ� := ���ü۸�ȼ�_In; 
  If n_�Ƿ��ü۸�ȼ� < 0 Then 
    Select Nvl(Max(1), 0) Into n_�Ƿ��ü۸�ȼ� From �շѼ۸�ȼ�Ӧ�� Where Rownum < 2; 
  End If; 
  --ÿ��5����ǰ������¼ʱ��Ǽ�Ϊ���죬����Ǽ�Ϊ��ʱ 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_�Ǽ�ʱ�� 
  From Dual; 
 
  v_���ʽ�۸�ȼ� := Null; 
  If Nvl(n_�Ƿ��ü۸�ȼ�, 0) = 1 Then 
    Select Max(�۸�ȼ�) 
    Into v_���ʽ�۸�ȼ� 
    From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B 
    Where a.�۸�ȼ� = b.���� And a.���� = 1 And a.ҽ�Ƹ��ʽ = Nvl(r_Pati.���ʽ, '-') And Nvl(b.�Ƿ�������ͨ��Ŀ, 0) = 1 And 
          Nvl(b.����ʱ��, Sysdate + 1) > Sysdate; 
  End If; 
 
  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(160), '0')) 
  Into n_Dec, n_����۸����� 
  From Dual; 
 
  n_����۸����� := 1; 
  --ÿ��5����ǰ������¼ʱ��Ǽ�Ϊ���죬����Ǽ�Ϊ��ʱ 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_�Ǽ�ʱ�� 
  From Dual; 
  -------------------------------------------------------------------------------- 
 
  --�����ò��˵ļ�¼,�����ظ����� 
  Update ������ҳ Set ״̬ = ״̬ Where ����id = ����id_In And ��ҳid = ��ҳid_In; 
 
  -------------------------------------------------------------------------------- 
  --2. �Ƚ��䶯��Ϣ����¼��,�Ա���ȡ����ҽʦ�����λ�ʿ 
  d_����ʱ��_Max := Null; 
  d_����ʱ��_Max := Null; 
  d_��λʱ��_Max := Null; 
  For c_�䶯 In (Select ID, ��ʼʱ��, Nvl(��ֹʱ��, Sysdate + 1) As ��ֹʱ��, ����id, ����id, ����ȼ�id, ��λ�ȼ�id, ����ҽʦ, ���λ�ʿ, ҽ��С��id 
               From ���˱䶯��¼ A 
               Where ����id = ����id_In And ��ҳid = ��ҳid_In And ����id Is Not Null And 
                     Nvl(��ֹʱ��, Sysdate) >= (Select Nvl(Min(�ϴμ���ʱ��), Sysdate - 1000) 
                                            From �����Զ����� 
                                            Where ����id = ����id_In And ��ҳid = ��ҳid_In) 
               Order By ����id, ����id, ��ʼʱ�� Desc) Loop 
 
    If c_�䶯.����ȼ�id Is Not Null And Nvl(d_����ʱ��_Max, c_�䶯.��ֹʱ�� - 1) <= c_�䶯.��ֹʱ�� Then 
      d_����ʱ��_Max := c_�䶯.��ֹʱ��; 
    End If; 
 
    If c_�䶯.��λ�ȼ�id Is Not Null And Nvl(d_��λʱ��_Max, c_�䶯.��ֹʱ�� - 1) <= c_�䶯.��ֹʱ�� Then 
      d_��λʱ��_Max := c_�䶯.��ֹʱ��; 
    End If; 
 
    If c_�䶯.����id Is Not Null And Nvl(d_����ʱ��_Max, c_�䶯.��ֹʱ�� - 1) <= c_�䶯.��ֹʱ�� Then 
      d_����ʱ��_Max := c_�䶯.��ֹʱ��; 
    End If; 
    r_���˱䶯.Extend; 
    r_���˱䶯(r_���˱䶯.Count).Id := c_�䶯.Id; 
    r_���˱䶯(r_���˱䶯.Count).��ʼʱ�� := c_�䶯.��ʼʱ��; 
    r_���˱䶯(r_���˱䶯.Count).��ֹʱ�� := c_�䶯.��ֹʱ��; 
    r_���˱䶯(r_���˱䶯.Count).����id := c_�䶯.����id; 
    r_���˱䶯(r_���˱䶯.Count).����id := c_�䶯.����id; 
    r_���˱䶯(r_���˱䶯.Count).����ҽʦ := c_�䶯.����ҽʦ; 
    r_���˱䶯(r_���˱䶯.Count).���λ�ʿ := c_�䶯.���λ�ʿ; 
    r_���˱䶯(r_���˱䶯.Count).ҽ��С��id := c_�䶯.ҽ��С��id; 
  End Loop; 
 
  --����12:00,��12:00Ϊ׼ 
  d_����ʱ��_Max := Zl_Date_Half(d_����ʱ��_Max, 1); 
  d_��λʱ��_Max := Zl_Date_Half(d_��λʱ��_Max, 1); 
  d_����ʱ��_Max := Zl_Date_Half(d_����ʱ��_Max, 1); 
 
  ----------------------------------------------------------------- 
  --ѭ���������������������ȷ���¼���ļ�¼ 
  ----------------------------------------------------------------- 
 
  d_Start_Date := Sysdate + 1000; 
  d_Temp       := Sysdate - 1000;  

  --1.���㴲λ�� 
  For c_�Զ����� In c_Pati_Change(����id_In, ��ҳid_In, r_Pati.����, �ڼ�_In) Loop 
 
    If v_���ʽ�۸�ȼ� Is Null Then 
      If Nvl(n_�Ƿ��ü۸�ȼ�, 0) = 1 And Nvl(r_Pati_Change_Pre.վ��, '-') <> Nvl(c_�Զ�����.վ��, '-') Then 
        v_Temp     := Nvl(Zl_Get_Pricegrade(c_�Զ�����.վ��, ����id_In, ��ҳid_In, r_Pati.���ʽ), '|||') || '||||'; 
        v_�۸�ȼ� := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
      End If; 
    Else 
      v_�۸�ȼ� := v_���ʽ�۸�ȼ�; 
    End If; 
 
    r_Pati_Change := c_�Զ�����; 
    If d_Start_Date > r_Pati_Change.��ʼ���� Then 
      d_Start_Date := r_Pati_Change.��ʼ����; 
    End If; 
 
    If Nvl(r_Pati_Change.����, 0) <> Nvl(n_�������, 0) Then 
      n_������� := Nvl(r_Pati_Change.����, 0); 
      Update סԺ���ü�¼ 
      Set ���ӱ�־ = 5 
      Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ <> 8 And Nvl(ҽ�����, 0) = 0 And 
            ����ʱ�� >= r_Pati_Change.��ʼ���� And ���ӱ�־ <> 5 And (��ҩ���� Is Null Or ��ҩ���� = Nvl(n_�������, 0)); 
    End If; 
 
    --����ÿ����� 
    For I In 0 .. r_Pati_Change.���� Loop 
      v_����վ��    := Null; 
      r_Pati_Change := c_�Զ�����; 
      d_����ʱ��    := Greatest(c_�Զ�����.��ʼ����, Trunc(c_�Զ�����.��ʼ���� + I)); 
      n_Dates       := Least(Trunc(c_�Զ�����.��ʼ���� + I + 1), c_�Զ�����.��ֹ����) - Greatest(c_�Զ�����.��ʼ����, Trunc(c_�Զ�����.��ʼ���� + I)); 
 
      If r_Pati_Change.���� <> 2 Then 
        ----������������:  ��Ժ���찴һ�����,��Ժ��������12��֮ǰ����죬12��֮����һ�� 
        If (r_Pati_Change.���� = 1 And Trunc(d_����ʱ��_Max) = Trunc(d_����ʱ��) And d_����ʱ��_Max = r_Pati_Change.��ֹ����) Or 
           (r_Pati_Change.���� = 3 And Trunc(d_����ʱ��_Max) = Trunc(d_����ʱ��) And d_����ʱ��_Max = r_Pati_Change.��ֹ���� And 
           r_Pati_Change.�����־ = 7) Or (r_Pati_Change.���� = 3 And Trunc(d_��λʱ��_Max) = Trunc(d_����ʱ��) And 
           d_��λʱ��_Max = r_Pati_Change.��ֹ���� And r_Pati_Change.�����־ = 6) Then 
          If To_Char(r_Pati_Change.��ֹ����, 'hh24') >= 12 Then 
            n_Dates := 1; 
          Else 
            n_Dates    := 0.5; 
            d_����ʱ�� := Trunc(d_����ʱ��) + 0.5; 
          End If; 
        Else 
          n_Dates    := Least(Trunc(c_�Զ�����.��ʼ���� + I + 1), Trunc(c_�Զ�����.��ֹ����)) - 
                        Greatest(c_�Զ�����.��ʼ����, Trunc(c_�Զ�����.��ʼ���� + I)); 
          d_����ʱ�� := Trunc(d_����ʱ��); 
        End If; 
      End If; 
 
      If n_����۸����� = 1 And c_�Զ�����.���� = 1 Then 
        If d_����ʱ�� <> d_Temp Or Nvl(n_����, 0) <> Nvl(c_�Զ�����.����, 0) Then 
          l_Mulit_ϸĿid.Delete; 
          d_Temp := d_����ʱ��; 
          n_���� := Nvl(c_�Զ�����.����, 0); 
        End If; 
 
        n_Finded := 0; 
        For J In 1 .. l_Mulit_ϸĿid.Count Loop 
          If l_Mulit_ϸĿid(J) = c_�Զ�����.�շ�ϸĿid Then 
            n_Finded := 1; 
            Exit; 
          End If; 
        End Loop; 
        If n_Finded = 0 Then 
          l_Mulit_ϸĿid.Extend; 
          l_Mulit_ϸĿid(l_Mulit_ϸĿid.Count) := c_�Զ�����.�շ�ϸĿid; 
        End If; 
      End If; 
 
      n_Last         := 0; 
      n_�Ƿ������� := 1; 
      If d_����ʱ�� > r_Pati_Change.��Ŀ����ʱ�� Or n_Dates <= 0 Or (d_����ʱ�� > r_Pati_Change.��ֹ����) Then 
        Select Nvl(Max(1), 0) 
        Into n_Exists 
        From �����Զ����� A 
        Where (a.��ֹԭ�� = 1 Or a.��ֹԭ�� = 10) And a.Id = r_Pati_Change.Id And r_Pati_Change.���� <> 2; 
        If n_Exists = 0 Or n_Dates <= 0 Then 
          n_�Ƿ������� := 0; 
        Else 
          n_Last := 0.5; 
        End If; 
      End If; 
 
      Select Nvl(Max(1), 0) 
      Into n_Exists 
      From �����Զ����� A 
      Where a.��ֹԭ�� = 1 And a.Id = r_Pati_Change.Id And Exists 
       (Select 1 
             From �����Զ����� 
             Where ����id = a.����id And ��ҳid = a.��ҳid And ��ʼԭ�� = 2 And Trunc(��ʼʱ��) = Trunc(a.��ֹʱ��)); 
 
      If n_Exists = 1 Then 
        n_�Ƿ������� := 1; 
        n_Dates        := 1; 
      End If; 
 
      If n_�Ƿ������� = 1 Then 
 
        If d_����ʱ�� = Trunc(r_Pati_Change.��ʼʱ��) And Nvl(r_Pati_Change.���Ӵ�λ, 0) = 0 Then 
          --12����ǰ����ת�������Ϊ׼;12���Ժ���ת��Ϊ׼ 
          If To_Char(r_Pati_Change.��ʼʱ��, 'hh24') >= 12 Then 
            --�����䶯��¼�Ĵ��� 
            Begin 
              n_Find := 1; 
              Select ����id, ����id, Decode(r_Pati_Change.����, 1, ����ȼ�id, 2, ��λ�ȼ�id, r_Pati_Change.�շ�ϸĿid), ����, ��λ�ȼ�id 
              Into n_ǰ����id, n_ǰ����id, n_ǰ�շ�ϸĿid, v_ǰ����, n_ǰ��λ�ȼ�id 
              From �����Զ����� 
              Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���� = r_Pati_Change.���� And 
                    ��ʼʱ�� = (Select Max(��ʼʱ��) 
                            From �����Զ����� 
                            Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���� = r_Pati_Change.���� And 
                                  ��ʼʱ�� <= To_Date(To_Char(r_Pati_Change.��ʼʱ��, 'yyyy-mm-dd') || ' 12:00:00', 
                                                  'yyyy-mm-dd hh24:mi:ss')); 
            Exception 
              When Others Then 
                n_Find := 0; 
            End; 
            If n_Find = 1 And n_ǰ�շ�ϸĿid Is Not Null And n_ǰ����id Is Not Null And n_ǰ����id Is Not Null And 
               Not (r_Pati_Change.���� = 3 And r_Pati_Change.�����־ = 6 And n_ǰ��λ�ȼ�id Is Null) Then 
              r_Pati_Change.����id     := n_ǰ����id; 
              r_Pati_Change.����id     := n_ǰ����id; 
              r_Pati_Change.�շ�ϸĿid := n_ǰ�շ�ϸĿid; 
              r_Pati_Change.����       := v_ǰ����; 
 
              Select Nvl(a.վ��, b.վ��) 
              Into v_����վ�� 
              From ���ű� A, ���ű� B 
              Where a.Id = n_ǰ����id And b.Id = n_ǰ����id; 
            End If; 
          End If; 
        End If; 
 
        If v_����վ�� Is Not Null Then 
          If v_���ʽ�۸�ȼ� Is Null Then 
            If Nvl(n_�Ƿ��ü۸�ȼ�, 0) = 1 Then 
              v_Temp     := Nvl(Zl_Get_Pricegrade(v_����վ��, ����id_In, ��ҳid_In, r_Pati.���ʽ), '|||') || '||||'; 
              v_�۸�ȼ� := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
            End If; 
          Else 
            v_�۸�ȼ� := v_���ʽ�۸�ȼ�; 
          End If; 
        Else 
          If v_���ʽ�۸�ȼ� Is Null Then 
            If Nvl(n_�Ƿ��ü۸�ȼ�, 0) = 1 Then 
              v_Temp     := Nvl(Zl_Get_Pricegrade(r_Pati_Change.վ��, ����id_In, ��ҳid_In, r_Pati.���ʽ), '|||') || '||||'; 
              v_�۸�ȼ� := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
            End If; 
          Else 
            v_�۸�ȼ� := v_���ʽ�۸�ȼ�; 
          End If; 
        End If; 
       --�ж��Ƿ��ֶ�����
       Select Count(1)
       Into n_Count
       From סԺ���ü�¼
       Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 2 And Nvl(���ӱ�־, 0) = 1 And
             �շ���� = Decode(r_Pati_Change.����, 1, 'H', 2, 'J', �շ����) And ����ʱ�� = d_����ʱ�� And
             �շ�ϸĿid = Decode(r_Pati_Change.����, 3, r_Pati_Change.�շ�ϸĿid, �շ�ϸĿid);
       
        If v_�۸�ȼ� Is Null Then 
          Open c_�۸�_Rec For 
            Select b.�ּ� As ��׼����, b.������Ŀid, c.�վݷ�Ŀ, m.���㵥λ, m.��� 
            From �շѼ�Ŀ B, ������Ŀ C, �շ���ĿĿ¼ M 
            Where b.�շ�ϸĿid = m.Id And b.�շ�ϸĿid = r_Pati_Change.�շ�ϸĿid And b.������Ŀid = c.Id And 
                  (c.����ʱ�� Is Null Or c.����ʱ�� > d_����ʱ��) And Trunc(d_����ʱ��) Between Trunc(b.ִ������) And 
                  Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And b.�۸�ȼ� Is Null; 
        Else 
          Open c_�۸�_Rec For 
            Select b.�ּ� As ��׼����, b.������Ŀid, c.�վݷ�Ŀ, m.���㵥λ, m.��� 
            From �շѼ�Ŀ B, ������Ŀ C, �շ���ĿĿ¼ M 
            Where b.�շ�ϸĿid = m.Id And b.�շ�ϸĿid = r_Pati_Change.�շ�ϸĿid And b.������Ŀid = c.Id And 
                  (c.����ʱ�� Is Null Or c.����ʱ�� > d_����ʱ��) And Trunc(d_����ʱ��) Between Trunc(b.ִ������) And 
                  Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                  (b.�۸�ȼ� = v_�۸�ȼ� Or 
                  (b.�۸�ȼ� Is Null And Not Exists 
                   (Select 1 
                     From �շѼ�Ŀ 
                     Where �շ�ϸĿid = r_Pati_Change.�շ�ϸĿid And �۸�ȼ� = Nvl(v_�۸�ȼ�, '-') And 
                           Trunc(d_����ʱ��) Between Trunc(ִ������) And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))))); 
        End If; 
 
        Loop 
          Fetch c_�۸�_Rec 
            Into n_��׼�۸�, n_������Ŀid, v_�վݷ�Ŀ, v_���㵥λ, v_���; 
          Exit When c_�۸�_Rec%NotFound; 
          --For c_�۸� In c_�۸�_Rec(n_�շ�ϸĿid, d_����ʱ��, v_�۸�ȼ�) Loop 
          --��ȡ��ǰ������Ŀ���շѱ��� 
          n_Exsetax := Get_Discount_Rate(r_Pati.�ѱ�, r_Pati_Change.�շ�ϸĿid, n_������Ŀid, Abs(n_��׼�۸� * r_Pati_Change.����)); 
 
          --����Ѿ����㣬ԭ��¼������ȫ��ȷ����ֱ���޸Ľ���־���� 
          Update סԺ���ü�¼ 
          Set ���ӱ�־ = 0, ��ҩ���� = r_Pati_Change.���� 
          Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And 
                Nvl(�Ӱ��־, 0) = Nvl(r_Pati_Change.���Ӵ�λ, 0) And ���˿���id = r_Pati_Change.����id And 
                ���˲���id = Nvl(r_Pati_Change.����id, 0) And Nvl(����, 0) = Nvl(r_Pati_Change.����, 0) And 
                �շ�ϸĿid = r_Pati_Change.�շ�ϸĿid And ������Ŀid = n_������Ŀid And ����ʱ�� = d_����ʱ�� And 
                ���� = r_Pati_Change.���� * n_Dates And ��׼���� = n_��׼�۸� And 
                Ӧ�ս�� = Round(n_��׼�۸� * r_Pati_Change.���� * n_Dates, n_Dec) And 
                ʵ�ս�� = Round(n_��׼�۸� * r_Pati_Change.���� * n_Dates * n_Exsetax / 100, n_Dec); 
                          
          If Sql%RowCount = 0 And n_Count =0 Then 
            --���δ�������������������ȷ�ļ����¼ 
            r_�䶯_Cur.Delete; 
            r_�䶯_Cur.Extend; 
            For Q In 1 .. r_���˱䶯.Count Loop 
              If r_���˱䶯(Q).����id = r_Pati_Change.����id And r_���˱䶯(Q).����id = r_Pati_Change.����id And 
                  d_����ʱ�� - n_Last Between Trunc(r_���˱䶯(Q).��ʼʱ��) And r_���˱䶯(Q).��ֹʱ�� Then 
                r_�䶯_Cur(r_�䶯_Cur.Count).Id := r_���˱䶯(Q).Id; 
                r_�䶯_Cur(r_�䶯_Cur.Count).��ʼʱ�� := r_���˱䶯(Q).��ʼʱ��; 
                r_�䶯_Cur(r_�䶯_Cur.Count).��ֹʱ�� := r_���˱䶯(Q).��ֹʱ��; 
                r_�䶯_Cur(r_�䶯_Cur.Count).����id := r_���˱䶯(Q).����id; 
                r_�䶯_Cur(r_�䶯_Cur.Count).����id := r_���˱䶯(Q).����id; 
                r_�䶯_Cur(r_�䶯_Cur.Count).����ҽʦ := r_���˱䶯(Q).����ҽʦ; 
                r_�䶯_Cur(r_�䶯_Cur.Count).���λ�ʿ := r_���˱䶯(Q).���λ�ʿ; 
                r_�䶯_Cur(r_�䶯_Cur.Count).ҽ��С��id := r_���˱䶯(Q).ҽ��С��id; 
                Exit; 
              End If; 
            End Loop; 
 
            If v_Billno Is Null Then 
              v_Billno := Nextno(17); 
            End If; 
            Insert Into סԺ���ü�¼ 
              (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ����, �Ա�, 
               ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, ������, 
               ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id, ��ҩ����) 
              Select ���˷��ü�¼_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(r_Pati_Change.��ҳid, Null, 1, 2), r_Pati_Change.����id, r_Pati_Change.��ҳid, r_Pati_Change.����id, 
                     r_Pati_Change.����id, r_Pati_Change.����id, r_Pati_Change.����id, r_Pati.����, r_Pati.�Ա�, r_Pati.����, 
                     r_Pati.סԺ��, r_Pati_Change.����, r_Pati.�ѱ�, 1, r_Pati_Change.�շ�ϸĿid, n_������Ŀid, 0, n_��׼�۸�, 1, 
                     r_Pati_Change.���� * n_Dates, Round(n_��׼�۸� * r_Pati_Change.���� * n_Dates, n_Dec), 
                     Round(n_��׼�۸� * r_Pati_Change.���� * n_Dates * n_Exsetax / 100, n_Dec), v_���, v_���㵥λ, 
                     r_Pati_Change.���Ӵ�λ, v_�վݷ�Ŀ,r_�䶯_Cur(1).����ҽʦ,r_�䶯_Cur(1).���λ�ʿ, r_Pati_Change.����Ա���, 
                     r_Pati_Change.����Ա����, d_����ʱ��, d_�Ǽ�ʱ��, Decode(r_Pati_Change.����, Null, 0, 1), r_Pati_Change.����id, 
                     Decode(Nvl(n_�㷨, 0), 1, 
                             Round(n_��׼�۸� * r_Pati_Change.���� * n_Dates * n_Exsetax / 100 * n_ͳ��ȶ� / 100, n_Dec), 2, n_ͳ��ȶ�, 
                             0),r_�䶯_Cur(1).ҽ��С��id, r_Pati_Change.���� 
              From Dual; 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End Loop; 
        Close c_�۸�_Rec; 
      End If; 
      r_Pati_Change_Pre := c_�Զ�����; 
    End Loop; 
  End Loop; 
 
  ----------------------------------------------------------------- 
  --������ǰ����Ĵ����¼ 
  ----------------------------------------------------------------- 
  Insert Into סԺ���ü�¼ 
    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ����, �Ա�, ����, ��ʶ��, 
     ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, ������, ����Ա���, ����Ա����, ����ʱ��, 
     �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id, ��ҩ����) 
    Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, 2, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, 
           ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, 0, ��׼����, ����, -����, -Ӧ�ս��, -ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, 
           ������, ����Ա���, ����Ա����, ����ʱ��, d_�Ǽ�ʱ��, ������Ŀ��, ���մ���id, -ͳ����, ҽ��С��id, ��ҩ���� 
    From סԺ���ü�¼ 
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 5 And ����ʱ�� >= d_Start_Date; 
 
  ----------------------------------------------------------------- 
  --��д������� 
  ----------------------------------------------------------------- 
  Select Sum(Decode(���ӱ�־, 0, 1, -1) * ʵ�ս��) As ʵ�ս�� 
  Into n_Summoney 
  From סԺ���ü�¼ 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And 
        (NO = v_Billno Or ���ӱ�־ = 5 And ����ʱ�� >= d_Start_Date); 
 
  Update ������� 
  Set ������� = Nvl(�������, 0) + Nvl(n_Summoney, 0) 
  Where ����id = ����id_In And ���� = 1 And ���� = 2 
  Returning ������� Into n_����ֵ; 
 
  If Sql%RowCount = 0 Then 
    Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 2, n_Summoney, 0); 
    n_����ֵ := n_Summoney; 
  End If; 
 
  If Nvl(n_����ֵ, 0) = 0 Then 
    Delete From ������� Where ���� = 1 And ����id = ����id_In And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --��д���˻��ܷ��� 
  ----------------------------------------------------------------- 
  n_Delete := 0; 
  For v_Currrow In c_Sumcur_Rec(v_Billno, d_Start_Date) Loop 
    Update ����δ����� 
    Set ��� = Nvl(���, 0) + Nvl(v_Currrow.ʵ�ս��, 0) 
    Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(v_Currrow.���˲���id, 0) And 
          Nvl(���˿���id, 0) = Nvl(v_Currrow.���˿���id, 0) And Nvl(��������id, 0) = Nvl(v_Currrow.��������id, 0) And 
          Nvl(ִ�в���id, 0) = Nvl(v_Currrow.ִ�в���id, 0) And Nvl(������Ŀid, 0) = Nvl(v_Currrow.������Ŀid, 0) And ��Դ;�� + 0 = 2 
    Returning ��� Into n_����ֵ; 
 
    If Sql%RowCount = 0 Then 
      Insert Into ����δ����� 
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���) 
      Values 
        (����id_In, ��ҳid_In, v_Currrow.���˲���id, v_Currrow.���˿���id, v_Currrow.��������id, v_Currrow.ִ�в���id, v_Currrow.������Ŀid, 2, 
         v_Currrow.ʵ�ս��); 
      n_����ֵ := v_Currrow.ʵ�ս��; 
    End If; 
    If Nvl(n_����ֵ, 0) = 0 Then 
      n_Delete := 1; 
    End If; 
  End Loop; 
 
  If Nvl(n_Delete, 0) = 1 Then 
    Delete From ����δ����� Where ����id = ����id_In And ��� = 0; 
  End If; 
  ----------------------------------------------------------------- 
  --�������޸ĵĸ��ӱ�־��ԭΪ������־ 
  ----------------------------------------------------------------- 
  Update סԺ���ü�¼ 
  Set ���ӱ�־ = 0, ��¼״̬ = 3 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 5 And ����ʱ�� >= d_Start_Date; 
  ----------------------------------------------------------------- 
  --�޸ļ���ʱ���־ 
  ----------------------------------------------------------------- 
  Update �����Զ����� 
  Set �ϴμ���ʱ�� = Greatest(Sysdate, Nvl(��ֹʱ��, Sysdate)) 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(��ֹʱ��, Sysdate) > d_Start_Date; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl1_Autocalc_Pati_Charge_Nm;
/

--123419:����,2018-05-10,�����ֹ����Զ����ʷ��ø��ӱ�־����
CREATE OR REPLACE Procedure Zl1_Autocalc_Pati_Charge
( 
  ����id_In       In Number, 
  ��ҳid_In       In Number, 
  �ڼ�_In         In Varchar2, 
  ǿ�Ƽ���_In     In Number := 0, 
  ���ü۸�ȼ�_In In Number := -1 
) As 
  ------------------------------------------------------------------------- 
  --����˵�������ָ������ָ���ڼ���Զ����� 
  --          1��ϵͳ���ȸ���ϵͳ����"���������Զ��Ʒ�"���޸������ò����Զ����ʼ�¼��־; 
  --          2���ۺϲ��˵Ĵ�λ�仯�����ת�������������ȶ������أ�����ڼ��ȡ����˷� 
  --             �����ɷ��õ���ȷ���㣺 
  --             ��������Ѿ����㣬���޸ı�־Ϊ����;���δ���㣬������µ��Զ����ʼ�¼; 
  --             ������ǰ�Ĵ������ļ�¼; 
  --             ͳ�Ʊ��α䶯(����������)����д����ͻ��ܱ�; 
  --��ڲ����� 
  --       ����ID_IN  number    �������ID 
  --       ��ҳID_IN  number    ������ҳID������������ͬȷ����Ҫ����Ĳ��� 
  --       �ڼ�_IN  varchar2     ��Ҫ�������С�ڼ� 
  --       ǿ�Ƽ���_IN number   Ϊ1ʱ,���ܲ�����ҳ.��ֹ�Զ��������Կ��� 
  --       ���ü۸�ȼ�_In number ��-1��ʾδ�жϼ۸�ȼ�,�ڲ����Զ�ȥ���;0-�����ü۸�ȼ�;1-�����˼۸�ȼ��� 
  --���ù�ϵ��zl1_AutoCptPati/zl1_AutoCptWard/zl1_AutoCptAll ���ñ����� 
  ------------------------------------------------------------------------- 
  v_�۸�ȼ�         �շѼ۸�ȼ�.����%Type; 
  v_վ��_Pre         �շѼ۸�ȼ�.����%Type; 
  v_���ʽ�۸�ȼ� �շѼ۸�ȼ�.����%Type; 
 
  v_Temp      Varchar2(500); 
  v_Billno    Varchar2(8); --���ñ�ʵ�ʵ��Զ����ʺ��� 
  n_Billcount Number(5) := 0; --������ż����� 
 
  n_Exsetax  Number(16, 2) := 0; --������ȡ���� 
  n_Summoney Number(16, 2) := 0; --��� 
 
  n_Dec        Number; --���С��λ�� 
  n_Dates      Number(6, 1); --��ǰ��¼��������ȫ��Ϊ1 
  n_Delete     Number; 
  n_����ֵ     �������.Ԥ�����%Type; 
  n_�շ�ϸĿid �շ���ĿĿ¼.Id%Type; 
 
  n_סԺ״̬       ������ҳ.״̬%Type; 
  n_������˷�ʽ   Number(2); 
  n_δ��ƽ�ֹ���� Number(2); 
  n_Count          Number(5);
  n_��λ����ģʽ Number(2); 
  n_�������ģʽ Number(2); 
  n_��������ģʽ Number(2); 
 
  n_��λ�۸�����   Number(2); 
  n_����۸�����   Number(2); 
  n_������۸����� Number(2); 
  n_�Ƿ��ü۸�ȼ� Number(2); --0-δ����;1-���� 
  n_�Ƿ�������   Number(2); 
  n_����           Number(2); --1-����;2- ��λ;3-���� 
  n_Finded         Number(2); 
  v_����ģʽ       Varchar2(50); 
  n_��������id     ���˱䶯��¼.����id%Type; 
  n_���˲���id     ���˱䶯��¼.����id%Type; 
  n_�������       Number(3); 
 
  d_Start_Date Date; 
  d_�Ǽ�ʱ��   Date; --�Ǽ�ʱ�� 
  d_����ʱ��   Date; --����ʱ�� 
  d_Temp       Date; 
 
  l_Mulit_ϸĿid t_Numlist := t_Numlist(); 
 
  Type t_���˱䶯_Rec Is Record( 
    ID         ���˱䶯��¼.Id%Type, 
    ��ʼʱ��   ���˱䶯��¼.��ʼʱ��%Type, 
    ��ֹʱ��   ���˱䶯��¼.��ʼʱ��%Type, 
    ����id     ���˱䶯��¼.����id%Type, 
    ����id     ���˱䶯��¼.����id%Type, 
    ����ҽʦ   ���˱䶯��¼.����ҽʦ%Type, 
    ���λ�ʿ   ���˱䶯��¼.���λ�ʿ%Type, 
    ҽ��С��id ���˱䶯��¼.ҽ��С��id%Type); 
 
  Type c_���˱䶯_Rec Is Table Of t_���˱䶯_Rec; 
  r_���˱䶯 c_���˱䶯_Rec := c_���˱䶯_Rec(); 
  r_�䶯_Cur c_���˱䶯_Rec := c_���˱䶯_Rec(); 
 
  Cursor c_Sumcur_Rec 
  ( 
    Billno    Varchar2, 
    Datestart Date 
  ) Is 
    Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Decode(���ӱ�־, 0, 1, -1) * Ӧ�ս��) As Ӧ�ս��, 
           Sum(Decode(���ӱ�־, 0, 1, -1) * ʵ�ս��) As ʵ�ս�� 
    From סԺ���ü�¼ 
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And 
          (NO = Billno Or ���ӱ�־ = 5 And ����ʱ�� >= Datestart) 
    Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid; 
 
  n_������Ŀid ������Ŀ.Id%Type; 
  v_�վݷ�Ŀ   ������Ŀ.�վݷ�Ŀ%Type; 
  v_���㵥λ   �շ���ĿĿ¼.���㵥λ%Type; 
  v_���       �շ���ĿĿ¼.���%Type; 
  n_��׼�۸�   �շѼ�Ŀ.�ּ�%Type; 
 
  n_�㷨     ����֧������.�㷨%Type; 
  n_ͳ��ȶ� ����֧������.ͳ��ȶ�%Type; 
 
  Type t_�۸�_Rec Is Ref Cursor; 
  c_�۸�_Rec t_�۸�_Rec; 
 
  Cursor c_Pati Is 
    Select a.����id, a.��ҳid, Nvl(a.����, i.����) As ����, Nvl(a.�Ա�, i.�Ա�) As �Ա�, Nvl(a.����, i.����) As ����, Nvl(a.סԺ��, i.סԺ��) As סԺ��, 
           a.�ѱ�, Nvl(a.����, 0) As ����, Nvl(a.��˱�־, 0) As ��˱�־, Nvl(a.״̬, 0) As סԺ״̬, Nvl(a.�Ƿ��ֹ�Զ�����, 0) As �Ƿ��ֹ�Զ�����, 
           a.ҽ�Ƹ��ʽ As ���ʽ 
    From ������ҳ A, ������Ϣ I 
    Where a.����id = i.����id And a.����id = ����id_In And a.��ҳid = ��ҳid_In; 
 
  r_Pati c_Pati%RowType; 
 
  Function Get_Discount_Rate 
  ( 
    �ѱ�_In       ������Ϣ.�ѱ�%Type, 
    �շ�ϸĿid_In �ѱ���ϸ.�շ�ϸĿid%Type, 
    ������Ŀid_In �ѱ���ϸ.������Ŀid%Type, 
    ���_In       �ѱ���ϸ.Ӧ�ն���ֵ%Type 
  ) Return Number As 
    n_Discount_Rate Number(16, 5); 
  Begin 
    Begin 
      Select ʵ�ձ��� 
      Into n_Discount_Rate 
      From (Select ʵ�ձ��� 
             From �ѱ���ϸ 
             Where �ѱ� = Nvl(�ѱ�_In, '-') And �շ�ϸĿid = Nvl(�շ�ϸĿid_In, 0) And (���_In Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) 
             Union All 
             Select ʵ�ձ��� 
             From �ѱ���ϸ 
             Where �ѱ� = Nvl(�ѱ�_In, '-') And ������Ŀid = Nvl(������Ŀid_In, 0) And (���_In Between Ӧ�ն���ֵ And Ӧ�ն�βֵ) And Not Exists 
              (Select 1 From �ѱ���ϸ Where �ѱ� = Nvl(�ѱ�_In, '-') And �շ�ϸĿid = Nvl(�շ�ϸĿid_In, 0))); 
    Exception 
      When Others Then 
        n_Discount_Rate := 100.00; 
    End; 
    n_Discount_Rate := Nvl(n_Discount_Rate, 100); 
    Return n_Discount_Rate; 
  End Get_Discount_Rate; 
 
Begin 
 
  --��ȡ������Ϣ 
  Begin 
    Open c_Pati; 
    Fetch c_Pati 
      Into r_Pati; 
  Exception 
    When Others Then 
      Return; 
  End; 
  If Nvl(ǿ�Ƽ���_In, 0) = 0 And Nvl(r_Pati.�Ƿ��ֹ�Զ�����, 0) = 1 Then 
    Return; 
  End If; 
 
  n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0); 
  n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0); 
 
  If n_������˷�ʽ = 1 And Nvl(r_Pati.��˱�־, 0) >= 1 Then 
    Return; 
  End If; 
 
  If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then 
    Return; 
  End If; 
 
  -------------------------------------------------------------------------------- 
  --1.��ʼ����صĲ��� 
  v_����ģʽ := Nvl(zl_GetSysParameter(100), '0'); 
  If Length(v_����ģʽ) = 3 Then 
    n_��λ����ģʽ := To_Number(Substr(v_����ģʽ, 1, 1)); 
    n_�������ģʽ := To_Number(Substr(v_����ģʽ, 2, 1)); 
    n_��������ģʽ := To_Number(Substr(v_����ģʽ, 3, 1)); 
  Else 
    n_��λ����ģʽ := To_Number(v_����ģʽ); 
    n_�������ģʽ := To_Number(v_����ģʽ); 
    n_��������ģʽ := To_Number(v_����ģʽ); 
  End If; 
 
  n_��λ�۸�����   := 0; 
  n_������۸����� := 0; 
 
  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(160), '0')) 
  Into n_Dec, n_����۸����� 
  From Dual; 
 
  n_�Ƿ��ü۸�ȼ� := ���ü۸�ȼ�_In; 
  If n_�Ƿ��ü۸�ȼ� < 0 Then 
    Select Nvl(Max(1), 0) Into n_�Ƿ��ü۸�ȼ� From �շѼ۸�ȼ�Ӧ�� Where Rownum < 2; 
  End If; 
  --ÿ��5����ǰ������¼ʱ��Ǽ�Ϊ���죬����Ǽ�Ϊ��ʱ 
  Select Decode(Sign(To_Number(To_Char(Sysdate, 'HH24')) - 5), -1, Trunc(Sysdate) - 1 / 24 / 60, Sysdate) 
  Into d_�Ǽ�ʱ�� 
  From Dual; 
 
  v_���ʽ�۸�ȼ� := Null; 
  If Nvl(n_�Ƿ��ü۸�ȼ�, 0) = 1 Then 
    Select Max(�۸�ȼ�) 
    Into v_���ʽ�۸�ȼ� 
    From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B 
    Where a.�۸�ȼ� = b.���� And a.���� = 1 And a.ҽ�Ƹ��ʽ = Nvl(r_Pati.���ʽ, '-') And Nvl(b.�Ƿ�������ͨ��Ŀ, 0) = 1 And 
          Nvl(b.����ʱ��, Sysdate + 1) > Sysdate; 
  End If; 
  -------------------------------------------------------------------------------- 
 
  --�����ò��˵ļ�¼,�����ظ����� 
  Update ������ҳ Set ״̬ = ״̬ Where ����id = ����id_In And ��ҳid = ��ҳid_In; 
 
  -------------------------------------------------------------------------------- 
  --2. �Ƚ��䶯��Ϣ����¼��,�Ա���ȡ����ҽʦ�����λ�ʿ 
  For c_�䶯 In (Select ID, ��ʼʱ��, Nvl(��ֹʱ��, Trunc(Sysdate) + 1) As ��ֹʱ��, ����id, ����id, ����ҽʦ, ���λ�ʿ, ҽ��С��id 
               From ���˱䶯��¼ A 
               Where ����id = ����id_In And ��ҳid = ��ҳid_In And ����id Is Not Null And 
                     Nvl(��ֹʱ��, Sysdate) >= (Select Nvl(Min(�ϴμ���ʱ��), Sysdate - 1000) 
                                            From �����Զ����� 
                                            Where ����id = ����id_In And ��ҳid = ��ҳid_In) 
               Order By ����id, ����id, ��ʼʱ�� Desc) Loop 
    r_���˱䶯.Extend; 
    r_���˱䶯(r_���˱䶯.Count).Id := c_�䶯.Id; 
    r_���˱䶯(r_���˱䶯.Count).��ʼʱ�� := c_�䶯.��ʼʱ��; 
    r_���˱䶯(r_���˱䶯.Count).��ֹʱ�� := c_�䶯.��ֹʱ��; 
    r_���˱䶯(r_���˱䶯.Count).����id := c_�䶯.����id; 
    r_���˱䶯(r_���˱䶯.Count).����id := c_�䶯.����id; 
    r_���˱䶯(r_���˱䶯.Count).����ҽʦ := c_�䶯.����ҽʦ; 
    r_���˱䶯(r_���˱䶯.Count).���λ�ʿ := c_�䶯.���λ�ʿ; 
    r_���˱䶯(r_���˱䶯.Count).ҽ��С��id := c_�䶯.ҽ��С��id; 
  End Loop; 
 
  ----------------------------------------------------------------- 
  --ѭ���������������������ȷ���¼���ļ�¼ 
  ----------------------------------------------------------------- 
  d_Start_Date := Sysdate + 1000; 
  d_Temp       := Sysdate - 1000; 
 
  --1.���㴲λ�� 
  For c_�Զ����� In (Select a.����, a.����id, a.��ҳid, a.����id, a.����id, a.����, a.���Ӵ�λ, a.�շ�ϸĿid, a.����Ա���, a.����Ա����, a.��ʼʱ��, 
                        Nvl(a.��ֹʱ��, Sysdate) As ��ֹʱ��, a.��������, a.����, Greatest(a.��ʼ����, Trunc(p.��ʼ����)) As ��ʼ����, 
                        Nvl(a.��ֹ����, Trunc(Sysdate)) As ��ֹ����, 
                        Nvl(a.��ֹ����, Trunc(Sysdate)) - Greatest(a.��ʼ����, Trunc(p.��ʼ����)) As ����, Nvl(Q1.վ��, Q2.վ��) As վ��, 
                        m.���㵥λ, m.���, i.����, i.����id, k.�㷨, k.ͳ��ȶ�, 
                        Nvl(m.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��Ŀ����ʱ��, a.�����־, a.Id 
                 From (Select b.Id, 2 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, b.��λ�ȼ�id As �շ�ϸĿid, b.����Ա���, 
                               b.����Ա����, b.��ʼʱ��, b.��ֹʱ��, a.��������, b.����, Zl_Date_Half(b.��ʼʱ��, n_��λ����ģʽ) As ��ʼ����, 
                               Zl_Date_Half(b.��ֹʱ��, n_��λ����ģʽ) As ��ֹ����, 0 As �����־ 
                        From �Զ��Ƽ���Ŀ A, 
                             (Select a.Id, a.����id, a.��ҳid, a.��ʼʱ��, a.���Ӵ�λ, a.����id, a.����id, a.����, a.��λ�ȼ�id, 1 As ����, a.��ֹʱ��, 
                                      a.����Ա���, a.����Ա����, a.�ϴμ���ʱ�� 
                               From �����Զ����� A 
                               Where a.���� = 2 And a.����id = ����id_In And a.��ҳid = ��ҳid_In And 
                                     Nvl(a.�ϴμ���ʱ��, a.��ʼʱ��) <= Nvl(a.��ֹʱ��, Sysdate) 
                               Union All 
                               Select b.Id, b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, 
                                      b.��ֹʱ��, b.����Ա���, b.����Ա����, b.�ϴμ���ʱ�� 
                               From �����Զ����� B, �շѴ�����Ŀ I 
                               Where b.���� = 2 And b.����id = ����id_In And b.��ҳid = ��ҳid_In And b.��λ�ȼ�id = i.����id And i.���д��� > 0 And 
                                     Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��) <= Nvl(b.��ֹʱ��, Sysdate)) B 
                        Where a.����id = b.����id And a.�����־ = 1 
                        Union All 
                        Select b.Id, 1 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, b.����ȼ�id As �շ�ϸĿid, b.����Ա���, 
                               b.����Ա����, b.��ʼʱ��, b.��ֹʱ��, a.��������, b.����, Zl_Date_Half(b.��ʼʱ��, n_�������ģʽ) As ��ʼ����, 
                               Zl_Date_Half(b.��ֹʱ��, n_�������ģʽ) As ��ֹ����, 0 As �����־ 
                        From �Զ��Ƽ���Ŀ A, 
                             (Select a.Id, a.����id, a.��ҳid, a.��ʼʱ��, a.���Ӵ�λ, a.����id, a.����id, a.����, a.����ȼ�id, 1 As ����, a.��ֹʱ��, 
                                      a.����Ա���, a.����Ա����, a.�ϴμ���ʱ�� 
                               From �����Զ����� A 
                               Where a.���� = 1 And a.����id = ����id_In And a.��ҳid = ��ҳid_In And 
                                     Nvl(a.�ϴμ���ʱ��, a.��ʼʱ��) <= Nvl(a.��ֹʱ��, Sysdate) 
                               Union All 
                               Select b.Id, b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, 
                                      b.��ֹʱ��, b.����Ա���, b.����Ա����, b.�ϴμ���ʱ�� 
                               From �����Զ����� B, �շѴ�����Ŀ I 
                               Where b.���� = 1 And b.����id = ����id_In And b.��ҳid = ��ҳid_In And b.����ȼ�id = i.����id And i.���д��� > 0 And 
                                     Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��) <= Nvl(b.��ֹʱ��, Sysdate)) B 
                        Where a.����id = b.����id And a.�����־ = 2 
                        Union All 
                        Select b.Id, 3 As ����, b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, a.�շ�ϸĿid, b.����Ա���, b.����Ա����, 
                               b.��ʼʱ��, b.��ֹʱ��, a.��������, a.����, Zl_Date_Half(b.��ʼʱ��, n_��������ģʽ) As ��ʼ����, 
                               Zl_Date_Half(b.��ֹʱ��, n_��������ģʽ) As ��ֹ����, a.�����־ 
                        From (Select ����id, �����־, �շ�ϸĿid, 1 As ����, �������� 
                               From �Զ��Ƽ���Ŀ 
                               Union All 
                               Select ����id, �����־, ����id, i.�������� As ����, �������� 
                               From �Զ��Ƽ���Ŀ A, �շѴ�����Ŀ I 
                               Where a.�շ�ϸĿid = i.����id And i.���д��� > 0) A, �����Զ����� B 
                        Where a.����id = b.����id And b.����id = ����id_In And b.��ҳid = ��ҳid_In And b.���� = 3 And 
                              Nvl(b.���Ӵ�λ, 0) = 0 And (a.�����־ = 6 And b.��λ�ȼ�id Is Not Null Or a.�����־ = 7) And 
                              Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��) <= Nvl(b.��ֹʱ��, Sysdate)) A, 
                      (Select Min(��ʼ����) As ��ʼ���� From �ڼ�� Where �ڼ� >= �ڼ�_In) P, ����֧����Ŀ I, ����֧������ K, �շ���ĿĿ¼ M, ���ű� Q1, 
                      ���ű� Q2 
                 Where Trunc(Nvl(a.��ֹʱ��, Greatest(a.��ʼʱ��, Sysdate))) >= Trunc(p.��ʼ����) And a.�շ�ϸĿid = i.�շ�ϸĿid(+) And 
                       i.����(+) = Nvl(r_Pati.����, 0) And i.����id = k.Id(+) And a.�շ�ϸĿid = m.Id And a.����id = Q1.Id(+) And 
                       a.����id = Q2.Id(+) 
                 Order By ����, ��ʼʱ��) Loop 
    --�������� 
    If v_���ʽ�۸�ȼ� Is Null Then 
      If Nvl(n_�Ƿ��ü۸�ȼ�, 0) = 1 And Nvl(v_վ��_Pre, '-') <> Nvl(c_�Զ�����.վ��, '-') Then 
        v_Temp     := Nvl(Zl_Get_Pricegrade(c_�Զ�����.վ��, ����id_In, ��ҳid_In, r_Pati.���ʽ), '|||') || '||||'; 
        v_�۸�ȼ� := Substr(v_Temp, 1, Instr(v_Temp, '|') - 1); 
      End If; 
    Else 
      v_�۸�ȼ� := v_���ʽ�۸�ȼ�; 
    End If; 
 
    v_վ��_Pre := c_�Զ�����.վ��; 
 
    If d_Start_Date > c_�Զ�����.��ʼ���� Then 
      d_Start_Date := c_�Զ�����.��ʼ����; 
    End If; 
 
    If Nvl(c_�Զ�����.����, 0) <> Nvl(n_�������, 0) Then 
      n_������� := Nvl(c_�Զ�����.����, 0); 
      Update סԺ���ü�¼ 
      Set ���ӱ�־ = 5 
      Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ <> 8 And Nvl(ҽ�����, 0) = 0 And 
            ����ʱ�� >= c_�Զ�����.��ʼ���� And ���ӱ�־ <> 5 And (��ҩ���� Is Null Or ��ҩ���� = Nvl(n_�������, 0)); 
    End If; 
 
    --����ÿ����� 
    For I In 0 .. c_�Զ�����.���� Loop 
      d_����ʱ�� := Greatest(c_�Զ�����.��ʼ����, Trunc(c_�Զ�����.��ʼ���� + I)); 
      n_Dates    := Least(Trunc(c_�Զ�����.��ʼ���� + I + 1), c_�Զ�����.��ֹ����) - Greatest(c_�Զ�����.��ʼ����, Trunc(c_�Զ�����.��ʼ���� + I)); 
      If n_Dates < 0 Then 
        n_Dates := 0; 
      End If; 
      If (n_����۸����� = 1 And c_�Զ�����.���� = 1) Or (n_��λ�۸����� = 1 And c_�Զ�����.���� = 2) Or (n_������۸����� = 1 And c_�Զ�����.���� = 3) Then 
 
        If d_����ʱ�� <> d_Temp Or Nvl(n_����, 0) <> Nvl(c_�Զ�����.����, 0) Then 
          l_Mulit_ϸĿid.Delete; 
          d_Temp := d_����ʱ��; 
          n_���� := Nvl(c_�Զ�����.����, 0); 
        End If; 
 
        n_Finded := 0; 
        For J In 1 .. l_Mulit_ϸĿid.Count Loop 
          If l_Mulit_ϸĿid(J) = c_�Զ�����.�շ�ϸĿid Then 
            n_Finded := 1; 
            Exit; 
          End If; 
        End Loop; 
        If n_Finded = 0 Then 
          l_Mulit_ϸĿid.Extend; 
          l_Mulit_ϸĿid(l_Mulit_ϸĿid.Count) := c_�Զ�����.�շ�ϸĿid; 
        End If; 
      End If; 
 
      n_�Ƿ������� := 1; 
      If d_����ʱ�� > c_�Զ�����.��Ŀ����ʱ�� Or n_Dates = 0 Then 
        n_�Ƿ������� := 0; 
      End If; 
 
      Select Decode(Nvl(Max(1), 0), 0, n_�Ƿ�������, 0) 
      Into n_�Ƿ������� 
      From �����Զ����� A 
      Where a.��ֹԭ�� = 1 And a.Id = c_�Զ�����.Id And Exists 
       (Select 1 
             From �����Զ����� 
             Where ����id = a.����id And ��ҳid = a.��ҳid And ��ʼԭ�� = 2 And Trunc(��ʼʱ��) = Trunc(a.��ֹʱ��)); 
 
      If n_�Ƿ������� = 1 Then 
        --��Ҫ����Ƿ���ָ�����ڱ�ͣ���˵� 
        n_�շ�ϸĿid := c_�Զ�����.�շ�ϸĿid; 
        n_�㷨       := c_�Զ�����.�㷨; 
        n_ͳ��ȶ�   := c_�Զ�����.ͳ��ȶ�; 
 
        If (n_����۸����� = 1 And c_�Զ�����.���� = 1) Or (n_��λ�۸����� = 1 And c_�Զ�����.���� = 2) Or 
           (n_������۸����� = 1 And c_�Զ�����.���� = 3) And l_Mulit_ϸĿid.Count > 1 Then 
          --ȡ��߼۸���շ���Ŀ 
          If v_�۸�ȼ� Is Null Then 
            Open c_�۸�_Rec For 
              Select b.�շ�ϸĿid, Sum(b.�ּ�) As ��׼���� 
              From �շѼ�Ŀ B, ������Ŀ C 
              Where b.�շ�ϸĿid In (Select Column_Value From Table(l_Mulit_ϸĿid)) And b.������Ŀid = c.Id And 
                    (c.����ʱ�� Is Null Or c.����ʱ�� > d_����ʱ��) And Trunc(d_����ʱ��) Between Trunc(b.ִ������) And 
                    Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And b.�۸�ȼ� Is Null 
 
              Group By �շ�ϸĿid 
              Order By ��׼���� Desc; 
          Else 
            Open c_�۸�_Rec For 
              Select b.�շ�ϸĿid, Sum(b.�ּ�) As ��׼���� 
              From �շѼ�Ŀ B, ������Ŀ C 
              Where b.�շ�ϸĿid In (Select Column_Value From Table(l_Mulit_ϸĿid)) And b.������Ŀid = c.Id And 
                    (c.����ʱ�� Is Null Or c.����ʱ�� > d_����ʱ��) And Trunc(d_����ʱ��) Between Trunc(b.ִ������) And 
                    Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                    (b.�۸�ȼ� = Nvl(v_�۸�ȼ�, '-') Or 
                    (b.�۸�ȼ� Is Null And Not Exists 
                     (Select 1 
                       From �շѼ�Ŀ 
                       Where �շ�ϸĿid = b.�շ�ϸĿid And �۸�ȼ� = Nvl(v_�۸�ȼ�, '-') And Trunc(d_����ʱ��) Between Trunc(ִ������) And 
                             Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))))) 
              Group By �շ�ϸĿid 
              Order By ��׼���� Desc; 
          End If; 
 
          Begin 
            Fetch c_�۸�_Rec 
              Into n_�շ�ϸĿid, n_��׼�۸�; 
          Exception 
            When Others Then 
              n_�շ�ϸĿid := c_�Զ�����.�շ�ϸĿid; 
          End; 
          Close c_�۸�_Rec; 
        End If; 
        n_��������id := c_�Զ�����.����id; 
        n_���˲���id := c_�Զ�����.����id; 
        If c_�Զ�����.�շ�ϸĿid <> n_�շ�ϸĿid Then 
          --��߼۸���շ�ϸĿ���ԣ�����ͳ��ȶһ�� 
          Select Max(k.�㷨), Max(k.ͳ��ȶ�) 
          Into n_�㷨, n_ͳ��ȶ� 
          From ����֧����Ŀ I, ����֧������ K 
          Where i.�շ�ϸĿid = n_�շ�ϸĿid And i.����(+) = Nvl(r_Pati.����, 0) And i.����id = k.Id(+); 
          If n_����۸����� = 1 And c_�Զ�����.���� = 1 Then 
            For c_�䶯��¼ In (Select ����id, ����id 
                           From ���˱䶯��¼ 
                           Where ��ʼԭ�� <> 10 And ����id = ����id_In And ��ҳid = ��ҳid_In And ����ȼ�id + 0 = n_�շ�ϸĿid And 
                                 (Trunc(��ʼʱ��) = Trunc(d_����ʱ��) Or Trunc(Nvl(��ֹʱ��, Sysdate)) = Trunc(d_����ʱ��)) 
                           Order By ��ʼʱ�� Desc) Loop 
              n_��������id := c_�䶯��¼.����id; 
              n_���˲���id := c_�䶯��¼.����id; 
              Exit; 
            End Loop; 
          End If; 
        End If; 
      End If; 
      --�ж��Ƿ��ֹ�����
      Select Count(1)
      Into n_Count
      From סԺ���ü�¼
      Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 2 And Nvl(���ӱ�־, 0) = 1 And
           �շ���� = Decode(c_�Զ�����.����, 1, 'H', 2, 'J', �շ����) And ����ʱ�� = d_����ʱ�� And
           �շ�ϸĿid = Decode(c_�Զ�����.����, 3, c_�Զ�����.�շ�ϸĿid, �շ�ϸĿid); 

      If n_�Ƿ������� = 1 Then 
        If v_�۸�ȼ� Is Null Then 
          Open c_�۸�_Rec For 
            Select b.�ּ� As ��׼����, b.������Ŀid, c.�վݷ�Ŀ, m.���㵥λ, m.��� 
            From �շѼ�Ŀ B, ������Ŀ C, �շ���ĿĿ¼ M 
            Where b.�շ�ϸĿid = m.Id And b.�շ�ϸĿid = n_�շ�ϸĿid And b.������Ŀid = c.Id And (c.����ʱ�� Is Null Or c.����ʱ�� > d_����ʱ��) And 
                  Trunc(d_����ʱ��) Between Trunc(b.ִ������) And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                  b.�۸�ȼ� Is Null; 
        Else 
          Open c_�۸�_Rec For 
            Select b.�ּ� As ��׼����, b.������Ŀid, c.�վݷ�Ŀ, m.���㵥λ, m.��� 
            From �շѼ�Ŀ B, ������Ŀ C, �շ���ĿĿ¼ M 
            Where b.�շ�ϸĿid = m.Id And b.�շ�ϸĿid = n_�շ�ϸĿid And b.������Ŀid = c.Id And (c.����ʱ�� Is Null Or c.����ʱ�� > d_����ʱ��) And 
                  Trunc(d_����ʱ��) Between Trunc(b.ִ������) And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And 
                  (b.�۸�ȼ� = v_�۸�ȼ� Or 
                  (b.�۸�ȼ� Is Null And Not Exists 
                   (Select 1 
                     From �շѼ�Ŀ 
                     Where �շ�ϸĿid = n_�շ�ϸĿid And �۸�ȼ� = Nvl(v_�۸�ȼ�, '-') And Trunc(d_����ʱ��) Between Trunc(ִ������) And 
                           Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))))); 
        End If; 
 
        Loop 
          Fetch c_�۸�_Rec 
            Into n_��׼�۸�, n_������Ŀid, v_�վݷ�Ŀ, v_���㵥λ, v_���; 
          Exit When c_�۸�_Rec%NotFound; 
          --For c_�۸� In c_�۸�_Rec(n_�շ�ϸĿid, d_����ʱ��, v_�۸�ȼ�) Loop 
          --��ȡ��ǰ������Ŀ���շѱ��� 
          n_Exsetax := Get_Discount_Rate(r_Pati.�ѱ�, n_�շ�ϸĿid, n_������Ŀid, Abs(n_��׼�۸� * c_�Զ�����.����)); 
 
          --����Ѿ����㣬ԭ��¼������ȫ��ȷ����ֱ���޸Ľ���־���� 
          Update סԺ���ü�¼ 
          Set ���ӱ�־ = 0 
          Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And Nvl(�Ӱ��־, 0) = Nvl(c_�Զ�����.���Ӵ�λ, 0) And 
                ���˿���id = n_��������id And ���˲���id = Nvl(n_���˲���id, 0) And Nvl(����, 0) = Nvl(c_�Զ�����.����, 0) And 
                �շ�ϸĿid = n_�շ�ϸĿid And ������Ŀid = n_������Ŀid And ����ʱ�� = d_����ʱ�� And ���� = c_�Զ�����.���� * n_Dates And 
                ��׼���� = n_��׼�۸� And Ӧ�ս�� = Round(n_��׼�۸� * c_�Զ�����.���� * n_Dates, n_Dec) And 
                ʵ�ս�� = Round(n_��׼�۸� * c_�Զ�����.���� * n_Dates * n_Exsetax / 100, n_Dec); 
 
          If Sql%RowCount = 0 And n_Count=0 Then 
            --���δ�������������������ȷ�ļ����¼ 
            r_�䶯_Cur.Delete; 
            r_�䶯_Cur.Extend; 
            For Q In 1 .. r_���˱䶯.Count Loop 
              If r_���˱䶯(Q) 
               .����id = c_�Զ�����.����id And r_���˱䶯(Q).����id = c_�Զ�����.����id And d_����ʱ�� Between Trunc(r_���˱䶯(Q).��ʼʱ��) And r_���˱䶯(Q).��ֹʱ�� Then 
                r_�䶯_Cur(r_�䶯_Cur.Count).Id := r_���˱䶯(Q).Id; 
                r_�䶯_Cur(r_�䶯_Cur.Count).��ʼʱ�� := r_���˱䶯(Q).��ʼʱ��; 
                r_�䶯_Cur(r_�䶯_Cur.Count).��ֹʱ�� := r_���˱䶯(Q).��ֹʱ��; 
                r_�䶯_Cur(r_�䶯_Cur.Count).����id := r_���˱䶯(Q).����id; 
                r_�䶯_Cur(r_�䶯_Cur.Count).����id := r_���˱䶯(Q).����id; 
                r_�䶯_Cur(r_�䶯_Cur.Count).����ҽʦ := r_���˱䶯(Q).����ҽʦ; 
                r_�䶯_Cur(r_�䶯_Cur.Count).���λ�ʿ := r_���˱䶯(Q).���λ�ʿ; 
                r_�䶯_Cur(r_�䶯_Cur.Count).ҽ��С��id := r_���˱䶯(Q).ҽ��С��id; 
                Exit; 
              End If; 
            End Loop; 
 
            If v_Billno Is Null Then 
              v_Billno := Nextno(17); 
            End If; 
            Insert Into סԺ���ü�¼ 
              (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ����, �Ա�, 
               ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, ������, 
               ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id, ��ҩ����) 
              Select ���˷��ü�¼_Id.Nextval, 3, v_Billno, 1, Rownum + n_Billcount, Null, Null, 0, Null, 
                     Decode(c_�Զ�����.��ҳid, Null, 1, 2), c_�Զ�����.����id, c_�Զ�����.��ҳid, c_�Զ�����.����id, n_��������id, n_��������id, 
                     n_���˲���id, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.סԺ��, c_�Զ�����.����, r_Pati.�ѱ�, 1, n_�շ�ϸĿid, n_������Ŀid, 
                     0, n_��׼�۸�, 1, c_�Զ�����.���� * n_Dates, Round(n_��׼�۸� * c_�Զ�����.���� * n_Dates, n_Dec), 
                     Round(n_��׼�۸� * c_�Զ�����.���� * n_Dates * n_Exsetax / 100, n_Dec), v_���, v_���㵥λ, c_�Զ�����.���Ӵ�λ, v_�վݷ�Ŀ, 
                     r_�䶯_Cur(1).����ҽʦ,r_�䶯_Cur(1).���λ�ʿ, c_�Զ�����.����Ա���, c_�Զ�����.����Ա����, d_����ʱ��, d_�Ǽ�ʱ��, 
                     Decode(c_�Զ�����.����, Null, 0, 1), c_�Զ�����.����id, 
                     Decode(Nvl(n_�㷨, 0), 1, Round(n_��׼�۸� * c_�Զ�����.���� * n_Dates * n_Exsetax / 100 * n_ͳ��ȶ� / 100, n_Dec), 
                             2, n_ͳ��ȶ�, 0),r_�䶯_Cur(1).ҽ��С��id, n_������� 
              From Dual; 
            n_Billcount := n_Billcount + Sql%RowCount; 
          End If; 
        End Loop; 
        Close c_�۸�_Rec; 
      End If; 
    End Loop; 
  End Loop; 
 
  ----------------------------------------------------------------- 
  --������ǰ����Ĵ����¼ 
  ----------------------------------------------------------------- 
  Insert Into סԺ���ü�¼ 
    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ����, �Ա�, ����, ��ʶ��, 
     ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, ���ӱ�־, ��׼����, ����, ����, Ӧ�ս��, ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, ������, ����Ա���, ����Ա����, ����ʱ��, 
     �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ҽ��С��id, ��ҩ����) 
    Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, 2, ���, ��������, �۸񸸺�, �ಡ�˵�, ҽ�����, �����־, ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, 
           ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���ʷ���, �շ�ϸĿid, ������Ŀid, 0, ��׼����, ����, -����, -Ӧ�ս��, -ʵ�ս��, �շ����, ���㵥λ, �Ӱ��־, �վݷ�Ŀ, ������, 
           ������, ����Ա���, ����Ա����, ����ʱ��, d_�Ǽ�ʱ��, ������Ŀ��, ���մ���id, -ͳ����, ҽ��С��id, ��ҩ���� 
    From סԺ���ü�¼ 
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 5 And ����ʱ�� >= d_Start_Date; 
 
  ----------------------------------------------------------------- 
  --��д������� 
  ----------------------------------------------------------------- 
  Select Sum(Decode(���ӱ�־, 0, 1, -1) * ʵ�ս��) As ʵ�ս�� 
  Into n_Summoney 
  From סԺ���ü�¼ 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And 
        (NO = v_Billno Or ���ӱ�־ = 5 And ����ʱ�� >= d_Start_Date); 
 
  Update ������� 
  Set ������� = Nvl(�������, 0) + Nvl(n_Summoney, 0) 
  Where ����id = ����id_In And ���� = 1 And ���� = 2 
  Returning ������� Into n_����ֵ; 
 
  If Sql%RowCount = 0 Then 
    Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 2, n_Summoney, 0); 
    n_����ֵ := n_Summoney; 
  End If; 
 
  If Nvl(n_����ֵ, 0) = 0 Then 
    Delete From ������� Where ���� = 1 And ����id = ����id_In And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0; 
  End If; 
 
  ----------------------------------------------------------------- 
  --��д���˻��ܷ��� 
  ----------------------------------------------------------------- 
  n_Delete := 0; 
  For v_Currrow In c_Sumcur_Rec(v_Billno, d_Start_Date) Loop 
    Update ����δ����� 
    Set ��� = Nvl(���, 0) + Nvl(v_Currrow.ʵ�ս��, 0) 
    Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(v_Currrow.���˲���id, 0) And 
          Nvl(���˿���id, 0) = Nvl(v_Currrow.���˿���id, 0) And Nvl(��������id, 0) = Nvl(v_Currrow.��������id, 0) And 
          Nvl(ִ�в���id, 0) = Nvl(v_Currrow.ִ�в���id, 0) And Nvl(������Ŀid, 0) = Nvl(v_Currrow.������Ŀid, 0) And ��Դ;�� + 0 = 2 
    Returning ��� Into n_����ֵ; 
 
    If Sql%RowCount = 0 Then 
      Insert Into ����δ����� 
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���) 
      Values 
        (����id_In, ��ҳid_In, v_Currrow.���˲���id, v_Currrow.���˿���id, v_Currrow.��������id, v_Currrow.ִ�в���id, v_Currrow.������Ŀid, 2, 
         v_Currrow.ʵ�ս��); 
      n_����ֵ := v_Currrow.ʵ�ս��; 
    End If; 
    If Nvl(n_����ֵ, 0) = 0 Then 
      n_Delete := 1; 
    End If; 
  End Loop; 
 
  If Nvl(n_Delete, 0) = 1 Then 
    Delete From ����δ����� Where ����id = ����id_In And ��� = 0; 
  End If; 
  ----------------------------------------------------------------- 
  --�������޸ĵĸ��ӱ�־��ԭΪ������־ 
  ----------------------------------------------------------------- 
  Update סԺ���ü�¼ 
  Set ���ӱ�־ = 0, ��¼״̬ = 3 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 5 And ����ʱ�� >= d_Start_Date; 
  ----------------------------------------------------------------- 
  --�޸ļ���ʱ���־ 
  ----------------------------------------------------------------- 
  Update �����Զ����� 
  Set �ϴμ���ʱ�� = Greatest(Sysdate, Nvl(��ֹʱ��, Sysdate)) 
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(��ֹʱ��, Sysdate) > d_Start_Date; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl1_Autocalc_Pati_Charge;
/

--123419:����,2018-05-10,�����ֹ����Զ����ʷ��ø��ӱ�־����
Create Or Replace Procedure Zl_סԺ���ʼ�¼_Delete
(
  No_In           סԺ���ü�¼.No%Type,
  ���_In         Varchar2,
  ����Ա���_In   סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In   סԺ���ü�¼.����Ա����%Type,
  ��¼����_In     סԺ���ü�¼.��¼����%Type := 2,
  ����״̬_In     Number := 0,
  ��Һ��ҩ���_In Number := 1,
  �Ǽ�ʱ��_In     סԺ���ü�¼.�Ǽ�ʱ��%Type := Sysdate
) As
  --���ܣ�����һ��סԺ���ʵ�����ָ�������
  --��ţ���ʽ��"1,3,5,7,8",��"1:2:33456,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,�м�����ֱ�ʾ�˵�����,��������ֱ�ʾ��ҩ��¼��ID,Ŀǰ�����������ʱ�Ŵ���
  --      Ϊ�ձ�ʾ�������пɳ�����
  --��¼����:    2-�˹����ʵ�,3-�Զ����ʵ�
  --��Һ��ҩ���:    0-ҽ�����ã������ҩƷ�Ƿ������Һ��ҩ���ģ�1-��ҽ�����ã����ҩƷ�Ƿ������ҩ����
  --�ù����������ָ��������
  --����״̬_In:0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill Is
    Select a.Id, a.�۸񸸺�, a.���, a.ִ��״̬, a.��¼����, a.�շ����, a.ҽ�����, a.�շ�ϸĿid, a.����id, a.��ҳid, a.������Ŀid, a.��������id, a.���˿���id,
           a.ִ�в���id, a.���˲���id, a.����, a.����, m.��������
    From סԺ���ü�¼ A, �������� M
    Where a.No = No_In And a.��¼���� = ��¼����_In And a.��¼״̬ In (0, 1, 3) And a.�����־ = 2 And a.�շ�ϸĿid + 0 = m.����id(+)
    Order By �շ�ϸĿid, ���;

  --���α����ڴ���ҩƷ����������
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
  Cursor c_Stock(v_���_In Varchar2) Is
    Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, ����, ʵ������, ���Ч��, Ч��, ����, ����, ��������, ����id, ��Ʒ����, �ڲ�����
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From סԺ���ü�¼
                   Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And
                         �����־ = 2 And (Instr(',' || v_���_In || ',', ',' || ��� || ',') > 0 Or v_���_In Is Null))
    Order By ҩƷid, �������� Desc;

  r_Stock c_Stock%RowType;

  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ���, �۸񸸺�
    From סԺ���ü�¼
    Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3)
    Order By ���;

  v_ҽ��id ����ҽ����¼.Id%Type;
  n_����   Number;
  v_����   סԺ���ü�¼.�۸񸸺�%Type;
  v_���   Varchar2(2000);
  v_Tmp    Varchar2(4000);

  v_ҽ��ids Varchar2(4000);
  l_����    t_Numlist := t_Numlist();
  n_����    Number;
  n_����ֵ  Number;
  --�����˷Ѽ������
  v_ʣ������ Number;
  v_ʣ��Ӧ�� Number;
  v_ʣ��ʵ�� Number;
  v_ʣ��ͳ�� Number;

  v_׼������ Number;
  v_�˷Ѵ��� Number;
  v_Ӧ�ս�� Number;
  v_ʵ�ս�� Number;
  v_ͳ���� Number;
  n_�������� Number;
  v_Dec      Number;
  n_Count    Number;
  v_Curdate  Date;
  Err_Item Exception;
  v_Err_Msg        Varchar2(255);
  n_����id         ������ҳ.����id%Type;
  n_��ҳid         ������ҳ.��ҳid%Type;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);
  v_��ҩid         Varchar2(4000);

  n_δִ������ ҩƷ�շ���¼.ʵ������%Type;
  n_��ִ������ ҩƷ�շ���¼.ʵ������%Type;
Begin
  --�������ʱ,��ҩƷ�ᴫ���кŵ���������
  If Not ���_In Is Null Then
    If Instr(���_In, ':') > 0 Then
      v_Tmp := ���_In || ',';
      While Not v_Tmp Is Null Loop
        v_��� := v_��� || ',' || Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
        If Instr(Substr(v_Tmp, Instr(v_Tmp, ':') + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':') - 1), ':') > 0 Then
          v_��ҩid := v_��ҩid || ',' ||
                    Substr(v_Tmp, Instr(v_Tmp, ':', 1, 2) + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':', 1, 2) - 1);
        End If;
        v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End Loop;
      v_��� := Substr(v_���, 2);
      If v_��ҩid Is Not Null Then
        v_��ҩid := Substr(v_��ҩid, 2);
      End If;
    Else
      v_��� := ���_In;
    End If;
  End If;

  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0), Nvl(Max(����id), 0), Nvl(Max(��ҳid), 0)
  Into n_Count, n_����id, n_��ҳid
  From סԺ���ü�¼
  Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1 And �����־ = 2;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
  n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
  If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
  
    Begin
      Select ��˱�־, ״̬ Into n_��˱�־, n_סԺ״̬ From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
    Exception
      When Others Then
        n_��˱�־ := 0;
        n_סԺ״̬ := 0;
    End;
    If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then
      v_Err_Msg := '����δ���,��ֹ�Բ�����ط��õĲ���!';
      Raise Err_Item;
    End If;
  
    If n_������˷�ʽ = 1 Then
    
      If Nvl(n_��˱�־, 0) = 1 Then
        v_Err_Msg := '�ò���Ŀǰ������˷���,���ܽ��з�����ص���!';
        Raise Err_Item;
      End If;
      If Nvl(n_��˱�־, 0) = 2 Then
        v_Err_Msg := '�ò���Ŀǰ�Ѿ�����˷������,���ܽ��з�����ص���!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From סԺ���ü�¼
                Where NO = No_In And ��¼���� = ��¼����_In And �����־ = 2 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From סԺ���ü�¼
                       Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
    Raise Err_Item;
  End If;

  --ҽ�����ã��������ִ�е�ҽ��(ע����ִ�е������������,��Ϊ���� ���_IN ����������ý���������)
  If Nvl(����״̬_In, 0) = 0 Then
    --�������������̵ģ������ҽ��ִ��״̬
    Select Nvl(Count(*), 0)
    Into n_Count
    From ����ҽ������
    Where ִ��״̬ = 3 And (NO, ��¼����, ҽ��id) In
          (Select NO, ��¼����, ҽ�����
                        From סԺ���ü�¼
                        Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And ҽ����� Is Not Null And
                              (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null));
    If n_Count > 0 Then
      v_Err_Msg := 'Ҫ���ʵķ����д��ڶ�Ӧ��ҽ������ִ�е�������������ʣ�';
      Raise Err_Item;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --�ȴ�ҩƷ��Ӧ���ݼ�,��ȷ����ǰ������������,Ϊ�˴������ж�
  --�������α�������ȡ��"����� is Null"��������Ϊ�����ҩ���ܲ������ѷ�
  Open c_Stock(v_���);

  --���ñ���
  Select �Ǽ�ʱ��_In Into v_Curdate From Dual;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;

  For c_��Ŀ���� In (Select a.����
                 From ������Ϣ A, ������ҳ B
                 Where a.����id = b.����id And b.��Ŀ���� Is Not Null And
                       (b.����id, b.��ҳid) In
                       (Select Distinct ����id, ��ҳid
                        From סԺ���ü�¼
                        Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �����־ = 2)) Loop
    v_Err_Msg := '���ˡ�' || c_��Ŀ����.���� || '�� �Ѿ���������Ŀ,���ܱ����ʣ�';
    Raise Err_Item;
  End Loop;
  v_ҽ��ids := Null;
  --ѭ������ÿ�з���(������Ŀ��)
  For r_Bill In c_Bill Loop
    --����Ѿ����ڲ�����Ŀ��,���ܽ������ʴ���
    If Instr(',' || v_��� || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or v_��� Is Null Then
      Select Decode(��¼״̬, 0, 1, 0) Into n_���� From סԺ���ü�¼ Where ID = r_Bill.Id;
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into v_ʣ������, v_ʣ��Ӧ��, v_ʣ��ʵ��, v_ʣ��ͳ��
        From סԺ���ü�¼
        Where NO = No_In And ��¼���� = ��¼����_In And ��� = r_Bill.���;
        n_�������� := 0;
        If v_ʣ������ = 0 Then
          If v_��� Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ�����ʣ�';
            Raise Err_Item;
          End If;
          --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
        Else
        
          If Instr(���_In, ':') > 0 Then
            v_Tmp := ',' || ���_In;
            v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || r_Bill.��� || ':') + Length(',' || r_Bill.��� || ':'));
            v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
            If Instr(v_Tmp, ':') > 0 Then
              v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            End If;
            v_׼������ := v_Tmp;
            n_�������� := 1;
          End If;
        
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Then
            If Instr(���_In, ':') = 0 Or ���_In Is Null Then
              v_׼������ := v_ʣ������;
            End If;
          Else
            --ҽ�������ջ�ʱ,���Ŀ���û�з���,���������ʵ��ǲ�������,����Ҫ�������Ϊ׼
            If Instr(���_In, ':') = 0 Or ���_In Is Null Then
              Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
              Into v_׼������, n_Count
              From ҩƷ�շ���¼
              Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = r_Bill.Id;
            End If;
          
            --��ʣ��������׼�������������������
            --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������
            --2.��������,��ʱ�ѷ�ҩ����
            If v_׼������ = 0 Then
              If r_Bill.�շ���� = '4' Then
                If n_Count > 0 Or Nvl(r_Bill.��������, 0) = 1 Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                  Raise Err_Item;
                Else
                  v_׼������ := v_ʣ������;
                End If;
              Else
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --����סԺ���ü�¼
          If Nvl(n_����, 0) = 0 Then
            --����ʱ,ֱ�Ӹ�������,���Բ���黮��������
            --�ñ���Ŀ�ڼ�������
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into v_�˷Ѵ���
            From סԺ���ü�¼
            Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ = 2 And ��� = r_Bill.��� And �����־ = 2;
          End If;
        
          --���=ʣ����*(׼����/ʣ����)
          v_Ӧ�ս�� := Round(v_ʣ��Ӧ�� * (v_׼������ / v_ʣ������), v_Dec);
          v_ʵ�ս�� := Round(v_ʣ��ʵ�� * (v_׼������ / v_ʣ������), v_Dec);
          v_ͳ���� := Round(v_ʣ��ͳ�� * (v_׼������ / v_ʣ������), v_Dec);
          If Nvl(n_����, 0) = 1 Then
            If Nvl(n_��������, 0) = 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
              n_����ֵ := 0;
            Else
              --��������
              --���۵�,�Ƚ���ص����ݴ������ڲ�����
              n_���� := 0;
              If r_Bill.���� > 1 Then
                --�������ҩ,���ڻ��տ϶��ǻ��յĸ���,�����Ǵ���.���,��Ҫ���׼�������Ƿ������ ��
                If Trunc(v_׼������ / r_Bill.����) <> (v_׼������ / r_Bill.����) Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з���Ϊ��ҩ,�밴���������˷ѣ�';
                  Raise Err_Item;
                End If;
                n_���� := Trunc(v_׼������ / r_Bill.����);
                If Nvl(r_Bill.����, 0) - n_���� < 0 Then
                  v_׼������ := r_Bill.����;
                Else
                  v_׼������ := 0;
                End If;
              End If;
              Update סԺ���ü�¼
              Set ���� = ���� - n_����, ���� = ���� - v_׼������, Ӧ�ս�� = Nvl(Ӧ�ս��, 0) - v_Ӧ�ս��, ʵ�ս�� = Nvl(ʵ�ս��, 0) - v_ʵ�ս��,
                  �Ǽ�ʱ�� = v_Curdate, ͳ���� = Nvl(ͳ����, 0) - v_ͳ����
              Where ID = r_Bill.Id
              Returning Nvl(����, 0) * Nvl(����, 0) Into n_����ֵ;
            End If;
            If Nvl(n_����ֵ, 0) <= 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
            End If;
            If r_Bill.ҽ����� Is Not Null Then
              If Instr(',' || Nvl(v_ҽ��ids, '') || ',', ',' || r_Bill.ҽ����� || ',') = 0 Then
                v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || r_Bill.ҽ�����;
              End If;
              --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
              If v_ҽ��id Is Null Then
                v_ҽ��id := r_Bill.ҽ�����;
              End If;
            End If;
          
          End If;
        
          If Nvl(n_����, 0) = 0 Then
            --����ʱ,ֱ�Ӹ�������,���Բ���黮��������
            --�����˷Ѽ�¼
            Insert Into סԺ���ü�¼
              (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���˲���id,
               ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������,
               ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���,
               ����, ҽ��С��id)
              Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��,
                     ����, �ѱ�, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(v_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                     Decode(Sign(v_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * v_׼������), �Ӱ��־, Decode(��¼����,3,1,���ӱ�־) ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                     ��׼����, -1 * v_Ӧ�ս��, -1 * v_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * v_�˷Ѵ���, ִ��ʱ��, ����Ա���_In,
                     ����Ա����_In, ����ʱ��, v_Curdate, ������Ŀ��, ���մ���id, -1 * v_ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���, ����, ҽ��С��id
              From סԺ���ü�¼
              Where ID = r_Bill.Id;
          
            --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
            If v_ҽ��id Is Null And r_Bill.ҽ����� Is Not Null Then
              v_ҽ��id := r_Bill.ҽ�����;
            End If;
          
            Update ����������Ŀ
            Set �������� = Nvl(��������, 0) - v_׼������
            Where ����id = r_Bill.����id And ��ҳid = r_Bill.��ҳid And ��Ŀid = r_Bill.�շ�ϸĿid And Nvl(ʹ������, 0) <> 0;
          
            --�������
            Update �������
            Set ������� = Nvl(�������, 0) - v_ʵ�ս��
            Where ����id = r_Bill.����id And ���� = 2 And ���� = 1;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, ����, ����, �������, Ԥ�����)
              Values
                (r_Bill.����id, 2, 1, -1 * v_ʵ�ս��, 0);
            End If;
          
            --����δ�����
            Update ����δ�����
            Set ��� = Nvl(���, 0) - v_ʵ�ս��
            Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = Nvl(r_Bill.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_Bill.���˲���id, 0) And
                  Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
                  Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = 2;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_Bill.����id, r_Bill.��ҳid, r_Bill.���˲���id, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid, 2,
                 -1 * v_ʵ�ս��);
            End If;
          
            --���ԭ���ü�¼
            --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,���򱣳�ԭ״̬
            If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Then
              --һ�������ҩƷ�����ĵ���Ŀ,�����ڲ������ʵ����,ֻ������������������ʱ,�Ż���ֲ�������,����
              --ִ��״ֻ̬������:0.δִ��;1��ִ��;
              --������������˹����н���ִ��ǿ�Ƹ�Ϊ��2����ִ��,�����Ҫ�ڴ˴���Ϊ1��ִ��.δִ�еĲ���.
              Update סԺ���ü�¼
              Set ��¼״̬ = 3,���ӱ�־=Decode(��¼����,3,1,���ӱ�־), ִ��״̬ = Decode(Sign(v_׼������ - v_ʣ������), 0, 0, Decode(ִ��״̬, 2, 1, ִ��״̬))
              Where ID = r_Bill.Id;
            Else
              Select Nvl(Sum(Decode(�����, Null, 1, 0) * Nvl(����, 1) * ʵ������), 0),
                     Nvl(Sum(Decode(�����, Null, 0, 1) * Nvl(����, 1) * ʵ������), 0)
              Into n_δִ������, n_��ִ������
              From ҩƷ�շ���¼
              Where NO = No_In And ���� In (9, 10, 25, 26) And ����id = r_Bill.Id;
            
              Update סԺ���ü�¼
              Set ��¼״̬ = 3,
                  ִ��״̬ = Decode(Sign(v_׼������ - v_ʣ������), 0, 0,
                                 Decode(Sign(n_δִ������ - v_׼������), 1, Decode(n_��ִ������, 0, 0, 2), 1))
              Where ID = r_Bill.Id;
            End If;
          End If;
        End If;
      Else
        If v_��� Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
          Raise Err_Item;
        End If;
        --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
      End If;
    End If;
  End Loop;

  If Nvl(����״̬_In, 0) = 2 Then
    --ת��������ʱ:
    --1.ҩƷ���������õ����Ĳ�����øù���
    --2.���ۼ��˵�Ҳ������øù���
    --3.����Ҫ����ҽ����Ϣ
    For r_Bill In c_Bill Loop
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        b_Message.Zlhis_Charge_008(r_Bill.�շ����, r_Bill.Id);
      End If;
    End Loop;
    Return;
  End If;

  --��������ҩID,����ҩƷ�Ƿ�����Һ��ҩ����
  If v_��ҩid Is Null And ��Һ��ҩ���_In = 1 Then
    For v_���� In (Select ID
                 From סԺ���ü�¼
                 Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = 2 And
                       (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From ��Һ��ҩ���� A, ҩƷ�շ���¼ B
        Where a.�շ�id = b.Id And b.����id = v_����.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.���� || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '�����Ѿ�������Һ��ҩ���ĵĴ�����ҩƷ���޷�������ʣ�';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  ---------------------------------------------------------------------------------
  --ҩƷ��ش���:��Ҫ�Ƕ����������Ч.(�����ǲ���)
  For v_���� In (Select ID, ���, �շ����
               From סԺ���ü�¼
               Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = 2 And
                     (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)
               Order By �շ�ϸĿid) Loop
    --���ݷ���ID��������صĴ���
    v_׼������ := 0;
    If Instr(���_In, ':') > 0 Then
      v_Tmp := ',' || ���_In;
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || v_����.��� || ':') + Length(',' || v_����.��� || ':'));
      v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
      If Instr(v_Tmp, ':') > 0 Then
        v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      End If;
      v_׼������ := v_Tmp;
    End If;
  
    Zl_ҩƷ�շ���¼_�����˷�(v_����.Id, v_׼������, v_��ҩid, 1);
  End Loop;

  ---------------------------------------------------------------------------------
  --����ǻ���,ֱ��ɾ�����ü�¼(ҩƷ�����)
  n_Count := l_����.Count;
  --ɾ�����ۼ�¼
  Forall I In 1 .. l_����.Count
    Delete From סԺ���ü�¼ Where ID = l_����(I);

  --ɾ��֮����ͳһ�������
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        v_���� := n_Count;
      End If;
    
      Update סԺ���ü�¼
      Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, v_����)
      Where NO = No_In And ��¼���� = ��¼����_In And ��� = r_Serial.���;
    
      Update סԺ���ü�¼
      Set �������� = n_Count
      Where NO = No_In And ��¼���� = ��¼����_In And �������� = r_Serial.���;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;
  --���ŵ���ȫ������ʱ��ɾ������ҽ������
  For c_ҽ�� In (Select Distinct ҽ�����
               From סԺ���ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From סԺ���ü�¼
                  Where ��¼���� = 2 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 2 And NO = No_In;
    End If;
  End Loop;

  If v_ҽ��ids Is Not Null Then
    --ҽ������
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(1, 2, 0, No_In, v_ҽ��ids);
  Else
    Zl_ҽ������_�Ʒ�״̬_Update(1, 2, 2, No_In);
  End If;
  For r_Bill In c_Bill Loop
    --����ҩƷ������Ϣ�ŵ�Zl_ҩƷ�շ���¼_�����˷��з���
    If Nvl(r_Bill.ִ��״̬, 0) <> 1 And Instr(',4,5,6,7,', ',' || r_Bill.�շ���� || ',') = 0 Then
      b_Message.Zlhis_Charge_008(r_Bill.�շ����, r_Bill.Id);
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺ���ʼ�¼_Delete;
/

--124924:Ƚ����,2018-05-08,���ѿ���ÿһ�ſ�Ƭ�����������
Create Or Replace Procedure Zl_���ѿ����Ŀ¼_Update
(
  ���_In             ���ѿ����Ŀ¼.���%Type,
  ����_In             ���ѿ����Ŀ¼.����%Type,
  ���㷽ʽ_In         ���ѿ����Ŀ¼.���㷽ʽ%Type,
  ǰ׺�ı�_In         ���ѿ����Ŀ¼.ǰ׺�ı�%Type,
  ���ų���_In         ���ѿ����Ŀ¼.���ų���%Type,
  �Ƿ�����_In         ���ѿ����Ŀ¼.�Ƿ�����%Type,
  �Ƿ�����_In         ���ѿ����Ŀ¼.�Ƿ�����%Type,
  �Ƿ�ȫ��_In         ���ѿ����Ŀ¼.�Ƿ�����%Type,
  ����_In             ���ѿ����Ŀ¼.����%Type,
  ���볤��_In         ���ѿ����Ŀ¼.���볤��%Type,
  ���볤������_In     ���ѿ����Ŀ¼.���볤������%Type,
  �������_In         ���ѿ����Ŀ¼.�������%Type,
  ������ʽ_In         Integer,
  ��������_In         ���ѿ����Ŀ¼.��������%Type,
  ���̿��Ʒ�ʽ_In     ���ѿ����Ŀ¼.���̿��Ʒ�ʽ%Type,
  �������_In         ���ѿ����Ŀ¼.�������%Type,
  �Ƿ��ϸ����_In     ���ѿ����Ŀ¼.�Ƿ��ϸ����%Type,
  �Ƿ��ض�����_In     ���ѿ����Ŀ¼.�Ƿ��ض�����%Type,
  �Ƿ�������_In     ���ѿ����Ŀ¼.�Ƿ�������%Type,
  �Ƿ�������_In     ���ѿ����Ŀ¼.�Ƿ�������%Type,
  �Ƿ���������˿�_In ���ѿ����Ŀ¼.�Ƿ���������˿�%Type,
  Ӧ�ó���_In         ���ѿ����Ŀ¼.Ӧ�ó���%Type
) Is
  --������ʽ_In 0-����,else-�޸� 
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_������ Varchar2(200);
Begin

  If ���㷽ʽ_In Is Null Then
    v_Err_Msg := '���㷽ʽ����Ϊ�գ�';
    Raise Err_Item;
  End If;

  Begin
    Select ����
    Into v_������
    From (Select ����
           From ҽ�ƿ����
           Where ���㷽ʽ = ���㷽ʽ_In
           Union All
           Select ���� From ���ѿ����Ŀ¼ Where ��� <> ���_In And ���㷽ʽ = ���㷽ʽ_In)
    Where Rownum < 2;
  Exception
    When Others Then
      v_������ := Null;
  End;
  If v_������ Is Not Null Then
    v_Err_Msg := '���㷽ʽ��' || ���㷽ʽ_In || '���ѱ�' || v_������ || 'ʹ�ã��ظ�ʹ�û���ɲ����������ң�������ѡ��һ�ֽ��㷽ʽ��';
    Raise Err_Item;
  End If;

  If ������ʽ_In = 0 Then
    Insert Into ���ѿ����Ŀ¼
      (���, ����, ���㷽ʽ, ����, ���ƿ�, ǰ׺�ı�, ���ų���, �Ƿ�����, �Ƿ�����, �Ƿ�ȫ��, ���볤��, ���볤������, �������, ��������, ���̿��Ʒ�ʽ, �������, �Ƿ��ϸ����, �Ƿ��ض�����,
       �Ƿ�������, �Ƿ�������, �Ƿ���������˿�, Ӧ�ó���)
    Values
      (���_In, ����_In, ���㷽ʽ_In, ����_In, 1, ǰ׺�ı�_In, ���ų���_In, �Ƿ�����_In, �Ƿ�����_In, �Ƿ�ȫ��_In, ���볤��_In, ���볤������_In, �������_In,
       ��������_In, ���̿��Ʒ�ʽ_In, �������_In, �Ƿ��ϸ����_In, �Ƿ��ض�����_In, �Ƿ�������_In, �Ƿ�������_In, �Ƿ���������˿�_In, Ӧ�ó���_In);
  Else
    Update ���ѿ����Ŀ¼
    Set ���� = ����_In, ���㷽ʽ = ���㷽ʽ_In, ���� = ����_In, ǰ׺�ı� = ǰ׺�ı�_In, ���ų��� = ���ų���_In, �Ƿ����� = �Ƿ�����_In, �Ƿ����� = �Ƿ�����_In,
        �Ƿ�ȫ�� = �Ƿ�ȫ��_In, ���볤�� = ���볤��_In, ���볤������ = ���볤������_In, ������� = �������_In, �������� = ��������_In, ���̿��Ʒ�ʽ = ���̿��Ʒ�ʽ_In,
        ������� = �������_In, �Ƿ��ϸ���� = �Ƿ��ϸ����_In, �Ƿ��ض����� = �Ƿ��ض�����_In, �Ƿ������� = �Ƿ�������_In, �Ƿ������� = �Ƿ�������_In,
        �Ƿ���������˿� = �Ƿ���������˿�_In, Ӧ�ó��� = Ӧ�ó���_In
    Where ��� = ���_In;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ѿ����Ŀ¼_Update;
/

--124924:Ƚ����,2018-05-08,���ѿ���ÿһ�ſ�Ƭ�����������
Create Or Replace Procedure Zl_���ѿ���Ϣ_Update
(
  Id_In         ���ѿ���Ϣ.Id%Type,
  ������_In     ���ѿ���Ϣ.������%Type,
  �ɷ��ֵ_In   ���ѿ���Ϣ.�ɷ��ֵ%Type,
  ��Ч��_In     ���ѿ���Ϣ.��Ч��%Type,
  ����ԭ��_In   ���ѿ���Ϣ.����ԭ��%Type,
  �쿨��_In     ���ѿ���Ϣ.�쿨��%Type,
  ����id_In     ���ѿ���Ϣ.����id%Type,
  �쿨����id_In ���ѿ���Ϣ.�쿨����id%Type,
  ��ע_In       ���ѿ���Ϣ.��ע%Type,
  �������_In   ���ѿ���Ϣ.�������%Type
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_����     ���ѿ���Ϣ.����%Type;
  v_������   ���ѿ����Ŀ¼.����%Type;
  n_�ɷ��ֵ ���ѿ���Ϣ.�ɷ��ֵ%Type;
  d_����ʱ�� Date;
  d_ͣ��ʱ�� Date;
  n_���     ���ѿ���Ϣ.���%Type;
  n_������ ���ѿ���Ϣ.���%Type;
  n_Count    Number(2);
Begin
  Begin
    Select b.����, a.����, a.�ɷ��ֵ, a.����ʱ��, a.ͣ������, a.���,
           (Select Max(���) From ���ѿ���Ϣ B Where a.���� = b.���� And a.�ӿڱ�� = b.�ӿڱ��)
    Into v_������, v_����, n_�ɷ��ֵ, d_����ʱ��, d_ͣ��ʱ��, n_���, n_������
    From ���ѿ���Ϣ A, ���ѿ����Ŀ¼ B
    Where a.�ӿڱ�� = b.��� And a.Id = Id_In;
  Exception
    When Others Then
      v_Err_Msg := 'δ�ҵ�����Ϣ�������޸ģ�';
      Raise Err_Item;
  End;

  If Nvl(n_���, 0) < Nvl(n_������, 0) Then
    v_Err_Msg := '�����޸���ʷ������Ϣ(����Ϊ��' || v_���� || '��)��';
    Raise Err_Item;
  End If;

  If Nvl(d_����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
    v_Err_Msg := '����Ϊ��' || v_���� || '����' || v_������ || '�Ѿ����գ������޸ģ�';
    Raise Err_Item;
  End If;
  If Nvl(d_ͣ��ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
    v_Err_Msg := '����Ϊ��' || v_���� || '����' || v_������ || '�Ѿ�ֹͣʹ�ã��������޸ģ�';
    Raise Err_Item;
  End If;

  If Nvl(�ɷ��ֵ_In, 0) = 0 And Nvl(n_�ɷ��ֵ, 0) = 1 Then
    --��Ҫ����Ƿ����˳�ֵ��¼ 
    Select Count(1)
    Into n_Count
    From ���˿������¼��where ���ѿ�id = Id_In And ��¼���� = 2 And ��¼״̬ = 1 And Rownum < 2;
    If n_Count <> 0 Then
      v_Err_Msg := '����Ϊ��' || v_���� || '����' || v_������ || 'ԭ���ǳ�ֵ���ҷ����˳�ֵ��¼�����ܸ���Ϊ�ǳ�ֵ����';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(��Ч��_In, To_Date('3000-01-01', 'yyyy-mm-dd')) < Sysdate Then
    v_Err_Msg := '����Ϊ��' || v_���� || '����' || v_������ || '��Ч�ڲ���С�ڵ�ǰϵͳʱ�䣡';
    Raise Err_Item;
  End If;

  Update ���ѿ���Ϣ
  Set ������ = ������_In, �ɷ��ֵ = �ɷ��ֵ_In, ��Ч�� = Decode(��Ч��_In, Null, To_Date('3000-01-01', 'yyyy-mm-dd'), ��Ч��_In),
      ����ԭ�� = ����ԭ��_In, �쿨�� = �쿨��_In, �쿨����id = �쿨����id_In, ��ע = ��ע_In, ����id = ����id_In, ������� = �������_In
  Where ID = Id_In;

  --��������ֵ��Ч��,�˿���ȡ���˿����ж�����ֵ��¼ 
  --������������ǰ������ And ������� > 0 
  Update �ʻ��ɿ����
  Set ��Ч�� = Decode(��Ч��_In, Null, To_Date('3000-01-01', 'yyyy-mm-dd'), ��Ч��_In)
  Where ������� In (Select ������� From ���˿������¼ A Where a.���ѿ�id = Id_In And a.��¼���� = 1) And ���ѿ�id = Id_In And ������� > 0;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ѿ���Ϣ_Update;
/


--125261:������,2018-05-08,ת��ҽ��У�Է��ʹ����Զ�ͣ����
CREATE OR REPLACE Procedure Zl_����ҽ������_Insert
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
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� <= v_Stoptime
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

--125261:������,2018-05-08,ת��ҽ��У�Է��ʹ����Զ�ͣ����
CREATE OR REPLACE Procedure Zl_����ҽ����¼_У��
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
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� <= v_Stoptime
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
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_У��;
/

------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0010' Where ���=&n_System;
Commit;