Create Or Replace View ��Ժ�����Զ����� as
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

Create Or Replace View ��Ժ�����Զ����� as
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

create or replace view xw_ris_studyinfo as
  Select b.ҽ��ID As OrderID,a.����ID As PatientID,a.����� As OutPatientID,a.סԺ�� As InPatientID,a.������ As HealthID,
        a.���� As Name,a.�Ա� As Sex ,a.���� As Age,a.�������� As DateOfBirth,b.Ӣ���� As PYName,b.Ӱ����� As Modality,
        b.���� As StudyID,c.������Դ As Source,c.ִ�п���ID As DeptID,c.ҽ������ As MedicalOrder,
        c.����ʱ�� As ApplyTime,d.�״�ʱ�� As CheckInTime
  From ������Ϣ a,Ӱ�����¼ b,����ҽ����¼ c,����ҽ������ d
  Where b.ҽ��ID = c.Id And c.Id = d.ҽ��ID And c.����ID =a.����ID;

create or replace view xw_pacs_imagepath as
  Select a.ҽ��ID As OrderID,b.IP��ַ As ServerIP,b.FTPĿ¼ As RootPath,b.����Ŀ¼�û��� As ServerUserName,
         b.����Ŀ¼���� As ServerPassWord,Decode(a.��������,Null,'',to_Char(a.��������,'YYYYMMDD')||'\')||a.���UID||'\'||d.ͼ��Uid As ImagePathName,
         a.���UID As StudyUID,c.����UID As SeriesUID,d.ͼ��Uid As ImageUID,
         'FTP[;]'||b.IP��ַ||'[;]21[;]'||b.FTP�û���||'[;]'||b.FTP����||'[;]'||'\'||b.FTPĿ¼||'\[;]'||d.ͼ��Uid As FTPString
  From Ӱ�����¼ a,Ӱ���豸Ŀ¼ b,Ӱ�������� c,Ӱ����ͼ�� d
  Where a.λ��һ =b.�豸�� And C.���UID = A.���UID And D.����UID = C.����UID;

create or replace view xw_ris_wlm_info as
  Select '' As F_MACHINE_AET,'' As F_MACHINE_NAME , Ӱ����� As F_MODALITY_DCMTYPE ,to_char(��������,'YYYYMMDD') As F_PAT_BIRTH ,Ӣ���� As F_PAT_NAME ,Ӣ���� As F_PAT_NAME_EN
         ,'' As F_PAT_ADDRESS,���� As F_PAT_NO,'' As F_PAT_OT_ID,'' As F_PAT_LOCATION,'' As F_ADD_HISTORY,'' As F_PAT_REGION,'' As F_MEDICAL_ALERTS,'' As F_CONTRAST,'' As F_PLACE_NO
         ,decode(�Ա�,'��','M','Ů','F','O') As F_SEX,���� As F_WEIGHT,��� As F_HEIGHT,'' As F_PERFORM_DOC,'' As F_REQUEST_DOC,'' As F_DIAGNOSES,'' As F_STU_REASON,'' As F_STU_COMMENT
         ,'' As F_MEN_DATE,'' As F_LATERALITY,to_char(b.�״�ʱ��,'YYYYMMDD') As F_STU_DATE_DCM,a.ҽ��ID As F_STU_ID,a.ҽ��ID As F_STU_NO,to_char(b.�״�ʱ��,'hh24:mi:ss') As F_STU_TIME_DCM
         ,a.ҽ��ID || '.' || a.���ͺ� As F_STU_UID,b.ִ�м� as F_ROOM_NAME, c.���� as F_DEPT_NAME 
  From  Ӱ�����¼ a ,����ҽ������ b,���ű� c
  Where a.ҽ��ID=b.ҽ��ID And a.���ͺ� = b.���ͺ� And b.ִ�в���id=c.Id And  b.ִ��״̬=3 And b.ִ�й���=2 And a.���UID IS Null AND b.�״�ʱ��>=SysDate-10;

create or replace view xw_ris_wlm_info_cn as
  Select '' As F_MACHINE_AET,'' As F_MACHINE_NAME , Ӱ����� As F_MODALITY_DCMTYPE ,to_char(��������,'YYYYMMDD') As F_PAT_BIRTH ,���� As F_PAT_NAME ,Ӣ���� As F_PAT_NAME_EN
         ,'' As F_PAT_ADDRESS,���� As F_PAT_NO,'' As F_PAT_OT_ID,'' As F_PAT_LOCATION,'' As F_ADD_HISTORY,'' As F_PAT_REGION,'' As F_MEDICAL_ALERTS,'' As F_CONTRAST,'' As F_PLACE_NO
         ,decode(�Ա�,'��','M','Ů','F','O') As F_SEX,���� As F_WEIGHT,��� As F_HEIGHT,'' As F_PERFORM_DOC,'' As F_REQUEST_DOC,'' As F_DIAGNOSES,'' As F_STU_REASON,'' As F_STU_COMMENT
         ,'' As F_MEN_DATE,'' As F_LATERALITY,to_char(b.�״�ʱ��,'YYYYMMDD') As F_STU_DATE_DCM,a.ҽ��ID As F_STU_ID,a.ҽ��ID As F_STU_NO,to_char(b.�״�ʱ��,'hh24:mi:ss') As F_STU_TIME_DCM
         ,a.ҽ��ID || '.' || a.���ͺ� As F_STU_UID,b.ִ�м� as F_ROOM_NAME, c.���� as F_DEPT_NAME
  From  Ӱ�����¼ a ,����ҽ������ b,���ű� c
  Where a.ҽ��ID=b.ҽ��ID And a.���ͺ� = b.���ͺ�  And b.ִ�в���id=c.Id And  b.ִ��״̬=3 And b.ִ�й���=2 And a.���UID IS Null AND b.�״�ʱ��>=SysDate-10;


CREATE OR REPLACE VIEW �շ���� AS 
    SELECT ����,���� AS ���,���� AS ˵��,�̶� AS ϵͳ��־,0 AS �����༭
    FROM �շ���Ŀ���;

CREATE OR REPLACE VIEW �շ�ϸĿ AS
SELECT ���,ID, NULL AS �ϼ�id,1 AS ĩ��,����,����,���||'��'||���� AS ���,���㵥λ,˵��,
        ��������,�������,0 AS �����༭,���ηѱ�,�Ƿ���,�Ӱ�Ӽ�,����ժҪ,
        decode(ִ�п���,1,2,2,2,3,3,4,1,0) As ִ�п���,��ʶ����,��ʶ����,����ʱ��,����ʱ��
FROM �շ���ĿĿ¼;

CREATE OR REPLACE VIEW �շѱ��� AS 
    SELECT �շ�ϸĿid,����,����
    FROM �շ���Ŀ����;

CREATE OR REPLACE VIEW �շ�ִ�в��� AS 
    SELECT �շ�ϸĿid,ִ�п���id as ִ�в���id
    FROM �շ�ִ�п���;

CREATE OR REPLACE VIEW �Һ���Ŀ AS 
    SELECT I.ID AS ���, I.����, I.����, I.���㵥λ, N.����, I.��Ŀ���� AS ������, I.˵��, I.����ʱ��, I.����ʱ��
    FROM �շ���ĿĿ¼ I, �շ���Ŀ���� N
    WHERE I.ID=N.�շ�ϸĿid And I.���='1' And N.����=1 And N.����=1;

CREATE OR REPLACE VIEW ��λ�ȼ� AS 
    SELECT I.ID AS ���, I.����, I.����, I.���㵥λ, N.����, I.˵��, I.����ʱ��, I.����ʱ��
    FROM �շ���ĿĿ¼ I, �շ���Ŀ���� N
    WHERE I.ID=N.�շ�ϸĿid And I.���='J' And N.����=1 And N.����=1;

CREATE OR REPLACE VIEW ����ȼ� AS 
    SELECT I.ID AS ���, I.����, I.����, I.���㵥λ, N.����, I.��Ŀ����-1 AS ��������, I.˵��, I.����ʱ��, I.����ʱ��
    FROM �շ���ĿĿ¼ I, �շ���Ŀ���� N
    WHERE I.ID=N.�շ�ϸĿid And I.���='H' And I.��Ŀ����>=1 And N.����=1 And N.����=1;
    
CREATE OR REPLACE VIEW ҩƷ���ʷ��� AS 
    SELECT decode(����,'5','1','6','2','3') AS ����,����, ����
    FROM ������Ŀ���
    WHERE ���� in ('5','6','7');

CREATE OR REPLACE VIEW ҩƷ��;���� AS 
    SELECT decode(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') AS ����, ID, ����, ����, ����, �ϼ�id,1 AS ĩ��
    FROM ���Ʒ���Ŀ¼
    WHERE ���� in (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'));

CREATE OR REPLACE VIEW ҩƷ��Ϣ AS 
    SELECT decode(I.���,'5','����ҩ','6','�г�ҩ','�в�ҩ') AS ���ʷ���, S.ҩ��ID, null as ҩ��id,I.����ID AS ��;����ID,
        K.���� AS ����,I.����, I.���� AS ͨ������, I.���㵥λ AS ������λ, S.�������, S.��Դ���, S.��ֵ����, S.��ҩ�ݴ�, 
        S.����ְ��, S.����ҩ��, S.�Ƿ���ҩ, S.�Ƿ�Ƥ��, S.�Ƿ�ԭ��, S.��������, S.ҩƷ����, I.����ʱ��, I.����ʱ��
    FROM ������ĿĿ¼ I, ҩƷ���� S,ҩƷ���� K
    WHERE I.ID=S.ҩ��ID AND S.ҩƷ����=K.����(+) And I.��� In ('5','6','7');

CREATE OR REPLACE VIEW ҩƷĿ¼ AS 
    SELECT S.ҩƷid, S.ҩ��ID, I.����, I.����, I.���, I.����, I.���㵥λ AS �ۼ۵�λ, 
            S.����ϵ��, S.���ﵥλ, S.�����װ, S.סԺ��λ, S.סԺ��װ, S.ҩ�ⵥλ, S.ҩ���װ, 
            S.���Ч��, S.ҩƷ��Դ, S.Э��ҩƷ, S.����ҩƷ, S.��׼�ĺ�, S.��ʶ��, 
            S.ҩ�ۼ���, I.�Ƿ���, S.ָ��������, S.ָ�����ۼ�, S.ָ�������, S.����, S.סԺ�ɷ����, S.����ɷ����,
            S.ҩ����� AS ��������, S.ҩ������ AS ҩ����������, I.��������, decode(I.�������,1,'100',2,'010',3,'110','000') AS �������,
            S.�б�ҩƷ,S.���������,S.GMP��֤,I.����ʱ��, I.����ʱ��
    FROM �շ���ĿĿ¼ I, ҩƷ���  S
    WHERE I.ID=S.ҩƷid And I.��� In ('5','6','7');

CREATE OR REPLACE VIEW ҩƷ������� AS 
    SELECT R.����, R.��ĿID AS ҩ��ID, R.����
    FROM ���ƻ�����Ŀ  R, ������ĿĿ¼  I
    WHERE R.��ĿID=I.ID And I.��� In ('5','6','7');

Create Or Replace View ҩƷ���� As 
    Select T.ҩ��id,N.����,N.����,decode(����,3,N.�շ�ϸĿid,Null) As ҩƷid,
           decode(N.����,3,2,1) As ����
    From �շ���Ŀ���� N,ҩƷ��� T
    Where N.�շ�ϸĿid=T.ҩƷid And N.����<>2;

Create Or Replace View �ҺŲ��� As
  Select Distinct NO As �Һŵ���, ����id, ����, �Ա�, ����, �շ�ϸĿid As �Һ���Ŀ, �Ӱ��־ As ����, �Ǽ�ʱ�� As ����, ִ�в���id As ����id, ��ҩ���� As ����,
                  ִ���� As ҽ��, ִ��״̬ As ״̬
  From ������ü�¼
  Where ��¼���� = 4 And ��¼״̬ = 1 And �շ���� = '1' And ����id Is Not Null And
        �Ǽ�ʱ�� > (Select Trunc(Sysdate) - To_Number(����ֵ)
                From zlParameters
                Where ϵͳ = (Select ���
                            From zlSystems
                            Where Upper(������) = (Select Username From All_Users Where User_Id = Userenv('SchemaID')) And
                                  Trunc(��� / 100) = 1) And ģ�� Is Null And Nvl(˽��, 0) = 0 And ������ = 21);

--�ṩ���������µĲ��˷�����ϸ��
Create Or Replace View ���˷�����ϸ As 
Select L.����id,
	   L.Id As ˳���,
	   S.��Ŀ����,
	   S.��Ŀ����,
	   L.��׼����,
	   L.�Ǽ�ʱ�� As �շ�����
  From  (Select Id, ����ID,�շ�ϸĿID,��׼����,�Ǽ�ʱ�� From  ������ü�¼ 
	     Union ALL Select Id, ����ID,�շ�ϸĿID,��׼����,�Ǽ�ʱ�� From  סԺ���ü�¼ 
	     )  L, �շ���ĿĿ¼ I, ��׼ҽ�۹淶 S
 Where L.�շ�ϸĿid = I.Id And I.��ʶ���� = S.��Ŀ����(+) And
	   I.��� Not In ('4', '5', '6', '7');

--��ҽ���ӿ�Ҫ��ZLHIS9���ݶ�����
Create Or Replace View ������ As 
Select ����id, ��ҳId,����id,������� As ������Ϣ,�������, 
   ��Ժ���,��ϴ���,�������, �Ƿ�δ��, �Ƿ�����, ¼�����,�������
From ������ϼ�¼ Where ��¼��Դ=2;

--Ϊ��������ǰ�汾����
Create OR REPLACE View ҽ���˶Ա� AS Select ����ID,���㷽ʽ,��� From ���ս�����ϸ Where ��־=1;

CREATE OR REPLACE VIEW �����ʻ� AS 
	SELECT A.����ID,B.* FROM ҽ�����˹����� A,ҽ�����˵��� B 
	WHERE A.����=B.���� AND A.����=B.���� AND A.ҽ����=B.ҽ���� AND A.��־=1; 

--���Ӳ������鵵
create or replace view �������ֱ�׼��ͼ as
select decode(T.�ϼ����,null,���,T.�ϼ����) as �ϼ����, decode(T.���,null,T.ID,T.���) as ���,T.ID,T.�ϼ�ID,T.����ID,T.��Ŀ,T.��׼��ֵ,T.����Ҫ��,T.ȱ������,T.�۷ֱ�׼,decode(T.�������,0,'��','��') as ����,T.����ȼ�
from
(
  select B.�ϼ����,A.���,A.����ID,
  A.ID,
  A.�ϼ�ID,
  decode(A.�������,0,decode(A.�ϼ�ID,Null,A.����,B.����),A.����) as ��Ŀ,
  decode(A.�������,0,decode(A.�ϼ�ID,Null,A.��׼��ֵ,B.��׼��ֵ),B.��׼��ֵ) as ��׼��ֵ,
  decode(A.�������,0,decode(A.�ϼ�ID,Null,A.����,B.����),A.����) as ����Ҫ��,
  decode(A.�ϼ�ID,Null,'',A.����) as ȱ������,
  DECODE(A.ȱ�ݵȼ�,NULL,decode(sign(A.��׼��ֵ-1),-1,To_CHAR(A.��׼��ֵ,'0.9'),To_Char(A.��׼��ֵ))||decode(A.���ֵ�λ,NULL,'','/'||A.���ֵ�λ),A.ȱ�ݵȼ�) as �۷ֱ�׼,
  A.�������,
  A.����ȼ�
  from
      (
          select AA.���,AA.ID,AA.����ID,AA.�ϼ�ID,AA.����,AA.����,AA.��׼��ֵ,AA.ȱ�ݵȼ�,AA.���ֵ�λ,AA.����ȼ�,count(BB.ID) as �������
          from �������ֱ�׼ AA,�������ֱ�׼ BB
          where AA.ID=BB.�ϼ�ID(+)
          group by AA.���,AA.ID,AA.����ID,AA.�ϼ�ID,AA.����,AA.����,AA.��׼��ֵ,AA.ȱ�ݵȼ�,AA.���ֵ�λ,AA.����ȼ�
      ) A,
      (
          select ��� as �ϼ����,ID,����,��׼��ֵ,���� from �������ֱ�׼
      ) B
  where A.�ϼ�ID=B.ID(+)
) T
order by decode(T.�ϼ����,null,���,T.�ϼ����),decode(T.���,null,T.ID,T.���);

create or replace view ��������������ͼ as
Select   Tb.����, Tb.�Ա�, Ta."����ID",Ta."��ҳID",Ta."סԺ��",Ta."��Ժ����",Ta."��Ժ����",Ta."��Ժ����",Ta."��Ժ����",Ta."����ҽʦ",Ta."���λ�ʿ",Ta."סԺҽʦ",Ta."��Ŀ����",Ta."���ID",Ta."����ID",Ta."�ܷ�",Ta."�ȼ�",Ta."������",Ta."����ʱ��",Ta."�����",Ta."���ʱ��",Ta."�����޸�",Ta."��ע",Ta."��������" 
From (Select T1.����id, T1.��ҳid,T1.סԺ��, T1.��Ժ����, T1.��Ժ����, T2.���� As ��Ժ����, T3.���� As ��Ժ����, T1.����ҽʦ, 
              T1.���λ�ʿ, T1.סԺҽʦ, T1.��Ŀ����, T1.���id, T1.����id, T1.�ܷ�, T1.�ȼ�, T1.������, 
              To_Char(T1.����ʱ��, 'YYYY-MM-DD') As ����ʱ��, T1.�����, To_Char(T1.���ʱ��, 'YYYY-MM-DD') As ���ʱ��, 
              T1.�����޸�, T1.��ע ,T1.��������
       From (Select A.����id, A.��ҳid, A.��Ժ����id, A.��Ժ����id, A.��Ժ����, A.��Ժ����, A.����ҽʦ, A.���λ�ʿ, 
                     A.סԺҽʦ, A.��Ŀ����, B.ID As ���id, B.����id, B.�ܷ�, B.�ȼ�, B.������, B.����ʱ��, B.�����, 
                     B.���ʱ��, B.�����޸�, B.��ע,B.��������,A.סԺ�� 
              From ������ҳ A, �������ֽ�� B 
              Where A.����id = B.����id(+) And A.��ҳid = B.��ҳid(+)) T1, ���ű� T2, ���ű� T3 
       Where T1.��Ժ����id = T2.ID And T1.��Ժ����id = T3.ID) Ta, ������Ϣ Tb 
Where Ta.����id = Tb.����id;