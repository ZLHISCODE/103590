----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--139433:������,2019-04-09,����΢���ﱨ���޸�
alter table ҽ���������� add �Ƿ��ֹ��ӡ number(1);

--120692:����,2019-04-03,�����¼֧�ּ�����Ŀ����
create table �������ݵ��붨��
(
��� number(1),
���� varchar2(100),
��ʽ varchar2(500)
)tablespace zl9BaseItem;
alter table �������ݵ��붨�� add constraint �������ݵ��붨��_PK primary key (���) using index tablespace zl9Indexhis;

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Alter Table ������ü�¼ Add ���˲���id Number(18);


------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--120692:����,2019-04-03,�����¼֧�ּ�����Ŀ����
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'�������ݵ��붨��','ZL9BASEITEM','A2');

--110283:����,2019-04-02,����ϵͳ����ָ�����ϲ���ʱ����ʾ�޿�������������Ƿ���ʾ�޿������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -Null, 0, 0, 0, 0, 0, 0, 316, 'ָ�����ϲ���ʱ����ʾ�޿������', Null, '0',
         '�������շѣ����ۡ�����)��סԺ���ʡ����ۻ�ҽ��վ���ѵ�¼����������ѡ��ʱ�����������ȱʡ���ϲ��ŵģ�����ѡ��������ʾ�޿����������ϡ�', '1-ָ����ȱʡ�ⷿʱ��ʾ�п�������;0-������ ',
         '��������Ҫ��ϲ���"ȱʡ���ϲ���"���ʹ�ã������ˡ�ȱʡ���ϲ��š�ʱ������������Ч��', '�����ڵ�ǰ�����޿������ʾ���ĵ������', Null
  From Dual;

--115787:Ƚ����,2019-04-01,����һ��˽��ģ���������������������Ʋ�ѯ���˷�����Ϣʱ�Ƿ��ȡ�������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1139, 1, 0, 0, 0, 0, 1, 21, '�����������', Null, '1',
         '���ò��˷��ò�ѯģ��鿴���˷�����Ϣʱ���Ƿ���������������', '0-������������ã�1-�����������', '', '�����ڵ��ò��˷��ò�ѯģ��鿴���˷�����Ϣʱֻ�鿴סԺ������Ϣ�����', Null
  From Dual;


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--120692:����,2019-04-03,�����¼֧�ּ�����Ŀ����
Insert Into Zlprogprivs(ϵͳ, ���, ����, ������, ����, Ȩ��)Values(&n_System, 1255, '����', User, '�������ݵ��붨��', 'SELECT');
Insert Into Zlprogprivs(ϵͳ, ���, ����, ������, ����, Ȩ��) Values (&n_System, 1255, '�����¼�Ǽ�', User, 'Zl_�������ݵ��붨��_Update', 'EXECUTE');


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--137062:Ƚ����,2019-04-08,��ȡHIS�������ݣ����ؽ�����ϸ
Create Or Replace Procedure Zl_Third_Getsettlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡHIS��������
  --���:Xml_In:
  --<IN>
  -- <BRID></BRID>       //����ID
  -- <XM></XM>          //����
  -- <SFZH></SFZH>       //���֤��
  -- <ZYID></ZYID>         //��ҳID
  -- <JSLX></JSLX>       //�������͡�1-����,2-סԺ���̶���2
  -- <JSKLB></JSKLB>       //���㿨���
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  --<JBXX>              //������Ϣ
  --   <XM></XM>           //����
  --   <XB></XB>           //�Ա�
  --   <NL></NL>         //����
  --   <ZYH></ ZYH>        //סԺ��
  --   <ZYKS></ ZYKS>          //סԺ����
  --   <KSID></KSID>         //����ID
  --   <ZZYS></ ZZYS>          //����ҽ��
  --   <RYSJ></ RYSJ>          //��Ժʱ��
  --   <CYSJ></ CYSJ >         //��Ժʱ��
  --   <JZSJ></JZSJ>         //����ʱ��(δ����Ϊ��)
  --   <DJH></DJH>         //���ݺ�(δ����Ϊ��)
  --   <JSZFY></JSZFY>         //�����ܷ���
  --</JBXX>
  --<YJKLIST>              //���Ԥ�ɿ��
  --   <ITEM>
  --     <DJH><DJH>        //Ԥ����ݺ�
  --     <JSFS></JSFS>     //���㷽ʽ��Ϊ���ƣ�����ʲô��ȡʲô��
  --     <JE></JE>           //Ԥ�ɿ���
  --     <JYLSH></JYLSH>       //������ˮ�ţ����ڳ���ʹ�ã�
  --     <JYSM></JYSM>        //����˵��
  --     <SFJSK></SFJSK>       //�Ƿ���㿨��1-�ǣ�0-��������ɴ���Ŀ����ɷѣ�����1�����򷵻�0
  --     <ZFZT></ZFZT>        //֧��״̬��0-��֧����1-����֧��
  --   </ITEM>
  --</YJKLIST >
  --<TBQK>               //�˲����
  --   <TBLX></TBLX>         //�˲�����(1:���˲��2:ҽԺ�˿�)
  --   <TBJE></TBJE>         //�˲����
  --</TBQK>
  --<JSMX>                 //������ϸ
  --  <ITEM>
  --    <JSFS></JSFS>         //���㷽ʽ
  --    <JSJE></JSJE>         //������
  --    <SFYB></SFYB>         //�Ƿ�ҽ�����㷽ʽ,1-�ǣ�0-��
  --    <SFYJK></SFYJK>         //�Ƿ�Ԥ����,1-�ǣ�0-��
  --  </ITEM>
  --</JSMX>
  -- <ERROR><MSG></MSG></ERROR>    //���ִ���ʱ���ؾ���ԭ��error�ڵ�Ϊ�ձ�ʾ�ɹ�
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_����id     ������Ϣ.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_��ҳid     ������Ϣ.��ҳid%Type;
  n_��������   Number(3);
  n_�����id   ҽ�ƿ����.Id%Type;
  v_���㿨��� Varchar2(200);
  n_�Ƿ����   Number(3); -- 1-δ����,0-����
  n_���ʽ��   סԺ���ü�¼.���ʽ��%Type;
  v_Temp       Varchar2(32767); --��ʱXML
  v_Subtemp    Varchar2(32767);
  n_�˲����   ����Ԥ����¼.��Ԥ��%Type;
  n_�������   ����Ԥ����¼.���%Type;
  n_����id     ����Ԥ����¼.����id%Type;

  n_Number  Number(2);
  x_Templet Xmltype; --ģ��XML
  x_Temp    Xmltype;

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/ZYID'), Nvl(Extractvalue(Value(A), 'IN/JSLX'), 2),
         Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_����id, n_��ҳid, n_��������, v_���㿨���, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_�������� = 1 And Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����!';
    Raise Err_Item;
  End If;

  Select Decode(Translate(Nvl(v_���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0) Into n_Number From Dual;
  If Nvl(n_Number, 0) = 1 Then
    Select Max(ID) Into n_�����id From ҽ�ƿ���� Where ID = To_Number(v_���㿨���);
  Else
    Select Max(ID) Into n_�����id From ҽ�ƿ���� Where ���� = v_���㿨���;
  End If;
  If Nvl(n_�����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ�ϴ���Ľ��㿨,����!';
    Raise Err_Item;
  End If;

  If n_�������� = 2 Then
    Select Count(1)
    Into n_�Ƿ����
    From (Select 1
           From סԺ���ü�¼
           Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
           Group By ����id, ��ҳid
           Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0)
    Where Rownum < 2;
  
    If n_�Ƿ���� = 0 Then
      --����,��ȡ��������
      For r_���� In (Select ����, �Ա�, ����, סԺ��, סԺ����, ����id, ����ҽ��, To_Char(��Ժʱ��, 'yyyy-mm-dd') As ��Ժʱ��,
                          To_Char(��Ժʱ��, 'yyyy-mm-dd') As ��Ժʱ��, To_Char(����ʱ��, 'yyyy-mm-dd') As ����ʱ��, ���ݺ�, �����ܷ���, ����id
                   From (Select c.����, c.�Ա�, c.����, c.סԺ��, d.���� As סԺ����, c.��Ժ����id As ����id, c.סԺҽʦ As ����ҽ��, c.��Ժ���� As ��Ժʱ��,
                                 c.��Ժ���� As ��Ժʱ��, a.�շ�ʱ�� As ����ʱ��, a.No As ���ݺ�, a.���ʽ�� As �����ܷ���, a.Id As ����id
                          From ���˽��ʼ�¼ A, ������ҳ C, ���ű� D
                          Where a.��¼״̬ = 1 And Nvl(a.����״̬, 0) In (0, 2) And a.����id = c.����id And a.����id = n_����id And
                                a.��ҳid = n_��ҳid And a.��ҳid = c.��ҳid And c.��Ժ����id = d.Id(+) And Exists
                           (Select 1 From ����Ԥ����¼ Where ����id = a.Id And �����id = n_�����id)
                          Order By ����ʱ�� Desc)
                   Where Rownum < 2) Loop
        v_Temp := '<XM>' || r_����.���� || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_����.�Ա� || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_����.���� || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_����.סԺ�� || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_����.סԺ���� || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_����.����id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_����.����ҽ�� || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_����.��Ժʱ�� || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_����.��Ժʱ�� || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || r_����.����ʱ�� || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || r_����.���ݺ� || '</DJH>';
        v_Temp := v_Temp || '<JSZFY>' || r_����.�����ܷ��� || '</JSZFY>';
        v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        n_����id := r_����.����id;
      End Loop;
    
      If n_����id Is Null Then
        v_Err_Msg := '�ò���û�н�������!';
        Raise Err_Item;
      End If;
    
      --���Ԥ�ɿ��
      v_Temp := '';
      For r_Ԥ�� In (Select NO As ���ݺ�, ���㷽ʽ, Sum(��Ԥ��) As ���, �����id, ������ˮ��, ����˵��, Decode(Nvl(У�Ա�־, 0), 0, 0, 1) As ֧��״̬
                   From ����Ԥ����¼
                   Where ����id = n_����id And Mod(��¼����, 10) = 1
                   Group By NO, ���㷽ʽ, �����id, ������ˮ��, ����˵��, Nvl(У�Ա�־, 0)
                   Order By ���ݺ� Desc) Loop
        v_Temp := '<DJH>' || r_Ԥ��.���ݺ� || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_Ԥ��.���㷽ʽ || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_Ԥ��.��� || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_Ԥ��.������ˮ�� || '</JYLSH>';
        v_Temp := v_Temp || '<JYSM>' || r_Ԥ��.����˵�� || '</JYSM>';
        If n_�����id = r_Ԥ��.�����id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp    := v_Temp || '<ZFZT>' || r_Ԥ��.֧��״̬ || '</ZFZT>';
        v_Temp    := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp := v_Subtemp || v_Temp;
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
    
      --�˲����
      Select Nvl(Sum(��Ԥ��), 0)
      Into n_�˲����
      From ����Ԥ����¼
      Where ����id = n_����id And Mod(��¼����, 10) = 2 And Nvl(У�Ա�־, 0) = 0;
      If n_�˲���� < 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(n_�˲����) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      --������ϸ
      Select Xmlelement("JSMX",
                         Xmlagg(Xmlelement("ITEM",
                                            Xmlforest(���㷽ʽ As "JSFS", ������ As "JSJE", �Ƿ�ҽ�� As "SFYB", �Ƿ�Ԥ���� As "SFYJK"))))
      Into x_Temp
      From (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ������, Decode(Mod(a.��¼����, 10), 1, 1, 0) As �Ƿ�Ԥ����,
                    Max(Decode(Nvl(b.����, 0), 3, 1, 4, 1, 0)) As �Ƿ�ҽ��
             From ����Ԥ����¼ A, ���㷽ʽ B
             Where a.���㷽ʽ = b.����(+) And a.����id = n_����id
             Group By Decode(Mod(a.��¼����, 10), 1, 1, 0), a.���㷽ʽ);
      Select Appendchildxml(x_Templet, '/OUTPUT', x_Temp) Into x_Templet From Dual;
    Else
      --δ���壬��ȡδ������
      For r_Info In (Select c.����, c.�Ա�, c.����, c.סԺ��, d.���� As סԺ����, c.��Ժ����id As ����id, c.סԺҽʦ As ����ҽ��,
                            To_Char(c.��Ժ����, 'yyyy-mm-dd') As ��Ժʱ��, To_Char(c.��Ժ����, 'yyyy-mm-dd') As ��Ժʱ��
                     From ������ҳ C, ���ű� D
                     Where c.����id = n_����id And c.��Ժ����id = d.Id(+) And c.��ҳid = n_��ҳid And Rownum < 2) Loop
        v_Temp := '<XM>' || r_Info.���� || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_Info.�Ա� || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_Info.���� || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_Info.סԺ�� || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_Info.סԺ���� || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_Info.����id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_Info.����ҽ�� || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_Info.��Ժʱ�� || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_Info.��Ժʱ�� || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || '' || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || '' || '</DJH>';
      End Loop;
    
      Begin
        Select Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0))
        Into n_���ʽ��
        From סԺ���ü�¼
        Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1;
      Exception
        When Others Then
          n_���ʽ�� := 0;
      End;
      v_Temp := v_Temp || '<JSZFY>' || n_���ʽ�� || '</JSZFY>';
      v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      --���Ԥ�ɿ��
      v_Subtemp := '';
      For r_Ԥ�� In (Select a.No As ���ݺ�, a.���㷽ʽ, Sum(Nvl(a.���, 0)) - Sum(Nvl(a.��Ԥ��, 0)) As ���, a.�����id, a.������ˮ��, a.����˵��,
                          Decode(Nvl(a.У�Ա�־, 0), 0, 0, 1) As ֧��״̬
                   From ����Ԥ����¼ A, ���㷽ʽ B
                   Where a.���㷽ʽ = b.����(+) And a.����id = n_����id And Mod(a.��¼����, 10) = 1 And Nvl(a.Ԥ�����, 2) = 2 And
                         (a.��ҳid = n_��ҳid Or a.��ҳid Is Null) And Nvl(b.����, 0) <> 5
                   Group By a.No, a.���㷽ʽ, a.�����id, a.������ˮ��, a.����˵��, Nvl(a.У�Ա�־, 0)
                   Having Sum(Nvl(a.���, 0)) - Sum(Nvl(a.��Ԥ��, 0)) <> 0
                   Order By ���ݺ�) Loop
        v_Temp := '<DJH>' || r_Ԥ��.���ݺ� || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_Ԥ��.���㷽ʽ || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_Ԥ��.��� || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_Ԥ��.������ˮ�� || '</JYLSH>';
        v_Temp := v_Temp || '<JYSM>' || r_Ԥ��.����˵�� || '</JYSM>';
        If n_�����id = r_Ԥ��.�����id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp     := v_Temp || '<ZFZT>' || r_Ԥ��.֧��״̬ || '</ZFZT>';
        v_Temp     := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp  := v_Subtemp || v_Temp;
        n_������� := Nvl(n_�������, 0) + Nvl(r_Ԥ��.���, 0);
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
    
      --�˲����
      If Nvl(n_�������, 0) - Nvl(n_���ʽ��, 0) > 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(Nvl(n_�������, 0) - Nvl(n_���ʽ��, 0)) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getsettlement;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_����תסԺ_������ת��
(
  No_In         ���ò����¼.No%Type,
  ���ó���id_In ����Ԥ����¼.����id%Type,
  �������id_In ����Ԥ����¼.����id%Type,
  �������_In   ����Ԥ����¼.�������%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
  ��Ժ����id_In ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����_In     ����Ԥ����¼.��Ԥ��%Type := Null
) As
  --���ܣ��Է��ò�������������ý���תסԺ���ô���
  --��Σ�
  --  ���㷽ʽ_In ��Ϊ�գ���ʾ���г�Ԥ����ķ�ҽ�����ȫ����Ϊָ���Ľ��㷽ʽ��
  --              Ϊ�գ���ʾ���г�Ԥ����ķ�ҽ�����ȫ��תΪסԺԤ����
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_����ֵ  ����Ԥ����¼.��Ԥ��%Type;

  n_��id   ����ɿ����.Id%Type;
  v_���� ���㷽ʽ.����%Type;
  n_���� ����Ԥ����¼.��Ԥ��%Type;
  n_Dec    Number; --���С��λ�� 

  v_Nos    Varchar2(4000);
  n_����id ����Ԥ����¼.����id%Type;

  n_���˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_δ�˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_���㷽ʽ Varchar2(4000);
  v_Ԥ��no   ����Ԥ����¼.No%Type;

  --����Ԥ�����
  Procedure ����Ԥ����¼_Insert
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    ���_In       ����Ԥ����¼.���%Type,
    ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    ����_In       ����Ԥ����¼.����%Type := Null,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
  ) As
    v_Ԥ��no ����Ԥ����¼.No%Type;
    n_����ֵ ����Ԥ����¼.���%Type;
  Begin
    If Nvl(���_In, 0) = 0 Or ���㷽ʽ_In Is Null Then
      Return;
    End If;
  
    --һ��ͨ��ÿһ�ʶ�����һ��Ԥ�����¼
    --������ͬһ�ֽ��㷽ʽֻ����һ��Ԥ�����¼
    Update ����Ԥ����¼
    Set ��� = Nvl(���, 0) + ���_In
    Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �տ�ʱ��_In And ����id + 0 = ����id_In And ���㷽ʽ = ���㷽ʽ_In And Nvl(�����id, 0) = 0;
    If Sql%RowCount = 0 Or Nvl(�����id_In, 0) <> 0 Then
      v_Ԥ��no := Nextno(11);
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����,
         �����id, ����, ����˵��, ������ˮ��, �������)
      Values
        (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, ����id_In, ��ҳid_In, ��Ժ����id_In, ���_In, ���㷽ʽ_In, �տ�ʱ��_In, Null, Null, Null,
         ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, 2, �����id_In, ����_In, ����˵��_In, ������ˮ��_In, �������_In);
    End If;
  
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ���� = 1 And ����id = ����id_In And ���� = 2
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (����id_In, 1, 2, ���_In, 0);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
  End;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --����
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'û�з��������㷽ʽ�������Ƿ���ȷ���ã�';
      Raise Err_Item;
  End;
  n_���� := Nvl(����_In, 0);

  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  Select f_List2str(Cast(Collect(a.No) As t_Strlist), ',', 1), Max(a.����id)
  Into v_Nos, n_����id
  From ������ü�¼ A, ���ò����¼ B
  Where a.����id = b.�շѽ���id And b.��¼���� = 1 And b.���ӱ�־ = 0 And b.No = No_In;
  If v_Nos Is Null Then
    v_Err_Msg := 'δ�ҵ�ԭҽ�����������ݣ�����ת��ʧ��!';
    Raise Err_Item;
  End If;

  --1.���·�����˼�¼ 
  Update ������˼�¼
  Set ��¼״̬ = 2
  Where ���� = 1 And ����id In (Select /*+cardinality(b,10)*/
                             a.Id
                            From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
                            Where a.No = b.No And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3));

  --2.����������ü�¼ 
  Update ������ü�¼
  Set ��¼״̬ = 3
  Where Mod(��¼����, 10) = 1 And ��¼״̬ = 1 And NO In (Select Column_Value As NO From Table(f_Str2list(v_Nos)));

  For c_���� In (Select /*+cardinality(b,10)*/
                a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����, a.��ʶ��, a.���ʽ, a.���˿���id, a.�ѱ�,
                a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.��ҩ����, Sum(Nvl(a.����, 1) * a.����) As ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ,
                a.��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, a.������, a.��������id, a.������, a.����ʱ��, a.ִ�в���id, a.ִ����,
                Min(Decode(a.��¼״̬, 2, a.ִ��״̬, 0)) - 1 As ִ��״̬, a.����, Sum(a.���ʽ��) As ���ʽ��, Max(���մ���id) As ���մ���id,
                Max(������Ŀ��) As ������Ŀ��, Max(���ձ���) As ���ձ���, Max(��������) As ��������, Sum(a.ͳ����) As ͳ����, Max(�Ƿ��ϴ�) As �Ƿ��ϴ�, �Ƿ���,
                a.�Һ�id, a.��ҳid, a.���˲���id
               From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2list(v_Nos))) B
               Where a.No = b.No And a.��¼���� In (1, 11)
               Group By a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����, a.��ʶ��, a.���ʽ, a.���˿���id,
                        a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.��ҩ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, a.������,
                        a.��������id, a.������, a.����ʱ��, a.ִ�в���id, a.ִ����, a.����, �Ƿ���, a.�Һ�id, a.��ҳid, a.���˲���id
               Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0) Loop
  
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��,
       ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, �Һ�id, ��ҳid, ���˲���id)
    Values
      (���˷��ü�¼_Id.Nextval, 1, c_����.No, 2, c_����.���, c_����.��������, c_����.�۸񸸺�, c_����.����id, c_����.ҽ�����, c_����.�����־, c_����.����,
       c_����.�Ա�, c_����.����, c_����.��ʶ��, c_����.���ʽ, c_����.���˿���id, c_����.�ѱ�, c_����.�շ����, c_����.�շ�ϸĿid, c_����.���㵥λ, 1, c_����.��ҩ����,
       -1 * c_����.����, c_����.�Ӱ��־, c_����.���ӱ�־, c_����.Ӥ����, c_����.������Ŀid, c_����.�վݷ�Ŀ, c_����.��׼����, -1 * c_����.Ӧ�ս��, -1 * c_����.ʵ�ս��,
       c_����.������, c_����.��������id, c_����.������, c_����.����ʱ��, �˷�ʱ��_In, c_����.ִ�в���id, c_����.ִ����, c_����.ִ��״̬, Null, c_����.����, ����Ա���_In,
       ����Ա����_In, ���ó���id_In, -1 * c_����.���ʽ��, c_����.���մ���id, c_����.������Ŀ��, c_����.���ձ���, c_����.��������, -1 * c_����.ͳ����, c_����.�Ƿ��ϴ�, '',
       c_����.�Ƿ���, n_��id, 0, c_����.�Һ�id, c_����.��ҳid, c_����.���˲���id);
  End Loop;
  Zl_�����˷ѽ���_Modify(1, n_����id, ���ó���id_In, Null);

  --3.���ϲ�������¼��ͬʱ�ѽ�����Ʊ�ݻ��պ�ҽ��ԭ���ˣ�
  Zl_���ò����¼_Delete(No_In, �������id_In, Null, �������_In, ���ó���id_In, ����Ա���_In, ����Ա����_In, �˷�ʱ��_In);
  Update ���ò����¼ Set ����״̬ = 0 Where ������� = �������_In;
  --����Ϊҽ���ӿ��ѵ��óɹ�
  Update ����Ԥ����¼
  Set У�Ա�־ = 2
  Where ��¼���� = 6 And ����id = �������id_In And ���㷽ʽ In (Select ���� From ���㷽ʽ Where ���� In (3, 4));

  --4.�������ݴ���
  Select -1 * Nvl(Sum(a.��Ԥ��), 0)
  Into n_δ�˽��
  From ����Ԥ����¼ A
  Where a.������� = �������_In And a.���㷽ʽ Is Null;
  If Nvl(n_����, 0) = 0 Then
    n_���� := Round(n_δ�˽��, n_Dec) - n_δ�˽��;
  End If;
  n_δ�˽�� := n_δ�˽�� - n_����;

  For r_Ԥ�� In (Select Case
                        When Mod(a.��¼����, 10) = 1 Then
                         1
                        When Nvl(a.�����id, 0) <> 0 Then
                         2
                        Else
                         0
                      End As ����, a.����id, Nvl(a.��Ԥ��, 0) As ��Ԥ��, a.No, a.����id, a.���㷽ʽ, a.�����id, a.����, a.������ˮ��, a.����˵��,
                      a.�������
               From ����Ԥ����¼ A, ���㷽ʽ B
               Where a.���㷽ʽ = b.���� And a.��¼״̬ In (1, 3) And b.���� Not In (3, 4, 9) And
                     a.����id In (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ���ӱ�־ = 0 And NO = No_In)) Loop
  
    --���ǵ��ֽ��㷽ʽ
    If r_Ԥ��.���� = 1 Then
      --Ԥ����
      Zl_���ò������_����˷�(�������id_In, Null, Null, Null, Null, Null, n_����, 0, 0, -1 * n_δ�˽��);
      Exit;
    Elsif r_Ԥ��.���� = 2 Then
      --һ��ͨ
      Select Nvl(Sum(���), 0) Into n_���˽�� From �����˿���Ϣ Where ��¼id = r_Ԥ��.����id;
      If r_Ԥ��.��Ԥ�� - n_���˽�� > 0 Then
        If r_Ԥ��.��Ԥ�� - n_���˽�� > n_δ�˽�� Then
          n_��Ԥ�� := n_δ�˽��;
        Else
          n_��Ԥ�� := r_Ԥ��.��Ԥ�� - n_���˽��;
        End If;
      
        v_���㷽ʽ := r_Ԥ��.���㷽ʽ || '|' || -1 * n_��Ԥ�� || '| | ';
        Zl_���ò������_����˷�(�������id_In, v_���㷽ʽ, r_Ԥ��.�����id, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��, n_����, 0, 1);
        Zl_�����˿���Ϣ_Insert(�������_In, r_Ԥ��.����id, n_��Ԥ��, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��);
      
        --תΪסԺԤ����
        ����Ԥ����¼_Insert(r_Ԥ��.����id, n_��Ԥ��, r_Ԥ��.���㷽ʽ, �˷�ʱ��_In, r_Ԥ��.�������, r_Ԥ��.�����id, r_Ԥ��.����, r_Ԥ��.������ˮ��, r_Ԥ��.����˵��);
      
        n_δ�˽�� := n_δ�˽�� - n_��Ԥ��;
        n_����   := 0;
      End If;
      If n_δ�˽�� = 0 Then
        Exit;
      End If;
    Else
      --������ҽ�����㷽ʽ
      --���㷽ʽ|������|�������|����ժҪ
      v_���㷽ʽ := r_Ԥ��.���㷽ʽ || '|' || n_δ�˽�� || '| | ';
      Zl_���ò������_����˷�(�������id_In, v_���㷽ʽ, Null, Null, Null, Null, n_����, 0);
    
      --תΪסԺԤ����
      ����Ԥ����¼_Insert(r_Ԥ��.����id, n_δ�˽��, r_Ԥ��.���㷽ʽ, �˷�ʱ��_In, r_Ԥ��.�������);
      Exit;
    End If;
  End Loop;

  --5.ת����ɴ���   
  Delete From ����Ԥ����¼ Where ����id = �������id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Raise Err_Item;
  End If;
  Delete From ����Ԥ����¼ Where ����id = ���ó���id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '������δ�ɿ�����ݣ�������ɽ��㣡';
    Raise Err_Item;
  End If;
  Update ����Ԥ����¼ Set У�Ա�־ = 0, �Ự�� = Null Where ������� = �������_In;

  --��Ա�ɿ�����Ҫ��ҽ����
  For c_Ԥ�� In (Select a.���㷽ʽ, a.����Ա����, Nvl(Sum(a.��Ԥ��), 0) As ��Ԥ��
               From ����Ԥ����¼ A, ���㷽ʽ B
               Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And a.������� = �������_In
               Group By a.���㷽ʽ, a.����Ա����) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + c_Ԥ��.��Ԥ��
    Where �տ�Ա = c_Ԥ��.����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (c_Ԥ��.����Ա����, c_Ԥ��.���㷽ʽ, 1, c_Ԥ��.��Ԥ��);
      n_����ֵ := c_Ԥ��.��Ԥ��;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = c_Ԥ��.����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ And Nvl(���, 0) = 0;
    End If;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_������ת��;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_����תסԺ_�շ�ת��
(
  No_In         סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  �����˷�_In   Number := 0,
  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����id_In     ����Ԥ����¼.����id%Type := Null,
  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
  ����_In     ����Ԥ����¼.��Ԥ��%Type := Null
) As
  --�����˷�_In:0-����תסԺ��������;1-�����˷�ģʽ
  -- �����˷�_InΪ1ʱ:��Ժ����id_In����ҳID_IN���Բ�����
  n_Count      Number(5);
  n_ԭ����id   סԺ���ü�¼.����id%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  n_ʵ�ʳ���   ����Ԥ����¼.��Ԥ��%Type;
  n_��id       ����ɿ����.Id%Type;
  n_����id     ������Ϣ.����id%Type;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_Ԥ�����   ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid     Ʊ��ʹ����ϸ.��ӡid%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  v_������     ������ü�¼.������%Type;
  n_����id     ������ü�¼.����id%Type;
  v_����     ���㷽ʽ.����%Type;
  n_����ֵ     �������.�������%Type;
  v_���㷽ʽ   ���㷽ʽ.����%Type;
  v_Nos        Varchar2(3000);
  v_����ids    Varchar2(3000);
  v_ԭ����ids  Varchar2(3000);
  n_Tempid     ����Ԥ����¼.Id%Type;
  n_ҽ��       Number;
  n_����       Number;
  n_����       Number;
  n_�����˷�   Number;
  n_�˷�����   Number;
  n_����״̬   ������ü�¼.����״̬%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --����
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'û�з��������㷽ʽ�������Ƿ���ȷ���ã�';
      Raise Err_Item;
  End;

  If ԭ����id_In Is Null Then
  
    Select Count(NO), Sum(ʵ�ս��)
    Into n_Count, n_ʵ�ս��
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = 1;
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '����' || No_In || '�����շѵ��ݻ��򲢷�ԭ�����˲����˸õ���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
    --1.1���Ϸ��ü�¼
    If ����id_In Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
  
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
       ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��,
       ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id, ����״̬, ��ҳid, ���˲���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
             �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -1 * Ӧ�ս��, -1 * ʵ�ս��, ��������id,
             ������, ִ�в���id, ������, ִ����, -1, ִ��ʱ��, ����Ա���_In, ����Ա����_In, ����ʱ��, �˷�ʱ��_In, n_����id, -1 * ���ʽ��, ������Ŀ��, ���մ���id, ͳ����,
             ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id, 0, ��ҳid, ���˲���id
      From ������ü�¼
      Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 1;
  
    --Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
  
    --1.2����Ԥ����¼
    --���ϳ�Ԥ������
    For r_����id In (Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select ����id
                                               From ����Ԥ����¼
                                               Where ������� In (Select b.�������
                                                              From ������ü�¼ A, ����Ԥ����¼ B
                                                              Where a.No = No_In And b.������� < 0 And Mod(a.��¼����, 10) = 1 And
                                                                    a.��¼״̬ <> 0 And a.����id = b.����id))) And
                         Mod(��¼����, 10) = 1 And ��¼״̬ <> 0
                   Union
                   Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select a.����id
                                               From ������ü�¼ A, ����Ԥ����¼ B
                                               Where a.No = No_In And b.������� > 0 And Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And
                                                     a.����id = b.����id)) And Mod(��¼����, 10) = 1 And ��¼״̬ <> 0) Loop
      v_ԭ����ids := v_ԭ����ids || ',' || r_����id.����id;
    End Loop;
    v_ԭ����ids := Substr(v_ԭ����ids, 2);
  
    Begin
      Select 1
      Into n_ҽ��
      From ���ս����¼
      Where ��¼id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And Rownum < 2;
    Exception
      When Others Then
        n_ҽ�� := 0;
    End;
  
    If n_ҽ�� = 1 Then
      Begin
        Select 1
        Into n_����
        From ҽ��������ϸ
        Where NO = No_In And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '��ǰ����' || No_In || '������ҽ��������ϸ,�޷���������תסԺ!';
          Raise Err_Item;
      End;
    End If;
  
    --ҽ���˿�
    For r_ҽ�� In (Select ����id, NO, ���㷽ʽ, ���, ��ע
                 From ҽ��������ϸ
                 Where NO = No_In And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids)))) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) - r_ҽ��.���
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_ҽ��.���㷽ʽ, 1, -1 * r_ҽ��.���);
        n_����ֵ := r_ҽ��.���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ And Nvl(���, 0) = 0;
      End If;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + (-1 * r_ҽ��.���)
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
           �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_ҽ��.���, r_ҽ��.���㷽ʽ, Null, �˷�ʱ��_In,
           Null, Null, Null, ����Ա���_In, ����Ա����_In, r_ҽ��.��ע, n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id,
           0, 3);
      End If;
    
      Update ����Ԥ����¼
      Set ��¼״̬ = 3
      Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
            ���㷽ʽ = r_ҽ��.���㷽ʽ;
    
      Update ҽ��������ϸ
      Set ��� = ��� + (-1 * r_ҽ��.���)
      Where NO = No_In And ����id = n_����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���)
        Values
          (n_����id, No_In, r_ҽ��.���㷽ʽ, -1 * r_ҽ��.���);
      End If;
      n_ʵ�ս�� := n_ʵ�ս�� - r_ҽ��.���;
    End Loop;
  
    Begin
      Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
    Exception
      When Others Then
        Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
    End;
  
    If n_ʵ�ս�� <> 0 Then
      For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���,
                              ����, ������ˮ��, ����˵��, ������λ
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids)))
                       Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �����id, ���㿨���, ����,
                                ������ˮ��, ����˵��, ������λ) Loop
        If n_ʵ�ս�� <> 0 Then
          If r_Prepay.��Ԥ�� >= n_ʵ�ս�� Then
            Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, �ɿ���id)
              Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                     r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                     ����Ա���_In, -1 * n_ʵ�ս��, n_����id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                     r_Prepay.����˵��, r_Prepay.������λ, 1, -1 * n_����id, n_��id
              From Dual;
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(n_ʵ�ս��, 0)
            Where ����id = n_����id And ���� = 1 And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_ʵ�ս��, 1);
              n_����ֵ := n_ʵ�ս��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
            n_ʵ�ս�� := 0;
          Else
            Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, �ɿ���id)
              Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                     r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                     ����Ա���_In, -1 * r_Prepay.��Ԥ��, n_����id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                     r_Prepay.����˵��, r_Prepay.������λ, 1, -1 * n_����id, n_��id
              From Dual;
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Prepay.��Ԥ��, 0)
            Where ����id = n_����id And ���� = 1 And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, r_Prepay.��Ԥ��, 1);
              n_����ֵ := r_Prepay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
            n_ʵ�ս�� := n_ʵ�ս�� - r_Prepay.��Ԥ��;
          End If;
        End If;
      End Loop;
    End If;
    --2.Ʊ���ջ�
    --������ǰû�д�ӡ,���ջ�
    Select Nvl(Max(ID), 0)
    Into n_��ӡid
    From (Select b.Id
           From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
           Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = No_In
           Order By a.ʹ��ʱ�� Desc)
    Where Rownum < 2;
    If n_��ӡid > 0 Then
      --���ŵ���ѭ������ʱֻ���ջ�һ��
      Select Count(��ӡid) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
      If n_Count = 0 Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, �˷�ʱ��_In, ����Ա����_In, Ʊ�ݽ��
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
      End If;
    End If;
  
    --3.�ɿ����ݴ���(
    --   �����������:
    --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
    --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
    If Nvl(�����˷�_In, 0) = 1 Then
      For c_Ԥ�� In (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, Min(a.������ˮ��) As ������ˮ��,
                          Min(a.����˵��) As ����˵��, Min(a.������λ) As ������λ, b.����
                   From ����Ԥ����¼ A, ���㷽ʽ B
                   Where a.��¼���� = 3 And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                         a.���㷽ʽ = b.���� And b.���� In (1, 2, 7, 8) And a.���㷽ʽ Is Not Null
                   Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����
                   Having Sum(a.��Ԥ��) <> 0
                   Order By a.�����id, ���� Desc) Loop
        If n_ʵ�ս�� <> 0 Then
          Begin
            Select �Ƿ����� Into n_���� From ҽ�ƿ���� Where ID = c_Ԥ��.�����id;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If (c_Ԥ��.���� = 7 Or (c_Ԥ��.���� = 8 And c_Ԥ��.�����id Is Not Null)) And n_���� = 0 Then
            If c_Ԥ��.��Ԥ�� > n_ʵ�ս�� Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * n_ʵ�ս�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = c_Ԥ��.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := 0;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = c_Ԥ��.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := n_ʵ�ս�� - c_Ԥ��.��Ԥ��;
            End If;
          Else
            n_ʵ�ʳ��� := 0;
            If c_Ԥ��.���� In (3, 4) Or (c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null) Then
              v_���㷽ʽ := c_Ԥ��.���㷽ʽ;
            Else
              If ���㷽ʽ_In Is Null Then
                Begin
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
                Exception
                  When Others Then
                    Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
                End;
              Else
                v_���㷽ʽ := ���㷽ʽ_In;
              End If;
            End If;
          
            If c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null Then
              If n_ʵ�ս�� >= c_Ԥ��.��Ԥ�� Then
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, c_Ԥ��.��Ԥ��, c_Ԥ��.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null,
                     �˷�ʱ��_In, Null, Null, Null, ����Ա���_In, ����Ա����_In,
                     '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id, Null, Null, Null, Null, Null, Null,
                     n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := c_Ԥ��.��Ԥ��;
              Else
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_ʵ�ս��, c_Ԥ��.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                     Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                     Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := n_ʵ�ս��;
              End If;
            Else
              If c_Ԥ��.��Ԥ�� > n_ʵ�ս�� Then
                n_ʵ�ʳ��� := n_ʵ�ս��;
              Else
                n_ʵ�ʳ��� := c_Ԥ��.��Ԥ��;
              End If;
            End If;
          
            If c_Ԥ��.���㿨��� Is Null Then
              Update ��Ա�ɿ����
              Set ��� = Nvl(���, 0) - n_ʵ�ʳ���
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
              Returning ��� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ��Ա�ɿ����
                  (�տ�Ա, ���㷽ʽ, ����, ���)
                Values
                  (����Ա����_In, v_���㷽ʽ, 1, -1 * n_ʵ�ʳ���);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From ��Ա�ɿ����
                Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
              End If;
            
              --��ԭԤ����¼
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ʳ���)
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, c_Ԥ��.������λ, n_����id,
                   -1 * n_����id, 0, 3);
              End If;
            End If;
            Update ����Ԥ����¼
            Set ��¼״̬ = 3
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                  ���㷽ʽ = c_Ԥ��.���㷽ʽ;
            n_ʵ�ս�� := n_ʵ�ս�� - n_ʵ�ʳ���;
          End If;
        End If;
      End Loop;
    
      --���·�����˼�¼
      Update ������˼�¼
      Set ��¼״̬ = 2
      Where ����id In (Select ID From ������ü�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3)) And ���� = 1;
      --���������¼
      Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 1;
      For r_Clinic In (Select ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
                              ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                              Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id, ������, Max(���ʵ�id) As ���ʵ�id, ����ʱ��,
                              ʵ��Ʊ��
                       From ������ü�¼
                       Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (2, 3) And Nvl(���ӱ�־, 0) Not In (8, 9)
                       Group By ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���,
                                ��������, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ����ʱ��, ʵ��Ʊ��
                       Having Sum(����) <> 0) Loop
        Insert Into ������ü�¼
          (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
           ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �ɿ���id, ����id, ���ʽ��, ����״̬)
        Values
          (���˷��ü�¼_Id.Nextval, 1, No_In, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1, r_Clinic.����id,
           '', r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid,
           r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����,
           -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����,
           -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��,
           �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', n_��id, n_����id,
           -1 * r_Clinic.ʵ�ս��, 0);
      End Loop;
    Else
      --4.�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
      For r_Pay In (Select Min(a.Id) As Ԥ��id, a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��,
                           a.����˵��, a.������λ, b.����
                    From ����Ԥ����¼ A, ���㷽ʽ B
                    Where a.��¼���� = 3 And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                          a.���㷽ʽ = b.���� And (b.���� In (1, 2, 7, 8)) And a.���㷽ʽ Is Not Null
                    Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����, a.������ˮ��, a.����˵��, a.������λ

                    
                    Having Sum(a.��Ԥ��) <> 0
                    Order By a.�����id, ���� Desc) Loop
        --4.1����Ԥ����� (�����ڲ����˷ѵ����)
        --���е���,����������Ԥ�����
        --��Ϊ�տ�������ɿ�,������Ա�ɿ�����ޱ仯
        If n_ʵ�ս�� <> 0 Then
          If r_Pay.���� = 7 Or (r_Pay.���� = 8 And r_Pay.�����id Is Not Null) Then
            If r_Pay.��Ԥ�� > n_ʵ�ս�� Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * n_ʵ�ս�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = r_Pay.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := 0;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|',
                   n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = r_Pay.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := n_ʵ�ս�� - r_Pay.��Ԥ��;
            End If;
          Else
            n_ʵ�ʳ��� := 0;
            If r_Pay.���� In (3, 4) Or (r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null) Then
              v_���㷽ʽ := r_Pay.���㷽ʽ;
            Else
              If ���㷽ʽ_In Is Null Then
                Begin
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
                Exception
                  When Others Then
                    Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
                End;
              Else
                v_���㷽ʽ := ���㷽ʽ_In;
              End If;
            End If;
          
            If r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null Then
              If n_ʵ�ս�� >= r_Pay.��Ԥ�� Then
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, r_Pay.��Ԥ��, r_Pay.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null,
                     �˷�ʱ��_In, Null, Null, Null, ����Ա���_In, ����Ա����_In,
                     '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id, Null, Null, Null, Null, Null,
                     Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := r_Pay.��Ԥ��;
              Else
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_ʵ�ս��, r_Pay.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                     Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || r_Pay.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                     Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := n_ʵ�ս��;
              End If;
            Else
              If r_Pay.��Ԥ�� > n_ʵ�ս�� Then
                n_ʵ�ʳ��� := n_ʵ�ս��;
              Else
                n_ʵ�ʳ��� := r_Pay.��Ԥ��;
              End If;
            End If;
          
            If r_Pay.���� Not In (3, 4, 7, 8) Then
              Update ����Ԥ����¼
              Set ��� = ��� + n_ʵ�ʳ���
              Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                v_Ԥ��no := Nextno(11);
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, Ԥ�����)
                Values
                  (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, r_Pay.Ԥ�����);
              End If;
            
              --�������
              Update �������
              Set Ԥ����� = Nvl(Ԥ�����, 0) + n_ʵ�ʳ���
              Where ���� = 1 And ����id = n_����id And ���� = 2
              Returning Ԥ����� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, n_ʵ�ʳ���, 0);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From �������
                Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
              End If;
            End If;
            --4.2�ɿ����ݴ���
            --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
            --�����˷��������ԭԤ����¼
            If r_Pay.���� In (3, 4) Then
              Update ��Ա�ɿ����
              Set ��� = Nvl(���, 0) - n_ʵ�ʳ���
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ
              Returning ��� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ��Ա�ɿ����
                  (�տ�Ա, ���㷽ʽ, ����, ���)
                Values
                  (����Ա����_In, r_Pay.���㷽ʽ, 1, -1 * n_ʵ�ʳ���);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From ��Ա�ɿ����
                Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(���, 0) = 0;
              End If;
            End If;
          
            If r_Pay.���� <> 8 Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ʳ���)
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��,
                   r_Pay.����˵��, r_Pay.������λ, n_����id, -1 * n_����id, 0, 3);
              End If;
            End If;
          
            Update ����Ԥ����¼
            Set ��¼״̬ = 3
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                  ���㷽ʽ = r_Pay.���㷽ʽ;
            n_ʵ�ս�� := n_ʵ�ս�� - n_ʵ�ʳ���;
          
          End If;
        End If;
      End Loop;
    End If;
  
    If ����_In Is Not Null Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, ����_In, v_����, Null, �˷�ʱ��_In, Null, Null,
         Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3);
    End If;
    Delete From ����Ԥ����¼
    Where ����id = n_����id And ��¼���� = 3 And ��¼״̬ = 2 And ��Ԥ�� = 0 And ���㷽ʽ Is Not Null;
    Delete From ����Ԥ����¼ Where ����id = n_ԭ����id And ժҪ = 'Ԥ����ʱ��¼' And ��¼���� = 3;
    Update ������ü�¼ Set ����״̬ = Nvl(n_����״̬, 0) Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
  Else
    --ҽ��������ת��
    For r_Nos In (Select Distinct a.No
                  From ������ü�¼ A
                  Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And a.����id = ԭ����id_In) Loop
      v_Nos := v_Nos || ',' || r_Nos.No;
    End Loop;
    v_Nos := Substr(v_Nos, 2);
  
    For r_����ids In (Select Distinct a.����id
                    From ������ü�¼ A
                    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                          a.��¼״̬ <> 0) Loop
      v_����ids := v_����ids || ',' || r_����ids.����id;
    End Loop;
    v_����ids := Substr(v_����ids, 2);
    Select Count(a.No), Sum(a.ʵ�ս��)
    Into n_Count, n_ʵ�ս��
    From ������ü�¼ A
    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1;
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '���ν��㲻���շѻ��򲢷�ԭ�����˲����˸ý���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where ����id = ԭ����id_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
    Begin
      Select 1
      Into n_�����˷�
      From ������ü�¼ A
      Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 2 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            Rownum < 2;
    Exception
      When Others Then
        n_�����˷� := 0;
    End;
  
    Begin
      Select 0
      Into n_�����˷�
      From ������ü�¼ A
      Where ��¼���� = 11 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select Count(Avg(1))
      Into n_�˷�����
      From ����Ԥ����¼ A
      Where a.��¼���� = 3 And a.��¼״̬ <> 0 And ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))
      Group By a.���㷽ʽ;
    Exception
      When Others Then
        n_�˷����� := 0;
    End;
    --1.1���Ϸ��ü�¼
    If ����id_In Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
       ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��,
       ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id, ����״̬)
      Select ���˷��ü�¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����,
             a.��ʶ��, a.���ʽ, a.�ѱ�, a.���˿���id, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.����, a.��ҩ����, -1 * a.����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid,
             a.�վݷ�Ŀ, a.���ʷ���, a.��׼����, -1 * a.Ӧ�ս��, -1 * a.ʵ�ս��, a.��������id, a.������, a.ִ�в���id, a.������, a.ִ����, -1, a.ִ��ʱ��,
             ����Ա���_In, ����Ա����_In, a.����ʱ��, �˷�ʱ��_In, n_����id, -1 * a.���ʽ��, a.������Ŀ��, a.���մ���id, a.ͳ����, a.ժҪ,
             Decode(Nvl(a.���ӱ�־, 0), 9, 1, 0), a.���ձ���, a.��������, n_��id, 0
      From ������ü�¼ A
      Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 1;
  
    --����ҽ��
    For r_ҽ�� In (Select ����id, NO, ���㷽ʽ, ���, ��ע
                 From ҽ��������ϸ
                 Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And
                       ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
      Update ҽ��������ϸ
      Set ��� = ��� + (-1 * r_ҽ��.���)
      Where NO = r_ҽ��.No And ����id = r_ҽ��.����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���)
        Values
          (r_ҽ��.����id, r_ҽ��.No, r_ҽ��.���㷽ʽ, -1 * r_ҽ��.���);
      End If;
    End Loop;
  
    --Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
    --1.2����Ԥ����¼
    --���ϳ�Ԥ������
    If n_�����˷� = 0 And Nvl(�����˷�_In, 0) = 0 Then
      For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, -1 * Sum(��Ԥ��) As ��Ԥ��,
                              �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                             Nvl(��Ԥ��, 0) <> 0
                       Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���,
                                ����, ������ˮ��, ����˵��, ������λ, ��������) Loop
        Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, Ԥ�����, ��������)
          Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                 r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                 ����Ա���_In, r_Prepay.��Ԥ��, n_����id, n_��id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                 r_Prepay.����˵��, r_Prepay.������λ, -1 * n_����id, 1, r_Prepay.��������
          From Dual;
      End Loop;
    
      For v_Ԥ�� In (Select ����id, Nvl(Ԥ�����, 2) As Ԥ�����, Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                   From ����Ԥ����¼ A
                   Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                         a.����id <> n_����id
                   Group By ����id, Nvl(Ԥ�����, 2)
                   Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
      
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
        Where ����id = v_Ԥ��.����id And ���� = Nvl(v_Ԥ��.Ԥ�����, 2) And ���� = 1
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, Ԥ�����, ����)
          Values
            (v_Ԥ��.����id, Nvl(v_Ԥ��.Ԥ�����, 2), v_Ԥ��.Ԥ�����, 1);
          n_����ֵ := v_Ԥ��.Ԥ�����;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From �������
          Where ����id = v_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      End Loop;
    Else
      If n_�˷����� = 0 And Nvl(�����˷�_In, 0) = 0 Then
        --ֻʹ����Ԥ����ԭ���˻�Ԥ��
        For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, Max(���㷽ʽ) As ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��,
                                -1 * Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                         From ����Ԥ����¼ A
                         Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                               Nvl(��Ԥ��, 0) <> 0
                         Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���, ����,
                                  ������ˮ��, ����˵��, ������λ, ��������) Loop
          Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
             ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, Ԥ�����, ��������)
            Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                   r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                   ����Ա���_In, r_Prepay.��Ԥ��, n_����id, n_��id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                   r_Prepay.����˵��, r_Prepay.������λ, -1 * n_����id, 1, r_Prepay.��������
            From Dual;
          Select -1 * ��Ԥ�� Into n_Ԥ����� From ����Ԥ����¼ Where ID = n_Tempid;
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(n_Ԥ�����, 0)
          Where ����id = r_Prepay.����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_Ԥ�����, 1);
            n_����ֵ := n_Ԥ�����;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Prepay.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
          End If;
        End Loop;
      Else
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
        Exception
          When Others Then
            Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
        End;
        Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
          Select n_Tempid, Max(NO), Max(ʵ��Ʊ��), 3, 3, ����id, ��ҳid, ����id, Null, v_���㷽ʽ, Max(�������), 'Ԥ����ʱ��¼', Null, Null,
                 Null, Max(�տ�ʱ��), ����Ա����_In, ����Ա���_In, Sum(��Ԥ��), n_ԭ����id, Null, Null, Null, Null, Null, Null,
                 -1 * n_ԭ����id, 3
          From ����Ԥ����¼ A
          Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                Nvl(��Ԥ��, 0) <> 0
          Group By n_Tempid, 3, 3, ����id, ��ҳid, ����id, Null, v_���㷽ʽ, 'Ԥ����ʱ��¼', ����Ա����_In, ����Ա���_In, n_ԭ����id;
      End If;
    End If;
  
    --��������ɷѼ�ҽ������
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˷�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             0, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * n_����id, ��������
      From ����Ԥ����¼ A, ���㷽ʽ B
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            a.���㷽ʽ = b.���� And b.���� Not In (7, 8);
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������, У�Ա�־)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˷�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             0, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * n_����id, ��������, 1
      From ����Ԥ����¼ A, ���㷽ʽ B
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            a.���㷽ʽ = b.���� And b.���� = 7;
    If Sql%RowCount <> 0 Then
      n_����״̬ := 1;
    End If;
  
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_����ids)));
  
    --2.Ʊ���ջ�
    --������ǰû�д�ӡ,���ջ�
    For r_Nos In (Select Distinct a.No
                  From ������ü�¼ A
                  Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And
                        a.����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = r_Nos.No
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
      If n_��ӡid > 0 Then
        --���ŵ���ѭ������ʱֻ���ջ�һ��
        Select Count(��ӡid) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, �˷�ʱ��_In, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        End If;
      End If;
    End Loop;
  
    --3.�ɿ����ݴ���(
    --   �����������:
    --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
    --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
    If Nvl(�����˷�_In, 0) = 1 Then
      For c_Ԥ�� In (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, Min(a.������ˮ��) As ������ˮ��,
                          Min(a.����˵��) As ����˵��, Min(a.������λ) As ������λ, b.����
                   From ����Ԥ����¼ A, ���㷽ʽ B
                   Where a.��¼���� = 3 And a.��¼״̬ In (2, 3) And
                         a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And a.���㷽ʽ = b.���� And
                         b.���� In (1, 2, 3, 4, 7, 8) And a.���㷽ʽ Is Not Null
                   Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����
                   Having Sum(a.��Ԥ��) <> 0) Loop
        Begin
          Select �Ƿ����� Into n_���� From ҽ�ƿ���� Where ID = c_Ԥ��.�����id;
        Exception
          When Others Then
            n_���� := 0;
        End;
        If (c_Ԥ��.���� = 7 Or (c_Ԥ��.���� = 8 And c_Ԥ��.�����id Is Not Null)) And n_���� = 0 Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
               Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
          End If;
          n_����״̬ := 1;
        Else
          If c_Ԥ��.���� In (3, 4) Or (c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null) Then
            v_���㷽ʽ := c_Ԥ��.���㷽ʽ;
          Else
            If ���㷽ʽ_In Is Null Then
              Begin
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
              Exception
                When Others Then
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
              End;
            Else
              v_���㷽ʽ := ���㷽ʽ_In;
            End If;
          End If;
        
          If c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null Then
            --Zl_Square_Update(v_����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, c_Ԥ��.��Ԥ��, c_Ԥ��.���㿨���);
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
                 Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
            End If;
            n_����״̬ := 1;
          End If;
          If c_Ԥ��.���㿨��� Is Null Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - c_Ԥ��.��Ԥ��
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, v_���㷽ʽ, 1, -1 * c_Ԥ��.��Ԥ��);
              n_����ֵ := c_Ԥ��.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
            End If;
            --�����˷��������ԭԤ����¼
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, c_Ԥ��.������λ, n_����id,
                 -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    
      --���·�����˼�¼
      Update ������˼�¼
      Set ��¼״̬ = 2
      Where ����id In (Select a.Id
                     From ������ü�¼ A
                     Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                           a.��¼״̬ In (1, 3)) And ���� = 1;
      --���������¼
      For r_Nos In (Select Distinct NO
                    From ������ü�¼
                    Where Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And
                          ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
        Update ������ü�¼ Set ��¼״̬ = 3 Where NO = r_Nos.No And Mod(��¼����, 10) = 1 And ��¼״̬ = 1;
      End Loop;
      For r_Clinic In (Select Min(a.��¼����) As ��¼����, a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.����, a.�Ա�, a.����, a.���˿���id, a.�ѱ�,
                              a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����, Sum(a.����) As ����,
                              a.�Ӱ��־, a.���ӱ�־, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��,
                              Sum(a.ͳ����) As ͳ����, a.��������id, a.������, a.ִ�в���id, a.������, Max(a.���ʵ�id) As ���ʵ�id,
                              Max(a.�Ƿ���) As �Ƿ���, a.����ʱ��, Min(a.ʵ��Ʊ��) As ʵ��Ʊ��
                       From ������ü�¼ A
                       Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                             a.��¼״̬ In (2, 3) And Nvl(a.���ӱ�־, 0) Not In (8, 9)
                       Group By a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.����, a.�Ա�, a.����, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid,
                                a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid, a.�վݷ�Ŀ,
                                a.��׼����, a.��������id, a.������, a.ִ�в���id, a.������, a.����ʱ��
                       Having Sum(a.����) <> 0) Loop
        Insert Into ������ü�¼
          (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
           ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ����id, ���ʽ��, ִ��״̬, ����״̬)
        Values
          (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, r_Clinic.No, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�,
           1, r_Clinic.����id, '', r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����,
           r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����,
           r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����,
           -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��,
           �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', r_Clinic.�Ƿ���, n_��id, n_����id,
           -1 * r_Clinic.ʵ�ս��, -1, 0);
      End Loop;
    Else
      --4.�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
    
      For r_Pay In (Select Min(a.Id) As Ԥ��id, a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��,
                           a.����˵��, a.������λ, b.����
                    From ����Ԥ����¼ A, ���㷽ʽ B
                    Where a.��¼���� = 3 And a.��¼״̬ In (2, 3) And
                          a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And a.���㷽ʽ = b.���� And
                          b.���� In (1, 2, 3, 4, 7, 8) And a.���㷽ʽ Is Not Null
                    Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����, a.������ˮ��, a.����˵��, a.������λ

                    
                    Having Sum(a.��Ԥ��) <> 0) Loop
        --4.1����Ԥ����� (�����ڲ����˷ѵ����)
        --���е���,����������Ԥ�����
        --��Ϊ�տ�������ɿ�,������Ա�ɿ�����ޱ仯
        If r_Pay.���� = 7 Or (r_Pay.���� = 8 And r_Pay.�����id Is Not Null) Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|'
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id,
               Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
          End If;
          n_����״̬ := 1;
        Else
          If r_Pay.���� In (3, 4) Or (r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null) Then
            v_���㷽ʽ := r_Pay.���㷽ʽ;
          Else
            Begin
              Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
            Exception
              When Others Then
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
            End;
          End If;
        
          If r_Pay.���� = 8 Then
            --Zl_Square_Update(v_����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, r_Pay.��Ԥ��, r_Pay.���㿨���);
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|'
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id,
                 Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
            End If;
            n_����״̬ := 1;
          End If;
          If r_Pay.���� Not In (3, 4, 7, 8) Then
            Update ����Ԥ����¼
            Set ��� = ��� + r_Pay.��Ԥ��
            Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              v_Ԥ��no := Nextno(11);
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, Ԥ�����)
              Values
                (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, r_Pay.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, r_Pay.Ԥ�����);
            End If;
          
            --�������
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + r_Pay.��Ԥ��
            Where ���� = 1 And ����id = n_����id And ���� = 2
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, r_Pay.��Ԥ��, 0);
              n_����ֵ := r_Pay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End If;
          --4.2�ɿ����ݴ���
          --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
          --�����˷��������ԭԤ����¼
          If r_Pay.���� In (3, 4) Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - r_Pay.��Ԥ��
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, r_Pay.���㷽ʽ, 1, -1 * r_Pay.��Ԥ��);
              n_����ֵ := r_Pay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(���, 0) = 0;
            End If;
          End If;
        
          If r_Pay.���㿨��� Is Null Then
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��,
                 r_Pay.����˵��, r_Pay.������λ, n_����id, -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    End If;
    If ����_In Is Not Null Then
      Begin
        Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
      Exception
        When Others Then
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
      End;
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - ����_In
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + ����_In
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_����;
      If Sql%RowCount = 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
           �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, ����_In, v_����, Null, �˷�ʱ��_In, Null, Null,
           Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3);
      End If;
    End If;
    Delete From ����Ԥ����¼ Where ����id = n_ԭ����id And ժҪ = 'Ԥ����ʱ��¼' And ��¼���� = 3;
    Delete From ����Ԥ����¼
    Where ����id = n_����id And ��¼���� = 3 And ��¼״̬ = 2 And ��Ԥ�� = 0 And ���㷽ʽ Is Not Null;
    Update ������ü�¼
    Set ����״̬ = Nvl(n_����״̬, 0)
    Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_�շ�ת��;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_����תסԺ_����ת��
(
  No_In         סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  ��������_In   Number := 0
) As
  --��������_In:0-����תסԺ��������;1-��������˷�ģʽ
  n_Count      Number(5);
  n_ʵ�ս��   סԺ���ü�¼.ʵ�ս��%Type;
  n_����id     סԺ���ü�¼.����id%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  v_������     ������ü�¼.������%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);
Begin

  Select Count(NO), Sum(ʵ�ս��) Into n_Count, n_ʵ�ս�� From ������ü�¼ Where NO = No_In And ��¼���� = 2;
  If n_Count = 0 Then
    v_Err_Msg := '����' || No_In || '���Ǽ��ʵ��ݻ��򲢷�ԭ�����˲����˸õ���,����תΪסԺ����.';
    Raise Err_Item;
  End If;

  Select ����id, ��������id, ������
  Into n_����id, n_��������id, v_������
  From ������ü�¼
  Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (1, 3) And Rownum = 1;

  --���������
  Begin
    Select Nvl(Sum(ʵ�ս��), 0)
    Into n_ʵ�ս��
    From ������ü�¼
    Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (1, 2, 3) And Nvl(�����־, 0) <> 4 And ����id Is Null
    Group By NO, ��¼����;
  Exception
    When Others Then
      n_ʵ�ս�� := 0;
  End;

  Update ������� Set ������� = Nvl(�������, 0) - n_ʵ�ս�� Where ����id = n_����id And ���� = 1 And ���� = 1;
  If Sql%RowCount = 0 And n_ʵ�ս�� <> 0 Then
    Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (n_����id, 1, 1, -1 * n_ʵ�ս��, 0);
  End If;

  --����δ�����
  For v_δ�� In (Select ��������id, ����id, ���˿���id, ִ�в���id, ������Ŀid, �����־, -1 * Nvl(Sum(ʵ�ս��), 0) As ʵ�ս��, ��ҳid, ���˲���id
               From ������ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (1, 2, 3)
               Group By ��������id, ����id, ���˿���id, ִ�в���id, ������Ŀid, �����־, ��ҳid, ���˲���id) Loop
    Update ����δ�����
    Set ��� = Nvl(���, 0) + v_δ��.ʵ�ս��
    Where ����id = v_δ��.����id And Nvl(��ҳid, 0) = Nvl(v_δ��.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(v_δ��.���˲���id, 0) And
          Nvl(���˿���id, 0) = Nvl(v_δ��.���˿���id, 0) And Nvl(��������id, 0) = Nvl(v_δ��.��������id, 0) And
          Nvl(ִ�в���id, 0) = Nvl(v_δ��.ִ�в���id, 0) And ������Ŀid + 0 = v_δ��.������Ŀid And ��Դ;�� + 0 = v_δ��.�����־;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (v_δ��.����id, v_δ��.��ҳid, v_δ��.���˲���id, v_δ��.���˿���id, v_δ��.��������id, v_δ��.ִ�в���id, v_δ��.������Ŀid, v_δ��.�����־, v_δ��.ʵ�ս��);
    End If;
  End Loop;

  --���Ϸ��ü�¼
  Insert Into ������ü�¼
    (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ,
     ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��, ����Ա���,
     ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����, ��ҳid, ���˲���id)
    Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
           �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -1 * Ӧ�ս��, -1 * ʵ�ս��, ��������id,
           ������, ִ�в���id, ������, ִ����, -1, ִ��ʱ��, ����Ա���_In, ����Ա����_In, ����ʱ��, �˷�ʱ��_In, ������Ŀ��, ���մ���id, -1 * ͳ����, ���ʵ�id, ժҪ, ���ձ���,
           �Ƿ���, ����, ��ҳid, ���˲���id
    From ������ü�¼
    Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 1;

  --Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 1;

  --ҩƷ����(δ����,��Ҫ����Ϊֱ��ת������ص�ҩ������.)
  If Nvl(��������_In, 0) = 1 Then
    Update ������˼�¼
    Set ��¼״̬ = 2
    Where ����id In (Select ID From ������ü�¼ Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (1, 3)) And ���� = 1;
    --���������¼
    Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 1;
    For r_Clinic In (Select ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
                            ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                            Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id, ������, ���ʵ�id, �Ƿ���, �ɿ���id, ����ʱ��,
                            ʵ��Ʊ��, ��ҳid, ���˲���id
                     From ������ü�¼
                     Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (2, 3) And ���ӱ�־ Not In (8, 9)
                     Group By ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
                              ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ���ʵ�id, �Ƿ���, �ɿ���id,
                              ����ʱ��, ʵ��Ʊ��, ��ҳid, ���˲���id
                     Having Sum(����) <> 0) Loop
      Insert Into ������ü�¼
        (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
         ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������,
         ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ��ҳid, ���˲���id)
      Values
        (���˷��ü�¼_Id.Nextval, 2, No_In, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1, r_Clinic.����id, '',
         r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid,
         r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����,
         -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.Ӥ����, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����,
         -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 1, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��,
         �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', r_Clinic.�Ƿ���, r_Clinic.�ɿ���id,
         r_Clinic.��ҳid, r_Clinic.���˲���id);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_����ת��;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_���˽��ʼ�¼_Cancel
(
  No_In         ���˽��ʼ�¼.No%Type,
  ����id_In     ���˽��ʼ�¼.Id%Type,
  ����Ա���_In ���˽��ʼ�¼.����Ա���%Type,
  ����Ա����_In ���˽��ʼ�¼.����Ա����%Type,
  ����ʱ��_In   ���˽��ʼ�¼.�շ�ʱ��%Type := Null,
  Ʊ�ݺ�_In     ���˽��ʼ�¼.ʵ��Ʊ��%Type := Null,
  ����id_In     Ʊ�����ü�¼.Id%Type := Null,
  Ʊ��_In       Ʊ��ʹ����ϸ.Ʊ��%Type := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --���α�����Ԥ����¼�����Ϣ

  --���α����ڴ��������ػ��ܱ�
  Cursor c_Money(v_Id ����Ԥ����¼.����id%Type) Is
    Select NO, ��������id, ���˿���id, ִ�в���id, ���˲���id, ����id, ��ҳid, ������Ŀid, �����־, ���ʽ��
    From סԺ���ü�¼
    Where ����id = v_Id
    Union All
    Select NO, ��������id, ���˿���id, ִ�в���id, 0 As ���˲���id, ����id, 0 As ��ҳid, ������Ŀid, �����־, ���ʽ��
    From ������ü�¼
    Where ����id = v_Id;

  r_Moneyrow c_Money%RowType;

  --���α�������˵������Ϣ
  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����, a.�Ա�, a.����, a.סԺ��, a.�����, b.��ҳid, b.��Ժ����, b.��ǰ����id, b.��Ժ����id, Nvl(b.�ѱ�, a.�ѱ�) As �ѱ�, a.����, c.���� As ���ʽ
    From ������Ϣ A, ������ҳ B, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.����id = b.����id(+) And Nvl(a.��ҳid, 0) = b.��ҳid(+) And a.ҽ�Ƹ��ʽ = c.����(+);
  r_Pati c_Pati%RowType;

  --���̱���
  v_ʵ��Ʊ�� ����Ԥ����¼.ʵ��Ʊ��%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_����id   ������Ϣ.����id%Type;

  n_ԭid    ���˽��ʼ�¼.Id%Type;
  n_����id  ���˽��ʼ�¼.Id%Type;
  v_��ӡids Varchar2(5000);
  v_��ӡid  Ʊ�ݴ�ӡ����.Id%Type;

  n_��Դ     Number; --1-����;2-סԺ;3-�����סԺ
  n_����ֵ   �������.Ԥ�����%Type;
  n_��id     ����ɿ����.Id%Type;
  n_Ԥ����� Number;
  d_Date     Date;

Begin
  n_��id := Zl_Get��id(����Ա����_In);
  Begin
    Select ID, ����id, ʵ��Ʊ�� Into n_ԭid, n_����id, v_ʵ��Ʊ�� From ���˽��ʼ�¼ Where ��¼״̬ = 1 And NO = No_In;
    --��ӡ������
    Begin
      Select ID
      Into v_��ӡids
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 3 And b.No = No_In
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  Exception
    When Others Then
      Begin
        v_Err_Msg := 'û�з���Ҫ���ϵĽ��ʵ���,�����Ѿ����ϣ�';
        Raise Err_Item;
      End;
  End;

  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 3, No_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��_In, Ʊ�ݺ�_In, 1, 6, ����id_In, v_��ӡid, ����ʱ��_In, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;

  Open c_Pati(n_����id);
  Fetch c_Pati
    Into r_Pati; --���ϵͳ���ô˹���,�������ʱû�в�����Ϣ
  d_Date := ����ʱ��_In;
  If d_Date Is Null Then
    Select Sysdate Into d_Date From Dual;
  End If;
  n_����id := ����id_In;
  If Nvl(n_����id, 0) = 0 Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  End If;

  --���˽��ʼ�¼
  Insert Into ���˽��ʼ�¼
    (ID, NO, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, ��ʼ����, ��������, �շ�ʱ��, ��ע, ԭ��, �ɿ���id, ��������, ����״̬, ��ҳid, סԺ����, ���ʽ��)
    Select n_����id, NO, ʵ��Ʊ��, 2, ��;����, ����id, ����Ա���_In, ����Ա����_In, ��ʼ����, ��������, d_Date, ��ע, ԭ��, n_��id, ��������, 1, ��ҳid, סԺ����,
           -1 * ���ʽ��
    From ���˽��ʼ�¼
    Where ID = n_ԭid;

  Update ���˽��ʼ�¼ Set ��¼״̬ = 3 Where ID = n_ԭid;

  --�����ջ�Ʊ��(������ǰû��ʹ��Ʊ��,�޷��ջ�)
  If v_��ӡids Is Not Null Then
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
      Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
      From Ʊ��ʹ����ϸ
      Where ��ӡid In (Select Column_Value From Table(f_Str2list(v_��ӡids))) And Ʊ�� In (1, 3) And ���� = 1;
  End If;

  Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
  --������㷽ʽΪNULL�Ľ��㷽ʽ
  Insert Into ����Ԥ����¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
     �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, У�Ա�־)
    Select n_Ԥ��id, No_In, v_ʵ��Ʊ��, 12, 1, n_����id, Max(Decode(Mod(��¼����, 10), 1, Null, ��ҳid)),
           Max(Decode(Mod(��¼����, 10), 1, Null, ����id)), Null, Null, Null, Null, Null, Null, Null, d_Date, ����Ա����_In,
           ����Ա���_In, -1 * Sum(��Ԥ��), n_����id, n_��id, Null, Null, Null, Null As ����, Null As ������ˮ��, Null As ����˵��,
           Null As ������λ, 2, 1
    From ����Ԥ����¼
    Where ����id = n_ԭid;

  --ȷ�����ʵķ��ü�¼��Դ
  Begin
    Select Case
             When Nvl(Max(סԺ), 0) = 1 And Nvl(Max(����), 0) = 1 Then
              3
             When Nvl(Max(סԺ), 0) = 1 Then
              2
             Else
              1
           End
    Into n_��Դ
    From (Select 1 As סԺ, 0 As ����
           From סԺ���ü�¼
           Where ����id = n_ԭid And Rownum = 1
           Union All
           Select 0 As סԺ, 1 As ����
           From ������ü�¼
           Where ����id = n_ԭid And Rownum = 1);
  
  Exception
    When Others Then
      n_��Դ := 3;
  End;

  If n_��Դ = 2 Or n_��Դ = 3 Then
    --���Ͻ��ʶ�Ӧ�ķ��ü�¼:������ԭʼ���ʲ����������Ŀ
    Insert Into סԺ���ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id,
       ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id,
       ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ,
       �ɿ���id, ҽ��С��id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�,
             ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����,
             �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Null, Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����,
             ִ��ʱ��, ����Ա����, ����Ա���, -1 * ���ʽ��, n_����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id, ҽ��С��id
      From סԺ���ü�¼
      Where ����id = n_ԭid;
  End If;

  If n_��Դ = 1 Or n_��Դ = 3 Then
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
       �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
       ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id, ��ҳid, ���˲���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id,
             ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����,
             ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Null, Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���,
             -1 * ���ʽ��, n_����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id, ��ҳid, ���˲���id
      From ������ü�¼
      Where ����id = n_ԭid;
  End If;

  For r_Moneyrow In c_Money(n_����id) Loop
    --������� ,���Բ���Ҫ�������������ܱ�
  
    If Nvl(r_Moneyrow.�����־, 0) = 1 Or Nvl(r_Moneyrow.�����־, 0) = 2 Then
      n_Ԥ����� := r_Moneyrow.�����־;
    Elsif Nvl(r_Moneyrow.��ҳid, 0) = 0 Or Nvl(r_Moneyrow.�����־, 0) = 4 Then
      --���:���ﲡ��
      n_Ԥ����� := 1;
    Else
      n_Ԥ����� := 2;
    End If;
  
    If Nvl(r_Moneyrow.�����־, 0) <> 4 Then
      Update �������
      Set ������� = Nvl(�������, 0) - r_Moneyrow.���ʽ�� --ע:�µĽ���ID�������Ǹ������
      Where ����id = r_Moneyrow.����id And ���� = n_Ԥ����� And ���� = 1
      Returning ������� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (r_Moneyrow.����id, 1, n_Ԥ�����, 0, -1 * r_Moneyrow.���ʽ��);
        n_����ֵ := -1 * r_Moneyrow.���ʽ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete �������
        Where ����id = r_Moneyrow.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) - r_Moneyrow.���ʽ��
    Where ����id = r_Moneyrow.����id And Nvl(��ҳid, 0) = Nvl(r_Moneyrow.��ҳid, 0) And
          Nvl(���˲���id, 0) = Nvl(r_Moneyrow.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Moneyrow.���˿���id, 0) And
          Nvl(��������id, 0) = Nvl(r_Moneyrow.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Moneyrow.ִ�в���id, 0) And
          ������Ŀid + 0 = r_Moneyrow.������Ŀid And ��Դ;�� + 0 = r_Moneyrow.�����־;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (r_Moneyrow.����id, Decode(r_Moneyrow.��ҳid, Null, Null, 0, Null, r_Moneyrow.��ҳid),
         Decode(r_Moneyrow.���˲���id, Null, Null, 0, Null, r_Moneyrow.���˲���id), r_Moneyrow.���˿���id, r_Moneyrow.��������id,
         r_Moneyrow.ִ�в���id, r_Moneyrow.������Ŀid, r_Moneyrow.�����־, -1 * r_Moneyrow.���ʽ��);
    End If;
  
  End Loop;
  Close c_Pati;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽��ʼ�¼_Cancel;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_���˽��ʼ�¼_Delete
(
  No_In           ���˽��ʼ�¼.No%Type,
  ����Ա���_In   ���˽��ʼ�¼.����Ա���%Type,
  ����Ա����_In   ���˽��ʼ�¼.����Ա����%Type,
  �����_In     ����Ԥ����¼.��Ԥ��%Type := 0, --ҽ����Ԥ�����ֽ���������
  �������Ͻ���_In Varchar2 := Null, --���㷽ʽ|������|�������||......
  Ԥ�����ֽ�_In   Number := 0, --��Ԥ�������ֽ�ʱ�����㷽ʽ�����ͨ�������������Ͻ���_In����
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  ����ʱ��_In     Date := Null,
  ��Ԥ��id_In     ����Ԥ����¼.Id%Type := Null, --������ʱ����صĽ���ֵ��Ԥ����ʱ��д
  Ʊ�ݺ�_In       ���˽��ʼ�¼.ʵ��Ʊ��%Type := Null,
  ����id_In       Ʊ�����ü�¼.Id%Type := Null,
  Ʊ��_In         Ʊ��ʹ����ϸ.Ʊ��%Type := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  --���α�����Ԥ����¼�����Ϣ
  Cursor c_Deposit(v_Id ����Ԥ����¼.����id%Type) Is
    Select ����id, ��¼����, ���㷽ʽ, ��Ԥ��, Ԥ����� From ����Ԥ����¼ Where ����id = v_Id;
  r_Depositrow c_Deposit%RowType;

  --���α����ڴ��������ػ��ܱ�
  Cursor c_Money(v_Id ����Ԥ����¼.����id%Type) Is
    Select NO, ��������id, ���˿���id, ִ�в���id, ���˲���id, ����id, ��ҳid, ������Ŀid, �����־, ���ʽ��
    From סԺ���ü�¼
    Where ����id = v_Id
    Union All
    Select NO, ��������id, ���˿���id, ִ�в���id, 0 As ���˲���id, ����id, 0 As ��ҳid, ������Ŀid, �����־, ���ʽ��
    From ������ü�¼
    Where ����id = v_Id;

  r_Moneyrow c_Money%RowType;

  --���α�������˵������Ϣ
  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����, a.�Ա�, a.����, a.סԺ��, a.�����, b.��ҳid, b.��Ժ����, b.��ǰ����id, b.��Ժ����id, Nvl(b.�ѱ�, a.�ѱ�) As �ѱ�, a.����, c.���� As ���ʽ
    From ������Ϣ A, ������ҳ B, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.����id = b.����id(+) And Nvl(a.��ҳid, 0) = b.��ҳid(+) And a.ҽ�Ƹ��ʽ = c.����(+);
  r_Pati c_Pati%RowType;

  --���̱���
  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(50);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_ʵ��Ʊ�� ����Ԥ����¼.ʵ��Ʊ��%Type;
  v_���no   סԺ���ü�¼.No%Type;
  v_���     ���㷽ʽ.����%Type;
  n_����id   ������Ϣ.����id%Type;

  n_ԭid   ���˽��ʼ�¼.Id%Type;
  n_����id ���˽��ʼ�¼.Id%Type;
  n_��ӡid Ʊ�ݴ�ӡ����.Id%Type;

  n_��Դ     Number; --1-����;2-סԺ;3-�����סԺ
  n_����ֵ   �������.Ԥ�����%Type;
  n_��id     ����ɿ����.Id%Type;
  n_Ԥ����� Number;
  d_Date     Date;
  v_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
Begin
  n_��id := Zl_Get��id(����Ա����_In);

  Select ���� Into v_��� From ���㷽ʽ Where ���� = 9 And Rownum = 1;

  Begin
    Select ID, ����id, ʵ��Ʊ�� Into n_ԭid, n_����id, v_ʵ��Ʊ�� From ���˽��ʼ�¼ Where ��¼״̬ = 1 And NO = No_In;
    --���һ�δ�ӡ������
    Select Max(ID)
    Into n_��ӡid
    From (Select b.Id
           From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
           Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 3 And b.No = No_In
           Order By a.ʹ��ʱ�� Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      Begin
        v_Err_Msg := 'û�з���Ҫ���ϵĽ��ʵ���,�����Ѿ����ϣ�';
        Raise Err_Item;
      End;
  End;

  Open c_Pati(n_����id);
  Fetch c_Pati
    Into r_Pati; --���ϵͳ���ô˹���,�������ʱû�в�����Ϣ

  d_Date := ����ʱ��_In;
  If d_Date Is Null Then
    Select Sysdate Into d_Date From Dual;
  End If;
  n_����id := ����id_In;
  If Nvl(n_����id, 0) = 0 Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  End If;

  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 3, No_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��_In, Ʊ�ݺ�_In, 1, 6, ����id_In, v_��ӡid, d_Date, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;

  --���˽��ʼ�¼
  Insert Into ���˽��ʼ�¼
    (ID, NO, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, ��ʼ����, ��������, �շ�ʱ��, ��ע, ԭ��, �ɿ���id, ��������)
    Select n_����id, NO, ʵ��Ʊ��, 2, ��;����, ����id, ����Ա���_In, ����Ա����_In, ��ʼ����, ��������, d_Date, ��ע, ԭ��, n_��id, ��������
    From ���˽��ʼ�¼
    Where ID = n_ԭid;

  Update ���˽��ʼ�¼ Set ��¼״̬ = 3 Where ID = n_ԭid;

  --�����ջ�Ʊ��(������ǰû��ʹ��Ʊ��,�޷��ջ�)
  If n_��ӡid Is Not Null Then
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
      Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
      From Ʊ��ʹ����ϸ
      Where ��ӡid = n_��ӡid And Ʊ�� In (1, 3) And ���� = 1;
  End If;

  --����Ԥ����¼(��Ԥ�����ɿ�)
  If �������Ͻ���_In Is Null Then
    For c_Ԥ�� In (Select ����Ԥ����¼_Id.Nextval As Ԥ��id, ID, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)) As ��¼����,
                        ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, ��Ԥ��, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��,
                        ����˵��, ������λ
                 From ����Ԥ����¼
                 Where ����id = n_ԭid And (��¼���� In (1, 11) And Nvl(��Ԥ��, 0) <> 0 Or ��¼���� Not In (1, 11))) Loop
    
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
      Values
        (c_Ԥ��.Ԥ��id, c_Ԥ��.No, c_Ԥ��.ʵ��Ʊ��, c_Ԥ��.��¼����, c_Ԥ��.��¼״̬, c_Ԥ��.����id, c_Ԥ��.��ҳid, c_Ԥ��.����id, Null, c_Ԥ��.���㷽ʽ,
         c_Ԥ��.�������, c_Ԥ��.ժҪ, c_Ԥ��.�ɿλ, c_Ԥ��.��λ������, c_Ԥ��.��λ�ʺ�, d_Date, ����Ա����_In, ����Ա���_In, -1 * c_Ԥ��.��Ԥ��, n_����id, n_��id,
         c_Ԥ��.Ԥ�����, c_Ԥ��.�����id, c_Ԥ��.���㿨���, c_Ԥ��.����, c_Ԥ��.������ˮ��, c_Ԥ��.����˵��, c_Ԥ��.������λ, 2);
    
      --���ѿ�����
      For c_��¼ In (Select c.�ӿڱ��, c.���ѿ�id, c.����, -1 * Sum(c.Ӧ�ս��) As ������
                   From ���˿������¼ C
                   Where c.����id = c_Ԥ��.Id And c.��¼״̬ = 1
                   Group By c.�ӿڱ��, c.���ѿ�id, c.����) Loop
        Zl_���˿������¼_�˿�(c_��¼.�ӿڱ��, c_��¼.����, c_��¼.���ѿ�id, c_��¼.������, c_Ԥ��.Id, c_Ԥ��.Ԥ��id, ����Ա���_In, ����Ա����_In, d_Date);
      End Loop;
    End Loop;
  Else
    --1.�ȴ����Ԥ������
    If Ԥ�����ֽ�_In = 0 Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ����id, ��ҳid, ����id,
               Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date, ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2
        From ����Ԥ����¼
        Where ����id = n_ԭid And ��¼���� In (1, 11) And Nvl(��Ԥ��, 0) <> 0;
    End If;
  
    --2.�ٴ�����ʽ���,����ҽ���ͷ�ҽ��
    v_�������� := �������Ͻ���_In || ' ||'; --�Կո�ֿ���|��β,û�н�������
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_������� := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, No_In, v_ʵ��Ʊ��, 12, 1, n_����id, r_Pati.��ҳid, r_Pati.��Ժ����id, Null, v_���㷽ʽ, v_�������, '���������˿�',
         Null, Null, Null, d_Date, ����Ա����_In, ����Ա���_In, -1 * n_������, n_����id, n_��id, 2);
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;
  --ȷ�����ʵķ��ü�¼��Դ
  Begin
    Select Case
             When Nvl(Max(סԺ), 0) = 1 And Nvl(Max(����), 0) = 1 Then
              3
             When Nvl(Max(סԺ), 0) = 1 Then
              2
             Else
              1
           End
    Into n_��Դ
    From (Select 1 As סԺ, 0 As ����
           From סԺ���ü�¼
           Where ����id = n_ԭid And Rownum = 1
           Union All
           Select 0 As סԺ, 1 As ����
           From ������ü�¼
           Where ����id = n_ԭid And Rownum = 1);
  
  Exception
    When Others Then
      n_��Դ := 3;
  End;

  If �����_In <> 0 Then
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� + �����_In
    Where NO = No_In And ��¼���� = 12 And ��¼״̬ = 1 And ����id = n_����id;
    If Sql%RowCount = 0 Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, No_In, v_ʵ��Ʊ��, 12, 1, n_����id, r_Pati.��ҳid, r_Pati.��Ժ����id, Null, v_���, Null, '���������˿�', Null,
         Null, Null, d_Date, ����Ա����_In, ����Ա���_In, �����_In, n_����id, n_��id, 2);
    End If;
  End If;

  If n_��Դ = 2 Or n_��Դ = 3 Then
    --���Ͻ��ʶ�Ӧ�ķ��ü�¼:������ԭʼ���ʲ����������Ŀ
    Insert Into סԺ���ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id,
       ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id,
       ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ,
       �ɿ���id, ҽ��С��id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�,
             ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����,
             �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Null, Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����,
             ִ��ʱ��, ����Ա����, ����Ա���, -1 * ���ʽ��, n_����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id, ҽ��С��id
      From סԺ���ü�¼
      Where ����id = n_ԭid And Nvl(���ӱ�־, 0) <> 9;
  End If;

  If n_��Դ = 1 Or n_��Դ = 3 Then
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
       �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
       ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id, ��ҳid, ���˲���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, To_Number('1' || Substr(��¼����, Length(��¼����), 1)), ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id,
             ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����,
             ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Null, Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���,
             -1 * ���ʽ��, n_����id, ������Ŀ��, ���մ���id, ͳ����, �Ƿ���, ���ձ���, ��������, ժҪ, �ɿ���id, ��ҳid, ���˲���id
      From ������ü�¼
      Where ����id = n_ԭid And Nvl(���ӱ�־, 0) <> 9;
  End If;
  --��ػ��ܱ���
  For r_Depositrow In c_Deposit(n_����id) Loop
    If r_Depositrow.��¼���� In (1, 11) Then
    
      --�������(Ԥ��)
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - r_Depositrow.��Ԥ�� --ע:�µĽ���ID�������Ǹ������
      Where ����id = r_Depositrow.����id And ���� = Nvl(r_Depositrow.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (r_Depositrow.����id, 1, Nvl(r_Depositrow.Ԥ�����, 2), -1 * r_Depositrow.��Ԥ��, 0);
        n_����ֵ := -1 * r_Depositrow.��Ԥ��;
      End If;
    
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete �������
        Where ���� = 1 And ����id = r_Depositrow.����id And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    
    Else
      --��Ա�ɿ����,ҽ����֧�����ϵĽ��㷽ʽ���µ�Ԥ���������ѱ�����Ϊ�����ֽ�,
      --�˴��ü�,��ʾ�ջ��˸����˵��ֽ�(����ʱ,�˿��Ǹ�,����ʱ����)
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + r_Depositrow.��Ԥ��
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Depositrow.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Depositrow.���㷽ʽ, 1, r_Depositrow.��Ԥ��);
        n_����ֵ := -1 * r_Depositrow.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Depositrow.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End If;
  End Loop;

  For r_Moneyrow In c_Money(n_����id) Loop
    --������� ,������ѽ���,���Բ���Ҫ�������������ܱ�
    If Nvl(v_���no, 'sc') <> Nvl(r_Moneyrow.No, 'sc') Then
      If Nvl(r_Moneyrow.�����־, 0) = 1 Or Nvl(r_Moneyrow.�����־, 0) = 2 Then
        n_Ԥ����� := r_Moneyrow.�����־;
      Elsif Nvl(r_Moneyrow.��ҳid, 0) = 0 Or Nvl(r_Moneyrow.�����־, 0) = 4 Then
        --���:���ﲡ��
        n_Ԥ����� := 1;
      Else
        n_Ԥ����� := 2;
      End If;
    
      If Nvl(r_Moneyrow.�����־, 0) <> 4 Then
        Update �������
        Set ������� = Nvl(�������, 0) - r_Moneyrow.���ʽ�� --ע:�µĽ���ID�������Ǹ������
        Where ����id = r_Moneyrow.����id And ���� = n_Ԥ����� And ���� = 1
        Returning ������� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, ����, Ԥ�����, �������)
          Values
            (r_Moneyrow.����id, 1, n_Ԥ�����, 0, -1 * r_Moneyrow.���ʽ��);
          n_����ֵ := -1 * r_Moneyrow.���ʽ��;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete �������
          Where ����id = r_Moneyrow.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - r_Moneyrow.���ʽ��
      Where ����id = r_Moneyrow.����id And Nvl(��ҳid, 0) = Nvl(r_Moneyrow.��ҳid, 0) And
            Nvl(���˲���id, 0) = Nvl(r_Moneyrow.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Moneyrow.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(r_Moneyrow.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Moneyrow.ִ�в���id, 0) And
            ������Ŀid + 0 = r_Moneyrow.������Ŀid And ��Դ;�� + 0 = r_Moneyrow.�����־;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (r_Moneyrow.����id, Decode(r_Moneyrow.��ҳid, Null, Null, 0, Null, r_Moneyrow.��ҳid),
           Decode(r_Moneyrow.���˲���id, Null, Null, 0, Null, r_Moneyrow.���˲���id), r_Moneyrow.���˿���id, r_Moneyrow.��������id,
           r_Moneyrow.ִ�в���id, r_Moneyrow.������Ŀid, r_Moneyrow.�����־, -1 * r_Moneyrow.���ʽ��);
      End If;
    End If;
  End Loop;

  If Nvl(��Ԥ��id_In, 0) <> 0 Then
    --����ʱ���˿����ֵ��Ԥ�����ʻ�,��������Ǳ��ν��ʽɴ��
    Update ����Ԥ����¼ Set ����id = ����id_In Where ID = ��Ԥ��id_In And ����id Is Null;
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ���Ӧ��Ԥ�����¼��';
      Raise Err_Item;
    End If;
  End If;
  Close c_Pati;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽��ʼ�¼_Delete;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_���ʷ��ü�¼_Insert
(
  Id_In       סԺ���ü�¼.Id%Type,
  No_In       סԺ���ü�¼.No%Type,
  ��¼����_In סԺ���ü�¼.��¼����%Type,
  ��¼״̬_In סԺ���ü�¼.��¼״̬%Type,
  ִ��״̬_In סԺ���ü�¼.ִ��״̬%Type,
  ���_In     סԺ���ü�¼.���%Type,
  ���ʽ��_In סԺ���ü�¼.���ʽ��%Type,
  ����id_In   סԺ���ü�¼.����id%Type
) As
  n_Next_Id    סԺ���ü�¼.Id%Type;
  n_����id     סԺ���ü�¼.����id%Type;
  n_��ҳid     סԺ���ü�¼.��ҳid%Type;
  n_���˲���id סԺ���ü�¼.���˲���id%Type;
  n_���˿���id סԺ���ü�¼.���˿���id%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  n_ִ�в���id סԺ���ü�¼.ִ�в���id%Type;
  n_������Ŀid סԺ���ü�¼.������Ŀid%Type;
  n_�����־   סԺ���ü�¼.�����־%Type;
  n_���ʷ���   סԺ���ü�¼.���ʷ���%Type;
  v_����Ա     סԺ���ü�¼.����Ա����%Type;
  v_����Ա���� סԺ���ü�¼.����Ա����%Type;

  n_���ʽ�� סԺ���ü�¼.���ʽ��%Type;
  n_ʵ�ս�� סԺ���ü�¼.ʵ�ս��%Type;
  n_����ֵ   �������.Ԥ�����%Type;
  n_���     Number(18);
  v_Temp     Varchar2(500);

  Err_Custom  Exception;
  Err_Special Exception;
  v_Error Varchar2(255);
  n_��Դ  Number;
Begin
  --��Աid,��Ա���,��Ա����
  v_Temp := Zl_Identity(1);
  If Not (Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_') Then
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_����Ա���� := v_Temp;
  End If;

  If Id_In <> 0 Then
    Begin
      Select 2 Into n_��Դ From סԺ���ü�¼ Where ID = Id_In;
    Exception
      When Others Then
        n_��Դ := 1;
    End;
  
    --��һ�ν��ʵ����ֽ�
    If n_��Դ = 1 Then
      Update ������ü�¼ Set ���ʽ�� = ���ʽ��_In, ����id = ����id_In Where ID = Id_In And ����id Is Null;
    Else
      Update סԺ���ü�¼ Set ���ʽ�� = ���ʽ��_In, ����id = ����id_In Where ID = Id_In And ����id Is Null;
    End If;
  
    If Sql%RowCount = 0 Then
      If n_��Դ = 1 Then
        Select Max(b.����Ա����)
        Into v_����Ա
        From ������ü�¼ A, ���˽��ʼ�¼ B
        Where a.Id = Id_In And b.Id = a.����id;
      Else
        Select Max(b.����Ա����)
        Into v_����Ա
        From סԺ���ü�¼ A, ���˽��ʼ�¼ B
        Where a.Id = Id_In And b.Id = a.����id;
      End If;
      If v_����Ա Is Null Then
        v_Error := 'δ���ֽ��ʵķ���,��ǰ���ʲ������ܼ�����';
        Raise Err_Custom;
      Else
        If v_����Ա���� = v_����Ա Then
          v_Error := '�����Ѿ������ʵķ���,��ǰ���ʲ������ܼ�����';
          Raise Err_Special;
        Else
          v_Error := '�����Ѿ��������˽��ʵķ���,��ǰ���ʲ������ܼ�����';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  
    n_Next_Id := Id_In;
  Else
    --����ǰ������
    Select ���˷��ü�¼_Id.Nextval Into n_Next_Id From Dual;
  
    If Mod(��¼����_In, 10) = 3 Or Mod(��¼����_In, 10) = 5 Then
      --�Զ����ʻ���￨;�϶���סԺ
      n_��Դ := 2;
    Else
      Begin
        Select 2
        Into n_��Դ
        From סԺ���ü�¼
        Where NO = No_In And ��� = ���_In And ��¼״̬ In (1, 2, 3) And Nvl(ִ��״̬, 0) = Nvl(ִ��״̬_In, 0) And
              Substr(��¼����, Length(��¼����), 1) = ��¼����_In And Rownum < 2;
      Exception
        When Others Then
          n_��Դ := 1;
      End;
    End If;
  
    If n_��Դ = 1 Then
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
         �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
         ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, ���ձ���, ��������, �Ƿ���, ժҪ, ��ҳid, ���˲���id)
        Select n_Next_Id, NO, ʵ��Ʊ��, To_Number('1' || ��¼����_In), ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ����, �Ա�, ����,
               ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Null,
               Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��_In, ����id_In, ������Ŀ��,
               ���մ���id, ͳ����, ���ձ���, ��������, �Ƿ���, ժҪ, ��ҳid, ���˲���id
        From ������ü�¼
        Where NO = No_In And ��� = ���_In And (��¼״̬ = ��¼״̬_In Or ��¼״̬ = Decode(��¼״̬_In, 1, 3, ��¼״̬_In)) And
              Nvl(ִ��״̬, 0) = Nvl(ִ��״̬_In, 0) And Substr(��¼����, Length(��¼����), 1) = ��¼����_In And Rownum < 2;
    
      --����ν��ʺ���ʽ���Ƿ����ԭ���
      Select Nvl(Sum(ʵ�ս��), 0), Nvl(Sum(���ʽ��), 0)
      Into n_ʵ�ս��, n_���ʽ��
      From ������ü�¼
      Where NO = No_In And ��� = ���_In And (��¼״̬ = ��¼״̬_In Or ��¼״̬ = Decode(��¼״̬_In, 1, 3, ��¼״̬_In)) And
            Substr(��¼����, Length(��¼����), 1) = ��¼����_In And Nvl(ִ��״̬, 0) = ִ��״̬_In;
    Else
      Insert Into סԺ���ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id,
         ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������,
         ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��, ����id, ������Ŀ��, ���մ���id, ͳ����, ���ձ���, ��������,
         �Ƿ���, ժҪ, ҽ��С��id)
        Select n_Next_Id, NO, ʵ��Ʊ��, To_Number('1' || ��¼����_In), ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־,
               ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ���ʷ���, ������Ŀid,
               �վݷ�Ŀ, ��׼����, Null, Null, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ����, ִ��ʱ��, ����Ա����, ����Ա���, ���ʽ��_In,
               ����id_In, ������Ŀ��, ���մ���id, ͳ����, ���ձ���, ��������, �Ƿ���, ժҪ, ҽ��С��id
        From סԺ���ü�¼
        Where NO = No_In And ��� = ���_In And (��¼״̬ = ��¼״̬_In Or ��¼״̬ = Decode(��¼״̬_In, 1, 3, ��¼״̬_In)) And
              Nvl(ִ��״̬, 0) = Nvl(ִ��״̬_In, 0) And Substr(��¼����, Length(��¼����), 1) = ��¼����_In And Rownum < 2;
    
      --����ν��ʺ���ʽ���Ƿ����ԭ���
      Select Nvl(Sum(ʵ�ս��), 0), Nvl(Sum(���ʽ��), 0)
      Into n_ʵ�ս��, n_���ʽ��
      From סԺ���ü�¼
      Where NO = No_In And ��� = ���_In And (��¼״̬ = ��¼״̬_In Or ��¼״̬ = Decode(��¼״̬_In, 1, 3, ��¼״̬_In)) And
            Substr(��¼����, Length(��¼����), 1) = ��¼����_In And Nvl(ִ��״̬, 0) = ִ��״̬_In;
    End If;
  
    If n_���ʽ�� > n_ʵ�ս�� Then
      If n_��Դ = 1 Then
        Select Max(b.����Ա����)
        Into v_����Ա
        From ������ü�¼ A, ���˽��ʼ�¼ B
        Where a.Id = Id_In And b.Id = a.����id;
      Else
        Select Max(b.����Ա����)
        Into v_����Ա
        From סԺ���ü�¼ A, ���˽��ʼ�¼ B
        Where a.Id = Id_In And b.Id = a.����id;
      End If;
    
      If v_����Ա Is Null Then
        v_Error := 'δ���ֽ��ʵķ���,��ǰ���ʲ������ܼ�����';
        Raise Err_Custom;
      Else
        If v_����Ա���� = v_����Ա Then
          v_Error := '�����Ѿ������ʵķ���,��ǰ���ʲ������ܼ�����';
          Raise Err_Special;
        Else
          v_Error := '�����Ѿ��������˽��ʵķ���,��ǰ���ʲ������ܼ�����';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  End If;
  If n_��Դ = 1 Then
    Select ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, �����־, ���ʷ���
    Into n_����id, n_��ҳid, n_���˲���id, n_���˿���id, n_��������id, n_ִ�в���id, n_������Ŀid, n_�����־, n_���ʷ���
    From ������ü�¼
    Where ID = n_Next_Id;
    n_��� := 1;
  Else
    Select ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, �����־, ���ʷ���
    Into n_����id, n_��ҳid, n_���˲���id, n_���˿���id, n_��������id, n_ִ�в���id, n_������Ŀid, n_�����־, n_���ʷ���
    From סԺ���ü�¼
    Where ID = n_Next_Id;
  
    If Nvl(n_�����־, 0) = 1 Or Nvl(n_�����־, 0) = 2 Then
      n_��� := n_�����־;
    Elsif Nvl(n_��ҳid, 0) = 0 Or Nvl(n_�����־, 0) = 4 Then
      n_��� := 1;
    Else
      n_��� := 2;
    End If;
  End If;

  If Nvl(n_�����־, 0) <> 4 Then
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) - ���ʽ��_In
    Where ����id = n_����id And ���� = 1 And ���� = n_���
    Returning ������� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, n_���, 0, -1 * ���ʽ��_In);
      n_����ֵ := -1 * ���ʽ��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0 And ����id = n_����id;
    End If;
  End If;

  --����δ�����
  Update ����δ�����
  Set ��� = Nvl(���, 0) - ���ʽ��_In
  Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(n_��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(n_���˲���id, 0) And
        Nvl(���˿���id, 0) = Nvl(n_���˿���id, 0) And Nvl(��������id, 0) = Nvl(n_��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(n_ִ�в���id, 0) And
        ������Ŀid + 0 = n_������Ŀid And ��Դ;�� + 0 = n_�����־
  Returning ��� Into n_����ֵ;
  If Sql%RowCount = 0 Then
    Insert Into ����δ�����
      (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
    Values
      (n_����id, Decode(n_��ҳid, 0, Null, n_��ҳid), Decode(n_���˲���id, 0, Null, n_���˲���id), n_���˿���id, n_��������id, n_ִ�в���id,
       n_������Ŀid, n_�����־, -1 * ���ʽ��_In);
    n_����ֵ := -1 * ���ʽ��_In;
  End If;
  If Nvl(n_����ֵ, 0) = 0 Then
    Delete From ����δ����� Where ����id = n_����id And Nvl(���, 0) = 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ʷ��ü�¼_Insert;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_�����շѼ�¼_����
(
  No_In         ������ü�¼.No%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ���_In       Varchar2 := Null,
  �˷�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �˷�ժҪ_In   ������ü�¼.ժҪ%Type := Null,
  ����id_In     ����Ԥ����¼.����id%Type := Null,
  ����Ʊ��_In   Number := 0
) As
  --���ܣ�ɾ��һ�������շѵ���
  --������
  --        ���_IN           =Ҫ�˷ѵ���Ŀ���,��ʽΪ"1,3,5,6...",ȱʡNULL��ʾ��"δ�˵�"�����С�
  --        ����Ʊ��_In       =0:ȫ�˻����һ��ȫ��ʱ,�ջ�Ʊ�ݡ�
  --                           1:�����˷Ѳ�����Ʊ��,ͨ���ش���õ�������
  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼

  --ҽ��ȫ�˵�ĳ�ֽ������ֽ�Ӷ��������µ����ʱ,�ſ��˴�������,ִ���걾���̺�,��������е������������
  Cursor c_Bill Is
    Select a.Id, a.No, a.���ӱ�־, a.�շ�ϸĿid, a.���, a.�۸񸸺�, a.ִ��״̬, a.�շ����, a.����, a.����, a.ҽ�����, j.�������, m.��������,
           Nvl(a.���ӱ�־, 0) As ���, Nvl(j.ҽ��״̬, 0) As ҽ��״̬
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.No = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.�շ�ϸĿid + 0 = m.����id(+)
    Order By a.�շ�ϸĿid, a.���;

  --:����ԭʼ�������,��Ӧ�ø��ݵ�ǰ�˷Ѳ������������д���
  -- Decode(Sign(���_In), 0, 999, 9)

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��

  n_����id ������ü�¼.����id%Type;
  n_��ӡid Ʊ�ݴ�ӡ����.Id%Type;

  --�����˷Ѽ������
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;
  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;
  n_�ܽ��   Number;
  n_��id     ����ɿ����.Id%Type;

  l_ʹ��id   t_Numlist := t_Numlist();
  l_���     t_Numlist := t_Numlist();
  l_ִ��״̬ t_Numlist := t_Numlist();

  n_Dec   Number;
  d_Date  Date;
  n_Count Number;

  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_����ģʽ Number(3);
  v_Para     Varchar2(1000);

Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�Ǹõ������ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  --ִ��״̬��ԭʼ��¼���ж�
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����, ����id
                From ������ü�¼
                Where NO = No_In And Mod(��¼����, 10) = 1 And Nvl(���ӱ�־, 0) <> 9 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���), ����id)
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --���ñ���
  If �˷�ʱ��_In Is Not Null Then
    d_Date := �˷�ʱ��_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;

  If ����id_In Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Else
    n_����id := ����id_In;
  End If;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --ѭ������ÿ�з���(������Ŀ��)
  n_�ܽ�� := 0;
  For r_Bill In c_Bill Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
        From ������ü�¼
        Where NO = No_In And Mod(��¼����, 10) = 1 And ��� = r_Bill.���;
      
        If n_ʣ������ = 0 Then
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ���˷ѣ�';
            Raise Err_Item;
          End If;
        Else
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            --@@@
            --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
            --: 1.����ҽ��ִ�мƼ۵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ),��ִ�еĲ������˷�
            --: 2.������ҽ��ִ�мƼ۵�,����ʣ������Ϊ׼
            --: 3.ҽ�������˵�,����ʣ������Ϊ׼(����ҽ����¼.ҽ��״̬=4��ʾ����ҽ������ɾ��"����ҽ������",����ҩ�������Ϻ���ҩ)
            --: 4.����ҽ������.ִ��״̬=1�����ִ�У�ʱ��׼����Ϊ0�����ٸ���ҽ��ִ�мƼ���ͳ��׼����
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null And r_Bill.ҽ��״̬ <> 4 Then
              Select Nvl(Sum(Decode(b.ִ��״̬, 1, 0, 1) * Decode(c.ִ��״̬, 0, 1, 0) * c.����), 0), Count(1)
              Into n_׼������, n_Count
              From ����ҽ������ B, ҽ��ִ�мƼ� C
              Where b.ҽ��id = r_Bill.ҽ����� And b.No = r_Bill.No And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And
                    c.�շ�ϸĿid + 0 = r_Bill.�շ�ϸĿid And b.��¼���� = 1;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_׼������ = 0 Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з�����ִ�У��������˷ѣ�';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              If r_Bill.ִ��״̬ = 2 Then
                --��ҽ��ִ�мƼ۵Ĳ����˷��޷��ж�׼���������������˷�
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѳ���ִ�У��޷��ж�׼���������������˷ѣ�';
                Raise Err_Item;
              Else
                n_׼������ := n_ʣ������;
              End If;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
            Into n_׼������, n_Count
            From ҩƷ�շ���¼
            Where NO = No_In And ���� In (8, 24) And Mod(��¼״̬, 3) = 1 --@@@
                  And ����� Is Null And ����id = r_Bill.Id;
          
            --��ʣ��������׼�������������������
            --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������
            --2.��������,��ʱ�ѷ�ҩ����
            If n_׼������ = 0 Then
              If r_Bill.�շ���� = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                  Raise Err_Item;
                Else
                  n_׼������ := n_ʣ������;
                End If;
              Else
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          If n_׼������ > n_ʣ������ Then
            v_Err_Msg := '����[' || No_In || '] �е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з��õ��˷�����(' || n_׼������ ||
                         ')������ʣ������(' || n_ʣ������ || ')���������˷ѣ�';
            Raise Err_Item;
          End If;
          --�շѵ�ʱ���Ǹ��������Ĳ����׼�������Ƿ�С����
          If n_׼������ < 0 And Nvl(r_Bill.����, 1) * r_Bill.���� > 0 Then
            v_Err_Msg := '����[' || No_In || '] �е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з��õ��˷�����(' || n_׼������ ||
                         ')С�����㣬�������˷ѣ�';
            Raise Err_Item;
          End If;
        
          --�ñ���Ŀ�ڼ����˷�
          Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
          Into n_�˷Ѵ���
          From ������ü�¼
          Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 2 And Nvl(ִ��״̬, 0) < 0 And ��� = r_Bill.���;
        
          --���=ʣ����*(׼����/ʣ����)
          n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
          n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
          n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
          n_�ܽ��   := n_�ܽ�� + n_ʵ�ս��;
        
          --�����˷Ѽ�¼
          Insert Into ������ü�¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����,
             ִ��״̬, ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����,
             �ɿ���id, �Һ�id, ��ҳid, ���˲���id)
            Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                   ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                   Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����,
                   -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, 1, ִ��ʱ��, ����Ա���_In, ����Ա����_In,
                   ����ʱ��, d_Date, n_����id, -1 * n_ʵ�ս��, ������Ŀ��, ���մ���id, -1 * n_ͳ����, Nvl(�˷�ժҪ_In, ժҪ),
                   Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, ����, n_��id, �Һ�id, ��ҳid, ���˲���id
            From ������ü�¼
            Where ID = r_Bill.Id;
        
          --���ԭ���ü�¼
          l_���.Extend;
          l_���(l_���.Count) := r_Bill.���;
          l_ִ��״̬.Extend;
          l_ִ��״̬(l_ִ��״̬.Count) := Case
                                    When Sign(n_׼������ - n_ʣ������) = 0 Then
                                     0
                                    Else
                                     1
                                  End;
        End If;
      Else
        If ���_In Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�����˷ѣ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End Loop;
  --���ԭ���ü�¼
  Forall I In 1 .. l_���.Count
    Update ������ü�¼
    Set ��¼״̬ = 3, ִ��״̬ = l_ִ��״̬(I)
    Where Mod(��¼����, 10) = 1 And NO = No_In And ��� = l_���(I) And ��¼״̬ In (1, 3);

  l_���.Delete;
  For c_���� In (Select Distinct b.����id
               From ������ü�¼ A, ����Ԥ����¼ B
               Where a.����id = b.����id And a.No = No_In And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And
                     Nvl(b.��¼״̬, 0) = 1) Loop
    l_���.Extend;
    l_���(l_���.Count) := c_����.����id;
  End Loop;

  Forall I In 1 .. l_���.Count
    Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ����id = l_���(I) And Mod(��¼����, 10) <> 1;

  ---------------------------------------------------------------------------------
  --�˷�Ʊ�ݻ���(��ȫ��ʱ�Ż���,�����������ش�����л���)
  If ����Ʊ��_In = 1 Then
  
    --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
    v_Para     := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
    n_����ģʽ := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_����ģʽ <> 0 Then
      --�ջ�Ʊ��
      Select ʹ��id
      Bulk Collect
      Into l_ʹ��id
      From (Select Distinct b.ʹ��id From Ʊ�ݴ�ӡ��ϸ B Where b.No = No_In And Nvl(b.Ʊ��, 0) = 1);
    
      n_����ģʽ := l_ʹ��id.Count;
      If l_ʹ��id.Count <> 0 Then
        --������ռ�¼
        Forall I In 1 .. l_ʹ��id.Count
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, ����Ա����_In, d_Date, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ A
            Where ID = l_ʹ��id(I) And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = a.Ʊ�� And Nvl(����, 0) <> 1);
      
        Forall I In 1 .. l_ʹ��id.Count
          Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I) And Nvl(�Ƿ����, 0) = 0;
      
      End If;
    End If;
    If n_����ģʽ = 0 Then
      --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ)
      Begin
        --����=1��ԭ��=6Ϊ�˷Ѵ�ӡƱ��(��Ʊ)��������
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = No_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --������ǰû�д�ӡ,���ջ�
      If n_��ӡid Is Not Null Then
        --a.���ŵ���ѭ������ʱֻ���ջ�һ��
        Select Count(*) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        Else
          --b.�����˷Ѷ���ջ�ʱ,���һ��ȫ���ջ�Ҫ�ſ����ջص�
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ A
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 2);
        End If;
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --ҩƷ�����������
  --���밴�ա��շ�ϸĿid���������򣬷�ֹ��������ҩƷ��桱��
  For r_Expenses In (Select ID
                     From ������ü�¼
                     Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And �շ���� In ('4', '5', '6', '7') And
                           (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null)
                     Order By �շ�ϸĿid) Loop
    Zl_ҩƷ�շ���¼_�����˷�(r_Expenses.Id);
  End Loop;

  --ҽ������
  --ɾ������ҽ������(���һ��ɾ��ʱ)
  For c_ҽ�� In (Select Distinct ҽ�����
               From ������ü�¼
               Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, ִ��״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
                  Where Mod(��¼����, 10) = 1 And Nvl(���ӱ�־, 0) <> 9 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, ִ��״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 1 And NO = No_In;
    End If;
  End Loop;

  --����ҽ��ִ�мƼ�.ִ��״̬ NULL-��ʷ���ݣ�0-δִ�У�1-��ִ�У�2-���˷�
  For c_���� In (Select Distinct a.ҽ����� As ҽ��id, a.�շ�ϸĿid, b.���ͺ�
               From ������ü�¼ A, ����ҽ������ B
               Where a.ҽ����� = b.ҽ��id And a.No = b.No And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.No = No_In And
                     (Instr(',' || ���_In || ',', ',' || a.��� || ',') > 0 Or ���_In Is Null) And a.�۸񸸺� Is Null And
                     b.��¼���� = 1) Loop
    Update ҽ��ִ�мƼ�
    Set ִ��״̬ = 2
    Where ҽ��id = c_����.ҽ��id And ���ͺ� = c_����.���ͺ� And �շ�ϸĿid = c_����.�շ�ϸĿid And ִ��״̬ = 0;
  End Loop;

  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In Varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_����;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_�����շѼ�¼_����
(
  ԭ����id_In     ������ü�¼.����id%Type,
  ����id_In       ������ü�¼.����id%Type,
  ���ս���id_In   ������ü�¼.����id%Type,
  �ſ�ҽ������_In Varchar2 := Null
) As
  --�ſ�ҽ������_IN:����ö��ŷ���(ֻĳЩҽ������,�������ֽ�)
  Cursor c_Fee_Data Is
    Select ID
    From ������ü�¼ A
    Where ����id = ԭ����id_In And Not Exists
     (Select 1
           From ������ü�¼ B
           Where Mod(b.��¼����, 10) = 1 And a.No = b.No And a.��� = b.��� And ����id = ����id_In)
    Order By ID;

  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ������ü�¼.�Ǽ�ʱ��%Type;
  n_�ɿ���id   ������ü�¼.�ɿ���id%Type;
  n_����id     ������ü�¼.����id%Type;
  Err_Item Exception;
  v_Err_Msg    Varchar2(255);
  n_Array_Size Number := 200;
  t_����id     t_Numlist;
  n_������   ������ü�¼.ʵ�ս��%Type;
  n_�������   ����Ԥ����¼.��Ԥ��%Type;
  n_Count      Number(18);
Begin
  Begin
    Select ����Ա���, ����Ա����, �Ǽ�ʱ��, �ɿ���id, ����id
    Into v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_�ɿ���id, n_����id
    From ������ü�¼
    Where ����id = ����id_In And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'NO';
  End;

  If Nvl(v_Err_Msg, '-') = 'NO' Then
    v_Err_Msg := '���ڲ�������,�õ��ݿ����Ѿ��������˷ѻ�ɾ��,�����ٽ����˷Ѳ�����';
    Raise Err_Item;
  End If;

  --1.�������ѡ������ǲ����˻򲿷�ִ�е�
  Insert Into ������ü�¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ,
     ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ����״̬, ִ��ʱ��,
     ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id, �Һ�id, ��ҳid, ���˲���id)
    Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
           �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������,
           ִ����, ִ��״̬, ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����,
           �ɿ���id, �Һ�id, ��ҳid, ���˲���id
    From (Select NO, Max(ʵ��Ʊ��) As ʵ��Ʊ��, 11 As ��¼����, 1 As ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ,
                  �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, 1 As ����, Max(��ҩ����) As ��ҩ����, Sum(Nvl(����, 1) * Nvl(����, 0)) As ����,
                  Max(�Ӱ��־) As �Ӱ��־, Max(���ӱ�־) As ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, Avg(��׼����) As ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                  Sum(ʵ�ս��) As ʵ�ս��, ��������id, ������, ִ�в���id, Max(������) As ������, Max(ִ����) ִ����, Max(ִ��״̬) As ִ��״̬, 1 As ����״̬,
                  Max(ִ��ʱ��) ִ��ʱ��, v_����Ա��� As ����Ա���, v_����Ա���� As ����Ա����, ����ʱ��, d_�Ǽ�ʱ�� As �Ǽ�ʱ��, ���ս���id_In As ����id,
                  Sum(���ʽ��) As ���ʽ��, Max(������Ŀ��) As ������Ŀ��, ���մ���id, Sum(ͳ����) As ͳ����,
                  Max(Decode(��¼����, 1, ժҪ, 11, ժҪ, Null)) As ժҪ, 0 As �Ƿ��ϴ�, Max(���ձ���) As ���ձ���, Max(��������) As ��������,
                  Max(Decode(��¼����, 1, ����, 11, ����, Null)) As ����, n_�ɿ���id As �ɿ���id, Max(�Һ�id) As �Һ�id, Max(��ҳid) As ��ҳid,
                  Max(���˲���id) As ���˲���id
           From ������ü�¼
           Where Mod(��¼����, 10) = 1 And (NO, ���) In (Select NO, ��� From ������ü�¼ Where ����id = ����id_In)
           Group By NO, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀid,
                    �վݷ�Ŀ, ���ʷ���, ��������id, ������, ִ�в���id, ����ʱ��, ���մ���id
           Having Sum(Nvl(����, 1) * Nvl(����, 0)) <> 0);

  For c_���� In (Select NO, ���, ��������, �۸񸸺�, ������Ŀid, -1 * Sum(Nvl(����, 1) * Nvl(����, 0)) As ����, Sum(��׼����) As ��׼����,
                      -1 * Sum(Ӧ�ս��) As Ӧ�ս��, -1 * Sum(ʵ�ս��) As ʵ�ս��, -1 * Sum(ͳ����) As ͳ����, -1 * Sum(���ʽ��) As ���ʽ��
               From ������ü�¼
               Where ��¼���� = 11 And ����id = ���ս���id_In
               Group By NO, ���, ��������, �۸񸸺�, ������Ŀid) Loop
    Update ������ü�¼
    Set ���� = Nvl(����, 0) + Nvl(c_����.����, 0), ʵ�ս�� = Nvl(ʵ�ս��, 0) + Nvl(c_����.ʵ�ս��, 0),
        Ӧ�ս�� = Nvl(Ӧ�ս��, 0) + Nvl(c_����.Ӧ�ս��, 0), ���ʽ�� = Nvl(���ʽ��, 0) + Nvl(c_����.���ʽ��, 0),
        ͳ���� = Nvl(ͳ����, 0) + Nvl(c_����.ͳ����, 0)
    Where NO = c_����.No And ��� = c_����.��� And Nvl(��������, -1) = Nvl(c_����.��������, '-1') And
          Nvl(�۸񸸺�, -1) = Nvl(c_����.�۸񸸺�, '-1') And ������Ŀid = c_����.������Ŀid And ����id = ����id_In;
  End Loop;

  --2.�������δѡ�˷Ѳ���,��Ҫȫ���Ҳ���11�����ռ�¼
  Open c_Fee_Data;
  Loop
    Fetch c_Fee_Data Bulk Collect
      Into t_����id Limit n_Array_Size;
    Exit When t_����id.Count = 0;
  
    --�˷Ѽ�¼
    Forall I In 1 .. t_����id.Count
      Insert Into ������ü�¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
         ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ����״̬,
         ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id, �Һ�id, ��ҳid,
         ���˲���id)
        Select ���˷��ü�¼_Id.Nextval, a.No, a.ʵ��Ʊ��, 1, 2, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����,
               a.��ʶ��, a.���ʽ, a.�ѱ�, a.���˿���id, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.����, a.��ҩ����, -1 * a.����, a.�Ӱ��־, a.���ӱ�־,
               a.������Ŀid, a.�վݷ�Ŀ, a.���ʷ���, a.��׼����, -1 * a.Ӧ�ս��, -1 * a.ʵ�ս��, a.��������id, a.������, a.ִ�в���id, a.������, ִ����,
               Nvl(q.ִ��״̬, -1) As ִ��״̬, 1, a.ִ��ʱ��, v_����Ա���, v_����Ա����, a.����ʱ��, d_�Ǽ�ʱ��, ����id_In, -1 * a.���ʽ��, a.������Ŀ��,
               a.���մ���id, -1 * a.ͳ����, a.ժҪ, 0 As �Ƿ��ϴ�, a.���ձ���, a.��������, a.����, n_�ɿ���id As �ɿ���id, �Һ�id, ��ҳid, ���˲���id
        From ������ü�¼ A,
             (Select j.No, j.���, Nvl(Max(j.ִ��״̬), 0) - 1 As ִ��״̬
               From ������ü�¼ M, ������ü�¼ J
               Where m.Id = t_����id(I) And m.No = j.No And m.��� = j.��� And Mod(j.��¼����, 10) = 1 And j.��¼״̬ = 2
               Group By j.No, j.���) Q
        Where ID = t_����id(I) And a.No = q.No(+) And a.��� = q.���(+);
  
    --��ԭ��¼״̬��1��Ϊ3
    Forall I In 1 .. t_����id.Count
      Update ������ü�¼ Set ��¼״̬ = 3 Where ID = t_����id(I) And ��¼״̬ = 1;
  
    --�����շѼ�¼
    If Nvl(���ս���id_In, 0) <> 0 Then
      Forall I In 1 .. t_����id.Count
        Insert Into ������ü�¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
           ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬,
           ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id, �Һ�id,
           ��ҳid, ���˲���id)
          Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, 11, 1, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
                 �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id,
                 ������, ִ����, ִ��״̬, 1, ִ��ʱ��, v_����Ա���, v_����Ա����, ����ʱ��, d_�Ǽ�ʱ��, ���ս���id_In, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ,
                 0 As �Ƿ��ϴ�, ���ձ���, ��������, ����, n_�ɿ���id As �ɿ���id, �Һ�id, ��ҳid, ���˲���id
          From ������ü�¼
          Where ID = t_����id(I);
    End If;
  End Loop;
  Close c_Fee_Data;

  Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
  If n_Count = 0 Then
    --�˷ѽ��㷽ʽ
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
      Select ����Ԥ����¼_Id.Nextval, 3, Null, 2, n_����id, ���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, -1 * ��Ԥ��, ����id_In, n_�ɿ���id,
             -1 * ����id_In, 2, 3
      From ����Ԥ����¼
      Where ����id = ԭ����id_In And ���㷽ʽ In (Select ���� From ���㷽ʽ Where ���� In (3, 4)) And
            Instr(',' || �ſ�ҽ������_In || ',', ',' || ���㷽ʽ || ',') = 0 And Mod(��¼����, 10) <> 1;
    --��ԭ����ȫ������
    --Insert Into ����Ԥ����¼
    --  (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־)
    --  Select ����Ԥ����¼_Id.Nextval, 3, Null, 2, n_����id, ���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, -1 * ��Ԥ��, ����id_In, n_�ɿ���id,
    --         -1 * ����id_In, 2
    --  From ����Ԥ����¼
    --  Where ����id = ԭ����id_In And ���㷽ʽ = v_���� And Mod(��¼����, 10) <> 1;
  
    Select Sum(��Ԥ��) Into n_������� From ����Ԥ����¼ Where ����id = ����id_In;
    Select Sum(���ʽ��) Into n_������ From ������ü�¼ Where ����id = ����id_In;
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 2, n_����id, Null, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, -1 * (Nvl(n_�������, 0) - Nvl(n_������, 0)),
       ����id_In, n_�ɿ���id, -1 * ����id_In, 1, 3);
  
  End If;
  If Nvl(���ս���id_In, 0) <> 0 Then
    Select Count(1) Into n_Count From ����Ԥ����¼ Where ����id = ���ս���id_In And ���㷽ʽ Is Null;
    If n_Count = 0 Then
      Select Sum(���ʽ��) Into n_������ From ������ü�¼ Where ����id = ���ս���id_In;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, n_����id, Null, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_������, ���ս���id_In, n_�ɿ���id,
         -1 * ����id_In, 1, 3);
    End If;
  
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    For c_No In (Select Distinct NO From ������ü�¼ Where ��¼���� = 11 And ����id = ���ս���id_In) Loop
      Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, c_No.No);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_����;
/

--139063:Ƚ����,2019-04-08,�������۲��˰��������̾���
Create Or Replace Procedure Zl_�����շѼ�¼_Delete
(
  No_In           ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ҽ�����㷽ʽ_In Varchar2 := Null,
  ���_In         Varchar2 := Null,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
  ���_In         ������ü�¼.ʵ�ս��%Type := 0,
  �˷�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type := Null,
  ����Ʊ��_In     Number := 0,
  �˷�ժҪ_In     ������ü�¼.ժҪ%Type := Null,
  У�Ա�־_In     Number := 0,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  �������_In     ����Ԥ����¼.�������%Type := Null,
  һ��ͨ����_In   Varchar2 := Null,
  �˿����_In     Number := 0,
  �൥��ȫ��_In   Number := 0
) As
  --���ܣ�ɾ��һ�������շѵ���
  --������
  --        ҽ�����㷽ʽ_IN   =ҽ���˷�ʱ,��֧�ֽ������ϵĽ��㷽ʽ,���Ϊ�ձ�ʾ��ҽ���˷ѻ�ҽ���˷�ȫ�������������ϡ�
  --        ���_IN           =Ҫ�˷ѵ���Ŀ���,��ʽΪ"1,3,5,6...",ȱʡNULL��ʾ��"δ�˵�"�����С�
  --        ���㷽ʽ_IN       =��Ϊ�����˷�ʱ,�˷ѽ��Ľ��㷽ʽ��
  --        ���_IN           =ָ�˷�ʱ�²����������,�����˷ѻ�ҽ��ȫ�˵�ĳ�ֽ������ֽ�ʱ�Ż�����µ���
  --                           ��ʱ��������ڼ��㱾���˷ѵĽ�����,�����ü�¼�Ĵ����ڱ�����ִ��������Zl_�����շ����_Insert����
  --        ����Ʊ��_In       =0:����ȫ�˻����һ��ȫ��ʱ�ջ�Ʊ��,ע��,���ŵ����˷�ѭ����������ʱֻ�ջ�һ�Ρ�
  --                           1:�����˷Ѳ�����Ʊ��,ͨ���ش���õ�������
  --        У�Ա�־_IN:0-����Ҫ�϶�;1-��϶�(��������Ա�ɿ����,������Ʊ��,������Ԥ�����)
  --        һ��ͨ����_In ȫ��ʱ���벻ԭ���˻صĽ��㷽ʽ��ҽ�ƿ������˷�ʱ������"���㷽ʽ|���"
  --        �˿����_In:1-���в�����(���˿ʽ�˵�ָ���Ľ��㷽ʽ<���㷽ʽ_In>��,0-��ָ���˿ʽ)
  --        �൥��ȫ��_IN=1-�൥��ȫ��(���ŵ���ȫ��,ԭ����);0-��ԭ����
  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼

  --ҽ��ȫ�˵�ĳ�ֽ������ֽ�Ӷ��������µ����ʱ,�ſ��˴�������,ִ���걾���̺�,��������е������������
  Cursor c_Bill Is
    Select a.Id, a.No, a.���ӱ�־, a.�շ�ϸĿid, a.���, a.�۸񸸺�, a.ִ��״̬, a.�շ����, a.����, a.����, a.ҽ�����, j.�������, m.��������,
           Nvl(a.���ӱ�־, 0) As ���
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.No = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.�շ�ϸĿid + 0 = m.����id(+) And
          Nvl(a.���ӱ�־, 0) <> Decode(�൥��ȫ��_In, 1, 999, 9)
    Order By a.�շ�ϸĿid, a.���;
  --:����ԭʼ�������,��Ӧ�ø��ݵ�ǰ�˷Ѳ������������д���
  -- Decode(Sign(���_In), 0, 999, 9)

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Money(����id_In ����Ԥ����¼.����id%Type) Is
    Select ���㷽ʽ, ��Ԥ��
    From ����Ԥ����¼
    Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = ����id_In And ���㷽ʽ Is Not Null And Nvl(��Ԥ��, 0) <> 0 And Nvl(У�Ա�־, 0) = 0;

  --���α����ڲ����շ�ʱʹ�ù��ĳ�Ԥ�����¼
  Cursor c_Deposit(V����id ����Ԥ����¼.����id%Type) Is
    Select ID, ����id, ��Ԥ�� As ���, Ԥ�����
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ��¼״̬ In (1, 3) And ����id = V����id And Nvl(��Ԥ��, 0) <> 0
    Order By ID Desc;

  n_����id   ������Ϣ.����id%Type;
  n_����id   ������ü�¼.����id%Type;
  n_������� ����Ԥ����¼.�������%Type;
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;

  n_���˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ����Ԥ����¼.��Ԥ��%Type;
  n_ԭ���� ������ü�¼.ʵ�ս��%Type;
  --�����˷Ѽ������
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;
  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;
  n_�ܽ��   Number;
  n_����״̬ ������ü�¼.����״̬%Type;
  n_�����˷� Number; --�Ƿ��һ���˷���ȫ���˷�,��ÿ���˷ѹ������жϵõ���
  n_��id     ����ɿ����.Id%Type;

  v_�˷ѽ��� ���㷽ʽ.����%Type;
  v_�������� Varchar2(500);
  n_������   Number(2);
  v_��ǰ���� Varchar2(50);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;

  l_ʹ��id   t_Numlist := t_Numlist();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_ԭ����id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_����ģʽ     Number(3);
  v_Para         Varchar2(1000);
  n_ҽ��ִ�мƼ� Number;
  n_�Ự��       ����Ԥ����¼.�Ự��%Type; --��ʽ��SID+'_'+SERIAL#

Begin
  n_��id := Zl_Get��id(����Ա����_In);

  Begin
    Select Sid || '_' || Serial# Into n_�Ự�� From V$session Where Audsid = Userenv('sessionid');
  Exception
    When Others Then
      n_�Ự�� := Null;
  End;

  n_������ := 0;
  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�Ǹõ������ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  --ִ��״̬��ԭʼ��¼���ж�
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 1 And Nvl(���ӱ�־, 0) <> 9 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�';
    Raise Err_Item;
  End If;
  --ȷ���Ƿ���ҽ��ִ�мƼ��д�������,�����������,�����ҽ��ִ�мƼ۽����˷�,���򰴾ɷ�ʽ���д���
  Select Count(1)
  Into n_ҽ��ִ�мƼ�
  From ������ü�¼ A, ҽ��ִ�мƼ� B
  Where a.ҽ����� = b.ҽ��id And a.��¼���� = 1 And a.No = No_In And a.��¼״̬ In (1, 3) And Rownum = 1;

  ---------------------------------------------------------------------------------
  --���ñ���
  If �˷�ʱ��_In Is Not Null Then
    d_Date := �˷�ʱ��_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;
  If ����id_In Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Else
    n_����id := ����id_In;
  End If;
  n_������� := �������_In;
  If n_������� Is Null Then
    n_������� := ����id_In;
  End If;
  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --��ȡ���㷽ʽ����
  v_�˷ѽ��� := ���㷽ʽ_In;
  If v_�˷ѽ��� Is Null Then
    Begin
      Select ���� Into v_�˷ѽ��� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�˷ѽ��� := '�ֽ�';
    End;
  End If;
  --ѭ������ÿ�з���(������Ŀ��)
  n_�ܽ��   := 0;
  n_�����˷� := 1;
  For r_Bill In c_Bill Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
        From ������ü�¼
        Where NO = No_In And ��¼���� = 1 And ��� = r_Bill.���;
      
        If n_ʣ������ = 0 Then
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ���˷ѣ�';
            Raise Err_Item;
          End If;
          --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ���˷�(ִ��״̬=0��һ�ֿ���)
          n_�����˷� := 0;
        Else
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            --@@@
            --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
            --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ)
            --: 2.������ҽ����,����ʣ������Ϊ׼
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
              If n_ҽ��ִ�мƼ� = 1 Then
                Select Decode(Sign(Sum(����)), -1, 0, Sum(����)), Count(*)
                Into n_׼������, n_Count
                From (Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, Max(a.ҽ�����) As ҽ��id, Max(a.�շ�ϸĿid) As �շ�ϸĿid,
                              Sum(Nvl(a.����, 1) * Nvl(a.����, 1)) As ����,
                              Sum(Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1))) As ԭʼ����
                       From ������ü�¼ A, ����ҽ����¼ M
                       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Instr('5,6,7', a.�շ����) = 0 And a.No = No_In And a.��� = r_Bill.��� And a.��¼���� = 1 And
                             a.��¼״̬ In (1, 2, 3) And a.�۸񸸺� Is Null
                       Group By a.���
                       Union All
                       Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����
                       From ������ü�¼ A, ҽ��ִ�мƼ� B, ����ҽ����¼ M
                       Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And
                             Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And
                             (Exists
                              (Select 1
                               From ����ҽ��ִ��
                               Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1) Or Exists
                              (Select 1
                               From ����ҽ������
                               Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1)) And Not Exists
                        (Select 1
                              From ����ҽ������
                              Where a.ҽ����� = ҽ��id And a.No = NO And Mod(a.��¼����, 10) = ��¼����) And a.No = No_In And
                             a.��� = r_Bill.��� And a.��¼���� = 1 And a.��¼״̬ In (1, 3) ��and a.�۸񸸺� Is Null) Q1
                Where Not Exists (Select 1
                       From ҩƷ�շ���¼
                       Where ����id = Q1.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) Having
                 Max(ID) <> 0;
              Else
              
                Select Nvl(Sum(����), 0), Count(*)
                Into n_׼������, n_Count
                From (Select a.ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(b.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And a.ҽ��id = m.Id And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And
                             a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And j.��¼���� = 1 And j.��� = r_Bill.��� And
                             j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Exists
                        (Select 1
                              From ����ҽ���Ƽ� A
                              Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0)
                       Union All
                       Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And a.ҽ��id = m.Id And
                             Nvl(c.ִ�н��, 1) = 1 And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And
                             j.No = No_In And j.��¼���� = 1 And Nvl(a.�շѷ�ʽ, 0) = 0 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                             j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) And Not Exists
                        (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                       Union All
                       Select a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * a.���� As ����
                       From ������ü�¼ A, ����ҽ����¼ M
                       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And a.No = No_In And
                             a.��¼���� = 1 And a.��� = r_Bill.��� And a.��¼״̬ = 2 And a.�۸񸸺� Is Null And Not Exists
                        (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = a.�շ�ϸĿid));
              End If;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_׼������ = 0 Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з�����ִ��,�������˷ѣ�';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_׼������ := n_ʣ������;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
            Into n_׼������, n_Count
            From ҩƷ�շ���¼
            Where NO = No_In And ���� In (8, 24) And Mod(��¼״̬, 3) = 1 --@@@
                  And ����� Is Null And ����id = r_Bill.Id;
          
            --��ʣ��������׼�������������������
            --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������
            --2.��������,��ʱ�ѷ�ҩ����
            If n_׼������ = 0 Then
              If r_Bill.�շ���� = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                  Raise Err_Item;
                Else
                  n_׼������ := n_ʣ������;
                End If;
              Else
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --�Ƿ񲿷��˷�
          If r_Bill.ִ��״̬ = 2 Or n_׼������ <> Nvl(r_Bill.����, 1) * r_Bill.���� Then
            n_�����˷� := 0;
          End If;
        
          --����������ü�¼
          n_����״̬ := 0;
          --�ñ���Ŀ�ڼ����˷�
          If Nvl(У�Ա�־_In, 0) <> 0 Then
            n_�˷Ѵ��� := -9; --�ȱ���,�̶�Ϊ9
            n_����״̬ := 1;
          Else
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into n_�˷Ѵ���
            From ������ü�¼
            Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 2 And Nvl(ִ��״̬, 0) < 0 And ��� = r_Bill.���;
          End If;
        
          --���=ʣ����*(׼����/ʣ����)
          If Nvl(r_Bill.���, 0) = 9 Then
            --�����Գ������õ�С��λ(����:ҽ�����㳬��С��λ��,���Ϳ��ܳ���С��λ
            n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), 5);
            n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), 5);
            n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), 5);
          Else
            n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
          End If;
          n_�ܽ�� := n_�ܽ�� + n_ʵ�ս��;
        
          --�����˷Ѽ�¼
          Insert Into ������ü�¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����,
             ִ��״̬, ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����,
             �ɿ���id, ��ҳid, ���˲���id)
            Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                   ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                   Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����,
                   -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, n_����״̬, ִ��ʱ��, ����Ա���_In,
                   ����Ա����_In, ����ʱ��, d_Date, n_����id, -1 * n_ʵ�ս��, ������Ŀ��, ���մ���id, -1 * n_ͳ����, Nvl(�˷�ժҪ_In, ժҪ),
                   Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, ����, n_��id, ��ҳid, ���˲���id
            From ������ü�¼
            Where ID = r_Bill.Id;
        
          --���ԭ���ü�¼
          --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1,�쳣�շѵ�,���Ǳ���9
          Update ������ü�¼
          Set ��¼״̬ = 3, ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 9, 9, Decode(Sign(n_׼������ - n_ʣ������), 0, 0, 1))
          Where ID = r_Bill.Id;
        End If;
      Else
        If ���_In Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�����˷ѣ�';
          Raise Err_Item;
        End If;
        --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
        n_�����˷� := 0;
      End If;
    Else
      n_�����˷� := 0; --δָ���ñ�,���ڲ����˷�
    End If;
  End Loop;
  ---------------------------------------------------------------------------------
  --������Ԥ����¼

  --ԭ���ݵĽ���ID
  Select ����id, ����id
  Into n_ԭ����id, n_����id
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum = 1;

  If n_�����˷� = 1 And Nvl(�˿����_In, 0) = 0 Then
    --���ݵ�һ���˷���ȫ������
    --��Ԥ�����ּ�¼
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
      Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
             ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
             Decode(У�Ա�־_In, 1, 2, У�Ա�־_In), n_�������, 3, n_�Ự��
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
    If Nvl(У�Ա�־_In, 0) = 0 Then
      --������Ԥ�����
      For v_Ԥ�� In (Select Ԥ�����, Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����, ����id
                   From ����Ԥ����¼
                   Where ��¼���� In (1, 11) And ����id = n_ԭ����id
                   Group By Ԥ�����, ����id
                   Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
        Where ����id = v_Ԥ��.����id And ���� = 1 And ���� = Nvl(v_Ԥ��.Ԥ�����, 2)
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, Ԥ�����, ����)
          Values
            (v_Ԥ��.����id, Nvl(v_Ԥ��.Ԥ�����, 2), Nvl(v_Ԥ��.Ԥ�����, 0), 1);
          n_����ֵ := n_Ԥ�����;
        End If;
        If n_����ֵ = 0 Then
          Delete From �������
          Where ����id = v_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      End Loop;
    End If;
    --��ҽ��ȫ��,��ҽ�����н��㷽ʽ���������,ԭ���˻�(��Ԥ����ǰ���Ѵ���)
    If ҽ�����㷽ʽ_In Is Null Then
      v_�������� := ',' || Nvl(һ��ͨ����_In, '-Lxh') || ',' || Nvl(һ��ͨ����_In, 'Lxh') || ',';
    
      --һ��ͨ�����ѿ������п������������Ҫ���⴦��,��Ҫ���϶�.
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
        Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
               -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
               Case
                 When Nvl(�����id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(���㿨���, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(q.Ԥ��id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(j.����, '-') <> '-' Then
                  Decode(У�Ա�־_In, 1, 1, 0)
                 Else
                  Decode(У�Ա�־_In, 1, 2, 0)
               End As У�Ա�־, n_�������, 3, n_�Ự��
        From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J,
             (Select m.Id As Ԥ��id
               From ����Ԥ����¼ M, һ��ͨĿ¼ C
               Where m.����id = n_ԭ����id And m.���㷽ʽ = c.���㷽ʽ And m.��¼���� = 3 And m.��¼״̬ = 1) Q
        Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id = n_ԭ����id And a.Id = q.Ԥ��id(+) And a.���㷽ʽ = j.����(+) And
              Instr(v_��������, ',' || ���㷽ʽ || ',') = 0 And
              (Not Exists (Select 1 From ���˿������¼ Where ����id = a.Id) Or Nvl(a.���㿨���, 0) = 0);
    
      --�������ѿ�,���㿨��������Ѿ�������
      Select Count(1)
      Into n_Count
      From ����Ԥ����¼ A, ���˿������¼ B
      Where a.Id = b.����id And a.��¼���� = 3 And a.����id = n_ԭ����id And Rownum < 2;
      If n_Count <> 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id,
           Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
          Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
                 -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2, n_�������, Mod(��¼����, 10), n_�Ự��
          From ����Ԥ����¼ A
          Where a.��¼���� = 3 And a.����id = n_ԭ����id And Exists
           (Select 1 From ���˿������¼ Where ����id = a.Id) And Instr(Nvl(v_��������, '_LXH'), ',' || a.���㷽ʽ || ',') = 0;
      
        --�շ�ʱ����ʹ���˶������ѿ�
        For c_��¼ In (Select a.Id, c.�ӿڱ��, c.���ѿ�id, c.����, -1 * Sum(c.Ӧ�ս��) As ������
                     From ����Ԥ����¼ A, ���˿������¼ C
                     Where a.Id = c.����id And a.��¼���� = 3 And a.��¼״̬ In (1, 3) And a.����id = n_ԭ����id And
                           Instr(Nvl(v_��������, '_LXH'), ',' || a.���㷽ʽ || ',') = 0
                     Group By a.Id, c.�ӿڱ��, c.���ѿ�id, c.����) Loop
        
          Zl_���˿������¼_�˿�(c_��¼.�ӿڱ��, c_��¼.����, c_��¼.���ѿ�id, c_��¼.������, c_��¼.Id, n_Ԥ��id, ����Ա���_In, ����Ա����_In, d_Date);
        End Loop;
      End If;
    
      --b.���µľ��������ӿ�֧�ֵ�������,���������ϵĽ��㷽ʽ,���ϵ�ָ���Ľ��㷽ʽ��,�������(��Ϊ������������֮�������)
      If һ��ͨ����_In Is Not Null Then
        Begin
          Select -1 * Nvl(Sum(��Ԥ��), 0) Into n_���˽�� From ����Ԥ����¼ Where ����id = n_����id;
        Exception
          When Others Then
            n_���˽�� := 0;
        End;
      
        If (n_�ܽ�� - n_���˽��) <> 0 Then
          --��ʱ���ܽ�û�а������,��Ϊ����������ڵ��ñ����̺�Ų��������ü�¼
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
             ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
            Select ����Ԥ����¼_Id.Nextval, 3, NO, 2, ����id, ��ҳid, '�����˷ѽ���', v_�˷ѽ���, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * (n_�ܽ�� - n_���˽��), n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
                   Decode(У�Ա�־_In, 1, 2, 0), n_�������, 3, n_�Ự��
            From ����Ԥ����¼
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum = 1;
          n_������ := 1;
        End If;
      End If;
      --ҽ�����������ϵĽ��㷽ʽ��,�������,�˵�ָ���Ľ��㷽ʽ��
      --��Ҫ���������
    Else
      --a.ԭ���˻�
      v_�������� := ',' || ҽ�����㷽ʽ_In || ',' || Nvl(һ��ͨ����_In, '-Lxh') || ',' || v_�˷ѽ��� || ',';
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
        Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, ����Ա���_In, ����Ա����_In, -1 * ��Ԥ��, n_����id,
               n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
               
               Case
                 When Nvl(�����id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(���㿨���, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(q.Ԥ��id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(j.����, '-') <> '-' Then
                  Decode(У�Ա�־_In, 1, 1, 0)
                 Else
                  Decode(У�Ա�־_In, 1, 2, 0)
               End As У�Ա�־, n_�������, 3, n_�Ự��
        From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J,
             (Select m.Id As Ԥ��id
               From ����Ԥ����¼ M, һ��ͨĿ¼ C
               Where m.����id = n_ԭ����id And m.���㷽ʽ = c.���㷽ʽ And m.��¼���� = 3 And m.��¼״̬ = 1) Q
        Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.���㷽ʽ = j.����(+) And a.����id = n_ԭ����id And
              Instr(v_��������, ',' || a.���㷽ʽ || ',') = 0 And a.Id = q.Ԥ��id(+) And
              (Not Exists (Select 1 From ���˿������¼ Where ����id = a.Id) Or Nvl(a.���㿨���, 0) = 0);
    
      --�������ѿ�,���㿨��������Ѿ�������
      Select Count(1)
      Into n_Count
      From ����Ԥ����¼ A, ���˿������¼ B
      Where a.Id = b.����id And a.��¼���� = 3 And a.����id = n_ԭ����id And Rownum < 2;
      If n_Count <> 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id,
           Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
          Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
                 -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2, n_�������, Mod(��¼����, 10), n_�Ự��
          From ����Ԥ����¼ A
          Where a.��¼���� = 3 And a.����id = n_ԭ����id And Exists
           (Select 1 From ���˿������¼ Where ����id = a.Id) And Instr(Nvl(v_��������, '_LXH'), ',' || a.���㷽ʽ || ',') = 0;
      
        --�շ�ʱ����ʹ���˶������ѿ�
        For c_��¼ In (Select a.Id, c.�ӿڱ��, c.���ѿ�id, c.����, -1 * c.Ӧ�ս�� As ������
                     From ����Ԥ����¼ A, ���˿������¼ C
                     Where a.Id = c.����id And a.��¼���� = 3 And a.��¼״̬ In (1, 3) And a.����id = n_ԭ����id And
                           Instr(Nvl(v_��������, '_LXH'), ',' || a.���㷽ʽ || ',') = 0) Loop
        
          Zl_���˿������¼_�˿�(c_��¼.�ӿڱ��, c_��¼.����, c_��¼.���ѿ�id, c_��¼.������, c_��¼.Id, n_Ԥ��id, ����Ա���_In, ����Ա����_In, d_Date);
        End Loop;
      End If;
    
      --b.���µľ���ҽ�����������ϵĽ��㷽ʽ,���ϵ�ָ���Ľ��㷽ʽ��,�������(��Ϊ������������֮�������)
      Begin
        Select -1 * Nvl(Sum(��Ԥ��), 0) Into n_���˽�� From ����Ԥ����¼ Where ����id = n_����id;
      Exception
        When Others Then
          n_���˽�� := 0;
      End;
    
      If (n_�ܽ�� - n_���˽��) <> 0 Then
        --��ʱ���ܽ�û�а������,��Ϊ����������ڵ��ñ����̺�Ų��������ü�¼
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
          Select ����Ԥ����¼_Id.Nextval, 3, NO, 2, ����id, ��ҳid, Decode(һ��ͨ����_In, Null, '����ҽ���ӿ��˷�', '����ҽ���ӿں������ӿ��˷�'), v_�˷ѽ���,
                 d_Date, ����Ա���_In, ����Ա����_In, -1 * (n_�ܽ�� - n_���˽��), n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
                 ������λ, Decode(У�Ա�־_In, 1, 2, 0), n_�������, 3, n_�Ự��
          From ����Ԥ����¼
          Where ��¼���� = 3 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum = 1;
        n_������ := 1;
      End If;
    
    End If;
  Else
    -------------------------------------------------
    --�����˷�
    n_���˽�� := 0;
    --ҽ�ƿ������˷�ʱ������:���㷽ʽ|���
    If һ��ͨ����_In Is Not Null Then
      If Instr(һ��ͨ����_In, '|') > 0 Then
        v_��ǰ���� := һ��ͨ����_In;
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
        n_������ := Nvl(To_Number(v_��ǰ����), 0);
        If Not Nvl(v_���㷽ʽ, 'TMP') = 'TMP' Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
             �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
            Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '�����ӿڲ����˷�', v_���㷽ʽ, d_Date, Null, Null, Null, ����Ա���_In,
                   ����Ա����_In, -1 * (n_������), n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
                   Decode(У�Ա�־_In, 1, 1, 0), n_�������, 3, n_�Ự��
            From ����Ԥ����¼
            Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum < 2;
        End If;
        n_���˽�� := n_������;
      End If;
    End If;
    --����ֱ����Ϊָ�����㷽ʽ
    If (n_�ܽ�� - n_���˽�� + Nvl(���_In, 0)) <> 0 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
         ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
        Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '�����˷ѽ���', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In,
               ����Ա����_In, -1 * (n_�ܽ�� - n_���˽�� + Nvl(���_In, 0)), n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
               Decode(У�Ա�־, 1, 2, 0), n_�������, 3, n_�Ự��
        From ����Ԥ����¼
        Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
    End If;
  
    --����շ�ʱֻʹ����Ԥ����,��Ҫ��Ԥ��,���ҿ����ж�ʳ�Ԥ��
    If Sql%RowCount = 0 And һ��ͨ����_In Is Null Then
      n_Ԥ����� := n_�ܽ�� - n_���˽�� + Nvl(���_In, 0);
    
      For r_Deposit In c_Deposit(n_ԭ����id) Loop
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 d_Date, ����Ա����_In, ����Ա���_In, Decode(Sign(r_Deposit.��� - n_Ԥ�����), -1, -1 * r_Deposit.���, -1 * n_Ԥ�����),
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Decode(У�Ա�־_In, 1, 2, 0), n_�������, 3, n_�Ự��
          From ����Ԥ����¼
          Where ID = r_Deposit.Id;
      
        If Nvl(У�Ա�־_In, 0) = 0 Then
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + n_�ܽ�� + Nvl(���_In, 0)
          Where ����id = r_Deposit.����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, Ԥ�����, ����)
            Values
              (r_Deposit.����id, 1, n_�ܽ�� + Nvl(���_In, 0), 1);
            n_����ֵ := n_�ܽ�� + Nvl(���_In, 0);
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
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
    End If;
  End If;

  --����ԭ��¼
  Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id;

  If �൥��ȫ��_In <> 1 Then
    --���������,�൥��ȫ��ʱ ,��ԭ����������
    --�������ļ�¼״̬����Ϊ3
    If Nvl(���_In, 0) <> 0 Then
      n_Count := 1;
      If n_�����˷� = 1 And Nvl(�˿����_In, 0) = 0 Then
        n_ԭ���� := 0;
        --ԭ����,���������
        If n_������ = 0 Then
          Select -1 * Nvl(Sum(ʵ�ս��), 0)
          Into n_ԭ����
          From ������ü�¼ A
          Where NO = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And Nvl(a.���ӱ�־, 0) = 9;
        End If;
        If Nvl(n_ԭ����, 0) <> 0 Or Nvl(���_In, 0) <> 0 Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� - n_ԭ���� - Nvl(���_In, 0)
          Where ���㷽ʽ = v_�˷ѽ��� And ����id = n_����id;
          If Sql%NotFound Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
              Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '����', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In,
                     ����Ա����_In, -1 * n_ԭ���� - Nvl(���_In, 0), n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, Null, Null, 0,
                     n_�������, 3, n_�Ự��
              From ����Ԥ����¼
              Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
          End If;
        End If;
      End If;
    Elsif n_�����˷� = 1 And Nvl(�˿����_In, 0) = 0 Then
      --ԭ����ʱ,��Ҫ����Ԥ����¼��������
      Select Nvl(Sum(Nvl(���ʽ��, 0)), 0) Into n_ʵ�ս�� From ������ü�¼ Where ����id = n_����id;
      Select Nvl(Sum(Nvl(��Ԥ��, 0)), 0) Into n_����ֵ From ����Ԥ����¼ Where ����id = n_����id;
      If Abs(n_ʵ�ս��) <> Abs(n_����ֵ) Then
        n_ʵ�ս�� := n_ʵ�ս�� - n_����ֵ;
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + Nvl(n_ʵ�ս��, 0) Where ���㷽ʽ = v_�˷ѽ��� And ����id = n_����id;
        If Sql%NotFound Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
             �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������, �Ự��)
            Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '����', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In,
                   ����Ա����_In, Nvl(n_ʵ�ս��, 0), n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, Null, Null, 0, n_�������, 3,
                   n_�Ự��
            From ����Ԥ����¼
            Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
        End If;
      End If;
    End If;
  
    Select Nvl(Sum(Nvl(���ʽ��, 0)), 0) Into n_ʵ�ս�� From ������ü�¼ Where ����id = n_����id;
    Select Nvl(Sum(Nvl(��Ԥ��, 0)), 0) Into n_����ֵ From ����Ԥ����¼ Where ����id = n_����id;
  
    n_ʵ�ս�� := n_ʵ�ս�� - n_����ֵ;
  
    If n_ʵ�ս�� <> 0 Then
      --δ�ҵ����²��������
      Zl_�����շ����_Insert(No_In, n_ʵ�ս��, 1, 0);
    End If;
  End If;
  ---------------------------------------------------------------------------------
  --��Ա�ɿ����(ע����Ԥ����¼�����Ŵ������������ʻ��ȵĽ�����,�����˳�Ԥ����)
  --�������ҪУ�Ե�,�ݲ�������Ա�ɿ����
  If Nvl(У�Ա�־_In, 0) = 0 Then
    For r_Moneyrow In c_Money(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + r_Moneyrow.��Ԥ��
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Moneyrow.���㷽ʽ, 1, r_Moneyrow.��Ԥ��);
        n_����ֵ := r_Moneyrow.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;

  ---------------------------------------------------------------------------------
  --�˷�Ʊ�ݻ���(��ȫ��ʱ�Ż���,�����������ش�����л���)
  If ����Ʊ��_In = 0 Then
  
    --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
    v_Para     := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
    n_����ģʽ := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_����ģʽ <> 0 Then
      --�ջ�Ʊ��
      Select ʹ��id
      Bulk Collect
      Into l_ʹ��id
      From (Select Distinct b.ʹ��id From Ʊ�ݴ�ӡ��ϸ B Where b.No = No_In And Nvl(b.Ʊ��, 0) = 1);
    
      n_����ģʽ := l_ʹ��id.Count;
      If l_ʹ��id.Count <> 0 Then
        --������ռ�¼
        Forall I In 1 .. l_ʹ��id.Count
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, ����Ա����_In, d_Date, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ A
            Where ID = l_ʹ��id(I) And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = a.Ʊ�� And Nvl(����, 0) <> 1);
      
        Forall I In 1 .. l_ʹ��id.Count
          Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I) And Nvl(�Ƿ����, 0) = 0;
      
      End If;
    End If;
    If n_����ģʽ = 0 Then
      --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ)
      Begin
        --����=1��ԭ��=6Ϊ�˷Ѵ�ӡƱ��(��Ʊ)��������
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = No_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --������ǰû�д�ӡ,���ջ�
      If n_��ӡid Is Not Null Then
        --a.���ŵ���ѭ������ʱֻ���ջ�һ��
        Select Count(*) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        Else
          --b.�����˷Ѷ���ջ�ʱ,���һ��ȫ���ջ�Ҫ�ſ����ջص�
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In, Ʊ�ݽ��
            From Ʊ��ʹ����ϸ A
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 2);
        End If;
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --ҩƷ�����������
  --���밴�ա��շ�ϸĿid���������򣬷�ֹ��������ҩƷ��桱��
  For r_Expenses In (Select ID
                     From ������ü�¼
                     Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And �շ���� In ('4', '5', '6', '7') And
                           (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null)
                     Order By �շ�ϸĿid) Loop
    Zl_ҩƷ�շ���¼_�����˷�(r_Expenses.Id);
  End Loop;

  --ҽ������
  --ɾ������ҽ������(���һ��ɾ��ʱ)
  For c_ҽ�� In (Select Distinct ҽ�����
               From ������ü�¼
               Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, ִ��״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
                  Where ��¼���� = 1 And Nvl(���ӱ�־, 0) <> 9 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, ִ��״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 1 And NO = No_In;
    End If;
  End Loop;

  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In Varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_Delete;
/

--139063:Ƚ����,2019-04-03,�������۲��˰��������̾���
Create Or Replace Procedure Zl_���˷�������_Delete
(
  Ids_In    In Varchar2,
  ��ҩid_In In ��Һ��ҩ��¼.Id%Type := Null
) As
  n_Id  ���˷�������.����id%Type;
  v_Ids Varchar2(4000);

  n_ҽ��id   סԺ���ü�¼.Id%Type;
  v_No       סԺ���ü�¼.No%Type;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_�������� ��Һ��ҩ��¼.����״̬%Type;
  n_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
Begin
  If ��ҩid_In Is Not Null Then
    Select ����ʱ��
    Into n_����ʱ��
    From (Select ������Ա, ����ʱ��, ��������
           From ��Һ��ҩ״̬
           Where ��ҩid = ��ҩid_In And �������� = 9
           Order By ����ʱ�� Desc)
    Where Rownum = 1;
  End If;

  v_Ids := Ids_In || ',';
  While v_Ids Is Not Null Loop
    n_Id  := To_Number(Substr(v_Ids, 1, Instr(v_Ids, ',') - 1));
    v_Ids := Substr(v_Ids, Instr(v_Ids, ',') + 1);
  
    If n_����ʱ�� Is Null Then
      Delete ���˷������� Where ����id = n_Id And ״̬ = 0;
    
      Select NO, ҽ�����
      Into v_No, n_ҽ��id
      From (Select a.No, a.ҽ�����
             From סԺ���ü�¼ A
             Where a.Id = n_Id
             Union All
             Select a.No, a.ҽ�����
             From ������ü�¼ A
             Where a.Id = n_Id);
      If Not n_ҽ��id Is Null Then
        --��δ�ṩ����ҩ����ȡ���Ĺ��ܣ����������������һ��ȡ��
        For R In (Select d.Id
                  From ����ҽ����¼ A, ����ҽ������ B, ��Һ��ҩ��¼ D
                  Where a.Id = n_ҽ��id And a.Id = b.ҽ��id And b.No = v_No And a.���id = d.ҽ��id And b.���ͺ� = d.���ͺ� And
                        b.��¼���� = 2) Loop
          Select ������Ա, ����ʱ��, ��������
          Into v_������Ա, d_����ʱ��, n_��������
          From (Select ������Ա, ����ʱ��, ��������
                 From ��Һ��ҩ״̬
                 Where ��ҩid = r.Id And �������� <> 9
                 Order By ����ʱ�� Desc, �������� Desc)
          Where Rownum = 1;
          Update ��Һ��ҩ��¼ Set ������Ա = v_������Ա, ����ʱ�� = d_����ʱ��, ����״̬ = n_�������� Where ID = r.Id;
        End Loop;
      End If;
    Else
      Delete ���˷������� Where ����id = n_Id And ״̬ = 0 And ����ʱ�� = n_����ʱ��;
      Select ������Ա, ����ʱ��, ��������
      Into v_������Ա, d_����ʱ��, n_��������
      From (Select ������Ա, ����ʱ��, ��������
             From ��Һ��ҩ״̬
             Where ��ҩid = ��ҩid_In And �������� <> 9
             Order By ����ʱ�� Desc, �������� Desc)
      Where Rownum = 1;
      Update ��Һ��ҩ��¼ Set ������Ա = v_������Ա, ����ʱ�� = d_����ʱ��, ����״̬ = n_�������� Where ID = ��ҩid_In;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˷�������_Delete;
/

--120692:����,2019-04-03,�����¼֧�ּ�����Ŀ����
Create Or Replace Procedure Zl_�������ݵ��붨��_Update
(
  ���_In �������ݵ��붨��.���%Type,
  ����_In �������ݵ��붨��.����%Type,
  ��ʽ_In �������ݵ��붨��.��ʽ%Type
) Is
Begin
  Update �������ݵ��붨�� Set ���� = ����_In, ��ʽ = ��ʽ_In Where ��� = ���_In;
  If Sql%Rowcount = 0 Then
    Insert Into �������ݵ��붨�� (���, ����, ��ʽ) Values (���_In, ����_In, ��ʽ_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ݵ��붨��_Update;
/

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Create Or Replace Procedure Zl_���˷�������_Audit
(
  Id_In       ���˷�������.����id%Type,
  ����ʱ��_In ���˷�������.����ʱ��%Type,
  �����_In   ���˷�������.�����%Type,
  ���ʱ��_In ���˷�������.���ʱ��%Type,
  ״̬_In     ���˷�������.״̬%Type,
  Int�Զ����� Integer := 1,
  �������_In ���˷�������.�������%Type := 1
) As
  --���ܣ���˻�ȡ�������������
  --��Σ�
  --    ״̬_In 1-���ͨ��,2-���δͨ��
  --    �������_In ��ҩƷ��������Ч,ȱʡΪ��ִ�е�ҩƷ������ 
  --˵����
  --    ���ÿ���������סԺ���ü�¼��Ҳ����������������ü�¼
  n_ִ��״̬       סԺ���ü�¼.ִ��״̬%Type;
  n_�������       ���˷�������.�������%Type;
  v_�շ����       סԺ���ü�¼.�շ����%Type;
  v_No             סԺ���ü�¼.No%Type;
  n_ʵ������       ҩƷ�շ���¼.ʵ������%Type;
  n_����           ���˷�������.����%Type;
  n_�շ�id         ҩƷ�շ���¼.Id%Type;
  n_ҽ��id         סԺ���ü�¼.ҽ�����%Type;
  v_��������       ��������.��������%Type;
  n_�շ�ϸĿid     סԺ���ü�¼.�շ�ϸĿid%Type;
  n_��˲���id     ���˷�������.��˲���id%Type;
  n_ִ�в���id     סԺ���ü�¼.ִ�в���id%Type;
  n_����id         סԺ���ü�¼.����id%Type;
  n_��ҳid         סԺ���ü�¼.��ҳid%Type;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);

  n_Count   Number(18);
  n_Temp    Number(18);
  v_Err_Msg Varchar2(300);
  Err_Item Exception;
Begin
  n_������� := 0;
  Select a.ִ��״̬, a.�շ����, a.�շ�ϸĿid, a.ִ�в���id, a.No, Nvl(b.��������, 0), a.ҽ�����, ����id, ��ҳid
  Into n_ִ��״̬, v_�շ����, n_�շ�ϸĿid, n_ִ�в���id, v_No, v_��������, n_ҽ��id, n_����id, n_��ҳid
  From (Select �շ����, NO, �շ�ϸĿid, ���˲���id, ִ�в���id, ҽ�����, ����id, ��ҳid, ִ��״̬
         From סԺ���ü�¼
         Where ID = Id_In
         Union All
         Select �շ����, NO, �շ�ϸĿid, ���˲���id, ִ�в���id, ҽ�����, ����id, ��ҳid, ִ��״̬
         From ������ü�¼
         Where ID = Id_In) A, �������� B
  Where a.�շ�ϸĿid = b.����id(+);

  If Nvl(n_��ҳid, 0) <> 0 Then
    n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
    n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
    If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
      Begin
        Select ��˱�־, ״̬
        Into n_��˱�־, n_סԺ״̬
        From ������ҳ
        Where ����id = Nvl(n_����id, 0) And ��ҳid = Nvl(n_��ҳid, 0);
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
  End If;
  If Instr('567', v_�շ����) > 0 Or (v_�շ���� = '4' And Nvl(v_��������, 0) = 1) Then
    n_������� := �������_In;
  End If;

  Update ���˷�������
  Set ����� = �����_In, ���ʱ�� = ���ʱ��_In, ״̬ = ״̬_In
  Where ����id = Id_In And ������� = n_������� And ����ʱ�� = ����ʱ��_In And ״̬ = 0
  Returning ����, ��˲���id Into n_����, n_��˲���id;
  If Sql%RowCount = 0 Then
    v_Err_Msg := '�������ʧ��,��ǰ�����ļ�¼������Ϊ���������Ѿ������˴���,����ˢ����Ϣ!';
    Raise Err_Item;
  End If;

  If n_������� = 0 And (Instr(',5,6,7', ',' || v_�շ����) > 0 Or (v_�շ���� = '4' And Nvl(v_��������, 0) = 1)) Then
    --��Ҫ���δִ�е���������ȫ������,�Ż�ͨ�� 
    Select Sum(Nvl(����, 0) * Nvl(ʵ������, 0))
    Into n_ʵ������
    From ҩƷ�շ���¼
    Where ������� Is Null And ����id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0;
    If Nvl(n_ʵ������, 0) < Nvl(n_����, 0) Then
      Select '�ڵ��ݺ�<<' || v_No || '>>��' || Decode(v_�շ����, '4', '����', 'ҩƷ') || 'Ϊ:' || Chr(13) || ���� || '-' || ���� ||
              Chr(13) || '����������(' || LTrim(To_Char(n_����, '9999999990.99')) || ')�����˴���' || Decode(v_�շ����, '4', '��', 'ҩ') ||
              '����(' || LTrim(To_Char(Nvl(n_ʵ������, 0), '9999999990.99')) || '),���������!'
      Into v_Err_Msg
      From �շ���ĿĿ¼
      Where ID = n_�շ�ϸĿid;
      Raise Err_Item;
    End If;
  
    If n_ҽ��id <> 0 Then
      Select Nvl(Max(d.Id), 0)
      Into n_Count
      From ����ҽ����¼ A, ����ҽ������ B, ��Һ��ҩ��¼ D
      Where a.Id = n_ҽ��id And a.Id = b.ҽ��id And b.No = v_No And a.���id = d.ҽ��id And b.���ͺ� = d.���ͺ� And b.��¼���� = 2 And
            d.����ʱ�� = ����ʱ��_In And d.����״̬ = 9;
    
      If n_Count <> 0 Then
        Select Count(1)
        Into n_Temp
        From ��Һ��ҩ״̬
        Where ��ҩid = n_Count And �������� = 10 And ����ʱ�� = ���ʱ��_In;
        If n_Temp = 0 Then
          Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (n_Count, 10, �����_In, ���ʱ��_In);
        End If;
        Update ��Һ��ҩ��¼ Set ������Ա = �����_In, ����ʱ�� = ���ʱ��_In, ����״̬ = 10 Where ID = n_Count;
      End If;
    End If;
  End If;

  If n_ִ��״̬ <> 0 Then
    If Instr(',5,6,7,', ',' || v_�շ���� || ',') > 0 And n_������� = 1 Then
      If n_ִ�в���id <> n_��˲���id Then
        Begin
          Select '[' || ���� || ']' || ���� Into v_Err_Msg From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;
        Exception
          When Others Then
            v_Err_Msg := '';
        End;
        v_Err_Msg := '���������ʱ,ҩƷΪ' || v_Err_Msg || ' ���Ѿ���ִ�п���ִ��,�����ٽ����������,��ȡ�����!';
        Raise Err_Item;
      End If;
    End If;
  
    If v_�շ���� = '4' Then
      If v_�������� = 1 Then
        If n_ִ�в���id <> n_��˲���id And n_������� = 1 And Int�Զ����� <> 1 Then
          Begin
            Select '[' || ���� || ']' || ���� Into v_Err_Msg From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;
          Exception
            When Others Then
              v_Err_Msg := '';
          End;
          v_Err_Msg := '���������ʱ,����Ϊ' || v_Err_Msg || ' ���Ѿ���ִ�п���ִ��,�����ٽ����������,��ȡ�����!';
          Raise Err_Item;
        End If;
      
        If n_������� = 1 And Int�Զ����� = 1 Then
          n_�շ�id := -1;
          --���������ڶ������ 
          For c_�շ���¼ In (Select ID, ����, Nvl(Sum(Nvl(����, 1) * ʵ������), 0) As ����
                         From ҩƷ�շ���¼
                         Where ����id = Id_In And ���� In (25, 26) And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)
                         Group By ID, ����) Loop
            n_�շ�id := c_�շ���¼.Id;
            If n_���� = 0 Then
              Exit;
            End If;
          
            If n_���� > c_�շ���¼.���� Then
              n_Temp := c_�շ���¼.����;
              n_���� := n_���� - c_�շ���¼.����;
            Else
              n_Temp := n_����;
              n_���� := 0;
            End If;
            Zl_�����շ���¼_��������(c_�շ���¼.Id, �����_In, ���ʱ��_In, c_�շ���¼.����, Null, Null, n_Temp, 0);
          End Loop;
          If n_�շ�id = -1 Then
            v_Err_Msg := '���������ʱ,����Ϊ' || v_Err_Msg || ' ��δ�ҵ���ص�ҩƷ�շ���Ϣ,��������Ϊ��;' || Chr(13) ||
                         '���������ĵĸ�������,�����ٽ����������,��ȡ�����!';
            Raise Err_Item;
          End If;
        End If;
      Else
        --���Ǹ��ٵ����� 
        Update סԺ���ü�¼ Set ִ��״̬ = 0 Where ID = Id_In;
        If Sql%NotFound Then
          Update ������ü�¼ Set ִ��״̬ = 0 Where ID = Id_In;
        End If;
      End If;
    Elsif Instr(',5,6,7,', ',' || v_�շ���� || ',') = 0 Then
      --���ܴ��ڲ�������,�����Ƚ���ҩƷ�Ĵ���ɲ���ִ��,����������˹���(ZL_סԺ���ʼ�¼_Delete)�д���,�����������: 
      --�ڵ��ñ�����ʱ: 
      --   1.������Ѿ�ִ�е�,���Ϊ����ִ��(ִ��״̬=2);�������ʹ����д����ⲿ������(ZL_סԺ���ʼ�¼_Delete):��:���ִ��״̬=2,���Ҳ������ʵ�,���Ϊ1(��ִ��) 
      --      ԭ������Ϊ��ҩƷ��ֻ�ܴ�������״̬.��ִ��;2-δִ�� 
      --   2.�����δִ�е�,��ִ��״̬����Ϊ0,�������ʹ����м�¼״̬���ֲ��� 
    
      --��ҩƷ����û��ȡ��ִ�еĲ���,���Զ���ִ�е�Ҫ�ȸ�״̬���ܵ����� 
      Update סԺ���ü�¼ Set ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 0, 0, 2) Where ID = Id_In;
      If Sql%NotFound Then
        Update ������ü�¼ Set ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 0, 0, 2) Where ID = Id_In;
      End If;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˷�������_Audit;
/

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Create Or Replace Procedure Zl_���˷�������_Insert
(
  Id_In         In ���˷�������.����id%Type,
  �շ�ϸĿid_In In ���˷�������.�շ�ϸĿid%Type,
  ���벿��id_In In ���˷�������.���벿��id%Type,
  ����_In       In ���˷�������.����%Type,
  ������_In     In ���˷�������.������%Type,
  ����ʱ��_In   In ���˷�������.����ʱ��%Type,
  �������_In   In ���˷�������.�������%Type,
  ɾ����־_In   In Integer := 0,
  ��ҩid_In     In Integer := 0,
  ����ԭ��_In   In ���˷�������.����ԭ��%Type := Null,
  ��Һ����_In   In Number := 1
) As
  --���ܣ�������������
  --��Σ�
  --     �������_In ��ҩƷ��������Ч:0-δ��ҩ(��);1-�ѷ�ҩ(��);����Ϊ0
  --     ɾ����־_In ɾ�����˷�������ʱ������:1-ɾ��ʱ�����������,0-ɾ��ʱ,�����������������ɾ��(��Ϊ���ܳ�������������ʱ,������ִ�к�δִ������״̬)
  --     ��Һ����_In �Ƿ� ��Һ��ҩ��¼ ״̬�ֶΡ�1-Ҫ���£�0-������
  --˵����
  --    ���ÿ���������סԺ���ü�¼��Ҳ����������������ü�¼
  n_��˲���id   ���˷�������.��˲���id%Type;
  n_�������     ���˷�������.�������%Type;
  n_�������Ҳ��� ���˷�������.��˲���id%Type;
  n_ִ�п��Ҳ��� ���˷�������.��˲���id%Type;
  n_��������     ��������.��������%Type;
  n_ִ��״̬     סԺ���ü�¼.ִ��״̬%Type;
  v_�շ����     סԺ���ü�¼.�շ����%Type;
  n_ʵ������     ҩƷ�շ���¼.ʵ������%Type;
  n_ҽ��id       סԺ���ü�¼.ҽ�����%Type;
  n_��ҳid       סԺ���ü�¼.Id%Type;
  v_No           סԺ���ü�¼.No%Type;
  n_����id       סԺ���ü�¼.����id%Type;
  n_���˿���id   סԺ���ü�¼.���˿���id%Type;
  n_Icu����id    סԺ���ü�¼.���˿���id%Type;
  n_����������   ҩƷ�շ���¼.ʵ������%Type;

  n_Temp    Number;
  n_Icu     Number;
  n_Count   Number;
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);

Begin
  Select Count(1)
  Into n_Count
  From (Select 1
         From סԺ���ü�¼ A, סԺ���ü�¼ B
         Where a.No = b.No And Mod(a.��¼����, 10) = Mod(b.��¼����, 10) And a.��� = b.��� And b.Id = Id_In Having
          Nvl(Sum(a.���ʽ��), 0) <> 0
         Union All
         Select 1
         From ������ü�¼ A, ������ü�¼ B
         Where a.No = b.No And Mod(a.��¼����, 10) = Mod(b.��¼����, 10) And a.��� = b.��� And b.Id = Id_In Having
          Nvl(Sum(a.���ʽ��), 0) <> 0);
  If Nvl(n_Count, 0) > 0 Then
    v_Err_Msg := '�������ʵļ�¼�ѱ����˽���';
    Raise Err_Item;
  End If;

  Select a.�շ����, a.No, Nvl(b.��������, 0), Decode(Nvl(�������_In, 0), 0, a.���˲���id, a.ִ�в���id), ҽ�����, ����id, Nvl(��ҳid, 0)
  Into v_�շ����, v_No, n_��������, n_��˲���id, n_ҽ��id, n_����id, n_��ҳid
  From (Select �շ����, NO, �շ�ϸĿid, ���˲���id, ִ�в���id, ҽ�����, ����id, ��ҳid
         From סԺ���ü�¼
         Where ID = Id_In
         Union All
         Select �շ����, NO, �շ�ϸĿid, ���˲���id, ִ�в���id, ҽ�����, ����id, ��ҳid
         From ������ü�¼
         Where ID = Id_In) A, �������� B
  Where a.�շ�ϸĿid = b.����id(+);

  If Nvl(n_��ҳid, 0) <> 0 Then
    n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
    n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
    If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
      Begin
        Select ��˱�־, ״̬
        Into n_��˱�־, n_סԺ״̬
        From ������ҳ
        Where ����id = Nvl(n_����id, 0) And ��ҳid = Nvl(n_��ҳid, 0);
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
  
  End If;

  Select Max(��Ժ����id)
  Into n_���˿���id
  From ������Ϣ A, ������ҳ B
  Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = n_����id;

  Select Decode(Count(1), 0, 0, 1) Into n_Icu From ��������˵�� B Where b.����id = n_���˿���id And b.�������� = 'ICU';
  If n_Icu = 1 Then
    --����Ƿ��ǵ�ǰ����Ա������ICU
    Select Decode(Count(Distinct a.�û���), 0, 0, 1)
    Into n_Icu
    From �ϻ���Ա�� A, ��������˵�� B, ������Ա C
    Where a.�û��� = User And a.��Աid = c.��Աid And c.����id = b.����id And b.�������� = 'ICU';
  End If;

  If n_Icu = 1 Then
    n_Icu����id := n_���˿���id;
    If Nvl(�������_In, 0) = 0 Then
      n_��˲���id := n_���˿���id;
    End If;
  End If;

  If Instr(',5,6,7', ',' || v_�շ����) > 0 Or v_�շ���� = '4' And Nvl(n_��������, 0) = 1 Then
    n_������� := �������_In;
  Else
    n_������� := 0;
  End If;

  --ȡ����ǰ�������������(����������ʱ������ȡ������Ϊ����id��ͬ��ÿ�����οɷֱ�����)
  If ��ҩid_In = 0 Then
    If Nvl(ɾ����־_In, 0) = 1 Then
      Delete ���˷������� Where ����id = Id_In And ״̬ = 0;
    Else
      Delete ���˷������� Where ����id = Id_In And ������� = n_������� And ״̬ = 0;
    End If;
  End If;
  If ����_In <> 0 Then
    --��˿���
    --1.ҩƷ���û�������õ�����:
    --    a. ���δִ��,�򰴲��˲�����Ϊ��˲���;
    --    b. �����ִ��,��ִ�в���ID��Ϊ��˲���
    --2.ҽ�����ҿ����ķ���(����������<>���˿���)��������˿���Ϊ��������,
    --  ����������������ڲ������ٴ�����,�����ʿ���Ϊ��������(����ʿ�ǲ������������ҷ����ķ���)
    --  (���ִ�п���x����a��b����������a��b��������������Ϊ����ȷ�Ͽ���,ȡ��һ�������a����ͬʱ�ǲ��˲�������ֻ��a������ȷ��)��
    --3.���������ķ���,û�о���������˵�,������˿���Ϊ���˲���(����Ѿ���ִ�У���Ϊִ�в���,����Ϊ���˲���)
    --  ����������˵�,������˿���Ϊִ�п��ҡ�
    --  ���ִ�п��������ڲ������ٴ����ң���������˿���Ϊ��������
    --  (���ִ�п���x����a��b����������a��b��������������Ϊ����ȷ�Ͽ���,ȡ��һ�������a����ͬʱ�ǲ��˲�������ֻ��a������ȷ��)��
    --4.�����ǰ����Ա������ICU,���Ҳ��˵�ǰ����ҲΪICU�Լ�δִ�е���Ŀ,��ICU���������������.
  
    If Nvl(n_��������, 0) = 1 Then
      If Nvl(�������_In, 0) = 0 Then
        --Ҫ���δִ�е�����������ڵ�����������,�Ż�ͨ��
        Select Sum(Nvl(����, 0) * Nvl(ʵ������, 0))
        Into n_ʵ������
        From ҩƷ�շ���¼
        Where ������� Is Null And ����id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0;
      
        If Nvl(n_ʵ������, 0) < Nvl(����_In, 0) Then
          Select '�ڵ��ݺ�<<' || v_No || '>>��������Ϊ:' || Chr(13) || ���� || '-' || ���� || Chr(13) || '����������(' ||
                  LTrim(To_Char(����_In, '9999999990.99')) || ')�����˴���������(' ||
                  LTrim(To_Char(Nvl(n_ʵ������, 0), '9999999990.99')) || '),���ܽ�����������!'
          Into v_Err_Msg
          From �շ���ĿĿ¼
          Where ID = �շ�ϸĿid_In;
          Raise Err_Item;
        End If;
      End If;
    Else
      --a.ִ�п��һ�����������:��0:δִ��;1:��ȫִ��;2:����ִ��
      Select Decode(b.����id, Null, a.ִ�в���id, Nvl(c.����id, b.����id)), Decode(a.ִ��״̬, 1, 1, 2, 1, 0)
      Into n_ִ�п��Ҳ���, n_ִ��״̬
      From (Select ִ�в���id, ���˲���id, ִ��״̬
             From סԺ���ü�¼
             Where ID = Id_In
             Union All
             Select ִ�в���id, ���˲���id, ִ��״̬
             From ������ü�¼
             Where ID = Id_In) A, �������Ҷ�Ӧ B, �������Ҷ�Ӧ C
      Where a.ִ�в���id = b.����id(+) And a.ִ�в���id = c.����id(+) And a.���˲���id = c.����id(+) And Rownum < 2;
    
      --b.�������һ�����������
      Select Decode(b.����id, Null, a.��������id, Nvl(c.����id, b.����id))
      Into n_�������Ҳ���
      From (Select ��������id, ���˲���id
             From סԺ���ü�¼
             Where ID = Id_In
             Union All
             Select ��������id, ���˲���id
             From ������ü�¼
             Where ID = Id_In) A, �������Ҷ�Ӧ B, �������Ҷ�Ӧ C
      Where a.��������id = b.����id(+) And a.��������id = c.����id(+) And a.���˲���id = c.����id(+) And Rownum < 2;
    
      For v_���� In (Select �շ����, Nvl(ִ��״̬, 0) As ִ��״̬, ���˲���id, ִ�в���id, ��������id, ���˿���id, ������, ����Ա����
                   From סԺ���ü�¼
                   Where ID = Id_In
                   Union All
                   Select �շ����, Nvl(ִ��״̬, 0) As ִ��״̬, ���˲���id, ִ�в���id, ��������id, ���˿���id, ������, ����Ա����
                   From ������ü�¼
                   Where ID = Id_In) Loop
      
        If Instr('567', v_����.�շ����, 1) > 0 Then
          n_Temp       := Case
                            When �������_In Is Null Then
                             Nvl(v_����.ִ��״̬, 0)
                            Else
                             �������_In
                          End;
          n_��˲���id := Case
                        When n_Temp = 0 Then
                         v_����.���˲���id
                        Else
                         v_����.ִ�в���id
                      End;
          If n_Temp = 0 And n_Icu = 1 Then
            --ICUΪICU����
            n_��˲���id := n_Icu����id;
          End If;
        Else
          If v_����.��������id = v_����.���˿���id Then
            --�ٴ������ķ���
            If Nvl(v_����.������, '-') = v_����.����Ա���� Or v_����.������ Is Null Then
              --�������
              n_��˲���id := Case
                            When n_ִ��״̬ = 1 Then
                             v_����.ִ�в���id
                            Else
                             v_����.���˲���id
                          End;
            Else
              n_��˲���id := n_ִ�п��Ҳ���;
            End If;
          Else
            n_��˲���id := n_ִ�п��Ҳ���;
          End If;
          If n_ִ��״̬ = 0 And n_Icu = 1 Then
            --ICUΪICU����
            n_��˲���id := n_Icu����id;
          End If;
        End If;
      End Loop;
    
      If Instr(',5,6,7', ',' || v_�շ����) > 0 And Nvl(�������_In, Nvl(n_ִ��״̬, 0)) = 0 Then
        --��Ҫ���δִ�е�����������ڵ�����������,�Ż�ͨ��
        Select Sum(Nvl(����, 0) * Nvl(ʵ������, 0))
        Into n_ʵ������
        From ҩƷ�շ���¼
        Where ������� Is Null And ����id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0;
        If Nvl(n_ʵ������, 0) < Nvl(����_In, 0) Then
          Select '�ڵ��ݺ�<<' || v_No || '>>��ҩƷΪ:' || Chr(13) || ���� || '-' || ���� || Chr(13) || '����������(' ||
                  LTrim(To_Char(����_In, '9999999990.99')) || ')���ܴ��ڴ���ҩ����(' ||
                  LTrim(To_Char(Nvl(n_ʵ������, 0), '9999999990.99')) || '),���ܽ�����������!'
          Into v_Err_Msg
          From �շ���ĿĿ¼
          Where ID = �շ�ϸĿid_In;
          Raise Err_Item;
        End If;
      End If;
    End If;
    --�����������:��ǰ��������+�Ѿ������������ܴ���ĳ����������
    Select Sum(Nvl(����, 0)) Into n_���������� From ���˷������� Where ����id = Id_In And Nvl(״̬, 0) <> 2;
  
    Select Sum(Nvl(����, 1) * Nvl(����, 0))
    Into n_ʵ������
    From (Select a.����, a.����
           From סԺ���ü�¼ A, סԺ���ü�¼ B
           Where a.No = b.No And Mod(a.��¼����, 10) = Mod(b.��¼����, 10) And a.��¼״̬ In (0, 1, 3) And a.��� = b.��� And
                 b.Id = Id_In
           Union All
           Select a.����, a.����
           From ������ü�¼ A, ������ü�¼ B
           Where a.No = b.No And Mod(a.��¼����, 10) = Mod(b.��¼����, 10) And a.��¼״̬ In (0, 1, 3) And a.��� = b.��� And
                 b.Id = Id_In);
  
    If Nvl(n_ʵ������, 0) < Nvl(n_����������, 0) + Nvl(����_In, 0) Then
      Select '�ڵ��ݺ�<<' || v_No || '>>�շ���Ŀ:' || Chr(13) || ���� || '-' || ���� || Chr(13) || '����������(' ||
              LTrim(To_Char(Nvl(n_����������, 0) + Nvl(����_In, 0), '9999999990.99')) || ')���ܴ��ڼ�������(' ||
              LTrim(To_Char(Nvl(n_ʵ������, 0), '9999999990.99')) || '),���ܽ�����������!'
      Into v_Err_Msg
      From �շ���ĿĿ¼
      Where ID = �շ�ϸĿid_In;
      Raise Err_Item;
    End If;
  
    If n_ҽ��id <> 0 And ��ҩid_In <> 0 Then
      --�������Һ��ҩ���ĵģ��������ر��ֶ�
      If Nvl(��Һ����_In, 0) = 1 Then
        Select Count(1)
        Into n_Temp
        From ��Һ��ҩ״̬
        Where ��ҩid = ��ҩid_In And �������� = 9 And ����ʱ�� = ����ʱ��_In;
        If n_Temp = 0 Then
          Insert Into ��Һ��ҩ״̬
            (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
          Values
            (��ҩid_In, 9, ������_In, ����ʱ��_In, ����ԭ��_In);
        End If;
        Update ��Һ��ҩ��¼
        Set ������Ա = ������_In, ����ʱ�� = ����ʱ��_In, ����״̬ = 9
        Where ID = ��ҩid_In
        Returning ����id Into n_��˲���id;
      Else
        Select ����id Into n_��˲���id From ��Һ��ҩ��¼ Where ID = ��ҩid_In;
      End If;
    End If;
  
    Insert Into ���˷�������
      (����id, �������, �շ�ϸĿid, ��˲���id, ���벿��id, ����, ������, ����ʱ��, ״̬, ����ԭ��)
    Values
      (Id_In, n_�������, �շ�ϸĿid_In, n_��˲���id, ���벿��id_In, ����_In, ������_In, ����ʱ��_In, 0, ����ԭ��_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˷�������_Insert;
/

--139063:Ƚ����,2019-04-04,�������۲��˰��������̾���
CREATE OR REPLACE Procedure Zl_����δ�����_Recalc
(
  ����id_In סԺ���ü�¼.����id%Type,
  ��ҳid_In סԺ���ü�¼.��ҳid%Type
) As
  v_�ѱ�     �ѱ�.����%Type;
  n_�������� ������ҳ.��������%Type;

  v_No       סԺ���ü�¼.No%Type;
  n_ʵ�ս�� סԺ���ü�¼.ʵ�ս��%Type;
  n_������� �������.�������%Type;
  n_С��λ�� Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  Select �ѱ�, �������� Into v_�ѱ�, n_�������� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;

  --�����ж� 
  --a.��ǰ���ǰ���������ܼ����ۿ�ģʽ 
  v_Counter := To_Number(Nvl(zl_GetSysParameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '��ǰ�ѱ�ʹ����������ܼ����ۿ�ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --b.��ǰ�ѱ���ʹ��ҩƷ���ɱ��ۼ��մ��۵ķѱ� 
  v_Counter := 0;
  Select Count(�ѱ�) Into v_Counter From �ѱ���ϸ Where �ѱ� = v_�ѱ� And ���㷽�� = 1;
  If v_Counter > 0 Then
    v_Error := '��ǰ�ѱ�ʹ��ҩƷ���ɱ��ۼ��մ���ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --�������۲��˿���û��סԺ���ü�¼ 
  If Nvl(n_��������, 0) <> 1 Then
    --c.û��δ����� 
    Begin
      Select ������� Into n_������� From ������� Where ����id = ����id_In And ���� = 2 And ���� = 1;
    Exception
      When Others Then
        n_������� := 0;
    End;
    --������δ����ã������Ǳ���סԺ�����ģ��ں���ִ��ʱ���жϱ����Ƿ���δ����ϸ 
    If n_������� = 0 Then
      v_Counter := 0;
      --�������Ϊ0ʱ��Ҳ�����з��ã����з��ö����շѣ� 
      Select Count(ID) Into v_Counter From סԺ���ü�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And Rownum < 2;
      If v_Counter = 0 Then
        v_Error := '���˲�����δ�����,���ý��з�������!';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  --d.�������뱾��סԺ�ѱ�ͬ�ķ�����ϸ 
  v_Counter := 0;
  Select Count(ID)
  Into v_Counter
  From סԺ���ü�¼
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And �ѱ� <> v_�ѱ� And Rownum < 2;
  If v_Counter = 0 And Nvl(n_��������, 0) <> 1 Then
    v_Error := '���˲������뱾��סԺ�ѱ�ͬ�ķ�����ϸ ,���ý��з�������!';
    Raise Err_Custom;
  End If;

  --ִ�� 
  If Nvl(v_Counter, 0) <> 0 Then
    v_Counter  := 0;
    d_Sysdate  := Sysdate;
    n_С��λ�� := To_Number(Nvl(zl_GetSysParameter(9), 2));
    For r_Fee In (Select ����id, ��ҳid, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, �Ӱ��־, ���ӱ�־, Ӥ����,
                         ������Ŀid, �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��, ����Ա���, ����Ա����, ҽ��С��id, Nvl(Sum(Ӧ�ս��), 0) Ӧ�ս��,
                         Nvl(Sum(ʵ�ս��), 0) ʵ�ս��
                  From (Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ���ʷ���, ����, �Ա�,
                                ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
                                �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���,
                                ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, ҽ��С��id
                         From סԺ���ü�¼
                         Union All
                         Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ���ʷ���, ����, �Ա�,
                                ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
                                �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���,
                                ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, ҽ��С��id
                         From HסԺ���ü�¼)
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼״̬ <> 0 And ���ʷ��� = 1
                  Group By ����id, ��ҳid, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, �Ӱ��־, ���ӱ�־,
                           Ӥ����, ������Ŀid, �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��, ����Ա���, ����Ա����, ҽ��С��id
                  Having(Nvl(Sum(ʵ�ս��), 0) <> Nvl(Sum(���ʽ��), 0) Or Nvl(Sum(���ʽ��), 0) = 0) And Not(Nvl(Sum(Ӧ�ս��), 0) = 0 And Nvl(Sum(ʵ�ս��), 0) = 0)
                  Order By ��������id, ������, ����Ա����) Loop
      --          ������δ��ķ���,������ϸ���ֽ���,�Լ����ʺ�����,��Щ��¼�п�����ת��󱸱� 
      --          1.�ſ�����ȫ�����ʵļ�¼(Sum(Ӧ�ս��)=Sum(Ӧ�ս��)) 
      --          2.�ſ����޴��۳���ļ��ʺ������ʵļ�¼(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)=0) 
      --          3.���ſ����۳�������˵������ʵļ�¼��Ҫ��ԭ�����¼һ����������(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)<>0) 
      --          4.���ſ����۳���������ʵ�պͽ��ʶ�Ϊ��ļ�¼����Ϊ�Ļ�ԭ���ķѱ�ʱ��Ҫ�����ȥ 
      If r_Fee.Ӧ�ս�� <> 0 Then
        Begin
          Select ʵ�ս��
          Into n_ʵ�ս��
          From (Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
                 From �ѱ���ϸ
                 Where �շ�ϸĿid = r_Fee.�շ�ϸĿid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And
                       Nvl(���㷽��, 0) = 0
                 Union All
                 Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
                 From �ѱ���ϸ A
                 Where ������Ŀid = r_Fee.������Ŀid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And
                       Nvl(���㷽��, 0) = 0 And Not Exists
                  (Select 1 From �ѱ���ϸ B Where b.�ѱ� = a.�ѱ� And b.�շ�ϸĿid = r_Fee.�շ�ϸĿid));
        Exception
          When Others Then
            n_ʵ�ս�� := r_Fee.Ӧ�ս��;
        End;
      Else
        n_ʵ�ս�� := 0;
      End If;
      --�����������ԭʵ�յĲ�� 
      n_ʵ�ս�� := -1 * (r_Fee.ʵ�ս�� - n_ʵ�ս��);
    
      If n_ʵ�ս�� <> 0 Then
        --һ�ŵ��ݵĿ�������id,������,����Ա����,����Ҫ����ͬ���������֮һ����������µ��ݣ������û�б䣬һ�ŵ������100����ϸ 
        v_Thisinfo := r_Fee.��������id || r_Fee.������ || r_Fee.����Ա���� || r_Fee.����;
        If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
          v_No       := Nextno(14);
          v_Counter  := 1;
          v_Lastinfo := v_Thisinfo;
        Else
          v_Counter := v_Counter + 1;
        End If;
      
        Insert Into סԺ���ü�¼
          (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, �����־, ����id, ��ҳid, ��ʶ��, ����, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�,
           �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, ��ҩ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���,
           ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, ժҪ, �Ƿ���, ҽ�����, ҽ��С��id)
        Values
          (���˷��ü�¼_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, 0, Null, r_Fee.�����־, r_Fee.����id, r_Fee.��ҳid, r_Fee.��ʶ��,
           r_Fee.����, r_Fee.����, r_Fee.�Ա�, r_Fee.����, r_Fee.���˲���id, r_Fee.���˿���id, v_�ѱ�, r_Fee.�շ����, r_Fee.�շ�ϸĿid,
           r_Fee.���㵥λ, Null, Null, 0, 0, Null, r_Fee.�Ӱ��־, r_Fee.���ӱ�־, r_Fee.Ӥ����, r_Fee.������Ŀid, r_Fee.�վݷ�Ŀ, 0, 0, n_ʵ�ս��,
           Null, 1, Null, r_Fee.��������id, r_Fee.������, r_Fee.����ʱ��, d_Sysdate, r_Fee.ִ�в���id, 0, Null, Null, r_Fee.����Ա���,
           r_Fee.����Ա����, Decode(v_Counter, 1, 'ʵ��������', ''), 0, Null, r_Fee.ҽ��С��id);
      End If;
    End Loop;
  End If;

  If v_Counter = 0 Then
    If Nvl(n_��������, 0) <> 1 Then
      v_Error := '��������ԭ��֮һ,û�н��з�������:' || Chr(13) || Chr(13) || 'a.û�з��ֲ��˱���סԺ��δ�����.' || Chr(13) || 'b.����δ������ѽ����˷�������.' ||
                 Chr(13) || 'c.����ǰ�ѱ������ʵ�ճ����Ϊ��.';
      Raise Err_Custom;
    End If;
  Else
    --������� 
    n_ʵ�ս�� := 0;
    Select Sum(ʵ�ս��)
    Into n_ʵ�ս��
    From סԺ���ü�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 2 And �Ǽ�ʱ�� = d_Sysdate;
  
    Update ������� Set ������� = Nvl(�������, 0) + n_ʵ�ս�� Where ����id = ����id_In And ���� = 1 And ���� = 2;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, �������, Ԥ�����, ����) Values (����id_In, 1, n_ʵ�ս��, 0, 2);
    End If;
  
    --����δ����� 
    For r_Fee In (Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(ʵ�ս��) ʵ�ս��
                  From סԺ���ü�¼
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 2 And �Ǽ�ʱ�� = d_Sysdate
                  Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
      Update ����δ�����
      Set ��� = Nvl(���, 0) + r_Fee.ʵ�ս��
      Where ����id = ����id_In And Nvl(��ҳid, 0) = ��ҳid_In And Nvl(���˲���id, 0) = r_Fee.���˲���id And
            Nvl(���˿���id, 0) = r_Fee.���˿���id And Nvl(��������id, 0) = r_Fee.��������id And Nvl(ִ�в���id, 0) = r_Fee.ִ�в���id And
            ������Ŀid + 0 = r_Fee.������Ŀid And ��Դ;�� + 0 = 2;
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, ��ҳid_In, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2, r_Fee.ʵ�ս��);
      End If;
    End Loop;
  End If;

  --�������۲��������������
  If Nvl(n_��������, 0) = 1 Then
    Begin
      Zl_����δ���������_Recalc(����id_In, ��ҳid_In);
    Exception
      When Others Then
        Null; --����
    End;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����δ�����_Recalc;
/

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Create Or Replace Procedure Zl_����δ���������_Recalc
(
  ����id_In סԺ���ü�¼.����id%Type,
  ��ҳid_In ������ҳ.��ҳid%Type := Null
) As
  v_�ѱ�     �ѱ�.����%Type;
  v_No       ������ü�¼.No%Type;
  n_ʵ�ս�� ������ü�¼.ʵ�ս��%Type;
  n_������� �������.�������%Type;
  n_С��λ�� Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  If Nvl(��ҳid_In, 0) = 0 Then
    Select �ѱ� Into v_�ѱ� From ������Ϣ Where ����id = ����id_In;
  Else
    Select Nvl(b.�ѱ�, a.�ѱ�)
    Into v_�ѱ�
    From ������Ϣ A, ������ҳ B
    Where a.����id = b.����id And b.����id = ����id_In And b.��ҳid = ��ҳid_In;
  End If;

  --�����ж� 
  --a.��ǰ���ǰ���������ܼ����ۿ�ģʽ 
  v_Counter := To_Number(Nvl(zl_GetSysParameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '��ǰ�ѱ�ʹ����������ܼ����ۿ�ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --b.��ǰ�ѱ���ʹ��ҩƷ���ɱ��ۼ��մ��۵ķѱ� 
  v_Counter := 0;
  Select Count(�ѱ�) Into v_Counter From �ѱ���ϸ Where �ѱ� = v_�ѱ� And ���㷽�� = 1;
  If v_Counter > 0 Then
    v_Error := '��ǰ�ѱ�ʹ��ҩƷ���ɱ��ۼ��մ���ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --c.û��δ����� 
  Begin
    Select ������� Into n_������� From ������� Where ����id = ����id_In And ���� = 1 And ���� = 1;
  Exception
    When Others Then
      n_������� := 0;
  End;
  --������δ����ã������Ǳ���סԺ�����ģ��ں���ִ��ʱ���жϱ����Ƿ���δ����ϸ 
  If n_������� = 0 Then
    v_Counter := 0;
    --�������Ϊ0ʱ��Ҳ�����з��ã����з��ö����շѣ� 
    Select Count(ID) Into v_Counter From ������ü�¼ Where ����id = ����id_In And Rownum < 2;
    If v_Counter = 0 Then
      v_Error := '���˲�����δ�����,���ý��з�������!';
      Raise Err_Custom;
    End If;
  End If;

  --d.�������뱾��סԺ�ѱ�ͬ�ķ�����ϸ 
  v_Counter := 0;
  Select Count(ID) Into v_Counter From ������ü�¼ Where ����id = ����id_In And �ѱ� <> v_�ѱ� And Rownum < 2;
  If v_Counter = 0 Then
    v_Error := '���˲������뱾��סԺ�ѱ�ͬ�ķ�����ϸ ,���ý��з�������!';
    Raise Err_Custom;
  End If;

  --ִ�� 
  v_Counter  := 0;
  d_Sysdate  := Sysdate;
  n_С��λ�� := To_Number(Nvl(zl_GetSysParameter(9), 2));
  For r_Fee In (Select ����id, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
                       �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��, ����Ա���, ����Ա����, Nvl(Sum(Ӧ�ս��), 0) Ӧ�ս��, Nvl(Sum(ʵ�ս��), 0) ʵ�ս��,
                       ��ҳid, �Һ�id
                From (Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��,
                              ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����,
                              Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id,
                              ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, ��ҳid, �Һ�id
                       From ������ü�¼
                       Union All
                       Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��,
                              ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����,
                              Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id,
                              ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, ��ҳid, �Һ�id
                       From H������ü�¼)
                Where ����id = ����id_In And ��¼״̬ <> 0 And ���ʷ��� = 1
                Group By ����id, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
                         �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��, ����Ա���, ����Ա����, ��ҳid, �Һ�id
                Having(Nvl(Sum(ʵ�ս��), 0) <> Nvl(Sum(���ʽ��), 0) Or Nvl(Sum(���ʽ��), 0) = 0) And Not(Nvl(Sum(Ӧ�ս��), 0) = 0 And Nvl(Sum(ʵ�ս��), 0) = 0)
                Order By ��������id, ������, ����Ա����) Loop
    --          ������δ��ķ���,������ϸ���ֽ���,�Լ����ʺ�����,��Щ��¼�п�����ת��󱸱� 
    --          1.�ſ�����ȫ�����ʵļ�¼(Sum(Ӧ�ս��)=Sum(Ӧ�ս��)) 
    --          2.�ſ����޴��۳���ļ��ʺ������ʵļ�¼(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)=0) 
    --          3.���ſ����۳�������˵������ʵļ�¼��Ҫ��ԭ�����¼һ����������(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)<>0) 
    --          4.���ſ����۳���������ʵ�պͽ��ʶ�Ϊ��ļ�¼����Ϊ�Ļ�ԭ���ķѱ�ʱ��Ҫ�����ȥ 
    If r_Fee.Ӧ�ս�� <> 0 Then
      Begin
        Select ʵ�ս��
        Into n_ʵ�ս��
        From (Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
               From �ѱ���ϸ
               Where �շ�ϸĿid = r_Fee.�շ�ϸĿid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And Nvl(���㷽��, 0) = 0
               Union All
               Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
               From �ѱ���ϸ A
               Where ������Ŀid = r_Fee.������Ŀid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And Nvl(���㷽��, 0) = 0 And
                     Not Exists (Select 1 From �ѱ���ϸ B Where b.�ѱ� = a.�ѱ� And b.�շ�ϸĿid = r_Fee.�շ�ϸĿid));
      Exception
        When Others Then
          n_ʵ�ս�� := r_Fee.Ӧ�ս��;
      End;
    Else
      n_ʵ�ս�� := 0;
    End If;
    --�����������ԭʵ�յĲ�� 
    n_ʵ�ս�� := -1 * (r_Fee.ʵ�ս�� - n_ʵ�ս��);
  
    If n_ʵ�ս�� <> 0 Then
      --һ�ŵ��ݵĿ�������id,������,����Ա����,����Ҫ����ͬ���������֮һ����������µ��ݣ������û�б䣬һ�ŵ������100����ϸ 
      v_Thisinfo := r_Fee.��������id || r_Fee.������ || r_Fee.����Ա���� || ' ';
      If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
        v_No       := Nextno(14);
        v_Counter  := 1;
        v_Lastinfo := v_Thisinfo;
      Else
        v_Counter := v_Counter + 1;
      End If;
    
      Insert Into ������ü�¼
        (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
         ���մ���id, ����, ����, ��ҩ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
         ִ�в���id, ִ��״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, ժҪ, �Ƿ���, ҽ�����, ��ҳid, �Һ�id, ���˲���id)
      Values
        (���˷��ü�¼_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, Null, r_Fee.�����־, r_Fee.����id, r_Fee.��ʶ��, r_Fee.����,
         r_Fee.�Ա�, r_Fee.����, r_Fee.���˿���id, v_�ѱ�, r_Fee.�շ����, r_Fee.�շ�ϸĿid, r_Fee.���㵥λ, Null, Null, 0, 0, Null,
         r_Fee.�Ӱ��־, r_Fee.���ӱ�־, r_Fee.Ӥ����, r_Fee.������Ŀid, r_Fee.�վݷ�Ŀ, 0, 0, n_ʵ�ս��, Null, 1, Null, r_Fee.��������id,
         r_Fee.������, r_Fee.����ʱ��, d_Sysdate, r_Fee.ִ�в���id, 0, Null, Null, r_Fee.����Ա���, r_Fee.����Ա����,
         Decode(v_Counter, 1, 'ʵ��������', ''), 0, Null, r_Fee.��ҳid, r_Fee.�Һ�id, r_Fee.���˲���id);
    End If;
  End Loop;

  If v_Counter = 0 Then
    v_Error := '��������ԭ��֮һ,û�н��з�������:' || Chr(13) || Chr(13) || 'a.û�з��ֲ��˱���סԺ��δ�����.' || Chr(13) || 'b.����δ������ѽ����˷�������.' ||
               Chr(13) || 'c.����ǰ�ѱ������ʵ�ճ����Ϊ��.';
    Raise Err_Custom;
  Else
    --������� 
    n_ʵ�ս�� := 0;
    Select Sum(ʵ�ս��)
    Into n_ʵ�ս��
    From ������ü�¼
    Where ����id = ����id_In And ��¼���� = 2 And Nvl(�����־, 0) <> 4 And �Ǽ�ʱ�� = d_Sysdate;
    Update ������� Set ������� = Nvl(�������, 0) + n_ʵ�ս�� Where ����id = ����id_In And ���� = 1 And ���� = 1;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, �������, Ԥ�����, ����) Values (����id_In, 1, n_ʵ�ս��, 0, 1);
    End If;
  
    --����δ����� 
    For r_Fee In (Select ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(ʵ�ս��) ʵ�ս��
                  From ������ü�¼
                  Where ����id = ����id_In And ��¼���� = 2 And �Ǽ�ʱ�� = d_Sysdate
                  Group By ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
      Update ����δ�����
      Set ��� = Nvl(���, 0) + r_Fee.ʵ�ս��
      Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(r_Fee.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_Fee.���˲���id, 0) And
            Nvl(���˿���id, 0) = Nvl(r_Fee.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Fee.��������id, 0) And
            Nvl(ִ�в���id, 0) = Nvl(r_Fee.ִ�в���id, 0) And ������Ŀid + 0 = Nvl(r_Fee.������Ŀid, 0) And ��Դ;�� + 0 = 2;
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, r_Fee.��ҳid, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2, r_Fee.ʵ�ս��);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����δ���������_Recalc;
/

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Create Or Replace Procedure Zl_���ﻮ�ۼ�¼_Insert
(
  No_In         ������ü�¼.No%Type,
  ���_In       ������ü�¼.���%Type,
  ����id_In     ������ü�¼.����id%Type,
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type,
  ��ʶ��_In     ������ü�¼.��ʶ��%Type,
  ���ʽ_In   ������ü�¼.���ʽ%Type,
  ����_In       ������ü�¼.����%Type,
  �Ա�_In       ������ü�¼.�Ա�%Type,
  ����_In       ������ü�¼.����%Type,
  �ѱ�_In       ������ü�¼.�ѱ�%Type,
  �Ӱ��־_In   ������ü�¼.�Ӱ��־%Type,
  ���˿���id_In ������ü�¼.���˿���id%Type,
  ��������id_In ������ü�¼.��������id%Type,
  ������_In     ������ü�¼.������%Type,
  ��������_In   ������ü�¼.��������%Type,
  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
  �շ����_In   ������ü�¼.�շ����%Type,
  ���㵥λ_In   ������ü�¼.���㵥λ%Type,
  ��ҩ����_In   ������ü�¼.��ҩ����%Type,
  ����_In       ������ü�¼.����%Type,
  ����_In       ������ü�¼.����%Type,
  ���ӱ�־_In   ������ü�¼.���ӱ�־%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
  ������Ŀid_In ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
  ��׼����_In   ������ü�¼.��׼����%Type,
  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
  ҩƷժҪ_In   ҩƷ�շ���¼.ժҪ%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ����ժҪ_In   ������ü�¼.ժҪ%Type := Null,
  ҽ�����_In   ������ü�¼.ҽ�����%Type := Null,
  Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
  ����_In       ҩƷ�շ���¼.����%Type := Null,
  �÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
  ��Ч_In       ҩƷ�շ���¼.����%Type := Null,
  �Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
  ������Դ_In   Number := 1,
  ���ձ���_In   ������ü�¼.���ձ���%Type := Null,
  ��������_In   ������ü�¼.��������%Type := Null,
  ������Ŀ��_In ������ü�¼.������Ŀ��%Type := Null,
  ���մ���id_In ������ü�¼.���մ���id%Type := Null,
  ��ҩ��̬_In   ������ü�¼.����%Type := Null,
  ��������_In   Number := 0,
  ����_In       ҩƷ�շ���¼.����%Type := Null,
  ִ����_In     ������ü�¼.ִ����%Type := Null,
  ���˲���id_In ������ü�¼.���˲���id%Type := Null
) As
  --���ܣ�����һ�����ﻮ�۵���
  --������
  --   ������Դ_IN:1-���ﲡ��,2-סԺ����
  --     ��ҳID_IN:סԺ���˻���ʱ�á�
  --   ҩƷժҪ_IN:�޸ı����µ���ʱ�á�Ŀǰ�������ҩƷ�շ���¼��ժҪ�С�
  --         �µ���(��¼״̬=1)��¼���޸ĵ�ԭ���ݺš�
  v_����id ������ü�¼.Id%Type;
  n_����   ���˹Һż�¼.����%Type;
  n_�Һ�id ���˹Һż�¼.Id%Type;
  n_��ҳid ������ü�¼.��ҳid%Type;

  --��ʱ����
  v_�÷�       ҩƷ�շ���¼.�÷�%Type;
  v_�巨       ҩƷ�շ���¼.���%Type;
  n_Dec        Number;
  v_���ʽ   ҽ�Ƹ��ʽ.����%Type;
  v_�ѱ�����   �ѱ�.����%Type;
  n_�²���ģʽ Number;
  n_����С��   Number;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
Begin
  --������С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_Dec, n_����С��
  From Dual;

  n_��ҳid := ��ҳid_In;
  If Nvl(n_��ҳid, 0) = 0 Then
    Select Max(��ҳid) Into n_��ҳid From ������ҳ Where ����id = ����id_In And �������� = 1 And ��Ժ���� Is Null;
  End If;

  --������ü�¼
  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
  --�Ƿ��Ǽ���Һŵ�
  If Nvl(ҽ�����_In, 0) <> 0 Then
    Begin
      Select Nvl(Max(����), 0), Max(ID)
      Into n_����, n_�Һ�id
      From ���˹Һż�¼
      Where NO In (Select �Һŵ� From ����ҽ����¼ Where ID = Nvl(ҽ�����_In, 0)) And ����id = ����id_In;
    Exception
      When Others Then
        n_����   := Null;
        n_�Һ�id := Null;
    End;
  End If;

  Insert Into ������ü�¼
    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����, ��ҩ����,
     �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ժҪ, ҽ�����, ������Ŀ��, ���ձ���,
     ���մ���id, ��������, ����, �Ƿ���, �Һ�id, ��ҳid, ���ʽ, ִ����, ִ��ʱ��, ִ��״̬, ���˲���id)
  Values
    (v_����id, 1, No_In, 0, ���_In, Decode(��������_In, 0, Null, ��������_In), Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), Nvl(������Դ_In, 1),
     Decode(����id_In, 0, Null, ����id_In), Decode(��ʶ��_In, 0, Null, ��ʶ��_In), ����_In, �Ա�_In, ����_In, ���˿���id_In, �ѱ�_In, �շ����_In,
     �շ�ϸĿid_In, ���㵥λ_In, ����_In, ����_In, ��ҩ����_In, �Ӱ��־_In, ���ӱ�־_In, ������Ŀid_In, �վݷ�Ŀ_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In, 0,
     ����Ա����_In, ��������id_In, ������_In, ����ʱ��_In, �Ǽ�ʱ��_In, ִ�в���id_In, ����ժҪ_In, ҽ�����_In, ������Ŀ��_In, ���ձ���_In, ���մ���id_In, ��������_In,
     ��ҩ��̬_In, Nvl(n_����, 0), n_�Һ�id, n_��ҳid, ���ʽ_In, ִ����_In, Decode(ִ����_In, Null, Null, �Ǽ�ʱ��_In),
     Decode(ִ����_In, Null, 0, 2), ���˲���id_In);

  --ҩƷ���������ϲ���
  If �շ����_In In ('4', '5', '6', '7') Then
    --ҩƷ�÷��巨�ֽ�
    If �÷�_In Is Not Null Then
      If Instr(�÷�_In, '|') > 0 Then
        v_�÷� := Substr(�÷�_In, 1, Instr(�÷�_In, '|') - 1);
        v_�巨 := Substr(�÷�_In, Instr(�÷�_In, '|') + 1);
      Else
        v_�÷� := �÷�_In;
      End If;
    End If;
    Zl_ҩƷ�շ���¼_���۳���(v_����id, ҩƷժҪ_In, Ƶ��_In, ����_In, v_�÷�, v_�巨, ��Ч_In, �Ƽ�����_In, n_��ҳid, ��������_In, ����_In);
  End If;

  --���²��ݲ�����Ϣ
  If ���_In = 1 And ����id_In Is Not Null Then
  
    If ���ʽ_In Is Not Null And ������Դ_In = 1 Then
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
    End If;
  
    If �ѱ�_In Is Not Null Then
      Select Max(����) Into v_�ѱ����� From �ѱ� Where ���� = �ѱ�_In; --2-��̬�ѱ𲻸���
    End If;

    Update ������Ϣ
    Set �Ա� = Decode(����, '�²���', Nvl(�Ա�_In, �Ա�), �Ա�), ���� = Decode(����, '�²���', Nvl(����_In, ����), ����),
        ���� = Decode(����, '�²���', ����_In, ����), ҽ�Ƹ��ʽ = Nvl(v_���ʽ, ҽ�Ƹ��ʽ), �ѱ� = Decode(v_�ѱ�����, 1, �ѱ�_In, �ѱ�)
    Where ����id = ����id_In;

    Select Zl_To_Number(Nvl(zl_GetSysParameter('�Զ���������', '1111'), '0')) Into n_�²���ģʽ From Dual;
    If n_�²���ģʽ = 1 Then
      Update ���˹Һż�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ����id = ����id_In And ���� = '�²���';
      Update ������ü�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ����id = ����id_In And ���� = '�²���';
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ﻮ�ۼ�¼_Insert;
/

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Create Or Replace Procedure Zl_������ʼ�¼_Delete
(
  No_In           ������ü�¼.No%Type,
  ���_In         Varchar2,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ��Һ��ҩ���_In Number := 1,
  �Ǽ�ʱ��_In     סԺ���ü�¼.�Ǽ�ʱ��%Type := Sysdate
) As
  --���ܣ�����һ��������ʵ�����ָ������� 
  --��ţ���ʽ��"1,3,5,7,8",��"1:2:33456,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,�м�����ֱ�ʾ�˵�����,��������ֱ�ʾ��ҩ��¼��ID,Ŀǰ�����������ʱ�Ŵ��� 
  --      Ϊ�ձ�ʾ�������пɳ����� 

  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill(n_��־ Number) Is
    Select a.Id, a.�۸񸸺�, a.���, a.ִ��״̬, a.�շ����, a.ҽ�����, a.����id, a.��ҳid, a.������Ŀid, a.��������id, a.ִ�в���id, a.���˲���id, a.���˿���id,
           a.ʵ�ս��, Decode(a.��¼״̬, 0, 1, 0) As ����, j.�������, m.��������
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.�շ�ϸĿid + 0 = m.����id(+) And a.No = No_In And a.��¼���� = 2 And a.��¼״̬ In (0, 1, 3) And
          a.�����־ = n_��־
    Order By a.�շ�ϸĿid, a.���;

  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ���, �۸񸸺� From ������ü�¼ Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) Order By ���;
  l_���� t_Numlist := t_Numlist();

  v_ҽ��ids  Varchar2(4000);
  n_����     ������ü�¼.�۸񸸺�%Type;
  n_�����־ ������ü�¼.�����־%Type;

  --�����˷Ѽ������
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;

  n_׼������ Number;
  n_�˷Ѵ��� Number;
  n_�˷����� Number;
  n_�������� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;

  v_���   Varchar2(4000);
  v_��ҩid Varchar2(4000);
  v_Tmp    Varchar2(4000);

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --�������ʱ,��ҩƷ�ᴫ���кŵ��������� 
  If Not ���_In Is Null Then
    If Instr(���_In, ':') > 0 Then
      --��ʽ��1:2:33456,3:2,5:2,7:2,8:2
      For c_��� In (Select C1, C2 From Table(f_Str2list2(���_In, ',', ':'))) Loop
        v_��� := v_��� || ',' || c_���.C1;
        If Instr(c_���.C2, ':') > 0 Then
          v_��ҩid := v_��ҩid || ',' || Substr(c_���.C2, Instr(c_���.C2, ':') + 1);
        End If;
      End Loop;
      v_���   := Substr(v_���, 2);
      v_��ҩid := Substr(v_��ҩid, 2);
    Else
      v_��� := ���_In;
    End If;
  End If;

  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(1), 0), Max(Nvl(�����־, 1))
  Into n_Count, n_�����־
  From ������ü�¼
  Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  If Nvl(n_�����־, 0) = 0 Then
    n_�����־ := 1;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(1), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --���ñ���
  Select Nvl(�Ǽ�ʱ��_In, Sysdate) Into d_Curdate From Dual;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --ѭ������ÿ�з���(������Ŀ��)
  For r_Bill In c_Bill(n_�����־) Loop
    If Instr(',' || v_��� || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or v_��� Is Null Then
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
        From ������ü�¼
        Where NO = No_In And ��¼���� = 2 And ��� = r_Bill.���;
      
        n_�������� := 0;
        n_�˷����� := 0;
        If n_ʣ������ = 0 Then
          If v_��� Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ�����ʣ�';
            Raise Err_Item;
          End If;
          --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
        Else
          If Instr(���_In, ':') > 0 Then
            Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(���_In, ',', ':')) A Where a.C1 = r_Bill.���;
            If Instr(v_Tmp, ':') > 0 Then
              n_�˷����� := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            Else
              n_�˷����� := To_Number(v_Tmp);
            End If;
            n_�������� := 1;
          End If;
        
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            --@@@
            --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
            --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ)
            --: 2.���ڲ���ҽ�ԼƼ��е��շѷ�ʽΪ:0-������ȡ ��,��֧�ֲ�����;�����������,��ֻ��ȫ��
            --: 3.������ҽ����,����ʣ������Ϊ׼
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
              Select Nvl(Sum(����), 0), Count(*)
              Into n_׼������, n_Count
              From (Select j.ҽ����� As ҽ��id, j.�շ�ϸĿid, Nvl(j.����, 1) * Nvl(j.����, 1) As ����
                     From ������ü�¼ J, ����ҽ����¼ M
                     Where j.ҽ����� = m.Id And j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                           Exists
                      (Select 1
                            From ����ҽ������ A
                            Where a.ҽ��id = j.ҽ����� And Nvl(a.ִ��״̬, 0) <> 1 And a.No || '' = No_In) And Exists
                      (Select 1
                            From ����ҽ���Ƽ� A
                            Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And j.�۸񸸺� Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                           (j.��¼״̬ In (1, 3) And Not Exists
                            (Select 1
                             From ҩƷ�շ���¼
                             Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) Or
                            j.��¼״̬ = 2 And Not Exists
                            (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = j.�շ�ϸĿid))
                     Union All
                     Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                     From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And Nvl(a.�շѷ�ʽ, 0) = 0 And b.���ͺ� = c.���ͺ� And a.ҽ��id = m.Id And
                           Nvl(c.ִ�н��, 1) = 1 And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And
                           j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Not Exists
                      (Select 1
                            From ҩƷ�շ���¼
                            Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) And Not Exists
                      (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                     Union All
                     Select a.ҽ��id, a.�շ�ϸĿid, 0 As ����
                     From ����ҽ���Ƽ� A, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = m.Id And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) <> 0 And
                           j.No = No_In And j.��¼���� = 2 And Nvl(j.ִ��״̬, 0) = 2 And Not Exists
                      (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1) And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0);
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_׼������ := n_ʣ������;
            End If;
          Else
            Select Sum(Nvl(����, 1) * ʵ������)
            Into n_׼������
            From ҩƷ�շ���¼
            Where NO = No_In And ���� In (9, 25) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = r_Bill.Id;
          
            --���������õ���������
            If r_Bill.�շ���� = '4' And Nvl(n_׼������, 0) = 0 Then
              n_׼������ := n_ʣ������;
            End If;
          End If;
        
          If Nvl(n_�˷�����, 0) = 0 Then
            n_�˷����� := n_׼������;
          Else
            If n_׼������ < n_�˷����� Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з���׼���������㱾������������';
              Raise Err_Item;
            End If;
          End If;
        
          --���=ʣ����*(׼����/ʣ����)
          n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_�˷����� / n_ʣ������), n_Dec);
          n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_�˷����� / n_ʣ������), n_Dec);
          n_ͳ���� := Round(n_ʣ��ͳ�� * (n_�˷����� / n_ʣ������), n_Dec);
        
          If Nvl(r_Bill.����, 0) = 0 Then
            --�ñ���Ŀ�ڼ�������
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into n_�˷Ѵ���
            From ������ü�¼
            Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 2 And ��� = r_Bill.���;
          
            --�����˷Ѽ�¼
            Insert Into ������ü�¼
              (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������,
               ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����, �Һ�id, ��ҳid,
               ���˲���id)
              Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                     ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_�˷����� - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                     Decode(Sign(n_�˷����� - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_�˷�����), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                     ��׼����, -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, ִ��ʱ��, ����Ա���_In,
                     ����Ա����_In, ����ʱ��, d_Curdate, ������Ŀ��, ���մ���id, -1 * n_ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����, �Һ�id, ��ҳid,
                     ���˲���id
              From ������ü�¼
              Where ID = r_Bill.Id;
          
            --�������
            If n_�����־ <> 4 Then
              Update �������
              Set ������� = Nvl(�������, 0) - n_ʵ�ս��
              Where ����id = r_Bill.����id And ���� = 1 And ���� = 1;
              If Sql%RowCount = 0 Then
                Insert Into �������
                  (����id, ����, ����, �������, Ԥ�����)
                Values
                  (r_Bill.����id, 1, 1, -1 * n_ʵ�ս��, 0);
              End If;
            End If;
          
            --����δ�����
            Update ����δ�����
            Set ��� = Nvl(���, 0) - n_ʵ�ս��
            Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = Nvl(r_Bill.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_Bill.���˲���id, 0) And
                  Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
                  Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = n_�����־;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_Bill.����id, r_Bill.��ҳid, r_Bill.���˲���id, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid,
                 n_�����־, -1 * n_ʵ�ս��);
            End If;
          
            --���ԭ���ü�¼
            --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1
            Update ������ü�¼
            Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(n_�˷����� - n_ʣ������), 0, 0, 1)
            Where ID = r_Bill.Id;
          Else
            --���ۼ��˵�
            If Nvl(n_��������, 0) = 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
            Else
              --�������� 
              --���۵�,�Ƚ���ص����ݴ������ڲ�����
              Update סԺ���ü�¼
              Set ���� = 1, ���� = Nvl(����, 1) * ���� - n_�˷�����, Ӧ�ս�� = Nvl(Ӧ�ս��, 0) - n_Ӧ�ս��, ʵ�ս�� = Nvl(ʵ�ս��, 0) - n_ʵ�ս��,
                  �Ǽ�ʱ�� = d_Curdate, ͳ���� = Nvl(ͳ����, 0) - n_ͳ����
              Where ID = r_Bill.Id
              Returning ���� Into n_ʣ������;
              If Nvl(n_ʣ������, 0) <= 0 Then
                l_����.Extend;
                l_����(l_����.Count) := r_Bill.Id;
              End If;
            End If;
          
            If r_Bill.ҽ����� Is Not Null Then
              If Instr(',' || Nvl(v_ҽ��ids, '') || ',', ',' || r_Bill.ҽ����� || ',') = 0 Then
                v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || r_Bill.ҽ�����;
              End If;
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

  --��������ҩID,����ҩƷ�Ƿ�����Һ��ҩ���� 
  If v_��ҩid Is Null And ��Һ��ҩ���_In = 1 Then
    For v_���� In (Select ID
                 From ������ü�¼
                 Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = n_�����־ And
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
  --���밴�ա��շ�ϸĿid���������򣬷�ֹ��������ҩƷ��桱��
  For v_���� In (Select ID, ���
               From ������ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = n_�����־ And
                     (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)
               Order By �շ�ϸĿid) Loop
    --���ݷ���ID��������صĴ��� 
    n_�˷����� := 0;
    If Instr(���_In, ':') > 0 Then
      Select Max(a.C2) Into v_Tmp From Table(f_Str2list2(���_In, ',', ':')) A Where a.C1 = v_����.���;
      If Instr(v_Tmp, ':') > 0 Then
        n_�˷����� := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      Else
        n_�˷����� := To_Number(v_Tmp);
      End If;
    End If;
    Zl_ҩƷ�շ���¼_�����˷�(v_����.Id, n_�˷�����, v_��ҩid);
  End Loop;

  --ɾ�����ۼ�¼
  n_Count := l_����.Count;
  Forall I In 1 .. l_����.Count
    Delete From ������ü�¼ Where ID = l_����(I);

  --ɾ��֮����ͳһ�������
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        n_���� := n_Count;
      End If;
    
      Update ������ü�¼
      Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, n_����)
      Where NO = No_In And ��¼���� = 2 And ��� = r_Serial.���;
    
      Update ������ü�¼ Set �������� = n_Count Where NO = No_In And ��¼���� = 2 And �������� = r_Serial.���;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;

  --���ŵ���ȫ������ʱ��ɾ������ҽ������
  For c_ҽ�� In (Select Distinct ҽ�����
               From ������ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
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
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 0, No_In, v_ҽ��ids);
  Else
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 2, No_In, v_ҽ��ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ʼ�¼_Delete;
/

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Create Or Replace Procedure Zl_������ʼ�¼_Insert
(
  No_In         ������ü�¼.No%Type,
  ���_In       ������ü�¼.���%Type,
  ����id_In     ������ü�¼.����id%Type,
  ��ʶ��_In     ������ü�¼.��ʶ��%Type,
  ����_In       ������ü�¼.����%Type,
  �Ա�_In       ������ü�¼.�Ա�%Type,
  ����_In       ������ü�¼.����%Type,
  �ѱ�_In       ������ü�¼.�ѱ�%Type,
  �Ӱ��־_In   ������ü�¼.�Ӱ��־%Type,
  Ӥ����_In     ������ü�¼.Ӥ����%Type,
  ���˿���id_In ������ü�¼.���˿���id%Type,
  ��������id_In ������ü�¼.��������id%Type,
  ������_In     ������ü�¼.������%Type,
  ��������_In   ������ü�¼.��������%Type,
  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
  �շ����_In   ������ü�¼.�շ����%Type,
  ���㵥λ_In   ������ü�¼.���㵥λ%Type,
  ����_In       ������ü�¼.����%Type,
  ����_In       ������ü�¼.����%Type,
  ���ӱ�־_In   ������ü�¼.���ӱ�־%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
  ������Ŀid_In ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
  ��׼����_In   ������ü�¼.��׼����%Type,
  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
  ҩƷժҪ_In   ҩƷ�շ���¼.ժҪ%Type,
  ����_In       Number,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ���ʵ�id_In   ������ü�¼.���ʵ�id%Type := Null,
  ����ժҪ_In   ������ü�¼.ժҪ%Type := Null,
  ҽ�����_In   ������ü�¼.ҽ�����%Type := Null,
  Ƶ��_In       ҩƷ�շ���¼.Ƶ��%Type := Null,
  ����_In       ҩƷ�շ���¼.����%Type := Null,
  �÷�_In       ҩƷ�շ���¼.�÷�%Type := Null, --�÷�[|�巨]
  ��Ч_In       ҩƷ�շ���¼.����%Type := Null,
  �Ƽ�����_In   ҩƷ�շ���¼.����%Type := Null,
  �����־_In   ������ü�¼.�����־%Type := 1,
  ��ҩ��̬_In   ������ü�¼.����%Type := Null,
  ��������_In   Number := 0,
  ����_In       ҩƷ�շ���¼.����%Type := Null,
  ��ҳid_In     ������ü�¼.��ҳid%Type := Null,
  ���˲���id_In ������ü�¼.���˲���id%Type := Null
) As
  --���ܣ�����һ��������ʵ���
  --������
  --   ҩƷժҪ_IN:�޸ı����µ���ʱ�á�Ŀǰ�����ڴ����ҩƷ�շ���¼��ժҪ�С�
  --         ԭ����(��¼״̬=2)��¼�޸Ĳ������µ��ݺš�
  --         �µ���(��¼״̬=1)��¼���޸ĵ�ԭ���ݺš�
  v_����id ������ü�¼.Id%Type;
  n_����   ���˹Һż�¼.����%Type;
  n_��ҳid ������ü�¼.��ҳid%Type;

  --��ʱ����
  v_�÷�     ҩƷ�շ���¼.�÷�%Type;
  v_�巨     ҩƷ�շ���¼.���%Type;
  n_����С�� Number;
  n_�Һ�id   ���˹Һż�¼.Id%Type;

  n_Dec     Number;
  n_Count   Number;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_��ҩ���� ҩƷ�շ���¼.��ҩ����%Type;
  n_�������� ��������.��������%Type;

Begin
  n_�������� := 0;
  If �շ����_In = '4' Then
    --�������õ����ĲŴ���
    Select Nvl(��������, 0) Into n_�������� From �������� Where ����id = �շ�ϸĿid_In;
  End If;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
  Into n_Dec, n_����С��
  From Dual;

  n_��ҳid := ��ҳid_In;
  If Nvl(n_��ҳid, 0) = 0 Then
    Select Max(��ҳid) Into n_��ҳid From ������ҳ Where ����id = ����id_In And �������� = 1 And ��Ժ���� Is Null;
  End If;

  If (�շ����_In In ('5', '6', '7') Or �շ����_In = '4' And n_�������� = 1) And Nvl(����_In, 0) = 0 Then
    --ͬһ�ŵ���,����ͬһҩ��ͬһ����
    Begin
      Select ��ҩ����
      Into v_��ҩ����
      From ������ü�¼
      Where �շ���� In ('5', '6', '7', '4') And NO = No_In And ��¼���� = 2 And ִ�в���id = ִ�в���id_In And ��ҩ���� Is Not Null And
            Rownum <= 1;
    Exception
      When Others Then
        v_��ҩ���� := Null;
    End;
    If v_��ҩ���� Is Null Then
      --ͬһ��������ͨ�ŹҺ���Ч�Һ���������δ��ҩ�����ϰ��,�����һ�μ��˴���Ϊ׼
      n_Count := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
      If n_Count = 0 Then
        n_Count := 1;
      End If;
    
      Begin
        Select ��ҩ����
        Into v_��ҩ����
        From (Select �Ǽ�ʱ��, ��ҩ����
               From ������ü�¼ A
               Where �շ���� In ('5', '6', '7', '4') And ����id = ����id_In And �Ǽ�ʱ�� Between Sysdate - n_Count And Sysdate And
                     ��¼���� = 2 And ִ�в���id = ִ�в���id_In And ��ҩ���� Is Not Null And Exists
                (Select 1
                      From δ��ҩƷ��¼
                      Where a.No = NO And ���� In (9, 26) And �ⷿid + 0 = ִ�в���id_In And ����id + 0 = ����id_In) And Exists
                (Select 1
                      From ��ҩ����
                      Where Nvl(�ϰ��, 0) = 1 And ���� = a.��ҩ���� And Nvl(ר��, 0) = 0 And ҩ��id = ִ�в���id_In)
               Order By �Ǽ�ʱ�� Desc)
        Where Rownum <= 1;
      
      Exception
        When Others Then
          v_��ҩ���� := Null;
      End;
      If v_��ҩ���� Is Null Then
        v_��ҩ���� := Zl_Get��ҩ����(ִ�в���id_In);
      End If;
    End If;
  End If;
  --������ü�¼
  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;

  --�Ƿ��Ǽ���Һŵ�
  If Nvl(ҽ�����_In, 0) <> 0 Then
    Begin
      Select Nvl(Max(����), 0), Max(ID)
      Into n_����, n_�Һ�id
      From ���˹Һż�¼
      Where NO In (Select �Һŵ� From ����ҽ����¼ Where ID = Nvl(ҽ�����_In, 0)) And ����id = ����id_In;
    Exception
      When Others Then
        n_����   := Null;
        n_�Һ�id := Null;
    End;
  End If;

  Insert Into ������ü�¼
    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����, �Ӱ��־,
     ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ����Ա���, ����Ա����, Ӥ����, ���ʵ�id,
     ժҪ, ҽ�����, ����, ��ҩ����, �Ƿ���, ��ҳid, �Һ�id, ���˲���id)
  Values
    (v_����id, 2, No_In, Decode(����_In, 1, 0, 1), ���_In, Decode(��������_In, 0, Null, ��������_In),
     Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), �����־_In, ����id_In, Decode(��ʶ��_In, 0, Null, ��ʶ��_In), ����_In, �Ա�_In, ����_In,
     ���˿���id_In, �ѱ�_In, �շ����_In, �շ�ϸĿid_In, ���㵥λ_In, ����_In, ����_In, �Ӱ��־_In, ���ӱ�־_In, ������Ŀid_In, �վݷ�Ŀ_In, ��׼����_In, Ӧ�ս��_In,
     ʵ�ս��_In, 1, ����Ա����_In, ��������id_In, ������_In, ����ʱ��_In, �Ǽ�ʱ��_In, ִ�в���id_In, 0, Decode(����_In, 1, Null, ����Ա���_In),
     Decode(����_In, 1, Null, ����Ա����_In), Ӥ����_In, ���ʵ�id_In, ����ժҪ_In, ҽ�����_In, ��ҩ��̬_In, v_��ҩ����, Nvl(n_����, 0), n_��ҳid, n_�Һ�id,
     Decode(���˲���id_In, 0, Null, ���˲���id_In));

  --��ػ��ܱ�Ĵ���
  If Nvl(����_In, 0) = 0 Then
    --�������
    If Nvl(�����־_In, 0) <> 4 Then
      Update �������
      Set ������� = Nvl(�������, 0) + ʵ�ս��_In
      Where ����id = ����id_In And ���� = 1 And ���� = Decode(�����־_In, 2, 2, 1);
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, Decode(�����־_In, 2, 2, 1), ʵ�ս��_In, 0);
      End If;
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + ʵ�ս��_In
    Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(n_��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(���˲���id_In, 0) And
          Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And Nvl(��������id, 0) = Nvl(��������id_In, 0) And
          Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And ��Դ;�� + 0 = �����־_In;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, n_��ҳid, ���˲���id_In, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, �����־_In, ʵ�ս��_In);
    End If;
  
  End If;

  --ҩƷ���������ϲ���
  If �շ����_In In ('4', '5', '6', '7') Then
    --ҩƷ�÷��巨�ֽ�
    If �÷�_In Is Not Null Then
      If Instr(�÷�_In, '|') > 0 Then
        v_�÷� := Substr(�÷�_In, 1, Instr(�÷�_In, '|') - 1);
        v_�巨 := Substr(�÷�_In, Instr(�÷�_In, '|') + 1);
      Else
        v_�÷� := �÷�_In;
      End If;
    End If;
    Zl_ҩƷ�շ���¼_���۳���(v_����id, ҩƷժҪ_In, Ƶ��_In, ����_In, v_�÷�, v_�巨, ��Ч_In, �Ƽ�����_In, n_��ҳid, ��������_In, ����_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ʼ�¼_Insert;
/

--139063:Ƚ����,2019-04-01,�������۲��˰��������̾���
Create Or Replace Function Zl1_Getdef_Prepaymoney
(
  ����id_In ������Ϣ.����id%Type,
  ��ҳid_In ������ҳ.��ҳid%Type,
  ��Դ_In   Number := 2
  --���ܣ���ȡĬ�ϵ�Ԥ����ɿ�� 
  --     ��������Ҫ�����˷��ò�ѯʱ���ý�Ԥ��ʱ����,��Ҫ�Ƕ�ȡȱʡ��Ԥ����� 
  --     �û����Ը���ʵ�ʲ����Ĺ���,���������ȱʡֵ 
  --������ 
  --    ����ID_In������ID 
  --    ��ҳID_In:��ҳID 
  --    ��Դ_In:1-���2-סԺ
) Return Number Is
  Err_Custom Exception;
  n_Ԥ����� �������.Ԥ�����%Type;
  n_������� �������.�������%Type;
  n_Ԥ����� �������.Ԥ�����%Type;
  n_����Ԥ�� �������.Ԥ�����%Type;
Begin
  --Ŀǰ������:���ܷ���-Ԥ�����ܶ�-������� >0�� 
  Select Nvl(Sum(Ԥ�����), 0), Nvl(Sum(�������), 0), Nvl(Sum(Ԥ�����), 0)
  Into n_Ԥ�����, n_�������, n_Ԥ�����
  From (Select Nvl(Ԥ�����, 0) Ԥ�����, Nvl(�������, 0) �������, 0 As Ԥ�����
         From �������
         Where ���� = 1 And ����id = ����id_In And Nvl(����, 2) = Nvl(��Դ_In, 2)
         Union All
         Select 0 As Ԥ�����, 0 As �������, Sum(b.���) As Ԥ�����
         From ������Ϣ A, ����ģ����� B
         Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id = ����id_In);
  n_����Ԥ�� := n_Ԥ����� - n_������� + n_Ԥ�����;
  If n_����Ԥ�� > 0 Then
    n_����Ԥ�� := 0;
  End If;
  n_����Ԥ�� := Abs(n_����Ԥ��);
  Return n_����Ԥ��;

End Zl1_Getdef_Prepaymoney;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0056' Where ���=&n_System;
Commit;
