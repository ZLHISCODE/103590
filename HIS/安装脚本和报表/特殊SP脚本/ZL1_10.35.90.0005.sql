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
--117772:��͢��,2018-04-02,����ϵͳ������Ⱦ�����濨ǿ����д
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, -null, -null, 300, '��Ⱦ�����濨ǿ����д', '0',
         '0', '�����ò��������������д��ϵ����Ĵ�Ⱦ�����濨����ʾ�˳���ť�ҵ���ر�X��ťʱ���رմ��塣�������ò����򲻿���',
         '0-��ʾ������,1-��ʾ����', Null, '������ĳЩҽԺ����Ҫ��ҽ��ǿ����д��Ⱦ�����濨', Null
  From Dual;


--123734:������,2018-03-31,�°滤ʿվ���������ſ���Ϣ��ʾ����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1265, 1, 0, 0, 0, 0, 0, 14, '����λ������ʾ��λ״��', 1, NULL,
         '�����°滤ʿ����վ�����没��������Ϣ����λ��Ϣ�Ƿ񰴴�λ������ʾ��λʹ��״��', '0-��ʾռ�ô�λ�����Ϳմ�������1-��ʾÿ�ִ�λ���Ʒ���Ĵ�λ���Ϳմ���', Null, '������Ҫ�鿴��ϸ�Ĵ�λʹ�����', Null
  From Dual;
-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--122954:��ΰ��,2018-04-08,����������ҩ
Create Or Replace Procedure Zl_����������ҩ����_Update
(
  ����id_In In ������Ϣ.����id%Type,
  ��ҳid_In In ������ҳ.��ҳid%Type,
  �Һ�id_In In ���˹Һż�¼.Id%Type := Null
) As

  --------------------------------------------------------------------------------------------------
  --����:������ҩ��⴫��ֵ�����غ�����ҩ����ֵ
  --����:Xml_Return  ���ز����XML��
  -- <details_xml>
  --    <patient_info>
  --      <info name="��������" value="28114.45"/>
  --     <info name="��������" value="����"/>
  --      <info name="�Ա�" value="Ů"/>
  --      <info name="ְҵ" value="�˶�Ա"/>
  --      <info name="����" value="1"/>
  --      <info name="����" value="1"/>
  --      <info name="�ι��ܲ�ȫ" value="1">
  --      <info name="���ظι��ܲ�ȫ" value="1">
  --      <info name="�����ܲ�ȫ" value="1">
  --      <info name="���������ܲ�ȫ" value="1">
  --      <info name="���" value="J18.000"/> --��ϴ����룬�������Զ��ŷָ�
  --    </patient_info>
  --    <medicine_info>
  --      <medicine>
  --        <info name="ҽ��ID" value="1"/>
  --        <info name="��λ��" value="86900967000160" main="46d64420-8319-4768-9a11-f4b0f5e4ce7a"/> --mainֵ�ǹ̶���
  --        <info name="������ĿID" value="67232" main="4e19df1c-c1b9-4a43-a83d-0741a19961ab"/>
  --        <info name="��Һ���" value="1"/>
  --        <info name="������λ" value="ml"/>
  --        <info name="������" value="250"/>
  --        <info name="������-������" value="5.21"/>--������-������= trunc(������/��������,2)
  --        <info name="������-�����" value="170.3"/>--������-�����= trunc(������/(0.0061*�������+0.0128*��������-0.1529),2)
  --        <info name="ÿ����" value="250"/>
  --        1.ÿ����=������*��Ƶ��
  --        2.��Ƶ�μ��㣺
  --            a.���÷�Χ=-1����Ƶ��=1
  --            b.�����λ=�� and Ƶ�ʼ��=1����Ƶ��=Ƶ�ʴ���
  --            c.�����λ=�� and Ƶ�ʼ��>1 and Ƶ�ʴ���=1����Ƶ��=1
  --            d.�����λ=Сʱ and Ƶ�ʼ��<=24,��Ƶ��=24/Ƶ�ʼ��*Ƶ�ʴ���
  --            e.�����λ=Сʱ and Ƶ�ʼ��>24 and Ƶ�ʴ���=1����Ƶ��=1
  --            f.�����λ=�� and Ƶ�ʴ���=1����Ƶ��=1
  --        <info name="ÿ����-������" value="5.21"/>  --trunc(ÿ����/��������,2)
  --        <info name="ÿ����-�����" value="170.3"/>  --ÿ����-�����= trunc(ÿ����/(0.0061*�������+0.0128*��������-0.1529),2)
  --        <info name="��ҩƵ��" value="ÿ��һ��"/>
  --        <info name="��ҩ;��" value="001"/>
  --      </medicine>
  --    </medicine_info>
  --  </details_xml>
  --------------------------------------------------------------------------------------------------
  Xml_Ret             Xmltype;
  Xml_Document        Xmldom.Domdocument;
  Xml_Nodelist        Xmldom.Domnodelist;
  Xml_Domelement      Xmldom.Domelement;
  Xml_Domnamednodemap Xmldom.Domnamednodemap;
  Xml_Node_Med        Xmldom.Domnode;
  Xml_Node            Xmldom.Domnode;
  Xml_Node_New        Xmldom.Domnode;
  ----------------------------------
  n_��� Number(10, 2); --��λ:cm
  n_���� Number(10, 2); --����:KG

  l_Clob    Clob;
  v_Err_Msg Varchar2(2000);
  v_Temp    Varchar2(200);
  v_Value   Varchar2(200);
  n_Nodenum Number(5);
  Err_Item Exception;
Begin
  --��
  --��CLOB������ȡ��v_XML��
  Select �������� Into l_Clob From ����������ҩ����;
  Xml_Ret        := Xmltype(l_Clob); --���溯������ֵ
  Xml_Document   := Xmldom.Newdomdocument(Xml_Ret);
  Xml_Domelement := Xmldom.Getdocumentelement(Xml_Document);
  Xml_Nodelist   := Xmldom.Getelementsbytagname(Xml_Domelement, 'patient_info');
  --��ȡpatient_info/INfo�ڵ�
  Xml_Nodelist := Xmldom.Getchildnodes(Xmldom.Item(Xml_Nodelist, 0));
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Domnamednodemap := Xmldom.Getattributes(Xmldom.Item(Xml_Nodelist, I));
    v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
    If v_Temp = '���' Then
      n_��� := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
    If v_Temp = '����' Then
      n_���� := Nvl(To_Number(Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'))), 0);
    End If;
  End Loop;
  --��ȡmedicine/INfo�ڵ�

  Xml_Nodelist := Xmldom.Getelementsbytagname(Xml_Domelement, 'medicine');
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Node_Med := Xmldom.Item(Xml_Nodelist, I); --ȡ��һ�����ӽڵ�medicine
    Xml_Nodelist := Xmldom.Getchildnodes(Xml_Node_Med); --infos
    Xml_Node     := Xmldom.Getfirstchild(Xml_Node_Med); --ȡ��һ�����ӽڵ�
    While Not Xmldom.Isnull(Xml_Node) Loop
      Xml_Domnamednodemap := Xmldom.Getattributes(Xml_Node);
      v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
      v_Value             := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'));
      If v_Temp = '������' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '������-������');
        If n_���� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(To_Number(v_Value) / n_����, 'FM9999990.09'));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
        --������-�����trunc(ÿ����/(0.0061*�������+0.0128*��������-0.1529),2)
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', '������-�����');
        If n_���� > 0 And n_��� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(To_Number(v_Value) / (0.0061 * n_��� + 0.0128 * n_���� - 0.1529), 'FM9999990.09'));
        
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
    
      If v_Temp = 'ÿ����' Then
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', 'ÿ����-������');
        If n_���� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(To_Number(v_Value) / n_����, 'FM9999990.09'));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      
        --ÿ����-�����
        Xml_Node_New := Xmldom.Appendchild(Xml_Node_Med, Xmldom.Clonenode(Xml_Node, False));
        Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'name', 'ÿ����-�����');
        If n_���� > 0 And n_��� > 0 Then
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value',
                              To_Char(Trunc(To_Number(v_Value) / (0.0061 * n_��� + 0.0128 * n_���� - 0.1529), 2),
                                       'FM9999990.09'));
        Else
          Xmldom.Setattribute(Xmldom.Makeelement(Xml_Node_New), 'value', '');
        End If;
      End If;
      --ȡ��һ���ֵܽڵ�
      Xml_Node := Xmldom.Getnextsibling(Xml_Node);
    End Loop;
  End Loop;

  --����������ֵ������ʱ��,ZLHIS���������ǰ��ȡ����Ϊ���������Ʒ���ֵ���ܳ���4000���ƣ�
  Update ����������ҩ���� Set �������� = Xml_Ret.Getclobval();

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����������ҩ����_Update;
/

--123754:Ƚ����,2018-04-09,ҽ������վԤԼ�Һ���������
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan
(
  �Һ�ʱ��_In In Date := Null,
  ��Դid_In   �ٴ������Դ.Id%Type := Null
) As
  -------------------------------------------------------------------------
  --����˵�����Զ������ٴ������¼
  --          1�����ݺ�Դ�Զ�����ԤԼ���ڵ��ٴ������¼;
  --          2��ԤԼ������ȷ��:��ԴԤԼ����-->ԤԼ��ʽ��������ȡ���)-->ϵͳԤԼ����
  --���:�Һ�ʱ��_IN:NULLʱ���Զ�����;����ֻ���ָ�������Ƿ������˳����¼û��
  --    ��Դid_In:NULLʱ�������к�Դ������ֻ����ָ����Դ
  -------------------------------------------------------------------------
  n_ȱʡԤԼ���� �ٴ������Դ.ԤԼ����%Type;
  v_����Ա����   �ٴ����ﰲ��.����Ա����%Type;
  d_�Ǽ�����     �ٴ����ﰲ��.�Ǽ�ʱ��%Type;
  n_����id       �ٴ����ﰲ��.Id%Type;
  n_��Ŀid       �ٴ����ﰲ��.��Ŀid %Type;

  n_��¼id   �ٴ������¼.Id%Type;
  d_��ǰ���� �ٴ������¼.��������%Type;

  l_�̶�ʱ�� t_Strlist := t_Strlist();
  n_Count    Number(18);

  n_��ԤԼ���� Number := 0;
  d_��ʼʱ��   �ٴ������¼.��ʼʱ��%Type;
Begin

  Select Max(ԤԼ����) Into n_ȱʡԤԼ���� From ԤԼ��ʽ;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := To_Number(Nvl(zl_GetSysParameter('�Һ�����ԤԼ����'), '0'));
  End If;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := 7;
  End If;

  --�԰���Ϊ��λ,�����������Դ����ʱ�䡱��12:00:00-23:59:59�ڼ�ģ��򿪷�ԤԼ����+1��
  n_��ԤԼ���� := Zl_Fun_Getappointmentdays;

  d_��ǰ����   := Trunc(Nvl(�Һ�ʱ��_In, Sysdate));
  d_�Ǽ�����   := Sysdate;
  v_����Ա���� := Zl_Username;

  --��һ��ѭ������Դ��Ϣ
  For c_��Դ In (Select c.Id, c.����, c.����, c.��Ŀid, c.����id, c.ҽ������,
                      Decode(Nvl(c.ԤԼ����, 0), 0, n_ȱʡԤԼ����, c.ԤԼ����) + n_��ԤԼ���� As ԤԼ����, Nvl(b.վ��, '-') As վ��,
                      Nvl(c.�Ƿ���ջ���, 0) As �Ƿ���ջ���, Nvl(c.���տ���״̬, 0) As ���տ���״̬, Nvl(c.�Ű෽ʽ, 0) As �Ű෽ʽ
               From �ٴ������Դ C, ���ű� B, ��Ա�� A, �շ���ĿĿ¼ D
               Where c.����id = b.Id And c.ҽ��id = a.Id(+) And c.��Ŀid = d.Id And Nvl(c.�Ƿ�ɾ��, 0) = 0 And
                     Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (��Դid_In Is Null Or c.Id = ��Դid_In)
                    --
                     And Exists (Select 1
                      From �ٴ����ﰲ�� M, �ٴ������ N
                      Where m.����id = n.Id And m.��Դid = c.Id And Nvl(n.�Ű෽ʽ, 0) = 0 And n.����ʱ�� Is Not Null And
                            m.���ʱ�� Is Not Null And d_��ǰ���� <= m.��ֹʱ��)) Loop
  
    --��鵱ǰ�������ڵİ��ŵ��շ���Ŀ�Ƿ�Ϊ��Դ�е��շ���Ŀ��������ǣ�����º�Դ�е��շ���Ŀ
    Begin
      Select ��Ŀid
      Into n_��Ŀid
      From (Select a.��Ŀid
             From �ٴ����ﰲ�� A, �ٴ������ B
             Where a.����id = b.Id And a.��Դid = c_��Դ.Id And a.���ʱ�� Is Not Null And d_��ǰ���� Between a.��ʼʱ�� And a.��ֹʱ�� And
                   Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null
             Order By a.�Ǽ�ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_��Ŀid := Null;
    End;
    If Nvl(n_��Ŀid, 0) <> 0 Then
      If Nvl(c_��Դ.��Ŀid, 0) <> n_��Ŀid Then
        Update �ٴ������Դ Set ��Ŀid = n_��Ŀid Where ID = c_��Դ.Id;
        Commit;
      End If;
    End If;
  
    --�ڶ���ѭ������������
    --��ͷһ�쿪ʼ���ɣ�������ȫ��(8:00-7:59)��0:00-7:59û�г����¼
    --1.δָ����ԴID�������������ɳ����¼���г����¼�����ڽ����ٴ���
    --2.ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼
    For c_���� In (Select m.����,
                        Decode(To_Char(m.����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                                '����', Null) As ����
                 From (Select Trunc(d_��ǰ����) + ���� As ����
                        From (Select Level - 1 As ���� From Dual Connect By Level <= c_��Դ.ԤԼ���� + 1)
                        Where ��Դid_In Is Not Null
                        Union All
                        Select Trunc(d_��ǰ���� - 1) + ���� As ����
                        From (Select Level - 1 As ���� From Dual Connect By Level <= c_��Դ.ԤԼ���� + 2)
                        Where ��Դid_In Is Null And Not Exists
                         (Select 1
                               From �ٴ������¼ A
                               Where a.��Դid = c_��Դ.Id And a.�������� = Trunc(d_��ǰ���� - 1) + ����)) M
                 Where �Һ�ʱ��_In Is Null Or Trunc(�Һ�ʱ��_In) = m.����) Loop
    
      l_�̶�ʱ�� := t_Strlist();
      --��鵱���Ƿ�����/�ܳ������,���ڣ������ɳ����¼
      Select Count(1)
      Into n_Count
      From �ٴ����ﰲ�� A, �ٴ������ B
      Where a.����id = b.Id And a.��Դid = c_��Դ.Id And c_����.���� Between Trunc(a.��ʼʱ��) And Trunc(a.��ֹʱ��) And
            Nvl(b.�Ű෽ʽ, 0) In (1, 2) And Rownum < 2;
    
      --��ǰ��ԴΪ����/���Ű࣬�ҵ�ǰ����֮ǰ���а���/���Ű�ĳ����¼�Ͳ��ٰ��̶��������ɳ����¼��
      If n_Count = 0 And Nvl(c_��Դ.�Ű෽ʽ, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From �ٴ����ﰲ�� A, �ٴ������ B
        Where a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) In (1, 2) And a.��Դid = c_��Դ.Id And a.��ʼʱ�� < c_����.���� And Rownum < 2;
      End If;
    
      If n_Count = 0 Then
        If ��Դid_In Is Null Then
          --���ﰲ��,ȡ���Ǽǵ�һ��
          Begin
            Select ����id
            Into n_����id
            From (Select a.Id As ����id
                   From �ٴ����ﰲ�� A, �ٴ������ B
                   Where a.��Դid = c_��Դ.Id And a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null And
                         a.���ʱ�� Is Not Null And c_����.���� Between a.��ʼʱ�� And a.��ֹʱ��
                   Order By a.�Ǽ�ʱ�� Desc)
            Where Rownum < 2;
          Exception
            When Others Then
              n_����id := 0;
          End;
        Else
          --���ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼�����Ǽǵ�һ���϶��Ǳ��������ģ�
          --ֻ��Ҫ����������ż��ɣ��������������Чʱ�䷶Χ�ڵľͲ�����
          Begin
            Select ����id
            Into n_����id
            From (Select a.Id As ����id, a.��ʼʱ��, a.��ֹʱ��, Row_Number() Over(Order By a.�Ǽ�ʱ�� Desc) As �к�
                   From �ٴ����ﰲ�� A, �ٴ������ B
                   Where a.��Դid = c_��Դ.Id And a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null And
                         a.���ʱ�� Is Not Null And c_����.���� Between ��ʼʱ�� And ��ֹʱ��)
            Where �к� = 1;
          Exception
            When Others Then
              n_����id := 0;
          End;
        End If;
      
        If Nvl(n_����id, 0) <> 0 Then
          If ��Դid_In Is Not Null Then
            --2.ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼
            --�����г����¼����Ҫ�����´���
            For c_��¼ In (Select a.����id, a.Id As ��¼id, a.��������, a.�ϰ�ʱ��, a.�Ƿ��ʱ��, a.�Ƿ���ſ���
                         From �ٴ������¼ A
                         Where a.��Դid = c_��Դ.Id And a.�������� = c_����.����) Loop
            
              Select Count(1) Into n_Count From ���˹Һż�¼ Where �����¼id = c_��¼.��¼id;
              If n_Count = 0 Then
                --2.2.1���ʱ�β�����ԤԼ�Һ����ݣ���ɾ����������
                Zl_�ٴ������ϰ�ʱ��_Delete(c_��¼.����id, To_Char(c_��¼.��������, 'yyyy-mm-dd'), 1, c_��¼.�ϰ�ʱ��);
              Else
                --2.2.2���ʱ�δ���ԤԼ�Һ����ݣ���ֻ����������¼�İ���ID����
                Update �ٴ������¼ Set ����id = n_����id Where ID = c_��¼.��¼id;
                l_�̶�ʱ��.Extend();
                l_�̶�ʱ��(l_�̶�ʱ��.Count) := c_��¼.�ϰ�ʱ��;
              End If;
            End Loop;
          End If;
        
          --��������Ƿ����
          Select Count(1) Into n_Count From �ٴ��������� Where ����id = n_����id And ������Ŀ = c_����.����;
          If n_Count = 0 Then
            --����������ٴ������¼���������ٴ������¼(ʱ���ΪNULL �Ŀռ�¼)
            Insert Into �ٴ������¼
              (ID, ����id, ��Դid, ��������, �Ǽ���, �Ǽ�ʱ��)
              Select �ٴ������¼_Id.Nextval, n_����id, a.Id As ID, c_����.����, v_����Ա����, d_�Ǽ����� As �Ǽ�ʱ��
              From �ٴ������Դ A, �ٴ����ﰲ�� B
              Where a.Id = b.��Դid And b.Id = n_����id And Not Exists
               (Select 1 From �ٴ������¼ Where ��Դid = a.Id And �������� = c_����.����);
          Else
            For c_��¼ In (With c_ʱ��� As
                            (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��
                            From (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��,
                                          Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                                   From ʱ���
                                   Where Nvl(վ��, c_��Դ.վ��) = c_��Դ.վ�� And Nvl(����, c_��Դ.����) = c_��Դ.����)
                            Where ��� = 1)
                           Select n_����id As ����id, B1.��Դid, c_����.���� As ��������, m.�ϰ�ʱ��, m.Id As ����id,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                           'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ֹʱ��, 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.��ֹʱ�� <= j.��ʼʱ�� Then
                                     1
                                    Else
                                     0
                                  End As ��ֹʱ��, Null As ͣ�￪ʼʱ��, Null As ͣ����ֹʱ��, Null As ͣ��ԭ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.ȱʡʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.ȱʡʱ�� < j.��ʼʱ�� Then
                                     1
                                    Else
                                     0
                                  End As ȱʡԤԼʱ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.��ǰʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.��ʼʱ�� < j.��ǰʱ�� Then
                                     -1
                                    Else
                                     0
                                  End As ��ǰ�Һ�ʱ��, m.�޺���, 0 As �ѹ���, m.��Լ��, 0 As ��Լ��, 0 As �����ѽ���, m.�Ƿ���ſ���, m.�Ƿ��ʱ��, m.ԤԼ����,
                                  m.�Ƿ��ռ, B1.��Ŀid, B1.ҽ��id, B1.ҽ������, Null As ����ҽ��id, Null As ����ҽ������, m.���﷽ʽ, m.����id,
                                  0 As �Ƿ�����, 0 As �Ƿ���ʱ����, v_����Ա���� As ����Ա����, d_�Ǽ����� As �Ǽ�ʱ��, c_����.���� As ������Ŀ
                           From �ٴ����ﰲ�� B1, �ٴ��������� M, c_ʱ��� J
                           Where B1.Id = n_����id And B1.Id = m.����id And m.������Ŀ = c_����.���� And m.�ϰ�ʱ�� = j.ʱ��� And
                                 To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                         'yyyy-mm-dd hh24:mi:ss') >= B1.��ʼʱ�� And Not Exists
                            (Select 1 From Table(l_�̶�ʱ��) Where Column_Value = m.�ϰ�ʱ��)) Loop
            
              Select �ٴ������¼_Id.Nextval Into n_��¼id From Dual;
              Insert Into �ٴ������¼
                (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ͣ�￪ʼʱ��, ͣ����ֹʱ��, ͣ��ԭ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, �ѹ���, ��Լ��, ��Լ��,
                 �����ѽ���, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���﷽ʽ, ����id, �Ƿ�����, �Ƿ���ʱ����, �Ǽ���,
                 �Ǽ�ʱ��, �Ƿ񷢲�)
              Values
                (n_��¼id, c_��¼.����id, c_��¼.��Դid, c_��¼.��������, c_��¼.�ϰ�ʱ��, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, c_��¼.ͣ�￪ʼʱ��, c_��¼.ͣ����ֹʱ��,
                 c_��¼.ͣ��ԭ��, c_��¼.ȱʡԤԼʱ��, c_��¼.��ǰ�Һ�ʱ��, c_��¼.�޺���, c_��¼.�ѹ���, c_��¼.��Լ��, c_��¼.��Լ��, c_��¼.�����ѽ���, c_��¼.�Ƿ���ſ���,
                 c_��¼.�Ƿ��ʱ��, c_��¼.ԤԼ����, c_��¼.�Ƿ��ռ, c_��¼.��Ŀid, c_��Դ.����id, c_��¼.ҽ��id, c_��¼.ҽ������, c_��¼.����ҽ��id, c_��¼.����ҽ������,
                 c_��¼.���﷽ʽ, c_��¼.����id, c_��¼.�Ƿ�����, c_��¼.�Ƿ���ʱ����, c_��¼.����Ա����, d_�Ǽ�����, 1);
            
              d_��ʼʱ�� := c_��¼.��ʼʱ��;
              --�����ٴ�������ſ���
              If Nvl(c_��¼.�Ƿ��ʱ��, 0) = 1 And Nvl(c_��¼.�Ƿ���ſ���, 0) = 1 Then
                --��ʱ����������ſ��ƣ�ʹ��"ԤԼ˳���"��¼"�Ƿ�ԤԼ"
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, ԤԼ˳���)
                  Select n_��¼id, ���,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_��ʼʱ�� > To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_��ʼʱ�� >= To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End, ��������, �Ƿ�ԤԼ, �Ƿ�ԤԼ
                  From �ٴ�����ʱ��
                  Where ����id = c_��¼.����id;
              Else
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
                  Select n_��¼id, ���,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_��ʼʱ�� > To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                                 'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_��ʼʱ�� >= To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End, ��������, �Ƿ�ԤԼ
                  From �ٴ�����ʱ��
                  Where ����id = c_��¼.����id;
              End If;
            
              --���������λ�Һſ��Ƽ�¼
              Insert Into �ٴ�����Һſ��Ƽ�¼
                (����, ����, ����, ��¼id, ���, ���Ʒ�ʽ, ����)
                Select ����, ����, ����, n_��¼id, ���, ���Ʒ�ʽ, ����
                From �ٴ�����Һſ���
                Where ����id = c_��¼.����id;
            
              --�����ٴ��������Ҽ�¼
              Insert Into �ٴ��������Ҽ�¼
                (��¼id, ����id)
                Select n_��¼id, ����id From �ٴ��������� Where ����id = c_��¼.����id;
            End Loop;
          
            --����ͣ�ﰲ�źͷ����ڼ��յ��������¼�ĳ���/ԤԼ���
            Zl_Clinicvisitmodify(c_��Դ.Id, n_����id, c_����.����, v_����Ա����, d_�Ǽ�����);
          End If;
        End If;
      End If;
      --һ��һ�ύ
      Commit;
    End Loop;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Auto_Buildingregisterplan;
/

--123726:Ƚ����,2018-03-31,�������ķ�������ʱ����
Create Or Replace Procedure Zl_���˷�������_Audit
(
  Id_In       ���˷�������.����id%Type,
  ����ʱ��_In ���˷�������.����ʱ��%Type,
  �����_In   ���˷�������.�����%Type,
  ���ʱ��_In ���˷�������.���ʱ��%Type,
  ״̬_In     ���˷�������.״̬%Type,
  Int�Զ����� Integer := 1,
  �������_In ���˷�������.�������%Type := 1 --��ҩƷ��������Ч,ȱʡΪ��ִ�е�ҩƷ������ 
) As
  n_ִ��״̬       סԺ���ü�¼.ִ��״̬%Type;
  n_�������       ���˷�������.�������%Type;
  v_�շ����       סԺ���ü�¼.�շ����%Type;
  v_No             סԺ���ü�¼.No%Type;
  n_ʵ������       ҩƷ�շ���¼.ʵ������%Type;
  n_����           ���˷�������.����%Type;
  n_�շ�id         ҩƷ�շ���¼.Id%Type;
  n_ҽ��id         סԺ���ü�¼.Id%Type;
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

  n_Cnt     Number(18);
  n_Temp    Number(18);
  v_Err_Msg Varchar2(300);
  Err_Item Exception;
Begin

  n_������� := 0;
  Select a.ִ��״̬, a.�շ����, a.�շ�ϸĿid, a.ִ�в���id, a.No, Nvl(b.��������, 0), a.ҽ�����, ����id, ��ҳid
  Into n_ִ��״̬, v_�շ����, n_�շ�ϸĿid, n_ִ�в���id, v_No, v_��������, n_ҽ��id, n_����id, n_��ҳid
  From סԺ���ü�¼ A, �������� B
  Where a.Id = Id_In And a.�շ�ϸĿid = b.����id(+);

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
  If Instr(',5,6,7', ',' || v_�շ����) > 0 Or (v_�շ���� = '4' And Nvl(v_��������, 0) = 1) Then
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
      Into n_Cnt
      From ����ҽ����¼ A, ����ҽ������ B, ��Һ��ҩ��¼ D
      Where a.Id = n_ҽ��id And a.Id = b.ҽ��id And b.No = v_No And a.���id = d.ҽ��id And b.���ͺ� = d.���ͺ� And b.��¼���� = 2 And
            d.����ʱ�� = ����ʱ��_In And d.����״̬ = 9;
    
      If n_Cnt <> 0 Then
        Select Count(1)
        Into n_Temp
        From ��Һ��ҩ״̬
        Where ��ҩid = n_Cnt And �������� = 10 And ����ʱ�� = ���ʱ��_In;
        If n_Temp = 0 Then
          Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (n_Cnt, 10, �����_In, ���ʱ��_In);
        End If;
        Update ��Һ��ҩ��¼ Set ������Ա = �����_In, ����ʱ�� = ���ʱ��_In, ����״̬ = 10 Where ID = n_Cnt;
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
      End If;
    Elsif Instr(',5,6,7,', ',' || v_�շ���� || ',') = 0 Then
      --���ܴ��ڲ�������,�����Ƚ���ҩƷ�Ĵ���ɲ���ִ��,����������˹���(ZL_סԺ���ʼ�¼_Delete)�д���,�����������: 
      --�ڵ��ñ�����ʱ: 
      --   1.������Ѿ�ִ�е�,���Ϊ����ִ��(ִ��״̬=2);�������ʹ����д����ⲿ������(ZL_סԺ���ʼ�¼_Delete):��:���ִ��״̬=2,���Ҳ������ʵ�,���Ϊ1(��ִ��) 
      --      ԭ������Ϊ��ҩƷ��ֻ�ܴ�������״̬.��ִ��;2-δִ�� 
      --   2.�����δִ�е�,��ִ��״̬����Ϊ0,�������ʹ����м�¼״̬���ֲ��� 
      Update סԺ���ü�¼ Set ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 0, 0, 2) Where ID = Id_In; --��ҩƷ����û��ȡ��ִ�еĲ���,���Զ���ִ�е�Ҫ�ȸ�״̬���ܵ����� 
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˷�������_Audit;
/

--123942:��ҵ��,2018-04-04,�������жϲ��������ã�������۸�
Create Or Replace Procedure Zl_ҩƷ�շ���¼_���۳���
(
  Id_In           In ������ü�¼.Id%Type,
  ҩƷժҪ_In     ҩƷ�շ���¼.ժҪ%Type := Null,
  Ƶ��_In         ҩƷ�շ���¼.Ƶ��%Type := Null,
  ����_In         ҩƷ�շ���¼.����%Type := Null,
  �÷�_In         ҩƷ�շ���¼.�÷�%Type := Null,
  �巨_In         ҩƷ�շ���¼.���%Type := Null,
  ��Ч_In         ҩƷ�շ���¼.����%Type := Null,
  �Ƽ�����_In     ҩƷ�շ���¼.����%Type := Null,
  ��ҳid_In       δ��ҩƷ��¼.��ҳid%Type := Null,
  ��������_In     Number := 0,
  ������������_In ҩƷ�շ���¼.����%Type := Null,
  ��ҩ����_In     ҩƷ�շ���¼.�Է�����id%Type := Null
) Is
  ----------------------------------
  --���ܣ��շѡ�����ʱ���ղ������÷ֽ�ҩƷ��������Ӧ���շ���¼
  --����
  --      1��ѭ���α��ж��ܳ����������α���ÿ����¼�����Ƿ���㣬�����������������������㰤������ֱ������ֱ�������겢�˳�
  --      2�������㷽ʽ������ȡ�շѼ�Ŀ���ּۣ�ʱ�۷���ȡ�������ۼۣ�ʱ�۲����������۽��/ʵ�������������������εĽ���ۼ�����Ϊ�ܳ�����
  --������
  --      Id_In��������ü�¼����סԺ���ü�¼ID
  --      ��������_In��ֻ�и�ֵ���Ĳ���Ҫ���룬��0��ʾ�Ǹ�ֵ����ģʽ
  --      ������������_In��֧�ָ�ֵ����ɨ��ȷ�����γ��⣬����35.70֧�ֲ��ϷǱ�������ģʽ�����γ��⣻ҩƷ��֧������ģʽ����ҩƷ���ζ����գ����������жϣ���ʹ���˷ǿգ�ֻҪ��ҩƷ����������
  --      ҩƷժҪ_In����ѡ����
  --      Ƶ��_In������_In���÷�_In���巨_In����Ч_In���Ƽ�����_In����ѡ������ҽ����¼����
  -----------------------------------
  Cursor c_Stock
  (
    n_Outmode  Number,
    n_�ⷿid   ҩƷ�շ���¼.�ⷿid%Type,
    n_ҩƷid   ҩƷ�շ���¼.ҩƷid%Type,
    n_�������� ҩƷ�շ���¼.����%Type,
    n_���     Number --0-����,1-ҩƷ
  ) Is
    Select �ⷿid, ҩƷid, Nvl(����, 0) ����, Ч��, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, ��׼�ĺ�,
           ƽ���ɱ���, ���ۼ�, �ϴο���, ��Ʒ����, �ڲ�����, ԭ����
    From ҩƷ��� A
    Where ҩƷid = n_ҩƷid And �ⷿid = n_�ⷿid And ���� = 1 And Decode(n_���, 0, Decode(n_��������, Null, 0, Nvl(����, 0)), 0) =
          Decode(n_���, 0, Decode(n_��������, Null, 0, Nvl(n_��������, 0)), 0) And
          (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate)) And Nvl(��������, 0) > 0
    Order By Decode(n_Outmode, 1, Ч��, Null), Nvl(����, 0);
  r_Stock c_Stock%RowType;

  n_Outmode      Number;
  n_����         ҩƷ���.ҩ������%Type;
  n_ʱ��         �շ���ĿĿ¼.�Ƿ���%Type;
  n_��ǰ����     ҩƷ���.ʵ������%Type;
  n_���ý��С�� Number;
  n_���õ���С�� Number;
  n_��ͨ���С�� Number;
  n_��ͨ����С�� Number;
  n_��׼����     �շѼ�Ŀ.�ּ�%Type;
  n_��ǰ����     �շѼ�Ŀ.�ּ�%Type;
  n_���         ҩƷ��������.���id%Type;
  n_�ܽ��       Number;
  n_������       ҩƷ���.ʵ������%Type;
  n_����         ҩƷ��������.����%Type;
  n_��������     ��������.��������%Type;
  n_���         ������ü�¼.���%Type;
  v_����         �շ���ĿĿ¼.����%Type;
  n_����ⷿid   ���ű�.Id%Type;
  n_�ⷿid       ���ű�.Id%Type;
  n_���ȼ�       ���.���ȼ�%Type;
  n_Count        Number;
  Err_Custom Exception;
  v_Rust     Varchar2(300);
  v_Error    Varchar2(255);
  v_�������� ���ű�.����%Type;

  v_�������   Varchar2(10);
  v_No         ҩƷ�շ���¼.No%Type;
  n_�Է�����id ҩƷ�շ���¼.�Է�����id%Type;
  n_�շ�ϸĿid ҩƷ�շ���¼.ҩƷid%Type;
  n_�ܳ������� ҩƷ���.ʵ������%Type;
  n_��ҩ�ⷿid ҩƷ�շ���¼.�ⷿid%Type;
  n_��¼����   ������ü�¼.��¼����%Type;
  v_�շ����   ������ü�¼.�շ����%Type;
  n_�ಡ�˵�   סԺ���ü�¼.�ಡ�˵�%Type;
  n_ҽ�����   ������ü�¼.ҽ�����%Type;
  v_����       ������ü�¼.����%Type;
  n_����       ������ü�¼.����%Type;
  v_����Ա     ������ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ������ü�¼.�Ǽ�ʱ��%Type;
  n_�����־   ������ü�¼.�����־%Type;
  n_���˿���id ������ü�¼.���˿���id%Type;
  n_��ʶ��     ������ü�¼.��ʶ��%Type;
  v_�Ա�       ������ü�¼.�Ա�%Type;
  n_����       ������ü�¼.����%Type;
  n_����id     ������ü�¼.����id%Type;
  v_��ҩ����   ������ü�¼.��ҩ����%Type;
  n_��¼״̬   ������ü�¼.��¼״̬%Type;

  --ҩƷ�շ���¼
  n_�շ�id   ҩƷ�շ���¼.Id%Type;
  n_����     ҩƷ�շ���¼.����%Type;
  d_���Ч�� ҩƷ�շ���¼.���Ч��%Type;
  d_������� ҩƷ�շ���¼.�������%Type;

  v_��������no ҩƷ�շ���¼.No%Type;
  n_�������   ҩƷ�շ���¼.���%Type;
  n_�����ۼ�   �շѼ�Ŀ.�ּ�%Type;
  n_������   Number(1);
Begin
  Begin
    Select ���, NO, ���, �Է�����id, �շ�ϸĿid, �ܳ�������, ��ҩ�ⷿid, ��¼����, �շ����, �ಡ�˵�, ҽ�����, ����, ����, ������, �Ǽ�ʱ��, �����־, ���˿���id, ��ʶ��, �Ա�,
           ����, ����id, ��ҩ����, ��¼״̬, ��׼����
    Into v_�������, v_No, n_���, n_�Է�����id, n_�շ�ϸĿid, n_�ܳ�������, n_��ҩ�ⷿid, n_��¼����, v_�շ����, n_�ಡ�˵�, n_ҽ�����, v_����, n_����, v_����Ա,
         d_�Ǽ�ʱ��, n_�����־, n_���˿���id, n_��ʶ��, v_�Ա�, n_����, n_����id, v_��ҩ����, n_��¼״̬, n_��׼����
    From (Select '����' As ���, NO, ���, ���˿���id As �Է�����id, �շ�ϸĿid, ���� * ���� As �ܳ�������, ִ�в���id As ��ҩ�ⷿid, ��¼����, �շ����, 0 As �ಡ�˵�,
                  ҽ�����, ����, ����, ������, �Ǽ�ʱ��, �����־, ���˿���id, ��ʶ��, �Ա�, ����, ����id, ��ҩ����, ��¼״̬, ��׼����
           From ������ü�¼
           Where ID = Id_In
           Union All
           Select 'סԺ' As ���, NO, ���, ���˿���id As �Է�����id, �շ�ϸĿid, ���� * ���� As �ܳ�������, ִ�в���id As ��ҩ�ⷿid, ��¼����, �շ����, �ಡ�˵�, ҽ�����,
                  ����, ����, Nvl(������, ����Ա����) As ������, �Ǽ�ʱ��, �����־, ���˿���id, ��ʶ��, �Ա�, ����, ����id, ��ҩ����, ��¼״̬, ��׼����
           From סԺ���ü�¼
           Where ID = Id_In);
  Exception
    When Others Then
      v_No         := Null;
      n_�Է�����id := 0;
      n_�ܳ������� := 0;
  End;

  Zl_ҩƷ���_���������쳣����(n_��ҩ�ⷿid, n_�շ�ϸĿid);

  n_�������� := 0;
  If v_�շ���� = '4' Then
    --��������
    Select �������� Into n_�������� From �������� Where ����id = n_�շ�ϸĿid;
  End If;

  --ҩƷ������������Ĳż������洦��
  If v_�շ���� In ('5', '6', '7') Or (v_�շ���� = '4' And Nvl(n_��������, 0) = 1) Then
  
    --סԺ��ҩ����ȷ��
    If v_������� = 'סԺ' Then
      n_�Է�����id := ��ҩ����_In;
    End If;
  
    --ֻ������������
    If n_�ܳ������� <> 0 Then
      If v_�շ���� = '4' Then
        --���ķ������ⷽʽ 
        Select Zl_To_Number(Nvl(zl_GetSysParameter(156), 0)) Into n_Outmode From Dual;
      Else
        --ҩƷ�������ⷽʽ
        Select Zl_To_Number(Nvl(zl_GetSysParameter(150), 0)) Into n_Outmode From Dual;
      End If;
    
      --���С��λ��
      Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')), Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))
      Into n_���ý��С��, n_���õ���С��
      From Dual;
    
      --ȡ��ͨҵ�񾫶�λ��
      --���:1-ҩƷ 2-����
      --���ݣ�2-���ۼ� 4-���
      --��λ��ҩƷ:1-�ۼ� 5-��λ
      If v_�շ���� = '4' Then
        Select ���� Into n_��ͨ����С�� From ҩƷ���ľ��� Where ��� = 2 And ���� = 2 And ��λ = 1;
        Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 2 And ���� = 4 And ��λ = 5;
      Else
        Select ���� Into n_��ͨ����С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 2 And ��λ = 1;
        Select ���� Into n_��ͨ���С�� From ҩƷ���ľ��� Where ��� = 1 And ���� = 4 And ��λ = 5;
      End If;
    
      n_������ := n_�ܳ�������;
    
      If v_�շ���� = '4' Then
        --�շ����=4��ʾ�����ĵ���
        If v_������� = '����' Then
          If n_��¼���� = 1 Then
            n_���� := 24;
          Else
            n_���� := 25;
          End If;
        Elsif v_������� = 'סԺ' Then
          If n_�ಡ�˵� = 1 Then
            n_���� := 26;
          Else
            n_���� := 25;
          End If;
        End If;
      
        Select Nvl(a.���÷���, 0), Nvl(b.�Ƿ���, 0), b.����, c.�ּ�
        Into n_����, n_ʱ��, v_����, n_�����ۼ�
        From �������� A, �շ���ĿĿ¼ B, �շѼ�Ŀ C
        Where a.����id = b.Id And b.Id = n_�շ�ϸĿid And b.Id = c.�շ�ϸĿid And Sysdate Between c.ִ������ And c.��ֹ����;
      
        --����������Ҫ�ж��Ƿ�����������ⷿ����
        If Nvl(��������_In, 0) = 1 Then
          Begin
            Select ����ⷿid Into n_����ⷿid From ����ⷿ���� Where ����id = n_��ҩ�ⷿid And Rownum <= 1;
          Exception
            When Others Then
              n_����ⷿid := 0;
          End;
          If Nvl(n_����ⷿid, 0) = 0 Then
            Begin
              Select ���� Into v_Error From ���ű� Where ID = n_��ҩ�ⷿid;
            Exception
              When Others Then
                v_Error := '';
            End;
            v_Error := 'ִ�в���"' || Nvl(v_Error, '') || '"δ�������ⲿ��,�������Ĳ�������������.';
            Raise Err_Custom;
          End If;
        End If;
      Else
        --�շ����<>4��ʾ��ҩƷ���ݣ��շ������"5��6��7"
        If v_������� = '����' Then
          If n_��¼���� = 1 Then
            n_���� := 8;
          Else
            n_���� := 9;
          End If;
        Elsif v_������� = 'סԺ' Then
          If n_�ಡ�˵� = 1 Then
            n_���� := 10;
          Else
            n_���� := 9;
          End If;
        End If;
      
        Select Nvl(a.ҩ������, 0), Nvl(b.�Ƿ���, 0), b.����, c.�ּ�
        Into n_����, n_ʱ��, v_����, n_�����ۼ�
        From ҩƷ��� A, �շ���ĿĿ¼ B, �շѼ�Ŀ C
        Where a.ҩƷid = b.Id And b.Id = n_�շ�ϸĿid And b.Id = c.�շ�ϸĿid And Sysdate Between c.ִ������ And c.��ֹ����;
      End If;
    
      --���ܷ���ʱ��ҩƷ�ֽ�����α���
      If n_ʱ�� = 1 Then
        --ֻ��һ������ʱ,ֱ��ȡ�����εĵ���
        --������С��λ���и�ʽ��
      
        If Nvl(��������_In, 0) = 1 And v_�շ���� = '4' Then
          v_Rust := Zl_Fun_Getprice(n_�շ�ϸĿid, n_����ⷿid, n_�ܳ�������, ��������_In, ������������_In);
        Else
          v_Rust := Zl_Fun_Getprice(n_�շ�ϸĿid, n_��ҩ�ⷿid, n_�ܳ�������, ��������_In, ������������_In);
        End If;
        n_��ǰ���� := To_Number(Substr(v_Rust, 1, Instr(v_Rust, '|') - 1));
      
        If Round(n_��ǰ����, n_���õ���С��) <> Round(n_��׼����, n_���õ���С��) Then
          If n_ҽ����� Is Null Then
            If v_�շ���� = '4' Then
              v_Error := '�� ' || n_��� || ' �е�ʱ����������"' || v_���� || '"��ǰ���㵥�۲�һ��,�����������������㣡';
            Else
              v_Error := '�� ' || n_��� || ' �е�ʱ��ҩƷ"' || v_���� || '"��ǰ���㵥�۲�һ��,�����������������㣡';
            End If;
          Else
            If v_�շ���� = '4' Then
              v_Error := '�ڴ�����"' || v_���� || '"ʱ����ʱ����������"' || v_���� || '"��ǰ����ĵ��۷����仯��' || Chr(13) || Chr(10) ||
                         '����ò����Ƿ�ͬʱʹ����������ͬ��"' || v_���� || '"��';
            Else
              v_Error := '�ڴ�����"' || v_���� || '"ʱ����ʱ��ҩƷ"' || v_���� || '"��ǰ����ĵ��۷����仯��' || Chr(13) || Chr(10) ||
                         '����ò����Ƿ�ͬʱʹ����������ͬ��"' || v_���� || '"��';
            End If;
          End If;
          Raise Err_Custom;
        End If;
      End If;
    
      If v_�շ���� In ('5', '6', '7') Or (v_�շ���� = '4' And Nvl(n_��������, 0) = 1) Then
        If Nvl(��������_In, 0) = 1 And v_�շ���� = '4' Then
          n_�ⷿid := n_����ⷿid;
        Else
          n_�ⷿid := n_��ҩ�ⷿid;
        End If;
      
        Begin
          If v_�շ���� In ('5', '6', '7') Then
            Select ��鷽ʽ Into n_������ From ҩƷ������ Where �ⷿid = n_�ⷿid;
          Else
            Select ��鷽ʽ Into n_������ From ���ϳ����� Where �ⷿid = n_�ⷿid;
          End If;
        Exception
          When Others Then
            n_������ := 0;
        End;
      
        If v_�շ���� = '4' Then
          Select ���id Into n_��� From ҩƷ�������� Where ���� = n_���� + 16;
        Else
          Select ���id Into n_��� From ҩƷ�������� Where ���� = n_����;
        End If;
      
        n_�ܽ�� := 0;
        --���α�
        If v_�շ���� = '4' Then
          Open c_Stock(n_Outmode, n_�ⷿid, n_�շ�ϸĿid, ������������_In, 0);
        Else
          Open c_Stock(n_Outmode, n_�ⷿid, n_�շ�ϸĿid, ������������_In, 1);
        End If;
        --ѭ������
        While n_�ܳ������� <> 0 Loop
          Fetch c_Stock
            Into r_Stock;
          If c_Stock%NotFound Then
            --��һ�ξ�û�п��,������ʱ�۶�������
            --����ҩƷ�����ֽⲻ��,Ҳ���ǿ�治�㡣
            If n_���� = 1 Or n_ʱ�� = 1 Then
              Close c_Stock;
              If n_���� = 8 Or n_���� = 24 Then
                If v_�շ���� = '4' Then
                  v_Error := '�� ' || n_��� || ' �еķ�����ʱ����������"' || v_���� || '"û�п��õĿ�棡';
                Else
                  v_Error := '�� ' || n_��� || ' �еķ�����ʱ��ҩƷ"' || v_���� || '"û�п��õ�ҩƷ��棡';
                End If;
              Else
                --����=9��10��25��26�Ǽ��˵���ʾ��һ��
                If n_ҽ����� Is Null Then
                  If v_�շ���� = '4' Then
                    If Nvl(��������_In, 0) = 1 And Not (n_���� = 1 Or n_ʱ�� = 1) Then
                      v_Error := '�� ' || n_��� || ' �е���������"' || v_���� || '"û���㹻�Ĳ��Ͽ��,���ܽ��б������ʣ�';
                    Else
                      v_Error := '�� ' || n_��� || ' �еķ�����ʱ����������"' || v_���� || '"û���㹻�Ĳ��Ͽ��' || Case
                                   When Nvl(��������_In, 0) = 0 Then
                                    '��'
                                   Else
                                    ',���ܽ��б������ʣ�'
                                 End;
                    End If;
                  Else
                    v_Error := '�� ' || n_��� || ' �еķ�����ʱ��ҩƷ"' || v_���� || '"û���㹻�Ŀ�棡';
                  End If;
                Else
                  If v_�շ���� = '4' Then
                    If Nvl(��������_In, 0) = 1 And Not (n_���� = 1 Or n_ʱ�� = 1) Then
                      v_Error := '�ڴ�����"' || v_���� || '"ʱ������������"' || v_���� || '"û���㹻�Ĳ��Ͽ��,���ܽ��б������ʣ�';
                    Else
                      v_Error := '�ڴ�����"' || v_���� || '"ʱ���ַ�����ʱ����������"' || v_���� || '"û���㹻�Ĳ��Ͽ��' || Case
                                   When Nvl(��������_In, 0) = 0 Then
                                    '��'
                                   Else
                                    ',���ܽ��б������ʣ�'
                                 End;
                    End If;
                  Else
                    v_Error := '�ڴ�����"' || v_���� || '"ʱ���ַ�����ʱ��ҩƷ"' || v_���� || '"û���㹻�Ŀ�棡';
                  End If;
                End If;
              End If;
              Raise Err_Custom;
            End If;
          Elsif (n_���� = 1 And Nvl(r_Stock.����, 0) = 0) Or (n_���� = 0 And Nvl(r_Stock.����, 0) <> 0) Then
            Close c_Stock;
            If n_ҽ����� Is Null Then
              If v_�շ���� = '4' Then
                v_Error := '�� ' || n_��� || ' ����������"' || v_���� || '"�����÷������������¼�����,����������ݵ���ȷ�ԣ�';
              Else
                v_Error := '�� ' || n_��� || ' ��ҩƷ"' || v_���� || '"�ķ������������¼�����,����ҩƷ���ݵ���ȷ�ԣ�';
              End If;
            Else
              If v_�շ���� = '4' Then
                v_Error := '�ڴ�����"' || v_���� || '"ʱ������������"' || v_���� || '"�ķ������������¼�����,����������ݵ���ȷ�ԣ�';
              Else
                v_Error := '�ڴ�����"' || v_���� || '"ʱ����ҩƷ"' || v_���� || '"�ķ������������¼�����,����ҩƷ���ݵ���ȷ�ԣ�';
              End If;
            End If;
            Raise Err_Custom;
          End If;
        
          If c_Stock%Found Then
            If Nvl(r_Stock.ʵ������, 0) = 0 And (n_�ܳ������� > 0 Or n_ʱ�� = 1) And n_������ = 2 Then
              --ʵ������Ϊ��ʱ������ϸ���ƿ�棬���������
              --ʵ��������Ϊ�㣬���Ϊ�㣬��������������۸����
              --����������൱�����,�������Ӧ������ģ���ʱ����Ҫ����۸񣬱���Ҫ��ʵ��������
              Close c_Stock;
              If n_ҽ����� Is Null Then
                If v_�շ���� = '4' Then
                  v_Error := '�� ' || n_��� || ' �е���������"' || v_���� || '"��ǰ�޿��ʵ�����������ܴ�����δ���ϵļ�¼����ǰ���ܳ��⡣';
                Else
                  v_Error := '�� ' || n_��� || ' ��ҩƷ"' || v_���� || '"��ǰ�޿��ʵ�����������ܴ�����δ��ҩ�ļ�¼����ǰ���ܳ��⡣';
                End If;
              Else
                If v_�շ���� = '4' Then
                  v_Error := '�ڴ�����"' || v_���� || '"ʱ������������"' || v_���� || '"��ǰ�޿��ʵ�����������ܴ�����δ���ϵļ�¼����ǰ���ܳ��⡣';
                Else
                  v_Error := '�ڴ�����"' || v_���� || '"ʱ����ҩƷ"' || v_���� || '"��ǰ�޿��ʵ�����������ܴ�����δ��ҩ�ļ�¼����ǰ���ܳ��⡣';
                End If;
              End If;
              Raise Err_Custom;
            End If;
          End If;
        
          If n_���� = 1 Or n_ʱ�� = 1 Then
            --���ڲ�������ʱ��ֻ���ֽܷ�һ��,�ֽⲻ�������ж���.���ֽ���Ϊ�˼��㵥��.
            --ÿ�ηֽ�ȡС��,��治���ֽⲻ���������ж�.
            If n_�ܳ������� <= Nvl(r_Stock.��������, 0) Then
              n_��ǰ���� := n_�ܳ�������;
            Else
              n_��ǰ���� := Nvl(r_Stock.��������, 0);
            End If;
            If n_ʱ�� = 1 Then
              n_��ǰ���� := Nvl(r_Stock.���ۼ�, Nvl(r_Stock.ʵ�ʽ�� / r_Stock.ʵ������, 0));
            Elsif n_���� = 1 Then
              n_��ǰ���� := n_�����ۼ�;
            End If;
          Else
            --���۲�����
            --�����ﵥ�����Ǹ�ֵ������Ҫ�����
            If n_���� <> 8 Or n_���� <> 24 Then
              If Nvl(��������_In, 0) = 1 And v_�շ���� = '4' Then
                If n_�ܳ������� > Nvl(r_Stock.��������, 0) Then
                  --������, �����Ǳ������ķ�ʽ�����,����Ҫ��鵱ǰ����Ƿ����.
                  v_Error := '�� ' || n_��� || ' �е���������"' || v_���� || '"û���㹻�Ĳ��Ͽ��,���ܽ��б������ʣ�';
                  Raise Err_Custom;
                End If;
              End If;
            End If;
            n_��ǰ���� := n_�ܳ�������;
            n_��ǰ���� := n_�����ۼ�;
          End If;
        
          --ҩƷ�շ���¼
          If c_Stock%Found Then
            --�������Ч��:һ���Բ�������Ч��
            If v_�շ���� = '4' Then
              n_Count := 0;
              Begin
                Select ���Ч�� Into n_Count From �������� Where Nvl(һ���Բ���, 0) = 1 And ����id = n_�շ�ϸĿid;
              Exception
                When Others Then
                  Null;
              End;
              If Nvl(n_Count, 0) > 0 Then
                d_���Ч�� := r_Stock.���Ч��;
                d_������� := d_���Ч�� - n_Count * 30;
              End If;
            End If;
          End If;
        
          Select Nvl(Max(���), 0) + 1
          Into n_���
          From ҩƷ�շ���¼
          Where ���� = n_���� And ��¼״̬ = 1 And NO = v_No;
        
          n_���� := Null;
          If ��Ч_In Is Not Null Or �Ƽ�����_In Is Not Null Then
            n_���� := Nvl(��Ч_In, 0) || Nvl(�Ƽ�����_In, 0);
          End If;
        
          --����ҩƷ,�����ֻʹ����һ������,��Ҫ��д����
          If n_���� = 1 And n_��ǰ���� <> n_������ Then
            n_Count := 1;
          Else
            n_Count := 0;
          End If;
        
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
          --�޸ĵ�ԭ���ݺŴ����ժҪ��
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, ���ۼ�, ���۽��, ժҪ, ������,
             ��������, ����id, Ƶ��, ��ҩ����, ����, �÷�, ���, ����, ���Ч��, �������, ��ҩ��λid, ��������, ��׼�ĺ�, ��Ʒ����, �ڲ�����, ԭ����)
          Values
            (n_�շ�id, 1, n_����, v_No, n_���, n_��ҩ�ⷿid, n_�Է�����id, n_���, -1, n_�շ�ϸĿid, Nvl(r_Stock.����, 0), r_Stock.�ϴβ���,
             r_Stock.�ϴ�����, r_Stock.Ч��, Decode(n_Count, 1, 1, n_����), Decode(n_Count, 1, n_��ǰ����, n_��ǰ���� / n_����),
             Decode(n_Count, 1, n_��ǰ����, n_��ǰ���� / n_����), n_��ǰ����, Round(n_��ǰ���� * n_��ǰ����, n_��ͨ���С��), ҩƷժҪ_In, v_����Ա, d_�Ǽ�ʱ��,
             Id_In, Ƶ��_In, v_��ҩ����, ����_In, �÷�_In, �巨_In, n_����, d_���Ч��, d_�������, r_Stock.�ϴι�Ӧ��id, r_Stock.�ϴ���������,
             r_Stock.��׼�ĺ�, r_Stock.��Ʒ����, r_Stock.�ڲ�����, r_Stock.ԭ����);
        
          Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
        
          --ҩƷ���(��ͨ�������û�м�¼)
          Zl_ҩƷ���_Update(n_�շ�id, 0, 1);
        
          --�����������ⵥ ��ֻ�и�ֵ���Ĳ���Ҫ����
          If v_�շ���� = '4' And Nvl(��������_In, 0) = 1 Then
            Begin
              Select Max(a.No), Max(a.���)
              Into v_��������no, n_�������
              From ҩƷ�շ���¼ A, סԺ���ü�¼ B
              Where a.����id = b.Id And b.No = v_No And ��¼���� = 2 And b.�����־ = n_�����־ And
                    Instr(',8,9,10,21,24,25,26,', ',' || a.���� || ',') > 0;
            Exception
              When Others Then
                v_��������no := Null;
            End;
            If v_��������no Is Null Then
              v_��������no := Nextno(74, n_����ⷿid, Null, 1);
            End If;
            If v_��������no Is Null Then
              v_Error := '�������������ϵ��������ⵥʱ,��ȡ��صĳ���NO����,������ⵥ�Ĺ����Ƿ�����!';
              Raise Err_Custom;
            End If;
            If Nvl(n_���˿���id, 0) <> 0 Then
              Select ���� Into v_�������� From ���ű� Where ID = n_���˿���id;
            End If;
            v_Error := LPad(' ', 4);
            v_Error := Substr('��������:' || v_���� || v_Error || '�Ա�:' || v_�Ա� || v_Error || '����' || n_���� || v_Error ||
                              '�����:' || Nvl(n_��ʶ��, '') || v_Error || '���˿���:' || v_��������, 1, 100);
          
            n_������� := Nvl(n_�������, 0) + 1;
            Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
          
            --��ֵ�������idĬ��19��Ϊ�˷���ͳ�ƣ���Ϊ��������������úܶ��������Ĭ��19
            Insert Into ҩƷ�շ���¼
              (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, ���ۼ�, ���۽��, ժҪ,
               ������, ��������, ����id, Ƶ��, ��ҩ����, ����, �÷�, ���, ����, ���Ч��, �������, ��ҩ��λid, ��������, ��׼�ĺ�, ��Ʒ����, �ڲ�����, ԭ����)
            Values
              (n_�շ�id, 1, 21, v_��������no, n_�������, n_����ⷿid, n_�Է�����id, 19, -1, n_�շ�ϸĿid, Nvl(r_Stock.����, 0), r_Stock.�ϴβ���,
               r_Stock.�ϴ�����, r_Stock.Ч��, 1, n_��ǰ����, n_��ǰ����, n_��ǰ����, Round(n_��ǰ���� * n_��ǰ����, n_��ͨ���С��), v_Error, v_����Ա,
               d_�Ǽ�ʱ��, Id_In, Ƶ��_In, v_��ҩ����, ����_In, �÷�_In, �巨_In, n_����, d_���Ч��, d_�������, r_Stock.�ϴι�Ӧ��id, r_Stock.�ϴ���������,
               r_Stock.��׼�ĺ�, r_Stock.��Ʒ����, r_Stock.�ڲ�����, r_Stock.ԭ����);
          
            Zl_δ��ҩƷ��¼_Insert(n_�շ�id);
          
            --ҩƷ���(��ͨ�������û�м�¼)
            Zl_ҩƷ���_Update(n_�շ�id, 0, 1);
          End If;
        
          v_Error      := '';
          n_�ܳ������� := n_�ܳ������� - n_��ǰ����;
          n_�ܽ��     := n_�ܽ�� + n_��ǰ���� * n_��ǰ����;
        End Loop;
      
        --δ��ҩƷ��¼
        Update δ��ҩƷ��¼
        Set ����id = n_����id, ���� = v_����, ��ҩ���� = v_��ҩ����, ��ҳid = ��ҳid_In
        Where ���� = n_���� And NO = v_No And Nvl(�ⷿid, 0) = Nvl(n_��ҩ�ⷿid, 0);
        If Sql%RowCount = 0 Then
          --ȡ������ȼ�
          Begin
            Select b.���ȼ� Into n_���ȼ� From ������Ϣ A, ��� B Where a.��� = b.����(+) And a.����id = n_����id;
          Exception
            When Others Then
              Null;
          End;
          Insert Into δ��ҩƷ��¼
            (����, NO, ����id, ��ҳid, ����, ���ȼ�, �ⷿid, �Է�����id, ��������, ���շ�, ��ӡ״̬, ��ҩ����)
          Values
            (n_����, v_No, n_����id, ��ҳid_In, v_����, n_���ȼ�, n_��ҩ�ⷿid, n_�Է�����id, d_�Ǽ�ʱ��, n_��¼״̬, 0, v_��ҩ����);
        End If;
      
        --����δ��ҩ��¼״̬
        Zl_Prescription_Type_Update(v_No, n_��¼����, n_�շ�ϸĿid, v_�շ����);
      
        Close c_Stock;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_���۳���;
/


------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0005' Where ���=&n_System;
Commit;
