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
--127820:��ΰ��,2018-06-26,������ҩ����ṩ�����󷽽ӿ�
Insert Into ������������Ŀ¼ (ϵͳ��ʶ, ��������) Values ('ҩʦ�������', '�������ѯ');

Insert Into ������������Ŀ¼ (ϵͳ��ʶ, ��������) Values ('ҩʦ�������', '��дҽ���ܾ�����');

--127571:��͢��,2018-06-25,���ڼ�¼���Ӳ������Ĵ����ϴ�ѡ����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1259, 1, 1, 0, 0, 0, 0, 1, 'ȱʡ��ʾ��Ϣ', Null, Null,
         '���ڼ�¼���Ӳ������Ĵ����ϴ�ѡ��ڵ����Keyֵ,�����´δ򿪵��Ӳ�������ʱ�ָ��ϴ�ѡ��Ľڵ�', '��¼���Ӳ������Ĵ����ϴ�ѡ��ڵ����Keyֵ', '', '�����´δ򿪵��Ӳ�������ʱ�ָ��ϴ�ѡ��Ľڵ�', Null
  From Dual;

--126645:������,2018-06-22,���ﲡ��ԤԼ��Ժ����޸�
Insert Into Zlprocedure(Id, ����, ����, ״̬, ������, ˵��) Values (Zlprocedure_Id.Nextval,2,'Zl_Third_Outpatireg',3,User,'���ڲ���Ԥ��Ժ��¼/ȡ��Ԥ��Ժ��������Ρ����Ρ�����ֵ˵���������vssData/DataStructure/���������ӿ�˵��(Oracle).xlsx��');

--126645:������,2018-06-22,���ﲡ��ԤԼ��Ժ����޸�
Insert Into ������������Ŀ¼(ϵͳ��ʶ,��������) 
Select 'ԤԼ����','סԺ����' From Dual Union All
Select 'ԤԼ����','סԺ����ȡ��' From Dual;

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--124567:������,2018-06-26,����ҽ��վ������ʾ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select 100,1260,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select '�ٴ������¼','SELECT' From Dual
Union All Select '�ٴ������Դ','SELECT' From Dual) A;


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--127819:��ΰ��,2018-06-27,����������ҩ�����ϴ�
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
  Xml_Node_Pati       Xmldom.Domnode;
  Xml_Node            Xmldom.Domnode;
  Xml_Node_New        Xmldom.Domnode;
  ----------------------------------
  n_��� Number(10, 2); --��λ:cm
  n_���� Number(10, 2); --����:KG
  v_Type Varchar2(200);

  l_Clob    Clob;
  v_Err_Msg Varchar2(2000);
  v_Temp    Varchar2(200);
  v_Value   Varchar2(200);
  n_Nodenum Number(5);
  Err_Item Exception;

  Procedure Addpatiinfo
  (
    Nodeparent Xmldom.Domnode,
    Nodecopy   Xmldom.Domnode,
    Nodename   Varchar2,
    Nodevalue  Varchar2
  ) Is
    Nodenew Xmldom.Domnode;
  Begin
    Nodenew := Xmldom.Appendchild(Nodeparent, Xmldom.Clonenode(Nodecopy, False));
    Xmldom.Setattribute(Xmldom.Makeelement(Nodenew), 'name', Nodename);
    Xmldom.Setattribute(Xmldom.Makeelement(Nodenew), 'value', Nodevalue);
  End;
Begin

  --��
  --��CLOB������ȡ��v_XML��
  Select �������� Into l_Clob From ����������ҩ����;
  Xml_Ret        := Xmltype(l_Clob); --���溯������ֵ
  Xml_Document   := Xmldom.Newdomdocument(Xml_Ret);
  Xml_Domelement := Xmldom.Getdocumentelement(Xml_Document);
  Xml_Nodelist   := Xmldom.Getelementsbytagname(Xml_Domelement, 'patient_info');
  Xml_Node_Pati  := Xmldom.Item(Xml_Nodelist, 0);
  --��ȡpatient_info/INfo�ڵ�
  Xml_Nodelist := Xmldom.Getchildnodes(Xml_Node_Pati);
  n_Nodenum    := Xmldom.Getlength(Xml_Nodelist);
  For I In 0 .. n_Nodenum - 1 Loop
    Xml_Node            := Xmldom.Item(Xml_Nodelist, I);
    Xml_Domnamednodemap := Xmldom.Getattributes(Xml_Node);
    v_Temp              := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'name'));
    If v_Temp = '�ύ����' Then
      v_Type := Xmldom.Getnodevalue(Xmldom.Getnameditem(Xml_Domnamednodemap, 'value'));
      If v_Type = '2' Then
        --1-�¿�δ����;2-����ҽ����
        Addpatiinfo(Xml_Node_Pati, Xml_Node, '����ID', ����id_In);
        If Nvl(�Һ�id_In, 0) = 0 Then
          Addpatiinfo(Xml_Node_Pati, Xml_Node, '����ID', ��ҳid_In);
        Else
          Addpatiinfo(Xml_Node_Pati, Xml_Node, '����ID', �Һ�id_In);
        End If;
      End If;
    End If;
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

--126645:������,2018-06-22,���ﲡ��ԤԼ��Ժ����޸�
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
  ԭҺƤ��_In   In Varchar2 := Null,
  ԤԼ����_In   In Number := 0
  --���ܣ���д����ҽ�����ͼ�¼
  --������First_IN=��ʾ�Ƿ�һ��ҽ���ĵ�һҽ����,�Ա㴦��ҽ���������(���ҩ,�䷽�ĵ�һ��,��Ϊ��ҩ;��,�䷽�巨,�÷�����Ϊ����������)
  --      ԴҺƤ��_In ԭҺƤ��ҽ��ID�������7107/bug115972���ڹ���ҩƷҽ���к�Ƥ��ҽ���С������ֶ�Ϊ ����ҽ������.�걾�������� ����ҩƷ�е�ҽ��IDֵ
  --      ԤԼ����_in �Ƿ�������ԺԤԼ���ģ��ɳ����ⲿ����
  --      ��ʽ��1ҽ��ID,2ҽ��ID ǰ��һ��ΪƤ��ҽ����ҽ��ID���ڶ���ΪҩƷ��ҽ����ҽ��ID  
) Is
  --�������˼�ҽ��(һ��ҽ���е�һ��)�����Ϣ���α�
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.���id, a.���, a.����id, a.�Һŵ�, a.Ӥ��, b.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��,
           a.��ʼִ��ʱ��, a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, Nvl(a.������־, 0) As ������־, a.������Ŀid, a.�շ�ϸĿid
    From ����ҽ����¼ A, ������Ϣ B, ������ĿĿ¼ C
    Where a.����id = b.����id And a.������Ŀid = c.Id And a.Id = ҽ��id_In
    Group By a.���id, a.Id, a.���, a.����id, a.�Һŵ�, a.Ӥ��, b.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��, a.��ʼִ��ʱ��, a.ִ��ʱ�䷽��,
             a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.������־, a.������Ŀid, a.�շ�ϸĿid;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select * From ������Ϣ Where ����id = v_����id;
  r_Pati c_Pati%RowType;

  --������ʱ����
  v_Temp       Varchar2(255);
  v_Count      Number;
  v_��������   ������ҳ.��������%Type;
  v_��Ա���   ��Ա��.���%Type;
  v_��Ա����   ��Ա��.����%Type;
  v_��Ժ��ʽ   ��Ժ��ʽ.����%Type;
  n_�Һ�id     ���˹Һż�¼.Id%Type;
  d_��ʼʱ��   ����ҽ����¼.��ʼִ��ʱ��%Type;
  n_ҽ��״̬   ����ҽ����¼.ҽ��״̬%Type;
  n_Ƥ�Ա��   ����ҽ������.ҽ��id%Type;
  n_Ƥ��ҽ��id ����ҽ������.ҽ��id%Type;
  v_Error      Varchar2(255);
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
  Select ��ʼִ��ʱ��, ҽ��״̬ Into d_��ʼʱ��, n_ҽ��״̬ From ����ҽ����¼ Where ID = ҽ��id_In;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  --��һ��ҽ���ĵ�һ��ʱ����ҽ������
  If Nvl(First_In, 0) = 1 Or n_ҽ��״̬ = 1 Then
    --�����������
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.ҽ��״̬, 0) <> 1 Then
      v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ��������˷��͡�' || Chr(13) || Chr(10) ||
                 '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
      Raise Err_Custom;
    End If;
  
    --���ͺ��ҽ������:�������ͺ��Զ�ֹͣ
    ---------------------------------------------------------------------------------------
    Update ����ҽ����¼
    Set ҽ��״̬ = 8, ִ����ֹʱ�� = ĩ��ʱ��_In,
        --����û��
        ͣ��ʱ�� = ����ʱ��_In,
        --Ҫ��Ϊ����ʱ����ʾ
        ͣ��ҽ�� = v_��Ա���� --Ҫ��Ϊ��������ʾ,��ͬ��סԺ,����ҽ���޻�ʿ����
    Where ID = r_Advice.��id Or ���id = r_Advice.��id;
  
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��)
      Select ID, 8, v_��Ա����, ����ʱ��_In From ����ҽ����¼ Where ID = r_Advice.��id Or ���id = r_Advice.��id;
  
    --����ҽ���Ĵ���
    ---------------------------------------------------------------------------------------
    If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      --1-����;2-סԺ;    
      --סԺҽ����ԤԼ����Ӱ���Ƚ��������ж�
      If r_Advice.�������� = '1' And ִ�в���id_In Is Not Null Then
        v_Count := 1;
      Elsif r_Advice.�������� = '2' And ִ�в���id_In Is Not Null Then
        v_Count := 1;
        If ԤԼ����_In = 1 Then
          v_Count := 0;
        End If;
      Else
        v_Count := 0;
      End If;
    
      If v_Count = 1 Then
        --��������µ�ԤԼ�Ǽǵ�������1.��ǰ��ԤԼ,2.��ǰ����Ժ,3-��Ҫ��ԤԼʱ���ڵ�סԺ��¼
      
        --ɾ�������Һ���Ч������ԤԼ�Ǽ�
        Begin
          Select Count(*) Into v_Count From ������ҳ Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0;
        Exception
          When Others Then
            v_Count := 0;
        End;
        If Nvl(v_Count, 0) > 0 Then
          Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0, 0, 0);
          v_Count := 0;
        End If;
      
        If v_Count = 0 Then
          Select Count(*) Into v_Count From ������ҳ Where ����id = r_Advice.����id And ��Ժ���� Is Null;
        End If;
        If v_Count = 0 Then
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And (��Ժ���� >= r_Advice.��ʼִ��ʱ�� Or ��Ժ���� >= r_Advice.��ʼִ��ʱ��);
        End If;
        If v_Count = 0 Then
          If r_Advice.�������� = '1' Then
            --����ҽ��,��������"��ʼʱ��"���۵��ٴ�ִ�п���
            Begin
              v_�������� := 2;
              Select Decode(�������, 1, 1, 2)
              Into v_��������
              From ��������˵��
              Where �������� = '�ٴ�' And ����id = ִ�в���id_In;
            Exception
              When Others Then
                Null;
            End;
          Elsif r_Advice.�������� = '2' Then
            --סԺҽ��,��������"��ʼʱ��"�Ǽǵ��ٴ�ִ�п���
            v_�������� := 0;
          End If;
        
          Open c_Pati(r_Advice.����id);
          Fetch c_Pati
            Into r_Pati;
        
          v_��Ժ��ʽ := Null;
          If r_Advice.������־ = 1 Then
            v_��Ժ��ʽ := '����';
            Select Max(ID)
            Into n_�Һ�id
            From ���˹Һż�¼
            Where NO = r_Advice.�Һŵ� And ��¼���� = 1 And ��¼״̬ = 1;
          Else
            Select Decode(����, 1, '����', Null), ID
            Into v_��Ժ��ʽ, n_�Һ�id
            From ���˹Һż�¼
            Where NO = r_Advice.�Һŵ� And ��¼���� = 1 And ��¼״̬ = 1;
          End If;
          If v_�������� = 1 Then
            Zl_��Ժ������ҳ_Insert(1, v_��������, r_Pati.����id, r_Pati.�����, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�,
                             r_Pati.��������, r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���,
                             r_Pati.���֤��, r_Pati.�����ص�, r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ,
                             r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ, r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ,
                             r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������, r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������,
                             r_Pati.��������, ִ�в���id_In, Null, Null, v_��Ժ��ʽ, Null, Null, r_Advice.����ҽ��, r_Pati.����, r_Pati.����,
                             r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null, Null, r_Pati.����,
                             v_��Ա���, v_��Ա����, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_�Һ�id);
          Else
            Zl_��Ժ������ҳ_Insert(1, v_��������, r_Pati.����id, r_Pati.סԺ��, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�,
                             r_Pati.��������, r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���,
                             r_Pati.���֤��, r_Pati.�����ص�, r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ,
                             r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ, r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ,
                             r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������, r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������,
                             r_Pati.��������, ִ�в���id_In, Null, Null, v_��Ժ��ʽ, Null, Null, r_Advice.����ҽ��, r_Pati.����, r_Pati.����,
                             r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null, Null, r_Pati.����,
                             v_��Ա���, v_��Ա����, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null, Null, n_�Һ�id);
          End If;
          Close c_Pati;
        End If;
      End If;
    End If;
  End If;
  Close c_Advice;

  If ԭҺƤ��_In Is Not Null Then
    v_Count      := Instr(ԭҺƤ��_In, ',');
    n_Ƥ��ҽ��id := Substr(ԭҺƤ��_In, 1, v_Count - 1);
    n_Ƥ�Ա��   := Substr(ԭҺƤ��_In, v_Count + 1);
    Update ����ҽ������ Set �걾�������� = n_Ƥ�Ա�� Where ҽ��id = n_Ƥ��ҽ��id;
  End If;
  --��д���ͼ�¼
  ---------------------------------------------------------------------------------------
  Insert Into ����ҽ������
    (ҽ��id, ���ͺ�, ��¼����, NO, ��¼���, ��������, ������, ����ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬, �״�ʱ��, ĩ��ʱ��, ��������, �������, �걾��������)
  Values
    (ҽ��id_In, ���ͺ�_In, ��¼����_In, No_In, ��¼���_In, ��������_In, v_��Ա����, ����ʱ��_In, ִ��״̬_In, ִ�в���id_In, �Ʒ�״̬_In,
     Nvl(�״�ʱ��_In, d_��ʼʱ��), Nvl(ĩ��ʱ��_In, d_��ʼʱ��), ��������_In, Decode(��¼����_In, 2, 1, Null), n_Ƥ�Ա��);

  --�����ͼ��ҽ��ͬ��������ҽ���ļƷ�״̬
  If �Ʒ�״̬_In = 1 And r_Advice.��id <> ҽ��id_In And (r_Advice.������� = 'D' Or r_Advice.������� = 'F') Then
    Update ����ҽ������ Set �Ʒ�״̬ = 1 Where ҽ��id = r_Advice.��id And ���ͺ� = ���ͺ�_In;
  End If;

  --�Զ���Ϊ��ִ��ʱ����Ҫͬ���������ִ��״̬����˻���״̬
  If ִ��״̬_In = 1 Then
    Zl_����ҽ��ִ��_Finish(ҽ��id_In, ���ͺ�_In, Null, Null, v_��Ա���, v_��Ա����, ִ�в���id_In);
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin zl_������Ϣ_����(:1,:2); End;'
      Using 3, ���ͺ�_In;
  Exception
    When Others Then
      Null;
  End;

  If r_Advice.������� = 'E' And r_Advice.�������� = '6' Then
    --������Ŀ
    b_Message.Zlhis_Cis_016(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id, 1);
  Elsif r_Advice.������� = 'D' And r_Advice.���id Is Null Then
    b_Message.Zlhis_Cis_017(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id, 1);
  Elsif r_Advice.������� = 'F' And r_Advice.���id Is Null Then
    b_Message.Zlhis_Cis_018(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id);
  Elsif r_Advice.������� = 'K' Then
    b_Message.Zlhis_Cis_019(r_Advice.����id, Null, r_Advice.�Һŵ�, ���ͺ�_In, r_Advice.��id);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ������_Insert;
/

--126645:������,2018-06-28,���ﲡ��ԤԼ��Ժ����޸�

Create Or Replace Procedure Zl_Third_Outpatireg
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ����ڲ���Ԥ��Ժ��¼/ȡ��Ԥ��Ժ    ����д��
  --��Σ�xml_in
  --<IN>
  -- <TYPE>1</TYPE>   --�������ͣ�1-����Ԥ��Ժ��¼��0-ȡ��Ԥ��Ժ
  -- <GHID>1162695</GHID>       --�Һ�id
  -- <RYKSID>202704</RYKSID>    --��Ժ����ID
  -- <RYBQID>202704</RYBQID>    --��Ժ����ID
  -- <CH>5</CH>   --���� 
  -- <YZID>3</YZID> --ҽ��id
  -- <CZYBH></CZYBH> --����Ա���
  -- <CZYXM></CZYXM> --����Ա����  
  --</IN>

  --���Σ�Xml_Out
  --<OUTPUT>
  --   <RESULT>true</RESULT>
  --</OUTPUT>

  --ʧ�ܣ�
  --<OUTPUT>
  --   <RESULT>false</RESULT>
  --   <ERROR>
  --     <MSG>��ϸ������ʾ</MSG>
  --   </ERROR>
  --</OUTPUT>

  n_ҽ��id ����ҽ����¼.Id%Type;
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.���id, a.���, a.����id, a.�Һŵ�, a.Ӥ��, a.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��,
           a.��ʼִ��ʱ��, a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, Nvl(a.������־, 0) As ������־, a.������Ŀid, a.�շ�ϸĿid
    From ����ҽ����¼ A, ������ĿĿ¼ C
    Where a.������Ŀid = c.Id And a.������� = 'Z' And c.�������� = '2' And a.Id = n_ҽ��id;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.סԺ��, a.����, a.�Ա�, a.����, a.�ѱ�, a.��������, a.����, a.����, a.ѧ��, a.����״��, a.ְҵ, a.���, a.���֤��, a.�����ص�, a.��ͥ��ַ,
           a.��ͥ��ַ�ʱ�, a.��ͥ�绰, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.������λ, a.��ͬ��λid, a.��λ�绰, a.��λ�ʱ�,
           a.��λ������, a.��λ�ʺ�, a.������, a.������, a.��������, a.����, a.����, a.ҽ�Ƹ��ʽ, a.����
    From ������Ϣ A
    Where a.����id = v_����id;
  r_Pati c_Pati%RowType;

  n_Type   Number;
  n_�Һ�id ����ҽ����¼.Id%Type;
  n_����id ����ҽ����¼.Id%Type;
  n_����id ����ҽ����¼.Id%Type;
  v_����   ������ҳ.��Ժ����%Type;

  n_����id ������ҳ.����id%Type;
  v_No     ���˹Һż�¼.No%Type;
  n_Count  Number;

  v_��Ժ��ʽ ������ҳ.��Ժ��ʽ%Type;
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Temp     Varchar2(4000);

Begin
  Select Extractvalue(Value(A), 'IN/TYPE'), Extractvalue(Value(A), 'IN/GHID') As �Һ�id,
         Extractvalue(Value(A), 'IN/RYKSID') As ��Ժ����id, Extractvalue(Value(A), 'IN/RYBQID') As ��Ժ����id,
         Extractvalue(Value(A), 'IN/CH') As ����, Extractvalue(Value(A), 'IN/CZYBH') As ���,
         Extractvalue(Value(A), 'IN/CZYXM') As ����, Extractvalue(Value(A), 'IN/YZID') As ҽ��id
  Into n_Type, n_�Һ�id, n_����id, n_����id, v_����, v_��Ա���, v_��Ա����, n_ҽ��id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If n_Type = 1 Then
    --סԺԤԼ�Ǽ�
    Select a.����id, a.No, Decode(a.����, 1, '����', Null)
    Into n_����id, v_No, v_��Ժ��ʽ
    From ���˹Һż�¼ A
    Where a.Id = n_�Һ�id;
  
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
  
    If r_Advice.������־ = 1 Then
      v_��Ժ��ʽ := '����';
    End If;
  
    Open c_Pati(n_����id);
    Fetch c_Pati
      Into r_Pati;
  
    --��ǰ������Ա
    If v_��Ա��� Is Null Or v_��Ա���� Is Null Then
      v_Temp     := Zl_Identity;
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
      v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    End If;
  
    --ɾ�����ۼ�¼��סԺԤԼ��¼���ܲ���
    Begin
      Select Count(1) Into n_Count From ������ҳ Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) > 0 Then
      Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0, 0, 0);
      n_Count := 0;
    End If;
  
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ������ҳ Where ����id = r_Advice.����id And ��Ժ���� Is Null;
    End If;
    If n_Count = 0 Then
      Select Count(1)
      Into n_Count
      From ������ҳ
      Where ����id = r_Advice.����id And (��Ժ���� >= r_Advice.��ʼִ��ʱ�� Or ��Ժ���� >= r_Advice.��ʼִ��ʱ��);
    End If;
  
    If n_Count = 0 Then
      Zl_��Ժ������ҳ_Insert(1, 0, r_Pati.����id, r_Pati.סԺ��, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�, r_Pati.��������,
                       r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���, r_Pati.���֤��, r_Pati.�����ص�,
                       r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ, r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ,
                       r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ, r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������,
                       r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������, r_Pati.��������, n_����id, Null, Null, v_��Ժ��ʽ, Null, Null,
                       r_Advice.����ҽ��, r_Pati.����, r_Pati.����, r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null,
                       Null, Null, Null, r_Pati.����, v_��Ա���, v_��Ա����, 0, Null, Null, 0, Null, Null, Null, Null, Null, Null,
                       Null, n_�Һ�id);
    End If;
  
    --���²����ʹ���
    Update ������ҳ
    Set ��Ժ���� = v_����, ��Ժ���� = v_����, ��Ժ����id = n_����id, ��ǰ����id = n_����id
    Where ����id = r_Pati.����id;
  
    --����λ����ռ��
    Update ��λ״����¼
    Set ״̬ = 'ռ��', ����id = r_Pati.����id, ����id = Decode(����, 1, n_����id, ����id)
    Where ����id = n_����id And ���� = v_����;
  Else
    --ȡ���Ǽ�
  
    Select b.����id, b.��Ժ����id, b.��Ժ����, b.��Ժ����id
    Into n_����id, n_����id, v_����, n_����id
    From ������ҳ B
    Where b.�Һ�id = n_�Һ�id;
  
    --���²����ʹ���
    Update ������ҳ
    Set ��Ժ���� = Null, ��Ժ���� = Null, ��Ժ����id = Null, ��ǰ����id = Null
    Where ����id = r_Pati.����id;
  
    --����λ����ȡ��ռ��
    Update ��λ״����¼
    Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id)
    Where ����id = n_����id And ���� = v_����;
  
    Zl_��Ժ������ҳ_Delete(n_����id, 0);
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Outpatireg;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0018' Where ���=&n_System;
Commit;
