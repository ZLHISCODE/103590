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
--136629:����,2019-01-16,����ϵͳ�������ڿ���סԺ�����Զ����Ź���
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��)
Select zlParameters_ID.Nextval,&n_System,-Null,-Null,-Null,-Null,-Null,A.* From (
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0 Union All 
  Select 0,0,320,'�Զ����ű���������','0','0','���ô˲����󣬱��뿪��������ִ�п���һ�²��Զ�����','0-������,1-����','����ϵͳ������סԺ�����Զ����ϡ��󣬴˲�������������','�����ڿ������ұ�����ִ�п���һ�µ����',Null From Dual Union All 
Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0) A;


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--134441:���ϴ�,2019-01-15,�Һż����Ŀ�Ƿ�һ��
Create Or Replace Procedure Zl_����ԤԼ�Һż�¼_Update
(
  ���ݺ�_In     ������ü�¼.No%Type,
  ���_In       ������ü�¼.���%Type,
  �۸񸸺�_In   ������ü�¼.�۸񸸺�%Type,
  ��������_In   ������ü�¼.��������%Type,
  �շ����_In   ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In ������ü�¼.�շ�ϸĿid%Type,
  ����_In       ������ü�¼.����%Type,
  ��׼����_In   ������ü�¼.��׼����%Type,
  ������Ŀid_In ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In   ������ü�¼.�վݷ�Ŀ%Type,
  Ӧ�ս��_In   ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In   ������ü�¼.ʵ�ս��%Type,
  ������_In     Number, --������¼�Ƿ���������
  ���մ���id_In ������ü�¼.���մ���id%Type,
  ������Ŀ��_In ������ü�¼.������Ŀ��%Type,
  ͳ����_In   ������ü�¼.ͳ����%Type,
  ���ձ���_In   ������ü�¼.���ձ���%Type,
  ���˿���id_In ������ü�¼.���˿���id%Type,
  ִ�в���id_In ������ü�¼.ִ�в���id%Type,
  ժҪ_In       ������ü�¼.ժҪ%Type := Null,
  �Ƿ�Һ���_In Number := 0
) As
  v_����id ������ü�¼.Id%Type;
  v_Error  Varchar2(255);
  Err_Custom Exception;
  Cursor c_���� Is
    Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�,
           �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���
    From ������ü�¼
    Where NO = ���ݺ�_In And ��¼���� = 4 And ��� = 1 And ��¼״̬ = 0;
Begin

  If Nvl(���_In, 1) = 1 Then
    --��һ����¼,ֻ��������
    Update ������ü�¼
    Set �۸񸸺� = Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), �������� = Decode(��������_In, 0, Null, ��������_In), ���ӱ�־ = ������_In,
        �շ���� = �շ����_In, �շ�ϸĿid = �շ�ϸĿid_In, ������Ŀid = ������Ŀid_In, �վݷ�Ŀ = �վݷ�Ŀ_In, ���� = 1, ���� = ����_In, ��׼���� = ��׼����_In,
        Ӧ�ս�� = Ӧ�ս��_In, ʵ�ս�� = ʵ�ս��_In, ���մ���id = ���մ���id_In, ������Ŀ�� = ������Ŀ��_In, ���ձ��� = ���ձ���_In, ͳ���� = ͳ����_In,
        ���˿���id =  Decode(�Ƿ�Һ���_In, 1, ���˿���id, ���˿���id_In), ִ�в���id = Decode(�Ƿ�Һ���_In, 1, ִ�в���id, ִ�в���id_In), ժҪ = Nvl(ժҪ_In, ժҪ)
    Where NO = ���ݺ�_In And ��� = 1 And ��¼״̬ = 0 And ��¼���� = 4;
    --ɾ����Ŵ���1������;
    Delete ������ü�¼ Where NO = ���ݺ�_In And ��� > 1 And ��¼���� = 4;
  Else
    --��������
    If Nvl(������_In, 0) <> 3 Then
      Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual; --Ӧ��ͨ������õ�
      For r_���� In c_���� Loop
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����)
        Values
          (v_����id, 4, 0, ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, r_����.ʵ��Ʊ��, 1, r_����.�Ӱ��־, ������_In,
           r_����.��ҩ����, r_����.����id, r_����.��ʶ��, r_����.���ʽ, r_����.����, r_����.�Ա�, r_����.����, r_����.�ѱ�, ���˿���id_In, �շ����_In, r_����.���㵥λ,
           �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In, Null, Null, 0, r_����.��������id, r_����.����Ա����,
           r_����.����Ա����, ִ�в���id_In, r_����.ִ����, r_����.����Ա���, r_����.����Ա����, r_����.����ʱ��, r_����.�Ǽ�ʱ��, ���մ���id_In, ������Ŀ��_In, ���ձ���_In,
           ͳ����_In, Nvl(ժҪ_In, r_����.ժҪ), r_����.����);
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ԤԼ�Һż�¼_Update;
/

--134441:���ϴ�,2019-01-15,�Һż����Ŀ�Ƿ�һ��
Create Or Replace Procedure Zl_Third_Registercheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:HIS�Һż��
  --���:Xml_In:
  --<IN>
  --  <BRID>1</BRID>                    //����ID
  --  <XM>����</XM>                     //����
  --  <SFZH>510221197008184710</SFZH>   //���֤��
  --  <HM>0100</HM>                     //����
  --  <CZJLID>100</CZJLID>              //�����¼ID,�ƻ��Ű�ģʽ���Բ���
  --  <GHSJ>2016-08-10 09:52:00</GHSJ>  //�Һ�ʱ��
  --  <KSID>1</KSID>                    //����ID
  --  <YSXM>����</YSXM>                 //ҽ������
  --  <GHXMID>1</GHXMID>                 //�Һ�����Ŀ������ʱ�����
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <ERROR><MSG></MSG></ERROR> //Ϊ�ձ�ʾ���ɹ�
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_����id         ������Ϣ.����id%Type;
  v_����           ������Ϣ.����%Type;
  v_���֤��       ������Ϣ.���֤��%Type;
  v_����           �ҺŰ���.����%Type;
  n_��ĿID         �ҺŰ���.��ĿID%Type;
  n_�����¼id     Number(18);
  d_����ʱ��       ���˹Һż�¼.����ʱ��%Type;
  v_Para           Varchar2(500);
  d_����ʱ��       Date;
  n_�Һ�ģʽ       Number(3);
  n_ͬ���޺���     Number;
  n_ͬ����Լ��     Number;
  n_ͬԴ�޺���     Number;
  n_���˹Һſ����� Number;
  n_����ԤԼ������ Number;
  n_ר�ҺŹҺ����� Number;
  n_ר�Һ�ԤԼ���� Number;
  n_Exists         Number;
  n_Count          Number;
  n_����id         ���˹Һż�¼.ִ�в���id%Type;
  v_ҽ������       ���˹Һż�¼.ִ����%Type;
  v_�Ա�           ������Ϣ.�Ա�%Type;
  v_����           ������Ϣ.����%Type;
  n_��Լ����       Number;
  v_Checkresult    Varchar2(500);
  v_Temp           Varchar2(32767); --��ʱXML
  x_Templet        Xmltype; --ģ��XML
  v_Err_Msg        Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/CZJLID'),
         To_Date(Extractvalue(Value(A), 'IN/GHSJ'), 'yyyy-mm-dd hh24:mi:ss'),
         To_Number(Extractvalue(Value(A), 'IN/KSID')), Extractvalue(Value(A), 'IN/YSXM'),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'),
         To_Number(Extractvalue(Value(A), 'IN/GHXMID'))
  Into n_����id, v_����, n_�����¼id, d_����ʱ��, n_����id, v_ҽ������, v_���֤��, v_����, n_��ĿID
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����';
    Raise Err_Item;
  End If;

  v_Para := zl_GetSysParameter(256);
  If v_Para Is Not Null Then
    n_�Һ�ģʽ := Substr(v_Para, 1, 1);
    Begin
      d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  
    If Sysdate - 10 > Nvl(d_����ʱ��, Sysdate - 30) Then
      If n_�Һ�ģʽ = 1 And Nvl(d_����ʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
        v_Temp := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    Else
      If n_�Һ�ģʽ = 1 And Nvl(d_����ʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
        Begin
          Select a.Id
          Into n_�����¼id
          From �ٴ������¼ A, �ٴ������Դ B
          Where a.��Դid = b.Id And b.���� = v_���� And Nvl(d_����ʱ��, Sysdate) Between a.��ʼʱ�� And a.��ֹʱ��;
        Exception
          When Others Then
            v_Temp := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
            v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
            Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
            Xml_Out := x_Templet;
            Return;
        End;
      End If;
    End If;
  End If;

  If n_�����¼id Is Not Null Then
    Select �Ա�, ���� Into v_�Ա�, v_���� From ������Ϣ Where ����id = n_����id And Rownum < 2;
    v_Checkresult := Zl_�ٴ���������_Check(n_�����¼id, v_����, v_�Ա�);
    If Substr(Nvl(v_Checkresult, '0'), 1, 1) <> '0' Then
      v_Temp := '���˲����øñ��ű�,���飡';
      v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Xml_Out := x_Templet;
      Return;
    End If;
    
    For C_�Ű� In (Select a.����, b.����id, a.��Ŀid From �ٴ������Դ a, �ٴ������¼ b
                    where a.id = b.��Դid And b.Id = n_�����¼id) loop
      v_Temp  := Null;
      n_Count := 1;
      if v_���� <> C_�Ű�.���� then
        v_Temp := '�Һ���Ϣ�ĺ������,���飡';
      Elsif n_����id <> C_�Ű�.����id then
        v_Temp := '�Һ���Ϣ�Ŀ��Ҵ���,���飡';
      Elsif n_��Ŀid <> C_�Ű�.��Ŀid And Nvl(n_��Ŀid, 0) <> 0 then
        v_Temp := '�Һ���Ϣ���շ���Ŀ����,���飡';
      end IF;
      IF v_Temp is not null Then
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End if;
    End loop;

    IF NVL(n_Count, 0) <> 1 Then
      v_Temp := '�Һ���Ϣ����,�����ԣ�';
      v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Xml_Out := x_Templet;
      Return;
    End IF;
  End If;

  If Trunc(Sysdate) > Trunc(d_����ʱ��) Then
    v_Temp := '���ܹ���ǰ�ĺ�(' || To_Char(d_����ʱ��, 'yyyy-mm-dd') || ')��';
    v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    v_Temp := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
    v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xml_Out := x_Templet;
    Return;
  End If;

  v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
  n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
  n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
  n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
  n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
  n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
  n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
  n_ͬԴ�޺���     := To_Number(Nvl(zl_GetSysParameter('����ͬһ��Դ�޹�N����', 1111), '0'));

  If Trunc(Sysdate) <> Trunc(d_����ʱ��) Then
    If Nvl(n_����ԤԼ������, 0) <> 0 Then
      n_��Լ���� := 0;
      For c_Chkitem In (Select Distinct ִ�в���id
                        From ���˹Һż�¼
                        Where ����id = n_����id And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(d_����ʱ��) And
                              Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
        n_��Լ���� := n_��Լ���� + 1;
      End Loop;
      If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
        v_Temp := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
    If Nvl(n_ͬ����Լ��, 0) <> 0 Then
      Select Count(1)
      Into n_Count
      From ���˹Һż�¼
      Where ����id = n_����id And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(d_����ʱ��) And
            Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
      If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
        v_Temp := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
  Else
    If Nvl(n_���˹Һſ�����, 0) <> 0 Then
      n_��Լ���� := 0;
      For c_Chkitem In (Select Distinct ִ�в���id
                        From ���˹Һż�¼
                        Where ����id = n_����id And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(d_����ʱ��) And
                              Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
        n_��Լ���� := n_��Լ���� + 1;
      End Loop;
      If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
        v_Temp := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
    If Nvl(n_ͬ���޺���, 0) <> 0 Then
      Select Count(1)
      Into n_Count
      From ���˹Һż�¼
      Where ����id = n_����id And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(d_����ʱ��) And
            Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
      If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
        v_Temp := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
        v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        Xml_Out := x_Templet;
        Return;
      End If;
    End If;
  End If;

  If Nvl(n_ͬԴ�޺���, 0) <> 0 Then
    If n_�����¼id Is Null Then
      Select Count(1)
      Into n_Count
      From ���˹Һż�¼
      Where ����id = n_����id And ��¼״̬ = 1 And ��¼���� In (1, 2) And ����ʱ�� Between Trunc(d_����ʱ��) And
            Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And �ű� = v_����;
    Else
      Select Count(1)
      Into n_Count
      From ���˹Һż�¼
      Where ����id = n_����id And ��¼״̬ = 1 And ��¼���� In (1, 2) And �����¼id = n_�����¼id;
    End If;
    If n_Count >= Nvl(n_ͬԴ�޺���, 0) And Nvl(n_ͬԴ�޺���, 0) > 0 Then
      v_Temp := 'ͬһ���������ͬʱ��(ԤԼ)[' || Nvl(n_ͬԴ�޺���, 0) || ']����ͬ�ű�ĺ�,�����ٹҺ�(ԤԼ)��';
      v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      Xml_Out := x_Templet;
      Return;
    End If;
  End If;

  If Trunc(Sysdate) = Trunc(d_����ʱ��) Then
    --�Һ�
    If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 And v_ҽ������ Is Not Null Then
      If n_�����¼id Is Null Then
        --�޳����¼��Ӧ
        Begin
          Select Count(1)
          Into n_Exists
          From ���˹Һż�¼
          Where ����id = n_����id And �ű� = v_���� And ����ʱ�� Between Trunc(d_����ʱ��) And Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And
                ��¼״̬ = 1 And ��¼���� = 1;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_ר�ҺŹҺ����� Then
          v_Temp := '�ò����Ѿ��������ŹҺ�����,�����ٴιҺţ�';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      Else
        --��Ӧ�����¼
        Begin
          Select Count(1)
          Into n_Exists
          From ���˹Һż�¼
          Where ����id = n_����id And �����¼id = n_�����¼id And ��¼״̬ = 1 And ��¼���� = 1;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_ר�ҺŹҺ����� Then
          v_Temp := '�ò����Ѿ��������ŹҺ�����,�����ٴιҺţ�';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      End If;
    End If;
  Else
    --ԤԼ
    If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 And v_ҽ������ Is Not Null Then
      If n_�����¼id Is Null Then
        --�޳����¼��Ӧ
        Begin
          Select Count(1)
          Into n_Exists
          From ���˹Һż�¼
          Where ����id = n_����id And �ű� = v_���� And ����ʱ�� Between Trunc(d_����ʱ��) And Trunc(d_����ʱ��) + 1 - 1 / 24 / 60 / 60 And
                ��¼״̬ = 1 And ��¼���� = 2;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_ר�Һ�ԤԼ���� Then
          v_Temp := '�ò����Ѿ���������ԤԼ����,�����ٴ�ԤԼ��';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      Else
        --��Ӧ�����¼
        Begin
          Select Count(1)
          Into n_Exists
          From ���˹Һż�¼
          Where ����id = n_����id And �����¼id = n_�����¼id And ��¼״̬ = 1 And ��¼���� = 2;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists >= n_ר�Һ�ԤԼ���� Then
          v_Temp := '�ò����Ѿ���������ԤԼ����,�����ٴ�ԤԼ��';
          v_Temp := '<ERROR><MSG>' || v_Temp || '</MSG></ERROR>';
          Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
          Xml_Out := x_Templet;
          Return;
        End If;
      End If;
    End If;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registercheck;
/

--134969:���ϴ�,2019-01-14,Ԥ��֧�����
Create Or Replace Procedure Zl_���˹Һż�¼_Insert
(
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0,
  ������������_In  Number := 0,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null,
  ���½������_In  Number := 1
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  --     ���½������_In:0-��zl_��Ա�ɿ����_Update �и��� 1-�ڱ������и���
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, Id, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, Null)) As �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(Ԥ�����, 2) = 1 Having
     Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By No, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), �տ�ʱ��;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ����
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;

  n_����id        ������ü�¼.Id%Type;
  n_�������      ����Ԥ����¼.���%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_��ǰ���      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_��id           ����ɿ����.Id%Type;
  n_�����         ������Ϣ.�����%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  d_�Ŷ�ʱ��       Date;
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type := 0;
  v_����           �ҺŰ�������.������Ŀ%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;

  n_�ҳ��������� Number(4) := 0;
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_Temp           Varchar2(3000);
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   ʱ���.��ʼʱ��%Type;
  d_������ʱ��   ʱ���.��ֹʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  n_��ʱ����ʾ     Number(3);
  d_����ʱ��       Date;
Begin
  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);

  If �ѱ�_In Is Null Then
    Begin
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
        Raise Err_Item;
    End;
  Else
    v_�ѱ� := �ѱ�_In;
    If Nvl(�������˷ѱ�_In, 0) = 1 Then
      Begin
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(������������_In, 0) = 1 Then
    Begin
      Update ������Ϣ Set ���� = ����_In Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
        Raise Err_Item;
    End;
  End If;

  If �����_In Is Not Null Then
    Begin
      Select Nvl(�����, 0) Into n_����� From ������Ϣ Where ����id = ����id_In;
    Exception
      When Others Then
        n_����� := 0;
    End;
    If n_����� = 0 Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In;
    End If;
  End If;

  Begin
    Delete From �Һ����״̬
    Where ���� = �ű�_In And ���� = ����ʱ��_In And ��� = ����_In And ״̬ = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;
  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
    If d_����ʱ�� Is Not Null Then
      If ����ʱ��_In > d_����ʱ�� Then
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
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
  End If;

  n_��� := ����_In;
  Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
  Into v_����
  From Dual;

  --�ҺŻ�ȡ����
  Begin
    Select a.Id, a.��ſ���, Nvl(b.�޺���, 0), Nvl(b.��Լ��, 0)
    Into n_����id, n_��ſ���, n_�޺���, n_��Լ��
    From �ҺŰ��� A, �ҺŰ������� B
    Where a.Id = b.����id(+) And b.������Ŀ(+) = v_���� And a.���� = �ű�_In;
  
  Exception
    When Others Then
      n_����id := -1;
  End;

  --����ǲ����ѻ��ߺű�Ϊ��ʱ�����
  If Nvl(������_In, 0) = 0 Or �ű�_In Is Not Null Then
    If n_����id = -1 Then
      v_Err_Msg := '������Ӧ�ĹҺŰ�������,����';
      Raise Err_Item;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
    --���Ȼ�ȡ�ƻ�
    Begin
      Select ID
      Into n_�ƻ�id
      From �ҺŰ��żƻ�
      Where ����id = n_����id And ���ʱ�� Is Not Null And
            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
            (Select Max(a.��Чʱ��) As ��Ч
             From �ҺŰ��żƻ� A
             Where a.���ʱ�� Is Not Null And ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                   a.ʧЧʱ�� And a.����id = n_����id) And
            ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And ʧЧʱ��;
    
    Exception
      When Others Then
        n_�ƻ�id := 0;
    End;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Begin
        --��ȡ�ƻ�������
        Select a.Id, a.��ſ���, Nvl(b.�޺���, 0) As �޺���, Nvl(b.��Լ��, 0) As ��Լ��
        Into n_�ƻ�id, n_��ſ���, n_�޺���, n_��Լ��
        From �ҺŰ��żƻ� A, �Һżƻ����� B
        Where a.���� = �ű�_In And a.Id = n_�ƻ�id And a.���ʱ�� Is Not Null And a.Id = b.�ƻ�id(+) And b.������Ŀ(+) = v_����;
      Exception
        When Others Then
          v_Err_Msg := '������Ӧ�ĹҺŰ��Ż�ƻ�����,����';
          Raise Err_Item;
      End;
    End If;
  End If;

  --��ȡ�Ƿ��ʱ��
  Begin
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum <= 1;
      Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum <= 1;
      Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null And ���_In = 1 Then
    --����Ƿ��ģʽ�ҺŰ���
    Select To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_��鿪ʼʱ��, d_������ʱ��
    From ʱ���
    Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
    If d_��鿪ʼʱ�� > d_������ʱ�� Then
      d_������ʱ�� := d_������ʱ�� + 1;
    End If;
    If d_��鿪ʼʱ�� < d_����ʱ�� And d_������ʱ�� > d_����ʱ�� Then
      --��ȡ�����¼id
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = �ű�_In And �ϰ�ʱ�� = v_ʱ��� And ����ʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
    End If;
  End If;

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    --����ʱ��_in>Sysdate ����ʱ��>����ʱ��ʱ��--����_in is null
    Begin
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And Nvl(��������, 0) <> 0;
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_׷�Ӻ� := 0;
    End;
  End If;
  d_ʱ��ʱ�� := ����ʱ��_In;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 Then
    --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
    Begin
      Select Nvl(���, 0),
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
      Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And
            (���, ����id, ����) In (Select Nvl(Max(���), -1), ����id, ����
                               From �ҺŰ���ʱ��
                               Where ����id = n_����id And ���� = v_���� And
                                     Decode(��������_In + n_׷�Ӻ�, 0, To_Char(����ʱ��_In, 'hh24:mi'),
                                            To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) =
                                     To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
                               Group By ����id, ����);
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ�� > 0 Then
    --ԤԼ��,ȡ�ƻ�
    Begin
      If Nvl(n_�ƻ�id, 0) = 0 Then
        --û�ƻ���Ч,ȡ���ŵ�����
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ҺŰ���ʱ�� C
        Where ����id = n_����id And ���� = v_���� And
              (���, ����id, ����) In
              (Select Nvl(Max(c.���), -1), ����id, ����
               From �ҺŰ���ʱ�� C
               Where ����id = n_����id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By ����id, ����);
      Else
        --�мƻ���Чȡ�ƻ�
        --û��Ч�������ǴӹҺżƻ�ʱ�β�ѯ
        Select Nvl(���, -1),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��,
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �Һżƻ�ʱ�� C
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And
              (���, �ƻ�id, ����) In
              (Select Nvl(Max(c.���), -1), �ƻ�id, ����
               From �Һżƻ�ʱ�� C
               Where �ƻ�id = n_�ƻ�id And c.���� = v_���� And
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'),
                            To_Char(����ʱ��_In, 'hh24:mi')) =
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')
               Group By �ƻ�id, ����);
      End If;
    Exception
      When Others Then
        n_ʱ����� := -1;
        n_��ʱ��   := 0;
        d_ʱ��ʱ�� := ����ʱ��_In;
        n_ʱ���޺� := 0;
        n_ʱ����Լ := 0;
    End;
  End If;

  If ���_In = 1 Then
  
    --��ȡ��ǰδʹ�õ����
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>
      Begin
        --������
        If �˺�����_In = 1 Then
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0))
          Into n_�������, n_��������, n_��Լ��
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
      End;
      If n_��� Is Null Then
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end>
    
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 Then
        --<����ԤԼ�Һ�-->
        Begin
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.���� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
          Into n_�ѹ���, n_��Լ��
          From �Һ����״̬ A
          Where a.���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And
                ״̬ Not In (4, 5);
        Exception
          When Others Then
            n_�ѹ��� := 0;
            n_��Լ�� := 0;
        End;
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
        If n_�޺��� <= n_�ѹ��� Then
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
          Raise Err_Item;
        End If;
      End If;
    
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷��
    
      --��ȡ����ҳ���������
      If �˺�����_In = 1 Then
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ Not In (4, 5);
      Else
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �Һ����״̬ A
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ <> 5;
      End If;
    
      --�������
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_��Լ�� + 1;
      If n_��� <= Nvl(n_�ҳ���������, 0) Then
        n_��� := Nvl(n_�ҳ���������, 0) + 1;
      End If;
    
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        If �˺�����_In = 1 Then
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5);
        Else
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0))
          Into n_�������, n_�ѹ���, n_��������
          From �Һ����״̬
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5;
        End If;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
          n_�ѹ���   := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111));
        n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111));
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ����), 1, 1, 0))
            Into n_ʧЧ��
            From �Һ����״̬
            Where ���� = �ű�_In And ���� Between Trunc(Sysdate) And Sysdate And Nvl(ԤԼ, 0) = 1 And ״̬ = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      Begin
        Select �ѹ���, ��Լ��
        Into n_��������, n_��Լ��
        From ���˹ҺŻ���
        Where ���� = Trunc(����ʱ��_In) And ���� = �ű�_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  End If;

  --���¹Һ����״̬
  If ���_In = 1 And Not n_��� Is Null Then
    If n_��ʱ�� = 1 Then
      d_���ʱ�� := ����ʱ��_In;
    Else
      d_���ʱ�� := Trunc(����ʱ��_In);
    End If;
    --������ŵĴ���
    Begin
      Select ����Ա����, ������
      Into v_��Ų���Ա, v_��Ż�����
      From �Һ����״̬
      Where ״̬ = 5 And ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
      n_���� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_����       := 0;
    End;
    If n_���� = 0 Then
      Update �Һ����״̬
      Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
      Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 3 And ����Ա���� = ����Ա����_In;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_��ʱ��, 0) = 0 Or Nvl(ԤԼ�Һ�_In, 0) = 1 Or (Nvl(n_��ſ���, 0) = 0 And Nvl(����_In, 0) = 0) Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
            Values
              (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          Elsif Nvl(n_��ʱ��, 0) > 0 Then
            --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
            Update �Һ����״̬
            Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In, ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
            Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 2;
            If Sql%NotFound Then
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
              Values
                (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate);
            End If;
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        Update �Һ����״̬
        Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate
        Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 5 And ����Ա���� = ����Ա����_In And ������ = v_������;
        If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
          Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���;
        End If;
      End If;
    End If;
  End If;

  If n_�����¼id Is Not Null Then
    Update �ٴ�������ſ���
    Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
    Where ��¼id = n_�����¼id And ��� = n_���;
    If ԤԼ�Һ�_In = 1 Then
      Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
    Else
      If ԤԼ����_In = 1 Then
        Update �ٴ������¼
        Set ��Լ�� = ��Լ�� + 1, �ѹ��� = �ѹ��� + 1, �����ѽ��� = �����ѽ��� + 1
        Where ID = n_�����¼id;
      Else
        Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
      End If;
    End If;
  End If;

  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, Decode(�շѵ�_In, Null, ժҪ_In, '����:' || �շѵ�_In), ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
  
    If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), �Ǽ�ʱ��_In,
         ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    
      If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
        Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, �ֽ�֧��_In, n_Ԥ��id, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In);
      End If;
    End If;
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
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
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2);
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
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 And Nvl(���½������_In, 1) = 1 Then
      If Nvl(�ֽ�֧��_In, 0) <> 0 Then
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
          Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End If;
    
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �շѵ�)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �շѵ�_In);
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(ִ�в���id_In, 0)));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
        n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then
          n_��ʱ����ʾ := 1;
        Else
          n_��ʱ����ʾ := Null;
        End If;
      
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         n_��ʱ����ʾ, v_�Ŷ����);
      
        --�Һ������Ŷ�
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
        End If;
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
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
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) >= Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_Insert;
/

--134441:���ϴ�,2019-01-15,�Һż����Ŀ�Ƿ�һ��
--134969:���ϴ�,2019-01-14,Ԥ��֧�����
Create Or Replace Procedure Zl_���˹Һż�¼_����_Insert
(
  �����¼id_In    �ٴ������¼.Id%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      Varchar2,
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0,
  ԤԼ˳���_In    �ٴ�������ſ���.ԤԼ˳���%Type := Null,
  ������������_In  Number := 0,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null,
  ���½������_In  Number := 1 --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, No, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, Id, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, Null)) As �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select /*+cardinality(d,10)*/
                                        d.Column_Value
                                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(Ԥ�����, 2) = 1 Having
     Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By No, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), �տ�ʱ��;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ԭʼ��ʱ��   Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ����
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;

  n_����id        ������ü�¼.Id%Type;
  n_�������      ����Ԥ����¼.���%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_��ǰ���      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_��id           ����ɿ����.Id%Type;
  n_�����         ������Ϣ.�����%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  d_�Ŷ�ʱ��       Date;
  v_���㷽ʽ��¼   Varchar2(1000);
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  v_���㷽ʽ       ���㷽ʽ.����%Type;
  v_��������       Varchar2(1000);
  v_��ǰ����       Varchar2(200);
  v_�������       ����Ԥ����¼.�������%Type;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_��������־     Number(2);
  n_ԤԼ˳���     �ٴ�������ſ���.ԤԼ˳���%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;
  n_Exists         Number;
  n_�ҳ��������� Number(4) := 0;
  n_��ʱ����ʾ     Number(3);
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_״̬           �ٴ�������ſ���.�Һ�״̬%Type;
Begin
  --��¼�����ж�
  If Nvl(���_In, 0) = 1 Then
    If �����¼id_In Is Not Null Then
      Begin
        Select 1
        Into n_Exists
        From �ٴ������¼ a, �ٴ������Դ b
        Where a.Id = �����¼id_In And a.��Դid = b.Id And b.���� = �ű�_In And a.����id = ִ�в���id_In And Nvl(a.�Ƿ񷢲�, 0) = 1 And
              Nvl(a.�Ƿ�����, 0) = 0;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�������¼����������¼�Ƿ���ڻ�������';
          Raise Err_Item;
      End;
    End If;
  End if;

  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);

  If �ѱ�_In Is Null Then
    Begin
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
        Raise Err_Item;
    End;
  Else
    v_�ѱ� := �ѱ�_In;
    If Nvl(�������˷ѱ�_In, 0) = 1 Then
      Begin
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If Nvl(������������_In, 0) = 1 Then
    Begin
      Update ������Ϣ Set ���� = ����_In Where ����id = ����id_In;
    Exception
      When Others Then
        v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�';
        Raise Err_Item;
    End;
  End If;

  If �����_In Is Not Null Then
    Begin
      Select Nvl(�����, 0) Into n_����� From ������Ϣ Where ����id = ����id_In;
    Exception
      When Others Then
        n_����� := 0;
    End;
    If n_����� = 0 Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In;
    End If;
  End If;

  Begin
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0
    Where ��¼id = �����¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
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
  End If;

  n_��� := ����_In;

  --��ȡ�Ƿ��ʱ��
  Begin
    Select Nvl(�Ƿ��ʱ��, 0), Nvl(�Ƿ���ſ���, 0), �޺���, ��Լ��
    Into n_��ʱ��, n_��ſ���, n_�޺���, n_��Լ��
    From �ٴ������¼
    Where ID = �����¼id_In;
    n_ԭʼ��ʱ�� := n_��ʱ��;
  Exception
    When Others Then
      n_��ʱ��     := 0;
      n_ԭʼ��ʱ�� := n_��ʱ��;
      n_��ſ���   := 0;
      n_�޺���     := Null;
      n_��Լ��     := Null;
  End;

  If n_��� Is Null And n_��ʱ�� = 1 And n_��ſ��� = 0 Then
    Begin
      Select ���
      Into n_���
      From �ٴ�������ſ���
      Where ��¼id = �����¼id_In And ��ʼʱ�� = ����ʱ��_In And Rownum < 2;
    Exception
      When Others Then
        n_��� := Null;
    End;
  End If;

  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then
    Begin
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'))
      Into d_������ʱ��
      From �ٴ�������ſ���
      Where ��¼id = �����¼id_In And Nvl(����, 0) <> 0;
    
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
                 When -1 Then
                  0
                 Else
                  1
               End;
    Exception
      When Others Then
        n_׷�Ӻ� := 0;
    End;
  End If;
  d_ʱ��ʱ�� := ����ʱ��_In;

  If ���_In = 1 And n_��ʱ�� > 0 Then
    If Nvl(n_��ſ���, 0) = 1 Then
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_���;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := ����ʱ��_In;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    Else
      --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� Is Null;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := ����ʱ��_In;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    End If;
  End If;

  If ���_In = 1 Then
    --��ȡ��ǰδʹ�õ����
    If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
      n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111));
      n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111));
    End If;
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>
      Begin
        --������
        Select Count(1) Into n_�������� From ���˹Һż�¼ Where �����¼id = �����¼id_In And ��¼״̬ = 1;
        Select Max(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
      End;
      Begin
        --������
        Select Sum(Nvl(����, 0))
        
        Into n_��Լ��
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) = 2;
      Exception
        When Others Then
          n_��Լ�� := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ԤԼʱ��), 1, 1, 0))
            Into n_ʧЧ��
            From ���˹Һż�¼
            Where �����¼id = �����¼id_In And ��¼״̬ = 1 And ��¼���� = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If n_ԭʼ��ʱ�� = 0 Then
        Select Min(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) = 0;
        If n_��� Is Null Then
          n_��� := Nvl(n_�������, 0);
        End If;
        IF nvl(n_���,0)=0 THEN 
          Select Nvl(Max(���), 0) + 1 Into n_��� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
        END IF;
      Else
        Select Max(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
        If n_��� Is Null Then
          n_��� := Nvl(n_�������, 0) + 1;
        End If;
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end>
    
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 Then
        --<����ԤԼ�Һ�-->
        Begin
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.��ʼʱ�� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
          Into n_�ѹ���, n_��Լ��
          From �ٴ�������ſ��� A
          Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
        Exception
          When Others Then
            n_�ѹ��� := 0;
            n_��Լ�� := 0;
        End;
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
        If n_�޺��� <= n_�ѹ��� Then
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
          Raise Err_Item;
        End If;
      End If;
    
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷��
    
      --��ȡ����ҳ���������
      Select Nvl(Max(���), 0)
      Into n_�ҳ���������
      From �ٴ�������ſ��� A
      Where ��¼id = �����¼id_In And ԤԼ˳��� Is Null And �Һ�״̬ Not In (0, 5);
      If ԤԼ˳���_In Is Not Null Then
        n_ԤԼ˳��� := ԤԼ˳���_In;
      Else
        Begin
          Select Nvl(Max(ԤԼ˳���), 0) + 1
          Into n_ԤԼ˳���
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Not Null;
        Exception
          When Others Then
            n_ԤԼ˳��� := Null;
        End;
      End If;
      --�������
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_ԤԼ˳���;
      If n_ԤԼ˳��� Is Null Then
        n_��� := Nvl(n_�ҳ���������, 0) + 1;
      End If;
    
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(��ʼʱ�� - d_ʱ��ʱ��), 0, 1, 0))
        Into n_�������, n_�ѹ���, n_��������
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
      Exception
        When Others Then
          n_������� := 0;
          n_�������� := 0;
          n_�ѹ���   := 0;
      End;
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
          Begin
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ��ʼʱ��), 1, 1, 0))
            Into n_ʧЧ��
            From �ٴ�������ſ���
            Where ��¼id = �����¼id_In And ��ʼʱ�� Between Trunc(Sysdate) And Sysdate And Nvl(�Һ�״̬, 0) = 2;
          Exception
            When Others Then
              n_ʧЧ�� := 0;
          End;
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        
        Else
          --�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      Begin
        Select �ѹ���, ��Լ�� Into n_��������, n_��Լ�� From �ٴ������¼ Where ID = �����¼id_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_��Լ��   := 0;
      End;
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  End If;

  --���¹Һ����״̬
  If ���_In = 1 And Not n_��� Is Null Then
    If n_��ʱ�� = 1 Then
      d_���ʱ�� := ����ʱ��_In;
    Else
      d_���ʱ�� := Trunc(����ʱ��_In);
    End If;
    --������ŵĴ���
    Begin
      If n_ԤԼ˳��� Is Null Then
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_���;
      Else
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳���;
      End If;
      n_���� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_����       := 0;
    End;
    If n_���� = 0 Then
      If n_ԤԼ˳��� Is Null Then
        Update �ٴ�������ſ���
        Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
        Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
        Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
      End If;
      If Sql%RowCount = 0 Then
        Begin
          If Nvl(n_��ʱ��, 0) > 0 Then
            If Nvl(n_��ſ���, 0) = 1 Then
              --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
              Update �ٴ�������ſ���
              Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
              Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) In (0, 2);
              If Sql%NotFound Then
                Begin
                  Select �Һ�״̬ Into n_״̬ From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��� = n_���;
                Exception
                  When Others Then
                    n_״̬ := -1;
                End;
                If n_״̬ = -1 Then
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                    Select �����¼id_In, n_���, d_���ʱ��, d_���ʱ��, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1), Null,
                           Null, Null, ����Ա����_In, '׷�Ӻ�'
                    From Dual;
                Else
                  v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
                  Raise Err_Item;
                End If;
              End If;
            Else
              If Nvl(ԤԼ����_In, 0) = 1 Then
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע, ԤԼ˳���)
                  Select ��¼id, ���, ��ʼʱ��, ��ֹʱ��, 1, 1, Decode(ԤԼ�Һ�_In, 1, 2, 1), Null, Null, Null, ����Ա����_In, n_���, n_ԤԼ˳���
                  From �ٴ�������ſ���
                  Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Null;
              End If;
            End If;
          Else
            If Nvl(n_��ſ���, 0) = 1 Then
              Update �ٴ�������ſ���
              Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
              Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 0;
            
              If Sql%RowCount = 0 Then
                Begin
                  Select �Һ�״̬ Into n_״̬ From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��� = n_���;
                Exception
                  When Others Then
                    n_״̬ := -1;
                End;
                If n_״̬ = -1 Then
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                    Select �����¼id_In, n_���, ����ʱ��_In, ����ʱ��_In, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1),
                           Null, Null, Null, ����Ա����_In, '׷�Ӻ�'
                    From Dual;
                Else
                  v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
                  Raise Err_Item;
                End If;
              End If;
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        If n_ԤԼ˳��� Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And ����վ���� = v_������;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And
                ����վ���� = v_������;
        End If;
      End If;
    End If;
  End If;

  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, Decode(�շѵ�_In, Null, ժҪ_In, '����:' || �շѵ�_In), ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 And ���_In = 1 Then
      v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      v_���㷽ʽ��¼ := '';
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If Instr('|' || v_���㷽ʽ��¼ || '|', '|' || Nvl(v_���㷽ʽ, v_�ֽ�) || '|') <> 0 Then
          v_Err_Msg := 'ʹ�����ظ��Ľ��㷽ʽ,����!';
          Raise Err_Item;
        Else
          v_���㷽ʽ��¼ := v_���㷽ʽ��¼ || '|' || Nvl(v_���㷽ʽ, v_�ֽ�);
        End If;
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4,
             v_�������);
        
          If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
            Zl_���˿������¼_֧��(���㿨���_In, ����_In, 0, Nvl(n_������, 0), n_Ԥ��id, ����Ա���_In, ����Ա����_In, �Ǽ�ʱ��_In);
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
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
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
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2);
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
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 And Nvl(���½������_In, 1) = 1 Then
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �����¼id, �շѵ�)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �����¼id_In, �շѵ�_In);
  
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
  
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(ִ�в���id_In, 0)));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
        n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then
          n_��ʱ����ʾ := 1;
        Else
          n_��ʱ����ʾ := Null;
        End If;
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         n_��ʱ����ʾ, v_�Ŷ����);
      
        --�Һ������Ŷ�
        If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
          Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
        End If;
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
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
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Exists (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) > Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
    b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Insert;
/

--134969:���ϴ�,2019-01-14,Ԥ��֧�����
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
          Begin
            Select 1 Into n_���� From �Һ����״̬ Where ���� = v_�ű� And ���� = Trunc(Sysdate) And ��� = v_����;
          Exception
            When Others Then
              n_���� := 0;
          End;
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

--134441:���ϴ�,2019-01-16,�Һż����Ŀ�Ƿ�һ��
--134969:���ϴ�,2019-01-14,Ԥ��֧�����
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
        
          Begin
            Select 1
            Into n_����
            From �ٴ�������ſ���
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Exception
            When Others Then
              n_���� := 0;
          End;
        
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
          Where ��� = v_���� And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
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
        Where ��� = ����_In And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
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

--134441:���ϴ�,2019-01-15,�Һż����Ŀ�Ƿ�һ��
--134969:���ϴ�,2019-01-14,Ԥ��֧�����
Create Or Replace Procedure Zl_���������Һ�_Insert
(
  ������ʽ_In      Integer,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          �ҺŰ���.����%Type,
  ����_In          �Һ����״̬.���%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���㷽ʽ_In      Varchar2,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ������λ_In      �Һź�����λ.����%Type,
  �ҺŽ��ϼ�_In  ������ü�¼.ʵ�ս��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type,
  ����˵��_In      ����Ԥ����¼.����˵��%Type,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  �������״̬_In  Number := 0,
  �Ƿ������豸_In  Number := 0,
  ����id_In        ������ü�¼.����id%Type := Null,
  ��������_In      Number := 0,
  ���ս���_In      Varchar2 := Null,
  ��Ԥ��_In        Number := Null,
  ֧������_In      ����Ԥ����¼.����%Type := Null,
  �˺�����_In      Number := 1,
  �ѱ�_In          ������ü�¼.�ѱ�%Type := Null,
  ��Ԥ������ids_In Varchar2 := Null,
  ������_In        �Һ����״̬.������%Type := Null,
  ��������_In      Number := 0,
  ������_In      Number := 0,
  �����¼id_In    �ٴ������¼.Id%Type := Null,
  ���ʷ���_In      Number := 0,
  ���ʽ_In      ҽ�Ƹ��ʽ.����%Type := Null
) As
  --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�)
  --���:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
  --      ���㷽ʽ_IN:֧�ֶ��ֽ��㷽ʽ,���ֽ��㷽ʽʱ�������ʽ����:���㷽ʽ����1,���,�������,��������־|���㷽ʽ����2,���,�������,��������־|...
  --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
  --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
  --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
  --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  --      ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
  Err_Item Exception;
  Err_Special Exception;
  v_Err_Msg            Varchar2(255);
  n_��ӡid             Ʊ�ݴ�ӡ����.Id%Type;
  n_����ֵ             ����Ԥ����¼.���%Type;
  v_�ŶӺ���           Varchar2(20);
  v_��������           �ŶӽкŶ���.��������%Type;
  n_Ԥ��id             ����Ԥ����¼.Id%Type;
  n_�Һ�id             ���˹Һż�¼.Id%Type;
  v_��������           Varchar2(3000);
  v_��ǰ����           Varchar2(150);
  d_����ʱ��           Date;
  v_���㷽ʽ           ����Ԥ����¼.���㷽ʽ%Type;
  n_������           ����Ԥ����¼.��Ԥ��%Type;
  n_����ϼ�           Number(16, 5);
  n_Ԥ�����           ����Ԥ����¼.��Ԥ��%Type;
  n_�������           �������.Ԥ�����%Type;
  n_��id               ����ɿ����.Id%Type;
  d_�Ŷ�ʱ��           Date;
  n_����               Number;
  n_����ԤԼ������     Number(18);
  n_��Լ����           Number(18);
  n_������λ����       Number(18);
  n_�Ƿ񿪷�           Number(1);
  n_Count              Number(18);
  n_�к�               Number(18);
  n_���               ���˹Һż�¼.����%Type;
  n_����id             ������ü�¼.Id%Type;
  n_�۸񸸺�           Number(18);
  n_ԭ��Ŀid           �շ���ĿĿ¼.Id%Type;
  n_ԭ������Ŀid       �շ���ĿĿ¼.Id%Type;
  v_����               ���˹Һż�¼.����%Type;
  n_����id             �ҺŰ���.Id%Type;
  n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
  n_��������id         ������ü�¼.��������id%Type;
  n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_����id             ���˽��ʼ�¼.Id%Type;
  v_Temp               Varchar2(500);
  n_ԤԼʱ�����       Number;
  n_ԤԼ����           Number;
  n_Exists             Number;
  n_��ʱ����ʾ         Number;
  d_ʱ�ο�ʼʱ��       Date;
  v_��Ԥ������ids      Varchar2(4000);
  v_�շ���Ŀids        Varchar2(300);
  n_ԤԼ����           ������λ�ҺŻ���.��Լ��%Type;
  n_����               ���˹Һż�¼.����%Type;
  d_�Ǽ�ʱ��           Date;
  v_����Ա���         ��Ա��.���%Type;
  v_����Ա����         ��Ա��.����%Type;
  n_����               ���˹Һż�¼.����%Type;
  n_ԤԼ               Integer;
  v_����               �ҺŰ���ʱ��.����%Type;
  n_���÷�ʱ��         Integer;
  n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
  n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
  n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ���ɶ���       Number;
  d_Date               Date;
  n_�Һ����           Number;
  v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
  v_������             �Һ����״̬.������%Type;
  v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
  v_��Ż�����         �Һ����״̬.������%Type;
  n_�������           Number := 0;
  n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
  v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_�ѱ�               ������ü�¼.�ѱ�%Type;
  n_���ηѱ�           Number(3) := 0;
  n_Tmp����id          �ҺŰ���.Id%Type;
  n_�ƻ�id             �ҺŰ��żƻ�.Id%Type;
  v_����               ������Ϣ.����%Type;
  n_������λ������ģʽ Number;
  n_�����¼id         �ٴ������¼.Id%Type;
  n_�Һ�ģʽ           Number(3);
  n_ͬ���޺���         Number;
  n_ͬ����Լ��         Number;
  n_ͬԴ�޺���         Number;
  n_���˹Һſ�����     Number;
  d_����ʱ��           Date;
  v_Para               Varchar2(2000);
  n_ר�ҺŹҺ�����     Number;
  n_ר�Һ�ԤԼ����     Number;
  v_վ��               ���ű�.վ��%Type;
  v_��ͨ�ȼ�           Varchar2(100);
  v_Pricegrade         Varchar2(500);
  v_ʱ���             ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��       ʱ���.��ʼʱ��%Type;
  d_������ʱ��       ʱ���.��ֹʱ��%Type;
  v_����               Varchar2(100);
  n_������Ŀid         �ҺŰ���.��Ŀid%Type;
  n_��Ŀid             �ҺŰ���.��Ŀid%Type;

  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ, a.��������, a.���֤��
    From ������Ϣ A, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);

  r_Pati c_Pati%RowType;

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

  Cursor c_����
  (
    v_����        �ҺŰ���.����%Type,
    d_����ʱ��_In Date
  ) Is
    Select *
    From (With ����ʱ��� As (Select ʱ���
                         From (Select ʱ���,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��,
                                       To_Date('3000-01-10 ' || To_Char(d_����ʱ��_In, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ǰʱ��,
                                       To_Date('3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��1,
                                       To_Date('3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��1
                                From ʱ���)
                         Where ��ǰʱ�� Between ��ʼʱ�� And ��ֹʱ��1 Or ��ǰʱ�� Between ��ʼʱ��1 And ��ֹʱ��)
           Select Distinct p.Id, p.����, p.����, p.����id, b.���� As ���ұ���, b.���� As ��������, p.��Ŀid, c.���� As ��Ŀ����, c.���� As ��Ŀ����,
                           p.ҽ��id, d.��� As ҽ�����, p.ҽ������, p.�޺���, p.��Լ��, p.���� As ��, p.��һ As һ, p.�ܶ� As ��, p.���� As ��,
                           p.���� As ��, p.���� As ��, p.���� As ��, p.��ſ���, p.�ƻ�id
           From (Select p.Id, p.����, p.����, p.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(p.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�, Null As �ƻ�id
                  From �ҺŰ��� P, �ҺŰ������� B
                  Where p.ͣ������ Is Null And p.Id = b.����id(+) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And
                        d_����ʱ��_In Between Nvl(p.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                        Nvl(p.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                   (Select 1
                         From �ҺŰ��żƻ�
                         Where ����id = p.Id And (d_����ʱ��_In Between ��Чʱ�� + 0 And ʧЧʱ��) And ���ʱ�� Is Not Null) And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = p.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����
                  Union All
                  Select c.Id, c.����, c.����, c.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(c.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�, p.Id As �ƻ�id
                  From �ҺŰ��żƻ� P, �ҺŰ��� C, �Һżƻ����� B,
                       (Select Max(a.��Чʱ��) As ��Ч, ����id
                         From �ҺŰ��żƻ� A, �ҺŰ��� B
                         Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                               ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                               a.ʧЧʱ�� And b.���� = ����_In
                         Group By ����id) E
                  Where p.����id = c.Id And p.Id = b.�ƻ�id(+) And p.��Чʱ�� = e.��Ч And p.����id = e.����id And
                        Nvl(p.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And (d_����ʱ��_In Between p.��Чʱ�� + 0 And p.ʧЧʱ��) And
                        p.���ʱ�� Is Not Null And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = c.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����) P, ���ű� B, �շ���ĿĿ¼ C,
                ��Ա�� D
           Where p.����id = b.Id And p.ҽ��id = d.Id(+) And p.��Ŀid = c.Id And
                 (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                 (Nvl(p.ҽ��id, 0) = 0 Or Exists
                  (Select 1
                   From ��Ա�� Q
                   Where p.ҽ��id = q.Id And (q.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or q.����ʱ�� Is Null))) And Exists
            (Select 1 From ����ʱ��� Where ʱ��� = p.�Ű�))
           Order By ����;


  r_���� c_����%RowType;

  Function Zl_����(����_In �ҺŰ���.����%Type) Return Varchar2 As
    n_���﷽ʽ �ҺŰ���.���﷽ʽ%Type;
    n_����id   �ҺŰ���.Id%Type;
    v_����     ���˹Һż�¼.����%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    If ��������_In = 2 Then
      --�Ե��ݽ��н���,���ȼ���Ƿ��������
      Select Count(Rowid) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      If n_���� = 0 Then
        v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
        Raise Err_Item;
      End If;
      Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
    End If;
  
    Begin
      Select ID, Nvl(���﷽ʽ, 0) Into n_����id, n_���﷽ʽ From �ҺŰ��� Where ���� = ����_In;
    Exception
      When Others Then
        n_����id := -1;
    End;
  
    If n_����id = -1 Then
      v_Err_Msg := '����(' || ����_In || ')δ�ҵ�!';
      Raise Err_Item;
    End If;
    --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
    v_���� := Null;
    If n_���﷽ʽ = 1 Then
      --1-ָ������
      Begin
        Select �������� Into v_���� From �ҺŰ������� Where �ű�id = n_����id;
      Exception
        When Others Then
          v_���� := Null;
      End;
    End If;
    If n_���﷽ʽ = 2 Then
      --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
      For c_���� In (Select ��������, Sum(Num) As Num
                   From (Select ��������, 0 As Num
                          From �ҺŰ�������
                          Where �ű�id = n_����id
                          Union All
                          Select ����, Count(����) As Num
                          From ���˹Һż�¼
                          Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
                                ���� In (Select �������� From �ҺŰ������� Where �ű�id = n_����id)
                          Group By ����)
                   Group By ��������
                   Order By Num) Loop
        v_���� := c_����.��������;
        Exit;
      End Loop;
    End If;
    If n_���﷽ʽ = 3 Then
    
      --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
      n_Next  := 0;
      n_First := 1;
      For c_���� In (Select Rowid As Rid, �ű�id, ��������, ��ǰ���� From �ҺŰ������� Where �ű�id = n_����id) Loop
        If n_First = 1 Then
          v_Rowid := c_����.Rid;
        End If;
        If n_Next = 1 Then
          v_���� := c_����.��������;
          Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
          Exit;
        End If;
        If Nvl(c_����.��ǰ����, 0) = 1 Then
          Update �ҺŰ������� Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_���� Is Null Then
        Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning �������� Into v_����;
      End If;
    End If;
  
    Return v_����;
  End;

  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

  Procedure Zl_���������Һ�_����_Insert
  (
    ��¼id_In        �ٴ������¼.Id%Type,
    ������ʽ_In      Integer,
    ����id_In        ������ü�¼.����id%Type,
    ����_In          �ҺŰ���.����%Type,
    ����_In          �Һ����״̬.���%Type,
    ���ݺ�_In        ������ü�¼.No%Type,
    Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
    ���㷽ʽ_In      Varchar2,
    ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
    ����ʱ��_In      ������ü�¼.����ʱ��%Type,
    �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
    ������λ_In      �Һź�����λ.����%Type,
    �ҺŽ��ϼ�_In  ������ü�¼.ʵ�ս��%Type,
    ����id_In        Ʊ��ʹ����ϸ.����id%Type,
    �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
    ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In      ����Ԥ����¼.����˵��%Type,
    ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
    Ԥ��id_In        ����Ԥ����¼.Id%Type := Null,
    �����id_In      ����Ԥ����¼.�����id%Type := Null,
    �������״̬_In  Number := 0,
    �Ƿ������豸_In  Number := 0,
    ����id_In        ������ü�¼.����id%Type := Null,
    ��������_In      Number := 0,
    ���ս���_In      Varchar2 := Null,
    ��Ԥ��_In        Number := Null,
    ֧������_In      ����Ԥ����¼.����%Type := Null,
    �ѱ�_In          ������ü�¼.�ѱ�%Type := Null,
    ��Ԥ������ids_In Varchar2 := Null,
    ������_In        �Һ����״̬.������%Type := Null,
    ��������_In      Number := 0,
    ������_In      Number := 0,
    ���ʷ���_In      Number := 0,
    ���ʽ_In      ҽ�Ƹ��ʽ.����%Type := Null
  ) As
    --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�),������Ű�ģʽ��ʹ��
    --���: ������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
    --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
    --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
    --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
    --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
    --      ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
    Err_Item Exception;
    Err_Special Exception;
    v_Err_Msg  Varchar2(255);
    n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
    n_����ֵ   ����Ԥ����¼.���%Type;
    v_�ŶӺ��� Varchar2(20);
    v_�������� �ŶӽкŶ���.��������%Type;
    n_Ԥ��id   ����Ԥ����¼.Id%Type;
    n_�Һ�id   ���˹Һż�¼.Id%Type;
    v_�������� Varchar2(3000);
    v_��ǰ���� Varchar2(150);
  
    v_���㷽ʽ           ����Ԥ����¼.���㷽ʽ%Type;
    n_������           ����Ԥ����¼.��Ԥ��%Type;
    n_����ϼ�           Number(16, 5);
    n_Ԥ�����           ����Ԥ����¼.��Ԥ��%Type;
    n_�������           �������.Ԥ�����%Type;
    n_��id               ����ɿ����.Id%Type;
    d_�Ŷ�ʱ��           Date;
    n_����               Number;
    n_����ԤԼ������     Number(18);
    n_��Լ����           Number(18);
    d_����ʱ��           Date;
    n_������λ����       Number(18);
    n_�Ƿ񿪷�           Number(1);
    n_Count              Number(18);
    n_�к�               Number(18);
    n_����id             ������ü�¼.Id%Type;
    n_�۸񸸺�           Number(18);
    n_ԭ��Ŀid           �շ���ĿĿ¼.Id%Type;
    n_ԭ������Ŀid       �շ���ĿĿ¼.Id%Type;
    v_����               ���˹Һż�¼.����%Type;
    n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
    n_��������id         ������ü�¼.��������id%Type;
    n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_����               ���˹Һż�¼.����%Type;
    n_����id             ���˽��ʼ�¼.Id%Type;
    v_Temp               Varchar2(500);
    v_���㷽ʽ��¼       Varchar2(1000);
    n_ԤԼʱ�����       Number;
    n_��ſ���           �ٴ������¼.�Ƿ���ſ���%Type;
    n_��Լ��             �ٴ������¼.��Լ��%Type;
    n_��Ŀid             �ٴ������¼.��Ŀid%Type;
    n_����id             �ٴ������¼.����id%Type;
    d_��ֹʱ��           �ٴ������¼.��ֹʱ��%Type;
    v_ҽ������           �ٴ������¼.ҽ������%Type;
    n_ҽ��id             �ٴ������¼.ҽ��id%Type;
    n_ԤԼ˳���         �ٴ�������ſ���.ԤԼ˳���%Type;
    n_ԤԼ����           Number;
    d_ʱ�ο�ʼʱ��       Date;
    d_ʱ����ֹʱ��       Date;
    v_�շ���Ŀids        Varchar2(300);
    n_��������־         Number;
    v_�ű�               ���˹Һż�¼.�ű�%Type;
    n_����               ���˹Һż�¼.����%Type;
    d_�Ǽ�ʱ��           Date;
    n_���ʽ��           ����Ԥ����¼.��Ԥ��%Type;
    v_�������           ����Ԥ����¼.�������%Type;
    v_����Ա���         ��Ա��.���%Type;
    v_����Ա����         ��Ա��.����%Type;
    n_ԤԼ               Integer;
    n_��ʱ����ʾ         Number;
    v_�ֽ�               ����Ԥ����¼.���㷽ʽ%Type;
    n_���÷�ʱ��         Integer;
    n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
    n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
    n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
    n_ԤԼ���ɶ���       Number;
    n_�޺���             �ٴ������¼.�޺���%Type;
    d_Date               Date;
    n_�Һ����           Number;
    v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
    v_������             �Һ����״̬.������%Type;
    v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
    v_��Ż�����         �Һ����״̬.������%Type;
    n_�������           Number := 0;
    n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
    v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
    v_�ѱ�               ������ü�¼.�ѱ�%Type;
    n_���ηѱ�           Number(3) := 0;
    v_����               ������Ϣ.����%Type;
    n_������λ������ģʽ Number;
    n_ͬ���޺���         Number;
    n_ͬ����Լ��         Number;
    n_ͬԴ�޺���         Number;
    n_���˹Һſ�����     Number;
    n_Exists             Number(5);
    v_Exists             Varchar2(4000);
    v_��Ԥ������ids      Varchar2(4000);
    n_����ҽ��id         �ٴ������¼.����ҽ��id%Type;
    v_����ҽ������       �ٴ������¼.����ҽ������%Type;
    d_���￪ʼʱ��       �ٴ������¼.���￪ʼʱ��%Type;
    d_������ֹʱ��       �ٴ������¼.������ֹʱ��%Type;
    n_ר�ҺŹҺ�����     Number;
    n_ר�Һ�ԤԼ����     Number;
    v_վ��               ���ű�.վ��%Type;
    v_��ͨ�ȼ�           Varchar2(100);
    v_Pricegrade         Varchar2(500);
    v_����               Varchar2(100);
    n_������Ŀid         �ҺŰ���.��Ŀid%Type;
  
    Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
      Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ, a.��������, a.���֤��
      From ������Ϣ A, ҽ�Ƹ��ʽ C
      Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);
  
    r_Pati c_Pati%RowType;
  
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
  
    Function Zl_����(��¼id_In �ٴ������¼.Id%Type) Return Varchar2 As
      n_���﷽ʽ �ٴ������¼.���﷽ʽ%Type;
      v_����     ���˹Һż�¼.����%Type;
      v_Rowid    Varchar2(500);
      n_Next     Integer;
      n_First    Integer;
    Begin
    
      If ��������_In = 2 Then
        --�Ե��ݽ��н���,���ȼ���Ƿ��������
        Select Count(Rowid)
        Into n_����
        From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
        If n_���� = 0 Then
          v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
          Raise Err_Item;
        End If;
        Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      End If;
    
      Begin
        Select Nvl(���﷽ʽ, 0) Into n_���﷽ʽ From �ٴ������¼ Where ID = ��¼id_In;
      Exception
        When Others Then
          v_Err_Msg := '�����¼(' || ��¼id_In || ')δ�ҵ�!';
          Raise Err_Item;
      End;
    
      --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
      v_���� := Null;
      If n_���﷽ʽ = 1 Then
        --1-ָ������
        Begin
          Select b.���� Into v_���� From �ٴ��������Ҽ�¼ A, �������� B Where a.����id = b.Id And a.��¼id = ��¼id_In;
        Exception
          When Others Then
            v_���� := Null;
        End;
      End If;
      If n_���﷽ʽ = 2 Then
        --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
        For c_���� In (Select ��������, Sum(Num) As Num
                     From (Select b.���� As ��������, 0 As Num
                            From �ٴ��������Ҽ�¼ A, �������� B
                            Where a.����id = b.Id And a.��¼id = ��¼id_In
                            Union All
                            Select ����, Count(����) As Num
                            From ���˹Һż�¼
                            Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
                                  ���� In (Select d.����
                                         From �ٴ��������Ҽ�¼ C, �������� D
                                         Where c.����id = d.Id And c.��¼id = ��¼id_In)
                            Group By ����)
                     Group By ��������
                     Order By Num) Loop
          v_���� := c_����.��������;
          Exit;
        End Loop;
      End If;
      If n_���﷽ʽ = 3 Then
        --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
        n_Next  := 0;
        n_First := 1;
        For c_���� In (Select a.Rowid As Rid, b.���� As ��������, a.��ǰ����
                     From �ٴ��������Ҽ�¼ A, �������� B
                     Where a.����id = b.Id And a.��¼id = ��¼id_In) Loop
          If n_First = 1 Then
            v_Rowid := c_����.Rid;
          End If;
          If n_Next = 1 Then
            v_���� := c_����.��������;
            Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
            Exit;
          End If;
          If Nvl(c_����.��ǰ����, 0) = 1 Then
            Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
            n_Next := 1;
          End If;
        End Loop;
        If v_���� Is Null Then
          Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning ����id Into v_����;
          Select ���� Into v_���� From �������� Where ID = v_����;
        End If;
      End If;
      Return v_����;
    End;
  
    Function Zl_����Ա
    (
      Type_In     Integer,
      Splitstr_In Varchar2
    ) Return Varchar2 As
      n_Step Number(18);
      v_Sub  Varchar2(1000);
      --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
      -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
    Begin
      If Type_In = 0 Then
        --ȱʡ����
        n_Step := Instr(Splitstr_In, ',');
        v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 1 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, 1, n_Step - 1);
        Return v_Sub;
      End If;
      If Type_In = 2 Then
        --����Ա����
        n_Step := Instr(Splitstr_In, ';');
        v_Sub  := Substr(Splitstr_In, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        n_Step := Instr(v_Sub, ',');
        v_Sub  := Substr(v_Sub, n_Step + 1);
        Return v_Sub;
      End If;
    End;
  
  Begin
    d_����ʱ�� := ����ʱ��_In;
  
    If d_����ʱ�� Is Null Then
      d_����ʱ�� := Sysdate;
    End If;
  
    If ���ʽ_In Is Null Then
      Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
    Else
      Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
      If v_���ʽ Is Null Then
        v_���ʽ := ���ʽ_In;
      End If;
    End If;
    v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
  
    If �ѱ�_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, ����id_In) Into v_�ѱ� From Dual;
    Else
      v_�ѱ� := �ѱ�_In;
    End If;
    If v_�ѱ� Is Null Then
      n_���ηѱ� := 1;
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
    End If;
    Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
  
    If ��������_In = 1 Then
      Select Zl_Age_Calc(����id_In) Into v_���� From Dual;
      If v_���� Is Not Null Then
        Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;
      End If;
    End If;
    --��ȡ��ǰ��������
    If ������_In Is Not Null Then
      v_������ := ������_In;
    Else
      Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_ʵ�ս��ϼ� := 0;
  
    Select Count(*) + 1
    Into n_�Һ����
    From ���˹Һż�¼
    Where �����¼id = ��¼id_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
  
    If �Ǽ�ʱ��_In Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    Else
      d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
    End If;
    If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
      v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
      Raise Err_Item;
    End If;
  
    v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
    n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
    n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
    n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
    n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
    n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
    n_ͬԴ�޺���     := To_Number(Nvl(zl_GetSysParameter('����ͬһ��Դ�޹�N����', 1111), '0'));
  
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Item;
    End If;
    n_��������id := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
    n_��id       := Zl_Get��id(v_����Ա����);
  
    --֧���������ύ���
    Select Nvl(Max(1), 0)
    Into n_Exists
    From ���˹Һż�¼
    Where ����id = ����id_In And �ű� = ����_In And ���� = ����_In And ����Ա���� = v_����Ա���� And Nvl(��¼id_In, 0) = Nvl(�����¼id, 0) And
          �Ǽ�ʱ�� > Sysdate - 0.01 And ��¼״̬ = 1 And ����ʱ�� = ����ʱ��_In;
    If n_Exists = 1 Then
      v_Err_Msg := '�����Ѿ��Һ�,�����ظ�����ͬ�ĺţ�';
      Raise Err_Special;
    End If;
  
    If ������ʽ_In <> 1 Then
      --ԤԼ����Ƿ���Ӻ�����λ����
      --��������˺�����λ���� ��
      Begin
        Select 1
        Into n_������λ����
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = 1 And ���� = 1 And ���Ʒ�ʽ <> 4 And Rownum < 2;
      Exception
        When Others Then
          n_������λ���� := 0;
      End;
    End If;
  
    If ������ʽ_In <> 2 Then
      v_���� := Zl_����(��¼id_In);
    End If;
  
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*) Into n_Count From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ In (1, 3) And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    Open c_Pati(����id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '����δ�ҵ������ܼ�����';
      Raise Err_Item;
    End If;
  
    Begin
      Select Nvl(a.�Ƿ��ʱ��, 0), a.�޺���, a.�ѹ���, a.�����ѽ���, a.��Լ��, a.�Ƿ���ſ���, a.��Լ��, a.��Ŀid, a.����id, a.ҽ��id, a.ҽ������, a.����ҽ��id,
             a.����ҽ������, a.���￪ʼʱ��, a.������ֹʱ��, b.����
      Into n_���÷�ʱ��, n_�޺���, n_�ѹ���, n_�����ѽ���, n_��Լ��, n_��ſ���, n_��Լ��, n_��Ŀid, n_����id, n_ҽ��id, v_ҽ������, n_����ҽ��id, v_����ҽ������,
           d_���￪ʼʱ��, d_������ֹʱ��, v_�ű�
      From �ٴ������¼ a, �ٴ������Դ b
      Where a.ID = ��¼id_In and a.��Դid = b.id And Nvl(a.�Ƿ�����, 0) = 0;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
    
    IF v_�ű� <> ����_In Then
      v_Err_Msg := '��ǰ�ű�������¼�в�һ�£����ܼ�����';
      Raise Err_Item;
    End IF;
  
    Select Min(վ��) Into v_վ�� From ���ű� Where ID = n_����id;
    v_Pricegrade := Zl_Get_Pricegrade(v_վ��, ����id_In, Null, v_���ʽ);
    v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    If ����ʱ��_In Between Nvl(d_���￪ʼʱ��, Sysdate) And Nvl(d_������ֹʱ��, Sysdate - 1) And v_����ҽ������ Is Not Null Then
      n_ҽ��id   := n_����ҽ��id;
      v_ҽ������ := v_����ҽ������;
    End If;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And �����¼id = ��¼id_In;
        If n_Count >= Nvl(n_ר�Һ�ԤԼ����, 0) And Nvl(n_ר�Һ�ԤԼ����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ���������ԤԼ����,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And �����¼id = ��¼id_In;
        If n_Count >= Nvl(n_ר�ҺŹҺ�����, 0) And Nvl(n_ר�ҺŹҺ�����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��������ŹҺ�����,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_ͬԴ�޺���, 0) <> 0 Then
      If �����¼id_In Is Null Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = ����_In;
      Else
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And �����¼id = �����¼id_In;
      End If;
      If n_Count >= Nvl(n_ͬԴ�޺���, 0) And Nvl(n_ͬԴ�޺���, 0) > 0 Then
        v_Err_Msg := 'ͬһ���������ͬʱ��(ԤԼ)[' || Nvl(n_ͬԴ�޺���, 0) || ']����ͬ�ű�ĺ�,�����ٹҺ�(ԤԼ)��';
        Raise Err_Item;
      End If;
    End If;
  
    d_Date         := Null;
    d_ʱ�ο�ʼʱ�� := Null;
  
    If Nvl(n_�޺���, 0) >= 0 Or n_�޺��� Is Null Then
      If n_���÷�ʱ�� = 1 Then
        If Nvl(n_��ſ���, 0) = 1 Then
          If Nvl(�Ƿ������豸_In, 0) = 0 Then
            Select Count(*), Max(��ʼʱ��)
            Into n_Count, d_ʱ�ο�ʼʱ��
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0);
          
            v_Temp := '�Һ�';
            If ������ʽ_In > 1 Then
              v_Temp := 'ԤԼ�Һ�';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ����в��������Ϊ' || Nvl(����_In, 0) || '�İ���,������' || v_Temp || '��';
              Raise Err_Item;
            End If;
          End If;
        
          --�����,����ѡ��Һ�
          If Trunc(Sysdate) = Trunc(����ʱ��_In) Then
            --�ҵ���ĺ�
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                         To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��, ����, �Ƿ�ԤԼ
                         From �ٴ�������ſ���
                         Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0)) Loop
              If Sysdate > v_ʱ��.��ֹʱ�� Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                Raise Err_Item;
              End If;
            End Loop;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          For v_ʱ�� In (Select ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ
                       From �ٴ�������ſ���
                       Where ��¼id = ��¼id_In And
                             (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                             '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                             ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-11 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
            n_ԤԼʱ����� := v_ʱ��.���;
            d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            d_ʱ����ֹʱ�� := v_ʱ��.��ֹʱ��;
          
            Select Count(*), Max(���), Max(ԤԼ˳���) + 1
            Into n_Count, n_ԤԼ����, n_ԤԼ˳���
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) Not In (0, 4, 5);
          
            If Nvl(n_Count, 0) > Nvl(v_ʱ��.����, 0) And ��������_In <> 2 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                           To_Char(v_ʱ��.��ֹʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.����, 0) || '��,�����ٽ���ԤԼ�Һţ�';
              Raise Err_Item;
            End If;
            n_Count := 1;
          End Loop;
        
          If n_Count = 0 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�����û����صİ��żƻ�(' || To_Char(����ʱ��_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),���ܽ���ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If ������ʽ_In = 1 And ��������_In <> 2 Then
      --�ҺŹ���:
      --  �ѹ������ܴ����޺���
      If n_�ѹ��� >= Nvl(n_�޺���, 0) And n_�޺��� Is Not Null Then
        v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(n_�޺���, 0) || '�����ٹҺţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If ������ʽ_In > 1 Then
      --ԤԼ����ؼ��
      --����:
      --   1.����Լ���ܳ�����Լ��
      --   2.����Ƿ�����ʱ�ε�
      If n_��Լ�� >= Nvl(n_��Լ��, 0) And Nvl(n_��Լ��, 0) <> 0 And n_��Լ�� Is Not Null And ��������_In <> 2 Then
        v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(n_��Լ��, 0) || '������ԤԼ�Һţ�';
        Raise Err_Item;
      End If;
      If ԤԼ��ʽ_In Is Not Null Then
        Select Zl_Fun_Get�ٴ�����ԤԼ״̬(��¼id_In, ����ʱ��_In, ����_In, ԤԼ��ʽ_In, Null, 0, v_����Ա����, v_������)
        Into v_Exists
        From Dual;
        If To_Number(Substr(v_Exists, 1, 1)) <> 0 Then
          v_Err_Msg := '�����ԤԼ��ʽ' || ԤԼ��ʽ_In || '������,ԭ��:' || Substr(v_Exists, 3);
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
      If Nvl(n_��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
        v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
        Raise Err_Item;
      End If; --Nvl(r_����.��ſ���, 0) =0
    
      --������λ����ģʽ
      Begin
        Select Nvl(���Ʒ�ʽ, 0)
        Into n_������λ������ģʽ
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And Rownum < 2;
      Exception
        When Others Then
          n_������λ������ģʽ := 4;
      End;
    
      If n_������λ������ģʽ = 0 Then
        v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || 'δ����' || ������λ_In || '��ԤԼ,���ܼ�����';
        Raise Err_Item;
      End If;
      If n_������λ������ģʽ = 1 Or n_������λ������ģʽ = 2 Then
        Select ����
        Into n_Count
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1;
        If n_������λ������ģʽ = 1 Then
          n_Count := Round(Nvl(n_��Լ��, n_�޺���) * n_Count / 100);
        End If;
        Select Count(1)
        Into n_Exists
        From ���˹Һż�¼
        Where ��¼״̬ = 1 And �����¼id = ��¼id_In And ������λ = ������λ_In;
        If n_Exists >= n_Count Then
          v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
      --������ż��
      If n_������λ������ģʽ = 3 Then
        For c_������λ In (Select ���, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And ��� = ����_In) Loop
          If n_��ſ��� = 1 Then
            Begin
              Select 1
              Into n_Count
              From �ٴ�������ſ���
              Where ��¼id = ��¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 0;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 1 Then
              n_�Ƿ񿪷� := 1;
            Else
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            End If;
          Else
            Select Count(1)
            Into n_Count
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = ����_In And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
            If n_Count >= c_������λ.���� Then
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            Else
              n_�Ƿ񿪷� := 1;
            End If;
          End If;
        End Loop;
      
        If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
          v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    --����޺�������Լ��
    n_�к�         := 1;
    n_ԭ��Ŀid     := 0;
    n_ԭ������Ŀid := 0;
    n_ʵ�ս��ϼ� := 0;
    If ��������_In <> 1 Then
      If ������ʽ_In <> 2 Then
        If Nvl(����id_In, 0) = 0 Then
          --����Ӧ�ó�����
          Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
        Else
          n_����id := ����id_In;
        End If;
      Else
        n_����id := Null;
      End If;
    End If;
  
    If Nvl(��¼id_In, 0) <> 0 Then
      v_���� := '2|' || ��¼id_In;
    End If;
    If v_���� Is Null Then
      v_���� := '3|' || ����_In;
    End If;
  
    n_������Ŀid := Zl_Custom_Getregeventitem(r_Pati.����id, r_Pati.����, r_Pati.���֤��, r_Pati.��������, r_Pati.�Ա�, r_Pati.����, v_����);
    If Nvl(n_������Ŀid, 0) <> 0 Then
      n_��Ŀid := n_������Ŀid;
    End If;
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := n_��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := n_��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������,
                          Nvl(a.��Ŀ����, 0) As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_��Ŀid And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Order By ����, ��Ŀ����, �������) Loop
      If c_Item.���� = 1 Then
        n_���� := Nvl(c_Item.����, 0);
      End If;
      n_�۸񸸺� := Null;
      If n_ԭ��Ŀid = c_Item.��Ŀid Then
        If n_ԭ������Ŀid <> c_Item.������Ŀid Then
          n_�۸񸸺� := n_�к�;
        End If;
        n_ԭ������Ŀid := c_Item.������Ŀid;
      End If;
      n_ԭ��Ŀid := c_Item.��Ŀid;
      n_Ӧ�ս�� := Round(c_Item.���� * c_Item.����, 5);
      n_ʵ�ս�� := n_Ӧ�ս��;
      If Nvl(c_Item.���ηѱ�, 0) <> 1 And n_���ηѱ� = 0 Then
        --����:
        v_Temp     := Zl_Actualmoney(r_Pati.�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_ʵ�ս�� := Zl_To_Number(v_Temp);
      End If;
      n_ʵ�ս��ϼ� := Nvl(n_ʵ�ս��ϼ�, 0) + n_ʵ�ս��;
    
      --�������ݲ���������
      If ��������_In <> 1 Then
        --�������˹Һŷ���(���ܵ����ǻ������������)
        Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
        --:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
        Values
          (n_����id, 4, Decode(������ʽ_In, 2, 0, 1), n_�к�, n_�۸񸸺�, c_Item.��������, ���ݺ�_In, Ʊ�ݺ�_In, 1, n_����, Null,
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, n_����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, n_ʵ�ս��)),
           Decode(Nvl(���ʷ���_In, 0), 1, Null, n_����id), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), n_����id, v_ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null, Null,
           ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(����ʱ��_In);
    End If;
  
    --���¹Һ����״̬
    If ��������_In <> 2 Then
      n_���� := ����_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And ��� = n_���� And Nvl(�Һ�״̬, 0) Not In (0, 5);
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 Then
        n_���� := Null;
      End If;
      If n_���÷�ʱ�� = 1 And Nvl(n_��ſ���, 0) = 1 Then
        v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
      Select Nvl(Min(���), 0)
      Into n_����
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
      If n_���� = 0 Then
        Select Nvl(Min(���), 0) Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 0;
        If n_���� = 0 Then
          Select Nvl(Max(���), 0) + 1 Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
        End If;
      End If;
    End If;
  
    If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
      If ������ʽ_In > 1 And Nvl(n_��ſ���, 0) = 0 Then
        --����:ԤԼʱ�����||ԤԼ��
        If Nvl(n_ԤԼ����, 0) = 0 Then
          v_Temp := Nvl(n_��Լ��, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_ԤԼʱ����� || v_Temp;
          n_���� := To_Number(v_Temp);
        Else
          n_���� := n_ԤԼ���� + 1;
        End If;
      End If;
    End If;
  
    If Nvl(n_��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
      --������ŵĴ���
      Begin
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where �Һ�״̬ = 5 And ��¼id = ��¼id_In And ��� = n_����;
        n_������� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_�������   := 0;
      End;
      If n_������� = 0 Then
        If n_���÷�ʱ�� = 1 And n_��ſ��� = 0 Then
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����, ��ע)
            Select ��¼id_In, n_ԤԼʱ�����, n_ԤԼ˳���, d_ʱ�ο�ʼʱ��, d_ʱ����ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1),
                   1, ������λ_In, v_����Ա����, n_����
            From Dual;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����Ա���� = v_����Ա����
          Where ��¼id = ��¼id_In And ��� = n_����;
        End If;
        If Sql%RowCount = 0 Then
          Begin
            If n_���÷�ʱ�� = 1 Then
              --��ʱ��
              If n_��ſ��� = 1 Then
                --��ſ���
                Select Max(��ֹʱ��) Into d_��ֹʱ�� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
                If Sysdate > d_��ֹʱ�� Then
                  d_��ֹʱ�� := Sysdate;
                End If;
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                  Select ��¼id_In, n_����, d_��ֹʱ��, d_��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1,
                         ������λ_In, v_����Ա����
                  From Dual;
              Else
                --��ʱ��,����ſ���
                Null;
              End If;
            Else
              --����ʱ��
              Insert Into �ٴ�������ſ���
                (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                Select ��¼id_In, n_����, ��ʼʱ��, ��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1, ������λ_In,
                       v_����Ա����
                From �ٴ�������ſ���
                Where ��¼id = ��¼id_In And ��� = 1;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_���� || '�ѱ�����' || v_������ || '����,������ѡ��һ�����.';
          Raise Err_Item;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����ʱ�� = Null
          Where ��¼id = ��¼id_In And ��� = n_���� And �Һ�״̬ = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
        End If;
      End If;
    End If;
  
    --�������ݲ������κ� ����
    If ������ʽ_In <> 2 And ��������_In <> 1 And Nvl(���ʷ���_In, 0) = 0 Then
      --�Һ�,ԤԼ�Һ��Ѿ��ۿ��
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      n_����ϼ� := 0;
      If ���ս���_In Is Not Null Then
        --�������ս���
        v_�������� := ���ս���_In || '||';
        n_����ϼ� := 0;
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
          n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
          n_����ϼ� := Nvl(n_����ϼ�, 0) + Nvl(n_������, 0);
          v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(��Ԥ��_In, 0) <> 0 Then
        Select Nvl(Sum(Nvl(Ԥ�����, 0) - Nvl(�������, 0)), 0)
        Into n_�������
        From �������
        Where ����id In (Select /*+cardinality(d,10)*/
                        d.Column_Value
                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(����, 0) = 1 And Nvl(����, 0) = 1;
        If n_������� < ��Ԥ��_In Then
          v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || Ltrim(To_Char(n_�������, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                       Ltrim(To_Char(��Ԥ��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
          Raise Err_Item;
        End If;
        
        --������Ԥ��
        n_����ϼ� := n_����ϼ� + Nvl(��Ԥ��_In, 0);
        n_Ԥ����� := ��Ԥ��_In;
        For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
          n_������ := Case
                      When r_Deposit.��� - n_Ԥ����� < 0 Then
                       r_Deposit.���
                      Else
                       n_Ԥ�����
                    End;
          If r_Deposit.����id = 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (r_Deposit.����id, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          --����Ƿ��Ѿ�������
          If r_Deposit.��� <= n_������ Then
            n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
          Else
            n_Ԥ����� := 0;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
        If n_Ԥ����� > 0 Then
          v_Err_Msg := 'Ԥ������֧������֧�����,���ܼ���������';
          Raise Err_Item;
        End If;
      End If;
      --ʣ�����,��ָ�����㷽֧��
      n_������ := Nvl(n_ʵ�ս��ϼ�, 0) - Nvl(n_����ϼ�, 0);
      If Nvl(n_������, 0) < 0 Then
        v_Err_Msg := '�Һŵ���ؽ�������˵�ǰʵ����,���ܼ���������';
        Raise Err_Item;
      End If;
    
      If Nvl(n_������, 0) <> 0 Or (Nvl(n_������, 0) = 0 And Nvl(��Ԥ��_In, 0) = 0) Then
        If ���㷽ʽ_In Is Null Then
          v_Err_Msg := 'δ����ָ���Ľ��㷽ʽ,���ܼ���������';
          Raise Err_Item;
        End If;
      
        If Nvl(Ԥ��id_In, 0) <> 0 Then
          --�����Ԥ��ID_In��Ҫ��Ϊ�˽����������,���ҽ������վ���˸�ID,��Ҫ���µ�ID���и���,����������ת���ID
          Update ����Ԥ����¼ Set ID = n_Ԥ��id Where ID = Nvl(Ԥ��id_In, 0);
          n_Ԥ��id := Nvl(Ԥ��id_In, 0);
        End If;
        If Instr(���㷽ʽ_In, ',') = 0 Then
          --ֻ����һ�ֽ��㷽ʽ��
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
        Else
          v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
          n_Exists       := 0;
          v_���㷽ʽ��¼ := '';
          While v_�������� Is Not Null Loop
            v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
            v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            n_���ʽ�� := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            n_��������־ := To_Number(v_��ǰ����);
          
            If Instr('|' || v_���㷽ʽ��¼ || '|', '|' || Nvl(v_���㷽ʽ, v_�ֽ�) || '|') <> 0 Then
              v_Err_Msg := 'ʹ�����ظ��Ľ��㷽ʽ,����!';
              Raise Err_Item;
            Else
              v_���㷽ʽ��¼ := v_���㷽ʽ��¼ || '|' || Nvl(v_���㷽ʽ, v_�ֽ�);
            End If;
          
            If n_��������־ = 0 Then
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
            Else
              If n_Exists = 1 Then
                v_Err_Msg := 'Ŀǰ�ҺŽ�֧��һ���������㷽ʽ,���ܼ���������';
                Raise Err_Item;
              End If;
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
              n_Exists := 1;
            End If;
          
            v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
          End Loop;
        End If;
      End If;
    
      --������Ա�ɿ�����
    
      For v_�ɿ� In (Select ���㷽ʽ, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ A
                   Where a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And Nvl(����id, 0) = Nvl(����id_In, 0)
                   Group By ���㷽ʽ) Loop
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(v_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_�ɿ�.���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, v_�ɿ�.���㷽ʽ, 1, Nvl(v_�ɿ�.��Ԥ��, 0));
          n_����ֵ := Nvl(v_�ɿ�.��Ԥ��, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = v_�ɿ�.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      End Loop;
    End If;
  
    --����Һż�¼
    If ��������_In = 2 Then
      Begin
        Select ID Into n_�Һ�id From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    End If;
  
    Update ���˹Һż�¼
    Set ��¼���� = Decode(������ʽ_In, 2, 2, 1), ��¼״̬ = Decode(��������_In, 1, 0, 1), ����� = r_Pati.�����, ����Ա���� = v_����Ա����,
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
        ����ʱ�� = Case ��������_In
                  When 1 Then
                   Null
                  Else
                   Case ������ʽ_In
                     When 2 Then
                      Null
                     Else
                      d_�Ǽ�ʱ��
                   End
                End, ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), ������λ = Nvl(������λ_In, ������λ),
        ԤԼ����Ա = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա, Null), Nvl(ԤԼ����Ա, v_����Ա����)),
        ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���)), �����¼id = ��¼id_In
    Where ID = n_�Һ�id;
    If Sql%NotFound Then
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = r_Pati.���ʽ And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
         ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���, �����¼id)
      Values
        (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
         r_Pati.�Ա�, r_Pati.����, ����_In, n_����, v_����, Null, n_����id, v_ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
         Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���), ��¼id_In);
    End If;
    --�������ݲ��ܲ�������
    If ��������_In <> 1 Then
      n_ԤԼ���ɶ��� := 0;
      If ������ʽ_In > 1 Then
        n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
      End If;
      --�Һź��շѵ�ԤԼ��ֱ�ӽ������(�շ�ԤԼȱ�ٽ��չ���,����ֱ�Ӻ͹Һ�һ��ֱ�ӽ������)
      If ������ʽ_In <> 2 Or n_ԤԼ���ɶ��� = 1 Then
        If Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113)) <> 0 Then
          --�Ŷӽк�ģʽ:-0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
          If Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(n_����id, 0))) = 0 Or n_ԤԼ���ɶ��� = 1 Then
            n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(������ʽ_In, 0) > 1 And n_��ʱ����ʾ = 1 And n_���÷�ʱ�� = 1 Then
              n_��ʱ����ʾ := 1;
            Else
              n_��ʱ����ʾ := Null;
            End If;
            --��������
            --.����ִ�в��š� �ķ�ʽ���ɶ���
            v_�������� := n_����id;
            v_�ŶӺ��� := Zlgetnextqueue(n_����id, n_�Һ�id, ����_In || '|' || ����_In);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
            --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
            Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, n_����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, v_ҽ������, d_�Ŷ�ʱ��,
                             ԤԼ��ʽ_In, n_��ʱ����ʾ, v_�Ŷ����);
          End If;
        End If;
      End If;
    
      If Nvl(������ʽ_In, 0) = 1 And Nvl(���ʷ���_In, 0) = 0 Then
        --����Ʊ��ʹ�����
        If Ʊ�ݺ�_In Is Not Null Then
          Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
          --����Ʊ��
          Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Values
            (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�Ǽ�ʱ��, v_����Ա����, �ҺŽ��ϼ�_In);
          --״̬�Ķ�
          Update Ʊ�����ü�¼
          Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
          Where ID = Nvl(����id_In, 0);
        End If;
        --���˱��ξ���(�Է���ʱ��Ϊ׼)
        If Nvl(r_Pati.����id, 0) <> 0 Then
          Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = v_���� Where ����id = r_Pati.����id;
        End If;
      End If;
    End If;
    --���˹ҺŻ���
    --��������ʱ�����ٶԻ��ܵ��ݽ���ͳ���� ����������ʱ�Ѿ������˻���
    If ��������_In <> 2 Then
      --������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
      --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
      --����zl_third_lockno�������ţ�������ʹ�ñ���������
      n_ԤԼ := Case
                When Nvl(������ʽ_In, 0) = 1 Then
                 0
                When Nvl(������ʽ_In, 0) = 2 Then
                 1
                When Nvl(������ʽ_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_��Ŀid, n_����id, ����ʱ��_In, n_ԤԼ, ����_In, 0, ��¼id_In);
    End If;
  
    If ��������_In <> 1 Then
      --��Ϣ����,����ʱ��������Ϣ
      Begin
        Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
          Using 1, n_�Һ�id;
      Exception
        When Others Then
          Null;
      End;
      b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Err_Special Then
      Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_�����¼id := �����¼id_In;
  v_Para       := zl_GetSysParameter(256);
  n_�Һ�ģʽ   := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  d_����ʱ�� := ����ʱ��_In;
  If d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;

  If ���ʽ_In Is Null Then
    Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
  Else
    Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
    If v_���ʽ Is Null Then
      v_���ʽ := ���ʽ_In;
    End If;
  End If;

  If Sysdate - 10 > Nvl(d_����ʱ��, Sysdate - 30) Then
    If n_�Һ�ģʽ = 1 And Nvl(����ʱ��_In, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
      v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
      Raise Err_Item;
    End If;
  Else
    If n_�Һ�ģʽ = 1 And Nvl(����ʱ��_In, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = ����_In And Nvl(����ʱ��_In, Sysdate) Between a.��ʼʱ�� And a.��ֹʱ��;
      Exception
        When Others Then
          v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If n_�����¼id Is Not Null Then
    --������Ű�ģʽ
    Zl_���������Һ�_����_Insert(n_�����¼id, ������ʽ_In, ����id_In, ����_In, ����_In, ���ݺ�_In, Ʊ�ݺ�_In, ���㷽ʽ_In, ժҪ_In, ����ʱ��_In, �Ǽ�ʱ��_In,
                        ������λ_In, �ҺŽ��ϼ�_In, ����id_In, �շ�Ʊ��_In, ������ˮ��_In, ����˵��_In, ԤԼ��ʽ_In, Ԥ��id_In, �����id_In, �������״̬_In,
                        �Ƿ������豸_In, ����id_In, ��������_In, ���ս���_In, ��Ԥ��_In, ֧������_In, �ѱ�_In, ��Ԥ������ids_In, ������_In, ��������_In,
                        ������_In, ���ʷ���_In, ���ʽ_In);
  Else
    v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
    v_Temp          := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          Null;
      End;
      If ����ʱ��_In > d_����ʱ�� Then
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
        Raise Err_Item;
      End If;
    End If;
    If �ѱ�_In Is Null Then
      Select Zl_Custom_Getpatifeetype(1, ����id_In) Into v_�ѱ� From Dual;
    Else
      v_�ѱ� := �ѱ�_In;
    End If;
    If v_�ѱ� Is Null Then
      n_���ηѱ� := 1;
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
    End If;
    Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
  
    If ��������_In = 1 Then
      Select Zl_Age_Calc(����id_In) Into v_���� From Dual;
      If v_���� Is Not Null Then
        Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;
      End If;
    End If;
    --��ȡ��ǰ��������
    If ������_In Is Not Null Then
      v_������ := ������_In;
    Else
      Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
    End If;
    n_ʵ�ս��ϼ� := 0;
    Select Count(*) + 1
    Into n_�Һ����
    From ���˹Һż�¼
    Where �ű� = ����_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
    --Begin
    v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
    n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
    n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
    n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
    n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
    n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
    n_ͬԴ�޺���     := To_Number(Nvl(zl_GetSysParameter('����ͬһ��Դ�޹�N����', 1111), '0'));
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If �Ǽ�ʱ��_In Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    Else
      d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
    End If;
    If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
      v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
      Raise Err_Item;
    End If;
    n_��������id := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
    n_��id       := Zl_Get��id(v_����Ա����);
  
    --֧���������ύ���
    Select Nvl(Max(1), 0)
    Into n_Exists
    From ���˹Һż�¼
    Where ����id = ����id_In And �ű� = ����_In And ���� = ����_In And ����Ա���� = v_����Ա���� And Nvl(n_�����¼id, 0) = Nvl(�����¼id, 0) And
          �Ǽ�ʱ�� > Sysdate - 0.01 And ��¼״̬ = 1 And ����ʱ�� = ����ʱ��_In;
    If n_Exists = 1 Then
      v_Err_Msg := '�����Ѿ��Һ�,�����ظ�����ͬ�ĺţ�';
      Raise Err_Special;
    End If;
  
    If ������ʽ_In <> 1 Then
      --ԤԼ����Ƿ���Ӻ�����λ����
      --��������˺�����λ���� ��
      Begin
        Select ID
        Into n_�ƻ�id
        From �ҺŰ��żƻ�
        Where ���� = ����_In And ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
              ʧЧʱ�� And Rownum < 2
        Order By ��Чʱ�� Desc;
      Exception
        When Others Then
          Select ID Into n_Tmp����id From �ҺŰ��� Where ���� = ����_In;
      End;
      If Nvl(n_�ƻ�id, 0) <> 0 Then
        Select Count(0)
        Into n_������λ����
        From ������λ�ƻ�����
        Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And Rownum < 2;
      Else
        Select Count(0)
        Into n_������λ����
        From ������λ���ſ���
        Where ������λ = ������λ_In And ����id = n_Tmp����id And Rownum < 2;
      End If;
    End If;
  
    If ������ʽ_In <> 2 Then
      v_���� := Zl_����(����_In);
    End If;
    If ������ʽ_In <> 2 And ���㷽ʽ_In Is Not Null Then
      --�����㷽ʽ�Ƿ��걸
      Select Count(*) Into n_Count From ���㷽ʽ Where ���� = Nvl(���㷽ʽ_In, 'Lxh') And ���� In (2, 7, 8);
      If Nvl(�����id_In, 0) <> 0 And n_Count = 0 Then
        Select Count(1)
        Into n_Count
        From ҽ�ƿ����
        Where ID = Nvl(�����id_In, 0) And ���㷽ʽ = Nvl(���㷽ʽ_In, 'lxh');
      End If;
      If n_Count = 0 Then
        v_Err_Msg := '���㷽ʽ(' || ���㷽ʽ_In || ')δ����,���ڽ��㷽ʽ���������á�';
        Raise Err_Item;
      End If;
    End If;
  
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*) Into n_Count From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ In (1, 3) And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    Open c_Pati(����id_In);
    n_Count := 0;
    Begin
      Fetch c_Pati
        Into r_Pati;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '����δ�ҵ������ܼ�����';
      Raise Err_Item;
    End If;
  
    Open c_����(����_In, ����ʱ��_In);
    Begin
      Fetch c_����
        Into r_����;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
  
    Select Min(վ��) Into v_վ�� From ���ű� Where ID = r_����.����id;
    v_Pricegrade := Zl_Get_Pricegrade(v_վ��, ����id_In, Null, v_���ʽ);
    v_��ͨ�ȼ�   := Substr(v_Pricegrade, 1, Instr(v_Pricegrade, '|') - 1);
  
    Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                   '����')
    Into v_����
    From Dual;
    Begin
      If r_����.�ƻ�id Is Null Then
        Select Max(1) Into n_���÷�ʱ�� From �ҺŰ���ʱ�� Where ����id = r_����.Id And ���� = v_���� And Rownum < 2;
        Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
        Into v_ʱ���
        From �ҺŰ���
        Where ID = r_����.Id;
      Else
        Select Max(1)
        Into n_���÷�ʱ��
        From �Һżƻ�ʱ��
        Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And Rownum < 2;
        Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
        Into v_ʱ���
        From �ҺŰ��żƻ�
        Where ID = r_����.�ƻ�id;
      End If;
    Exception
      When Others Then
        n_���÷�ʱ�� := 0;
    End;
  
    If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
      --����Ƿ��ģʽ�ҺŰ���
      Select To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_��鿪ʼʱ��, d_������ʱ��
      From ʱ���
      Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
      If d_��鿪ʼʱ�� > d_������ʱ�� Then
        d_������ʱ�� := d_������ʱ�� + 1;
      End If;
      If d_������ʱ�� > d_����ʱ�� Then
        --��ȡ�����¼id
        Begin
          Select a.Id
          Into n_�����¼id
          From �ٴ������¼ A, �ٴ������Դ B
          Where a.��Դid = b.Id And b.���� = ����_In And �ϰ�ʱ�� = v_ʱ��� And ����ʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
        Exception
          When Others Then
            n_�����¼id := Null;
        End;
      End If;
    End If;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> r_����.����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = r_����.����id;
        If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = r_����.����;
        If n_Count >= Nvl(n_ר�Һ�ԤԼ����, 0) And Nvl(n_ר�Һ�ԤԼ����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ���������ԤԼ����,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> r_����.����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = r_����.����id;
        If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = r_����.����;
        If n_Count >= Nvl(n_ר�ҺŹҺ�����, 0) And Nvl(n_ר�ҺŹҺ�����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��������ŹҺ�����,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_ͬԴ�޺���, 0) <> 0 Then
      If �����¼id_In Is Null Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = ����_In;
      Else
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� In (1, 2) And �����¼id = �����¼id_In;
      End If;
      If n_Count >= Nvl(n_ͬԴ�޺���, 0) And Nvl(n_ͬԴ�޺���, 0) > 0 Then
        v_Err_Msg := 'ͬһ���������ͬʱ��(ԤԼ)[' || Nvl(n_ͬԴ�޺���, 0) || ']����ͬ�ű�ĺ�,�����ٹҺ�(ԤԼ)��';
        Raise Err_Item;
      End If;
    End If;
  
    d_Date         := Null;
    d_ʱ�ο�ʼʱ�� := Null;
  
    If Nvl(r_����.�޺���, 0) >= 0 Or r_����.�޺��� Is Null Then
    
      Select Nvl(Sum(Nvl(b.�ѹ���, 0)), 0), Nvl(Sum(Nvl(b.�����ѽ���, 0)), 0), Nvl(Sum(Nvl(b.��Լ��, 0)), 0)
      Into n_�ѹ���, n_�����ѽ���, n_��Լ��
      From �ҺŰ��� A, ���˹ҺŻ��� B
      Where a.����id = b.����id And a.��Ŀid = b.��Ŀid And a.���� = ����_In And b.���� Between Trunc(����ʱ��_In) And
            Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And (a.���� = b.���� Or b.���� Is Null) And Nvl(a.ҽ��id, 0) = Nvl(b.ҽ��id, 0) And
            Nvl(a.ҽ������, 'ҽ��') = Nvl(b.ҽ������, 'ҽ��');
    
      If n_���÷�ʱ�� = 1 Then
        If Nvl(r_����.��ſ���, 0) = 1 Then
          If Nvl(�Ƿ������豸_In, 0) = 0 Then
            If r_����.�ƻ�id Is Null Then
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �ҺŰ���ʱ��
              Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0);
            Else
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �Һżƻ�ʱ��
              Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0);
            End If;
            v_Temp := '�Һ�';
            If ������ʽ_In > 1 Then
              v_Temp := 'ԤԼ�Һ�';
            End If;
          
            If n_Count = 0 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ����в��������Ϊ' || Nvl(����_In, 0) || '�İ���,������' || v_Temp || '��';
              Raise Err_Item;
            End If;
          End If;
          --�����,����ѡ��Һ�
          If Trunc(Sysdate) = Trunc(����ʱ��_In) Then
            --�ҵ���ĺ�
            v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
            If r_����.�ƻ�id Is Null Then
              For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                           From �ҺŰ���ʱ��
                           Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
                If Sysdate > v_ʱ��.����ʱ�� Then
                  v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                  Raise Err_Item;
                End If;
              End Loop;
            Else
              For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                           From �Һżƻ�ʱ��
                           Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
                If Sysdate > v_ʱ��.����ʱ�� Then
                  v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                  Raise Err_Item;
                End If;
              End Loop;
            End If;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          If r_����.�ƻ�id Is Null Then
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �ҺŰ���ʱ��
                         Where ����id = r_����.Id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_ԤԼʱ����� := v_ʱ��.���;
              d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            
              Select Count(*), Max(���)
              Into n_Count, n_ԤԼ����
              From �Һ����״̬
              Where ���� = ����_In And ���� = ����ʱ��_In And ״̬ Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_ʱ��.��������, 0) And ��������_In <> 2 Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                             To_Char(v_ʱ��.����ʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.��������, 0) || '��,�����ٽ���ԤԼ�Һţ�';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          Else
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �Һżƻ�ʱ��
                         Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                               '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                               ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-11 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                        '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
              n_ԤԼʱ����� := v_ʱ��.���;
              d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            
              Select Count(*), Max(���)
              Into n_Count, n_ԤԼ����
              From �Һ����״̬
              Where ���� = ����_In And ���� = ����ʱ��_In And ״̬ Not In (4, 5);
            
              If Nvl(n_Count, 0) > Nvl(v_ʱ��.��������, 0) And ��������_In <> 2 Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                             To_Char(v_ʱ��.����ʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.��������, 0) || '��,�����ٽ���ԤԼ�Һţ�';
                Raise Err_Item;
              End If;
              n_Count := 1;
            End Loop;
          End If;
        
          If n_Count = 0 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�����û����صİ��żƻ�(' || To_Char(����ʱ��_In, 'YYYY-mm-dd HH24:MI:SS') ||
                         '),���ܽ���ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
    End If;
  
    If ������ʽ_In = 1 And ��������_In <> 2 Then
      --�ҺŹ���:
      --  �ѹ������ܴ����޺���
      If n_�ѹ��� >= Nvl(r_����.�޺���, 0) And r_����.�޺��� Is Not Null Then
        v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(r_����.�޺���, 0) || '�����ٹҺţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If ������ʽ_In > 1 Then
      --ԤԼ����ؼ��
      --����:
      --   1.����Լ���ܳ�����Լ��
      --   2.����Ƿ�����ʱ�ε�
      If n_��Լ�� >= Nvl(r_����.��Լ��, 0) And Nvl(r_����.��Լ��, 0) <> 0 And r_����.��Լ�� Is Not Null And ��������_In <> 2 Then
        v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(r_����.��Լ��, 0) || '������ԤԼ�Һţ�';
        Raise Err_Item;
      End If;
    End If;
    If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
    
      If Nvl(r_����.��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
        v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
        Raise Err_Item;
      End If; --Nvl(r_����.��ſ���, 0) =0
    
      n_��� := Case
                When Nvl(r_����.��ſ���, 0) = 1 Or n_���÷�ʱ�� = 1 And ������ʽ_In > 1 Then
                 Nvl(����_In, 0)
                Else
                 0
              End;
    
      --������λ������ģʽ
      Begin
        If Nvl(n_�ƻ�id, 0) <> 0 Then
          Select 0
          Into n_���
          From ������λ�ƻ�����
          Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
        Else
          Select 0
          Into n_���
          From ������λ���ſ���
          Where ������λ = ������λ_In And ����id = n_Tmp����id And
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
        End If;
        n_������λ������ģʽ := 1;
      Exception
        When Others Then
          n_������λ������ģʽ := 0;
      End;
      --������ż��
      For c_������λ In (Select c.���, ����
                     From �ҺŰ��� A, ������λ���ſ��� C
                     Where a.���� = ����_In And Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                   '����', '6', '����', '7', '����', Null) = c.������Ŀ(+) And a.Id = c.����id And
                           c.������λ = ������λ_In And c.��� = n_��� And Not Exists
                      (Select 1
                            From �ҺŰ��żƻ� D
                            Where d.����id = a.Id And d.���ʱ�� Is Not Null And
                                  ����ʱ��_In Between Nvl(d.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  d.ʧЧʱ��)
                     Union All
                     Select c.���, ����
                     From �ҺŰ��żƻ� A, �ҺŰ��� D, ������λ�ƻ����� C,
                          (Select Max(a.��Чʱ��) As ��Ч, ����id
                            From �ҺŰ��żƻ� A, �ҺŰ��� B
                            Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                                  ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  a.ʧЧʱ�� And b.���� = ����_In
                            Group By ����id) E
                     Where a.����id = d.Id And a.���ʱ�� Is Not Null And d.���� = ����_In And a.����id = e.����id And
                           Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                           Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) = c.������Ŀ(+) And a.Id = c.�ƻ�id And c.������λ = ������λ_In And c.��� = n_��� And
                           ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           a.ʧЧʱ��) Loop
      
        If Nvl(r_����.��ſ���, 0) = 1 And c_������λ.��� = n_��� And n_������λ������ģʽ = 0 Then
          n_�Ƿ񿪷� := 1;
          Exit;
        Elsif (Nvl(r_����.��ſ���, 0) = 0 And c_������λ.��� = n_���) Or n_������λ������ģʽ = 1 Then
          Begin
            Select Nvl(��Լ��, 0)
            Into n_ԤԼ����
            From ������λ�ҺŻ���
            Where ������λ = ������λ_In And ���� = Trunc(����ʱ��_In) And ���� = ����_In;
          Exception
            When Others Then
              n_ԤԼ���� := 0;
          End;
          If c_������λ.���� <= n_ԤԼ���� And Nvl(c_������λ.����, 0) > 0 And ��������_In <> 2 Then
            v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(c_������λ.����, 0) || '������ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
          n_�Ƿ񿪷� := 1;
          Exit;
        End If;
      
      End Loop;
    
      If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
        v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
        Raise Err_Item;
      End If;
    End If;
  
    --����޺�������Լ��
    n_�к�         := 1;
    n_ԭ��Ŀid     := 0;
    n_ԭ������Ŀid := 0;
    n_ʵ�ս��ϼ� := 0;
    If ��������_In <> 1 Then
      If ������ʽ_In <> 2 Then
        If Nvl(����id_In, 0) = 0 Then
          --����Ӧ�ó�����
          Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
        Else
          n_����id := ����id_In;
        End If;
      Else
        n_����id := Null;
      End If;
    End If;
    n_��Ŀid := r_����.��Ŀid;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      v_���� := '1|' || n_�ƻ�id;
    Else
      If Nvl(r_����.Id, 0) <> 0 Then
        v_���� := '0|' || r_����.Id;
      End If;
    End If;
    If v_���� Is Null Then
      v_���� := '3|' || ����_In;
    End If;
  
    n_������Ŀid := Zl_Custom_Getregeventitem(r_Pati.����id, r_Pati.����, r_Pati.���֤��, r_Pati.��������, r_Pati.�Ա�, r_Pati.����, v_����);
    If Nvl(n_������Ŀid, 0) <> 0 Then
      n_��Ŀid := n_������Ŀid;
    End If;
  
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := n_��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := n_��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������,
                          Nvl(a.��Ŀ����, 0) As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = r_����.��Ŀid And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������, 0 As ����
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And
                         (b.�۸�ȼ� = v_��ͨ�ȼ� Or
                         (b.�۸�ȼ� Is Null And Not Exists
                          (Select 1
                            From �շѼ�Ŀ
                            Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_��ͨ�ȼ� And d_����ʱ�� Between ִ������ And
                                  Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))
                   Order By ����, ��Ŀ����, �������) Loop
      If c_Item.���� = 1 Then
        n_���� := Nvl(c_Item.����, 0);
      End If;
      n_�۸񸸺� := Null;
      If n_ԭ��Ŀid = c_Item.��Ŀid Then
        If n_ԭ������Ŀid <> c_Item.������Ŀid Then
          n_�۸񸸺� := n_�к�;
        End If;
        n_ԭ������Ŀid := c_Item.������Ŀid;
      End If;
      n_ԭ��Ŀid := c_Item.��Ŀid;
      n_Ӧ�ս�� := Round(c_Item.���� * c_Item.����, 5);
      n_ʵ�ս�� := n_Ӧ�ս��;
      If Nvl(c_Item.���ηѱ�, 0) <> 1 And n_���ηѱ� = 0 Then
        --����:
        v_Temp     := Zl_Actualmoney(r_Pati.�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
        n_ʵ�ս�� := Zl_To_Number(v_Temp);
      End If;
      n_ʵ�ս��ϼ� := Nvl(n_ʵ�ս��ϼ�, 0) + n_ʵ�ս��;
    
      --�������ݲ���������
      If ��������_In <> 1 Then
        --�������˹Һŷ���(���ܵ����ǻ������������)
        Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
        --:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
        Insert Into ������ü�¼
          (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
           �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
           ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
        Values
          (n_����id, 4, Decode(������ʽ_In, 2, 0, 1), n_�к�, n_�۸񸸺�, c_Item.��������, ���ݺ�_In, Ʊ�ݺ�_In, 1, n_����, Null,
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, r_����.����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, n_ʵ�ս��)),
           Decode(Nvl(���ʷ���_In, 0), 1, Null, n_����id), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), r_����.����id, r_����.ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null,
           Null, ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
    Else
      d_Date := Trunc(����ʱ��_In);
    End If;
  
    --���¹Һ����״̬
    If ��������_In <> 2 Then
      n_���� := ����_In;
    End If;
    Begin
      Select 1
      Into n_Count
      From �Һ����״̬
      Where Trunc(����) = Trunc(����ʱ��_In) And ���� = ����_In And ��� = n_���� And ״̬ <> 5;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 Then
        n_���� := Null;
      End If;
      If n_���÷�ʱ�� = 1 And Nvl(r_����.��ſ���, 0) = 1 Then
        v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
        Raise Err_Item;
      End If;
    End If;
    n_Count := 0;
    If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
      If �˺�����_In = 1 Then
        Select Nvl(Max(���), 0) + 1
        Into n_����
        From �Һ����״̬
        Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ Not In (4, 5);
      Else
        Select Nvl(Max(���), 0) + 1
        Into n_����
        From �Һ����״̬
        Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ <> 5;
      End If;
    End If;
    If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
    
      If ������ʽ_In > 1 And Nvl(r_����.��ſ���, 0) = 0 Then
        --����:ԤԼʱ�����||ԤԼ��
        If Nvl(n_ԤԼ����, 0) = 0 Then
          v_Temp := Nvl(r_����.��Լ��, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_ԤԼʱ����� || v_Temp;
          n_���� := To_Number(v_Temp);
        Else
          n_���� := n_ԤԼ���� + 1;
        End If;
      End If;
    End If;
  
    If Nvl(r_����.��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
      --������ŵĴ���
      Begin
        Select ����Ա����, ������
        Into v_��Ų���Ա, v_��Ż�����
        From �Һ����״̬
        Where ״̬ = 5 And ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_����;
        n_������� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_�������   := 0;
      End;
      If n_������� = 0 Then
        Update �Һ����״̬
        Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
        Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ����Ա���� = v_����Ա����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
            Values
              (����_In, d_Date, n_����, Decode(������ʽ_In, 2, 2, 1), v_����Ա����, Decode(������ʽ_In, 1, 0, 1), Sysdate);
          
            If n_������λ���� > 0 And ������ʽ_In > 1 And Nvl(n_�Ƿ񿪷�, 0) = 1 Then
              Update ������λ�ҺŻ���
              Set ��Լ�� = ��Լ�� + Decode(������ʽ_In, 2, 1, 0), �ѽ��� = �ѽ��� + Decode(������ʽ_In, 3, 1, 0)
              Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ������λ = ������λ_In;
              If Sql%NotFound Then
                Insert Into ������λ�ҺŻ���
                  (����, ����, ���, ������λ, ��Լ��, �ѽ���)
                Values
                  (����_In, d_Date, n_����, ������λ_In, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 3, 1, 0));
              End If;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_���� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
          Raise Err_Item;
        Else
          Update �Һ����״̬
          Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
          Where ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_���� And ״̬ = 5 And ����Ա���� = v_����Ա���� And ������ = v_������;
        End If;
      End If;
    End If;
  
    If n_�����¼id Is Not Null Then
      Update �ٴ�������ſ���
      Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����Ա���� = v_����Ա����
      Where ��¼id = n_�����¼id And ��� = n_���;
      If ������ʽ_In = 2 Then
        Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
      Else
        If ������ʽ_In <> 1 Then
          Update �ٴ������¼
          Set ��Լ�� = ��Լ�� + 1, �ѹ��� = �ѹ��� + 1, �����ѽ��� = �����ѽ��� + 1
          Where ID = n_�����¼id;
        Else
          Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
        End If;
      End If;
    End If;
  
    --�������ݲ������κ� ����
    If ������ʽ_In <> 2 And ��������_In <> 1 And Nvl(���ʷ���_In, 0) = 0 Then
      --�Һ�,ԤԼ�Һ��Ѿ��ۿ��
      n_Ԥ��id := Ԥ��id_In;
      If Nvl(n_Ԥ��id, 0) = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      End If;
      n_����ϼ� := 0;
      If ���ս���_In Is Not Null Then
        --�������ս���
        v_�������� := ���ս���_In || '||';
        n_����ϼ� := 0;
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
          n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
          If Nvl(n_������, 0) <> 0 Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
            Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          End If;
          n_����ϼ� := Nvl(n_����ϼ�, 0) + Nvl(n_������, 0);
          v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
        End Loop;
      End If;
    
      If Nvl(��Ԥ��_In, 0) <> 0 Then
        Select Nvl(Sum(Nvl(Ԥ�����, 0) - Nvl(�������, 0)), 0)
        Into n_�������
        From �������
        Where ����id In (Select /*+cardinality(d,10)*/
                        d.Column_Value
                       From Table(f_Num2list(v_��Ԥ������ids)) d) And Nvl(����, 0) = 1 And Nvl(����, 0) = 1;
        If n_������� < ��Ԥ��_In Then
          v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || Ltrim(To_Char(n_�������, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                       Ltrim(To_Char(��Ԥ��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
          Raise Err_Item;
        End If;
        
        --������Ԥ��
        n_����ϼ� := n_����ϼ� + Nvl(��Ԥ��_In, 0);
        n_Ԥ����� := ��Ԥ��_In;
        For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
          n_������ := Case
                      When r_Deposit.��� - n_Ԥ����� < 0 Then
                       r_Deposit.���
                      Else
                       n_Ԥ�����
                    End;
          If r_Deposit.����id = 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (r_Deposit.����id, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          --����Ƿ��Ѿ�������
          If r_Deposit.��� <= n_������ Then
            n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
          Else
            n_Ԥ����� := 0;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
        If n_Ԥ����� > 0 Then
          v_Err_Msg := 'Ԥ������֧������֧�����,���ܼ���������';
          Raise Err_Item;
        End If;
      End If;
      --ʣ�����,��ָ�����㷽֧��
      n_������ := Nvl(n_ʵ�ս��ϼ�, 0) - Nvl(n_����ϼ�, 0);
      If Nvl(n_������, 0) < 0 Then
        v_Err_Msg := '�Һŵ���ؽ�������˵�ǰʵ����,���ܼ���������';
        Raise Err_Item;
      End If;
      If Nvl(n_������, 0) <> 0 Or (Nvl(n_������, 0) = 0 And Nvl(��Ԥ��_In, 0) = 0) Then
        If ���㷽ʽ_In Is Null Then
          v_Err_Msg := 'δ����ָ���Ľ��㷽ʽ,���ܼ���������';
          Raise Err_Item;
        End If;
      
        If Nvl(Ԥ��id_In, 0) <> 0 Then
          --�����Ԥ��ID_In��Ҫ��Ϊ�˽����������,���ҽ������վ���˸�ID,��Ҫ���µ�ID���и���,����������ת���ID
          Update ����Ԥ����¼ Set ID = n_Ԥ��id Where ID = Nvl(Ԥ��id_In, 0);
          n_Ԥ��id := Nvl(Ԥ��id_In, 0);
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ������ˮ��, ����˵��, �������, ������λ, �����id, ����,
           ��������)
        Values
          (n_Ԥ��id, 4, 1, ���ݺ�_In, r_Pati.����id, ���㷽ʽ_In, Nvl(n_������, 0), d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_����id,
           ������λ_In || '�ɿ�', n_��id, ������ˮ��_In, ����˵��_In, n_����id, ������λ_In, �����id_In, ֧������_In, 4);
      End If;
    
      --������Ա�ɿ�����
    
      For v_�ɿ� In (Select ���㷽ʽ, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ A
                   Where a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And Nvl(����id, 0) = Nvl(����id_In, 0)
                   Group By ���㷽ʽ) Loop
      
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(v_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_�ɿ�.���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, v_�ɿ�.���㷽ʽ, 1, Nvl(v_�ɿ�.��Ԥ��, 0));
          n_����ֵ := Nvl(v_�ɿ�.��Ԥ��, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = ���㷽ʽ_In And ���� = 1 And Nvl(���, 0) = 0;
        End If;
      
      End Loop;
    
    End If;
  
    --����Һż�¼
    If ��������_In = 2 Then
      Begin
        Select ID Into n_�Һ�id From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    Else
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    End If;
  
    Update ���˹Һż�¼
    Set ��¼���� = Decode(������ʽ_In, 2, 2, 1), ��¼״̬ = Decode(��������_In, 1, 0, 1), ����� = r_Pati.�����, ����Ա���� = v_����Ա����,
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
        ����ʱ�� = Case ��������_In
                  When 1 Then
                   Null
                  Else
                   Case ������ʽ_In
                     When 2 Then
                      Null
                     Else
                      d_�Ǽ�ʱ��
                   End
                End, ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), ������λ = Nvl(������λ_In, ������λ),
        ԤԼ����Ա = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա, Null), Nvl(ԤԼ����Ա, v_����Ա����)),
        ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���))
    Where ID = n_�Һ�id;
    If Sql%NotFound Then
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = r_Pati.���ʽ And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
         ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���)
      Values
        (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
         r_Pati.�Ա�, r_Pati.����, ����_In, n_����, v_����, Null, r_����.����id, r_����.ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
         Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���));
    End If;
    --�������ݲ��ܲ�������
    If ��������_In <> 1 Then
      n_ԤԼ���ɶ��� := 0;
      If ������ʽ_In > 1 Then
        n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
      End If;
      --�Һź��շѵ�ԤԼ��ֱ�ӽ������(�շ�ԤԼȱ�ٽ��չ���,����ֱ�Ӻ͹Һ�һ��ֱ�ӽ������)
      If ������ʽ_In <> 2 Or n_ԤԼ���ɶ��� = 1 Then
        If Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113)) <> 0 Then
          --�Ŷӽк�ģʽ:-0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
          If Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(r_����.����id, 0))) = 0 Or n_ԤԼ���ɶ��� = 1 Then
            n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(������ʽ_In, 0) > 1 And n_��ʱ����ʾ = 1 And n_���÷�ʱ�� = 1 Then
              n_��ʱ����ʾ := 1;
            Else
              n_��ʱ����ʾ := Null;
            End If;
            --��������
            --.����ִ�в��š� �ķ�ʽ���ɶ���
            v_�������� := r_����.����id;
            v_�ŶӺ��� := Zlgetnextqueue(r_����.����id, n_�Һ�id, ����_In || '|' || ����_In);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
            --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
            Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, r_����.����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, r_����.ҽ������,
                             d_�Ŷ�ʱ��, ԤԼ��ʽ_In, n_��ʱ����ʾ, v_�Ŷ����);
          End If;
        End If;
      End If;
    
      If Nvl(������ʽ_In, 0) = 1 Then
        --����Ʊ��ʹ�����
        If Ʊ�ݺ�_In Is Not Null Then
          Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
          --����Ʊ��
          Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Values
            (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�Ǽ�ʱ��, v_����Ա����, �ҺŽ��ϼ�_In);
          --״̬�Ķ�
          Update Ʊ�����ü�¼
          Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
          Where ID = Nvl(����id_In, 0);
        End If;
        --���˱��ξ���(�Է���ʱ��Ϊ׼)
        If Nvl(r_Pati.����id, 0) <> 0 Then
          Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = v_���� Where ����id = r_Pati.����id;
        End If;
      End If;
    End If;
    --���˹ҺŻ���
    --��������ʱ�����ٶԻ��ܵ��ݽ���ͳ���� ����������ʱ�Ѿ������˻���
    If ��������_In <> 2 Then
      --������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
      --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
      --����zl_third_lockno�������ţ�������ʹ�ñ���������
      n_ԤԼ := Case
                When Nvl(������ʽ_In, 0) = 1 Then
                 0
                When Nvl(������ʽ_In, 0) = 2 Then
                 1
                When Nvl(������ʽ_In, 0) = 3 Then
                 3
                Else
                 0
              End;
      Zl_���˹ҺŻ���_Update(r_����.ҽ������, r_����.ҽ��id, r_����.��Ŀid, r_����.����id, ����ʱ��_In, n_ԤԼ, ����_In);
    End If;
  
    If ��������_In <> 1 Then
      --��Ϣ����,����ʱ��������Ϣ
      Begin
        Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
          Using 1, n_�Һ�id;
      Exception
        When Others Then
          Null;
      End;
      b_Message.Zlhis_Regist_001(n_�Һ�id, ���ݺ�_In);
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Insert;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0046' Where ���=&n_System;
Commit;
