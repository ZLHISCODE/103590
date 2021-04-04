----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.35.90������ v10.35.90
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
----------------------------------------------------------------------------------------------------------------



------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--126717:������,2018-07-06,΢����PDF�����ӡ
Alter Table ҽ���������� Add ��ӡ���� Number(5); 




------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--127912:��ΰ��,2018-07-03,ԤԼ������ס����סȡ�������ַ����
Insert Into ������������Ŀ¼ (ϵͳ��ʶ, ��������) Values ('ԤԼ����', '��ס����סȡ��');

--127796,��¶¶,2018-07-04,�������¼�����Բ�¼�벡���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval,&n_System,-Null,-Null,-Null,-Null,-Null,A.* From (
  Select ����,����,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵�� From zlParameters Where 1 = 0
  Union All Select 0, 0, 307, '����д�����','0', '0','������ҳ����ʱ��д������Ϻ��Ƿ���Բ���д�����', '0-��д������Ϻ������д����ţ�1-��д������Ϻ���Բ���д�����', '', '�����ڲ�����ҳ����ʱ', Null
  From Dual)A;


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--126717:������,2018-07-06,΢����PDF�����ӡ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1252,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_ҽ����������_Print','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1253,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_ҽ����������_Print','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--128177:Ƚ����,2018-07-06,����ƽ̨�ҺŽӿ�֧��ʹ��Ԥ����
Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:HIS�Һ� 
  --���:Xml_In: 
  --<IN> 
  --   <CZFS>3</CZFS>    //������ʽ 
  --   <CZJLID>1</CZJLID>    //�����¼ID 
  --   <HM>����</HM>    //���� 
  --   <HX>����</HX>     //���� 
  --   <JKFS>0</JKFS>  //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ� 
  --   <YYSJ>2014-10-21 </YYSJ>    //ԤԼ���� YYYY-MM-DD,��ʱ�η���ſ�����Ҫ����ʱ�� 
  --   <JE>���</JE>     //��� 
  --   <JSLIST> 
  --     <JS>            //������Ϣ���Һŷ�ҽ������Ŀǰ��֧��һ�����ṹ���շ�һ�� 
  --       <JSKLB>���㿨���</JSKLB>    //���㿨��� 
  --       <JSKH>֧�����ʺ�</JSKH>           //���㿨��(֧�����ʺ�) 
  --       <JYSM>����˵��</JYSM>            //˵�����̶���֧���� 
  --       <JYLSH>��ˮ��</JYLSH>           //��ˮ�ţ��������� 
  --       <JSFS>���㷽ʽ</JSFS>            //���㷽ʽ:�ֽ�֧Ʊ�������������,���Դ��� 
  --       <JSJE>������</JSJE>            //������ 
  --       <ZY>ժҪ</ZY>                  //ժҪ 
  --       <SFCYJ></SFCYJ>              //�Ƿ��Ԥ�����Һ�Ŀǰ���� 
  --       <SFXFK></SFXFK>              //�Ƿ����ѿ�,�Һ�Ŀǰ���� 
  --       <EXPENDLIST>                 //��չ��Ϣ 
  --         <EXPEND> 
  --           <JYMC>��������</JYMC>        //�������� 
  --           <JYLR>��������<JYLR>         //�������� 
  --         </EXPEND> 
  --         <EXPEND> 
  --           ... 
  --         </EXPEND> 
  --       </EXPENDLIST> 
  --     </JS> 
  --   </JSLIST> 
  --   <HZDW>������λ</HZDW>        //������λ���� 
  --   <YYFS>֧����<YYFS>    //ԤԼ��ʽ,����������֧���� 
  --   <BRID>����ID</BRID>     //����ID 
  --   <SFZH>���֤��</SFZH>     //���֤�� 
  --   <XM>����</XM>            //���� 
  --   <BRLX></BRLX>             //ҽ���������� 
  --   <FB>��ͨ</FB>               //���˷ѱ𣬿��Բ��� 
  --   <JQM>������</JQM>            //������ 
  --</IN> 

  --����:Xml_Out 
  --<OUTPUT> 
  -- <GHDH>�Һŵ���</GHDH>          //�Һŵ��� 
  -- <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ�� 
  -- <JZID>����ID</JZID>          //���ν���ID 
  -- <ERROR><MSG>������Ϣ</MSG></ERROR>  //����ʱ���� 
  --</ OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  v_����     �ҺŰ���.����%Type;
  d_����ʱ�� Date;
  d_ԭʼʱ�� Date;
  d_�Ǽ�ʱ�� Date;

  n_Ӧ�ս��   ������ü�¼.Ӧ�ս��%Type;
  v_��ˮ��     ����Ԥ����¼.������ˮ��%Type;
  v_˵��       ������ü�¼.ժҪ%Type;
  n_����id     ������Ϣ.����id%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  v_ԤԼ��ʽ   ԤԼ��ʽ.����%Type;
  v_��������� ҽ�ƿ����.����%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  n_�����     ������ü�¼.��ʶ��%Type;
  v_����       ������ü�¼.����%Type;
  v_�Ա�       ������ü�¼.�Ա�%Type;
  v_����       ������ü�¼.����%Type;
  v_���ʽ   ������ü�¼.���ʽ%Type;
  v_�ѱ�       ������ü�¼.�ѱ�%Type;
  v_No         ���˹Һż�¼.No%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_�շ�ϸĿid ������ü�¼.�շ�ϸĿid%Type;
  n_���˿���id ������ü�¼.���˿���id%Type;
  n_��������id ������ü�¼.��������id%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_ҽ������   �ҺŰ���.ҽ������%Type;
  n_ҽ��id     �ҺŰ���.ҽ��id%Type;
  n_����id     ������ü�¼.����id%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_�Ű�       �ҺŰ���.����%Type;
  n_����id     �ҺŰ���.Id%Type;
  n_�ƻ�id     �ҺŰ��żƻ�.Id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_��ſ���   �ҺŰ���.��ſ���%Type;
  n_����       �Һ����״̬.���%Type;
  v_����       �ҺŰ�������.������Ŀ%Type;
  v_��������   ������Ϣ.��������%Type;
  n_����       Number(3);
  v_�ֽ�       ���㷽ʽ.����%Type;
  n_��ʱ��     Number(3);
  v_��������   Varchar2(3000);
  v_������λ   ���˹Һż�¼.������λ%Type;
  v_������     �Һ����״̬.������%Type;
  n_�ɿʽ   Number(3);
  n_�Һ�ģʽ   Number(3);
  n_Exists     Number(3);
  v_���ս���   Varchar2(1000);
  n_��¼id     �ٴ������¼.Id%Type;
  v_Temp       Varchar2(32767); --��ʱXML 
  x_Templet    Xmltype; --ģ��XML 
  v_Err_Msg    Varchar2(200);
  d_����ʱ��   Date;
  n_Count      Number(3);
  v_�����     �������׼�¼.���%Type;
  n_��Ԥ��     ����Ԥ����¼.��Ԥ��%Type;
  v_Para       Varchar2(2000);
  Err_Item Exception;
  Err_Special Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS'), Extractvalue(Value(A), 'IN/CZJLID'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into v_����, n_����, d_ԭʼʱ��, n_Ӧ�ս��, v_ԤԼ��ʽ, v_������λ, n_����id, v_��������, v_�ѱ�, v_������, n_�ɿʽ, n_��¼id, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����!';
    Raise Err_Item;
  End If;

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  If Sysdate - 10 > Nvl(d_����ʱ��, Sysdate - 30) Then
    If n_�Һ�ģʽ = 1 And Nvl(d_ԭʼʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_��¼id Is Null Then
      v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
      Raise Err_Item;
    End If;
  Else
    If n_�Һ�ģʽ = 1 And Nvl(d_ԭʼʱ��, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_��¼id Is Null Then
      Begin
        Select a.Id
        Into n_��¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_���� And Nvl(d_ԭʼʱ��, Sysdate) Between a.��ʼʱ�� And a.��ֹʱ��;
      Exception
        When Others Then
          v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  d_�Ǽ�ʱ�� := Sysdate;
  d_����ʱ�� := Trunc(d_ԭʼʱ��);

  For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    --��Ԥ������Ҫ���������� 
    If Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 0 Then
      If c_���׼�¼.���㿨��� Is Null Then
        v_����� := c_���׼�¼.���㷽ʽ;
      Else
        Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
      
        If Nvl(n_Count, 0) = 1 Then
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
        Else
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
        End If;
      End If;
    
      If v_����� Is Null Then
        v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 4) = 0 Then
        v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  If v_�������� Is Not Null Then
    Begin
      Select 1 Into n_���� From �������� Where ���� = v_��������;
    Exception
      When Others Then
        v_Err_Msg := 'û�з���Ϊ(' || v_�������� || ')�Ĳ�������';
        Raise Err_Item;
    End;
    Update ������Ϣ Set �������� = Nvl(��������, v_��������) Where ����id = n_����id;
  End If;

  Select a.�����, a.����, a.�Ա�, a.����, Nvl(b.����, c.����)
  Into n_�����, v_����, v_�Ա�, v_����, v_���ʽ
  From ������Ϣ A, ҽ�Ƹ��ʽ B, (Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = '1' And Rownum < 2) C
  Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = b.����(+);

  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_��������id From Dual;

  v_No := Nextno(12);
  Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;

  If n_��¼id Is Null Then
    For r_���� In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_����.�Ƿ��Ԥ��, 0) = 0 Then
        If r_����.���㷽ʽ Is Null Then
          Begin
            Select b.���㷽ʽ, b.Id
            Into v_���㷽ʽ, n_�����id
            From ҽ�ƿ���� B
            Where b.���� = r_����.���㿨��� And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
              Raise Err_Item;
          End;
        Else
          Select Nvl(Max(1), 0) Into n_Exists From ���㷽ʽ Where ���� = r_����.���㷽ʽ And ���� In (3, 4);
          If n_Exists = 1 Then
            v_���ս��� := v_���ս��� || '||' || r_����.���㷽ʽ || '|' || r_����.������;
          Else
            If v_���㷽ʽ Is Null Then
              v_���㷽ʽ := r_����.���㷽ʽ;
            Else
              v_Err_Msg := 'Ŀǰ�ƻ��Ű�ҺŲ�֧�ַ�ҽ����Ķ��ֽ��㷽ʽ,����!';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        If r_����.���㿨��� Is Not Null Then
          v_��������� := r_����.���㿨���;
          v_���㿨��   := r_����.���㿨��;
          v_��ˮ��     := r_����.������ˮ��;
          v_˵��       := r_����.����˵��;
        
          If n_�����id Is Null Then
            Begin
              Select b.Id Into n_�����id From ҽ�ƿ���� B Where b.���� = r_����.���㿨��� And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_����.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(r_����.���㿨���);
          Else
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = r_����.���㿨���;
          End If;
        Else
          v_����� := r_����.���㷽ʽ;
        End If;
      
        Update �������׼�¼
        Set ҵ�����id = n_����id
        Where ��ˮ�� = r_����.������ˮ�� And ��� = v_����� And ҵ������ = 4;
      Else
        n_��Ԥ�� := r_����.������;
      End If;
    End Loop;
  
    If v_���ս��� Is Not Null Then
      v_���ս��� := Substr(v_���ս���, 3);
    End If;
  
    Select Decode(To_Char(d_ԭʼʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                   Null)
    Into v_����
    From Dual;
  
    Begin
      Select ID
      Into n_�ƻ�id
      From (Select ID
             From �ҺŰ��żƻ�
             Where ���� = v_���� And d_ԭʼʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                   Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And ���ʱ�� Is Not Null
             Order By ��Чʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Select ID Into n_����id From �ҺŰ��� Where ���� = v_����;
    End;
  
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      --�Ӽƻ���ȡ��Ϣ 
      Select a.��Ŀid, b.����id, a.ҽ������, a.ҽ��id,
             Decode(To_Char(d_����ʱ��, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7', a.����,
                     Null), Nvl(a.��ſ���, 0)
      Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
      From �ҺŰ��żƻ� A, �ҺŰ��� B
      Where a.Id = n_�ƻ�id And b.Id = a.����id;
      Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum < 2;
    
      --������λ��� 
      If v_������λ Is Not Null Then
        Begin
          Select 1 Into n_���� From ������λ�ƻ����� Where �ƻ�id = n_�ƻ�id And ���� = 0 And ������λ = v_������λ;
        Exception
          When Others Then
            n_���� := 0;
        End;
      End If;
    
      If n_���� = 1 Then
        v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
        Raise Err_Item;
      End If;
    
      If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
        d_����ʱ�� := d_ԭʼʱ��;
        Select ���
        Into n_����
        From �Һżƻ�ʱ��
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_����ʱ��
          From �Һżƻ�ʱ��
          Where �ƻ�id = n_�ƻ�id And ���� = v_���� And ��� = Nvl(n_����, 0);
        Exception
          When Others Then
            If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From �Һżƻ�ʱ��
              Where �ƻ�id = n_�ƻ�id And ���� = v_����;
            Else
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From ʱ���
              Where ʱ��� = v_�Ű�;
            End If;
            If d_����ʱ�� < d_�Ǽ�ʱ�� Then
              d_����ʱ�� := d_�Ǽ�ʱ��;
            End If;
        End;
      End If;
    Else
      --�Ӱ��Ŷ�ȡ��Ϣ 
      Select b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id,
             Decode(To_Char(d_����ʱ��, 'D'), '1', b.����, '2', b.��һ, '3', b.�ܶ�, '4', b.����, '5', b.����, '6', b.����, '7', b.����,
                     Null), Nvl(b.��ſ���, 0)
      Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
      From �ҺŰ��� B
      Where b.Id = n_����id;
      Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum < 2;
    
      --������λ��� 
      If v_������λ Is Not Null Then
        Begin
          Select 1 Into n_���� From ������λ���ſ��� Where ����id = n_����id And ���� = 0 And ������λ = v_������λ;
        Exception
          When Others Then
            n_���� := 0;
        End;
      End If;
    
      If n_���� = 1 Then
        v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
        Raise Err_Item;
      End If;
    
      If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
        d_����ʱ�� := d_ԭʼʱ��;
        Select ���
        Into n_����
        From �ҺŰ���ʱ��
        Where ����id = n_����id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
      Else
        Begin
          Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
          Into d_����ʱ��
          From �ҺŰ���ʱ��
          Where ����id = n_����id And ���� = v_���� And ��� = Nvl(n_����, 0);
        Exception
          When Others Then
            If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                              'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From �ҺŰ���ʱ��
              Where ����id = n_����id And ���� = v_����;
            Else
              Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
              Into d_����ʱ��
              From ʱ���
              Where ʱ��� = v_�Ű�;
            End If;
            If d_����ʱ�� < d_�Ǽ�ʱ�� Then
              d_����ʱ�� := d_�Ǽ�ʱ��;
            End If;
        End;
      End If;
    End If;
  
    If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(3, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      End If;
    Else
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(1, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1);
      End If;
    End If;
  
    For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --������Ű�ģʽ 
    For r_���� In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_����.�Ƿ��Ԥ��, 0) = 0 Then
        If r_����.���㷽ʽ Is Null Then
          Begin
            Select b.���㷽ʽ, b.Id
            Into v_���㷽ʽ, n_�����id
            From ҽ�ƿ���� B
            Where b.���� = r_����.���㿨��� And Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
              Raise Err_Item;
          End;
          v_�������� := v_�������� || '|' || v_���㷽ʽ || ',' || r_����.������ || ',,';
        Else
          v_�������� := v_�������� || '|' || r_����.���㷽ʽ || ',' || r_����.������ || ',,';
        End If;
      
        If r_����.���㿨��� Is Not Null Then
          v_��������   := v_�������� || '1';
          v_��������� := r_����.���㿨���;
          v_���㿨��   := r_����.���㿨��;
          v_��ˮ��     := r_����.������ˮ��;
          v_˵��       := r_����.����˵��;
          If n_�����id Is Null Then
            Begin
              Select b.Id Into n_�����id From ҽ�ƿ���� B Where b.���� = r_����.���㿨��� And Rownum < 2;
            Exception
              When Others Then
                v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
                Raise Err_Item;
            End;
          End If;
        
          Select Decode(Translate(Nvl(r_����.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(r_����.���㿨���);
          Else
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = r_����.���㿨���;
          End If;
        Else
          v_�������� := v_�������� || '0';
          v_�����   := r_����.���㷽ʽ;
        End If;
      
        Update �������׼�¼
        Set ҵ�����id = n_����id
        Where ��ˮ�� = r_����.������ˮ�� And ��� = v_����� And ҵ������ = 4;
      Else
        n_��Ԥ�� := r_����.������;
      End If;
    End Loop;
  
    If v_�������� Is Not Null Then
      v_�������� := Substr(v_��������, 2);
    Else
      Begin
        Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
      Exception
        When Others Then
          v_�ֽ� := '�ֽ�';
      End;
      v_�������� := v_�ֽ� || ',' || 0 || ',,0';
    End If;
  
    Select ��Ŀid, ����id, ҽ������, ҽ��id, �Ƿ���ſ���, �Ƿ��ʱ��
    Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, n_��ſ���, n_��ʱ��
    From �ٴ������¼
    Where ID = n_��¼id;
  
    Begin
      Select ��ʼʱ�� Into d_����ʱ�� From �ٴ�������ſ��� Where ��¼id = n_��¼id And ��� = n_����;
    Exception
      When Others Then
        d_����ʱ�� := d_ԭʼʱ��;
    End;
  
    If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(3, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      End If;
    Else
      If Nvl(n_�ɿʽ, 0) = 0 Then
        Zl_���������Һ�_Insert(1, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      Else
        Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_��������, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                         v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, v_���ս���, n_��Ԥ��, v_���㿨��, 1, v_�ѱ�, Null,
                         v_������, 1, 0, n_��¼id);
      End If;
    End If;
  
    For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                          Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
    End Loop;
  
    v_Temp := '<GHDH>' || v_No || '</GHDH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<JZID>' || n_����id || '</JZID>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

--126717:������,2018-07-06,΢����PDF�����ӡ
Create Or Replace Procedure Zl_ҽ����������_Print
(
  ����id_In In ҽ����������.Id%Type,
  ����_In   In Number
) Is
  --����_In:0-��ʾ��ӡ��
Begin
  If ����_In = 0 Then
    Update ҽ���������� Set ��ӡ���� = Nvl(��ӡ����, 0) + 1 Where ID = ����id_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҽ����������_Print;
/

--128391:����,2018-07-05,�������״̬Ϊ2�����
CREATE OR REPLACE PROCEDURE Zl_ҽ����˹���_Cancel
(
  ҽ��ids_In  VARCHAR2,
  ��˶���_In NUMBER := 1, --1=����ҽ����2=��Ѫҽ��
  ִ�����_In NUMBER := 0 --0=�ϰ�Ѫ�����̣���Ϊ0ʱ����ΪĿ�����״̬��1=����ˣ�7=��ǩ����4-��ǩ����3-�Ѿܾ���
) IS
  --ȡ�����
  CURSOR c_Advice IS
    SELECT * FROM TABLE(CAST(f_Num2list(ҽ��ids_In) AS t_Numlist));
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_ҽ��״̬ NUMBER;
  n_���״̬ NUMBER;
  n_�������� NUMBER;
BEGIN
  FOR r_Advice IN c_Advice LOOP
    SELECT COUNT(1), MAX(ҽ��״̬), Nvl(MAX(���״̬), 0)
    INTO n_Count, n_ҽ��״̬, n_���״̬
    FROM ����ҽ����¼
    WHERE Id = r_Advice.Column_Value;
  
    IF n_Count = 0 THEN
      v_Err_Msg := '��ҽ���Ѿ�ɾ��,���֤��';
      RAISE Err_Item;
    END IF;
  
    IF n_ҽ��״̬ <> 1 THEN
      v_Err_Msg := '��ѡ���ҽ���а�����У�Ե�ҽ��������ȡ����ˡ�';
      RAISE Err_Item;
    END IF;
  
    IF n_���״̬ = 1 THEN
      n_�������� := 19;
    ELSIF n_���״̬ = 7 THEN
      n_�������� := 18;
    ELSIF n_���״̬ = 3 THEN
      n_�������� := 12;
    ELSIF n_���״̬ = 4 OR n_���״̬ = 2 THEN
      n_�������� := 11;
    END IF;
  
    IF ��˶���_In = 1 OR ִ�����_In = 0 THEN
      UPDATE ����ҽ����¼ SET ���״̬ = 1 WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value;
      DELETE FROM ����ҽ��״̬
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value) AND
            �������� IN (11, 12) AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = r_Advice.Column_Value AND �������� IN (11, 12));
    ELSIF ��˶���_In = 2 AND ִ�����_In <> 0 THEN
      UPDATE ����ҽ����¼
      SET ���״̬ = ִ�����_In
      WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value;
      DELETE FROM ����ҽ��״̬
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value) AND
            �������� = n_�������� AND
            ����ʱ�� =
            (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = r_Advice.Column_Value AND �������� = n_��������);
    END IF;
  
  END LOOP;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_ҽ����˹���_Cancel;
/

--128391:����,2018-07-05,�������״̬Ϊ2�����
CREATE OR REPLACE PROCEDURE Zl_ҽ����˹���_Update
(
  ҽ��id_In   ����ҽ��״̬.ҽ��id%TYPE,
  ����ʱ��_In ����ҽ��״̬.����ʱ��%TYPE,
  ����˵��_In ����ҽ��״̬.����˵��%TYPE := NULL,
  ��˶���_In NUMBER := 1, --1=����ҽ����2=��Ѫҽ��
  ������Ա_In VARCHAR2 := NULL
) IS
  --�޸�ֻ��������˲�ͨ����ҽ�����޸������˵��
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_���״̬ NUMBER;
BEGIN
  SELECT COUNT(1), Nvl(MAX(���״̬), 0) INTO n_Count, n_���״̬ FROM ����ҽ����¼ WHERE Id = ҽ��id_In;
  IF n_Count = 0 THEN
    v_Err_Msg := '��ҽ���Ѿ�ɾ��,���֤��';
    RAISE Err_Item;
  END IF;

  IF ��˶���_In = 1 THEN
    UPDATE ����ҽ��״̬
    SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
    WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 12 AND
          ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 12);
  ELSE
    IF n_���״̬ = 1 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 19 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 19);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 19, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
    ELSIF n_���״̬ = 7 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 18 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 18);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 18, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
    ELSIF n_���״̬ = 3 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 12 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 12);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 12, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
    ELSIF n_���״̬ = 4 OR n_���״̬ = 2 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 11 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 11);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 11, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
    END IF;
  END IF;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_ҽ����˹���_Update;
/

--127912:��ΰ��,2018-07-03,ԤԼϵͳ����
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
  v_Error    Varchar2(2000);

  Err_Custom Exception;
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
    Where ����id = r_Pati.����id And ��ҳid = 0;
    --������鴲λ�Ƿ�Ϊ��
    Select Count(*) Into n_Count From ��λ״����¼ Where ����id = n_����id And ���� = v_���� And ״̬ = '�մ�';
    If n_Count = 0 Then
      v_Error := '����ʧ��,��λ ' || v_���� || ' ���ǿմ���' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�º����ԣ�';
      Raise Err_Custom;
    End If;
    --����λ����ռ��
    Update ��λ״����¼
    Set ״̬ = 'ռ��', ����id = r_Pati.����id, ����id = Decode(����, 1, n_����id, ����id)
    Where ����id = n_����id And ���� = v_����;
  Else
    --ȡ���Ǽ�
    Select b.����id Into n_����id From ������ҳ B Where b.�Һ�id = n_�Һ�id;
    Zl_��Ժ������ҳ_Delete(n_����id, 0);
  End If;
  Xml_Out := Xmltype('<OUTPUT><RESULT>true</RESULT></OUTPUT>');
Exception
  When Err_Custom Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || v_Error || '</MSG></ERROR></OUTPUT>');
  When Others Then
    Xml_Out := Xmltype('<OUTPUT><RESULT>false</RESULT><ERROR><MSG>' || SQLCode || '***' || SQLErrM ||
                       '</MSG></ERROR></OUTPUT>');
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Outpatireg;
/

--127912:��ΰ��,2018-07-03,ԤԼ���˽���
Create Or Replace Procedure Zl_��Ժ������ҳ_Insert
(
  �Ǽ�ģʽ_In       Number,
  ��������_In       ������ҳ.��������%Type,
  ����id_In         ������Ϣ.����id%Type,
  סԺ��_In         ������Ϣ.סԺ��%Type,
  ҽ����_In         �����ʻ�.ҽ����%Type,
  ����_In           ������Ϣ.����%Type,
  �Ա�_In           ������Ϣ.�Ա�%Type,
  ����_In           ������Ϣ.����%Type,
  �ѱ�_In           ������Ϣ.�ѱ�%Type,
  ��������_In       ������Ϣ.��������%Type,
  ����_In           ������Ϣ.����%Type,
  ����_In           ������Ϣ.����%Type,
  ѧ��_In           ������Ϣ.ѧ��%Type,
  ����״��_In       ������Ϣ.����״��%Type,
  ְҵ_In           ������Ϣ.ְҵ%Type,
  ���_In           ������Ϣ.���%Type,
  ���֤��_In       ������Ϣ.���֤��%Type,
  �����ص�_In       ������Ϣ.�����ص�%Type,
  ��ͥ��ַ_In       ������Ϣ.��ͥ��ַ%Type,
  ��ͥ��ַ�ʱ�_In   ������Ϣ.��ͥ��ַ�ʱ�%Type,
  ��ͥ�绰_In       ������Ϣ.��ͥ�绰%Type,
  ���ڵ�ַ_In       ������Ϣ.���ڵ�ַ%Type,
  ���ڵ�ַ�ʱ�_In   ������Ϣ.���ڵ�ַ�ʱ�%Type,
  ��ϵ������_In     ������Ϣ.��ϵ������%Type,
  ��ϵ�˹�ϵ_In     ������Ϣ.��ϵ�˹�ϵ%Type,
  ��ϵ�˵�ַ_In     ������Ϣ.��ϵ�˵�ַ%Type,
  ��ϵ�˵绰_In     ������Ϣ.��ϵ�˵绰%Type,
  ������λ_In       ������Ϣ.������λ%Type,
  ��ͬ��λid_In     ������Ϣ.��ͬ��λid%Type,
  ��λ�绰_In       ������Ϣ.��λ�绰%Type,
  ��λ�ʱ�_In       ������Ϣ.��λ�ʱ�%Type,
  ��λ������_In     ������Ϣ.��λ������%Type,
  ��λ�ʺ�_In       ������Ϣ.��λ�ʺ�%Type,
  ������_In         ������Ϣ.������%Type,
  ������_In         ������Ϣ.������%Type,
  ��������_In       ������Ϣ.��������%Type,
  ��Ժ����id_In     ������ҳ.��Ժ����id%Type,
  ����ȼ�id_In     ������ҳ.����ȼ�id%Type,
  ��Ժ����_In       ������ҳ.��Ժ����%Type,
  ��Ժ��ʽ_In       ������ҳ.��Ժ��ʽ%Type,
  סԺĿ��_In       ������ҳ.סԺĿ��%Type,
  ����Ժת��_In     ������ҳ.����Ժת��%Type,
  ����ҽʦ_In       ������ҳ.����ҽʦ%Type,
  ����_In           ������Ϣ.����%Type,
  ����_In           ������ҳ.����%Type,
  ��Ժʱ��_In       ������ҳ.��Ժ����%Type,
  �Ƿ����_In       ������ҳ.�Ƿ����%Type,
  ����_In           ������ҳ.��Ժ����%Type,
  ���ʽ_In       ������ҳ.ҽ�Ƹ��ʽ%Type,
  ����id_In         ������ϼ�¼.����id%Type,
  ���id_In         ������ϼ�¼.���id%Type,
  �������_In       ������ϼ�¼.�������%Type,
  ��ҽ����id_In     ������ϼ�¼.����id%Type,
  ��ҽ���id_In     ������ϼ�¼.���id%Type,
  ��ҽ���_In       ������ϼ�¼.�������%Type,
  ����_In           ������ҳ.����%Type,
  ����Ա���_In     ������ҳ.��ĿԱ���%Type,
  ����Ա����_In     ������ҳ.��ĿԱ����%Type,
  �²���_In         Number := 1,
  ��ע_In           ������ҳ.��ע%Type,
  ��Ժ����id_In     ������ҳ.��Ժ����id%Type,
  ����Ժ_In         ������ҳ.����Ժ%Type,
  ��Ժ����_In       ������ҳ.��Ժ����%Type := Null,
  ��ҳid_In         ������ҳ.��ҳid%Type := Null,
  סԺ����_In       ������Ϣ.סԺ����%Type := Null,
  ����֤��_In       ������Ϣ.����֤��%Type := Null,
  ��������_In       ������ҳ.��������%Type := Null,
  ��ϵ�����֤��_In ������Ϣ.��ϵ�����֤��%Type := Null,
  �ֻ���_In         ������Ϣ.�ֻ���%Type := Null,
  �Һ�id_In         ������ҳ.�Һ�id%Type := Null
) As
  -----------------------------------------------------------
  --���ܣ�����Ժ��������һ�Ų�����ҳ��ͬʱ���ܴ�����ơ�
  --������
  --      �Ǽ�ģʽ_IN=0-�����Ǽ�,1-ԤԼ�Ǽ�,2-����ԤԼ(�²���_IN=0)
  --      ��������_IN=��Ӧ"������ҳ.��������"
  --      ����_IN=Null:��ͬʱ���;'��ͥ����':�����ͥ����,��Ϊ��;����:������崲λ��
  --      �²���_IN=��������е����Ĳ�����Ժ,��ò���Ϊ0��ȱʡΪ�²���
  --      ��Ժ����ID_IN=ֻ�е�ʹ��[����������]ģʽ(������99)ʱ,������Ժͬʱ��Ʒִ�ʱ,����ֵ
  --      סԺ��_In = �Ǽ��������۲���ʱ סԺ��_In Ϊ���������
  -----------------------------------------------------------
  v_��ҳid   ������ҳ.��ҳid%Type;
  v_�ȼ�id   ��λ״����¼.�ȼ�id%Type;
  n_סԺ���� ������Ϣ.סԺ����%Type;

  v_�ѱ�      ������ҳ.�ѱ�%Type;
  v_����      ������ҳ.��Ժ����%Type;
  v_Count     Number;
  n_Uniqueid  Number;
  v_Date      Date;
  d_Indeptime Date;
  v_Error     Varchar2(255);
  Err_Custom Exception;
Begin
  --�жϲ����Ƿ�����
  Select Count(����id) Into v_Count From ������Ϣ Where ����id = ����id_In;
  If v_Count <> 0 Then
    Zl_������Ϣ_�������(����id_In);
  End If;

  Select Sysdate Into v_Date From Dual;
  Zl_������Ǽ�¼_Clear(����id_In);

  --���֤�Ų����ڿ�,����ϵͳ�����ж��Ƿ�Ψһ��������
  If ���֤��_In Is Not Null Then
    n_Uniqueid := Nvl(zl_GetSysParameter(279), 0);
    If n_Uniqueid = 1 Then
      Select Count(1) Into v_Count From ������Ϣ Where ���֤�� = ���֤��_In And ����id <> Nvl(����id_In, 0);
      If v_Count <> 0 Then
        v_Error := '�Ѿ��������֤��Ϊ' || ���֤��_In || '�Ĳ���,������¼����ͬ�����֤��!';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  --���˻�����Ϣ
  If ��������_In = 1 Then
    If �²���_In = 1 Then
      Insert Into ������Ϣ
        (����id, �����, סԺ��, ����, �Ա�, ����, �ѱ�, ҽ�Ƹ��ʽ, ��������, ����, ����, ����, ����, ѧ��, ����״��, ְҵ, ���, ���֤��, �����ص�, ��ͥ��ַ, ��ͥ��ַ�ʱ�, ��ͥ�绰,
         ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ������λ, ��ͬ��λid, ��λ�绰, ��λ�ʱ�, ��λ������, ��λ�ʺ�, ������, ������, ��������, ����, �Ǽ�ʱ��, ����֤��,
         ��������, ��ϵ�����֤��, �ֻ���)
      Values
        (����id_In, סԺ��_In, Null, ����_In, �Ա�_In, ����_In, �ѱ�_In, ���ʽ_In, ��������_In, ����_In, ����_In, ����_In, ����_In, ѧ��_In,
         ����״��_In, ְҵ_In, ���_In, ���֤��_In, �����ص�_In, ��ͥ��ַ_In, ��ͥ��ַ�ʱ�_In, ��ͥ�绰_In, ���ڵ�ַ_In, ���ڵ�ַ�ʱ�_In, ��ϵ������_In, ��ϵ�˹�ϵ_In,
         ��ϵ�˵�ַ_In, ��ϵ�˵绰_In, ������λ_In, Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In), ��λ�绰_In, ��λ�ʱ�_In, ��λ������_In, ��λ�ʺ�_In, ������_In,
         Decode(������_In, 0, Null, ������_In), ��������_In, ����_In, v_Date, ����֤��_In, ��������_In, ��ϵ�����֤��_In, �ֻ���_In);
    Else
      --�ϲ��˵�����ѱ𲻱�,�������������۲���
      Update ������Ϣ
      Set ����� = סԺ��_In, ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, �ѱ� = Decode(��������_In, 1, �ѱ�_In, �ѱ�), ҽ�Ƹ��ʽ = ���ʽ_In,
          �������� = ��������_In, ���� = ����_In, ���� = ����_In, ���� = ����_In, ���� = ����_In, ѧ�� = ѧ��_In, ����״�� = ����״��_In, ְҵ = ְҵ_In,
          ��� = ���_In, ���֤�� = ���֤��_In, �����ص� = �����ص�_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In, ��ͥ�绰 = ��ͥ�绰_In, ���ڵ�ַ = ���ڵ�ַ_In,
          ���ڵ�ַ�ʱ� = ���ڵ�ַ�ʱ�_In, ��ϵ������ = ��ϵ������_In, ��ϵ�˹�ϵ = ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ = ��ϵ�˵�ַ_In, ��ϵ�˵绰 = ��ϵ�˵绰_In, ������λ = ������λ_In,
          ��ͬ��λid = Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In), ��λ�绰 = ��λ�绰_In, ��λ�ʱ� = ��λ�ʱ�_In, ��λ������ = ��λ������_In,
          ��λ�ʺ� = ��λ�ʺ�_In, ������ = ������_In, ������ = Decode(������_In, 0, Null, ������_In), �������� = ��������_In, ���� = ����_In,
          ����֤�� = ����֤��_In, �������� = ��������_In, ��ϵ�����֤�� = ��ϵ�����֤��_In, �ֻ��� = Nvl(�ֻ���_In, �ֻ���)
      Where ����id = ����id_In;
    End If;
  Else
    If �²���_In = 1 Then
      Insert Into ������Ϣ
        (����id, סԺ��, ����, �Ա�, ����, �ѱ�, ҽ�Ƹ��ʽ, ��������, ����, ����, ����, ����, ѧ��, ����״��, ְҵ, ���, ���֤��, �����ص�, ��ͥ��ַ, ��ͥ��ַ�ʱ�, ��ͥ�绰,
         ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ������λ, ��ͬ��λid, ��λ�绰, ��λ�ʱ�, ��λ������, ��λ�ʺ�, ������, ������, ��������, ����, �Ǽ�ʱ��, ����֤��,
         ��������, ��ϵ�����֤��, �ֻ���)
      Values
        (����id_In, Decode(��������_In, 2, Null, סԺ��_In), ����_In, �Ա�_In, ����_In, �ѱ�_In, ���ʽ_In, ��������_In, ����_In, ����_In, ����_In,
         ����_In, ѧ��_In, ����״��_In, ְҵ_In, ���_In, ���֤��_In, �����ص�_In, ��ͥ��ַ_In, ��ͥ��ַ�ʱ�_In, ��ͥ�绰_In, ���ڵ�ַ_In, ���ڵ�ַ�ʱ�_In,
         ��ϵ������_In, ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ_In, ��ϵ�˵绰_In, ������λ_In, Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In), ��λ�绰_In, ��λ�ʱ�_In,
         ��λ������_In, ��λ�ʺ�_In, ������_In, Decode(������_In, 0, Null, ������_In), ��������_In, ����_In, v_Date, ����֤��_In, ��������_In,
         ��ϵ�����֤��_In, �ֻ���_In);
    Else
      --�ϲ��˵�����ѱ𲻱�,�������������۲���
      Update ������Ϣ
      Set סԺ�� = Decode(��������_In, 2, סԺ��, Decode(סԺ��_In, Null, סԺ��, סԺ��_In)), ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In,
          �ѱ� = Decode(��������_In, 1, �ѱ�_In, �ѱ�), ҽ�Ƹ��ʽ = ���ʽ_In, �������� = ��������_In, ���� = ����_In, ���� = ����_In, ���� = ����_In,
          ���� = ����_In, ѧ�� = ѧ��_In, ����״�� = ����״��_In, ְҵ = ְҵ_In, ��� = ���_In, ���֤�� = ���֤��_In, �����ص� = �����ص�_In, ��ͥ��ַ = ��ͥ��ַ_In,
          ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In, ��ͥ�绰 = ��ͥ�绰_In, ���ڵ�ַ = ���ڵ�ַ_In, ���ڵ�ַ�ʱ� = ���ڵ�ַ�ʱ�_In, ��ϵ������ = ��ϵ������_In, ��ϵ�˹�ϵ = ��ϵ�˹�ϵ_In,
          ��ϵ�˵�ַ = ��ϵ�˵�ַ_In, ��ϵ�˵绰 = ��ϵ�˵绰_In, ������λ = ������λ_In, ��ͬ��λid = Decode(��ͬ��λid_In, 0, Null, ��ͬ��λid_In),
          ��λ�绰 = ��λ�绰_In, ��λ�ʱ� = ��λ�ʱ�_In, ��λ������ = ��λ������_In, ��λ�ʺ� = ��λ�ʺ�_In, ������ = ������_In,
          ������ = Decode(������_In, 0, Null, ������_In), �������� = ��������_In, ���� = ����_In, ����֤�� = ����֤��_In, �������� = ��������_In,
          ��ϵ�����֤�� = ��ϵ�����֤��_In, �ֻ��� = Nvl(�ֻ���_In, �ֻ���)
      Where ����id = ����id_In;
    End If;
  End If;

  --������Ϣ
  Begin
    If �Ǽ�ģʽ_In = 1 Then
      v_��ҳid := 0; --ԤԼ�ǼǼ�¼����ҳID=0
    Else
      If ��ҳid_In Is Null Then
        Select Nvl(Max(��ҳid), 0) + 1 Into v_��ҳid From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0;
      Else
        v_��ҳid := ��ҳid_In;
      End If;
    End If;
  Exception
    When Others Then
      Null;
  End;
  v_���� := ����_In;
  If �Ǽ�ģʽ_In = 2 And v_���� Is Null Then
    Select Count(1) Into v_Count From ������ҳ Where ����id = ����id_In And ��ҳid = 0;
    If v_Count = 0 Then
      v_Error := '����ԤԼ��¼������,���ܼ�������!';
      Raise Err_Custom;
    End If;
    Select ��Ժ���� Into v_���� From ������ҳ Where ����id = ����id_In And ��ҳid = 0;
  End If;
  If �Ǽ�ģʽ_In <> 1 Then
    Update ������Ϣ
    Set ��ҳid = v_��ҳid, ��ǰ����id = ��Ժ����id_In, ��ǰ����id = ��Ժ����id_In, ��ǰ���� = Decode(v_����, '��ͥ����', Null, v_����), ��Ժʱ�� = ��Ժʱ��_In,
        ��Ժʱ�� = Null, ��Ժ = 1
    Where ����id = ����id_In;
  End If;

  --����סԺ����
  If �Ǽ�ģʽ_In <> 1 And ��������_In = 0 Then
    If Nvl(סԺ����_In, 0) = 0 Then
      Select Nvl(סԺ����, 0) + 1 Into n_סԺ���� From ������Ϣ Where ����id = ����id_In;
    Else
      n_סԺ���� := סԺ����_In;
    End If;
    Update ������Ϣ Set סԺ���� = n_סԺ���� Where ����id = ����id_In;
  End If;

  --ȡ���ʱ��
  If v_���� Is Null Then
    d_Indeptime := Null;
  Else
    d_Indeptime := ��Ժʱ��_In;
  End If;

  --״̬��0-������Ժ,1-�ȴ����,2-�ȴ�ת��
  If �Ǽ�ģʽ_In = 2 Then
    --��������ҳ�ӱ�
    Delete From ������ҳ�ӱ� Where ����id = ����id_In And Nvl(��ҳid, 0) = 0;
    --����ԤԼ
    Update ������ҳ
    Set ��ҳid = v_��ҳid, �������� = ��������_In, סԺ�� = Decode(��������_In, 1, Null, 2, Null, סԺ��_In),
        ���ۺ� = Decode(��������_In, 2, סԺ��_In, Null),
        --��ҳID���,�������ʿ��ܱ��
        �ѱ� = �ѱ�_In, ��Ժ����id = ��Ժ����id_In, ��Ժ����id = ��Ժ����id_In, ��Ժ���� = ��Ժʱ��_In, ���ʱ�� = d_Indeptime, ��Ժ���� = ��Ժ����_In,
        ��Ժ��ʽ = ��Ժ��ʽ_In, ��Ժ���� = ��Ժ����_In, ����Ժת�� = ����Ժת��_In, סԺĿ�� = סԺĿ��_In, ��Ժ���� = Decode(v_����, '��ͥ����', Null, v_����),
        �Ƿ���� = �Ƿ����_In, ��ǰ���� = ��Ժ����_In, ��ǰ����id = ��Ժ����id_In, ����ȼ�id = Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In),
        ��Ժ����id = ��Ժ����id_In, ��Ժ���� = Decode(v_����, '��ͥ����', Null, v_����), ����ҽʦ = ����ҽʦ_In, ��ĿԱ��� = ����Ա���_In, ��ĿԱ���� = ����Ա����_In,
        ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, ����״�� = ����״��_In, ְҵ = ְҵ_In, ���� = ����_In, ѧ�� = ѧ��_In, ��λ�绰 = ��λ�绰_In,
        ��λ�ʱ� = ��λ�ʱ�_In, ��λ��ַ = ������λ_In, ���� = ����_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In, ���ڵ�ַ = ���ڵ�ַ_In,
        ���ڵ�ַ�ʱ� = ���ڵ�ַ�ʱ�_In, ��ϵ������ = ��ϵ������_In, ��ϵ�˹�ϵ = ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ = ��ϵ�˵�ַ_In, ��ϵ�����֤�� = ��ϵ�����֤��_In, ��ϵ�˵绰 = ��ϵ�˵绰_In,
        ҽ�Ƹ��ʽ = ���ʽ_In, ��ע = ��ע_In, ���� = ����_In, ״̬ = Decode(v_����, Null, 1, 0), �Ǽ��� = ����Ա����_In, �Ǽ�ʱ�� = v_Date,
        ����Ժ = ����Ժ_In, �������� = ��������_In
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0;
    Update ����Ԥ����¼
    Set ��ҳid = ��ҳid_In
    Where ����id = ����id_In And ��ҳid Is Null And ����id = ��Ժ����id_In And Ԥ����� = 2 And ��Ԥ�� Is Null And
          Trunc(�տ�ʱ��) = Trunc(Sysdate);
  Else
    --��Ժ�Ǽǻ�ԤԼ�Ǽ�
    Insert Into ������ҳ
      (��������, ����id, ��ҳid, סԺ��, ���ۺ�, �ѱ�, ��Ժ����id, ��Ժ����id, ��Ժ����, ���ʱ��, ��Ժ����, ��Ժ��ʽ, ��Ժ����, ����Ժת��, סԺĿ��, ��Ժ����, �Ƿ����, ��ǰ����,
       ��ǰ����id, ����ȼ�id, ��Ժ����id, ��Ժ����, ����ҽʦ, ��ĿԱ���, ��ĿԱ����, ״̬, ����, �Ա�, ����, ����״��, ְҵ, ����, ѧ��, ��λ�绰, ��λ�ʱ�, ��λ��ַ, ����, ��ͥ��ַ,
       ��ͥ�绰, ��ͥ��ַ�ʱ�, ���ڵ�ַ, ���ڵ�ַ�ʱ�, ��ϵ������, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, ��ϵ�˵绰, ��ϵ�����֤��, ҽ�Ƹ��ʽ, ����, ��ע, �Ǽ���, �Ǽ�ʱ��, ����Ժ, ��������, �Һ�id)
    Values
      (��������_In, ����id_In, v_��ҳid, Decode(��������_In, 1, Null, 2, Null, סԺ��_In), Decode(��������_In, 2, סԺ��_In, Null), �ѱ�_In,
       ��Ժ����id_In, ��Ժ����id_In, ��Ժʱ��_In, d_Indeptime, ��Ժ����_In, ��Ժ��ʽ_In, ��Ժ����_In, ����Ժת��_In, סԺĿ��_In,
       Decode(v_����, '��ͥ����', Null, v_����), �Ƿ����_In, ��Ժ����_In, ��Ժ����id_In, Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In), ��Ժ����id_In,
       Decode(v_����, '��ͥ����', Null, v_����), ����ҽʦ_In, ����Ա���_In, ����Ա����_In, Decode(v_����, Null, 1, 0), ����_In, �Ա�_In, ����_In,
       ����״��_In, ְҵ_In, ����_In, ѧ��_In, ��λ�绰_In, ��λ�ʱ�_In, ������λ_In, ����_In, ��ͥ��ַ_In, ��ͥ�绰_In, ��ͥ��ַ�ʱ�_In, ���ڵ�ַ_In, ���ڵ�ַ�ʱ�_In,
       ��ϵ������_In, ��ϵ�˹�ϵ_In, ��ϵ�˵�ַ_In, ��ϵ�˵绰_In, ��ϵ�����֤��_In, ���ʽ_In, ����_In, ��ע_In, ����Ա����_In, v_Date, ����Ժ_In, ��������_In,
       �Һ�id_In);
  End If;

  Begin
    If �Ǽ�ģʽ_In <> 1 Then
      Update ��Ժ���� Set ����id = Nvl(��Ժ����id_In, 0), ����id = ��Ժ����id_In Where ����id = ����id_In;
      If Sql%RowCount = 0 Then
        Insert Into ��Ժ����
          (����id, ����id, ����id, ��ҳid)
        Values
          (����id_In, ��Ժ����id_In, Nvl(��Ժ����id_In, 0), Nvl(v_��ҳid, 0));
      End If;
    End If;
  Exception
    When Others Then
      Null;
  End;

  Select �ѱ� Into v_�ѱ� From ������Ϣ Where ����id = ����id_In;
  If v_�ѱ� Is Null Then
    Update ������Ϣ
    Set �ѱ� =
         (Select �ѱ� From ������ҳ Where ����id = ����id_In And ��ҳid = v_��ҳid)
    Where ����id = ����id_In;
  End If;

  --ҽ����
  If �Ǽ�ģʽ_In <> 1 Then
    Select Zl_סԺ�ձ�_Count(��Ժ����id_In, Trunc(��Ժʱ��_In)) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
      Raise Err_Custom;
    End If;
  
    If ҽ����_In Is Not Null Then
      Insert Into ������ҳ�ӱ� (����id, ��ҳid, ��Ϣ��, ��Ϣֵ) Values (����id_In, v_��ҳid, 'ҽ����', ҽ����_In);
    End If;
  
    --���˱䶯��¼
    --ͬʱ����ҷǼ�ͥ����ʱ�еȼ�
    If v_���� Is Not Null And v_���� <> '��ͥ����' Then
      Select �ȼ�id Into v_�ȼ�id From ��λ״����¼ Where ����id = ��Ժ����id_In And ���� = v_����;
    End If;
  
    --���ͬʱ���,����Ժ�������д��һ����Ժ�䶯
    Insert Into ���˱䶯��¼
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ���Ӵ�λ, ����id, ����id, ����ȼ�id, ��λ�ȼ�id, ����, ����, ����Ա���, ����Ա����)
    Values
      (���˱䶯��¼_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 0, ��Ժ����id_In, ��Ժ����id_In, Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In),
       v_�ȼ�id, Decode(v_����, '��ͥ����', Null, v_����), ��Ժ����_In, ����Ա���_In, ����Ա����_In);
  
    Insert Into �����Զ�����
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ����, ����id, ����id, ����ȼ�id, ����Ա���, ����Ա����)
    Values
      (�����Զ�����_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 1, ��Ժ����id_In, ��Ժ����id_In, Decode(����ȼ�id_In, 0, Null, ����ȼ�id_In),
       ����Ա���_In, ����Ա����_In);
    Insert Into �����Զ�����
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ����, ���Ӵ�λ, ����id, ����id, ��λ�ȼ�id, ����, ����Ա���, ����Ա����)
    Values
      (�����Զ�����_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 2, 0, ��Ժ����id_In, ��Ժ����id_In, v_�ȼ�id,
       Decode(v_����, '��ͥ����', Null, v_����), ����Ա���_In, ����Ա����_In);
    Insert Into �����Զ�����
      (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ����, ���Ӵ�λ, ����id, ����id, ��λ�ȼ�id, ����, ����Ա���, ����Ա����)
    Values
      (�����Զ�����_Id.Nextval, ����id_In, v_��ҳid, ��Ժʱ��_In, 1, 3, 0, ��Ժ����id_In, ��Ժ����id_In, v_�ȼ�id,
       Decode(v_����, '��ͥ����', Null, v_����), ����Ա���_In, ����Ա����_In);
  
    --ͬʱ����ҷǼ�ͥ����ʱ��λ��ռ��
    If v_���� Is Not Null And v_���� <> '��ͥ����' And �Ǽ�ģʽ_In <> 2 Then
      Select Count(*) Into v_Count From ��λ״����¼ Where ����id = ��Ժ����id_In And ���� = v_���� And ״̬ = '�մ�';
    
      If v_Count = 0 Then
        v_Error := '����ʧ��,��λ ' || v_���� || ' ���ǿմ���';
        Raise Err_Custom;
      End If;
    
      Update ��λ״����¼
      Set ״̬ = 'ռ��', ����id = ����id_In, ����id = Decode(����, 1, ��Ժ����id_In, ����id)
      Where ����id = ��Ժ����id_In And ���� = v_����;
    End If;
  
    --������ϼ�¼
    If �������_In Is Not Null Or ����id_In Is Not Null Then
      Insert Into ������ϼ�¼
        (ID, ����id, ��ҳid, ��¼��Դ, �������, ��ϴ���, ����id, ���id, �������, ��¼����, ��¼��)
      Values
        (������ϼ�¼_Id.Nextval, ����id_In, v_��ҳid, 2, 1, 1, ����id_In, ���id_In, �������_In, Sysdate, ����Ա����_In);
    End If;
    If ��ҽ���_In Is Not Null Or ��ҽ����id_In Is Not Null Then
      Insert Into ������ϼ�¼
        (ID, ����id, ��ҳid, ��¼��Դ, �������, ��ϴ���, ����id, ���id, �������, ��¼����, ��¼��)
      Values
        (������ϼ�¼_Id.Nextval, ����id_In, v_��ҳid, 2, 11, 1, ��ҽ����id_In, ��ҽ���id_In, ��ҽ���_In, Sysdate, ����Ա����_In);
    End If;
    --���˵�����¼
    Update ���˵�����¼
    Set ����ʱ�� = Sysdate
    Where ����id = ����id_In And ����ʱ�� Is Not Null And ����ʱ�� > Sysdate;
  
    --���˷���������Ŀ
    If �Ǽ�ģʽ_In <> 1 Then
      Delete From ����������Ŀ Where ����id = ����id_In;
      b_Message.Zlhis_Patient_001(����id_In, v_��ҳid);
    End If;
  
    If �Ǽ�ģʽ_In = 0 And ((�������_In Is Not Null Or ����id_In Is Not Null) Or (��ҽ���_In Is Not Null Or ��ҽ����id_In Is Not Null)) Then
      --����������дʱ��
      Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '���', ��Ժ����id_In, Null, Sysdate, Sysdate);
    End If;
  
    If �Ǽ�ģʽ_In = 0 And v_���� Is Not Null Then
      If ����Ժ_In = 0 Then
        Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '��Ժ', ��Ժ����id_In, Null, ��Ժʱ��_In, ��Ժʱ��_In);
      Else
        Zl_���Ӳ���ʱ��_Insert(����id_In, ��ҳid_In, 2, '�ٴ���Ժ', ��Ժ����id_In, Null, ��Ժʱ��_In, ��Ժʱ��_In);
      End If;
    End If;
  
    If v_���� Is Not Null Then
      --����׷����µ�
      Zl_�������µ�_Newfirst(����id_In, ��ҳid_In, ��Ժ����id_In);
    End If;
  
    --�����������
    Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��Ժ���� Is Null;
    If v_Count > 1 Then
      v_Error := '���ֲ��˴��ڷǷ��Ĳ�����¼,��ǰ�������ܼ�����' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�²���״̬�����ԣ�';
      Raise Err_Custom;
    End If;
  
    Select Count(*)
    Into v_Count
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = v_��ҳid And Nvl(���Ӵ�λ, 0) = 0 And ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null;
    If v_Count > 1 Then
      v_Error := '���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�²���״̬�����ԣ�';
      Raise Err_Custom;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ժ������ҳ_Insert;
/

--127912:��ΰ��,2018-07-03,ԤԼ���˴���
CREATE OR REPLACE Procedure Zl_��Ժ������ҳ_Delete
(
  ����id_In     ������ҳ.����id%Type,
  ��ҳid_In     ������ҳ.��ҳid%Type,
  ת����_In     Number := 0,
  ���סԺ��_In Number := 0
  --���ܣ�ȡ��������Ժ/ԤԼ�Ǽ�
  --     ��ҳID_IN:Ϊ0ʱ��ʾȡ��ԤԼ�Ǽ�
  --     ת����_IN:��������Ժ�Ǽǲ���תΪסԺ���۲���
  --     ���סԺ��_In:��һ��סԺ�Ĳ���ת����ʱ�Ƿ����סԺ��
) As
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_��Ժ����   ������ҳ.��Ժ����id%Type;
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_סԺ��     ������ҳ.סԺ��%Type;
  v_����Ժ     ������ҳ.����Ժ%Type;
  v_��Ժ����id ������ҳ.��Ժ����id%Type;
  v_��Ժ����   ������ҳ.��Ժ����id%Type;
  v_����       ������ҳ.��Ժ����%Type;

  n_�������� ������ҳ.��������%Type;
  n_��ҳid   ������ҳ.��ҳid%Type;

  v_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;

  Function Checkpatiadvice
  (
    ����id_In ������ҳ.����id%Type,
    ��ҳid_In ������ҳ.��ҳid%Type
  ) Return Varchar2 Is
    --����סԺ����ҽ����¼��������
    v_Err Varchar2(255);
  Begin
    v_Err := Null;
  
    For r_Row In (Select ����ҽ��, Decode(ҽ��״̬, -1, '�ݴ�', 1, '�¿�', 2, 'У������', 'δ����') As ״̬, ҽ������
                  From ����ҽ����¼
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(ҽ��״̬, 0) <> 4 And Rownum < 2) Loop
      v_Err := '��' || r_Row.����ҽ�� || '��ҽ����' || r_Row.״̬ || '��ҽ��û�д���,������ȡ���Ǽǣ�';
    End Loop;
    Return v_Err;
  End Checkpatiadvice;
Begin
  Select Nvl(״̬, 0), Nvl(��������, 0)
  Into v_Count, n_��������
  From ������ҳ
  Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_Count <> 1 Then
    v_Error := '�ò����Ѿ����,���Ƚ����˳�������Ժ״̬��';
    Raise Err_Custom;
  End If;

  --ɾ�����Ӳ���ʱ��
  Select ��Ժ����id, ����Ժ Into v_��Ժ����id, v_����Ժ From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_����Ժ = 0 Then
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', v_��Ժ����id);
  Else
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '�ٴ���Ժ', v_��Ժ����id);
  End If;

  --��ȡ���һ�β�Ϊ�յ�סԺ��
  Begin
    If ��ҳid_In = 0 Then
      --ԤԼ���ĵ�ԤԼ���˴��ڴ�λ��¼ 
      Select ��Ժ����id, ��Ժ���� Into v_��Ժ����, v_���� From ������ҳ Where ����id = ����id_In And ��ҳid = 0;
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0 And Nvl(סԺ��, 0) <> 0);
    Else
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And ��ҳid < ��ҳid_In And Nvl(סԺ��, 0) <> 0);
    End If;
  Exception
    When Others Then
      Null;
  End;

  b_Message.Zlhis_Patient_006(����id_In, ��ҳid_In, '��Ժ�Ǽ�');

  If ת����_In = 1 And Nvl(��ҳid_In, 0) <> 0 Then
    Update ������ҳ
    Set �������� = 2, סԺ�� = Decode(���סԺ��_In, 1, Null, סԺ��)
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(��������, 0) = 0;
  
    --����סԺ����
    Update ������Ϣ Set סԺ���� = Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null) Where ����id = ����id_In;
    If ���סԺ��_In = 1 Then
      Update ������Ϣ Set סԺ�� = v_סԺ�� Where ����id = ����id_In;
    End If;
  Else
    Begin
      Select b.��Ժ����, b.��Ժ����, b.��Ժ����id
      Into v_��Ժʱ��, v_��Ժʱ��, v_��Ժ����
      From ������Ϣ A, ������ҳ B
      Where a.����id = ����id_In And a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0;
    Exception
      When Others Then
        Null;
    End;
    --����ԤԼ�Ǽǲ��˲����סԺ�ձ�
    If Nvl(��ҳid_In, 0) <> 0 Then
      Select Zl_סԺ�ձ�_Count(v_��Ժ����, v_��Ժʱ��) Into v_Count From Dual;
      If v_Count > 0 Then
        v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
        Raise Err_Custom;
      End If;
    End If;
    --ԤԼ���ĵ�ԤԼ������Ҫ�ͷŴ�λ
    If v_���� Is Not Null Then
      Update ��λ״����¼ Set ״̬ = '�մ�', ����id = Null Where ����id = v_��Ժ���� And ���� = v_����;
    End If;
    --�������۲����´���Ժ֪ͨ�����������Ч�Ĳ�����ҳ��¼��36549��
    Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��Ժ���� Is Not Null And ��Ժ���� Is Null;
    If Not v_Count > 1 Then
      v_Count := 0;
      If Nvl(��ҳid_In, 0) <> 0 And Nvl(n_��������, 0) = 0 Then
        v_Count := 1;
      End If;
      --����Ժ����,ȡ����Ժ�Ǽ�ʱ,������Ϣ����Ժʱ��ͳ�Ժʱ��Ӧ�û��˵���һ����Ժ���ںͳ�Ժ����
      If v_����Ժ = 1 Then
        Begin
          Select ��Ժ����, ��Ժ����
          Into v_��Ժʱ��, v_��Ժʱ��
          From ������ҳ
          Where ����id = ����id_In And
                ��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And ��ҳid < ��ҳid_In);
        Exception
          When Others Then
            --�쳣������Ϊ������ȡ�������ݵ��쳣���
            Null;
        End;
      End If;
    
      Update ������Ϣ
      Set סԺ�� = v_סԺ��, סԺ���� = Decode(v_Count, 0, סԺ����, Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null)), ��ǰ����id = Null,
          ��ǰ����id = Null, ��ǰ���� = Null, ��Ժʱ�� = v_��Ժʱ��, ��Ժʱ�� = v_��Ժʱ��, ������ = Null, ������ = Null, �������� = Null, ��Ժ = Null
      Where ����id = ����id_In;
      Delete From ��Ժ���� Where ����id = ����id_In;
    End If;
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Delete From �����Զ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Delete From ������ϼ�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 2;
  
    --����סԺ�������Ԥ����,��Ϊ�������ｻ��
    Update ����Ԥ����¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  
    --���η�����,�ı����﷢��
    Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 5;
  
    --����סԺ�����з��ü�¼�޽�������ȫ���������򽫶�Ӧ���ü�¼�е�"��ҳID"�����
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From סԺ���ü�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1 And ����id Is Not Null;
  
    If v_Count = 0 Then
      Begin
        Select Nvl(Count(*), 0)
        Into v_Count
        From סԺ���ü�¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1
        Group By NO, ��¼����, ���
        Having Nvl(Sum(ʵ�ս��), 0) <> 0;
      Exception
        When Others Then
          v_Count := 0;
      End;
    
      If v_Count = 0 Then
        Delete ����δ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��� = 0;
        Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1;
      End If;
    End If;
  
    --����סԺ����ҽ����¼��������
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From ����ҽ����¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(ҽ��״̬, 0) <> 4;
    If v_Count = 0 Then
      Delete From ����ҽ����¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Else
      v_Error := Checkpatiadvice(����id_In, ��ҳid_In);
      If v_Error Is Not Null Then
        Raise Err_Custom;
      End If;
    End If;
  
    --���±�,û�н�������ҳ(����ID,��ҳID)�����,��Ϊ����ҳID�����ǹҺ�ID
    Delete From ���˹�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������ϼ�¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������������¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����ӡ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�����Ժ�����˾��￨,��ɾ����ʧ��(���˷��ü�¼��ҳID�����Լ��)
    Delete From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�޸Ĳ�����Ϣ����ҳID��סԺ����
    Select Max(��ҳid) Into n_��ҳid From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0;
    Update ������Ϣ Set ��ҳid = n_��ҳid Where ����id = ����id_In;
    If n_��ҳid Is Null Then
      Update ������Ϣ Set סԺ���� = Null Where ����id = ����id_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ժ������ҳ_Delete;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0019' Where ���=&n_System;
Commit;
