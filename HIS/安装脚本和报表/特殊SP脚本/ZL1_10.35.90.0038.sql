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



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--134222:����,2018-11-23,����Oracle����Zl_Third_Getregistalter,�ֶδ������ص�XML�ַ���
Create Or Replace Procedure Zl_Third_Getregistalter
(
  Xml_In  Xmltype,
  Xml_Out Out Xmltype
) Is
  -----------------------------------------------
  --���ܣ���ȡ���������ͣ���ﰲ��
  --��Σ�XML_IN
  --<IN>
  --  <JSKLB>���㿨���</JSKLB>
  --  <RQ>����</RQ>
  --</IN>
  --����:XML_OUT
  --<OUTPUT>
  --  <TZLISTS>          //ͣ���б�
  --    <ITEM>
  --      <HM>����</HM>
  --      <YSID>ҽ��ID</YSID>
  --      <YS>ҽ������</YS>
  --      <KSSJ>ͣ�￪ʼʱ��</KSSJ>
  --      <JSSJ>ͣ�����ʱ��</JSSJ>
  --      <BRLIST>
  --        <INFO>
  --          <YYNO>ԤԼ���ݺ�</YYNO>
  --          <BRID>����ID</BRID>
  --          <YYSJ>ԤԼʱ��</YYSJ>
  --          <CZSJ>����ʱ��</CZSJ>
  --          <YYKS>ԤԼ����</YYKS>
  --          <GHLX>����</GHLX>
  --          <YSXM>ҽ������</YSXM>
  --        </INFO>
  --      </BRLIST>
  --    </ITEM>
  --  </TZLISTS>
  --  <HZLISTS>          //�����б�
  --    <ITEM>
  --      <BRID>����ID</BRID>
  --      <YYSJ>ԤԼ�Ĳ���ʱ��</YYSJ>
  --      <YSJ>ԭԤԼʱ��</YSJ>
  --      <YHM>ԭ����</YHM>
  --      <YYS>ԭҽ��</YYS>
  --      <YZC>ԭҽ����ְ��</YZC>
  --      <XSJ>��ԤԼʱ��</XSJ>
  --      <XHM>�ֺ���</XHM>
  --      <XYS>��ҽ��</XYS>
  --      <XZC>��ҽ����ְ��</XZC>
  --    </ITEM>
  --  </HZLIST>
  --</OUTPUT>
  -----------------------------------------------------

  d_Date     Date;
  v_Jsklb    Varchar2(100);
  n_�����id ҽ�ƿ����.Id%Type;
  n_Cnt      Number(3);
  v_Temp     Clob;
  v_Brinfo   Varchar2(4000);
  d_����ʱ�� Date;
  v_Para     Varchar2(2000);
  n_Exists   Number(3);
  n_�Һ�ģʽ Number(3);
  x_Templet  Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/JSKLB') Into v_Jsklb From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd')
  Into d_Date
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select b.Id Into n_�����id From ҽ�ƿ���� B Where b.���� = v_Jsklb And Rownum < 2;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  If n_�Һ�ģʽ = 1 And Nvl(d_Date, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) Then
    --������Ű�ģʽ
    --��ȡͣ�ﰲ��
    For r_ͣ�� In (Select a.Id As ��¼id, b.����, a.ҽ��id, a.ҽ������, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                 From �ٴ������¼ A, �ٴ������Դ B, �ٴ�����ͣ���¼ C
                 Where a.Id = c.��¼id And a.��Դid = b.Id And a.ͣ�￪ʼʱ�� Is Not Null And c.����ʱ�� Between d_Date And
                       d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || r_ͣ��.���� || '</HM><YSID>' || r_ͣ��.ҽ��id || '</YSID><YS>' || r_ͣ��.ҽ������ ||
                '</YS><KSSJ>' || r_ͣ��.ͣ�￪ʼʱ�� || '</KSSJ><JSSJ>' || r_ͣ��.ͣ����ֹʱ�� || '</JSSJ><BRLIST>';
      For r_ͣ�ﲡ�� In (Select a.��¼����, a.No, a.����id, To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��,
                            To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, b.����, d.����, c.ҽ������ As ҽ������
                     From ���˹Һż�¼ A, ���ű� B, �ٴ������¼ C, �ٴ������Դ D
                     Where a.ִ�в���id = b.Id And a.�����¼id = c.Id And c.��Դid = d.Id And ��¼״̬ = 1 And
                           ����ʱ�� Between r_ͣ��.ͣ�￪ʼʱ�� And r_ͣ��.ͣ����ֹʱ�� And a.�����¼id = r_ͣ��.��¼id And Not Exists
                      (Select 1 From ����䶯��¼ Where �Һŵ� = a.No)) Loop
        --ͣ�ﲡ���б��������Ѿ������ȡ���˵Ĳ���
        If r_ͣ�ﲡ��.��¼���� = 2 Then
          v_Brinfo := '<INFO><YYNO>' || r_ͣ�ﲡ��.No || '</YYNO><BRID>' || r_ͣ�ﲡ��.����id || '</BRID><YYSJ>' || r_ͣ�ﲡ��.����ʱ�� ||
                      '</YYSJ><CZSJ>' || r_ͣ�ﲡ��.�Ǽ�ʱ�� || '</CZSJ>' || '<YYKS>' || r_ͣ�ﲡ��.���� || '</YYKS><GHLX>' ||
                      r_ͣ�ﲡ��.���� || '</GHLX><YSXM>' || r_ͣ�ﲡ��.ҽ������ || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        Else
          Begin
            Select 1
            Into n_Exists
            From ����Ԥ����¼
            Where NO = r_ͣ�ﲡ��.No And ��¼���� = 4 And �����id = n_�����id;
          Exception
            When Others Then
              n_Exists := 0;
          End;
          If n_Exists = 1 Then
            v_Brinfo := '<INFO><YYNO>' || r_ͣ�ﲡ��.No || '</YYNO><BRID>' || r_ͣ�ﲡ��.����id || '</BRID><YYSJ>' || r_ͣ�ﲡ��.����ʱ�� ||
                        '</YYSJ><CZSJ>' || r_ͣ�ﲡ��.�Ǽ�ʱ�� || '</CZSJ>' || '<YYKS>' || r_ͣ�ﲡ��.���� || '</YYKS><GHLX>' ||
                        r_ͣ�ﲡ��.���� || '</GHLX><YSXM>' || r_ͣ�ﲡ��.ҽ������ || '</YSXM></INFO>';
            v_Temp   := v_Temp || v_Brinfo;
          End If;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    --��ȡ�����б�
    v_Temp := '';
    For r_���� In (Select d.��¼����, d.No, a.����id, To_Char(d.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ԤԼʱ��,
                        To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.ԭ����, a.ԭҽ������, b.רҵ����ְ�� As ԭְ��, a.�ֺ���, a.��ҽ������,
                        c.רҵ����ְ�� As ��ְ��
                 From ����䶯��¼ A, ��Ա�� B, ��Ա�� C, ���˹Һż�¼ D
                 Where a.�Ǽ�ʱ�� Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.ԭҽ��id = b.Id And a.��ҽ��id = c.Id And
                       a.�Һŵ� = d.No) Loop
      --ֻ���ظÿ����ҺŵĲ���         
      If r_����.��¼���� = 2 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || r_����.����id || '</BRID><YYSJ>' || r_����.�Ǽ�ʱ�� || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || r_����.ԤԼʱ�� || '</YSJ><YHM>' || r_����.ԭ���� || '</YHM><YYS>' || r_����.ԭҽ������ ||
                  '</YYS><YZC>' || r_����.ԭְ�� || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || r_����.ԤԼʱ�� || '</XSJ><XHM>' || r_����.�ֺ��� || '</XHM><XYS>' || r_����.��ҽ������ ||
                  '</XYS><XZC>' || r_����.��ְ�� || '</XZC></ITEM>';
      Else
        Begin
          Select 1 Into n_Exists From ����Ԥ����¼ Where NO = r_����.No And ��¼���� = 4 And �����id = n_�����id;
        Exception
          When Others Then
            n_Exists := 0;
        End;
        If n_Exists = 1 Then
          v_Temp := v_Temp || '<ITEM><BRID>' || r_����.����id || '</BRID><YYSJ>' || r_����.�Ǽ�ʱ�� || '</YYSJ>';
          v_Temp := v_Temp || '<YSJ>' || r_����.ԤԼʱ�� || '</YSJ><YHM>' || r_����.ԭ���� || '</YHM><YYS>' || r_����.ԭҽ������ ||
                    '</YYS><YZC>' || r_����.ԭְ�� || '</YZC>';
          v_Temp := v_Temp || '<XSJ>' || r_����.ԤԼʱ�� || '</XSJ><XHM>' || r_����.�ֺ��� || '</XHM><XYS>' || r_����.��ҽ������ ||
                    '</XYS><XZC>' || r_����.��ְ�� || '</XZC></ITEM>';
        End If;
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Else
    --�ƻ��Ű�ģʽ
    --��ȡͣ�ﰲ��
    For Rs In (Select b.����, b.ҽ��id, b.ҽ������, To_Char(a.��ʼֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʼֹͣʱ��,
                      To_Char(a.����ֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ֹͣʱ��
               From �ҺŰ���ͣ��״̬ A, �ҺŰ��� B
               Where a.����id = b.Id And a.�ƶ����� Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60) Loop
      v_Temp := v_Temp || '<ITEM><HM>' || Rs.���� || '</HM><YSID>' || Rs.ҽ��id || '</YSID><YS>' || Rs.ҽ������ ||
                '</YS><KSSJ>' || Rs.��ʼֹͣʱ�� || '</KSSJ><JSSJ>' || Rs.����ֹͣʱ�� || '</JSSJ><BRLIST>';
      ----2015/7/28
      For Rs_Br In (Select a.No, a.����id, To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��,
                           To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, b.����, c.����, a.ִ���� As ҽ������
                    From ���˹Һż�¼ A, ���ű� B, �ҺŰ��� C
                    Where a.�ű� = Rs.���� And a.ִ��״̬ = 0 And a.ִ�в���id = b.Id And b.Id = c.����id And a.�ű� = c.���� And
                          Trunc(����ʱ��) Between Trunc(To_Date(Rs.��ʼֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss')) And
                          Trunc(To_Date(Rs.����ֹͣʱ��, 'yyyy-mm-dd hh24:mi:ss'))) Loop
        --ֻ���ظÿ����ҺŵĲ���
        Select Count(*)
        Into n_Cnt
        From (Select 1
               From ����Ԥ����¼ A
               Where a.No = Rs_Br.No And a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = Rs_Br.����id And �����id = n_�����id
               Union All
               Select 1 From ���˹Һż�¼ Where NO = Rs_Br.No And ��¼״̬ = 1 And ����˵�� = v_Jsklb);
        If n_Cnt > 0 Then
          v_Brinfo := '<INFO><YYNO>' || Rs_Br.No || '</YYNO><BRID>' || Rs_Br.����id || '</BRID><YYSJ>' || Rs_Br.����ʱ�� ||
                      '</YYSJ><CZSJ>' || Rs_Br.�Ǽ�ʱ�� || '</CZSJ>' || '<YYKS>' || Rs_Br.���� || '</YYKS><GHLX>' || Rs_Br.���� ||
                      '</GHLX><YSXM>' || Rs_Br.ҽ������ || '</YSXM></INFO>';
          v_Temp   := v_Temp || v_Brinfo;
        End If;
        v_Brinfo := '';
      End Loop;
      v_Temp := v_Temp || '</BRLIST></ITEM>';
    End Loop;
    v_Temp := '<TZLISTS>' || v_Temp || '</TZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  
    --��ȡ�����¼
    v_Temp := '';
    For Rs In (Select d.No, a.����id, To_Char(d.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ԤԼʱ��,
                      To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.ԭ����, a.ԭҽ������, b.רҵ����ְ�� As ԭְ��, a.�ֺ���, a.��ҽ������,
                      c.רҵ����ְ�� As ��ְ��
               From ����䶯��¼ A, ��Ա�� B, ��Ա�� C, ���˹Һż�¼ D
               Where a.�Ǽ�ʱ�� Between d_Date And d_Date + 1 - 1 / 24 / 60 / 60 And a.ԭҽ��id = b.Id And a.��ҽ��id = c.Id And
                     a.�Һŵ� = d.No) Loop
      --ֻ���ظÿ����ҺŵĲ���         
      Select Count(*)
      Into n_Cnt
      From (Select 1
             From ����Ԥ����¼ A
             Where a.No = Rs.No And a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = Rs.����id And �����id = n_�����id
             Union All
             Select 1 From ���˹Һż�¼ Where NO = Rs.No And ��¼״̬ = 1 And ����˵�� = v_Jsklb);
      If n_Cnt > 0 Then
        v_Temp := v_Temp || '<ITEM><BRID>' || Rs.����id || '</BRID><YYSJ>' || Rs.�Ǽ�ʱ�� || '</YYSJ>';
        v_Temp := v_Temp || '<YSJ>' || Rs.ԤԼʱ�� || '</YSJ><YHM>' || Rs.ԭ���� || '</YHM><YYS>' || Rs.ԭҽ������ || '</YYS><YZC>' ||
                  Rs.ԭְ�� || '</YZC>';
        v_Temp := v_Temp || '<XSJ>' || Rs.ԤԼʱ�� || '</XSJ><XHM>' || Rs.�ֺ��� || '</XHM><XYS>' || Rs.��ҽ������ || '</XYS><XZC>' ||
                  Rs.��ְ�� || '</XZC></ITEM>';
      End If;
    End Loop;
    v_Temp := '<HZLISTS>' || v_Temp || '</HZLISTS>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregistalter;
/

--134551:������,2018-11-22,�����ӿڹ��̲���ê��
Create Or Replace Procedure Zl_Third_Buildpatient
(
  Patiinfo_In  In Xmltype,
  Patiinfo_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------
  --����˵��:
  -- ��� Patiinfo_In:
  --<IN>
  --  <ZJH></ZJH>                 //֤���ţ�Ŀǰ��֧�����֤��
  --  <ZJLX></ZJLX>                       //֤������(Ŀǰ��֧�����֤,Ϊ��ʱĬ��Ϊ���֤)
  --  <XM></XM>                       //����
  --  <SJH></SJH>                      //�ֻ���
  --</IN>

  --���� Patiinfo_Out��
  --<OUTPUT>
  --       <BRID></BRID>                //����ID
  --       <MZH></MZH>                  //�����
  --     <ERROR></ERROR>         //����д��󷵻ظýڵ�
  --</OUTPUT>
  -------------------------------------------------------------------------------
  n_Pati_Id      ������Ϣ.����id%Type;
  n_Card_Type_Id ҽ�ƿ����.Id%Type;
  n_Count        Number(5);
  n_Sum          Number(5);
  v_У��λ       Varchar2(50);

  v_����         ������Ϣ.����%Type;
  v_���֤��     ������Ϣ.���֤��%Type;
  v_�ֻ���       ������Ϣ.��ͥ�绰%Type;
  v_�Ա�         ������Ϣ.�Ա�%Type;
  v_����         ������Ϣ.����%Type;
  v_����Ա       ��Ա��.����%Type;
  v_ҽ�Ƹ��ʽ ������Ϣ.ҽ�Ƹ��ʽ%Type;
  n_�����       ������Ϣ.�����%Type;
  v_֤������     ҽ�ƿ����.����%Type;
  v_֤����       ����ҽ�ƿ���Ϣ.����%Type;

  v_Pattern Varchar2(500);
  v_Temp    Varchar2(32767); --��ʱXML
  v_Err_Msg Varchar2(2000);
  n_����    Number(2);

  d_��������  ������Ϣ.��������%Type;
  d_Curr_Time Date;

  Err_Item Exception;
Begin
  Patiinfo_Out := Xmltype('<OUTPUT></OUTPUT>');
  Select Sysdate Into d_Curr_Time From Dual;

  --�½����ˣ����������֤�š��ֻ��ţ����ڼ�ͥ�绰�У����������ڡ��Ա�����(��������ɴ����֤�л�ȡ)��
  Select Extractvalue(Value(I), 'IN/XM'), Extractvalue(Value(I), 'IN/ZJH'), Extractvalue(Value(I), 'IN/SJH'),
         Extractvalue(Value(I), 'IN/ZJLX')
  Into v_����, v_֤����, v_�ֻ���, v_֤������
  From Table(Xmlsequence(Extract(Patiinfo_In, 'IN'))) I;

  Begin
    If v_֤������ Is Null Then
      Select ����id
      Into n_Pati_Id
      From ����ҽ�ƿ���Ϣ
      Where ���� = v_֤���� And �����id In (Select ID From ҽ�ƿ���� Where ���� Like '%���֤%') And Rownum < 2;
    Else
      Select ����id
      Into n_Pati_Id
      From ����ҽ�ƿ���Ϣ
      Where ���� = v_֤���� And �����id In (Select ID From ҽ�ƿ���� Where ���� = v_֤������) And Rownum < 2;
    End If;
    n_���� := 1;
  Exception
    When Others Then
      n_���� := 0;
  End;

  If Nvl(n_����, 0) = 1 Then
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    Select ����� Into n_����� From ������Ϣ Where ����id = n_Pati_Id;
    If n_����� Is Null Then
      n_����� := Nextno(3);
      Update ������Ϣ Set ����� = n_����� Where ����id = n_Pati_Id;
    End If;
    v_Temp := '<MZH>' || n_����� || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
  Else
    If v_���� Is Null Then
      v_Err_Msg := '��������Ϊ��!';
      Raise Err_Item;
    End If;
    If v_֤������ Like '%���֤%' Or v_֤������ Is Null Then
      v_���֤�� := v_֤����;
    Else
      v_Err_Msg := 'Ŀǰ��֧�����֤����ķ�ʽ������';
      Raise Err_Item;
    End If;
  
    If v_���֤�� Is Null Then
      v_Err_Msg := '�������֤��Ϊ��!';
      Raise Err_Item;
    Else
      --���֤�Ϸ���֤
      v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,83,91';
    
      --��������
      If Instr(v_Pattern, Substr(v_���֤��, 1, 2)) = 0 Then
        v_Err_Msg := '���֤ǰ��λ�����벻��ȷ!';
        Raise Err_Item;
      End If;
      --���֤���ȼ��
      If Length(v_���֤��) = 15 Then
        --������֤��:15λ���֤��Ҫ��ȫ��Ϊ����
        v_Pattern := '^\d{15}$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_���֤��, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�а����Ƿ��ַ�������!';
          Raise Err_Item;
        End If;
        --��ȡ�Ա�
        If Mod(To_Number(Substr(v_���֤��, 15, 1)), 2) = 1 Then
          v_�Ա� := '��';
        Else
          v_�Ա� := 'Ů';
        End If;
        --�������ڵĺϷ��Լ��
      
        v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(v_���֤��, 7, 6), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Raise Err_Item;
        Else
          --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
          If Instr(',0229,0230,', ',' || Substr(v_���֤��, 9, 4) || ',') > 0 Then
            v_Temp     := '19' || Substr(v_���֤��, 7, 2) || '0301';
            d_�������� := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_�������� := To_Date('19' || Substr(v_���֤��, 7, 6), 'yyyy-mm-dd');
          End If;
          If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '���֤�еĳ���������Ч������!';
            Raise Err_Item;
          End If;
        End If;
      Elsif Length(v_���֤��) = 18 Then
        -- 18 λ���֤��ǰ17 λȫ��Ϊ���֣����1λ��Ϊ���ֻ�x
        v_Pattern := '^\d{17}[0-9Xx]$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(v_���֤��, v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�а����Ƿ��ַ�!';
          Raise Err_Item;
        End If;
        --��ȡ�Ա�
        If Mod(To_Number(Substr(v_���֤��, 17, 1)), 2) = 1 Then
          v_�Ա� := '��';
        Else
          v_�Ա� := 'Ů';
        End If;
        --�������ڵĺϷ��Լ��
        v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
        Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(v_���֤��, 7, 8), v_Pattern);
        If n_Count = 0 Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Raise Err_Item;
        Else
          --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
          If Instr(',0229,0230,', ',' || Substr(v_���֤��, 11, 4) || ',') > 0 Then
            v_Temp     := Substr(v_���֤��, 7, 4) || '0301';
            d_�������� := To_Date(v_Temp, 'yyyy-mm-dd') - 1;
          Else
            d_�������� := To_Date(Substr(v_���֤��, 7, 8), 'yyyy-mm-dd');
          End If;
          If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
            v_Err_Msg := '���֤�еĳ���������Ч������!';
            Raise Err_Item;
          End If;
          --����У��λ
          n_Sum     := (To_Number(Substr(v_���֤��, 1, 1)) + To_Number(Substr(v_���֤��, 11, 1))) * 7 +
                       (To_Number(Substr(v_���֤��, 2, 1)) + To_Number(Substr(v_���֤��, 12, 1))) * 9 +
                       (To_Number(Substr(v_���֤��, 3, 1)) + To_Number(Substr(v_���֤��, 13, 1))) * 10 +
                       (To_Number(Substr(v_���֤��, 4, 1)) + To_Number(Substr(v_���֤��, 14, 1))) * 5 +
                       (To_Number(Substr(v_���֤��, 5, 1)) + To_Number(Substr(v_���֤��, 15, 1))) * 8 +
                       (To_Number(Substr(v_���֤��, 6, 1)) + To_Number(Substr(v_���֤��, 16, 1))) * 4 +
                       (To_Number(Substr(v_���֤��, 7, 1)) + To_Number(Substr(v_���֤��, 17, 1))) * 2 +
                       To_Number(Substr(v_���֤��, 8, 1)) * 1 + To_Number(Substr(v_���֤��, 9, 1)) * 6 +
                       To_Number(Substr(v_���֤��, 10, 1)) * 3;
          n_Count   := Mod(n_Sum, 11);
          v_Pattern := '10X98765432';
          v_У��λ  := Substr(v_Pattern, n_Count + 1, 1);
          If v_У��λ <> Upper(Substr(v_���֤��, 18, 1)) Then
            v_Err_Msg := '���֤���벻��ȷ�����顣';
            Raise Err_Item;
          End If;
        End If;
      Else
        v_Err_Msg := '���֤���Ȳ���,���顣';
        Raise Err_Item;
      End If;
    
      If Nvl(v_����, '_') = '_' Then
        v_���� := Zl_Age_Calc(0, d_��������, d_Curr_Time);
      End If;
    End If;
  
    Select ���� Into v_ҽ�Ƹ��ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
    n_Pati_Id := Nextno(1);
    n_�����  := Nextno(3);
    Insert Into ������Ϣ
      (����id, ����, ���֤��, ��ͥ�绰, ��������, �Ա�, ����, �Ǽ�ʱ��, �����, ҽ�Ƹ��ʽ, �ֻ���)
      Select n_Pati_Id, v_����, v_���֤��, v_�ֻ���, d_��������, v_�Ա�, v_����, d_Curr_Time, n_�����, v_ҽ�Ƹ��ʽ, v_�ֻ���




      From Dual;
    --������Ϣ����������ҽ�ƿ��󶨣��������֤�����İ󶨣�
    Begin
      If v_֤������ Is Null Then
        Select ID Into n_Card_Type_Id From ҽ�ƿ���� Where ���� Like '%���֤%' And Rownum < 2;
      Else
        Select ID Into n_Card_Type_Id From ҽ�ƿ���� Where ���� = v_֤������ And Rownum < 2;
      End If;
    Exception
      When No_Data_Found Then
        v_Err_Msg := '���֤����𲻴��ڣ�';
        Raise Err_Item;
    End;
    Select b.���� Into v_����Ա From �ϻ���Ա�� A, ��Ա�� B Where a.��Աid = b.Id And a.�û��� = User;
  
    Zl_ҽ�ƿ��䶯_Insert(11, n_Pati_Id, n_Card_Type_Id, Null, v_���֤��, '�������⿨', Null, v_����Ա, d_Curr_Time);
  
    v_Temp := '<BRID>' || n_Pati_Id || '</BRID>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
    v_Temp := '<MZH>' || n_����� || '</MZH>';
    Select Appendchildxml(Patiinfo_Out, '/OUTPUT', Xmltype(v_Temp)) Into Patiinfo_Out From Dual;
	 b_Message.Zlhis_Patient_015(n_Pati_Id);
  End If;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Buildpatient;
/

--129503:����,2018-11-22,������Ŀ����ӿ�
Create Or Replace Procedure Zl_Third_Tendfile_Itemsave
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  n_Fileid       Number(18);
  n_��ʽid       Number(18);
  n_Xh           Number(5);
  n_Brid         Number(18);
  n_Zyid         Number(5);
  n_Babby        Number(1);
  v_Czy          Varchar2(20);
  n_Newadd       Number(1);
  n_Kind         Number(1);
  n_Num          Number(1);
  Intins         Number(1);
  n_�鵵         Number(1);
  v_����id       Number(18);
  v_Name         Varchar2(20);
  d_Ӥ����Ժʱ�� Date;
  v_Error        Varchar2(255);
  v_Temp         Varchar2(32767);
  x_Templet      Xmltype; --ģ��XML
  Err_Custom Exception;
Begin

  Select To_Number(Extractvalue(Value(A), 'IN/BRID'))
  Into n_Brid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/ZYID'))
  Into n_Zyid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/YEXH'))
  Into n_Babby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Char(Extractvalue(Value(A), 'IN/CZY')) 
  Into v_Czy 
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Begin
    Select Max(1)
    Into n_�鵵
    From ���˻����ļ�
    Where ����id = n_Brid And ��ҳid = n_Zyid And Ӥ�� = n_Babby And �鵵ʱ�� Is Null;
  End;

  For r_Input In (Select ����ʱ��, Lx, Ly, Mc, Nr, Bw, Wj
                  From Xmltable('$a/IN/ITEMLIST/ITEM' Passing Xmlfilelist_In As "a" Columns ����ʱ�� Varchar2(20) Path
                                 'TIME', Lx Number(1) Path 'LX', Ly Number(2) Path 'LY', Mc Varchar2(20) Path 'MC',
                                 Nr Varchar2(20) Path 'NR', Bw Varchar2(10) Path 'BW', Wj Varchar2(4000) Path 'Wj') B) Loop
    If r_Input.Mc Is Null Then
      v_Error := 'δ¼�����ݣ���������������飡';
      Raise Err_Custom;
    Else
      Select Max(��Ŀ���) Into n_Xh From �����¼��Ŀ Where ��Ŀ���� = r_Input.Mc;
    End If;
  
    If n_Babby <> 0 Then
      Begin
        Select ��ʼִ��ʱ��
        Into d_Ӥ����Ժʱ��
        From ����ҽ����¼ B, ������ĿĿ¼ C
        Where b.������Ŀid + 0 = c.Id And b.ҽ��״̬ = 8 And Nvl(b.Ӥ��, 0) <> 0 And c.��� = 'Z' And
              Instr(',3,5,11,', ',' || c.�������� || ',', 1) > 0 And b.����id = n_Brid And b.��ҳid = n_Zyid And b.Ӥ�� = n_Babby;
      Exception
        When Others Then
          d_Ӥ����Ժʱ�� := Null;
      End;
    End If;
  
    If d_Ӥ����Ժʱ�� Is Null Then
      v_����id := 0;
      Begin
        Select a.����id
        Into v_����id
        From ���˱䶯��¼ A
        Where a.����id Is Not Null And a.����id = n_Brid And a.��ҳid = n_Zyid And
              (To_Date(To_Char(To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 'YYYY-MM-DD HH24:MI') || '59',
                       'YYYY-MM-DD HH24:MI:SS') >= a.��ʼʱ�� And
              (To_Date(To_Char(To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 'YYYY-MM-DD HH24:MI') || '00',
                        'YYYY-MM-DD HH24:MI:SS') <= Nvl(a.��ֹʱ��, Sysdate) Or a.��ֹʱ�� Is Null)) And Rownum < 2;
      Exception
        When Others Then
          v_����id := 0;
      End;
      If v_����id = 0 Then
        v_Error := '���ݷ���ʱ�� ' || To_Date(r_Input.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ������������';
        Raise Err_Custom;
      End If;
    End If;
  
    Select Max(a.Id)
    Into n_Fileid
    From ���˻����ļ� A, �����ļ��ṹ B, �����ļ��б� C
    Where a.��ʽid = c.Id And ���� = 0 And ����id = n_Brid And ��ҳid = n_Zyid And Ӥ�� = n_Babby And a.��ʽid = b.�ļ�id And
          Ҫ������ = r_Input.Mc And b.�ļ�id = c.Id And a.��ʼʱ�� < To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') And
          (����ʱ�� > To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') Or ����ʱ�� Is Null)
    Order By a.��ʼʱ��;
  
    If n_Fileid Is Null Then
      Select Max(a.Id), Max(c.����), Max(c.Id)
      Into n_Fileid, n_Kind, n_��ʽid
      From ���˻����ļ� A, �����ļ��б� C
      Where a.��ʽid = c.Id And ���� = -1 And ����id = n_Brid And ��ҳid = n_Zyid And Ӥ�� = n_Babby And
            a.��ʼʱ�� < To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') And
            (����ʱ�� > To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') Or ����ʱ�� Is Null)
      Order By a.��ʼʱ��;
      If n_Fileid <> 0 Then
        v_Name := r_Input.Mc;
        If n_Kind = '1' Then
          If n_Xh = 4 Or n_Xh = 5 Then
            Select Max(�����ı�)
            Into n_Num
            From ���˻����ļ� A, �����ļ��ṹ B
            Where a.��ʽid = b.�ļ�id And a.Id = n_Fileid And Ҫ������ = 'Ӥ�����µ�';
            If Not (n_Num = 1) Then
              v_Name := 'Ѫѹ';
            End If;
          End If;
          Begin
            Select 1
            Into Intins
            From (Select To_Char(f.��¼��) As ��Ŀ����, g.��Ŀ����
                   From ���¼�¼��Ŀ F, �����¼��Ŀ G
                   Where f.��Ŀ��� = g.��Ŀ��� And g.��Ŀ���� = 2 And
                         (g.���ÿ��� = 1 Or
                         (g.���ÿ��� = 2 And Exists
                          (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id))) And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And
                         (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(n_Babby, 0, 1, 2))
                   Union All
                   Select b.Ҫ������ As ��Ŀ����, 1 As ��Ŀ����
                   From �����ļ��ṹ A, �����ļ��ṹ B
                   Where a.�ļ�id = n_��ʽid And a.��id Is Null And a.������� In (2, 3) And b.��id = a.Id) H
            Where Instr(',' || h.��Ŀ���� || ',', ',' || v_Name || ',', 1) > 0;
          
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1
            Into Intins
            From ���¼�¼��Ŀ F, �����¼��Ŀ G
            Where f.��Ŀ��� = g.��Ŀ��� And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And g.����ȼ� >= 0 And
                  (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(n_Babby, 0, 1, 2)) And f.��Ŀ��� = n_Xh And
                  (g.���ÿ��� = 1 Or (g.���ÿ��� = 2 And Exists
                   (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id)));
          Exception
            When Others Then
              Intins := 0;
          End;
        End If;
      
        If Intins = 1 Then
        
          Zl_���µ�����_Update(n_Fileid, To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 1, n_Xh, r_Input.Nr, r_Input.Bw, 0,
                          Null, 1, r_Input.Ly, Null, 0, 0, Null, Null, v_Czy);
        End If;
      End If;
    Else
      Select Max(1)
      Into n_Newadd
      From ���˻�������
      Where ����ʱ�� = To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') And �ļ�id = n_Fileid;
      If n_Newadd = 1 Then
        Zl_���˻�������_Update(n_Fileid, To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 1, n_Xh, r_Input.Nr, r_Input.Bw, 1,
                         r_Input.Ly, 0, v_Czy, Null, Null, Null);
      Else
        Select Max(1)
        Into n_Num
        From ���˻�������
        Where ����ʱ�� = To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') And ǩ���� Is Not Null And �ļ�id = n_Fileid;
        If n_Num = 1 Then
          v_Error := '��ǰ���˵Ļ����ļ���ǩ�����������޸ģ����Ȼ���ǩ����';
          Raise Err_Custom;
        Else
          Zl_���˻�������_Update(n_Fileid, To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 1, n_Xh, r_Input.Nr, r_Input.Bw, 1,
                           r_Input.Ly, 0, v_Czy, Null, Null, Null);
          Zl_���˻����ӡ_Update(n_Fileid, To_Date(r_Input.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 1);
        End If;
      End If;
    End If;
  End Loop;
  v_Temp := '<RESULT>True</RESULT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xmlfilelist_Out := x_Templet;

Exception
  When Err_Custom Then
    v_Temp := '<RESULT>False</RESULT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<ERROR><MSG>' || v_Error || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xmlfilelist_Out := x_Templet;
  
  When Others Then
    v_Temp := '<RESULT>False</RESULT>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := '<ERROR><MSG>' || SQLErrM || '</MSG></ERROR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Xmlfilelist_Out := x_Templet;
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Itemsave;
/





------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0038' Where ���=&n_System;
Commit;
