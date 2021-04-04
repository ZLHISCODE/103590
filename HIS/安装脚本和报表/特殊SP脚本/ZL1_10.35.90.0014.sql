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
--125489:����,2018-05-28,���Ӳ���Ԥ������ģ�����Ԥ�����վ����ʾ
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1103, 0, 0, 0, 0, 0, 0, 25, 'Ԥ�����վ����ʾ', Null, '0',
         '��Ԥ��������п���Ԥ�����Ƿ��վ������ʾ,�����վ����ʾ,��ֻ�ܲ�ѯ�ʹ���վ��ɿ��Ԥ�������˿���⣩,���������ѯ�Ͳ�������վ���Ԥ���', '1-��վ����ʾ,0-����վ����ʾ', Null,
         '�������ܷ�Ժ��ʽ���ϸ�����Ԥ����վ����ʾԤ����ҵ��', Null
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
--126591:���˺�,2018-06-01,�������ѵĴ���
--126587:���˺�,2018-06-01,�޽������ݼ����ѿ�����Ԥ��������
Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --����:�����ӿ�֧��
  --���:Xml_In:
  --<IN>
  --        <BRID>����ID</BRID>         //����ID
  --        <XM>����</XM>               //����
  --        <SFZH>���֤��</SFZH>       //���֤��
  --        <ZYID>��ҳID</ZYID>         //��ҳID
  --        <JSLX>2</JSLX>         //��������,1-����,2-סԺ��Ĭ��Ϊ 2
  --        <JE></JE>         //���ν����ܽ��
  --        <NO></NO>         //���ʵķ��õ��ݺ�(������ʵ�),Ŀǰ����������=1ʱ��ʹ��
  --        <JZKNO></JZKNO>   //���ʵľ��￨���ݺ�,Ŀǰ����������=1ʱ��ʹ��
  --        <JZSJ></JZSJ>     //����ʱ��
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>֧�������</JSKLB >
  --              <JSKH>֧������</ JSKH >
  --              <JSFS>֧����ʽ</JSFS> //֧����ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --              <JSJE>������</JSJE> //������(�������˲������ҽԺ�˿�)<SFCYJ>Ϊ1ʱΪ��Ԥ�����
  --              <JYLSH>������ˮ��</JYLSH>
  --              <ZY>ժҪ</ZY>
  --              <SFCYJ>�Ƿ��Ԥ��</SFCYJ>  //�Ƿ��Ԥ����0-���㣬1-��Ԥ��.�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�
  --              <SFXFK>�Ƿ����ѿ�</SFXFK>  //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ�
  --              <EXPENDLIST>  //��չ������Ϣ
  --                  <EXPEND>
  --                        <JYMC>��������</JYMC> //��������   �˿�ʱ,�����Ԥ������ˮ��
  --                        <JYLR>��������</JYLR> //��������   �˿�ʱ,�����Ԥ���Ľ��
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --����:Xml_Out
  --  <OUT>
  --    <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --    �D�D�������д�������˵����ȷִ��
  --    <ERROR>
  --      <MSG>������Ϣ</MSG>
  --    </ERROR>
  --  </OUT>
  --------------------------------------------------------------------------------------------------
  n_��ҳid     ������ҳ.��ҳid%Type;
  n_����id     ������ҳ.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_�����ܶ�   ����Ԥ����¼.��Ԥ��%Type;
  n_�����ʽ�� ����Ԥ����¼.��Ԥ��%Type;
  n_��������   Number(3);
  v_����Ա���� ���˽��ʼ�¼.����Ա���%Type;
  v_����Ա���� ���˽��ʼ�¼.����Ա����%Type;
  n_����id     ���˽��ʼ�¼.Id%Type;
  n_��Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  d_����ʱ��   Date;
  d_��ʼ����   Date;
  d_��������   Date;
  d_��С����   Date;
  d_�������   Date;

  n_���㿨���   ���ѿ����Ŀ¼.���%Type;
  n_ʱ������     Number(3);
  v_No           ���˽��ʼ�¼.No%Type;
  n_�����id     ҽ�ƿ����.Id%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  v_Temp         Varchar2(500);
  v_Ids          Varchar2(20000);
  x_Templet      Xmltype; --ģ��XML
  v_Err_Msg      Varchar2(200);
  v_���ݺ�       Varchar2(20000);
  v_���￨���ݺ� Varchar2(20000);
  Err_Item    Exception;
  Err_Special Exception;

  v_�����     �������׼�¼.���%Type;
  v_���ѿ����� Varchar2(20000);
  n_Number     Number(2);
  n_����id     ������ü�¼.Id%Type;
  n_��¼����   ������ü�¼.��¼����%Type;
  v_����no     ������ü�¼.No%Type;
  n_���       ������ü�¼.���%Type;
  n_��¼״̬   ������ü�¼.��¼״̬%Type;
  n_ִ��״̬   ������ü�¼.ִ��״̬%Type;
  n_δ����   ������ü�¼.ʵ�ս��%Type;
  n_���ʽ��   ������ü�¼.ʵ�ս��%Type;
  n_����     ������ü�¼.ʵ�ս��%Type;

  Type t_���ý�����ϸ Is Ref Cursor;
  c_���ý�����ϸ t_���ý�����ϸ;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX')),
         To_Number(Extractvalue(Value(A), 'IN/NO')), To_Number(Extractvalue(Value(A), 'IN/JZKNO')),
         To_Date(Extractvalue(Value(A), 'IN/JZSJ'), 'yyyy-mm-dd hh24:mi:ss'), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into n_��ҳid, n_����id, n_�����ܶ�, n_��������, v_���ݺ�, v_���￨���ݺ�, d_����ʱ��, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  n_�������� := Nvl(n_��������, 2);
  If n_�������� = 1 And Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  --0.��ؼ��
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,���������!';
    Raise Err_Item;
  End If;

  --��Աid,��Ա���,��Ա����
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := 'ϵͳ�����ϱ���Ч�Ĳ���Ա,���������!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;
  v_Err_Msg    := Null;

  For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Not (c_���׼�¼.���㿨��� Is Null Or Nvl(c_���׼�¼.�Ƿ����ѿ�, '0') = '1' Or Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 1) Then
    
      Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Number
      From Dual;
    
      If Nvl(n_Number, 0) = 1 Then
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
      Else
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
      End If;
      If v_����� Is Null Then
        v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 2) = 0 Then
        v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
        Raise Err_Special;
      End If;
    
    End If;
  End Loop;
  n_ʱ������ := Zl_Getsysparameter('���ʷ���ʱ��', 1137);

  If n_�������� = 2 Then
    Open c_���ý�����ϸ For
      Select Max(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
             Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
             Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
      From סԺ���ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
      Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
      Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, ���;
  Else
  
    If v_���ݺ� Is Null And v_���￨���ݺ� Is Null Then
      Open c_���ý�����ϸ For
        Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
               Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
               Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
        From ������ü�¼
        Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1
        Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
        Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
        
        Union All
        Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
               Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
               Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
        From סԺ���ü�¼
        Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5
        Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
        Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, ���;
    
    Elsif v_���ݺ� Is Not Null And v_���￨���ݺ� Is Not Null Then
      Open c_���ý�����ϸ For
        Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
               Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
               Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
        From ������ü�¼
        Where ����id + 0 = n_����id And ��¼״̬ <> 0 And Mod(��¼����, 10) = 2 And ���ʷ��� = 1 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_���ݺ�)) B)
        Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
        Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
        Union All
        Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
               Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
               Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
        From סԺ���ü�¼
        Where ����id + 0 = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_���￨���ݺ�)) B)
        Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
        Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By ��¼����, NO, ���;
    Elsif v_���ݺ� Is Not Null Then
      Open c_���ý�����ϸ For
        Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
               Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
               Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
        From ������ü�¼
        Where ����id + 0 = n_����id And ��¼״̬ <> 0 And Mod(��¼����, 10) = 2 And ���ʷ��� = 1 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_���ݺ�)) B)
        Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
        Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, ���;
    Else
      Open c_���ý�����ϸ For
        Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
               Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
               Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
        From סԺ���ü�¼
        Where ����id + 0 = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5 And
              NO In (Select /*+cardinality(b,10) */
                      Column_Value
                     From Table(f_Str2list(v_���￨���ݺ�)) B)
        Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
        Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
        Order By NO, ���;
    End If;
  End If;

  Select ���˽��ʼ�¼_Id.Nextval, Sysdate, Nextno(15) Into n_����id, d_����ʱ��, v_No From Dual;
  n_�����ʽ�� := 0;
  Loop
    Fetch c_���ý�����ϸ
      Into n_����id, n_��¼����, v_����no, n_���, n_��¼״̬, n_ִ��״̬, d_��С����, d_�������, n_δ����, n_���ʽ��;
    Exit When c_���ý�����ϸ%NotFound;
  
    n_�����ʽ�� := n_�����ʽ�� + Nvl(n_δ����, 0);
    If d_��ʼ���� Is Null Then
      d_��ʼ���� := d_��С����;
    Elsif d_��ʼ���� > d_��С���� Then
      d_��ʼ���� := d_��С����;
    End If;
    If d_�������� Is Null Then
      d_�������� := d_�������;
    Elsif d_�������� < d_������� Then
      d_�������� := d_�������;
    End If;
  
    If Nvl(n_���ʽ��, 0) = 0 Then
      If n_����id Is Not Null Then
        If Length(v_Ids || ',' || n_����id) > 4000 Then
          v_Ids := Substr(v_Ids, 2);
          Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
          v_Ids := '';
        End If;
        v_Ids := v_Ids || ',' || n_����id;
      Else
        Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
      End If;
    Else
      Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
    End If;
  End Loop;

  If v_Ids Is Not Null Then
    v_Ids := Substr(v_Ids, 2);
    Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
  End If;

  n_�����ʽ�� := Round(n_�����ʽ��, 6);

  If n_�����ʽ�� <> Nvl(n_�����ܶ�, 0) Then
    v_Err_Msg := '����Ľ��ʽ����ʵ�ʽ��ʽ���,���������!';
    Raise Err_Item;
  End If;

  Zl_���˽��ʼ�¼_Insert(n_����id, v_No, n_����id, d_����ʱ��, d_��ʼ����, d_��������, 0, 0, n_��ҳid, Null, n_��������, Null, n_��������, 0, n_��ҳid,
                   n_�����ܶ�);

  For r_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    v_�����   := r_���㷽ʽ.���㷽ʽ;
    n_���ʽ�� := n_���ʽ�� + Nvl(r_���㷽ʽ.������, 0);
  
    If Nvl(r_���㷽ʽ.�Ƿ��Ԥ��, 0) = 0 Then
      --����
      n_�����id := Null;
      If r_���㷽ʽ.���㿨��� Is Not Null Then
        Select Decode(Translate(Nvl(r_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
        If Nvl(r_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(���), Max(���㷽ʽ), Max(����)
            Into n_���㿨���, v_���㷽ʽ, v_�����
            From ���ѿ����Ŀ¼
            Where ��� = n_�����id And Nvl(����, 0) = 1;
          Else
            Select Max(���), Max(���㷽ʽ), Max(����)
            Into n_���㿨���, v_���㷽ʽ, v_�����
            From ���ѿ����Ŀ¼
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
          End If;
          If n_���㿨��� Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ�����ѿ���Ϣ';
            Raise Err_Item;
          End If;
          n_�����id := Null;
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(���㷽ʽ), Max(����)
            Into n_�����id, v_���㷽ʽ, v_�����
            From ҽ�ƿ����
            Where ID = n_�����id And Nvl(�Ƿ�����, 0) = 1;
          Else
            Select Max(ID), Max(���㷽ʽ), Max(����)
            Into n_�����id, v_���㷽ʽ, v_�����
            From ҽ�ƿ����
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
          End If;
        
          If n_�����id Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ���Ϣ!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_�����id Is Not Null Then
        --������
        v_���㷽ʽ := v_���㷽ʽ || '|' || r_���㷽ʽ.������ || '|';
        Zl_���˽��ʽ���_Modify(1, n_����id, n_����id, v_���㷽ʽ, Null, 0, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, 0, 0, 0,
                         n_��������, Null, v_����Ա����, v_����Ա����, d_����ʱ��, Null, 0);
        For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                              Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                       From Table(Xmlsequence(Extract(r_���㷽ʽ.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
          Zl_�������㽻��_Insert(n_�����id, 0, r_���㷽ʽ.���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
        End Loop;
      Else
        If n_���㿨��� Is Not Null Then
          --���ѿ�
          v_���ѿ����� := Nvl(v_���ѿ�����, '') || '||' || n_���㿨��� || '|' || r_���㷽ʽ.���㿨�� || '|0|' || r_���㷽ʽ.������;
        Else
          --��������
          v_���㷽ʽ := r_���㷽ʽ.���㷽ʽ || '|' || r_���㷽ʽ.������ || '||';
          Zl_���˽��ʽ���_Modify(0, n_����id, n_����id, v_���㷽ʽ, Null, 0, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, 0, 0, 0,
                           n_��������, Null, v_����Ա����, v_����Ա����, d_����ʱ��, Null, 0);
        End If;
      End If;
    Else
      --��Ԥ��,ĿǰĬ��ȫ��
      n_��Ԥ����� := r_���㷽ʽ.������;
      Zl_���˽��ʽ���_Modify(0, n_����id, n_����id, Null, n_��Ԥ�����, 0, Null, Null, Null, Null, 0, 0, 0, n_��������, Null, v_����Ա����,
                       v_����Ա����, d_����ʱ��, Null, 0);
    End If;
  
    Update �������׼�¼
    Set ҵ�����id = n_����id
    Where ��ˮ�� = Nvl(r_���㷽ʽ.������ˮ��, '-') And ��� = v_����� And ҵ������ = 2;
  
  End Loop;

  --���ѿ�����
  If v_���ѿ����� Is Not Null Then
    v_���ѿ����� := Substr(v_���ѿ�����, 3);
    Zl_���˽��ʽ���_Modify(3, n_����id, n_����id, v_���ѿ�����, Null, 0, Null, Null, Null, Null, 0, 0, 0, n_��������, Null, v_����Ա����,
                     v_����Ա����, d_����ʱ��, Null, 0);
  End If;

  n_���� := Round(Nvl(n_�����ܶ�, 0) - Nvl(n_���ʽ��, 0), 6);

  If Abs(Nvl(n_����, 0)) > 1 Then
    v_Err_Msg := '���������������1.00��С��-1.00Ԫ,��������ʲ���,����!';
    Raise Err_Item;
  End If;
   
  Zl_���˽��ʽ���_Modify(0, n_����id, n_����id, '', Null, 0, Null, Null, Null, Null, 0, 0, n_����, n_��������, Null, v_����Ա����, v_����Ա����,
                   d_����ʱ��, Null, 1);

  Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = n_����id And Nvl(У�Ա�־, 0) <> 0;
  v_Temp := '<CZSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Third_Settlement;
/

--125779:����,2018-05-28,��ҩ��ҩƷid������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_�������
(
  ��ҩid_In   In Varchar2, --ID��:ID1,��˱�־1,ID2,��˱�־2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tansids    Varchar2(4000);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_��ҩid     ҩƷ�շ���¼.Id%Type;
  n_Count      Number(1);
  d_���ʱ��   ҩƷ�շ���¼.�������%Type;
  v_No         ҩƷ�շ���¼.No%Type;
  v_�ϴ�no     ҩƷ�շ���¼.No%Type;
  n_��˱�־   Number(1);
  n_����״̬   Number(2);
  v_�շ�ids    Varchar2(4000);
  v_��ҩ����id ҩƷ�շ���¼.Id%Type;
  v_ԭʼid     ҩƷ�շ���¼.Id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;

  Cursor c_���ʼ�¼ Is
    Select Distinct a.����id, b.����ʱ��
    From ҩƷ�շ���¼ A, ��Һ��ҩ��¼ B, ��Һ��ҩ���� C
    Where a.Id = c.�շ�id And b.Id = c.��¼id And b.Id = v_Tansid And b.����״̬ = 9;

  v_���ʼ�¼ c_���ʼ�¼%RowType;

  Cursor c_��ҩ��¼ Is
    Select /*+ rule*/
    Distinct a.Id As ��ҩid, c.�շ�id, c.����, a.ҩƷid, a.����,c.��¼id as ��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ��¼ c_��ҩ��¼%RowType;

  Cursor c_�������� Is
    Select /*+ rule*/
     a.No, a.��� || ':' || c.���� || ':' || c.��¼id As �������
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.����id And b.Id = c.�շ�id And Mod(b.��¼״̬, 3) = 1 And c.��¼id = d.Column_Value;

  v_�������� c_��������%RowType;

Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_��˱�־ := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_�շ�ids := Null;
  
    --ͳ�����ȷ�ϵ���Һ��(n_��˱�־ = 1)
    If n_��˱�־ = 1 Then
      If v_Tansids Is Null Then
        v_Tansids := v_Tansid;
      Else
        v_Tansids := v_Tansids || ',' || v_Tansid;
      End If;
    End If;
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ <> 9 Then
        v_Error := '�������ѱ����������ܽ���������ˣ�';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    If n_��˱�־ = 1 Then
      n_����״̬ := 10;
    Elsif n_��˱�־ = 2 Then
      n_����״̬ := 11;
    End If;
  
    --������Һ����Ӧ���շ�NO
    Begin
      Select NO
      Into v_No
      From ҩƷ�շ���¼
      Where ID In (Select �շ�id From ��Һ��ҩ���� Where ��¼id In (Select ID From ��Һ��ҩ��¼ Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --�շ�NO��ͬ����ҩID�����ʱ���Դ�����Ϊ�ӳ�1��
    If v_No = v_�ϴ�no Then
      d_���ʱ�� := d_���ʱ�� + 1 / 24 / 60 / 60;
    Else
      d_���ʱ�� := ����ʱ��_In;
      v_�ϴ�no   := v_No;
    End If;
  
    --���ʼ�¼����
    For v_���ʼ�¼ In c_���ʼ�¼ Loop
      Zl_���˷�������_Audit(v_���ʼ�¼.����id, v_���ʼ�¼.����ʱ��, ������Ա_In, d_���ʱ��, n_��˱�־);
    End Loop;
  
    Select Count(*) Into n_Count From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And ����ʱ�� = ����ʱ��_In;
  
    If n_Count <> 1 Then
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��)
      Values
        (v_Tansid, n_����״̬, ������Ա_In, ����ʱ��_In);
    End If;
    Update ��Һ��ҩ��¼ Set ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In, ����״̬ = n_����״̬ Where ID = v_Tansid;
  End Loop;

  --����ҩ
  For v_��ҩ��¼ In c_��ҩ��¼ Loop
    Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ��¼.��ҩid, ������Ա_In, ����ʱ��_In, Null, Null, Null, v_��ҩ��¼.����, Null, ������Ա_In);
  
    --ȡ��ҩ����id
    Select a.Id
    Into v_��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
  
    --��Һ��ҩ�����е��շ�ID����Ϊ��ҩ�������շ�ID
    Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩ��¼.�շ�id;
  
    If v_�շ�ids Is Null Then
      v_�շ�ids := v_��ҩid;
    Else
      v_�շ�ids := v_�շ�ids || ',' || v_��ҩid;
    End If;
  
    --ȡԭʼid
    Select a.Id
    Into v_ԭʼid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 0 And a.������� Is Not Null;
  
    Insert Into ��Һ��ҩ����
      (��¼id, �շ�id, ����)
      Select ��¼id, v_ԭʼid, ���� From ��Һ��ҩ���� Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩid;
  
    v_�շ�ids := v_�շ�ids || ',' || v_ԭʼid;
  End Loop;

  --��������
  For v_�������� In c_�������� Loop
    Zl_סԺ���ʼ�¼_Delete(v_��������.No, v_��������.�������, v_Usercode, Zl_Username, 2, 1, 1, d_���ʱ��);
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�������;
/

--126046:����,2018-05-28,��ҩ��ҩƷid������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_��ҩ
(
  ����id_In   In ��Һ��ҩ��¼.����id%Type,
  ��ҩid_In   In Varchar2, --ID��:ID1,ID2....
  ��ҩ����_In In ��Һ��ҩ��¼.��ҩ����%Type,
  ������Ա_In In ��Һ��ҩ״̬.������Ա%Type := Null,
  ����ʱ��_In In ��Һ��ҩ״̬.����ʱ��%Type := Null,
  �ƶ�����_In In Number := 0
) Is
  v_Tansid Varchar2(20);
  v_Tmp    Varchar2(4000);

  v_�շ�ids  Varchar2(4000);
  v_Error    Varchar2(255);
  n_�Ƿ��� ��Һ��ҩ��¼.�Ƿ���%Type;
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  v_��ҩ��   Varchar2(20);
  v_��ҩ̨   Varchar2(20);
  n_��ҩ̨id Number(4);
  n_����id   Number(18);
  n_����     Number(2);
  d_����     Date;
  Err_Custom Exception;
  Cursor c_�շ���¼ Is
    Select /*+ rule*/
     a.Id, Nvl(a.����, 0) As ����
    From ҩƷ�շ���¼ A,
         (Select Distinct �շ�id
           From ��Һ��ҩ���� A, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) B
           Where a.��¼id = b.Column_Value) B
    Where a.Id = b.�շ�id And a.����� Is Null
    Order By a.ҩƷid, a.����;

  v_�շ���¼ c_�շ���¼%RowType;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ > 1 Then
        v_Error := '�������ѱ����������ܽ��з�ҩ��';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Begin
      Select �Ƿ��� Into n_�Ƿ��� From ��Һ��ҩ��¼ Where ID = v_Tansid For Update Nowait;
    Exception
      When Others Then
        v_Error := '���������û���ִ�з�ҩ�������ظ�������';
        Raise Err_Custom;
    End;
  
    v_��ҩ̨   := '';
    n_��ҩ̨id := 0;
    n_����id   := 0;
    v_��ҩ��   := '';
    Begin
      Select ����, ID, ����id, ��ҩ����, ִ��ʱ��
      Into v_��ҩ̨, n_��ҩ̨id, n_����id, n_����, d_����
      From (Select f.����, f.Id, a.����id, a.��ҩ����, a.ִ��ʱ��
             From ��Һ��ҩ��¼ A, ��Һ��ҩ���� B, ҩƷ�շ���¼ C, ��Һ̨ҩƷ���� D, ��Һ̨ F
             Where a.Id = b.��¼id And b.�շ�id = c.Id And c.ҩƷid = d.ҩƷid And d.��ҩ̨id = f.Id And c.�ⷿid = d.����id And
                   a.Id = v_Tansid
             Order By d.��ҩ̨id)
      Where Rownum = 1;
    
      Select ��ҩ��
      Into v_��ҩ��
      From ��Һ��������
      Where ����id = n_����id And ��ҩ̨id = n_��ҩ̨id And ���� = n_���� And
            ���� = To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd');
    Exception
      When Others Then
        Null;
    End;
  
    Update ��Һ��ҩ��¼
    Set ����״̬ = 2, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In, ��ҩ���� = ��ҩ����_In, ��ҩ̨ = v_��ҩ̨
    Where ID = v_Tansid;
  
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ʵ�ʹ�����Ա)
    Values
      (v_Tansid, 2, ������Ա_In, ����ʱ��_In, v_��ҩ��);
    If n_�Ƿ��� <> 0 And �ƶ�����_In = 0 Then
      Update ��Һ��ҩ��¼ Set ����״̬ = 4, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In Where ID = v_Tansid;
      Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (v_Tansid, 4, ������Ա_In, ����ʱ��_In);
    End If;
  End Loop;

  For v_�շ���¼ In c_�շ���¼ Loop
    If v_�շ�ids Is Null Then
      v_�շ�ids := v_�շ���¼.Id || ',' || v_�շ���¼.����;
    Else
      If Length(v_�շ�ids || '|' || v_�շ���¼.Id || ',' || v_�շ���¼.����) > 3950 Then
        Zl_ҩƷ�շ���¼_������ҩ(v_�շ�ids, ����id_In, ������Ա_In, ����ʱ��_In, 4, ������Ա_In, ��ҩ����_In);
        v_�շ�ids := v_�շ���¼.Id || ',' || v_�շ���¼.����;
      Else
        v_�շ�ids := v_�շ�ids || '|' || v_�շ���¼.Id || ',' || v_�շ���¼.����;
      End If;
    End If;
  End Loop;

  If Not v_�շ�ids Is Null Then
    Zl_ҩƷ�շ���¼_������ҩ(v_�շ�ids, ����id_In, ������Ա_In, ����ʱ��_In, 4, ������Ա_In, ��ҩ����_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_��ҩ;
/

--126046:����,2018-05-28,��ҩ��ҩƷid������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ����ҩ(��ҩid_In In Varchar2 --ID��:��ҩID1,��ҩID2....
                                           ) Is
  v_Id       Varchar2(20);
  v_��ҩid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_Date     Date;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;

  Cursor c_��ҩ���� Is
    Select /*+ rule*/
    Distinct c.��¼id, a.Id As ��ҩid, c.�շ�id, a.����, a.Ч��, a.����, c.���� As ��ҩ��, a.ҩƷid, a.����
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ���� c_��ҩ����%RowType;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Id  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp := Replace(',' || v_Tmp, ',' || v_Id || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Id;
    
      If n_����״̬ != 2 Then
        v_Error := '�������ѱ����������ܽ���ȡ����ҩ������';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From ��Һ��ҩ״̬
    Where ��ҩid = v_Id And �������� = 1 And Rownum = 1;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 1, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Id;
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ����ҩ���Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Id, 1, v_������Ա, Sysdate, 'ȡ����ҩ');
  
  End Loop;

  Select Sysdate Into v_Date From Dual;

  For v_��ҩ���� In c_��ҩ���� Loop
    --������ҩ
    Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ����.��ҩid, Zl_Username, v_Date, v_��ҩ����.����, v_��ҩ����.Ч��, v_��ҩ����.����, v_��ҩ����.��ҩ��, Null, Zl_Username);
  
    Select Max(a.Id)
    Into v_��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ����.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
  
    --�滻��Һ��ҩ�����е��շ�ID
    Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩ����.��¼id And �շ�id = v_��ҩ����.�շ�id;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ����ҩ;
/

------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0014' Where ���=&n_System;
Commit;
