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
--129387:����,2018-07-26,���صĲ��˻�����Ϣ������������
CREATE OR REPLACE Procedure Zl_Third_Getpatiinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:��ȡ���˻�����Ϣ
  --���:Xml_In: 
  --  <IN> 
  --      <BRID></BRID>     --����ID
  --      <SFZH></SFZH>     --���֤��
  --      <CXKLB></CXKLB>   --��ѯ�����
  --      <MZH></MZH>       --�����
  --      <GHDH></GHDH>     --�Һŵ���
  --      <YLKLB></YLKLB>   --ҽ�ƿ����ID��������
  --      <YLKH></YLKH>     --ҽ�ƿ���
  --      <BRXM></BRXM>     --��������
  --  </IN> 
  --����:Xml_Out 
  -- <OUTPUT>
  --   <BR>
  --     <BRID></BRID>       --����ID
  --     <XM></XM>           --����
  --     <XB></XB>           --�Ա�
  --     <Nl></NL>           --����
  --     <CSRQ></CSRQ>       --��������
  --     <MZH></MZH>         --�����
  --     <HY></HY>           --����
  --     <GJ></GJ>           --����
  --     <MZ></MZ>           --����
  --     <XL></XL>           --ѧ��
  --     <SF></SF>           --���
  --     <ZY></ZY>           --ְҵ
  --     <SFZH></SFZH>       --���֤��
  --     <FKFS></FKFS>       --���ʽ
  --     <LXFS></LXFS>       --��ϵ��ʽ
  --     <LXRXM></LXRXM>     --��ϵ������
  --     <LXRDH></LXRDH>     --��ϵ�˵绰
  --     <LXRDZ></LXRDZ>     --��ϵ�˵�ַ
  --     <LXDH></LXDH>       --��ϵ�绰
  --     <XJZDZ></XJZDZ>     --�־�ס��ַ 
  --     <HJDZ></HJDZ>       --������ַ
  --     <CSDD></CSDD>       --�����ص�
  --     <KSID></KSID>       --����ID
  --     <CXKH></CXKH>       --��ѯ����
  --     <GMS></GMS>         --����ʷ         
  --     <GHD></GHD>         --�Һŵ���
  --     <GHSJ></GHSJ>       --�Һ�ʱ��
  --     <JZSJ></JZSJ>       --����ʱ��
  --     <JZKS></JZKS>       --�������
  --     <JZYS></JZYS>       --����ҽ��
  --   </BR>
  -- </OUTPUT>
  -------------------------------------------------------------------------------------------------- 

  v_����ids      varchar2(30000);
  v_ҽ�ƿ�       varchar2(500);
  v_�����       varchar2(500);
  v_�Һŵ�       ���˹Һż�¼.No%Type;
  v_����         ����ҽ�ƿ���Ϣ.����%Type;
  v_����         ������Ϣ.����%Type;
  v_���֤��     ������Ϣ.���֤��%Type;
  v_��ѯ�����   varchar2(20);
  n_��ѯ�����id ����ҽ�ƿ���Ϣ.�����id%Type;
  n_�����id     ҽ�ƿ����.Id%Type;
  v_No           ���˹Һż�¼.No%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  d_�Һ�ʱ��     ���˹Һż�¼.�Ǽ�ʱ��%Type;
  d_����ʱ��     ���˹Һż�¼.ִ��ʱ��%Type;
  v_�������     ���ű�.����%Type;
  v_����ҽ��     ���˹Һż�¼.ִ����%Type;
  v_����ʷ       ���˹�����¼.ҩ����%Type;
  v_Temp         varchar2(32767); --��ʱXML 
  x_Templet      Xmltype; --ģ��XML 
  v_Err_Msg      varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/YLKLB'), Extractvalue(Value(A), 'IN/YLKH'),
         Extractvalue(Value(A), 'IN/BRXM'), Extractvalue(Value(A), 'IN/MZH'), Extractvalue(Value(A), 'IN/GHDH'),
         Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/CXKLB')
  Into v_���֤��, v_ҽ�ƿ�, v_����, v_����, v_�����, v_�Һŵ�, v_����ids, v_��ѯ�����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_��ѯ����� Is Not Null Then
    Select Max(ID) Into n_��ѯ�����id From ҽ�ƿ���� Where ���� = v_��ѯ�����;
    If n_��ѯ�����id Is Null Then
      n_��ѯ�����id := To_Number(v_��ѯ�����);
    End If;
  End If;

  If v_����ids Is Null Then
  
    If v_���֤�� Is Null And v_ҽ�ƿ� Is Null And v_���� Is Null And v_���� Is Null And v_����� Is Null And v_�Һŵ� Is Null Then
      v_Err_Msg := 'δ�����κ�����,�޷���ɲ�ѯ!';
      Raise Err_Item;
    End If;
  
    If v_ҽ�ƿ� Is Not Null Then
      Select Max(ID) Into n_�����id From ҽ�ƿ���� Where ���� = v_ҽ�ƿ�;
      If n_�����id Is Null Then
        n_�����id := To_Number(v_ҽ�ƿ�);
      End If;
    End If;
  
    If v_�Һŵ� Is Null Then
      If Nvl(n_�����id, 0) = 0 Then
        For r_���� In (Select Distinct ����id
                     From ������Ϣ
                     Where Nvl(���֤��, '-') = Nvl(v_���֤��, Nvl(���֤��, '-')) And ���� = Nvl(v_����, ����) And
                           Nvl(�����, 0) = Nvl(v_�����, Nvl(�����, 0))) Loop
          v_����ids := v_����ids || ',' || r_����.����id;
        End Loop;
      Else
        For r_���� In (Select Distinct a.����id
                     From ������Ϣ A, ����ҽ�ƿ���Ϣ B
                     Where a.����id = b.����id And b.�����id = n_�����id And b.���� = v_���� And
                           Nvl(a.���֤��, '-') = Nvl(v_���֤��, Nvl(a.���֤��, '-')) And a.���� = Nvl(v_����, a.����) And
                           Nvl(�����, 0) = Nvl(v_�����, Nvl(�����, 0))) Loop
          v_����ids := v_����ids || ',' || r_����.����id;
        End Loop;
      End If;
    Else
      If Nvl(n_�����id, 0) = 0 Then
        For r_���� In (Select Distinct a.����id
                     From ������Ϣ A, ���˹Һż�¼ B
                     Where a.����id = b.����id And b.No = v_�Һŵ� And Nvl(a.���֤��, '-') = Nvl(v_���֤��, Nvl(a.���֤��, '-')) And
                           a.���� = Nvl(v_����, a.����) And Nvl(a.�����, 0) = Nvl(v_�����, Nvl(a.�����, 0))) Loop
          v_����ids := v_����ids || ',' || r_����.����id;
        End Loop;
      Else
        For r_���� In (Select Distinct a.����id
                     From ������Ϣ A, ����ҽ�ƿ���Ϣ B, ���˹Һż�¼ C
                     Where a.����id = c.����id And c.No = v_�Һŵ� And a.����id = b.����id And b.�����id = n_�����id And b.���� = v_���� And
                           Nvl(a.���֤��, '-') = Nvl(v_���֤��, Nvl(a.���֤��, '-')) And a.���� = Nvl(v_����, a.����) And
                           Nvl(a.�����, 0) = Nvl(v_�����, Nvl(a.�����, 0))) Loop
          v_����ids := v_����ids || ',' || r_����.����id;
        End Loop;
      End If;
    End If;
  
    If v_����ids Is Not Null Then
      v_����ids := Substr(v_����ids, 2);
    End If;
  End If;

  For r_�Һ� In (Select c.����id, c.��ǰ����id, c.�����, c.����, c.�Ա�,c.����,c.����״��, c.����,c.��������, c.���֤��, c.ְҵ, c.ѧ��, c.����, 
                        c.��ͥ�绰, c.��ͥ��ַ, c.���ڵ�ַ,c.���,c.�ֻ���,c.��ϵ������,c.��ϵ�˵绰,c.��ϵ�˵�ַ,c.�����ص�,
                      Max(f.����) As ����
               From ������Ϣ C, Table(f_Str2list(v_����ids)) E, ����ҽ�ƿ���Ϣ F
               Where c.����id = e.Column_Value And c.����id = f.����id(+) And f.�����id(+) = n_��ѯ�����id And Nvl(f.״̬, 0) = 0
               Group By c.����id, c.��ǰ����id, c.�����, c.����, c.�Ա�, c.��������, c.���֤��, c.ְҵ, c.ѧ��, c.����, c.��ͥ�绰, c.��ͥ��ַ, c.���ڵ�ַ
                        ,c.���,c.�ֻ���,c.��ϵ������,c.��ϵ�˵绰,c.��ϵ�˵�ַ,c.�����ص�,c.����,c.����״��, c.����
               ) Loop
    v_Temp := '<BR>';

    Select Max(No), Max(ҽ�Ƹ��ʽ), Max(�Ǽ�ʱ��), Max(ִ��ʱ��), Max(ִ����), Max(�������)
    Into v_No, v_���ʽ, d_�Һ�ʱ��, d_����ʱ��, v_����ҽ��, v_�������
    From (Select a.No, a.ҽ�Ƹ��ʽ, a.�Ǽ�ʱ��, a.ִ��ʱ��, a.ִ����, b.���� As �������
           From ���˹Һż�¼ a, ���ű� b
           Where a.ִ�в���id = b.Id(+)  And a.����id = r_�Һ�.����id And a.��¼���� = 1 And a.��¼״̬ = 1
           Order By a.�Ǽ�ʱ�� Desc)
    Where Rownum < 2;
    
    For R In(Select ҩ���� From ���˹�����¼ Where ����ID=r_�Һ�.����id)Loop
      v_����ʷ := v_����ʷ || ',' || r.ҩ����;      
    End Loop;
    v_����ʷ := Substr(v_����ʷ, 2);
    
    v_Temp := v_Temp || '<BRID>' || r_�Һ�.����id || '</BRID>';
    v_Temp := v_Temp || '<XM>' || r_�Һ�.���� || '</XM>';
    v_Temp := v_Temp || '<XB>' || r_�Һ�.�Ա� || '</XB>';
    v_Temp := v_Temp || '<NL>' || r_�Һ�.���� || '</NL>';
    v_Temp := v_Temp || '<CSRQ>' || To_Char(r_�Һ�.��������, 'yyyy-mm-dd hh24:mi:ss') || '</CSRQ>';
    v_Temp := v_Temp || '<MZH>' || r_�Һ�.����� || '</MZH>'; 
    v_Temp := v_Temp || '<HY>' || r_�Һ�.����״�� || '</HY>';
    v_Temp := v_Temp || '<GJ>' || r_�Һ�.���� || '</GJ>';     
    v_Temp := v_Temp || '<MZ>' || r_�Һ�.���� || '</MZ>';
    v_Temp := v_Temp || '<XL>' || r_�Һ�.ѧ�� || '</XL>';   
    v_Temp := v_Temp || '<SF>' || r_�Һ�.��� || '</SF>';   
    v_Temp := v_Temp || '<ZY>' || r_�Һ�.ְҵ || '</ZY>';      
    v_Temp := v_Temp || '<SFZH>' || r_�Һ�.���֤�� || '</SFZH>';
    v_Temp := v_Temp || '<FKFS>' || v_���ʽ || '</FKFS>';    
    v_Temp := v_Temp || '<LXFS>' || r_�Һ�.�ֻ��� || '</LXFS>';
    v_Temp := v_Temp || '<LXRXM>' || r_�Һ�.��ϵ������ || '</LXRXM>';
    v_Temp := v_Temp || '<LXRDH>' || r_�Һ�.��ϵ�˵绰 || '</LXRDH>';
    v_Temp := v_Temp || '<LXRDZ>' || r_�Һ�.��ϵ�˵�ַ || '</LXRDZ>';        
    v_Temp := v_Temp || '<LXDH>' || r_�Һ�.��ͥ�绰 || '</LXDH>';
    v_Temp := v_Temp || '<XJZDZ>' || r_�Һ�.��ͥ��ַ || '</XJZDZ>';
    v_Temp := v_Temp || '<HJDZ>' || r_�Һ�.���ڵ�ַ || '</HJDZ>';
    v_Temp := v_Temp || '<CSDD>' || r_�Һ�.�����ص� || '</CSDD>';                  
    v_Temp := v_Temp || '<KSID>' || r_�Һ�.��ǰ����id || '</KSID>';    
    v_Temp := v_Temp || '<CXKH>' || r_�Һ�.���� || '</CXKH>';
    v_Temp := v_Temp || '<GMS>' || v_����ʷ || '</GMS>';
    v_Temp := v_Temp || '<GHD>' || v_No || '</GHD>';
    v_Temp := v_Temp || '<GHSJ>' || To_Char(d_�Һ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</GHSJ>';
    v_Temp := v_Temp || '<JZSJ>' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</JZSJ>';
    v_Temp := v_Temp || '<JZKS>' || v_������� || '</JZKS>';
    v_Temp := v_Temp || '<JZYS>' || v_����ҽ�� || '</JZYS>';
    v_Temp := v_Temp || '</BR>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpatiinfo;
/

--127075:����,2018-07-24,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_���˹Һż�¼_ת��
(
  No_In         In ���˹Һż�¼.No%Type,
  ת��״̬_In   In ���˹Һż�¼.ת��״̬%Type,
  ת�����id_In In ���˹Һż�¼.ת�����id%Type := Null,
  ת������_In   In ���˹Һż�¼.ת������%Type := Null,
  ת��ҽ��_In   In ���˹Һż�¼.ת��ҽ��%Type := Null
  --���ܣ���ɲ���ת�ת����գ�ȡ��ת��ܾ�ת�﹦��
  --������
  ----ת��״̬_IN��0:ת��(��Ҫ������������),1:����,-1:�ܾ�,Null:ȡ��ת��
) As
  v_����id   ���˹Һż�¼.����id%Type;
  v_ת��״̬ ���˹Һż�¼.ת��״̬%Type;

  n_�ٴ�ǩ�������Ŷ� Number;
  n_����̨ǩ���Ŷ�   Number;
  v_Temp             Varchar2(255);
  v_��Ա����         ������ü�¼.����Ա����%Type;
  v_��������         �ŶӽкŶ���.��������%Type;
  v_�ֶ�������       �ŶӽкŶ���.��������%Type;
  v_��������         ���˹Һż�¼.����%Type;
  v_ҽ��             ���˹Һż�¼.ִ����%Type;
  v_����             ���˹Һż�¼.����%Type;
  n_�Һ�id           ���˹Һż�¼.Id%Type;
  n_ִ�в���id       ���˹Һż�¼.ִ�в���id%Type;
  v_�ű�             ���˹Һż�¼.�ű�%Type;
  n_����             ���˹Һż�¼.����%Type;
  d_Cur              Date;
  v_Error            Varchar2(255);
  Err_Custom Exception;
  n_�Ŷ�       Number(2);
  v_�ŶӺ���   �ŶӽкŶ���.�ŶӺ���%Type;
  v_���ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�Ŷ����   �ŶӽкŶ���.�Ŷ����%Type;
  d_���Ŷ�ʱ�� �ŶӽкŶ���.�Ŷ�ʱ��%Type;
Begin
  Begin
    Select ����id, ת��״̬, ID
    Into v_����id, v_ת��״̬, n_�Һ�id
    From ���˹Һż�¼
    Where NO = No_In And ��¼״̬ = 1 And ��¼���� = 1;
  Exception
    When Others Then
      Begin
        v_Error := '���˵ĹҺż�¼�����ڣ������Ѿ��˺š�';
        Raise Err_Custom;
      End;
  End;

  n_�ٴ�ǩ�������Ŷ� := Zl_To_Number(zl_GetSysParameter('�ٴ�ǩ���������Ŷ�', 1113));

  If ת��״̬_In Is Null Then
    --ȡ��ת��
    If Not (v_ת��״̬ = 0 Or v_ת��״̬ = -1) Or v_ת��״̬ Is Null Then
      v_Error := '���˵�ǰ������ת������ջ򱻾ܾ�״̬������ȡ��ת�';
      Raise Err_Custom;
    End If;
  
    Update ���˹Һż�¼
    Set ת��״̬ = Null, ת��ű� = Null, ת�����id = Null, ת������ = Null, ת��ҽ�� = Null
    Where NO = No_In;
  
    Begin
      Select 1 Into n_�Ŷ� From �ŶӽкŶ��� Where Nvl(ҵ������, 0) = 0 And ҵ��id = n_�Һ�id;
    Exception
      When Others Then
        n_�Ŷ� := -1;
    End;
  
    If Nvl(n_�Ŷ�, 0) <> 0 Then
      Update ���˹Һż�¼ Set ��¼��־ = 1 Where NO = No_In;
      Begin
        Select ID, ִ�в���id, ����, ִ����, ����, �ű�, Nvl(����, 0)
        Into n_�Һ�id, n_ִ�в���id, v_��������, v_ҽ��, v_����, v_�ű�, n_����
        From ���˹Һż�¼
        Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1 And Rownum = 1;
      Exception
        When Others Then
          n_�Һ�id := -1;
      End;
      If n_�Һ�id > 0 Then
        --ȡ��ת��Ҳֻ�����»�ȡ����
        v_�ֶ������� := n_ִ�в���id;
        --Zlgetnextqueue(ִ�в���id_In Number,ҵ��id_In     Number := Null)
        v_�ŶӺ��� := Zlgetnextqueue(n_ִ�в���id, n_�Һ�id, v_�ű� || '|' || n_����);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 1);
        --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In,�ŶӺ���_In
        Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_�Һ�id, n_ִ�в���id, v_��������, v_����, v_ҽ��, v_�ŶӺ���, v_�Ŷ����);
        --ת�����»�ȡ����
      End If;
    End If;
  
  Elsif ת��״̬_In = 0 Then
    --ת��
    If Not (v_ת��״̬ Is Null Or v_ת��״̬ = 1) Then
      v_Error := '���˵�ǰ�Ѿ�ת������������ٽ���ת�';
      Raise Err_Custom;
    End If;
  
    Update ���˹Һż�¼
    Set ת��״̬ = 0, ת��ű� = �ű�, ת�����id = ת�����id_In, ת������ = ת������_In, ת��ҽ�� = ת��ҽ��_In
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1
    Returning ID, ִ�в���id Into n_�Һ�id, n_ִ�в���id;
  
    Begin
      Select 1, �������� Into n_�Ŷ�, v_�������� From �ŶӽкŶ��� Where Nvl(ҵ������, 0) = 0 And ҵ��id = n_�Һ�id;
    Exception
      When Others Then
        n_�Ŷ� := -1;
    End;
    n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(n_ִ�в���id, 0)));
    If Nvl(n_�Ŷ�, 0) <> 0 Then
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 1 Then
        If Nvl(v_��������, 0) <> 0 And Nvl(n_�ٴ�ǩ�������Ŷ�, 0) = 1 Then
          --ɾ��ԭ���ŶӼ�¼�����Ŷӣ���������_IN��ҵ��ID_IN
          Zl_�ŶӽкŶ���_Delete(v_��������, n_�Һ�id);
        End If;
        Update ���˹Һż�¼ Set ��¼��־ = 0 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
      Else
        Begin
          Select ID, ִ�в���id, ����, �ű�, ����
          Into n_�Һ�id, n_ִ�в���id, v_��������, v_�ű�, n_����
          From ���˹Һż�¼
          Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1 And Rownum = 1;
        Exception
          When Others Then
            n_�Һ�id := -1;
        End;
      
        v_�ֶ������� := ת�����id_In;
        Begin
          Select �ŶӺ��� Into v_�ŶӺ��� From �ŶӽкŶ��� Where ҵ��id = n_�Һ�id And ҵ������ = 0;
        Exception
          When Others Then
            v_�ŶӺ��� := -1;
        End;
        If n_�Һ�id > 0 Then
          v_���ŶӺ��� := Zl_Get_Requeue(2, n_�Һ�id, ת�����id_In, ת��ҽ��_In, ת������_In);
          If v_�ŶӺ��� <> v_���ŶӺ��� Or Nvl(n_�ٴ�ǩ�������Ŷ�, 0) = 1 Then
            d_���Ŷ�ʱ�� := Zl_Get_Requeuedate(2, n_�Һ�id, ת�����id_In, ת��ҽ��_In, ת������_In);
            --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In
            Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_�Һ�id, ת�����id_In, v_��������, ת������_In, ת��ҽ��_In, v_���ŶӺ���, Null, d_���Ŷ�ʱ��);
          Else
            --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In
            Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_�Һ�id, ת�����id_In, v_��������, ת������_In, ת��ҽ��_In);
          End If;
          --ת���,�����Ŷ�
          Update �ŶӽкŶ��� Set �Ŷ�״̬ = 0 Where ҵ������ = 0 And ҵ��id = n_�Һ�id;
        End If;
      End If;
    End If;
  Elsif ת��״̬_In = 1 Then
    --����
    If v_ת��״̬ <> 0 Or v_ת��״̬ Is Null Then
      v_Error := '���˵�ǰ������ת�������״̬�����ܽ���ת�';
      Raise Err_Custom;
    End If;
  
    --��ǰ������Ա����Ȼת��ָ���ˣ���ʵ�ʽ���Ŀ��ܲ���ԭָ���ġ�
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    d_Cur      := Sysdate;
  
    --ת���������ǿ�ƽ���һ����ֻ����ִ�������Ϣ
    --ԭת��ʱָ���������ڽ���ʱ��һ����ָ����һ��
    --ԭת��ʱָ�����ݵ����ã�1.ȷ��ת����ٽ���ķ�Χ��2.���顣
    Insert Into ����ת���¼
      (�Һ�id, NO, �������id, ����ҽ��, ���տ���id, ����ҽ��, ����ʱ��)
      Select ID, No_In, ִ�в���id, ִ����, ת�����id, v_��Ա����, d_Cur From ���˹Һż�¼ Where NO = No_In;
  
    Update ������Ϣ Set ����״̬ = 2, ����ʱ�� = d_Cur Where ����id = v_����id;
    Update ���˹Һż�¼
    Set ִ���� = v_��Ա����, ִ�в���id = ת�����id, ִ��״̬ = 2, ִ��ʱ�� = d_Cur, ת��״̬ = 1
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1
    Returning ת�����id, ID Into n_ִ�в���id, n_�Һ�id;
  
    Update ������ü�¼
    Set ִ���� = v_��Ա����, ���˿���id = n_ִ�в���id, ִ�в���id = n_ִ�в���id, ִ��״̬ = 2, ִ��ʱ�� = d_Cur
    Where NO = No_In And ��¼���� = 4;
  
    --�����,�������
    Update �ŶӽкŶ��� Set �Ŷ�״̬ = 2 Where ҵ������ = 0 And ҵ��id = n_�Һ�id;
  
  Elsif ת��״̬_In = -1 Then
    --�ܾ�
    If v_ת��״̬ <> 0 Or v_ת��״̬ Is Null Then
      v_Error := '���˵�ǰ������ת�������״̬�����ܾܾ�ת�';
      Raise Err_Custom;
    End If;
    Update ���˹Һż�¼ Set ת��״̬ = -1 Where NO = No_In;
    --�����,�������
    Update �ŶӽкŶ��� Set �Ŷ�״̬ = 2 Where ҵ������ = 0 And ҵ��id = n_�Һ�id;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_ת��;


/

--127075:����,2018-07-24,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
Create Or Replace Procedure Zl_���˹Һż�¼_����
(
  No_In         ���˹Һż�¼.No%Type,
  �ű�_In       ���˹Һż�¼.�ű�%Type,
  ����_In       ���˹Һż�¼.����%Type,
  ����id_In     ���˹Һż�¼.ִ�в���id%Type,
  ԭҽ��_In     ���˹Һż�¼.ִ����%Type,
  ԭҽ��id_In   ���˹ҺŻ���.ҽ��id%Type,
  ��ҽ��_In     ���˹Һż�¼.ִ����%Type,
  ��ҽ��id_In   ���˹ҺŻ���.ҽ��id%Type,
  �����¼id_In �ٴ������¼.Id%Type := Null
  --���ܣ���ɲ��˻��Ź��ܣ��ڹҺ���ĿID��ͬ������¡�
) As
  Cursor c_Bill Is
    Select a.Id, a.��¼����, a.No, a.ʵ��Ʊ��, a.��¼״̬, b.����, a.���, a.��������, a.�۸񸸺�, a.���ʵ�id, a.����id, a.ҽ�����, a.�����־, a.���ʷ���, a.����,
           a.�Ա�, a.����, a.��ʶ��, a.���ʽ, a.���˿���id, a.�ѱ�, �շ����, a.�շ�ϸĿid, a.���㵥λ, a.����, a.��ҩ����, a.����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����,
           a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��, a.������, a.��������id, a.������, b.����ʱ��, a.�Ǽ�ʱ��, a.ִ�в���id, a.ִ����, a.ִ��״̬,
           a.ִ��ʱ��, a.����, a.����Ա���, a.����Ա����, a.����id, a.���ʽ��, a.���մ���id, a.������Ŀ��, a.���ձ���, a.��������, a.ͳ����, a.�Ƿ��ϴ�, a.ժҪ,
           a.�Ƿ���
    From ������ü�¼ A, ���˹Һż�¼ B
    Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.No = No_In And a.No = b.No
    Order By a.���;

  v_����id           ������ü�¼.Id%Type;
  v_��������         �ŶӽкŶ���.��������%Type;
  v_�ֶ�������       �ŶӽкŶ���.��������%Type;
  v_�Һ����ɶ���     Varchar2(2);
  n_����̨ǩ���Ŷ�   Number;
  n_�ٴ�ǩ�������Ŷ� Number;
  v_ԤԼ�Һ�         Number(2);
  n_ҵ��id           ���˹Һż�¼.Id%Type;
  v_�ŶӺ���         �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ű�             ���˹Һż�¼.�ű�%Type;
  n_����             ���˹Һż�¼.����%Type;
  v_�Ŷ����         �ŶӽкŶ���.�Ŷ����%Type;
  d_�Ŷ�ʱ��         �ŶӽкŶ���.�Ŷ�ʱ��%Type;
  v_Temp             Varchar2(500);
  v_����Ա���       ����䶯��¼.����Ա���%Type;
  v_����Ա����       ����䶯��¼.����Ա����%Type;
  n_ҽ��id           ��Ա��.Id%Type;
  n_����id           ��������.Id%Type;
  n_ԭ�����¼id     �ٴ������¼.Id%Type;
  n_�䶯id           ����䶯��¼.Id%Type;
  v_Error            Varchar2(255);
  n_Exists           Number(3);
  n_ԭ���           �ٴ�������ſ���.���%Type;
  n_ԭԤԼ˳���     �ٴ�������ſ���.ԤԼ˳���%Type;
  n_ԭ�Һ�״̬       �ٴ�������ſ���.�Һ�״̬%Type;
  v_ԭ����Ա         �ٴ�������ſ���.����Ա����%Type;
  Err_Custom Exception;
Begin
  v_����id := 0;
  If �����¼id_In Is Null Then
    Begin
      Select ����id Into v_����id From ���˹Һż�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
    Exception
      When Others Then
        Null;
    End;
    If v_����id = 0 Then
      v_Error := 'û���ҵ����˵ĹҺ���Ϣ��';
      Raise Err_Custom;
    Elsif v_����id Is Null Then
      v_Error := 'û���ҵ�������Ϣ��';
      Raise Err_Custom;
    End If;
  
    ---�ȸ��²�����Ϣ�ľ������Һ�״̬
    Update ������Ϣ Set �������� = ����_In, ����״̬ = 1 Where ����id = v_����id And ����״̬ In (1, 2);
  
    For r_Bill In c_Bill Loop
      If r_Bill.��� = 1 Then
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into v_ԤԼ�Һ� From ���˹Һż�¼ Where NO = r_Bill.No And Rownum = 1;
        Exception
          When Others Then
            v_ԤԼ�Һ� := 0;
        End;
      
        --�ָ���ǰ�ĹҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(����id, 0) = Nvl(r_Bill.ִ�в���id, 0) And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = r_Bill.���㵥λ Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.����ʱ��), r_Bill.ִ�в���id, r_Bill.�շ�ϸĿid, ԭҽ��_In, Decode(ԭҽ��id_In, 0, Null, ԭҽ��id_In), r_Bill.���㵥λ,
             -1, -1 * v_ԤԼ�Һ�, -1 * v_ԤԼ�Һ�);
        End If;
      
        ----Ȼ���ٸ��¹ҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) + v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(����id, 0) = ����id_In And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = �ű�_In Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.����ʱ��), ����id_In, r_Bill.�շ�ϸĿid, ��ҽ��_In, Decode(��ҽ��id_In, 0, Null, ��ҽ��id_In), �ű�_In, 1, v_ԤԼ�Һ�,
             v_ԤԼ�Һ�);
        End If;
      
        --�������״̬
        Select Count(1)
        Into n_Exists
        From �Һ����״̬
        Where ���� = �ű�_In And Trunc(����) = Trunc(r_Bill.����ʱ��) And ��� = r_Bill.���� And Nvl(״̬, 0) <> 0;
      
        If n_Exists = 0 Then
          Update �Һ����״̬
          Set ���� = �ű�_In
          Where Trunc(����) = Trunc(r_Bill.����ʱ��) And ���� = r_Bill.���㵥λ And ��� = r_Bill.����;
        Else
          Delete From �Һ����״̬
          Where Trunc(����) = Trunc(r_Bill.����ʱ��) And ���� = r_Bill.���㵥λ And ��� = r_Bill.����;
          Update ���˹Һż�¼ Set ���� = Null Where NO = r_Bill.No;
        End If;
      End If;
    
      ---���¹Һż�¼
      Update ������ü�¼
      Set ִ�в���id = ����id_In, ���˿���id = ����id_In, ���㵥λ = �ű�_In, ��ҩ���� = ����_In,
          --���˲���id = ����id_In,
          ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null
      Where ID = r_Bill.Id;
    
      --���²��˹Һż�¼
      If r_Bill.��� = 1 Then
        v_Temp := Zl_Identity(1);
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
        Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
        Begin
          Select ID Into n_ҽ��id From ��Ա�� Where ���� = ��ҽ��_In And Rownum < 2;
        Exception
          When Others Then
            n_ҽ��id := Null;
        End;
        Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
        Zl_����䶯��¼_Insert(r_Bill.No, 2, '���ﻻ��', v_����Ա����, v_����Ա���, �ű�_In, ����id_In, Null, n_ҽ��id, ��ҽ��_In, ����_In, n_����,
                         Null, n_�䶯id);
        v_�Һ����ɶ���     := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        n_����̨ǩ���Ŷ�   := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(����id_In, 0)));
        n_�ٴ�ǩ�������Ŷ� := Zl_To_Number(zl_GetSysParameter('�ٴ�ǩ���������Ŷ�', 1113));
      
        Select ID, �ű�, Nvl(����, 0)
        Into n_ҵ��id, v_�ű�, n_����
        From ���˹Һż�¼
        Where NO = r_Bill.No And Rownum = 1;
      
        If v_�Һ����ɶ��� <> 0 Then
          If Nvl(n_����̨ǩ���Ŷ�, 0) = 1 Then
            Select �������� Into v_�������� From �ŶӽкŶ��� Where ҵ��id = n_ҵ��id;
            If Nvl(v_��������, 0) <> 0 And Nvl(n_�ٴ�ǩ�������Ŷ�, 0) = 1 Then
              --ɾ��ԭ���ŶӼ�¼�����Ŷӣ���������_IN��ҵ��ID_IN
              Zl_�ŶӽкŶ���_Delete(v_��������, n_ҵ��id);
            Else
              Update �ŶӽкŶ��� Set �Ŷ�״̬ = 2 Where ҵ��id = n_ҵ��id And ҵ������ = 0;
            End If;
            Update ���˹Һż�¼ Set ��¼��־ = 0 Where ID = n_ҵ��id;
          Else
            v_�ֶ������� := ����id_In;
            --Zlgetnextqueue(ִ�в���id_In Number,ҵ��id_In     Number := Null)
            v_�ŶӺ��� := Zlgetnextqueue(����id_In, n_ҵ��id, v_�ű� || '|' || n_����);
            v_�Ŷ���� := Zlgetsequencenum(0, n_ҵ��id, 1);
            d_�Ŷ�ʱ�� := Zl_Get_Requeuedate(3, n_ҵ��id, ����id_In, ��ҽ��_In, ����_In);
            --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In
            Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_ҵ��id, ����id_In, r_Bill.����, ����_In, ��ҽ��_In, v_�ŶӺ���, v_�Ŷ����, d_�Ŷ�ʱ��);
            --���ź���¶�����Ϣ���Ŷ�״̬Ҳ����Ϊ�Ŷ���
            Update �ŶӽкŶ��� Set �Ŷ�״̬ = 0 Where ҵ��id = n_ҵ��id And ҵ������ = 0;
          End If;
        End If;
        --ɾ��ת����Ϣ
        Update ���˹Һż�¼
        Set ִ�в���id = ����id_In, �ű� = �ű�_In, ���� = ����_In, ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null, ת��ű� = Null, ת�����id = Null,
            ת������ = Null, ת��ҽ�� = Null, ת��״̬ = Null
        Where NO = r_Bill.No;
      End If;
    End Loop;
  Else
    --������Ű�ģʽ
    Begin
      Select ����id, �����¼id
      Into v_����id, n_ԭ�����¼id
      From ���˹Һż�¼
      Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
      Select ID Into n_����id From �������� Where ���� = ����_In;
    Exception
      When Others Then
        Null;
    End;
    If v_����id = 0 Then
      v_Error := 'û���ҵ����˵ĹҺ���Ϣ��';
      Raise Err_Custom;
    Elsif v_����id Is Null Then
      v_Error := 'û���ҵ�������Ϣ��';
      Raise Err_Custom;
    End If;
  
    ---�ȸ��²�����Ϣ�ľ������Һ�״̬
    Update ������Ϣ Set �������� = ����_In, ����״̬ = 1 Where ����id = v_����id And ����״̬ In (1, 2);
  
    For r_Bill In c_Bill Loop
      If r_Bill.��� = 1 Then
      
        --��Ҫȷ���Ƿ�ԤԼ�Һ�
        --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
        --2.����������Һ�,��ֻ���ѹ���
        Begin
          Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into v_ԤԼ�Һ� From ���˹Һż�¼ Where NO = r_Bill.No And Rownum = 1;
        Exception
          When Others Then
            v_ԤԼ�Һ� := 0;
        End;
      
        --�ָ���ǰ�ĹҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(ԭҽ��id_In, 0) And Nvl(ҽ������, '-') = Nvl(ԭҽ��_In, '-') And
              Nvl(����id, 0) = Nvl(r_Bill.ִ�в���id, 0) And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = r_Bill.���㵥λ Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.����ʱ��), r_Bill.ִ�в���id, r_Bill.�շ�ϸĿid, ԭҽ��_In, Decode(ԭҽ��id_In, 0, Null, ԭҽ��id_In), r_Bill.���㵥λ,
             -1, -1 * v_ԤԼ�Һ�, -1 * v_ԤԼ�Һ�);
        End If;
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - v_ԤԼ�Һ�
        Where ID = n_ԭ�����¼id;
      
        ----Ȼ���ٸ��¹ҺŻ���
        Update ���˹ҺŻ���
        Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) + v_ԤԼ�Һ�
        Where ���� = Trunc(r_Bill.����ʱ��) And Nvl(����id, 0) = ����id_In And Nvl(��Ŀid, 0) = Nvl(r_Bill.�շ�ϸĿid, 0) And
              (���� = �ű�_In Or ���� Is Null);
        If Sql%RowCount = 0 Then
          Insert Into ���˹ҺŻ���
            (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, ��Լ��, �����ѽ���)
          Values
            (Trunc(r_Bill.����ʱ��), ����id_In, r_Bill.�շ�ϸĿid, ��ҽ��_In, Decode(��ҽ��id_In, 0, Null, ��ҽ��id_In), �ű�_In, 1, v_ԤԼ�Һ�,
             v_ԤԼ�Һ�);
        End If;
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) + 1, �����ѽ��� = Nvl(�����ѽ���, 0) + v_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) + v_ԤԼ�Һ�
        Where ID = �����¼id_In;
      
        --������ſ���
        Select Max(���), Max(ԤԼ˳���), Max(�Һ�״̬), Max(����Ա����)
        Into n_ԭ���, n_ԭԤԼ˳���, n_ԭ�Һ�״̬, v_ԭ����Ա
        From �ٴ�������ſ���
        Where ��¼id = n_ԭ�����¼id And (��� = r_Bill.���� Or ��ע = To_Char(r_Bill.����));
      
        If n_ԭ��� Is Not Null Then
          Select Count(1)
          Into n_Exists
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��� = n_ԭ��� And Nvl(ԤԼ˳���, 0) = Nvl(n_ԭԤԼ˳���, 0) And Nvl(�Һ�״̬, 0) = 0;
          If n_Exists = 1 Then
            Update �ٴ�������ſ���
            Set �Һ�״̬ = n_ԭ�Һ�״̬, ����Ա���� = v_ԭ����Ա
            Where ��¼id = �����¼id_In And ��� = n_ԭ��� And Nvl(ԤԼ˳���, 0) = Nvl(n_ԭԤԼ˳���, 0) And Nvl(�Һ�״̬, 0) = 0;
          Else
            Update ���˹Һż�¼ Set ���� = Null Where NO = r_Bill.No;
          End If;
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = Null
          Where ��¼id = n_ԭ�����¼id And ��� = n_ԭ��� And Nvl(ԤԼ˳���, 0) = Nvl(n_ԭԤԼ˳���, 0);
        End If;
      End If;
    
      ---���¹Һż�¼
      Update ������ü�¼
      Set ִ�в���id = ����id_In, ���˿���id = ����id_In, ���㵥λ = �ű�_In, ��ҩ���� = ����_In,
          --���˲���id = ����id_In,
          ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null
      Where ID = r_Bill.Id;
    
      --���²��˹Һż�¼
      If r_Bill.��� = 1 Then
        v_Temp := Zl_Identity(1);
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
        Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
        Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
        Begin
          Select ID Into n_ҽ��id From ��Ա�� Where ���� = ��ҽ��_In And Rownum < 2;
        Exception
          When Others Then
            n_ҽ��id := Null;
        End;
        Select ����䶯��¼_Id.Nextval Into n_�䶯id From Dual;
        Zl_����䶯��¼_Insert(r_Bill.No, 2, '���ﻻ��', v_����Ա����, v_����Ա���, �ű�_In, ����id_In, Null, n_ҽ��id, ��ҽ��_In, ����_In, n_����,
                         Null, n_�䶯id);
        v_�Һ����ɶ���     := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
        n_����̨ǩ���Ŷ�   := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113, 1, Nvl(����id_In, 0)));
        n_�ٴ�ǩ�������Ŷ� := Zl_To_Number(zl_GetSysParameter('�ٴ�ǩ���������Ŷ�', 1113));
        Select ID, �ű�, Nvl(����, 0)
        Into n_ҵ��id, v_�ű�, n_����
        From ���˹Һż�¼
        Where NO = r_Bill.No And Rownum = 1;
        If v_�Һ����ɶ��� <> 0 Then
          If Nvl(n_����̨ǩ���Ŷ�, 0) = 1 Then
            Select �������� Into v_�������� From �ŶӽкŶ��� Where ҵ��id = n_ҵ��id;
            If Nvl(v_��������, 0) <> 0 And Nvl(n_�ٴ�ǩ�������Ŷ�, 0) = 1 Then
              --ɾ��ԭ���ŶӼ�¼�����Ŷӣ���������_IN��ҵ��ID_IN
              Zl_�ŶӽкŶ���_Delete(v_��������, n_ҵ��id);
            Else
              Update �ŶӽкŶ��� Set �Ŷ�״̬ = 2 Where ҵ��id = n_ҵ��id And ҵ������ = 0;
            End If;
            Update ���˹Һż�¼ Set ��¼��־ = 0 Where ID = n_ҵ��id;
          Else
            v_�ֶ������� := ����id_In;
            --Zlgetnextqueue(ִ�в���id_In Number,ҵ��id_In     Number := Null)
            v_�ŶӺ��� := Zlgetnextqueue(����id_In, n_ҵ��id, v_�ű� || '|' || n_����);
            v_�Ŷ���� := Zlgetsequencenum(0, n_ҵ��id, 1);
            d_�Ŷ�ʱ�� := Zl_Get_Requeuedate(3, n_ҵ��id, ����id_In, ��ҽ��_In, ����_In);
            --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In , ҽ������_In
            Zl_�ŶӽкŶ���_Update(v_�ֶ�������, 0, n_ҵ��id, ����id_In, r_Bill.����, ����_In, ��ҽ��_In, v_�ŶӺ���, v_�Ŷ����, d_�Ŷ�ʱ��);
            --���ź���¶�����Ϣ���Ŷ�״̬Ҳ����Ϊ�Ŷ���
            Update �ŶӽкŶ��� Set �Ŷ�״̬ = 0 Where ҵ��id = n_ҵ��id And ҵ������ = 0;
          End If;
        End If;
        Update ���˹Һż�¼
        Set ִ�в���id = ����id_In, �ű� = �ű�_In, ���� = ����_In, ִ���� = ��ҽ��_In, ִ��״̬ = 0, ִ��ʱ�� = Null, �����¼id = �����¼id_In,
            ת��ű� = Null, ת�����id = Null, ת������ = Null, ת��ҽ�� = Null, ת��״̬ = Null
        Where NO = r_Bill.No;
      End If;
    End Loop;
  End If;
  b_Message.Zlhis_Regist_005(No_In, 2, n_�䶯id);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����;
/

--127075:����,2018-07-24,����ʹ�õ��˲��Ų�������̨ǩ���Ŷӵ�Oracle����
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
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;

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
                         Where ����id = p.Id And (d_����ʱ��_In Between ��Чʱ�� And ʧЧʱ��) And ���ʱ�� Is Not Null) And Not Exists
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
                               Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.���� = ����_In
                         Group By ����id) E
                  Where p.����id = c.Id And p.Id = b.�ƻ�id(+) And p.��Чʱ�� = e.��Ч And p.����id = e.����id And
                        Nvl(p.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And (d_����ʱ��_In Between p.��Чʱ�� And p.ʧЧʱ��) And
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
      Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
             Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
       Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
      Group By NO, ����id
      Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;
  
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
      Select Nvl(�Ƿ��ʱ��, 0), �޺���, �ѹ���, �����ѽ���, ��Լ��, �Ƿ���ſ���, ��Լ��, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���￪ʼʱ��, ������ֹʱ��
      Into n_���÷�ʱ��, n_�޺���, n_�ѹ���, n_�����ѽ���, n_��Լ��, n_��ſ���, n_��Լ��, n_��Ŀid, n_����id, n_ҽ��id, v_ҽ������, n_����ҽ��id, v_����ҽ������,
           d_���￪ʼʱ��, d_������ֹʱ��
      From �ٴ������¼
      Where ID = ��¼id_In And Nvl(�Ƿ�����, 0) = 0;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
  
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
              Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Rownum < 2
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
                                  Nvl(d.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')))
                     Union All
                     Select c.���, ����
                     From �ҺŰ��żƻ� A, �ҺŰ��� D, ������λ�ƻ����� C,
                          (Select Max(a.��Чʱ��) As ��Ч, ����id
                            From �ҺŰ��żƻ� A, �ҺŰ��� B
                            Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                                  ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                  Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.���� = ����_In
                            Group By ����id) E
                     Where a.����id = d.Id And a.���ʱ�� Is Not Null And d.���� = ����_In And a.����id = e.����id And
                           Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                           Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                  '7', '����', Null) = c.������Ŀ(+) And a.Id = c.�ƻ�id And c.������λ = ������λ_In And c.��� = n_��� And
                           ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                           Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
      
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
Update zlSystems Set �汾��='10.35.90.0022' Where ���=&n_System;
Commit;
