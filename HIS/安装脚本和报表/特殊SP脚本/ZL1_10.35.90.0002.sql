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
--122063:����,2018-03-12,��Һ�������Ŀ��Ը���ѡ��Ĳ�������ʾ��Ӧ����Һ��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ, Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 1, 0, 0, 0, 0, 44, '��ʾ��Դ����', '', '', '�ڲ�ѯ��Һ��ʱ��Ϊ��ѯ�����������Һ���еĲ�������Դ�����б���ʱ�������ѯ����',
         '��Դ���������б�������ID1������ID2����ʽ����', '', '��Һ������������ж�̨�������Էֱ�Բ�ͬ�������в����������ò�ͬ����Դ�����ķ�ʽ����ȡ��Ӧ�Ĳ�������Һ����', Null
  From Dual;


--122724:������,2018-03-12,��Ѫ��ֱ�ӷ�Ѫ��ʾҽ��վ
insert into ҵ����Ϣ����(����,����,˵��,��������) values ('ZLHIS_BLOOD_004','��Ѫ��ֱ�ӷ�Ѫ����','��Ѫ��ֱ�ӷ�Ѫ��ɣ�����ҽ��վ����ҽ�����',7);


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--122937:������,2018-03-16,����ӿڳ�����ӵĽ������
Create Or Replace Procedure Zl_Third_Getadviceinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���ȡҽ��������Ϣ/��ѯ
  --������
  --����� Xml_In
  --<IN>
  --     <YZID>1156789</YZID>--��ҽ��ID
  --</IN>

  --���� Xml_Out
  --<OUTPUT>
  --    <YZ>
  --       <PATIID></PATIID>     --����ҽ����¼.����ID
  --       <PAGEID></PAGEID>     --����ҽ����¼.��ҳID
  --       <BABY></BABY>   --����ҽ����¼.Ӥ��
  --       <YZID>1145878</YZID>   --����ҽ����¼.ҽ��ID�� ��ҽ��ID
  --       <RELATEDID></RELATEDID>   --����ҽ����¼.���ID
  --       <ZXKSID></ZXKSID>   --����ҽ����¼.ִ�п���id
  --       <YZQX>0</YZQX>      --����ҽ����¼.ҽ����Ч
  --       <STATE>8</STATE>    --����ҽ����¼.ҽ��״̬
  --       <JJBZ>0</JJBZ>      --����ҽ����¼.������־
  --       <KZYS>����</KZYS>   --����ҽ����¼.����ҽ��
  --       <KZSJ>2015-03-25 16:37:00</KZSJ>   --����ҽ����¼.����ʱ��
  --       <ZLXMID></ZLXMID>   --������ĿĿ¼.ID
  --       <ZLLB>E</ZLLB>      --������ĿĿ¼.���
  --       <ZLXMMC></ZLXMMC>   --������ĿĿ¼.���� ����飬����(������ C)������(�������� F)����Ѫ(K)����ҩ�䷽(������ E)������(����)
  --       <ZLXMCZLX></<ZLXMCZLX>   --������ĿĿ¼.��������
  --       <ZLXMZXFL></ZLXMZXFL>   --������ĿĿ¼.ִ�з���
  --       <BZ>21</BZ> ������ĿĿ¼.��������||������ĿĿ¼.ִ�з���
  --       <YF>������ע</YF>   --����ҽ����¼.ҽ������ ����ҽ�����е�  ҽ������
  --       <PC>BID</PC>   --����Ƶ����Ŀ.Ӣ������
  --       <ZXSJFY>18-20</ZXSJFY>   --����ҽ����¼.ִ��ʱ�䷽��
  --       <PLCS>2</PLCS>   --����ҽ����¼.Ƶ�ʴ���
  --       <PLJG>1</PLJG>   --����ҽ����¼.Ƶ�ʼ��
  --       <PSJG></PSJG>   --����ҽ����¼.Ƥ�Խ��
  --       <YSZT></YSZT>   --����ҽ����¼.ҽ������
  --       <KSZXSJ>2015-03-25 16:35:00</KSZXSJ>  --����ҽ����¼.��ʼִ��ʱ��
  --       <ZXZZSJ></ZXZZSJ>   --����ҽ����¼.ִ����ֹʱ��
  --       <TZYS></TZYS>   --����ҽ����¼.ͣ��ҽ��
  --       <TZSJ></TZSJ>   --����ҽ����¼.ͣ��ʱ��
  --       <DW>��</DW>   --������ĿĿ¼.���㵥λ
  --       <DL></DL>   --����ҽ����¼.��������
  --       <ZL></ZL>   --����ҽ����¼.�ܸ�����

  --       <ITEMLIST> ����Ѫ��Ŀ����/��ҩҽ����Ŀ��ϸ�����Ϣ����Ѫ��Ѫ����Ϣ��ҩƷ����ϸ��Ϣ
  --        <ITEM>
  --         <YSZT></YSZT>   --����ҽ����¼.ҽ������
  --         <YZID>1145878</YZID>   --����ҽ����¼.ҽ��ID
  --         <RELATEDID></RELATEDID>   --����ҽ����¼.���ID
  --         <ZLXMID></ZLXMID>   --������ĿĿ¼.ID
  --         <SFXMID></SFXMID>   --�շ���ĿĿ¼.id
  --         <SFXMMC></SFXMMC>   --�շ���ĿĿ¼.����
  --         <SFXMGG></SFXMGG>   --�շ���ĿĿ¼.���
  --         <BM></BM>           --�շ���Ŀ����.���ƣ���Ʒ����
  --         <ZL></ZL>           --����ҽ����¼.�ܸ�����
  --         <DL>10</DL>         --����ҽ����¼.��������
  --         <DW>ml</DW>         --�շ���ĿĿ¼.���㵥λ
  --         <ZLDW>ml</ZLDW>   --������ĿĿ¼.���㵥λ
  --         <ZXXZ></ZXXZ>   --����ҽ����¼.ִ������
  --         <ZXKS></ZXKS>   --������ĿĿ¼.ִ�п���
  --         <XDBH></XDBH>   --ѪҺ�շ���¼.Ѫ�����
  --         <SXXH></SXXH>   --ѪҺ�շ���¼.���
  --        </ITEM>
  --        <ITEM/>...
  --       </ITEMLIST>
  --      </YZ>
  --</OUTPUT>

  n_ҽ��id  ����ҽ����¼.Id%Type;
  x_ҽ��    Xmltype;
  x_Item    Xmltype;
  v_Xtmp    Clob; --��ʱXML
  n_Cnt     Number;
  x_Templet Xmltype;

  v_Ӣ����     ����Ƶ����Ŀ.Ӣ������%Type;
  v_�Թ�����   ��Ѫ������.����%Type;
  v_��Ӽ�     ��Ѫ������.��Ӽ�%Type;
  v_�Թܹ��   ��Ѫ������.���%Type;
  n_�Թ���ɫ   ��Ѫ������.��ɫ%Type;
  v_�շ���Ʒ�� �շ���Ŀ����.����%Type;
  n_����Ѫ��   Number; 
  v_SqlѪ��    Varchar2(4000);
  n_Ѫ������id Number(18);
  v_Tmp��Ѫ    Varchar2(4000);

  Type Bloodlist_Type Is Ref Cursor;
  Cbloodlist Bloodlist_Type;

  Type t_Code Is Record(
    ID       �շ���ĿĿ¼.Id%Type,
    ����     �շ���ĿĿ¼.����%Type,
    ���     �շ���ĿĿ¼.���%Type,
    ��λ     �շ���ĿĿ¼.���㵥λ%Type,
    Ѫ����� Varchar2(50),
    ���     Number(5));
  r_b t_Code;

Begin

  Select Extractvalue(Value(A), 'IN/YZID') Into n_ҽ��id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_Cnt := 0;
  For R In (Select a.����id, a.��ҳid, a.Ӥ��, a.Id As ҽ��id, a.���id, a.ִ�п���id, a.ҽ����Ч, a.ҽ��״̬, a.������־, a.����ҽ��, a.����ʱ��, a.������Ŀid,
                   a.�������, a.ҽ������, a.ִ��ʱ�䷽��, a.ִ��Ƶ��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.Ƥ�Խ��, a.ҽ������, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.ͣ��ҽ��, a.ͣ��ʱ��,
                   b.���� As ��Ŀ����, b.��������, b.ִ�з���, b.���㵥λ As ���Ƶ�λ, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.�շ�ϸĿid, c.���� As �շ�����,
                   c.���, Null As �շ���Ʒ��, c.���㵥λ As �շѵ�λ, a.ִ������, b.ִ�п���, b.�Թܱ���
            From ����ҽ����¼ A, ������ĿĿ¼ B, �շ���ĿĿ¼ C
            Where a.������Ŀid = b.Id And a.�շ�ϸĿid = c.Id(+) And (a.Id = n_ҽ��id Or a.���id = n_ҽ��id)
            Order By a.���) Loop
    n_Cnt := n_Cnt + 1;
    If n_Cnt = 1 Then
      Select Max(a.Ӣ������) Into v_Ӣ���� From ����Ƶ����Ŀ A Where a.���� = r.ִ��Ƶ��;
    End If;
    v_�Թ����� := Null;
    v_��Ӽ�   := Null;
    v_�Թܹ�� := Null;
    n_�Թ���ɫ := Null;
    If r.�Թܱ��� Is Not Null Then
      Select Max(a.����), Max(a.��Ӽ�), Max(a.���), Max(a.��ɫ)
      Into v_�Թ�����, v_��Ӽ�, v_�Թܹ��, n_�Թ���ɫ
      From ��Ѫ������ A
      Where a.���� = r.�Թܱ���;
    End If;
    --��ҽ��
    If r.���id Is Null Then
      v_Xtmp := '<YZ>';
      v_Xtmp := v_Xtmp || '<PATIID>' || r.����id || '</PATIID>'; --����ҽ����¼.����ID
      v_Xtmp := v_Xtmp || '<PAGEID>' || r.��ҳid || '</PAGEID>'; --����ҽ����¼.��ҳID
      v_Xtmp := v_Xtmp || '<BABY>' || r.Ӥ�� || '</BABY>'; --����ҽ����¼.Ӥ��
      v_Xtmp := v_Xtmp || '<YZID>' || r.ҽ��id || '</YZID>'; --����ҽ����¼.ҽ��ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.���id || '</RELATEDID>'; --����ҽ����¼.���ID
      v_Xtmp := v_Xtmp || '<ZXKSID>' || r.ִ�п���id || '</ZXKSID>'; --����ҽ����¼.ִ�п���id
      v_Xtmp := v_Xtmp || '<YZQX>' || r.ҽ����Ч || '</YZQX>'; --����ҽ����¼.ҽ����Ч
      v_Xtmp := v_Xtmp || '<STATE>' || r.ҽ��״̬ || '</STATE>'; --����ҽ����¼.ҽ��״̬
      v_Xtmp := v_Xtmp || '<JJBZ>' || r.������־ || '</JJBZ>'; --����ҽ����¼.������־
      v_Xtmp := v_Xtmp || '<KZYS>' || r.����ҽ�� || '</KZYS>'; --����ҽ����¼.����ҽ��
      v_Xtmp := v_Xtmp || '<KZSJ>' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</KZSJ>'; --����ҽ����¼.����ʱ��
      v_Xtmp := v_Xtmp || '<BZ>' || r.�������� || r.ִ�з��� || '</BZ>'; -- ������ĿĿ¼.��������||������ĿĿ¼.ִ�з���
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.������Ŀid || '</ZLXMID>'; --������ĿĿ¼.ID
      v_Xtmp := v_Xtmp || '<ZLLB>' || r.������� || '</ZLLB>'; --������ĿĿ¼.���
      v_Xtmp := v_Xtmp || '<YZNR>' || r.ҽ������ || '</YZNR>'; --ҽ������
      v_Xtmp := v_Xtmp || '<YF>' || r.��Ŀ���� || '</YF>'; --����ҽ����¼.ҽ������
      v_Xtmp := v_Xtmp || '<PC>' || v_Ӣ���� || '</PC>'; --����Ƶ����Ŀ.Ӣ������
      v_Xtmp := v_Xtmp || '<ZXSJFY>' || r.ִ��ʱ�䷽�� || '</ZXSJFY>'; --����ҽ����¼.ִ��ʱ�䷽��
      v_Xtmp := v_Xtmp || '<PLCS>' || r.Ƶ�ʴ��� || '</PLCS>'; --����ҽ����¼.Ƶ�ʴ���
      v_Xtmp := v_Xtmp || '<PLJG>' || r.Ƶ�ʼ�� || '</PLJG>'; --����ҽ����¼.Ƶ�ʼ��
      v_Xtmp := v_Xtmp || '<PSJG>' || r.Ƥ�Խ�� || '</PSJG>'; --����ҽ����¼.Ƥ�Խ��
      v_Xtmp := v_Xtmp || '<YSZT>' || r.ҽ������ || '</YSZT>'; --����ҽ����¼.ҽ������
      v_Xtmp := v_Xtmp || '<KSZXSJ>' || To_Char(r.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</KSZXSJ>'; --����ҽ����¼.��ʼִ��ʱ��
      v_Xtmp := v_Xtmp || '<ZXZZSJ>' || To_Char(r.ִ����ֹʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</ZXZZSJ>'; --����ҽ����¼.ִ����ֹʱ��
      v_Xtmp := v_Xtmp || '<TZYS>' || r.ͣ��ҽ�� || '</TZYS>'; --����ҽ����¼.ͣ��ҽ��
      v_Xtmp := v_Xtmp || '<TZSJ>' || To_Char(r.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</TZSJ>'; --����ҽ����¼.ͣ��ʱ��
      v_Xtmp := v_Xtmp || '<ZLXMMC>' || r.��Ŀ���� || '</ZLXMMC>'; --������ĿĿ¼.����
      v_Xtmp := v_Xtmp || '<ZLXMCZLX>' || r.�������� || '</ZLXMCZLX>'; --������ĿĿ¼.��������
      v_Xtmp := v_Xtmp || '<ZLXMZXFL>' || r.ִ�з��� || '</ZLXMZXFL>'; --������ĿĿ¼.ִ�з���
      --       (����Ѫ�ܷ���)
      v_Xtmp := v_Xtmp || '<CXGMC>' || v_�Թ����� || '</CXGMC>'; --��Ѫ������
      v_Xtmp := v_Xtmp || '<CXGTJJ>' || v_��Ӽ� || '</CXGTJJ>'; --��Ѫ����Ӽ�
      v_Xtmp := v_Xtmp || '<CXGGG>' || v_�Թܹ�� || '</CXGGG>'; --��Ѫ�ܹ��
      v_Xtmp := v_Xtmp || '<CXGYS>' || n_�Թ���ɫ || '</CXGYS>'; --��Ѫ����ɫ
      v_Xtmp := v_Xtmp || '<DW>' || r.���Ƶ�λ || '</DW>'; --������ĿĿ¼.���㵥λ
      v_Xtmp := v_Xtmp || '<DL>' || r.�������� || '</DL>'; --����ҽ����¼.��������
      v_Xtmp := v_Xtmp || '<ZL>' || r.�ܸ����� || '</ZL>'; --����ҽ����¼.�ܸ�����
      v_Xtmp := v_Xtmp || '</YZ>';
      x_ҽ�� := Xmltype(v_Xtmp);
    End If;
  
    --��Ѫ
    If r.������� = 'K' Then
      --�ж��Ƿ�װѪ��
      Select Zl_Checkobject(1, 'ѪҺ�շ���¼') Into n_����Ѫ�� From Dual;
      If n_����Ѫ�� > 0 Then
        n_Ѫ������id := r.ҽ��id;
        --ҽ������
        v_Xtmp    := '<YSZT>' || r.ҽ������ || '</YSZT>'; --����ҽ����¼.ҽ������
        v_Xtmp    := v_Xtmp || '<YZID>' || r.ҽ��id || '</YZID>'; --����ҽ����¼.ҽ��ID
        v_Xtmp    := v_Xtmp || '<RELATEDID>' || r.���id || '</RELATEDID>'; --����ҽ����¼.���ID
        v_Xtmp    := v_Xtmp || '<ZLXMID>' || r.������Ŀid || '</ZLXMID>'; --������ĿĿ¼.ID
        v_Xtmp    := v_Xtmp || '<ZL>' || r.�ܸ����� || '</ZL>'; --����ҽ����¼.�ܸ�����
        v_Xtmp    := v_Xtmp || '<DL>' || r.�������� || '</DL>'; --����ҽ����¼.��������
        v_Xtmp    := v_Xtmp || '<ZLDW>' || r.���Ƶ�λ || '</ZLDW>'; --������ĿĿ¼.���㵥λ
        v_Xtmp    := v_Xtmp || '<ZXXZ>' || r.ִ������ || '</ZXXZ>'; --����ҽ����¼.ִ������
        v_Xtmp    := v_Xtmp || '<ZXKS>' || r.ִ�п��� || '</ZXKS>'; --������ĿĿ¼.ִ�п���
        v_Tmp��Ѫ := v_Xtmp;
        If r.��鷽�� = '1' Then
          v_SqlѪ�� := 'Select d.Id,d.����,d.���,d.���㵥λ as ��λ, a.Ѫ�����,a.���
                       From ѪҺ�շ���¼ a,ѪҺ���ͼ�¼ b,ѪҺ��Ѫ��¼ c,�շ���ĿĿ¼ d
                       Where a.Id = b.�շ�id And b.�䷢id = c.Id and a.ѪҺid =d.id  And c.����id =:1';
        End If;
      End If;
    Elsif r.���id Is Not Null And r.������� = 'E' And r.�������� = '8' And Nvl(r.ִ�з���, 0) = 0 And n_����Ѫ�� = 1 And
          v_SqlѪ�� Is Null Then
      v_SqlѪ�� := 'Select b.Id,b.����,  b.���,b.���㵥λ as ��λ, a.Ѫ�����,a.���
                  From ѪҺ�շ���¼ a,�շ���ĿĿ¼ b
                  Where a.ѪҺid =b.id and a.�䷢id = (Select Id From ѪҺ��Ѫ��¼ Where ����id=:1)';
    Else
      v_SqlѪ�� := Null;
    End If;
  
    If v_SqlѪ�� Is Not Null And n_Ѫ������id Is Not Null Then
      --��Ѫҽ����ֻ�з�ҽ����ſ�����Ѫ����Ϣ
      x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      Open Cbloodlist For v_SqlѪ��
        Using n_Ѫ������id;
      Loop
        Fetch Cbloodlist
          Into r_b.Id, r_b.����, r_b.���, r_b.��λ, r_b.Ѫ�����, r_b.���;
        Exit When Cbloodlist%NotFound;
        v_�շ���Ʒ�� := Null;
        If r_b.Id Is Not Null Then
          For Z In (Select a.����, a.����
                    From �շ���Ŀ���� A
                    Where a.�շ�ϸĿid = r_b.Id
                    Group By a.����, a.����
                    Order By a.����) Loop
            v_�շ���Ʒ�� := z.����;
            If z.���� = 3 Then
              v_�շ���Ʒ�� := z.����;
              Exit;
            End If;
          End Loop;
        End If;
      
        v_Xtmp := '<ITEM jsonArray="True" >';
      
        v_Xtmp := v_Xtmp || v_Tmp��Ѫ;
      
        --Ѫ�ⲿ��
        v_Xtmp := v_Xtmp || '<SFXMID>' || r_b.Id || '</SFXMID>'; --�շ���ĿĿ¼.id
        v_Xtmp := v_Xtmp || '<SFXMMC>' || r_b.���� || '</SFXMMC>'; --�շ���ĿĿ¼.����
        v_Xtmp := v_Xtmp || '<SFXMGG>' || r_b.��� || '</SFXMGG>'; --�շ���ĿĿ¼.���
        v_Xtmp := v_Xtmp || '<BM>' || v_�շ���Ʒ�� || '</BM>'; --�շ���Ŀ����.���ƣ���Ʒ����
        v_Xtmp := v_Xtmp || '<DW>' || r_b.��λ || '</DW>'; --�շ���ĿĿ¼.���㵥λ
        v_Xtmp := v_Xtmp || '<XDBH>' || r_b.Ѫ����� || '</XDBH>'; --ѪҺ�շ���¼.Ѫ�����
        v_Xtmp := v_Xtmp || '<SXXH>' || r_b.��� || '</SXXH>'; --ѪҺ�շ���¼.���
      
        v_Xtmp := v_Xtmp || '</ITEM>';
        Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
      End Loop;
      Close Cbloodlist;
    End If;
  
    --��ҩ��ҩҽ��
    If r.������� = '5' Or r.������� = '6' Then
      --��/�� ҩ
      If x_Item Is Null Then
        --ֻ��ʼ��һ��
        x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
      End If;
      v_�շ���Ʒ�� := Null;
      If r.�շ�ϸĿid Is Not Null Then
        For Z In (Select a.����, a.����
                  From �շ���Ŀ���� A
                  Where a.�շ�ϸĿid = r.�շ�ϸĿid
                  Group By a.����, a.����
                  Order By a.����) Loop
          v_�շ���Ʒ�� := z.����;
          If z.���� = 3 Then
            v_�շ���Ʒ�� := z.����;
            Exit;
          End If;
        End Loop;
      End If;
    
      v_Xtmp := '<ITEM jsonArray="True" >';
      v_Xtmp := v_Xtmp || '<YSZT>' || r.ҽ������ || '</YSZT>'; --����ҽ����¼.ҽ������
      v_Xtmp := v_Xtmp || '<YZID>' || r.ҽ��id || '</YZID>'; --����ҽ����¼.ҽ��ID
      v_Xtmp := v_Xtmp || '<RELATEDID>' || r.���id || '</RELATEDID>'; --����ҽ����¼.���ID
      v_Xtmp := v_Xtmp || '<ZLXMID>' || r.������Ŀid || '</ZLXMID>'; --������ĿĿ¼.ID
      v_Xtmp := v_Xtmp || '<SFXMID>' || r.�շ�ϸĿid || '</SFXMID>'; --�շ���ĿĿ¼.id
      v_Xtmp := v_Xtmp || '<SFXMMC>' || r.�շ����� || '</SFXMMC>'; --�շ���ĿĿ¼.����
      v_Xtmp := v_Xtmp || '<SFXMGG>' || r.��� || '</SFXMGG>'; --�շ���ĿĿ¼.���
      v_Xtmp := v_Xtmp || '<BM>' || v_�շ���Ʒ�� || '</BM>'; --�շ���Ŀ����.���ƣ���Ʒ����
      v_Xtmp := v_Xtmp || '<ZL>' || r.�ܸ����� || '</ZL>'; --����ҽ����¼.�ܸ�����
      v_Xtmp := v_Xtmp || '<DL>' || r.�������� || '</DL>'; --����ҽ����¼.��������
      v_Xtmp := v_Xtmp || '<DW>' || r.�շѵ�λ || '</DW>'; --�շ���ĿĿ¼.���㵥λ
      v_Xtmp := v_Xtmp || '<ZLDW>' || r.���Ƶ�λ || '</ZLDW>'; --������ĿĿ¼.���㵥λ
      v_Xtmp := v_Xtmp || '<ZXXZ>' || r.ִ������ || '</ZXXZ>'; --����ҽ����¼.ִ������
      v_Xtmp := v_Xtmp || '<ZXKS>' || r.ִ�п��� || '</ZXKS>'; --������ĿĿ¼.ִ�п���
      v_Xtmp := v_Xtmp || '<XDBH></XDBH>'; --ѪҺ�շ���¼.Ѫ�����
      v_Xtmp := v_Xtmp || '<SXXH></SXXH>'; --ѪҺ�շ���¼.���
      v_Xtmp := v_Xtmp || '</ITEM>';
      Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
    End If;
  End Loop;
  If x_Item Is Not Null Then
    Select Appendchildxml(x_ҽ��, '/YZ', x_Item) Into x_ҽ�� From Dual;
  End If;
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Appendchildxml(x_Templet, '/OUTPUT', x_ҽ��) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getadviceinfo;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getpathway
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ������ٴ�·����,�׶���Ϣ/��ѯ
  --��Σ�Xml_In
  --<IN>
  --     <PATIID>29</PATIID>     --����ID
  --     <PAGEID>1</PAGEID>     --��ҳID
  --</IN>

  --���Σ�Xml_Out
  --<OUTPUT>
  --  <LJID>43</LJID>     --�����ٴ�·��.·��ID
  --  <YSLJID></YSLJID>   --����·������.ԭ·��id
  --  <BBH></BBH>         --�����ٴ�·��.�汾��
  --  <YSBBH></YSBBH>     --����·������.ԭ·���汾
  --  <LJMC>���Ե�������β���ٴ�·��</LJMC>   --�ٴ�·��Ŀ¼.����
  --  <ZXZT>ִ����</ZXZT>   --�����ٴ�·��.״̬
  --  <BZZYR>6</BZZYR>   --�ٴ�·���汾.��׼סԺ��
  --  <DQTS>5</DQTS>   --�����ٴ�·��.��ǰ����

  --  <PHASELIST>
  --   <PHASES>
  --    <PHASE>
  --      <JDID>1</JDID>   --����·��ִ��.�׶�ID
  --      <JD>סԺ��1�� (סԺ��,������)</JD>   --�ٴ�·���׶�.����
  --      <DQJD>0</DQJD>   --�����ٴ�·��.��ǰ�׶�ID
  --      <DAYS>
  --        <DAY>
  --          <TS>1</TS>                     --����·��ִ��.����
  --          <RQ>2011-09-16 00:00:00</RQ>   --����·��ִ��.����
  --          <PGJG>����</PGJG>              --����·������.�������
  --          <PGSM>�ֹ�</PGSM>              --����·������.����˵��
  --          <PGR>����</PGR>                --����·������.������
  --          <PGSJ>2011-09-16 10:51:40</PGSJ>   --����·������.����ʱ��
  --          <BYYY></BYYY>                      --����·������.����ԭ�򣨱��쳣��ԭ��.���ƣ�
  --          <ITEMLIST>
  --             <ITEM>
  --                <FL>��Ҫ���ƹ���</FL>   --�ٴ�·������.����
  --                <TBID />   --����·��ִ��.ͼ��ID
  --                <ZXID>3366</ZXID>   --����·��ִ��.ID
  --                <XMID>1</XMID>   --����·��ִ��.��ĿID    
  --                <XMXH>1</XMXH>   --�ٴ�·����Ŀ.��Ŀ��ţ�XMIDΪ��ʱ��ȡ����·��ִ�� .��Ŀ��ţ�   
  --                <XMNR>ѯ�ʲ�ʷ�����</XMNR>   --�ٴ�·����Ŀ.��Ŀ���ݣ�XMIDΪ��ʱ��ȡ����·��ִ�� .��Ŀ���ݣ�
  --                <ZXFS>1</ZXFS>   --�ٴ�·����Ŀ.ִ�з�ʽ
  --                <ZXJG>�Ѿ�ִ��</ZXJG>   --����·��ִ��.ִ�н��
  --                <TJYY />   --����·��ִ��.���ԭ��
  --                <ZXBYYY />  ����ԭ��
  --              </ITEM>
  --           </ITEMLIST>
  --        </DAY>
  --        <DAY/>...
  --     </DAYS>
  --    </PHASE>
  --    <PHASE/>...
  --  </PHASES>
  -- </PHASELIST>
  --</OUTPUT>

  n_����id     ����ҽ����¼.����id%Type;
  n_��ҳid     ����ҽ����¼.��ҳid%Type;
  n_·��id     �ٴ�·���׶�.·��id%Type;
  n_�汾��     �ٴ�·���׶�.�汾��%Type;
  n_��ǰ�׶�id �����ٴ�·��.��ǰ�׶�id%Type;
  n_·����¼id �����ٴ�·��.Id%Type;
  v_Xtmp       Clob; --��ʱXML
  x_Templet    Xmltype;
  x_Phase      Xmltype;
  x_Day        Xmltype;
  x_Item       Xmltype;
  v_������Ϣ   Varchar2(4000);
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;

  Cursor c_Main Is
    Select a.Id, a.·��id, c.ԭ·��id, a.�汾��, c.ԭ·���汾, e.���� As ·������,
           Decode(a.״̬, 0, '�����ϵ�������', 1, 'ִ����', 2, '��������', 3, '�������', Null) As ״̬, f.��׼סԺ��, a.��ǰ����, a.��ǰ�׶�id
    From �����ٴ�·�� A, ����·������ C, �ٴ�·��Ŀ¼ E, �ٴ�·���汾 F
    Where a.����id = n_����id And a.��ҳid = n_��ҳid And a.·��id = e.Id And a.��ǰ�׶�id = c.�׶�id(+) And a.Id = c.·����¼id(+) And
          a.��ǰ���� = c.����(+) And a.·��id = f.·��id And a.�汾�� = f.�汾��;

  Cursor c_����׶� Is
    Select a.Id As �׶�id, a.���� As �׶�����
    From �ٴ�·���׶� A
    Where a.·��id = n_·��id And a.�汾�� = n_�汾�� And a.��id Is Null
    Order By a.���;

  Type t_����׶� Is Table Of c_����׶�%RowType;
  r_����׶� t_����׶�;

  --�����ɵĽ׶�
  Cursor c_Phase Is
    Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, To_Char(a.����, 'day') ����, b.���� As �׶���, b.���, b.˵��, b.��id,
           Decode(g.·��id, b.·��id, 1, 0) As ����
    From (Select a.�׶�id, a.����, a.����, a.·����¼id
           From ����·��ִ�� A
           Where a.·����¼id = n_·����¼id
           Group By a.�׶�id, a.����, a.����, a.·����¼id) A, �ٴ�·���׶� B, �ٴ�·���׶� C, �����ٴ�·�� G
    Where a.�׶�id = b.Id And b.��id = c.Id(+) And g.Id = a.·����¼id
    Order By ����, ����, Nvl(c.���, b.���);

  Type t_Phase Is Table Of c_Phase%RowType;
  r_Phase t_Phase;

  --��ϸ��Ŀ
  Cursor c_Item Is
    Select a.Id, Nvl(b.ͼ��id, a.ͼ��id) As ͼ��id, a.����, To_Char(a.����, 'yyyy-mm-dd') As ����, a.����, a.�׶�id,
           Nvl(a.��Ŀ���, b.��Ŀ���) As ��Ŀ���, Nvl(b.��Ŀ����, a.��Ŀ����) ��Ŀ����, a.��Ŀid, Decode(a.ִ����, Null, 0, 1) ִ��״̬,
           Nvl(b.ִ�з�ʽ, 1) ִ�з�ʽ, a.���ԭ��, Nvl(a.����ʱ������, 0) As ����ʱ������, c.���� As ����ԭ��, Nvl(b.��Ŀ���, a.��Ŀ���) As ��Ŀ���, a.ִ�н��,
           d.·��id, d.��֧id, Nvl(Nvl(a.������, b.������), 1) As ������, d.���� As �׶���
    From ����·��ִ�� A, �ٴ�·����Ŀ B, ���쳣��ԭ�� C, �ٴ�·���׶� D
    Where a.·����¼id = n_·����¼id And a.��Ŀid = b.Id(+) And a.����ԭ�� = c.����(+) And a.�׶�id + 0 = d.Id
    Order By a.����, ����, ��Ŀ���;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;

  --�׶�����
  Cursor c_Eval Is
    Select a.�׶�id, a.����, Decode(a.�������, 1, '����', -1, '����', Null) As �������, a.����˵��, a.������, a.����ʱ��, c.���� As ����ԭ��, a.���������,
           Nvl(a.ʱ�����, 0) ʱ�����, a.��ת�����, a.ԭ·��id
    From ����·������ A, ����·������ B, ���쳣��ԭ�� C
    Where a.·����¼id = b.·����¼id(+) And a.�׶�id = b.�׶�id(+) And a.���� = b.����(+) And a.·����¼id = n_·����¼id And b.����ԭ�� = c.����(+);
  Type t_Eval Is Table Of c_Eval%RowType;
  r_Eval t_Eval;
Begin
  Select Extractvalue(Value(A), 'IN/PATIID') As ����id, Extractvalue(Value(A), 'IN/PAGEID') As ��ҳid
  Into n_����id, n_��ҳid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  For R In c_Main Loop
    v_Xtmp := '<LJID>' || r.·��id || '</LJID>'; --�����ٴ�·��.·��ID
    v_Xtmp := v_Xtmp || '<YSLJID>' || r.ԭ·��id || '</YSLJID>'; --����·������.ԭ·��id
    v_Xtmp := v_Xtmp || '<BBH>' || r.�汾�� || '</BBH>'; --�����ٴ�·��.�汾��
    v_Xtmp := v_Xtmp || '<YSBBH>' || r.ԭ·���汾 || '</YSBBH>'; --����·������.ԭ·���汾
    v_Xtmp := v_Xtmp || '<LJMC>' || r.·������ || '</LJMC>'; --�ٴ�·��Ŀ¼.����
    v_Xtmp := v_Xtmp || '<ZXZT>' || r.״̬ || '</ZXZT>'; --�����ٴ�·��.״̬
    v_Xtmp := v_Xtmp || '<BZZYR>' || r.��׼סԺ�� || '</BZZYR>'; --�ٴ�·���汾.��׼סԺ��
    v_Xtmp := v_Xtmp || '<DQTS>' || r.��ǰ���� || '</DQTS>'; --�����ٴ�·��.��ǰ����;
  
    n_·��id     := r.·��id;
    n_�汾��     := r.�汾��;
    n_��ǰ�׶�id := r.��ǰ�׶�id;
    n_·����¼id := r.Id;
  End Loop;
  x_Templet := Xmltype('<OUTPUT>' || v_Xtmp || '<PHASELIST></PHASELIST></OUTPUT>');

  If n_·����¼id Is Null Then
    v_Err_Msg := 'δ�ҵ�·����Ϣ��';
    Raise Err_Item;
  End If;

  Open c_����׶�;
  Fetch c_����׶� Bulk Collect
    Into r_����׶�;
  Close c_����׶�;

  Open c_Phase;
  Fetch c_Phase Bulk Collect
    Into r_Phase;
  Close c_Phase;

  Open c_Item;
  Fetch c_Item Bulk Collect
    Into r_Item;
  Close c_Item;

  Open c_Eval;
  Fetch c_Eval Bulk Collect
    Into r_Eval;
  Close c_Eval;

  For I In 1 .. r_����׶�.Count Loop
    v_Xtmp  := '<PHASE jsonArray="True" ><JDID>' || r_����׶�(I).�׶�id || '</JDID><JD>' || r_����׶�(I).�׶����� || '</JD><DQJD>' || n_��ǰ�׶�id ||
               '</DQJD><DAYS></DAYS></PHASE>';
    x_Phase := Xmltype(v_Xtmp);
  
    For J In 1 .. r_Phase.Count Loop
      If r_����׶�(I).�׶�id = r_Phase(J).�׶�id Then
        --day 
        v_Xtmp     := '<DAY jsonArray="True" >';
        v_Xtmp     := v_Xtmp || '<TS>' || r_Phase(J).���� || '</TS>'; --����·��ִ��.����
        v_Xtmp     := v_Xtmp || '<RQ>' || r_Phase(J).���� || ' 00:00:00</RQ>'; --����·��ִ��.����      
        v_������Ϣ := '<PGJG></PGJG><PGSM></PGSM><PGR></PGR><PGSJ></PGSJ><BYYY></BYYY>';
        For K In 1 .. r_Eval.Count Loop
          If r_Phase(J).�׶�id = r_Eval(K).�׶�id And r_Phase(J).���� = r_Eval(K).���� Then
            v_������Ϣ := '<PGJG>' || r_Eval(K).������� || '</PGJG>'; --����·������.�������
            v_������Ϣ := v_������Ϣ || '<PGSM>' || r_Eval(K).����˵�� || '</PGSM>'; --����·������.����˵��
            v_������Ϣ := v_������Ϣ || '<PGR>' || r_Eval(K).������ || '</PGR>'; --����·������.������
            v_������Ϣ := v_������Ϣ || '<PGSJ>' || To_Char(r_Eval(K).����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</PGSJ>'; --����·������.����ʱ��
            v_������Ϣ := v_������Ϣ || '<BYYY>' || r_Eval(K).����ԭ�� || '</BYYY>'; --����·������.����ԭ�򣨱��쳣��ԭ��.���ƣ�         
          End If;
        End Loop;
        v_Xtmp := v_Xtmp || v_������Ϣ;
        v_Xtmp := v_Xtmp || '</DAY>';
        x_Day  := Xmltype(v_Xtmp);
      
        --item
        x_Item := Xmltype('<ITEMLIST></ITEMLIST>');
        For K In 1 .. r_Item.Count Loop
          If r_Phase(J).�׶�id = r_Item(K).�׶�id And r_Phase(J).���� = r_Item(K).���� Then
            v_Xtmp := '<ITEM jsonArray="True" >';
            v_Xtmp := v_Xtmp || '<FL>' || r_Item(K).���� || '</FL>'; --�ٴ�·������.����
            v_Xtmp := v_Xtmp || '<TBID>' || r_Item(K).ͼ��id || '</TBID>'; --����·��ִ��.ͼ��ID
            v_Xtmp := v_Xtmp || '<ZXID>' || r_Item(K).Id || '</ZXID>'; --����·��ִ��.ID
            v_Xtmp := v_Xtmp || '<XMID>' || r_Item(K).��Ŀid || '</XMID>'; --����·��ִ��.��ĿID    
            v_Xtmp := v_Xtmp || '<XMXH>' || r_Item(K).��Ŀ��� || '</XMXH>'; --�ٴ�·����Ŀ.��Ŀ��ţ�XMIDΪ��ʱ��ȡ����·��ִ�� .��Ŀ��ţ�   
            v_Xtmp := v_Xtmp || '<XMNR>' || r_Item(K).��Ŀ���� || '</XMNR>'; --�ٴ�·����Ŀ.��Ŀ���ݣ�XMIDΪ��ʱ��ȡ����·��ִ�� .��Ŀ���ݣ�
            v_Xtmp := v_Xtmp || '<ZXFS>' || r_Item(K).ִ�з�ʽ || '</ZXFS>'; --�ٴ�·����Ŀ.ִ�з�ʽ
            v_Xtmp := v_Xtmp || '<ZXJG>' || r_Item(K).ִ�н�� || '</ZXJG>'; --����·��ִ��.ִ�н��
            v_Xtmp := v_Xtmp || '<TJYY>' || r_Item(K).���ԭ�� || '</TJYY>'; --����·��ִ��.���ԭ��
            v_Xtmp := v_Xtmp || '<ZXBYYY>' || r_Item(K).����ԭ�� || '</ZXBYYY>'; --����ԭ��
            v_Xtmp := v_Xtmp || '</ITEM>';
            Select Appendchildxml(x_Item, '/ITEMLIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
          End If;
        End Loop;
        Select Appendchildxml(x_Day, '/DAY', x_Item) Into x_Day From Dual;
        Select Appendchildxml(x_Phase, '/PHASE/DAYS', x_Day) Into x_Phase From Dual;
      End If;
    End Loop;
    Select Appendchildxml(x_Templet, '/OUTPUT/PHASELIST', x_Phase) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpathway;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getpathwaydetail
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ��ٴ�·����ĳ����Ŀ�����ִ����ϸ��ҽ����Ϣ/��ѯ
  --��Σ�Xml_In
  --<IN>
  --     <ZXID>29</ZXID>     --����·��ִ��.ID
  --</IN>
  --���Σ�Xml_Out
  --<OUTPUT>
  -- <LJEXEC>
  --  <ZXQK>
  --   <ZXZ />   --����·��ִ��.ִ����
  --   <ZXR>����</ZXR>   --����·��ִ��.ִ����
  --   <ZXSJ>2011-10-24 17:28:53</ZXSJ>   --����·��ִ��.ִ��ʱ��
  --   <ZXJG>�Ѿ�ִ��</ZXJG>   --����·��ִ��.ִ�н��
  --   <ZXSM />   --����·��ִ��..ִ��˵��
  --  </ZXQK>
  --  <YZLIST>
  --   <YZXX>
  --    <YZQX>0</YZQX>   --����ҽ����¼.ҽ����Ч
  --    <YZNR>ע���ÿ���ù�� 0.6g/֧ ���ݵ�Ҽ��ҩ���޹�˾</YZNR>   --����ҽ����¼.ҽ������
  --    <DL>ÿ��1.2g</DL>   -����ҽ����¼.��������
  --    <ZL />   --����ҽ����¼.�ܸ�����
  --    <GYTJ>������ע�����</GYTJ>   --������ĿĿ¼.����
  --    <ZXPL>ÿ�����</ZXPL>   --����ҽ����¼.ִ��Ƶ��
  --    <ZXSJ>10-16</ZXSJ>   -����ҽ����¼.ִ��ʱ�䷽��
  --    <YSZT />   --����ҽ����¼..ҽ������
  --   </YZXX>
  --  </YZLIST>
  -- </LJEXEC>
  --</OUTPUT>

  n_ִ��id  ����·��ִ��.Id%Type;
  v_Xtmp    Clob; --��ʱXML
  x_Templet Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/ZXID') Into n_ִ��id From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  For R In (Select a.ִ����, a.ִ����, a.ִ��ʱ��, a.ִ�н��, a.ִ��˵�� From ����·��ִ�� A Where a.Id = n_ִ��id) Loop
    v_Xtmp := '<ZXQK>';
    v_Xtmp := v_Xtmp || '<ZXZ>' || r.ִ���� || '</ZXZ>';
    v_Xtmp := v_Xtmp || '<ZXR>' || r.ִ���� || '</ZXR>';
    v_Xtmp := v_Xtmp || '<ZXSJ>' || To_Char(r.ִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</ZXSJ>';
    v_Xtmp := v_Xtmp || '<ZXJG>' || r.ִ�н�� || '</ZXJG>';
    v_Xtmp := v_Xtmp || '<ZXSM>' || r.ִ��˵�� || '</ZXSM>';
    v_Xtmp := v_Xtmp || '</ZXQK>';
  End Loop;

  x_Templet := Xmltype('<OUTPUT><LJEXEC>' || v_Xtmp || '</LJEXEC><YZLIST></YZLIST></OUTPUT>');

  --(��/��ҩ������ҩƷ�У�����ҽ��������ҽ����)
  For R In (Select a.ҽ����Ч, a.ҽ������, a.����, a.����, a.��ҩ;��, a.ִ��Ƶ��, a.ʱ�䷽��, a.ҽ������
            From (Select a.Id, a.���id, a.�������, d.������� As �����, a.ҽ����Ч, a.ҽ������, a.�������� As ����, a.�ܸ����� As ����, e.���� As ��ҩ;��,
                          a.ִ��Ƶ��, a.ִ��ʱ�䷽�� As ʱ�䷽��, a.ҽ������, c.��������
                   From ����ҽ����¼ A, ����·��ҽ�� B, ������ĿĿ¼ C, ����ҽ����¼ D, ������ĿĿ¼ E
                   Where b.·��ִ��id = n_ִ��id And a.Id = b.����ҽ��id And a.������Ŀid = c.Id(+) And a.���id = d.Id(+) And
                         d.������Ŀid = e.Id(+)) A
            Where a.���id Is Null And Not (a.������� = 'E' And a.�������� = '2')) Loop
    v_Xtmp := '<YZXX jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<YZQX>' || r.ҽ����Ч || '</YZQX>';
    v_Xtmp := v_Xtmp || '<YZNR>' || r.ҽ������ || '</YZNR>';
    v_Xtmp := v_Xtmp || '<DL>' || r.���� || '</DL>';
    v_Xtmp := v_Xtmp || '<ZL>' || r.���� || '</ZL>';
    v_Xtmp := v_Xtmp || '<GYTJ>' || r.��ҩ;�� || '</GYTJ>';
    v_Xtmp := v_Xtmp || '<ZXPL>' || r.ִ��Ƶ�� || '</ZXPL>';
    v_Xtmp := v_Xtmp || '<ZXSJ>' || r.ʱ�䷽�� || '</ZXSJ>';
    v_Xtmp := v_Xtmp || '<YSZT>' || r.ҽ������ || '</YSZT>';
    v_Xtmp := v_Xtmp || '</YZXX>';
    Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST', Xmltype(v_Xtmp)) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpathwaydetail;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getdiagnosis
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���ȡ���������Ϣ/��ѯ
  --��Σ�Xml_In
  --<IN>
  --     <PATIID></PATIID>         --����ID
  --     <PAGEID></PAGEID>     --��ҳID
  --</IN>
  --���Σ�Xml_Out
  --<OUTPUT>
  --  <ZDLIST>
  --    <ZD>
  --      <ZDLX></ZDLX> --������͡����͵����ƣ�������ϡ���Ժ��ϡ���Ժ��ϵ�
  --      <ZDCX></ZDCX> --��ϴ���
  --      <ZDBM></ZDBM> --��ϱ���
  --      <ZDMC></ZDMC> --�������
  --    </ZD>
  --  </ZDLIST>
  --</OUTPUT>

  n_����id   ����ҽ����¼.����id%Type;
  n_��ҳid   ����ҽ����¼.��ҳid%Type;
  v_Xtmp     Varchar(5000); --��ʱXML
  v_Tmp      Varchar2(800);
  v_��ϱ��� Varchar2(1000);
  v_������� Varchar2(1000);
  x_Templet  Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/PATIID'), Extractvalue(Value(A), 'IN/PAGEID')
  Into n_����id, n_��ҳid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><ZDLIST></ZDLIST></OUTPUT>');

  For R In (Select Decode(a.�������, 1, '��ҽ�������', 2, '��ҽ��Ժ���', 3, '��ҽ��Ժ���', 5, 'Ժ�ڸ�Ⱦ', 6, '�������', 7, '�����ж���', 8, '��ǰ���', 9,
                           '�������', 10, '����֢', 11, '��ҽ�������', 12, '��ҽ��Ժ���', 13, '��ҽ��Ժ���', 21, '��ԭѧ���', 22, 'Ӱ��ѧ���') As �������,
                   a.��ϴ���, a.�������
            From ������ϼ�¼ A
            Where a.����id = n_����id And a.��ҳid = n_��ҳid
            Order By a.�������, a.��ϴ���) Loop
  
    v_��ϱ��� := Null;
    v_������� := r.�������;
    v_Tmp      := r.�������;
    If Substr(v_Tmp, 1, 1) = '(' Then
      v_��ϱ��� := Substr(v_Tmp, 2, Instr(v_Tmp, ')') - 2);
      v_������� := Substr(v_Tmp, Instr(v_Tmp, ')') + 1);
    End If;
  
    v_Xtmp := '<ZD jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<ZDLX>' || r.������� || '</ZDLX>';
    v_Xtmp := v_Xtmp || '<ZDCX>' || r.��ϴ��� || '</ZDCX>';
    v_Xtmp := v_Xtmp || '<ZDBM>' || v_��ϱ��� || '</ZDBM>';
    v_Xtmp := v_Xtmp || '<ZDMC>' || v_������� || '</ZDMC>';
    v_Xtmp := v_Xtmp || '</ZD>';
    
    Select Appendchildxml(x_Templet, '/OUTPUT/ZDLIST', Xmltype(v_Xtmp)) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdiagnosis;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getallergy
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���ȡ���˹�����¼/��ѯ
  --��Σ�Xml_In
  --<IN>
  --     <PATIID></PATIID>         --����ID
  --     <PAGEID></PAGEID>     --��ҳID
  --</IN>
  --���Σ�Xml_Out
  --<OUTPUT>
  --  <GMLIST>
  --    <GM>
  --      <GMYW></GMYW> --����ҩ��
  --      <GMSJ></GMSJ> --����ʱ��
  --      <JLSJ></JLSJ> --��¼ʱ��
  --      <JLR></JLR> --��¼��
  --    </GM>
  --  </GMLIST>
  --</OUTPUT>
  n_����id  ����ҽ����¼.����id%Type;
  n_��ҳid  ����ҽ����¼.��ҳid%Type;
  v_Xtmp    Varchar(5000); --��ʱXML
  x_Templet Xmltype;
Begin
  Select Extractvalue(Value(A), 'IN/PATIID'), Extractvalue(Value(A), 'IN/PAGEID')
  Into n_����id, n_��ҳid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><GMLIST></GMLIST></OUTPUT>');
  For R In (Select a.ҩ����, a.����ʱ��, a.��¼ʱ��, a.��¼��
            From ���˹�����¼ A
            Where a.����id = n_����id And a.��ҳid = n_��ҳid
            Order By a.��¼ʱ��) Loop
    v_Xtmp := '<GM jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<GMYW>' || r.ҩ���� || '</GMYW>'; --����ҩ��
    v_Xtmp := v_Xtmp || '<GMSJ>' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</GMSJ>'; --����ʱ��
    v_Xtmp := v_Xtmp || '<JLSJ>' || To_Char(r.��¼ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '</JLSJ>'; --��¼ʱ��
    v_Xtmp := v_Xtmp || '<JLR>' || r.��¼�� || '</JLR>'; --��¼��
    v_Xtmp := v_Xtmp || '</GM>';

    Select Appendchildxml(x_Templet, '/OUTPUT/GMLIST', Xmltype(v_Xtmp)) Into x_Templet From Dual;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getallergy;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
Create Or Replace Procedure Zl_Third_Getpatichange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���ȡ���˱䶯��¼/��ѯ
  --��Σ�Xml_In
  --<IN>
  --     <PATIID></PATIID>         --����ID
  --     <PAGEID></PAGEID>     --��ҳID
  --</IN>
  --���Σ�Xml_Out
  --<OUTPUT>
  --  <BDLIST>--�䶯����
  --    <ITEM>
  --      <SJ></SJ>--�䶯ʱ��
  --      <LSXLIST>--�䶯�����б�
  --         <ITEM>
  --           <MC></MC>--�䶯�������ƣ���������ȼ�/����סԺҽʦ/������ʿ/����תסԺ/��������ҽʦ/��������/����ҽ��С��/��������/��Ժ/��ס/����/������λ�ȼ�/Ԥ��Ժ/��������ҽʦ��
  --           <XXLIST>--�䶯��Ϣ�б�
  --              <ITEM>
  --                <XXM></XXM>  ----��Ϣ���ƣ�ҽ��С��/����/����/����/��λ�ȼ�/����ȼ�/��ʿ/סԺҽʦ/����ҽ��/����ҽ��/��ǰ����/���������         
  --                <YXX></YXX>  ----ԭ��Ϣֵ        
  --                <XXX></XXX>  ----����Ϣֵ             
  --              </ITEM> 
  --              ...
  --           </XXLIST>
  --         </ITEM>
  --         ...
  --      </LSXLIST>  
  --    </ITEM>
  --    ... 
  --  </BDLIST>
  --</OUTPUT>

  n_����id ����ҽ����¼.����id%Type;
  n_��ҳid ����ҽ����¼.��ҳid%Type;

  v_Tmp  Varchar(5000);
  v_Tmp1 Varchar(5000);

  v_Preʱ��  Varchar(500);
  v_Curʱ��  Varchar(500);
  v_�䶯���� Varchar(500);

  n_Preidx Number;
  n_Curidx Number;

  v_Value    Varchar(500);
  x_Templet  Xmltype;
  x_�䶯ʱ�� Xmltype;
  x_�䶯���� Xmltype;

  Cursor c_Pati Is
    Select a.Id, f.���� As ҽ��С����, b.���� As ����, c.���� As ����, a.���Ӵ�λ, Decode(a.���Ӵ�λ, 0, '����', '����') As ��λ����, a.����,
           d.���� As ��λ�ȼ�, e.���� As ����ȼ�, a.���λ�ʿ As ��ʿ, a.����ҽʦ As סԺҽʦ, a.����ҽʦ As ����ҽ��, a.����ҽʦ As ����ҽ��, a.���� As ��ǰ����,
           a.����Ա���� As ��ʼ����Ա,
           Decode(a.��ʼԭ��, 1, '��Ժ', 2, '��ס', 3, 'ת��', 4, '����', 5, '������λ�ȼ�', 6, '��������ȼ�', 7, '����סԺҽʦ', 8, '������ʿ', 9,
                   '����תסԺ', 10, 'Ԥ��Ժ', 11, '��������ҽʦ', 12, '��������ҽʦ', 13, '��������', 14, '����ҽ��С��', 15, '��������') As ��ʼԭ��,
           To_Char(a.��ʼʱ��, 'YYYY-MM-DD HH24:MI:SS') As ��ʼʱ��, a.��ֹ��Ա As ��ֹ����Ա,
           Decode(a.��ֹԭ��, 1, '��Ժ', 2, '��ס', 3, 'ת��', 4, '����', 5, '������λ�ȼ�', 6, '��������ȼ�', 7, '����סԺҽʦ', 8, '������ʿ', 9,
                   '����תסԺ', 10, 'Ԥ��Ժ', 11, '��������ҽʦ', 12, '��������ҽʦ', 13, '��������', 14, '����ҽ��С��', 15, '��������') As ��ֹԭ��,
           To_Char(a.��ֹʱ��, 'YYYY-MM-DD HH24:MI:SS') As ��ֹʱ��
    From ���˱䶯��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ D, �շ���ĿĿ¼ E, �ٴ�ҽ��С�� F
    Where a.����id = b.Id And a.����id = c.Id And a.��λ�ȼ�id = d.Id(+) And a.����ȼ�id = e.Id(+) And a.����id = n_����id And
          a.��ҳid = n_��ҳid And a.��ʼʱ�� Is Not Null And a.ҽ��С��id = f.Id(+)
    Order By a.��ֹʱ��, a.��ʼʱ��, a.���Ӵ�λ, a.����;

  Type t_Pati Is Table Of c_Pati%RowType;
  r_Pati t_Pati;
  r_Seek t_Pati;

Begin
  Select Extractvalue(Value(A), 'IN/PATIID') As ����id, Extractvalue(Value(A), 'IN/PAGEID') As ��ҳid
  Into n_����id, n_��ҳid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  Open c_Pati;
  Fetch c_Pati Bulk Collect
    Into r_Pati;
  Close c_Pati;
  r_Seek := r_Pati;

  v_Preʱ�� := '0';
  n_Preidx  := 0;
  n_Curidx  := 0;

  x_Templet := Xmltype('<OUTPUT><BDLIST></BDLIST></OUTPUT>');

  For I In 1 .. r_Pati.Count Loop
    --��ʱ���Ϊһ�α䶯
    If v_Preʱ�� <> r_Pati(I).��ʼʱ�� And r_Pati(I).���Ӵ�λ = 0 Then
      v_Curʱ�� := r_Pati(I).��ʼʱ��;
      --���м����䶯���
      x_�䶯ʱ�� := Xmltype('<ITEM jsonArray="True" ><SJ>' || v_Curʱ�� || '</SJ><LSXLIST></LSXLIST></ITEM>');
    
      v_�䶯���� := '0';
      For J In 1 .. r_Seek.Count Loop
        If r_Seek(J).��ʼʱ�� = v_Curʱ�� And r_Seek(J).���Ӵ�λ = 0 Then
          If v_�䶯���� <> r_Seek(J).��ʼԭ�� Then
            v_�䶯���� := r_Seek(J).��ʼԭ��;
            n_Curidx   := J;
            x_�䶯���� := Xmltype('<ITEM jsonArray="True" ><MC>' || v_�䶯���� || '</MC><XXLIST></XXLIST></ITEM>');
          
            --ҽ��С����
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).ҽ��С����, 'XXX') <> Nvl(r_Seek(n_Curidx).ҽ��С����, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>ҽ��С����</XXM><YXX>' || r_Seek(n_Preidx).ҽ��С���� || '</YXX><XXX>' || r_Seek(n_Curidx)
                          .ҽ��С���� || '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).ҽ��С���� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>ҽ��С����</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).ҽ��С���� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --����
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).����, 'XXX') <> Nvl(r_Seek(n_Curidx).����, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����</XXM><YXX>' || r_Seek(n_Preidx).���� || '</YXX><XXX>' || r_Seek(n_Curidx).���� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).���� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).���� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --����
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).����, 'XXX') <> Nvl(r_Seek(n_Curidx).����, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����</XXM><YXX>' || r_Seek(n_Preidx).���� || '</YXX><XXX>' || r_Seek(n_Curidx).���� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).���� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).���� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --����
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).����, 'XXX') <> Nvl(r_Seek(n_Curidx).����, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����</XXM><YXX>' || r_Seek(n_Preidx).���� || '</YXX><XXX>' || r_Seek(n_Curidx).���� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).���� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).���� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --��λ�ȼ�
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).��λ�ȼ�, 'XXX') <> Nvl(r_Seek(n_Curidx).��λ�ȼ�, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>��λ�ȼ�</XXM><YXX>' || r_Seek(n_Preidx).��λ�ȼ� || '</YXX><XXX>' || r_Seek(n_Curidx).��λ�ȼ� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).��λ�ȼ� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>��λ�ȼ�</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).��λ�ȼ� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --����ȼ�
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).����ȼ�, 'XXX') <> Nvl(r_Seek(n_Curidx).����ȼ�, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����ȼ�</XXM><YXX>' || r_Seek(n_Preidx).����ȼ� || '</YXX><XXX>' || r_Seek(n_Curidx).����ȼ� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).����ȼ� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����ȼ�</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).����ȼ� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --��ʿ
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).��ʿ, 'XXX') <> Nvl(r_Seek(n_Curidx).��ʿ, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>��ʿ</XXM><YXX>' || r_Seek(n_Preidx).��ʿ || '</YXX><XXX>' || r_Seek(n_Curidx).��ʿ ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).��ʿ Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>��ʿ</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).��ʿ || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --סԺҽʦ
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).סԺҽʦ, 'XXX') <> Nvl(r_Seek(n_Curidx).סԺҽʦ, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>סԺҽʦ</XXM><YXX>' || r_Seek(n_Preidx).סԺҽʦ || '</YXX><XXX>' || r_Seek(n_Curidx).סԺҽʦ ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).סԺҽʦ Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>סԺҽʦ</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).סԺҽʦ || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --����ҽ��
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).����ҽ��, 'XXX') <> Nvl(r_Seek(n_Curidx).����ҽ��, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����ҽ��</XXM><YXX>' || r_Seek(n_Preidx).����ҽ�� || '</YXX><XXX>' || r_Seek(n_Curidx).����ҽ�� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).����ҽ�� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����ҽ��</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).����ҽ�� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --����ҽ��
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).����ҽ��, 'XXX') <> Nvl(r_Seek(n_Curidx).����ҽ��, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����ҽ��</XXM><YXX>' || r_Seek(n_Preidx).����ҽ�� || '</YXX><XXX>' || r_Seek(n_Curidx).����ҽ�� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).����ҽ�� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>����ҽ��</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).����ҽ�� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            --��ǰ����
            v_Value := Null;
            If n_Preidx <> 0 Then
              If Nvl(r_Seek(n_Preidx).��ǰ����, 'XXX') <> Nvl(r_Seek(n_Curidx).��ǰ����, 'XXX') Then
                v_Value := '<ITEM jsonArray="True" ><XXM>��ǰ����</XXM><YXX>' || r_Seek(n_Preidx).��ǰ���� || '</YXX><XXX>' || r_Seek(n_Curidx).��ǰ���� ||
                           '</XXX></ITEM>';
              End If;
            Else
              If r_Seek(n_Curidx).��ǰ���� Is Not Null Then
                v_Value := '<ITEM jsonArray="True" ><XXM>��ǰ����</XXM><YXX></YXX><XXX>' || r_Seek(n_Curidx).��ǰ���� || '</XXX></ITEM>';
              End If;
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            ----�������  �ô���ƴ�����ɣ��򵥴���
            v_Value := Null;
            v_Tmp   := Null;
            v_Tmp1  := Null;
            If n_Preidx <> 0 Then
              --���֮ǰ�ı䶯�Ƿ��ǰ��˴�,�ô���ƴ�����ɣ��򵥴���
              For K In 1 .. r_Seek.Count Loop
                If v_�䶯���� = r_Seek(K).��ֹԭ�� And v_Curʱ�� = r_Seek(K).��ֹʱ�� And r_Seek(K).���Ӵ�λ = 1 Then
                  If v_Tmp Is Null Then
                    v_Tmp := r_Seek(K).����;
                  Else
                    v_Tmp := v_Tmp || ',' || r_Seek(K).����;
                  End If;
                End If;
              End Loop;
            End If;
            --��鵱ǰ�ı䶯�Ƿ��ǰ��˴�,�ô���ƴ�����ɣ��򵥴���
            For K In 1 .. r_Seek.Count Loop
              If v_�䶯���� = r_Seek(K).��ʼԭ�� And v_Curʱ�� = r_Seek(K).��ʼʱ�� And r_Seek(K).���Ӵ�λ = 1 Then
                If v_Tmp1 Is Null Then
                  v_Tmp1 := r_Seek(K).����;
                Else
                  v_Tmp1 := v_Tmp1 || ',' || r_Seek(K).����;
                End If;
              End If;
            End Loop;
            If Nvl(v_Tmp, 'XXX') <> Nvl(v_Tmp1, 'XXX') Then
              v_Value := '<ITEM jsonArray="True" ><XXM>�������</XXM><YXX>' || v_Tmp || '</YXX><XXX>' || v_Tmp1 || '</XXX></ITEM>';
            End If;
            If v_Value Is Not Null Then
              Select Appendchildxml(x_�䶯����, '/ITEM/XXLIST', Xmltype(v_Value)) Into x_�䶯���� From Dual;
            End If;
          
            Select Appendchildxml(x_�䶯ʱ��, '/ITEM/LSXLIST', x_�䶯����) Into x_�䶯ʱ�� From Dual;
          End If;
        End If;
      End Loop;
    
      v_Preʱ�� := v_Curʱ��;
      n_Preidx  := I;
    
      Select Appendchildxml(x_Templet, '/OUTPUT/BDLIST', x_�䶯ʱ��) Into x_Templet From Dual;
    End If;
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getpatichange;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getoperation
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���ȡ����������Ϣ/��ѯ
  --���ڻ�ȡָ�����˵�������Ϣ
  --�����������˷�������ҽ��ʱ����
  --�������������������ʱ����
  --��Ҫ�ֶ�ͬ��ʱ����
  --��ȡ���в�����Ϣ�����
  --��Σ�Xml_In
  --<IN>
  --  <PATIID></PATIID>     --����ID
  --  <PAGEID></PAGEID>     --��ҳID
  --</IN>

  --���Σ�Xml_Out 
  --<OUTPUT>
  --  <SSLIST>
  --    <SS>
  --      <SSMC></SSMC>  //��������
  --      <SSSJ></SSSJ>  //����ʱ��,yyyy-mm-dd hh24:mi
  --      <MZFS></MZFS>  //����ʽ
  --      <SSQK></SSQK>  //�������  ���ڡ��������
  --      <ZXKSID></ZXKSID>   //ִ�п���ID
  --      <ZXKSMC></ZXKSMC>    //ִ�п�������
  --      <FJSS></FJSS>     //��������   1-�ǣ�0-��
  --    <SS>
  --  <SSLIST>
  --</OUTPUT>

  n_����id   ����ҽ����¼.����id%Type;
  n_��ҳid   ����ҽ����¼.��ҳid%Type;
  n_��ҽ��id ����ҽ����¼.Id%Type;
  v_����     ������ĿĿ¼.����%Type;
  v_Xtmp     Clob; --��ʱXML 

  Cursor c_ҽ�� Is
    Select b.����, a.�걾��λ As ����ʱ��, Nvl(a.���id, a.Id) As ��ҽ��id, a.ִ�п���id, c.���� As ִ�п�������,
           Decode(a.���id, Null, 0, 1) As ��������, a.�������, Decode(a.�������, Null, '����', 1, '����', 2, '����', Null) As �������
    From ����ҽ����¼ A, ������ĿĿ¼ B, ���ű� C
    Where a.������Ŀid = b.Id And a.ִ�п���id = c.Id And a.������� In ('F', 'G') And a.ҽ��״̬ <> 4 And Nvl(a.ִ�б��, 0) <> -1 And
          a.����id = n_����id And a.��ҳid = n_��ҳid
    Order By a.������� Desc, a.���;
Begin

  Select Extractvalue(Value(A), 'IN/PATIID') As ����id, Extractvalue(Value(A), 'IN/PAGEID') As ��ҳid
  Into n_����id, n_��ҳid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  n_��ҽ��id := 0;

  For R In c_ҽ�� Loop
    If n_��ҽ��id <> r.��ҽ��id Then
      n_��ҽ��id := r.��ҽ��id;
      If r.������� = 'G' Then
        v_���� := r.����;
      End If;
    Else
      v_���� := Null;
    End If;
  
    If r.������� = 'F' Then
      v_Xtmp := v_Xtmp || '<SS jsonArray="True" >';
      v_Xtmp := v_Xtmp || '<SSMC>' || r.���� || '</SSMC>'; --  //��������
      v_Xtmp := v_Xtmp || '<SSSJ>' || r.����ʱ�� || '</SSSJ>'; --  //����ʱ��,yyyy-mm-dd hh24:mi
      v_Xtmp := v_Xtmp || '<MZFS>' || v_���� || '</MZFS>'; --  //����ʽ
      v_Xtmp := v_Xtmp || '<SSQK>' || r.������� || '</SSQK>'; --  //�������  ���ڡ��������
      v_Xtmp := v_Xtmp || '<ZXKSID>' || r.ִ�п���id || '</ZXKSID>'; --   //ִ�п���ID
      v_Xtmp := v_Xtmp || '<ZXKSMC>' || r.ִ�п������� || '</ZXKSMC>'; --    //ִ�п�������
      v_Xtmp := v_Xtmp || '<FJSS>' || r.�������� || '</FJSS>'; --    //��������   1-�ǣ�0-��    
      v_Xtmp := v_Xtmp || '</SS>';
    End If;
  End Loop;

  If v_Xtmp Is Not Null Then
    Xml_Out := Xmltype('<OUTPUT><SSLIST>' || v_Xtmp || '</SSLIST></OUTPUT>');
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getoperation;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getallpatiinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���ȡ�������в��˻�����Ϣ/��ѯ
  --���ڻ�ȡĳ�������в��˵Ļ�����Ϣ��
  --�ڵ�һ�λ�ȡ������Ϣʱ����
  --ÿ�������Զ�ͬ��ʱ����
  --��Ҫ�ֶ�ͬ��ʱ����

  --��Σ�Xml_In
  --<INPUT>
  --  <BQID></BQID>      --����ID
  --  <CYTS></CYTS>      --��Ժ��������ȡ������֮�ڳ�Ժ�Ĳ���
  --</INPUT>

  --���Σ�Xml_Out 
  --<OUTPUT>
  --  <PATILIST>
  --    <PATI>
  --      <JBXX>    --������Ϣ
  --        <PATIID></PATIID>         --����ID
  --        <PAGEID></PAGEID>  --��ҳID
  --        <BABY></BABY>  --Ӥ�����
  --        <XM></XM>   --����
  --        <XB></XB>   --�Ա�
  --        <NL></NL>   --����
  --        <CSRQ></CSRQ>  --��������
  --        <ZYH></ZYH>  --סԺ��
  --        <HY></HY>   --����
  --        <GJ></GJ>   --����
  --        <MZ></MZ>   --����
  --        <XL></XL>   --ѧ��
  --        <SF></SF>   --���
  --        <ZY></ZY>   --ְҵ
  --        <SFZH></SFZH>  --���֤��
  --        <FKFS></FKFS>  --���ʽ
  --        <LXFS></LXFS>  --��ϵ��ʽ
  --        <LXRXM></LXRXM>  --��ϵ������
  --        <LXRDH></LXRDH>  --��ϵ�˵绰
  --        <LXRDZ></LXRDZ>  --��ϵ�˵�ַ
  --        <JTDH></JTDH>  --��ͥ�绰
  --        <JTDZ></JTDZ>  --��ͥ��ַ
  --        <CSDD></CSDD>  --�����ص�
  --        <GMS></GMS>  --����ʷ
  --      </JBXX>
  --      <ZYXX>    --סԺ��Ϣ
  --        <RYRQ></RYRQ>  --��Ժ����
  --        <RKRQ></RKRQ>  --�������
  --        <CYRQ></CYRQ>  --��Ժ����
  --        <ZYTS></ZYTS>  --סԺ����
  --        <RYFS></RYFS>  --��Ժ��ʽ
  --        <KSID></KSID>  --����ID
  --        <KSMC></KSMC>  --��������
  --        <BQID></BQID>  --����ID
  --        <BQMC></BQMC>  --��������
  --        <CH></CH>   --����
  --        <BQ></BQ>   --����
  --        <ZZYS></ZZYS>  --����ҽʦ
  --        <ZRYS></ZRYS>  --����ҽʦ
  --        <ZYYS></ZYYS>  --סԺҽʦ
  --        <ZRHS></ZRHS>  --���λ�ʿ
  --        <HLDJ></HLDJ>  --����ȼ�
  --        <YLZ></YLZ>    --ҽ��С��id
  --        <YBH></YBH>    --ҽ����
  --        <YBMC></YBMC>  --ҽ������
  --      </ZYXX>
  --    </PATI>
  --  </PATILIST>
  --</OUTPUT>

  n_����id ���ű�.Id%Type;
  v_����   ���ű�.����%Type;
  n_����   Number;
  v_Xtmp   Clob; --��ʱXML 
  x_Item   Xmltype;
  d_��ʼ   Date;
  d_����   Date;

  v_������Ϣ Varchar2(5000);
  v_����ҽʦ Varchar2(500);
  v_����ҽʦ Varchar2(500);
  x_Templet  Xmltype;

  Cursor c_��Ժ Is
    Select a.����id, b.��ҳid, 0 As Ӥ�����, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, Nvl(b.����, a.����) As ����, a.��������, b.סԺ��,
           a.����״�� As ����, a.����, a.����, a.ѧ��, a.���, a.ְҵ, a.���֤��, a.ҽ�Ƹ��ʽ As ���ʽ, a.�ֻ��� As ��ϵ��ʽ, a.��ϵ������, a.��ϵ�˵绰,
           a.��ϵ�˵�ַ, a.��ͥ�绰, a.��ͥ��ַ, a.�����ص�, '����������ѯ' As ����ʷ, Decode(b.���ʱ��, Null, b.��Ժ����, b.���ʱ��) As ��Ժ����,
           b.���ʱ�� As �������, Null As ��Ժ����, (Trunc(Sysdate) - Trunc(Decode(b.���ʱ��, Null, b.��Ժ����, b.���ʱ��))) As סԺ����, b.��Ժ��ʽ,
           b.��Ժ����id As ����id, c.���� As ��������, r.����id, v_���� As ��������, b.��Ժ���� As ����, b.��ǰ���� As ����, '����������ҳ�ӱ�' As ����ҽʦ,
           '����������ҳ�ӱ�' As ����ҽʦ, b.סԺҽʦ, b.���λ�ʿ, e.���� As ����ȼ�, b.ҽ��С��id, a.ҽ����, d.���� As ҽ������

    From ������Ϣ A, ������ҳ B, ���ű� C, ������� D, �շ���ĿĿ¼ E, ��Ժ���� R
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��Ժ����id = c.Id And b.���� = d.���(+) And Nvl(b.״̬, 0) <> 1 And
          b.����ȼ�id = e.Id(+) And (r.����id = n_����id Or b.Ӥ������id = n_����id) And a.����id = r.����id And a.��ǰ����id + 0 = r.����id And
          Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null
    Order By b.��Ժ����;

  Cursor c_��Ժ Is
    Select a.����id, b.��ҳid, 0 As Ӥ�����, Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, Nvl(b.����, a.����) As ����, a.��������, b.סԺ��,
           a.����״�� As ����, a.����, a.����, a.ѧ��, a.���, a.ְҵ, a.���֤��, a.ҽ�Ƹ��ʽ As ���ʽ, a.�ֻ��� As ��ϵ��ʽ, a.��ϵ������, a.��ϵ�˵绰,
           a.��ϵ�˵�ַ, a.��ͥ�绰, a.��ͥ��ַ, a.�����ص�, '����������ѯ' As ����ʷ, Decode(b.���ʱ��, Null, b.��Ժ����, b.���ʱ��) As ��Ժ����,
           b.���ʱ�� As �������, b.��Ժ����, (Trunc(b.��Ժ����) - Trunc(Decode(b.���ʱ��, Null, b.��Ժ����, b.���ʱ��))) As סԺ����, b.��Ժ��ʽ,
           b.��Ժ����id As ����id, c.���� As ��������, b.��ǰ����id As ����id, v_���� As ��������, b.��Ժ���� As ����, b.��ǰ���� As ����,
           '����������ҳ�ӱ�' As ����ҽʦ, '����������ҳ�ӱ�' As ����ҽʦ, b.סԺҽʦ, b.���λ�ʿ, e.���� As ����ȼ�, b.ҽ��С��id, a.ҽ����, d.���� As ҽ������
    From ������Ϣ A, ������ҳ B, ���ű� C, ������� D, �շ���ĿĿ¼ E
    Where a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And b.״̬ = 0 And b.��Ժ����id = c.Id And b.���� = d.���(+) And
          b.����ȼ�id = e.Id(+) And b.��ǰ����id + 0 = n_����id And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And
          b.��Ժ���� Between d_��ʼ And d_����
    Order By b.��Ժ����;

  --�ŵ�ѭ����ִ�еģ���������������
  Procedure p_Getother
  (
    ����id_In    In ������Ϣ.����id%Type,
    ��ҳid_In    In ������ҳ.��ҳid%Type,
    ������Ϣ_Out Out Varchar2,
    ����ҽʦ_Out Out Varchar2,
    ����ҽʦ_Out Out Varchar2
  ) Is
  Begin
  
    ������Ϣ_Out := Null;
    ����ҽʦ_Out := Null;
    ����ҽʦ_Out := Null;
 
    For R In (Select a.��Ϣ��, a.��Ϣֵ
              From ������ҳ�ӱ� A
              Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And a.��Ϣ�� In ('����ҽʦ', '����ҽʦ')) Loop
      If r.��Ϣ�� = '����ҽʦ' Then
        ����ҽʦ_Out := r.��Ϣֵ;
      Elsif r.��Ϣ�� = '����ҽʦ' Then
        ����ҽʦ_Out := r.��Ϣֵ;
      End If;
    End Loop;
  
    For R In (Select a.ҩ����
              From ���˹�����¼ A, ���˹Һż�¼ B, ������ҳ C
              Where a.����id = b.����id(+) And a.��ҳid = b.Id(+) And b.��¼����(+) = 1 And b.��¼״̬(+) = 1 And a.����id = c.����id(+) And
                    a.��ҳid = c.��ҳid(+) And a.��� = 1 And ҩ���� Is Not Null And a.����id = 202 And Not Exists
               (Select ҩ��id
                     From ���˹�����¼
                     Where (Nvl(ҩ��id, 0) = Nvl(a.ҩ��id, 0) Or Nvl(ҩ����, 'Null') = Nvl(a.ҩ����, 'Null')) And Nvl(���, 0) = 0 And
                           ��¼ʱ�� > a.��¼ʱ�� And ����id = 202)
              Group By a.ҩ����
              Order By a.ҩ����) Loop
      ������Ϣ_Out := ������Ϣ_Out || ',' || r.ҩ����;
    End Loop;
    ������Ϣ_Out := Substr(������Ϣ_Out, 2);
  End;

Begin
  Select Extractvalue(Value(A), 'IN/BQID') As ����id, Extractvalue(Value(A), 'IN/CYTS') As ����
  Into n_����id, n_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Sysdate Into d_���� From Dual;

  d_��ʼ := Trunc(d_����) - n_����; --����� 00:00:00  
  d_���� := Trunc(d_����) + 1 - 1 / 24 / 60; --����� 23:59:59

  Select Max(a.����) Into v_���� From ���ű� A Where a.Id = n_����id;

  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  x_Item    := Xmltype('<PATILIST></PATILIST>');

  For R In c_��Ժ Loop
    p_Getother(r.����id, r.��ҳid, v_������Ϣ, v_����ҽʦ, v_����ҽʦ);
    v_Xtmp := '<PATI jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<JBXX>';
    v_Xtmp := v_Xtmp || '<PATIID>' || r.����id || '</PATIID>';
    v_Xtmp := v_Xtmp || '<PAGEID>' || r.��ҳid || '</PAGEID>';
    v_Xtmp := v_Xtmp || '<BABY>' || r.Ӥ����� || '</BABY>';
    v_Xtmp := v_Xtmp || '<XM>' || r.���� || '</XM>';
    v_Xtmp := v_Xtmp || '<XB>' || r.�Ա� || '</XB>';
    v_Xtmp := v_Xtmp || '<NL>' || r.���� || '</NL>';
    v_Xtmp := v_Xtmp || '<CSRQ>' || To_Char(r.��������, 'yyyy-mm-dd hh24:mi:ss') || '</CSRQ>';
    v_Xtmp := v_Xtmp || '<ZYH>' || r.סԺ�� || '</ZYH>';
    v_Xtmp := v_Xtmp || '<HY>' || r.���� || '</HY>';
    v_Xtmp := v_Xtmp || '<GJ>' || r.���� || '</GJ>';
    v_Xtmp := v_Xtmp || '<MZ>' || r.���� || '</MZ>';
    v_Xtmp := v_Xtmp || '<XL>' || r.ѧ�� || '</XL>';
    v_Xtmp := v_Xtmp || '<SF>' || r.��� || '</SF>';
    v_Xtmp := v_Xtmp || '<ZY>' || r.ְҵ || '</ZY>';
    v_Xtmp := v_Xtmp || '<SFZH>' || r.���֤�� || '</SFZH>';
    v_Xtmp := v_Xtmp || '<FKFS>' || r.���ʽ || '</FKFS>';
    v_Xtmp := v_Xtmp || '<LXFS>' || r.��ϵ��ʽ || '</LXFS>';
    v_Xtmp := v_Xtmp || '<LXRXM>' || r.��ϵ������ || '</LXRXM>';
    v_Xtmp := v_Xtmp || '<LXRDH>' || r.��ϵ�˵绰 || '</LXRDH>';
    v_Xtmp := v_Xtmp || '<LXRDZ>' || r.��ϵ�˵�ַ || '</LXRDZ>';
    v_Xtmp := v_Xtmp || '<JTDH>' || r.��ͥ�绰 || '</JTDH>';
    v_Xtmp := v_Xtmp || '<JTDZ>' || r.��ͥ��ַ || '</JTDZ>';
    v_Xtmp := v_Xtmp || '<CSDD>' || r.�����ص� || '</CSDD>';
    v_Xtmp := v_Xtmp || '<GMS>' || v_������Ϣ || '</GMS>'; -- r.����ʷ 
    v_Xtmp := v_Xtmp || '</JBXX>';
    v_Xtmp := v_Xtmp || '<ZYXX>';
    v_Xtmp := v_Xtmp || '<RYRQ>' || To_Char(r.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || '</RYRQ>';
    v_Xtmp := v_Xtmp || '<RKRQ>' || To_Char(r.�������, 'yyyy-mm-dd hh24:mi:ss') || '</RKRQ>';
    v_Xtmp := v_Xtmp || '<CYRQ>' || To_Char(r.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || '</CYRQ>';
    v_Xtmp := v_Xtmp || '<ZYTS>' || r.סԺ���� || '</ZYTS>';
    v_Xtmp := v_Xtmp || '<RYFS>' || r.��Ժ��ʽ || '</RYFS>';
    v_Xtmp := v_Xtmp || '<KSID>' || r.����id || '</KSID>';
    v_Xtmp := v_Xtmp || '<KSMC>' || r.�������� || '</KSMC>';
    v_Xtmp := v_Xtmp || '<BQID>' || r.����id || '</BQID>';
    v_Xtmp := v_Xtmp || '<BQMC>' || r.�������� || '</BQMC>';
    v_Xtmp := v_Xtmp || '<CH>' || r.���� || '</CH>';
    v_Xtmp := v_Xtmp || '<BQ>' || r.���� || '</BQ>';
    v_Xtmp := v_Xtmp || '<ZZYS>' || v_����ҽʦ || '</ZZYS>';
    v_Xtmp := v_Xtmp || '<ZRYS>' || v_����ҽʦ || '</ZRYS>';
    v_Xtmp := v_Xtmp || '<ZYYS>' || r.סԺҽʦ || '</ZYYS>';
    v_Xtmp := v_Xtmp || '<ZRHS>' || r.���λ�ʿ || '</ZRHS>';
    v_Xtmp := v_Xtmp || '<HLDJ>' || r.����ȼ� || '</HLDJ>';
    v_Xtmp := v_Xtmp || '<YLZ>' || r.ҽ��С��id || '</YLZ>';
    v_Xtmp := v_Xtmp || '<YBH>' || r.ҽ���� || '</YBH>';
    v_Xtmp := v_Xtmp || '<YBMC>' || r.ҽ������ || '</YBMC>';
    v_Xtmp := v_Xtmp || '</ZYXX>';
    v_Xtmp := v_Xtmp || '</PATI>';
    Select Appendchildxml(x_Item, '/PATILIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
  End Loop;

  For R In c_��Ժ Loop
    p_Getother(r.����id, r.��ҳid, v_������Ϣ, v_����ҽʦ, v_����ҽʦ);
    v_Xtmp := '<PATI jsonArray="True" >';
    v_Xtmp := v_Xtmp || '<JBXX>';
    v_Xtmp := v_Xtmp || '<PATIID>' || r.����id || '</PATIID>';
    v_Xtmp := v_Xtmp || '<PAGEID>' || r.��ҳid || '</PAGEID>';
    v_Xtmp := v_Xtmp || '<BABY>' || r.Ӥ����� || '</BABY>';
    v_Xtmp := v_Xtmp || '<XM>' || r.���� || '</XM>';
    v_Xtmp := v_Xtmp || '<XB>' || r.�Ա� || '</XB>';
    v_Xtmp := v_Xtmp || '<NL>' || r.���� || '</NL>';
    v_Xtmp := v_Xtmp || '<CSRQ>' || To_Char(r.��������, 'yyyy-mm-dd hh24:mi:ss') || '</CSRQ>';
    v_Xtmp := v_Xtmp || '<ZYH>' || r.סԺ�� || '</ZYH>';
    v_Xtmp := v_Xtmp || '<HY>' || r.���� || '</HY>';
    v_Xtmp := v_Xtmp || '<GJ>' || r.���� || '</GJ>';
    v_Xtmp := v_Xtmp || '<MZ>' || r.���� || '</MZ>';
    v_Xtmp := v_Xtmp || '<XL>' || r.ѧ�� || '</XL>';
    v_Xtmp := v_Xtmp || '<SF>' || r.��� || '</SF>';
    v_Xtmp := v_Xtmp || '<ZY>' || r.ְҵ || '</ZY>';
    v_Xtmp := v_Xtmp || '<SFZH>' || r.���֤�� || '</SFZH>';
    v_Xtmp := v_Xtmp || '<FKFS>' || r.���ʽ || '</FKFS>';
    v_Xtmp := v_Xtmp || '<LXFS>' || r.��ϵ��ʽ || '</LXFS>';
    v_Xtmp := v_Xtmp || '<LXRXM>' || r.��ϵ������ || '</LXRXM>';
    v_Xtmp := v_Xtmp || '<LXRDH>' || r.��ϵ�˵绰 || '</LXRDH>';
    v_Xtmp := v_Xtmp || '<LXRDZ>' || r.��ϵ�˵�ַ || '</LXRDZ>';
    v_Xtmp := v_Xtmp || '<JTDH>' || r.��ͥ�绰 || '</JTDH>';
    v_Xtmp := v_Xtmp || '<JTDZ>' || r.��ͥ��ַ || '</JTDZ>';
    v_Xtmp := v_Xtmp || '<CSDD>' || r.�����ص� || '</CSDD>';
    v_Xtmp := v_Xtmp || '<GMS>' || v_������Ϣ || '</GMS>'; -- r.����ʷ 
    v_Xtmp := v_Xtmp || '</JBXX>';
    v_Xtmp := v_Xtmp || '<ZYXX>';
    v_Xtmp := v_Xtmp || '<RYRQ>' || To_Char(r.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || '</RYRQ>';
    v_Xtmp := v_Xtmp || '<RKRQ>' || To_Char(r.�������, 'yyyy-mm-dd hh24:mi:ss') || '</RKRQ>';
    v_Xtmp := v_Xtmp || '<CYRQ>' || To_Char(r.��Ժ����, 'yyyy-mm-dd hh24:mi:ss') || '</CYRQ>';
    v_Xtmp := v_Xtmp || '<ZYTS>' || r.סԺ���� || '</ZYTS>';
    v_Xtmp := v_Xtmp || '<RYFS>' || r.��Ժ��ʽ || '</RYFS>';
    v_Xtmp := v_Xtmp || '<KSID>' || r.����id || '</KSID>';
    v_Xtmp := v_Xtmp || '<KSMC>' || r.�������� || '</KSMC>';
    v_Xtmp := v_Xtmp || '<BQID>' || r.����id || '</BQID>';
    v_Xtmp := v_Xtmp || '<BQMC>' || r.�������� || '</BQMC>';
    v_Xtmp := v_Xtmp || '<CH>' || r.���� || '</CH>';
    v_Xtmp := v_Xtmp || '<BQ>' || r.���� || '</BQ>';
    v_Xtmp := v_Xtmp || '<ZZYS>' || v_����ҽʦ || '</ZZYS>';
    v_Xtmp := v_Xtmp || '<ZRYS>' || v_����ҽʦ || '</ZRYS>';
    v_Xtmp := v_Xtmp || '<ZYYS>' || r.סԺҽʦ || '</ZYYS>';
    v_Xtmp := v_Xtmp || '<ZRHS>' || r.���λ�ʿ || '</ZRHS>';
    v_Xtmp := v_Xtmp || '<HLDJ>' || r.����ȼ� || '</HLDJ>';
    v_Xtmp := v_Xtmp || '<YLZ>' || r.ҽ��С��id || '</YLZ>';
    v_Xtmp := v_Xtmp || '<YBH>' || r.ҽ���� || '</YBH>';
    v_Xtmp := v_Xtmp || '<YBMC>' || r.ҽ������ || '</YBMC>';
    v_Xtmp := v_Xtmp || '</ZYXX>';
    v_Xtmp := v_Xtmp || '</PATI>';
    Select Appendchildxml(x_Item, '/PATILIST', Xmltype(v_Xtmp)) Into x_Item From Dual;
  End Loop;

  Select Appendchildxml(x_Templet, '/OUTPUT', x_Item) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getallpatiinfo;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
Create Or Replace Procedure Zl_Third_Getkfcws
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --���ܣ����Ŵ�λ��/��ѯ
  --���Ŵ�λ����ÿ��ҹ��12�㿪�Ų������ܺͣ����۸ô��Ƿ񱻲���ռ��,��Ӧ�������ڡ�����������С�������ͣʹ�õĲ�����
  --             �Լ���������ļӴ����������򲡷���������޶�ͣ�õĲ�������ʱ���財��

  --ƽ�����Ŵ�λ����ÿ��ƽ�����Ŵ�λ��=����ÿ�տ��Ŵ�λ��֮��/��������
  --����ZLHISϵͳ���ݽṹ����⣺��λ������¼�������ӵĴ�λ��¼
  --��Σ�xml_in
  --<IN>
  --    <BQID></BQID>    //����ID������ȡ���в���
  --    <KSRQ></KSRQ>  //��ʼ����   yyyy-mm-dd
  --    <JSRQ></JSRQ>   //��������  yyyy-mm-dd
  --</IN>

  --���Σ�xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //����ID
  --      <BQMC></BQMC>  //��������
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //�·�
  --           <KFCR></KFCR>  //���Ŵ�λ��
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT>

  n_����id   ���ű�.Id%Type;
  d_��ʼ     Date;
  d_����     Date;
  v_Xtmp     Varchar(5000);
  x_Tmp      Xmltype;
  x_Templet  Xmltype;
  n_����λ�� ����ҽ����¼.Id%Type;
  n_������   Number;
  v_�������� ���ű�.����%Type;

  v_�·�   Varchar(30);
  d_Tmp    Date;
  n_������ Number;

  Cursor c_Item(����id_In ���ű�.Id%Type) Is
    Select a.����id, a.��, Sum(a.�䶯) As ���Ŵ�λ��
    From (Select a.����id, a.�䶯, To_Char(a.����, 'yyyy-mm-dd') As ��
           From ��λ������¼ A
           Where a.���� Between d_��ʼ And d_���� And a.����id = ����id_In) A
    Group By a.����id, a.��
    Order By a.��;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;
Begin
  Select Extractvalue(Value(A), 'IN/BQID') As ����id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼ����,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��������
  Into n_����id, d_��ʼ, d_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');

  --��������
  If Nvl(n_����id, 0) <> 0 Then
    Select ���� Into v_�������� From ���ű� Where ID = n_����id;
    Select Nvl(Sum(a.�䶯), 0) Into n_����λ�� From ��λ������¼ A Where a.����id = n_����id And a.���� < d_��ʼ;
    Open c_Item(n_����id);
    Fetch c_Item Bulk Collect
      Into r_Item;
    Close c_Item;
    v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_����id || '</BQID><BQMC>' || v_�������� || '</BQMC>';
    v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
    x_Tmp    := Xmltype(v_Xtmp);
    n_������ := 0;
    d_Tmp    := d_��ʼ;
    v_�·�   := '-';
    --ѭ������
    While d_Tmp <= d_���� Loop
      For J In 1 .. r_Item.Count Loop
        If r_Item(J).�� = To_Char(d_Tmp, 'yyyy-mm-dd') Then
          n_����λ�� := n_����λ�� + r_Item(J).���Ŵ�λ��;
        End If;
      End Loop;
    
      If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_�·� Then
        --1.ƴ��֮ǰ��
        If v_�·� <> '-' Then
          n_������ := Round(n_������ / n_������);
          v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><KFCR>' || n_������ || '</KFCR></ITEM>';
          Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
        End If;
        v_�·� := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
        --2.����
        n_������ := 1;
        n_������ := n_����λ��;
      Else
        n_������ := n_������ + 1;
        n_������ := n_������ + n_����λ��;
      End If;
      d_Tmp := d_Tmp + 1;
    End Loop;
    n_������ := Round(n_������ / n_������);
    v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><KFCR>' || n_������ || '</KFCR></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
    Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
  Else
    --���в���  
    For R In (Select a.Id, a.����, a.����
              From ���ű� A, ��������˵�� B
              Where a.Id = b.����id And b.�������� = '����' And ������� = 2
              Group By a.Id, a.����, a.����
              Order By a.����) Loop
      v_�������� := r.����;
      n_����id   := r.Id;
    
      Select Nvl(Sum(a.�䶯), 0) Into n_����λ�� From ��λ������¼ A Where a.����id = n_����id And a.���� < d_��ʼ;
      Open c_Item(n_����id);
      Fetch c_Item Bulk Collect
        Into r_Item;
      Close c_Item;
    
      v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_����id || '</BQID><BQMC>' || v_�������� || '</BQMC>';
      v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp    := Xmltype(v_Xtmp);
      n_������ := 0;
      d_Tmp    := d_��ʼ;
      v_�·�   := '-';
      --ѭ������
      While d_Tmp <= d_���� Loop
        For J In 1 .. r_Item.Count Loop
          If r_Item(J).�� = To_Char(d_Tmp, 'yyyy-mm-dd') Then
            n_����λ�� := n_����λ�� + r_Item(J).���Ŵ�λ��;
          End If;
        End Loop;
      
        If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_�·� Then
          --1.ƴ��֮ǰ��
          If v_�·� <> '-' Then
            n_������ := Round(n_������ / n_������);
            v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><KFCR>' || n_������ || '</KFCR></ITEM>';
            Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
          End If;
          v_�·� := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
          --2.����
          n_������ := 1;
          n_������ := n_����λ��;
        Else
          n_������ := n_������ + 1;
          n_������ := n_������ + n_����λ��;
        End If;
        d_Tmp := d_Tmp + 1;
      End Loop;
      n_������ := Round(n_������ / n_������);
      v_Xtmp   := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><KFCR>' || n_������ || '</KFCR></ITEM>';
      Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
      Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
    End Loop;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getkfcws;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
Create Or Replace Procedure Zl_Third_Getzyrs
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ�סԺ��������/��ѯ
  --��Σ�Xml_In
  --<IN>
  --  <BQID></BQID>    //����ID������ȡ���в���
  --  <KSRQ></KSRQ>  //��ʼ����   yyyy-mm-dd
  --  <JSRQ></JSRQ>   //��������  yyyy-mm-dd
  --</IN>
  --���Σ�xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //����ID
  --      <BQMC></BQMC>  //��������
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //�·�
  --           <QCRS></QCRS>  //�ڳ�����������ʼʱ�����Ժ������
  --           <XRRS></XRRS>  //������������ʱ��������벡����������������Ժ��ת��
  --           <XCRS></XCRS>  //�³���������ʱ������³�������������������Ժ��ת��������
  --           <QMRS></QMRS>  //��ĩ������������ʱ�����Ժ������
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT>

  n_����id   ����ҽ����¼.ִ�п���id%Type;
  v_�������� ���ű�.����%Type;
  v_Xtmp     Varchar(5000); --��ʱXML
  x_Tmp      Xmltype;
  d_��ʼ     Date;
  d_����     Date;
  x_Templet  Xmltype;
  d_Tmp      Date;
  n_�ڳ����� Number;
  n_�������� Number;
  n_�³����� Number;
  n_��ĩ���� Number;
  d_s        Date;
  d_e        Date;

  v_�·� Varchar(50);

  --����ָ��ʱ��������
  Cursor c_��ǰ����
  (
    ʱ��_In   Date,
    ����id_In ����ҽ����¼.ִ�п���id%Type
  ) Is
    Select Count(1) As ����
    From ���˱䶯��¼ A
    Where ��ʼʱ�� < ʱ��_In And (��ֹʱ�� Is Null Or ��ֹʱ�� > ʱ��_In) And Nvl(a.���Ӵ�λ, 0) = 0 And ����id = ����id_In;

  r_��ǰ���� c_��ǰ����%RowType;

  Cursor c_������
  (
    ʱ����_In Date,
    ʱ��ֹ_In Date,
    ����id_In ����ҽ����¼.ִ�п���id%Type
  ) Is
    Select Count(1) As ����
    From ���˱䶯��¼ A
    Where (a.��ʼԭ�� In (2, 3, 15) Or a.��ʼԭ�� = 1 And Not Exists
           (Select 1 From ���˱䶯��¼ B Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��ʼԭ�� = 2)) And a.����id = ����id_In And
          a.��ʼʱ�� Between ʱ����_In And ʱ��ֹ_In And Nvl(a.���Ӵ�λ, 0) = 0;
  r_������ c_������%RowType;

Begin
  Select Extractvalue(Value(A), 'IN/BQID') As ����id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼ����,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��������
  Into n_����id, d_��ʼ, d_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');

  If Nvl(n_����id, 0) <> 0 Then
    Select ���� Into v_�������� From ���ű� Where ID = n_����id;
  
    v_Xtmp := '<ITEM jsonArray="True" ><BQID>' || n_����id || '</BQID><BQMC>' || v_�������� || '</BQMC>';
    v_Xtmp := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
    x_Tmp  := Xmltype(v_Xtmp);
  
    d_Tmp  := d_��ʼ;
    v_�·� := '-';
    d_s    := d_��ʼ;
  
    --ѭ������ȡ��ÿ���·�
    While d_Tmp <= d_���� Loop
      If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_�·� Then
        If v_�·� <> '-' Then
        
          d_e := Trunc(d_Tmp) - 1 / 24 / 60;
        
          Open c_��ǰ����(d_s, n_����id);
          Fetch c_��ǰ����
            Into r_��ǰ����;
          If c_��ǰ����%RowCount = 0 Then
            n_�ڳ����� := 0;
          Else
            n_�ڳ����� := r_��ǰ����.����;
          End If;
          Close c_��ǰ����;
        
          Open c_��ǰ����(d_e, n_����id);
          Fetch c_��ǰ����
            Into r_��ǰ����;
          If c_��ǰ����%RowCount = 0 Then
            n_��ĩ���� := 0;
          Else
            n_��ĩ���� := r_��ǰ����.����;
          End If;
          Close c_��ǰ����;
        
          Open c_������(d_s, d_e, n_����id);
          Fetch c_������
            Into r_������;
          If c_������%RowCount = 0 Then
            n_�������� := 0;
          Else
            n_�������� := r_������.����;
          End If;
          Close c_������;
        
          n_�³����� := n_�������� + n_�ڳ����� - n_��ĩ����;
        
          v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><QCRS>' || n_�ڳ����� || '</QCRS><XRRS>' || n_�������� ||
                    '</XRRS><XCRS>' || n_�³����� || '</XCRS><QMRS>' || n_��ĩ���� || '</QMRS></ITEM>';
          Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
        
        End If;
        v_�·� := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
        d_s    := d_Tmp;
      
      End If;
      d_Tmp := d_Tmp + 1;
    End Loop;
  
    d_e := d_����;
  
    Open c_��ǰ����(d_s, n_����id);
    Fetch c_��ǰ����
      Into r_��ǰ����;
    If c_��ǰ����%RowCount = 0 Then
      n_�ڳ����� := 0;
    Else
      n_�ڳ����� := r_��ǰ����.����;
    End If;
    Close c_��ǰ����;
  
    Open c_��ǰ����(d_e, n_����id);
    Fetch c_��ǰ����
      Into r_��ǰ����;
    If c_��ǰ����%RowCount = 0 Then
      n_��ĩ���� := 0;
    Else
      n_��ĩ���� := r_��ǰ����.����;
    End If;
    Close c_��ǰ����;
  
    Open c_������(d_s, d_e, n_����id);
    Fetch c_������
      Into r_������;
    If c_������%RowCount = 0 Then
      n_�������� := 0;
    Else
      n_�������� := r_������.����;
    End If;
    Close c_������;
  
    n_�³����� := n_�������� + n_�ڳ����� - n_��ĩ����;
  
    v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><QCRS>' || n_�ڳ����� || '</QCRS><XRRS>' || n_�������� ||
              '</XRRS><XCRS>' || n_�³����� || '</XCRS><QMRS>' || n_��ĩ���� || '</QMRS></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
    If x_Tmp Is Not Null Then
      Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
    End If;
  Else
    --���в���
    For R In (Select a.Id, a.����, a.����
              From ���ű� A, ��������˵�� B
              Where a.Id = b.����id And b.�������� = '����' And ������� = 2
              Group By a.Id, a.����, a.����
              Order By a.����) Loop
      v_�������� := r.����;
      n_����id   := r.Id;
    
      v_Xtmp := '<ITEM jsonArray="True" ><BQID>' || n_����id || '</BQID><BQMC>' || v_�������� || '</BQMC>';
      v_Xtmp := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp  := Xmltype(v_Xtmp);
    
      d_Tmp  := d_��ʼ;
      v_�·� := '-';
      d_s    := d_��ʼ;
    
      --ѭ������ȡ��ÿ���·�
      While d_Tmp <= d_���� Loop
        If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_�·� Then
          If v_�·� <> '-' Then
          
            d_e := Trunc(d_Tmp) - 1 / 24 / 60;
          
            Open c_��ǰ����(d_s, n_����id);
            Fetch c_��ǰ����
              Into r_��ǰ����;
            If c_��ǰ����%RowCount = 0 Then
              n_�ڳ����� := 0;
            Else
              n_�ڳ����� := r_��ǰ����.����;
            End If;
            Close c_��ǰ����;
          
            Open c_��ǰ����(d_e, n_����id);
            Fetch c_��ǰ����
              Into r_��ǰ����;
            If c_��ǰ����%RowCount = 0 Then
              n_��ĩ���� := 0;
            Else
              n_��ĩ���� := r_��ǰ����.����;
            End If;
            Close c_��ǰ����;
          
            Open c_������(d_s, d_e, n_����id);
            Fetch c_������
              Into r_������;
            If c_������%RowCount = 0 Then
              n_�������� := 0;
            Else
              n_�������� := r_������.����;
            End If;
            Close c_������;
          
            n_�³����� := n_�������� + n_�ڳ����� - n_��ĩ����;
          
            v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><QCRS>' || n_�ڳ����� || '</QCRS><XRRS>' || n_�������� ||
                      '</XRRS><XCRS>' || n_�³����� || '</XCRS><QMRS>' || n_��ĩ���� || '</QMRS></ITEM>';
            Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
          
          End If;
          v_�·� := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
          d_s    := d_Tmp;
        
        End If;
        d_Tmp := d_Tmp + 1;
      End Loop;
    
      d_e := d_����;
    
      Open c_��ǰ����(d_s, n_����id);
      Fetch c_��ǰ����
        Into r_��ǰ����;
      If c_��ǰ����%RowCount = 0 Then
        n_�ڳ����� := 0;
      Else
        n_�ڳ����� := r_��ǰ����.����;
      End If;
      Close c_��ǰ����;
    
      Open c_��ǰ����(d_e, n_����id);
      Fetch c_��ǰ����
        Into r_��ǰ����;
      If c_��ǰ����%RowCount = 0 Then
        n_��ĩ���� := 0;
      Else
        n_��ĩ���� := r_��ǰ����.����;
      End If;
      Close c_��ǰ����;
    
      Open c_������(d_s, d_e, n_����id);
      Fetch c_������
        Into r_������;
      If c_������%RowCount = 0 Then
        n_�������� := 0;
      Else
        n_�������� := r_������.����;
      End If;
      Close c_������;
    
      n_�³����� := n_�������� + n_�ڳ����� - n_��ĩ����;
    
      v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><QCRS>' || n_�ڳ����� || '</QCRS><XRRS>' || n_�������� ||
                '</XRRS><XCRS>' || n_�³����� || '</XCRS><QMRS>' || n_��ĩ���� || '</QMRS></ITEM>';
      Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
      If x_Tmp Is Not Null Then
        Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
      End If;
    End Loop;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getzyrs;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getzycws
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ�ʵ��ռ�ô�����/��ѯ
  --ʵ��ռ���ܴ�������ָÿ��ҹ��12��ʵ��ռ�ò�����(��ÿ��ҹ��12��סԺ����)�ܺ͡�
  --                   ����ʵ��ռ�õ���ʱ�Ӵ����ڡ�������Ժ���ڵ���12��ǰ��������ʳ�Ժ�Ĳ���, ��Ϊʵ��ռ�ô�λ1�����ͳ��
  --��Σ�Xml_In
  --<IN>
  --    <BQID></BQID>    //����ID������ȡ���в���
  --    <KSRQ></KSRQ>    //��ʼ����   yyyy-mm-dd
  --    <JSRQ></JSRQ>    //��������  yyyy-mm-dd
  --</IN>

  --���Σ�xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //����ID
  --      <BQMC></BQMC>  //��������
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //�·�
  --           <ZYCR></ZYCR>  //ʵ��ռ�ô�����
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT>
  n_����id     ���ű�.Id%Type;
  v_��������   ���ű�.����%Type;
  d_��ʼ       Date;
  d_����       Date;
  v_Xtmp       Varchar(5000); --��ʱXML
  x_Tmp        Xmltype;
  x_Templet    Xmltype;
  n_������     Number; --����סԺ����֮��
  n_���˴����� Number;
  n_���������� Number(18);
  v_�·�       Varchar2(50);
  d_Tmp        Date;
  v_Pre����    Varchar2(100);
  n_Index      Number;
  d_s          Date;
  d_e          Date;

  d_������    Date;
  d_����ֹ    Date;
  d_Pre����ֹ Date;

  Cursor c_Item
  (
    ʱ����_In Date,
    ʱ��ֹ_In Date,
    ����id_In ����ҽ����¼.ִ�п���id%Type
  ) Is
    Select Case
             When ��ֹԭ�� = 1 And ��ʼԭ�� In (1, 2, 3, 15) And ����id = ����id_In Then
              'ת���ת��'
             When ��ֹԭ�� = 1 Or ��ʼԭ�� In (3, 15) And ����id <> ����id_In Then
              'ת��'
             Else
              'ת��'
           End As ����, ����id, ��ҳid, Trunc(��ʼʱ��) AS ��ʼʱ��, Trunc(��ֹʱ��) AS ��ֹʱ��, ����id, ��ʼԭ��, ��ֹԭ��
    From (Select a.����id, a.��ҳid,
                  Case
                    When Trunc(a.��ʼʱ��) < ʱ����_In Then
                     ʱ����_In
                    Else
                     a.��ʼʱ��
                  End As ��ʼʱ��,
                  Case
                    When Trunc(a.��ֹʱ��) > ʱ��ֹ_In Then
                     ʱ��ֹ_In
                    Else
                     a.��ֹʱ��
                  End As ��ֹʱ��, a.����id, a.��ʼԭ��, a.��ֹԭ��
           From ���˱䶯��¼ A
           Where a.��ʼʱ�� Between ʱ����_In And ʱ��ֹ_In And Exists
            (Select 1 From ���˱䶯��¼ B Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����id = ����id_In) And
                 ((((a.��ʼԭ�� = 2 Or a.��ʼԭ�� = 1 And Not Exists
                  (Select 1 From ���˱䶯��¼ C Where a.����id = c.����id And a.��ҳid = c.��ҳid And c.��ʼԭ�� = 2)) And a.����id = ����id_In Or
                 a.��ʼԭ�� In (3, 15))) Or a.��ֹԭ�� = 1) And Nvl(a.���Ӵ�λ, 0) = 0
           Union All
           Select a.����id, a.��ҳid,
                  Case
                    When Trunc(a.��ʼʱ��) < ʱ����_In Then
                     ʱ����_In
                    Else
                     a.��ʼʱ��
                  End As ��ʼʱ��,
                  Case
                    When Trunc(a.��ֹʱ��) > ʱ��ֹ_In Then
                     ʱ��ֹ_In
                    Else
                     a.��ֹʱ��
                  End As ��ֹʱ��, a.����id, a.��ʼԭ��, a.��ֹԭ��
           From ���˱䶯��¼ A
           Where a.��ʼʱ�� < ʱ����_In And a.����id = ����id_In And
                 ((a.��ʼԭ�� = 2 Or a.��ʼԭ�� = 1 And Not Exists
                  (Select 1 From ���˱䶯��¼ C Where a.����id = c.����id And a.��ҳid = c.��ҳid And c.��ʼԭ�� = 2)) And a.����id = ����id_In) And
                 Nvl(a.���Ӵ�λ, 0) = 0 And Not Exists
            (Select 1
                  From ���˱䶯��¼ B
                  Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��ʼʱ�� < ʱ��ֹ_In And
                        (b.��ʼԭ�� In (3, 15) And b.����id <> ����id_In Or b.��ֹԭ�� = 1 And ����id = ����id_In)))
    Order By ����id, ��ҳid, ��ʼʱ��, ��ֹʱ��;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;
Begin

  Select Extractvalue(Value(A), 'IN/BQID') As ����id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼ����,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��������
  Into n_����id, d_��ʼ, d_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');

  If Nvl(n_����id, 0) <> 0 Then
    Select ���� Into v_�������� From ���ű� Where ID = n_����id;
    v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_����id || '</BQID><BQMC>' || v_�������� || '</BQMC>';
    v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
    x_Tmp    := Xmltype(v_Xtmp);
    d_Tmp    := d_��ʼ;
    v_�·�   := '-';
    d_s      := d_��ʼ;
    n_������ := 0;

    --ѭ������ȡ��ÿ���·�
    While d_Tmp <= d_���� Loop
      If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_�·� Then
        If v_�·� <> '-' Then
          d_e       := Trunc(d_Tmp) - 1 / 24 / 60;
          n_������  := 0;
          v_Pre���� := '-';
          Open c_Item(d_s, d_e, n_����id);
          Fetch c_Item Bulk Collect
            Into r_Item;
          Close c_Item;
          n_���������� := 0;
          For I In 1 .. r_Item.Count Loop
            If r_Item(I).����id || '_' || r_Item(I).��ҳid <> v_Pre���� Or v_Pre���� = '-' Then
              --�²��˿�ʼ
              If r_Item(I).���� = 'ת���ת��' Then
                d_������ := r_Item(I).��ʼʱ��;
                d_����ֹ := r_Item(I).��ֹʱ��;
                If (d_����ֹ - d_������)=0 Then
                  n_���˴����� :=1;
                Else
                  n_���˴����� :=(d_����ֹ - d_������);
                End If;
                n_���������� := n_���������� + n_���˴����� ;
              Elsif r_Item(I).���� = 'ת��' Then
                If I >=r_Item.Count Then
                  d_������ := r_Item(I).��ʼʱ��;
                  d_����ֹ := Trunc(d_e);
                  If (d_����ֹ - d_������)=0 Then
                    n_���˴����� :=1;
                  Else
                    n_���˴����� :=(d_����ֹ - d_������);
                  End If;
                  n_���������� := n_���������� + n_���˴����� ;
                Else
                  If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
                     d_������ := r_Item(I).��ʼʱ��;
                     d_����ֹ := r_Item(I+1).��ʼʱ��;
                     If (d_����ֹ - d_������)=0 Then
                       n_���˴����� :=1;
                     Else
                       n_���˴����� :=(d_����ֹ - d_������);
                     End If;
                     n_���������� := n_���������� + n_���˴����� ;
                  Else
                     d_������ := r_Item(I).��ʼʱ��;
                     d_����ֹ := Trunc(d_e);
                     If (d_����ֹ - d_������)=0 Then
                       n_���˴����� :=1;
                     Else
                       n_���˴����� :=(d_����ֹ - d_������);
                     End If;
                     n_���������� := n_���������� + n_���˴����� ;
                  End If;
                End If;
              Elsif r_Item(I).���� = 'ת��' Then
                d_������ := Trunc(d_s);
                If NVL(r_Item(I).��ֹԭ��,0) = 1 then
                   d_����ֹ := r_Item(I).��ֹʱ��;
                Else
                  d_����ֹ := r_Item(I).��ʼʱ��;
                End If;
                If (d_����ֹ - d_������)=0 Then
                  n_���˴����� :=1;
                Else
                  n_���˴����� :=(d_����ֹ - d_������);
                End If;
                n_���������� := n_���������� + n_���˴����� ;
              End If;
            Else
              If r_Item(I).���� = 'ת���ת��' Then
                --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
                if d_������ = r_Item(I).��ʼʱ�� Then
                  n_���������� := n_���������� - 1;
                End If;
                d_������ := r_Item(I).��ʼʱ��;
                d_����ֹ := r_Item(I).��ֹʱ��;
                If (d_����ֹ - d_������)=0 Then
                  n_���˴����� :=1;
                Else
                  n_���˴����� :=(d_����ֹ - d_������);
                End If;
                n_���������� := n_���������� + n_���˴����� ;
              Elsif r_Item(I).���� = 'ת��' Then
                --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
                if d_������ = r_Item(I).��ʼʱ�� Then
                  n_���������� := n_���������� - 1;
                End If;
                If I >=r_Item.Count Then
                  d_������ := r_Item(I).��ʼʱ��;
                  d_����ֹ := Trunc(d_e);
                  If (d_����ֹ - d_������)=0 Then
                    n_���˴����� :=1;
                  Else
                    n_���˴����� :=(d_����ֹ - d_������);
                  End If;
                  n_���������� := n_���������� + n_���˴����� ;
                Else
                  If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
                     d_������ := r_Item(I).��ʼʱ��;
                     d_����ֹ := r_Item(I+1).��ʼʱ��;
                     If (d_����ֹ - d_������)=0 Then
                       n_���˴����� :=1;
                     Else
                       n_���˴����� :=(d_����ֹ - d_������);
                     End If;
                     n_���������� := n_���������� + n_���˴����� ;
                  Else
                     d_������ := r_Item(I).��ʼʱ��;
                     d_����ֹ := Trunc(d_e);
                     If (d_����ֹ - d_������)=0 Then
                       n_���˴����� :=1;
                     Else
                       n_���˴����� :=(d_����ֹ - d_������);
                     End If;
                     n_���������� := n_���������� + n_���˴����� ;
                  End If;
                End If;
              End If;
            End If;
            v_Pre���� := r_Item(I).����id || '_' || r_Item(I).��ҳid;
          End Loop;
          v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><ZYCR>' || n_���������� || '</ZYCR></ITEM>';
          Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
        End If;
        v_�·� := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
        d_s    := d_Tmp;
      End If;
      d_Tmp := d_Tmp + 1;
    End Loop;

    d_e       := d_����;
    n_������  := 0;
    v_Pre���� := '-';
    Open c_Item(d_s, d_e, n_����id);
    Fetch c_Item Bulk Collect
      Into r_Item;
    Close c_Item;
    n_���������� := 0;
    For I In 1 .. r_Item.Count Loop
      If r_Item(I).����id || '_' || r_Item(I).��ҳid <> v_Pre���� Or v_Pre���� = '-' Then
        --�²��˿�ʼ
        If r_Item(I).���� = 'ת���ת��' Then
          d_������ := r_Item(I).��ʼʱ��;
          d_����ֹ := r_Item(I).��ֹʱ��;
          If (d_����ֹ - d_������)=0 Then
            n_���˴����� :=1;
          Else
            n_���˴����� :=(d_����ֹ - d_������);
          End If;
          n_���������� := n_���������� + n_���˴����� ;
        Elsif r_Item(I).���� = 'ת��' Then
          If I >=r_Item.Count Then
            d_������ := r_Item(I).��ʼʱ��;
            d_����ֹ := Trunc(d_e);
            If (d_����ֹ - d_������)=0 Then
              n_���˴����� :=1;
            Else
              n_���˴����� :=(d_����ֹ - d_������);
            End If;
            n_���������� := n_���������� + n_���˴����� ;
          Else
            If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
               d_������ := r_Item(I).��ʼʱ��;
               d_����ֹ := r_Item(I+1).��ʼʱ��;
               If (d_����ֹ - d_������)=0 Then
                 n_���˴����� :=1;
               Else
                 n_���˴����� :=(d_����ֹ - d_������);
               End If;
               n_���������� := n_���������� + n_���˴����� ;
            Else
               d_������ := r_Item(I).��ʼʱ��;
               d_����ֹ := Trunc(d_e);
               If (d_����ֹ - d_������)=0 Then
                 n_���˴����� :=1;
               Else
                 n_���˴����� :=(d_����ֹ - d_������);
               End If;
               n_���������� := n_���������� + n_���˴����� ;
            End If;
          End If;
        Elsif r_Item(I).���� = 'ת��' Then
          d_������ := Trunc(d_s);
          If NVL(r_Item(I).��ֹԭ��,0) = 1 then
             d_����ֹ := r_Item(I).��ֹʱ��;
          Else
            d_����ֹ := r_Item(I).��ʼʱ��;
          End If;
          If (d_����ֹ - d_������)=0 Then
            n_���˴����� :=1;
          Else
            n_���˴����� :=(d_����ֹ - d_������);
          End If;
          n_���������� := n_���������� + n_���˴����� ;
        End If;
      Else
        If r_Item(I).���� = 'ת���ת��' Then
          --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
          if d_������ = r_Item(I).��ʼʱ�� Then
            n_���������� := n_���������� - 1;
          End If;
          d_������ := r_Item(I).��ʼʱ��;
          d_����ֹ := r_Item(I).��ֹʱ��;
          If (d_����ֹ - d_������)=0 Then
            n_���˴����� :=1;
          Else
            n_���˴����� :=(d_����ֹ - d_������);
          End If;
          n_���������� := n_���������� + n_���˴����� ;
        Elsif r_Item(I).���� = 'ת��' Then
          --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
          if d_������ = r_Item(I).��ʼʱ�� Then
            n_���������� := n_���������� - 1;
          End If;
          If I >=r_Item.Count Then
            d_������ := r_Item(I).��ʼʱ��;
            d_����ֹ := Trunc(d_e);
            If (d_����ֹ - d_������)=0 Then
              n_���˴����� :=1;
            Else
              n_���˴����� :=(d_����ֹ - d_������);
            End If;
            n_���������� := n_���������� + n_���˴����� ;
          Else
            If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
               d_������ := r_Item(I).��ʼʱ��;
               d_����ֹ := r_Item(I+1).��ʼʱ��;
               If (d_����ֹ - d_������)=0 Then
                 n_���˴����� :=1;
               Else
                 n_���˴����� :=(d_����ֹ - d_������);
               End If;
               n_���������� := n_���������� + n_���˴����� ;
            Else
               d_������ := r_Item(I).��ʼʱ��;
               d_����ֹ := Trunc(d_e);
               If (d_����ֹ - d_������)=0 Then
                 n_���˴����� :=1;
               Else
                 n_���˴����� :=(d_����ֹ - d_������);
               End If;
               n_���������� := n_���������� + n_���˴����� ;
            End If;
          End If;
        End If;
      End If;
      v_Pre���� := r_Item(I).����id || '_' || r_Item(I).��ҳid;
    End Loop;
    v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><ZYCR>' || n_���������� || '</ZYCR></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
    If x_Tmp Is Not Null Then
      Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
    End If;
  Else
    --���в���
    For R In (Select a.Id, a.����, a.����
              From ���ű� A, ��������˵�� B
              Where a.Id = b.����id And b.�������� = '����' And ������� = 2
              Group By a.Id, a.����, a.����
              Order By a.����) Loop
      v_�������� := r.����;
      n_����id   := r.Id;

      v_Xtmp   := '<ITEM jsonArray="True" ><BQID>' || n_����id || '</BQID><BQMC>' || v_�������� || '</BQMC>';
      v_Xtmp   := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp    := Xmltype(v_Xtmp);
      d_Tmp    := d_��ʼ;
      v_�·�   := '-';
      d_s      := d_��ʼ;
      n_������ := 0;

      --ѭ������ȡ��ÿ���·�
      While d_Tmp <= d_���� Loop
        If Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7) <> v_�·� Then
          If v_�·� <> '-' Then
            d_e       := Trunc(d_Tmp) - 1 / 24 / 60;
            n_������  := 0;
            v_Pre���� := '-';
            Open c_Item(d_s, d_e, n_����id);
            Fetch c_Item Bulk Collect
              Into r_Item;
            Close c_Item;
            n_���������� := 0;
            For I In 1 .. r_Item.Count Loop
              If r_Item(I).����id || '_' || r_Item(I).��ҳid <> v_Pre���� Or v_Pre���� = '-' Then
                --�²��˿�ʼ
                If r_Item(I).���� = 'ת���ת��' Then
                  d_������ := r_Item(I).��ʼʱ��;
                  d_����ֹ := r_Item(I).��ֹʱ��;
                  If (d_����ֹ - d_������)=0 Then
                    n_���˴����� :=1;
                  Else
                    n_���˴����� :=(d_����ֹ - d_������);
                  End If;
                  n_���������� := n_���������� + n_���˴����� ;
                Elsif r_Item(I).���� = 'ת��' Then
                  If I >=r_Item.Count Then
                    d_������ := r_Item(I).��ʼʱ��;
                    d_����ֹ := Trunc(d_e);
                    If (d_����ֹ - d_������)=0 Then
                      n_���˴����� :=1;
                    Else
                      n_���˴����� :=(d_����ֹ - d_������);
                    End If;
                    n_���������� := n_���������� + n_���˴����� ;
                  Else
                    If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
                       d_������ := r_Item(I).��ʼʱ��;
                       d_����ֹ := r_Item(I+1).��ʼʱ��;
                       If (d_����ֹ - d_������)=0 Then
                         n_���˴����� :=1;
                       Else
                         n_���˴����� :=(d_����ֹ - d_������);
                       End If;
                       n_���������� := n_���������� + n_���˴����� ;
                    Else
                       d_������ := r_Item(I).��ʼʱ��;
                       d_����ֹ := Trunc(d_e);
                       If (d_����ֹ - d_������)=0 Then
                         n_���˴����� :=1;
                       Else
                         n_���˴����� :=(d_����ֹ - d_������);
                       End If;
                       n_���������� := n_���������� + n_���˴����� ;
                    End If;
                  End If;
                Elsif r_Item(I).���� = 'ת��' Then
                  d_������ := Trunc(d_s);
                  If NVL(r_Item(I).��ֹԭ��,0) = 1 then
                     d_����ֹ := r_Item(I).��ֹʱ��;
                  Else
                    d_����ֹ := r_Item(I).��ʼʱ��;
                  End If;
                  If (d_����ֹ - d_������)=0 Then
                    n_���˴����� :=1;
                  Else
                    n_���˴����� :=(d_����ֹ - d_������);
                  End If;
                  n_���������� := n_���������� + n_���˴����� ;
                End If;
              Else
                If r_Item(I).���� = 'ת���ת��' Then
                  --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
                  if d_������ = r_Item(I).��ʼʱ�� Then
                    n_���������� := n_���������� - 1;
                  End If;
                  d_������ := r_Item(I).��ʼʱ��;
                  d_����ֹ := r_Item(I).��ֹʱ��;
                  If (d_����ֹ - d_������)=0 Then
                    n_���˴����� :=1;
                  Else
                    n_���˴����� :=(d_����ֹ - d_������);
                  End If;
                  n_���������� := n_���������� + n_���˴����� ;
                Elsif r_Item(I).���� = 'ת��' Then
                  --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
                  if d_������ = r_Item(I).��ʼʱ�� Then
                    n_���������� := n_���������� - 1;
                  End If;
                  If I >=r_Item.Count Then
                    d_������ := r_Item(I).��ʼʱ��;
                    d_����ֹ := Trunc(d_e);
                    If (d_����ֹ - d_������)=0 Then
                      n_���˴����� :=1;
                    Else
                      n_���˴����� :=(d_����ֹ - d_������);
                    End If;
                    n_���������� := n_���������� + n_���˴����� ;
                  Else
                    If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
                       d_������ := r_Item(I).��ʼʱ��;
                       d_����ֹ := r_Item(I+1).��ʼʱ��;
                       If (d_����ֹ - d_������)=0 Then
                         n_���˴����� :=1;
                       Else
                         n_���˴����� :=(d_����ֹ - d_������);
                       End If;
                       n_���������� := n_���������� + n_���˴����� ;
                    Else
                       d_������ := r_Item(I).��ʼʱ��;
                       d_����ֹ := Trunc(d_e);
                       If (d_����ֹ - d_������)=0 Then
                         n_���˴����� :=1;
                       Else
                         n_���˴����� :=(d_����ֹ - d_������);
                       End If;
                       n_���������� := n_���������� + n_���˴����� ;
                    End If;
                  End If;
                End If;
              End If;
              v_Pre���� := r_Item(I).����id || '_' || r_Item(I).��ҳid;
            End Loop;
            v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><ZYCR>' || n_���������� || '</ZYCR></ITEM>';
            Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
          End If;
          v_�·� := Substr(To_Char(d_Tmp, 'yyyy-mm-dd'), 1, 7);
          d_s    := d_Tmp;
        End If;
        d_Tmp := d_Tmp + 1;
      End Loop;

      d_e       := d_����;
      n_������  := 0;
      v_Pre���� := '-';
      Open c_Item(d_s, d_e, n_����id);
      Fetch c_Item Bulk Collect
        Into r_Item;
      Close c_Item;
      n_���������� := 0;
      For I In 1 .. r_Item.Count Loop
        If r_Item(I).����id || '_' || r_Item(I).��ҳid <> v_Pre���� Or v_Pre���� = '-' Then
          --�²��˿�ʼ
          If r_Item(I).���� = 'ת���ת��' Then
            d_������ := r_Item(I).��ʼʱ��;
            d_����ֹ := r_Item(I).��ֹʱ��;
            If (d_����ֹ - d_������)=0 Then
              n_���˴����� :=1;
            Else
              n_���˴����� :=(d_����ֹ - d_������);
            End If;
            n_���������� := n_���������� + n_���˴����� ;
          Elsif r_Item(I).���� = 'ת��' Then
            If I >=r_Item.Count Then
              d_������ := r_Item(I).��ʼʱ��;
              d_����ֹ := Trunc(d_e);
              If (d_����ֹ - d_������)=0 Then
                n_���˴����� :=1;
              Else
                n_���˴����� :=(d_����ֹ - d_������);
              End If;
              n_���������� := n_���������� + n_���˴����� ;
            Else
              If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
                 d_������ := r_Item(I).��ʼʱ��;
                 d_����ֹ := r_Item(I+1).��ʼʱ��;
                 If (d_����ֹ - d_������)=0 Then
                   n_���˴����� :=1;
                 Else
                   n_���˴����� :=(d_����ֹ - d_������);
                 End If;
                 n_���������� := n_���������� + n_���˴����� ;
              Else
                 d_������ := r_Item(I).��ʼʱ��;
                 d_����ֹ := Trunc(d_e);
                 If (d_����ֹ - d_������)=0 Then
                   n_���˴����� :=1;
                 Else
                   n_���˴����� :=(d_����ֹ - d_������);
                 End If;
                 n_���������� := n_���������� + n_���˴����� ;
              End If;
            End If;
          Elsif r_Item(I).���� = 'ת��' Then
            d_������ := Trunc(d_s);
            If NVL(r_Item(I).��ֹԭ��,0) = 1 then
               d_����ֹ := r_Item(I).��ֹʱ��;
            Else
              d_����ֹ := r_Item(I).��ʼʱ��;
            End If;
            If (d_����ֹ - d_������)=0 Then
              n_���˴����� :=1;
            Else
              n_���˴����� :=(d_����ֹ - d_������);
            End If;
            n_���������� := n_���������� + n_���˴����� ;
          End If;
        Else
          If r_Item(I).���� = 'ת���ת��' Then
            --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
            if d_������ = r_Item(I).��ʼʱ�� Then
              n_���������� := n_���������� - 1;
            End If;
            d_������ := r_Item(I).��ʼʱ��;
            d_����ֹ := r_Item(I).��ֹʱ��;
            If (d_����ֹ - d_������)=0 Then
              n_���˴����� :=1;
            Else
              n_���˴����� :=(d_����ֹ - d_������);
            End If;
            n_���������� := n_���������� + n_���˴����� ;
          Elsif r_Item(I).���� = 'ת��' Then
            --�����һ����ʼʱ�����ε���ʼʱ����ͬ�����һ�����磺һ������ͬһ������ת�˶�Σ�
            if d_������ = r_Item(I).��ʼʱ�� Then
              n_���������� := n_���������� - 1;
            End If;
            If I >=r_Item.Count Then
              d_������ := r_Item(I).��ʼʱ��;
              d_����ֹ := Trunc(d_e);
              If (d_����ֹ - d_������)=0 Then
                n_���˴����� :=1;
              Else
                n_���˴����� :=(d_����ֹ - d_������);
              End If;
              n_���������� := n_���������� + n_���˴����� ;
            Else
              If r_Item(I+1).���� = 'ת��' And r_Item(I+1).����id || '_' || r_Item(I+1).��ҳid = r_Item(I).����id || '_' || r_Item(I).��ҳid Then
                 d_������ := r_Item(I).��ʼʱ��;
                 d_����ֹ := r_Item(I+1).��ʼʱ��;
                 If (d_����ֹ - d_������)=0 Then
                   n_���˴����� :=1;
                 Else
                   n_���˴����� :=(d_����ֹ - d_������);
                 End If;
                 n_���������� := n_���������� + n_���˴����� ;
              Else
                 d_������ := r_Item(I).��ʼʱ��;
                 d_����ֹ := Trunc(d_e);
                 If (d_����ֹ - d_������)=0 Then
                   n_���˴����� :=1;
                 Else
                   n_���˴����� :=(d_����ֹ - d_������);
                 End If;
                 n_���������� := n_���������� + n_���˴����� ;
              End If;
            End If;
          End If;
        End If;
        v_Pre���� := r_Item(I).����id || '_' || r_Item(I).��ҳid;
      End Loop;
      v_Xtmp := '<ITEM jsonArray="True" ><YF>' || Substr(v_�·�, 6, 2) || '</YF><ZYCR>' || n_���������� || '</ZYCR></ITEM>';
      Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;

      If x_Tmp Is Not Null Then
        Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
      End If;
    End Loop;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getzycws;
/

--122937:������,2018-03-15,����ӿڳ�����ӵĽ������
CREATE OR REPLACE Procedure Zl_Third_Getzyhzrrs
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --���ܣ���Ժ��ռ���ܴ�����/��ѯ
  --��Ժ��ռ���ܴ�������ָ�����Ժ���ߵ�סԺ����֮�ܺͣ���סԺ��������������Ժ���ڵ���12��ǰ��������ʳ�Ժ�Ĳ���, ��Ϊռ�ô�����1�����ͳ��
  --����ZLHISϵͳ����⣺��ָ��ʱ�䷶Χ��Ժ�Ĳ��˵�סԺ�������ܺ�
  --��Σ�Xml_In
  --<IN>
  --    <BQID></BQID>    //����ID������ȡ���в���
  --    <KSRQ></KSRQ>  //��ʼ����   yyyy-mm-dd
  --    <JSRQ></JSRQ>   //��������  yyyy-mm-dd
  --</IN>

  --���Σ�xml_out
  --<OUTPUT>
  --  <BQLIST>
  --    <ITEM>
  --      <BQID></BQID>  //����ID
  --      <BQMC></BQMC>  //��������
  --      <DATALIST>
  --        <ITEM>
  --           <YF></YF>  //�·�
  --           <CRS></CRS>  //��Ժ����ռ���ܴ�����
  --        </ITEM>
  --      </DATALIST>
  --    </ITEM>
  --  </BQLIST>
  --</OUTPUT> 

  n_����id    ���ű�.Id%Type;
  n_Pre����id ���ű�.Id%Type;
  d_��ʼ      Date;
  d_����      Date;
  v_Xtmp      Varchar(5000); --��ʱXML 
  x_Tmp       Xmltype;
  x_Templet   Xmltype;

  Cursor c_Item Is
    Select m.����id, m.��������, m.��, Sum(m.סԺ����) As ������
    From (Select b.��ǰ����id As ����id, b.����id, b.��ҳid, a.���� As ��������, To_Char(b.��Ժ����, 'mm') As ��,
                  (Trunc(b.��Ժ����) - Trunc(Decode(b.���ʱ��, Null, b.��Ժ����, b.���ʱ��))) As סԺ����
           From ������ҳ B, ���ű� A
           Where b.��ǰ����id = a.Id And b.��ǰ����id = n_����id And b.��Ժ���� Between d_��ʼ And d_����) M
    Group By m.����id, m.��������, m.��
    Having Sum(m.סԺ����) > 0;

  Type t_Item Is Table Of c_Item%RowType;
  r_Item t_Item;

  Cursor c_Itemall Is
    Select m.����id, m.��������, m.��, Sum(m.סԺ����) As ������
    From (Select b.��ǰ����id As ����id, b.����id, b.��ҳid, a.���� As ��������, To_Char(b.��Ժ����, 'mm') As ��,
                  (Trunc(b.��Ժ����) - Trunc(Decode(b.���ʱ��, Null, b.��Ժ����, b.���ʱ��))) As סԺ����
           From ������ҳ B, ���ű� A
           Where b.��ǰ����id = a.Id And b.��Ժ���� Between d_��ʼ And d_����) M
    Group By m.����id, m.��������, m.��
    Having Sum(m.סԺ����) > 0
    Order By m.����id;
Begin
  Select Extractvalue(Value(A), 'IN/BQID') As ����id,
         To_Date(Extractvalue(Value(A), 'IN/KSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼ����,
         To_Date(Extractvalue(Value(A), 'IN/JSRQ'), 'yyyy-mm-dd hh24:mi:ss') As ��������
  Into n_����id, d_��ʼ, d_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_Pre����id := -1;

  If n_����id Is Null Then
    Open c_Itemall;
    Fetch c_Itemall Bulk Collect
      Into r_Item;
    Close c_Itemall;
  Else
    Open c_Item;
    Fetch c_Item Bulk Collect
      Into r_Item;
    Close c_Item;
  End If;
  x_Templet := Xmltype('<OUTPUT><BQLIST></BQLIST></OUTPUT>');
  For I In 1 .. r_Item.Count Loop
    If n_Pre����id <> r_Item(I).����id Then
      If x_Tmp Is Not Null Then
        Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
      End If;
      n_Pre����id := r_Item(I).����id;
      v_Xtmp      := '<ITEM jsonArray="True" ><BQID>' || r_Item(I).����id || '</BQID><BQMC>' || r_Item(I).�������� || '</BQMC>';
      v_Xtmp      := v_Xtmp || '<DATALIST></DATALIST></ITEM>';
      x_Tmp       := Xmltype(v_Xtmp);
    End If;
    v_Xtmp := '<ITEM jsonArray="True" ><YF>' || r_Item(I).�� || '</YF><CRS>' || r_Item(I).������ || '</CRS></ITEM>';
    Select Appendchildxml(x_Tmp, '/ITEM/DATALIST', Xmltype(v_Xtmp)) Into x_Tmp From Dual;
  End Loop;
  If x_Tmp Is Not Null Then
    Select Appendchildxml(x_Templet, '/OUTPUT/BQLIST', x_Tmp) Into x_Templet From Dual;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getzyhzrrs;
/

--122832:����,2018-03-13,�ƶ�HIS�ӿ����ӽڵ���
Create Or Replace Procedure Zl_Third_Tendfile_Gettemphdata
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --����:��ȡĳ����ָ����Χ�ڵ����µ���ʷ����
  --���:Xml_In
  --<IN>
  --<BQ></BQ>       --����ID
  --<PATIID></PATIID>     --����ID
  --<PAGEID></PAGEID>     --��ҳID
  --<BABY></BABY>      --Ӥ��
  -- <FW></FW>   --��Χ�����졢���졢 һ��
  --</IN>
  -- ����:Xml_Out
  --<OUTPUT>
  -- <GROUPS>
  --  <GROUP>
  --   <SJ></SJ>   --����ʱ��
  --   <CZY></CZY>  --����Ա
  --   <ITEMS>
  --    <ITEM>
  --     <XH></XH>   --���
  --     <MC></MC>   --����
  --     <NR></NR>   --����
  --     <WJ />     --δ��˵��
  --     <BW />     --��λ
  --    </ITEM>
  --   </ITEMS>
  --  </GROUP>
  -- </GROUPS>
  --</OUTPUT>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Patiid   Number(18);
  n_Pageid   Number(18);
  n_Baby     Number(18);
  n_Areaid   Number(18);
  n_Fw       Number(18);
  d_��ʼʱ�� Date;
  d_����ʱ�� Date;
  v_Temp     Varchar2(32767);
  x_Templet  Xmltype; --ģ��XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/BQ'))
  Into n_Areaid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PATIID'))
  Into n_Patiid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_Pageid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/BABY'))
  Into n_Baby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/FW')) Into n_Fw From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;
  d_����ʱ�� := Sysdate;
  d_��ʼʱ�� := d_����ʱ�� - n_Fw;

  x_Templet := Xmltype('<OUTPUT><GROUPS></GROUPS></OUTPUT>');
  For r_File In (Select a.Id
                 From ���˻����ļ� A, �����ļ��б� B
                 Where ����id = n_Patiid And ��ҳid = n_Pageid And Ӥ�� = n_Baby And a.��ʽid = b.Id And ���� = -1 And
                       a.��ʼʱ�� > d_��ʼʱ�� And (a.����ʱ�� < d_����ʱ�� Or a.����ʱ�� Is Null)
                 Order By a.Id) Loop
    For r_Twd In (Select ID, To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��, ������
                  From ���˻�������
                  Where �ļ�id = r_File.Id And ����ʱ�� Between d_��ʼʱ�� And d_����ʱ��
                  Order By ����ʱ�� Desc) Loop
      v_Temp := '<GROUP jsonArray="True"><SJ>' || r_Twd.����ʱ�� || '</SJ><CZY>' || r_Twd.������ ||
                '</CZY><ITEMS jsonArray="True"></ITEMS></GROUP>';
      Select Appendchildxml(x_Templet, '/OUTPUT/GROUPS', Xmltype(v_Temp)) Into x_Templet From Dual;
      For r_Nr In (Select ��Ŀ���, ��Ŀ����, ��¼����, δ��˵��, ���²�λ From ���˻�����ϸ Where ��¼id = r_Twd.Id) Loop
        v_Temp := '<ITEM jsonArray="True"><XH>' || r_Nr.��Ŀ��� || '</XH><MC>' || r_Nr.��Ŀ���� || '</MC><NR>' || r_Nr.��¼���� ||
                  '</NR><WJ>' || r_Nr.δ��˵�� || '</WJ><BW>' || r_Nr.���²�λ || '</BW></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/GROUPS/GROUP/ITEMS', Xmltype(v_Temp)) Into x_Templet From Dual;
      End Loop;
    End Loop;
  End Loop;

  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Gettemphdata;
/

--122832:����,2018-03-13,�ƶ�HIS�ӿ����ӽڵ���
Create Or Replace Procedure Zl_Third_Tendfile_Getmitems
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --����:��ȡ���µ�/�����¼���еĿ��û��Ŀ
  --���:Xml_In
  --��
  -- ����:Xml_Out
  --<OUTPUT>
  -- <ITEMLIST>
  --  <ITEM>
  --    <XH/>      --���
  --    <MC/>      --����
  --    <LX/>      --��Ŀ����
  --    <BS/>      --��Ŀ��ʾ
  --    <CD/>      --��Ŀ����
  --    <XS/>      --��ĿС��
  --    <DW/>     --��Ŀ��λ
  --    <ZY/>      --��Ŀֵ��
  --    <SYBR/>      --���ò���
  --    <YYFS/>      --Ӧ�÷�ʽ
  --    <BW/>       --���Ŀ��λ
  --  <ITEM/>
  -- </ITEMLIST>
  --</OUTPUT>
  ----------------------------------------------------------------------------------------------------------------------------
  v_Temp    Varchar2(32767);
  v_Bw      Varchar2(100);
  x_Templet Xmltype; --ģ��XML
Begin

  x_Templet := Xmltype('<OUTPUT><ITEMLIST></ITEMLIST></OUTPUT>');

  For r_Hdxm In (Select a.��Ŀ���, a.��Ŀ����, a.��Ŀ����, a.��Ŀ��ʾ, a.��Ŀ����, a.��ĿС��, a.��Ŀ��λ, a.��Ŀֵ��, a.���ò���, a.Ӧ�÷�ʽ
                 From �����¼��Ŀ A
                 Where a.��Ŀ���� = 2 And Nvl(a.Ӧ�ó���, 0) <> 1) Loop
    Select f_List2str(Cast(Collect(��λ) As t_Strlist)) Into v_Bw From ���²�λ Where ��Ŀ��� = r_Hdxm.��Ŀ���;
  
    v_Temp := '<ITEM jsonArray="True"><XH>' || r_Hdxm.��Ŀ��� || '</XH><MC>' || r_Hdxm.��Ŀ���� || '</MC><LX>' || r_Hdxm.��Ŀ���� ||
              '</LX><BS>' || r_Hdxm.��Ŀ��ʾ || '</BS><CD>' || r_Hdxm.��Ŀ���� || '</CD><XS>' || r_Hdxm.��ĿС�� || '</XS><DW>' ||
              r_Hdxm.��Ŀ��λ || '</DW><ZY>' || r_Hdxm.��Ŀֵ�� || '</ZY><SYBR>' || r_Hdxm.���ò��� || '</SYBR><YYFS>' ||
              r_Hdxm.��Ŀ��ʾ || '</YYFS><BW>' || v_Bw || '</BW></ITEM>';
    Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getmitems;
/

--122832:����,2018-03-13,�ƶ�HIS�ӿ����ӽڵ���
Create Or Replace Procedure Zl_Third_Tendfile_Getdetail
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --����:��ȡָ���Ļ����¼���У�ĳһ��������Ŀ������ɴεļ�¼���ݣ���ʱ���ɽ���Զ����
  --���:Xml_In
  --<IN>
  -- <FILE></FILE>       --�ļ�id
  -- <XH></XH>   --��Ŀ���
  -- <FW></FW>   --��Χ��ֵ����3����ʾ���3��
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  -- <LISTS>
  --  <ITEM>
  --   <TIME></TIME>   --����ʱ��
  --   <DATA></DATA>   --����
  --  </ITEM>
  -- </LISTS>
  --</OUTPUT>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Fileid  Number(18);
  n_Xh      Number(18);
  n_Fw      Number(18);
  v_Temp    Varchar2(32767);
  x_Templet Xmltype; --ģ��XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/FILE'))
  Into n_Fileid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;
  Select To_Number(Extractvalue(Value(A), 'IN/XH')) Into n_Xh From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;
  Select To_Number(Extractvalue(Value(A), 'IN/FW')) Into n_Fw From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><LISTS></LISTS></OUTPUT>');
  For r_Hljl In (Select To_Char(ʱ��, 'YYYY-MM-DD hh24:mi:ss') ʱ��, ����
                 From (Select b.����ʱ�� ʱ��, Decode(c.��¼����, Null, c.δ��˵��, c.��¼����) ����,
                               Row_Number() Over(Partition By b.�ļ�id Order By b.����ʱ�� Desc) As Top
                        From ���˻������� B, ���˻�����ϸ C
                        Where b.Id = c.��¼id And ��Ŀ��� = n_Xh And �ļ�id = n_Fileid And ��¼���� = 1)
                 Where Top <= n_Fw) Loop
    v_Temp := '<ITEM jsonArray="True"><TIME>' || r_Hljl.ʱ�� || '</TIME><DATA>' || r_Hljl.���� || '</DATA></ITEM>';
    Select Appendchildxml(x_Templet, '/OUTPUT/LISTS', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;
  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getdetail;
/

--122832:����,2018-03-13,�ƶ�HIS�ӿ����ӽڵ���
Create Or Replace Procedure Zl_Third_Tendfile_Getitems
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --����:��ȡ����д�Ļ�����Ŀ�������Ŀ������Ŀ��
  --���:Xml_In
  --<IN>
  -- <BQ></BQ>       --����ID
  -- <PATIID></PATIID>     --����ID
  -- <PAGEID></PAGEID>     --��ҳID
  -- <BABY></BABY>      --Ӥ��
  -- <FILE></FILE>   --���ձ�ʾ��ȡ�������µ���Ŀ�������ȡid��Ӧ�Ļ����¼����Ŀ
  --</IN>
  --����:Xml_Out
  --<OUTPUT />
  -- <YH></YH>      --ҳ�ţ����ڰ󶨻��Ŀ
  -- <FILE></FILE>   --�ļ�id
  -- <ITEMLIST>
  --  <ITEM>
  --   <LH></LH>     --�кţ����ڰ󶨻��Ŀ
  --   <XH></XH>     --��Ŀ���
  --   <MC></MC>     --��Ŀ����
  --   <LX></LX>     --��Ŀ����0��ֵ1-�ı�
  --   <BS></BS>     --��Ŀ��ʾ
  --   <CD></CD>     --��Ŀ����
  --   <XS></XS>     --��ĿС��
  --   <DW</DW>      --��λ
  --   <ZY></ZY>     --ֵ��
  --   <SYBR></SYBR>    --���ò���0����1���˱���2Ӥ��
  --   <YYFS></YYFS>    --Ӧ�÷�ʽ0��ֹʹ��1����ʹ��2����������
  --   <BW></BW>   --��λ
  --   <XMXZ></XMXZ>  --��Ŀ����1-��ͨ2-���Ŀ,��ʾ����ĿΪԤ���Ļ��Ŀλ��
  --  </ITEM>
  -- </ITEMLIST>
  --<OUTPUT/>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Patiid  Number(18);
  n_Pageid  Number(18);
  n_Baby    Number(18);
  n_Areaid  Number(18);
  n_Fileid  Number(18);
  n_Format  Number(18);
  n_Yh      Number(18);
  v_Nnit    Varchar2(100);
  v_Hdlh    Varchar2(40);
  v_Temp    Varchar2(32767);
  v_Temp2   Varchar2(32767);
  x_Templet Xmltype; --ģ��XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/BQ'))
  Into n_Areaid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PATIID'))
  Into n_Patiid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_Pageid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/BABY'))
  Into n_Baby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/FILE'))
  Into n_Fileid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUTPUT><ITEMLIST></ITEMLIST></OUTPUT>');
  If Nvl(n_Fileid, 0) = 0 Then
    Select a.��ʽid
    Into n_Format
    From ���˻����ļ� A, �����ļ��б� B
    Where a.��ʽid = b.Id And b.���� = 3 And b.���� = -1 And a.����id = n_Patiid And a.��ҳid = n_Pageid And a.Ӥ�� = n_Baby And
          a.����ʱ�� Is Null;
    If n_Format = 30 Then
      For r_Twd In (Select b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��Ŀ��ʾ, b.��Ŀ����, b.��ĿС��, b.��Ŀ��λ, b.��Ŀֵ��, b.���ò���, b.Ӧ�÷�ʽ, b.��Ŀ����
                    From ���¼�¼��Ŀ F, �����¼��Ŀ B
                    Where f.��Ŀ��� = b.��Ŀ���) Loop
        Select f_List2str(Cast(Collect(��λ) As t_Strlist)) Into v_Nnit From ���²�λ Where ��Ŀ��� = r_Twd.��Ŀ���;
        v_Temp2 := '<ITEM jsonArray="True"><XH>' || r_Twd.��Ŀ��� || '</XH><MC>' || r_Twd.��Ŀ���� || '</MC><LX>' || r_Twd.��Ŀ���� || '</LX><BS>' ||
                   r_Twd.��Ŀ��ʾ || '</BS><CD>' || r_Twd.��Ŀ���� || '</CD><XS>' || r_Twd.��ĿС�� || '</XS><DW>' || r_Twd.��Ŀ��λ ||
                   '</DW><ZY>' || r_Twd.��Ŀֵ�� || '</ZY><SYBR>' || r_Twd.���ò��� || '</SYBR><YYFS>' || r_Twd.Ӧ�÷�ʽ ||
                   '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' || r_Twd.��Ŀ���� || '</XMXZ></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      End Loop;
    Else
      For r_Twd In (Select d.������� �к�, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��Ŀ��ʾ, b.��Ŀ����, b.��ĿС��, b.��Ŀ��λ, b.��Ŀֵ��, b.���ò���, b.Ӧ�÷�ʽ,
                           b.��Ŀ����
                    From �����ļ��ṹ C, �����ļ��ṹ D, �����¼��Ŀ B
                    Where c.�ļ�id = n_Format And c.��id Is Null And c.������� In (2, 3) And d.��id = c.Id And b.��Ŀ���� = d.Ҫ������) Loop
        Select f_List2str(Cast(Collect(��λ) As t_Strlist)) Into v_Nnit From ���²�λ Where ��Ŀ��� = r_Twd.��Ŀ���;
        v_Temp2 := '<ITEM jsonArray="True"><XH>' || r_Twd.��Ŀ��� || '</XH><MC>' || r_Twd.��Ŀ���� || '</MC><LX>' || r_Twd.��Ŀ���� || '</LX><BS>' ||
                   r_Twd.��Ŀ��ʾ || '</BS><CD>' || r_Twd.��Ŀ���� || '</CD><XS>' || r_Twd.��ĿС�� || '</XS><DW>' || r_Twd.��Ŀ��λ ||
                   '</DW><ZY>' || r_Twd.��Ŀֵ�� || '</ZY><SYBR>' || r_Twd.���ò��� || '</SYBR><YYFS>' || r_Twd.Ӧ�÷�ʽ ||
                   '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' || r_Twd.��Ŀ���� || '</XMXZ></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      End Loop;
    
    End If;
  Else
    Select Max(����ҳ��) Into n_Yh From ���˻����ӡ Where �ļ�id = n_Fileid;
    v_Temp := '<YH>' || n_Yh || '</YH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    v_Temp := ' <FILE>' || n_Fileid || '</FILE>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    For r_Jld In (Select a.ҳ��, a.�ļ�id, a.�к�, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��Ŀ��ʾ, b.��Ŀ����, b.��ĿС��, b.��Ŀ��λ, b.��Ŀֵ��, b.���ò���,
                         b.Ӧ�÷�ʽ, b.��Ŀ����
                  From ���˻�����Ŀ A, �����¼��Ŀ B
                  Where b.��Ŀ��� = a.��Ŀ��� And b.��Ŀ��� = a.��Ŀ��� And a.�ļ�id = n_Fileid And a.ҳ�� = n_Yh) Loop
    
      v_Temp2 := '<ITEM jsonArray="True"><LH>' || r_Jld.�к� || '</LH><XH>' || r_Jld.��Ŀ��� || '</XH><MC>' || r_Jld.��Ŀ���� || '</MC><LX>' ||
                 r_Jld.��Ŀ���� || '</LX><BS>' || r_Jld.��Ŀ��ʾ || '</BS><CD>' || r_Jld.��Ŀ���� || '</CD><XS>' || r_Jld.��ĿС�� ||
                 '</XS><DW>' || r_Jld.��Ŀ��λ || '</DW><ZY>' || r_Jld.��Ŀֵ�� || '</ZY><SYBR>' || r_Jld.���ò��� ||
                 '</SYBR><YYFS>' || r_Jld.Ӧ�÷�ʽ || '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' || r_Jld.��Ŀ���� ||
                 '</XMXZ></ITEM>';
      Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      v_Hdlh := v_Hdlh || ',' || r_Jld.�к�;
    End Loop;
  
    --��¼���Ѱ󶨵���Ŀ
    For r_Jldb In (Select d.������� �к�, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��Ŀ��ʾ, b.��Ŀ����, b.��ĿС��, b.��Ŀ��λ, b.��Ŀֵ��, b.���ò���, b.Ӧ�÷�ʽ,
                          b.��Ŀ����
                   From �����¼��Ŀ B, �����ļ��ṹ C, �����ļ��ṹ D, ���˻����ļ� E
                   Where c.�ļ�id = e.��ʽid And e.Id = n_Fileid And c.�����ı� = '���м���' And d.��id = c.Id And b.��Ŀ����(+) = d.Ҫ������) Loop
      Select f_List2str(Cast(Collect(��λ) As t_Strlist)) Into v_Nnit From ���²�λ Where ��Ŀ��� = r_Jldb.��Ŀ���;
      If Not Instr(v_Hdlh || ',', ',' || r_Jldb.�к� || ',') > 0 Then
        v_Temp2 := '<ITEM jsonArray="True"><LH>' || r_Jldb.�к� || '</LH><XH>' || r_Jldb.��Ŀ��� || '</XH><MC>' || r_Jldb.��Ŀ���� || '</MC><LX>' ||
                   r_Jldb.��Ŀ���� || '</LX><BS>' || r_Jldb.��Ŀ��ʾ || '</BS><CD>' || r_Jldb.��Ŀ���� || '</CD><XS>' ||
                   r_Jldb.��ĿС�� || '</XS><DW>' || r_Jldb.��Ŀ��λ || '</DW><ZY>' || r_Jldb.��Ŀֵ�� || '</ZY><SYBR>' ||
                   r_Jldb.���ò��� || '</SYBR><YYFS>' || r_Jldb.Ӧ�÷�ʽ || '</YYFS><BW>' || v_Nnit || '</BW><XMXZ>' ||
                   r_Jldb.��Ŀ���� || '</XMXZ></ITEM>';
        Select Appendchildxml(x_Templet, '/OUTPUT/ITEMLIST', Xmltype(v_Temp2)) Into x_Templet From Dual;
      End If;
    End Loop;
  
  End If;
  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getitems;
/

--122832:����,2018-03-13,�ƶ�HIS�ӿ����ӽڵ���
Create Or Replace Procedure Zl_Third_Tendfile_Getall
(
  Xmlfilelist_In  Xmltype,
  Xmlfilelist_Out Out Xmltype
) Is
  ---------------------------------------------------------------------------------------------------------------------------
  --����:��ȡ��ѡ���˵�ǰ�Ѵ����Ļ����¼���Ϳɴ����Ļ����¼���б�
  --���:Xml_In
  --<IN>
  -- <BQ></BQ>        --����ID
  -- <PATIID></PATIID>      --����ID
  -- <PAGEID></PAGEID>      --��ҳID
  -- <BABY></BABY>       --Ӥ��
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  -- <ITEMLIST>
  --  <ITEM>
  --   <ID></ID>   --�Ѵ�����Ϊ�ļ�ID��δ������Ϊ��ʽID
  --   <MC></MC>
  --   <TYPE></TYPE>   --0��ʾδ������1��ʾ�Ѵ���
  --  </ITEM>
  -- </ITEMLIST>
  --<OUTPUT/>
  ----------------------------------------------------------------------------------------------------------------------------
  n_Patiid  Number(18);
  n_Pageid  Number(18);
  n_Baby    Number(18);
  n_Areaid  Number(18);
  v_Temp    Varchar2(32767);
  x_Templet Xmltype; --ģ��XML
Begin
  Select To_Number(Extractvalue(Value(A), 'IN/BQ'))
  Into n_Areaid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PATIID'))
  Into n_Patiid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_Pageid
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  Select To_Number(Extractvalue(Value(A), 'IN/BABY'))
  Into n_Baby
  From Table(Xmlsequence(Extract(Xmlfilelist_In, 'IN'))) A;

  x_Templet := Xmltype('<OUT><ITEMLIST></ITEMLIST></OUT>');

  For r_Ycj In (Select a.Id, a.��ʽid, a.����id, c.���� As ����, a.�ļ�����, a.��ʼʱ��, a.����ʱ��, b.����, b.���
                From ���˻����ļ� A, �����ļ��б� B, ���ű� C
                Where a.��ʽid = b.Id And a.����id = c.Id And a.����id = n_Patiid And a.��ҳid = n_Pageid And a.Ӥ�� = n_Baby
                Order By b.����, a.��ʼʱ��) Loop
    v_Temp := '<ITEM jsonArray="True"><ID>' || r_Ycj.Id || '</ID><MC>' || r_Ycj.�ļ����� || '</MC><TYPE>1</TYPE></ITEM>';
    Select Appendchildxml(x_Templet, '/OUT/ITEMLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  For r_Kcj In (Select ID, ����, ���, ��ʽ
                From (Select ID, ����, ���, ���� As ��ʽ
                       From �����ļ��б�
                       Where ���� = 3 And ���� <> 1 And
                             (ͨ�� = 1 Or (ͨ�� = 2 And ID In (Select �ļ�id From ����Ӧ�ÿ��� Where ����id = n_Areaid))))
                Order By ����, ���) Loop
    v_Temp := '<ITEM jsonArray="True"><ID>' || r_Kcj.Id || '</ID><MC>' || r_Kcj.��ʽ || '</MC><TYPE>0</TYPE></ITEM>';
    Select Appendchildxml(x_Templet, '/OUT/ITEMLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
  End Loop;

  Xmlfilelist_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Tendfile_Getall;
/

--122763:Ƚ����,2018-03-12,XMLѭ���ڵ����� jsonArray ����
Create Or Replace Procedure Zl_Third_Getexessort
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:��ȡסԺ���˷��÷������ 
  --���:Xml_In 
  -- <IN> 
  --  <PATIID></PATIID>         --����ID 
  --  <PAGEID></PAGEID>     --��ҳID 
  --</IN> 
  --����:Xml_Out 
  --<OUTPUT> 
  --  <ZFY></ZFY>    --�ܷ��� 
  --  <FYLIST> 
  --    <ITEM> 
  --      <XM></XM> --�վݷ�Ŀ 
  --      <JE></JE> --��� 
  --    </ITEM> 
  --  <FYLIST> 
  --</OUTPUT> 
  -------------------------------------------------------------------------------------------------- 
  n_����id ������Ϣ.����id%Type;
  n_��ҳid ������ҳ.��ҳid%Type;
  n_�ܷ��� סԺ���ü�¼.ʵ�ս��%Type;

  x_Templet Xmltype; --ģ��XML 
Begin
  --��ȡ��� 
  Select Extractvalue(Value(A), 'IN/PATIID'),
         Decode(Extractvalue(Value(A), 'IN/PAGEID'), 0, Null, Extractvalue(Value(A), 'IN/PAGEID'))
  Into n_����id, n_��ҳid
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Nvl(Sum(a.ʵ�ս��), 0)
  Into n_�ܷ���
  From סԺ���ü�¼ A
  Where a.����id = n_����id And a.��ҳid = n_��ҳid And Nvl(a.�����־, 0) = 2;

  Select Xmlelement("OUTPUT",
                     Xmlforest(n_�ܷ��� As "ZFY",
                                Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                   Xmlforest(a.�վݷ�Ŀ As "XM", Nvl(Sum(a.ʵ�ս��), 0) As "JE"))) As "FYLIST"))
  Into x_Templet
  From סԺ���ü�¼ A
  Where a.����id = n_����id And a.��ҳid = n_��ҳid And Nvl(a.�����־, 0) = 2
  Group By a.�վݷ�Ŀ;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getexessort;
/

--122714:��˼��,2018-03-12,Zl_����ҽ��ִ��_�ܾ�ִ��  �����߼�����
Create Or Replace Procedure Zl_����ҽ��ִ��_�ܾ�ִ��
(
  ҽ��id_In     In ����ҽ��ִ��.ҽ��id%Type,
  ���ͺ�_In     In ����ҽ��ִ��.���ͺ�%Type,
  ����Ա���_In In ��Ա��.���%Type := Null,
  ����Ա����_In In ��Ա��.����%Type := Null,
  ִ�в���id_In In ������ü�¼.ִ�в���id%Type := 0,
  �ܾ�ԭ��_In   In ����ҽ������.ִ��˵��%Type := Null
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
) Is
  Cursor c_Advice Is
    Select a.Id, a.���id, a.�������, a.����id, a.��ҳid, a.�Һŵ�, b.No, a.������Դ
    From ����ҽ����¼ A, ����ҽ������ B
    Where ID = ҽ��id_In And a.Id = b.ҽ��id;
  r_Advice c_Advice%RowType;

  n_Temp     Number;
  v_Temp     Varchar2(255);
  v_��Ա���� ��Ա��.����%Type;

Begin
  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;

  If r_Advice.������� = 'C' And r_Advice.���id Is Not Null Then
    --����һ���ɼ������м�����Ŀ
    Update ����ҽ������
    Set ִ��״̬ = 2, ����� = v_��Ա����, ���ʱ�� = Sysdate, ִ��˵�� = �ܾ�ԭ��_In
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id In (Select ID From ����ҽ����¼ Where ���id = r_Advice.���id);
  Else
    --������������,���鲿λ,�Լ���������ҽ��;�������ҩ�巨�ǵ�������
    Update ����ҽ������
    Set ִ��״̬ = 2, ����� = v_��Ա����, ���ʱ�� = Sysdate, ִ��˵�� = �ܾ�ԭ��_In
    Where ���ͺ� + 0 = ���ͺ�_In And ҽ��id In (Select ID
                                        From ����ҽ����¼
                                        Where ID = ҽ��id_In
                                        Union All
                                        Select ID
                                        From ����ҽ����¼
                                        Where ���id = ҽ��id_In And ������� In ('F', 'D'));
  End If;
  If r_Advice.������� = 'D' Then
    Select Count(1) Into n_Temp From ��������˵�� Where ����id = ִ�в���id_In And �������� = '���';
    If n_Temp > 0 Then
      b_Message.Zlhis_Cis_037(r_Advice.����id, r_Advice.��ҳid, r_Advice.�Һŵ�, ���ͺ�_In, ҽ��id_In, r_Advice.No, r_Advice.������Դ);
    End If;
  End If;
  Close c_Advice;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ��ִ��_�ܾ�ִ��;
/

--122731:��˶,2018-03-09,�ƶ�����ӿ�ѭ���ڵ��������
Create Or Replace Procedure Zl_Third_Getfeeitem
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ�շ���ĿĿ¼,���������û�����Ŀ��Ӧ�շ���Ŀʱ��������롢���Ʋ�ѯHIS�շ���Ŀ
  --���:Xml_In:
  --<IN>
  --  <KEY></KEY>   //��ѯ�ؼ��֣����롢���ơ����룬����Ϊƴ�����룬������ƥ�䣬���������ȫƥ�䡣ΪNULL�򲻽���ƥ���ѯ
  --  <PAGENOW></PAGENOW>  //��ǰҳ������PAGESIZE��PAGENOWΪ�ջ�<1,�򷵻��������ݡ����򷵻�ָ��ҳ��������
  --  <PAGESIZE></PAGESIZE>  //��¼��������PAGESIZE��PAGENOWΪ�ջ�<1,�򷵻��������ݡ����򷵻�ָ��ҳ��������
  --</IN>
  --����:Xml_Out--�������������򷵻ط�ҳ
  --<OUTPUT>
  --  <XMLIST>
  --    <XM jsonArray="true">
  --      <LB></LB>   //������ơ������
  --      <ID></ID>   //�շ���ĿId
  --      <BM></BM>   //�շ���Ŀ����
  --      <MC></MC>   //�շ���Ŀ����
  --      <GG></GG>   //���
  --      <DW></DW>   //��λ
  --      <DJ></DJ>   //����
  --      <SM></SM>   //˵��
  --    </XM>
  --  </XMLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_Key      �շ���ĿĿ¼.����%Type;
  n_Cur_Page Number(5);
  n_Pagesize Number(5);
  x_Templet  Xmltype; --ģ��XML
Begin
  --��ȡ��ѯ����
  Select Max(b.Key), Max(b.Pagenow), Max(Pagesize)
  Into v_Key, n_Cur_Page, n_Pagesize
  From Xmltable('$a/IN' Passing Xml_In As "a" Columns Key Varchar2(100) Path 'KEY', Pagenow Number(5) Path 'PAGENOW',
                 Pagesize Number(5) Path 'PAGESIZE') B;

  --��ȡ���е����ݣ���ƥ��
  --��ѯSQL��Դ��������Ŀ�����������շ���Ŀʱ����ƥ�䣬�䶯��Ϊ����Sum�޸�ΪMax��
  If v_Key Is Null Then
    --�����з�ҳ��ֱ�ӷ�����������
    If Nvl(n_Cur_Page, 0) < 1 Or Nvl(n_Pagesize, 0) < 1 Then
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.������� As "LB", e.Id As "ID", e.���� As "BM", e.���� As "MC",
                                                                   e.��� As "GG", e.���㵥λ As "DW", e.�ۼ� As "DJ", e.˵�� As "SM"))))) ��������
      Into x_Templet
      From (Select b.���� �������, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.˵��,
                    Decode(Nvl(a.�Ƿ���, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.�ּ�, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', a.���), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.ȱʡ�۸�, 0)), '9999999990.0000'))),
                                    'ʱ��')) As �ۼ�
             From �շ���ĿĿ¼ A, �շ���Ŀ��� B, �շѼ�Ŀ D
             Where a.��� = b.���� And a.Id = d.�շ�ϸĿid(+) And a.��� Not In ('1', 'J') And
                   (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                   (a.������� = 1 Or a.������� = 2 Or a.������� = 3) And d.ִ������ <= Sysdate And
                   (d.��ֹ���� > Sysdate Or d.��ֹ���� Is Null) And d.�۸�ȼ� Is Null
             Group By b.����, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.˵��, a.�Ƿ���, a.���) E;
      --��ҳ��ѯ
    Else
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.������� As "LB", e.Id As "ID", e.���� As "BM", e.���� As "MC",
                                                                   e.��� As "GG", e.���㵥λ As "DW", e.�ۼ� As "DJ", e.˵�� As "SM"))))) ��������
      Into x_Templet
      From (Select b.���� �������, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.˵��,
                    Decode(Nvl(a.�Ƿ���, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.�ּ�, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', a.���), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.ȱʡ�۸�, 0)), '9999999990.0000'))),
                                    'ʱ��')) As �ۼ�, Row_Number() Over(Order By b.����, a.����) As Rn
             From �շ���ĿĿ¼ A, �շ���Ŀ��� B, �շѼ�Ŀ D
             Where a.��� = b.���� And a.Id = d.�շ�ϸĿid(+) And a.��� Not In ('1', 'J') And
                   (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                   (a.������� = 1 Or a.������� = 2 Or a.������� = 3) And d.ִ������ <= Sysdate And
                   (d.��ֹ���� > Sysdate Or d.��ֹ���� Is Null) And d.�۸�ȼ� Is Null
             Group By b.����, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.˵��, a.�Ƿ���, a.���) E
      Where Rn Between n_Pagesize * (n_Cur_Page - 1) + 1 And n_Pagesize * n_Cur_Page;
    End If;
    --��ȡָ����ƥ������
  Else
    --�����з�ҳ��ֱ�ӷ�����������
    If Nvl(n_Cur_Page, 0) < 1 Or Nvl(n_Pagesize, 0) < 1 Then
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.������� As "LB", e.Id As "ID", e.���� As "BM", e.���� As "MC",
                                                                   e.��� As "GG", e.���㵥λ As "DW", e.�ۼ� As "DJ", e.˵�� As "SM"))))) ��������
      Into x_Templet
      From (Select f.�������, f.Id, f.����, f.����, f.���, f.����, f.���㵥λ, f.˵��,
                    Decode(Nvl(f.�Ƿ���, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.�ּ�, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', f.���), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.ȱʡ�۸�, 0)), '9999999990.0000'))),
                                    'ʱ��')) As �ۼ�
             From (Select Distinct b.���� �������, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.˵��, a.�Ƿ���, a.���
                    
                    From �շ���ĿĿ¼ A, �շ���Ŀ��� B, �շ���Ŀ���� C
                    Where a.��� = b.���� And a.Id = c.�շ�ϸĿid And a.��� Not In ('1', 'J') And
                          (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                          (a.������� = 1 Or a.������� = 2 Or a.������� = 3) And
                          (a.���� Like v_Key || '%' Or c.���� Like '%' || v_Key || '%' Or c.���� Like '%' || v_Key || '%') And
                          c.���� = '1') F,
                  
                  �շѼ�Ŀ D
             Where f.Id = d.�շ�ϸĿid(+) And d.�۸�ȼ� Is Null And d.ִ������ <= Sysdate And (d.��ֹ���� > Sysdate Or d.��ֹ���� Is Null)
             Group By f.����, f.Id, f.����, f.����, f.���, f.����, f.���㵥λ, f.˵��, f.�Ƿ���, f.�������, f.���) E;
      --��ҳ��ѯ
    Else
      Select Xmlelement("OUTPUT",
                         Xmlelement("XMLIST",
                                     Xmlagg(Xmlelement("XM", Xmlattributes('true' As "jsonArray"),
                                                        Xmlforest(e.������� As "LB", e.Id As "ID", e.���� As "BM", e.���� As "MC",
                                                                   e.��� As "GG", e.���㵥λ As "DW", e.�ۼ� As "DJ", e.˵�� As "SM"))))) ��������
      Into x_Templet
      From (Select f.�������, f.Id, f.����, f.����, f.���, f.����, f.���㵥λ, f.˵��,
                    Decode(Nvl(f.�Ƿ���, 0), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.�ּ�, 0)), '9999999990.0000'))),
                            Decode(Instr('4567', f.���), 0, LTrim(RTrim(To_Char(Sum(Nvl(d.ȱʡ�۸�, 0)), '9999999990.0000'))),
                                    'ʱ��')) As �ۼ�, Row_Number() Over(Order By f.����, f.����) As Rn
             From (Select Distinct b.���� �������, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.˵��, a.�Ƿ���, a.���
                    From �շ���ĿĿ¼ A, �շ���Ŀ��� B, �շ���Ŀ���� C
                    Where a.��� = b.���� And a.Id = c.�շ�ϸĿid And a.��� Not In ('1', 'J') And
                          (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                          (a.������� = 1 Or a.������� = 2 Or a.������� = 3) And
                          (a.���� Like v_Key || '%' Or c.���� Like '%' || v_Key || '%' Or c.���� Like '%' || v_Key || '%') And
                          c.���� = '1') F,
                  
                  �շѼ�Ŀ D
             Where f.Id = d.�շ�ϸĿid(+) And d.�۸�ȼ� Is Null And d.ִ������ <= Sysdate And (d.��ֹ���� > Sysdate Or d.��ֹ���� Is Null)
             Group By f.����, f.Id, f.����, f.����, f.���, f.����, f.���㵥λ, f.˵��, f.�Ƿ���, f.�������, f.���) E
      Where Rn Between n_Pagesize * (n_Cur_Page - 1) + 1 And n_Pagesize * n_Cur_Page;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getfeeitem;
/

--122731:��˶,2018-03-09,�ƶ�����ӿ�ѭ���ڵ��������
Create Or Replace Procedure Zl_Third_Getdeptmatch
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:���ڻ�ȡ���Ҳ������չ�ϵ
  --���:Xml_In:
  --<IN>
  --  <BMID></BMID>         --����ID,����ʱΪ��ȡ���ж��չ�ϵ
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  --  <BMKSLIST>
  --    <ITEM jsonArray="true">
  --      <BQID></BQID>  --����ID
  --      <KSID></KSID>  --����ID
  --    </ITEM>
  --  <BMKSLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_����id  ���ű�.Id%Type;
  x_Templet Xmltype; --ģ��XML
Begin
  --��ȡ����ID
  Select Max(b.Bmid) Into n_����id From Xmltable('$a/IN' Passing Xml_In As "a" Columns Bmid Number(18) Path 'BMID') B;

  --��ȡ���ж�Ӧ��ϵ
  If n_����id Is Null Then
  
    Select Xmlelement("OUTPUT",
                       Xmlelement("BMKSLIS",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.����id As "BQID", a.����id As "KSID"))))) ��������
    
    Into x_Templet
    From �������Ҷ�Ӧ A;
    --��ȡָ�����Ҷ�Ӧ�Ĳ���
  Else
    Select Xmlelement("OUTPUT",
                       Xmlelement("BMKSLIS",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.����id As "BQID", a.����id As "KSID"))))) ��������
    
    Into x_Templet
    From �������Ҷ�Ӧ A
    Where a.����id = n_����id;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdeptmatch;
/

--122731:��˶,2018-03-09,�ƶ�����ӿ�ѭ���ڵ��������
Create Or Replace Procedure Zl_Third_Getdept(Xml_Out Out Xmltype) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ������Ϣ
  --���:��
  --����:Xml_Out
  --<OUTPUT>
  --  <BMLIST>
  --    <ITEM jsonArray=��true��>
  --      <ID></ID>  --ID
  --      <MC></MC>  --����
  --      <BH></BH>  --���
  --      <JM></JM>  --����
  --      <ZD></ZD>  --վ��
  --      <XZ></XZ>  --���ʣ���������á�,���ŷָ�
  --    </ITEM>
  --  <BMLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  x_Templet Xmltype; --ģ��XML
Begin
  --��ȡ���в���
  Select Xmlelement("OUTPUT",
                     Xmlelement("BMLIST",
                                 Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                    Xmlforest(a.Id As "ID", Max(a.����) As "MC", Max(a.����) As "BH",
                                                               Max(a.����) As "JM", Max(a.վ��) As "ZD",
                                                               f_List2str(Cast(Collect(b.��������) As t_Strlist)) As "XZ"))))) ��������

  
  Into x_Templet
  From ���ű� A, ��������˵�� B
  Where a.Id = b.����id
  Group By a.Id;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getdept;
/

--122731:��˶,2018-03-09,�ƶ�����ӿ�ѭ���ڵ���������
Create Or Replace Procedure Zl_Third_Getperson
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ��Ա��Ϣ
  --���:Xml_In:
  --<IN>
  --  <BMID></BMID>         --����ID,����ʱΪ��ȡ������Ա
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  --  <RYLIST>
  --    <ITEM jsonArray=��true��>
  --      <ID></ID>  --ID
  --      <XM></XM>  --����
  --      <BH></BH>  --���
  --      <JM></JM>  --����
  --      <XB></XB>  --�Ա�
  --      <CSRQ></CSRQ> --��������
  --      <SFZH></SFZH> --���֤��
  --      <MZ></MZ>  --����
  --      <XL></XL>  --ѧ��
  --      <ZYJSZW></ZYJSZW> --רҵ����ְ��
  --      <XZ></XZ>  --��Ա���ʣ��ַ�����ҽ��,��ʿ,������
  --      <BMLIST>  --���������б�
  --        <ITEM jsonArray=��true��>
  --          <ID></ID> --����ID
  --          <MC></MC> --��������
  --        </ITEM>
  --      </BMLIST>
  --    </ITEM>
  --  <RYLIST>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  n_����id  ���ű�.Id%Type;
  x_Templet Xmltype; --ģ��XML
Begin
  --��ȡ����ID
  Select Max(b.Bmid) Into n_����id From Xmltable('$a/IN' Passing Xml_In As "a" Columns Bmid Number(18) Path 'BMID') B;
  --��ȡ���в�����Ա
  If n_����id Is Null Then
    Select Xmlelement("OUTPUT",
                       Xmlelement("RYLIST",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.Id As "ID", Max(a.����) As "XM", Max(a.���) As "BH",
                                                                 Max(a.����) As "JM", Max(a.�Ա�) As "XB",
                                                                 To_Char(Max(a.��������), 'YYYY-MM-DD HH24:MI:SS') As "CSRQ",
                                                                 Max(a.���֤��) As "SFZH", Max(a.����) As "MZ", Max(a.ѧ��) As "XL",
                                                                 Max(a.רҵ����ְ��) As "ZYJSZW", Max(b.��Ա����) As "XZ",
                                                                 Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                                                    Xmlforest(g.Id As "ID", g.���� As "MC"))) As
                                                                  "BMLIST")))))
    Into x_Templet
    From ��Ա�� A,
         (Select e.��Աid, f_List2str(Cast(Collect(e.��Ա����) As t_Strlist)) ��Ա����
           From (Select c.Id ��Աid, d.��Ա����
                  From ��Ա�� C, ��Ա����˵�� D
                  Where d.��Աid = c.Id And d.��Ա���� In ('ҽ��', '��ʿ')
                  Union All
                  Select c.Id ��Աid, '����' ��Ա����
                  From ��Ա�� C, ��Ա����˵�� D
                  Where d.��Աid = c.Id And d.��Ա���� Not In ('ҽ��', '��ʿ')
                  Group By c.Id) E
           Group By e.��Աid) B, ������Ա F, ���ű� G
    Where a.Id = b.��Աid(+) And a.Id = f.��Աid(+) And f.����id = g.Id(+)
    Group By a.Id;
    --��ȡָ��������Ա
  Else
    With People As
     (Select r.Id, r.����, r.���, r.����, r.�Ա�, r.��������, r.���֤��, r.����, r.ѧ��, r.רҵ����ְ��
      From ��Ա�� R
      Where r.Id In (Select ��Աid From ������Ա H Where h.����id = n_����id))
    Select Xmlelement("OUTPUT",
                       Xmlelement("RYLIST",
                                   Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                      Xmlforest(a.Id As "ID", Max(a.����) As "XM", Max(a.���) As "BH",
                                                                 Max(a.����) As "JM", Max(a.�Ա�) As "XB",
                                                                 To_Char(Max(a.��������), 'YYYY-MM-DD HH24:MI:SS') As "CSRQ",
                                                                 Max(a.���֤��) As "SFZH", Max(a.����) As "MZ", Max(a.ѧ��) As "XL",
                                                                 Max(a.רҵ����ְ��) As "ZYJSZW", Max(b.��Ա����) As "XZ",
                                                                 Xmlagg(Xmlelement("ITEM", Xmlattributes('true' As "jsonArray"),
                                                                                    Xmlforest(g.Id As "ID", g.���� As "MC"))) As
                                                                  "BMLIST")))))
    Into x_Templet
    From People A,
         (Select e.��Աid, f_List2str(Cast(Collect(e.��Ա����) As t_Strlist)) ��Ա����
           From (Select c.Id ��Աid, d.��Ա����
                  From People C, ��Ա����˵�� D
                  Where d.��Աid = c.Id And d.��Ա���� In ('ҽ��', '��ʿ')
                  Union All
                  Select c.Id ��Աid, '����' ��Ա����
                  From People C, ��Ա����˵�� D
                  Where d.��Աid = c.Id And d.��Ա���� Not In ('ҽ��', '��ʿ')
                  Group By c.Id) E
           Group By e.��Աid) B, ������Ա F, ���ű� G
    Where a.Id = b.��Աid(+) And a.Id = f.��Աid(+) And f.����id = g.Id(+)
    Group By a.Id;
  End If;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getperson;
/







------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.35.90.0002' Where ���=&n_System;
Commit;
