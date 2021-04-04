Create Or Replace Procedure Zl_Patisvr_Batupdoutpativisit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�������²��˵ľ���״̬�;�������
  --��Σ�Json_In:��ʽ
  --    input
  --      visit_status      N 1 ����״̬
  --      pati_list[]      ����
  --        pati_ids       C 1 ����id,�����","�ָ�
  --        visit_room     C 1 ����

  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  j_Jsonin   Pljson;
  j_Json_Tmp Pljson;
  j_List     Pljson_List := Pljson_List();
  n_״̬     ������Ϣ.����״̬%Type;
  v_����ids  Varchar2(3000);
  v_����     ������Ϣ.��������%Type;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_״̬   := j_Json.Get_Number('visit_status');
  If n_״̬ Is Null Then
    Json_Out := Zljsonout('δ�������״̬�����飡');
    Return;
  End If;
  j_List := j_Json.Get_Pljson_List('pati_list');
  If j_List Is Null Then
    Json_Out := Zljsonout('δ���벡��id�����ң����飡');
    Return;
  End If;
  For I In 1 .. j_List.Count Loop
    j_Json_Tmp := Pljson();
    j_Json_Tmp := Pljson(j_List.Get(I));
    v_����ids  := j_Json_Tmp.Get_String('pati_ids');
    v_����     := j_Json_Tmp.Get_String('visit_room');
    Update ������Ϣ
    Set ����״̬ = n_״̬, �������� = v_����
    Where ����id In (Select /*+cardinality(x,10)*/
                    x.Column_Value As ����id
                   From Table(Cast(f_Num2list(v_����ids) As Zltools.t_Numlist)) X) And ����״̬ In (1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Batupdoutpativisit;
/
Create Or Replace Procedure Zl_Patisvr_Calc_Age
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --����:���ݳ������ڼ�������.����Ǽǲ���,�������䲻��.
  --��Σ�Json_In,��ʽ����
  --  input
  --    pati_id            N 1 ����ID
  --    birthdate          C 1 ��������
  --    calc_date          C 1 ��������
  --����: Json_Out,��ʽ����
  --  output
  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    age                C 1  ����:1�����ڣ�XСʱ[X����],1����1�����ڣ�X��[XСʱ],1����1�����ڣ�X��[X��],1������ͯ�������ޣ�X��[X��],>=��ͯ�������ޣ�X��
  --                            ˵��:1�����ڣ���ָ����������24Сʱ��;1�����ڣ���ָ������㣻����7.8�ճ�����8.8�ղ���1��;1�����ڣ�Ҳ�Ƕ�����㡣;�����ڡ�����ָ��<����
  --------------------------------------------------------------------------------------------------------------------
  j_Json      Pljson;
  j_Jsonin    Pljson;
  n_Pati_Id   ������Ϣ.����id%Type;
  d_Birthdate Date;
  d_Calc_Date Date;

  v_Age Varchar2(20); --���ڲ�����Ϣ����ر�������ֶ�Ϊ10���ַ��������������10���ַ���5������
Begin
  --�������
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_Pati_Id := j_Json.Get_Number('pati_id');

  d_Birthdate := To_Date(j_Json.Get_String('birthdate'), 'YYYY-MM-DD HH24:MI:SS');
  d_Calc_Date := To_Date(j_Json.Get_String('calc_date'), 'YYYY-MM-DD HH24:MI:SS');
  v_Age       := Zl_Age_Calc(n_Pati_Id, d_Birthdate, d_Calc_Date);
  Json_Out    := '{"output":{"code":1,"message":"�ɹ�","age":"' || Zljsonstr(v_Age) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Calc_Age;
/
Create Or Replace Procedure Zl_Patisvr_Checkcardexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��鿨����Ƿ����
  --��Σ�Json_In:��ʽ
  --    input
  --      card_type_id      N 1 �����ID
  --      pati_id           N 1 ����id
  --      card_no           C 1 ҽ�ƿ���

  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_id           N   1   ��ǰʹ�����ſ��Ĳ���id����Դ��뿨��ʱ��Ч
  --    exist             N   1   ��ǰ�����Ѿ�����ͬ���͵�ҽ�ƿ�����Դ��벡��idʱ��Ч
  ---------------------------------------------------------------------------

  j_Json       Pljson;
  j_Jsonin     Pljson;
  n_����id     ����ҽ�ƿ���Ϣ.����id%Type;
  n_�����id   ����ҽ�ƿ���Ϣ.�����id%Type;
  v_����       ����ҽ�ƿ���Ϣ.����%Type;
  n_����id_Out ����ҽ�ƿ���Ϣ.����id%Type;
  n_Exist      Number(2);

Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_�����id := j_Json.Get_Number('card_type_id');
  v_����     := j_Json.Get_String('card_no');

  If Nvl(n_�����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���뿨���id�����飡');
    Return;
  End If;

  If v_���� Is Not Null Then
    Select Nvl(Max(����id), 0) Into n_����id_Out From ����ҽ�ƿ���Ϣ Where ���� = v_���� And �����id = n_�����id;
  Else
    Select Count(1)
    Into n_Exist
    From ����ҽ�ƿ���Ϣ
    Where ����id = Nvl(n_����id, 0) And �����id = n_�����id And Nvl(״̬, 0) = 0 And Rownum < 2;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exist":' || Nvl(n_Exist, 0) || ',"pati_id":' || Nvl(n_����id_Out, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkcardexist;
/
Create Or Replace Procedure Zl_Patisvr_Checkdepositerrorno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݵ��ݺŻ�ȡ���ڲ��˽����쳣��¼�е�NO
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id            N 1 ����id
  --   bill_nos           C 1 ����Ԥ����¼.NO,����ö��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    bill_nos          C 1 ��Ч��Nos,����ö��ŷָ�
  --    occasion          N 1 ���ϣ�1-ҽ�ƿ�����;2-������Ϣ�Ǽǣ����ֻ��һ��NO����Ч��
  ---------------------------------------------------------------------------
  n_����id  ���˽����쳣��¼.����id%Type;
  v_Nos     Varchar2(3000);
  j_Json    Pljson;
  j_Jsonin  Pljson;
  v_Nos_Out Varchar2(3000);
  n_����    ���˽����쳣��¼.��������%Type;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_Nos    := j_Json.Get_String('bill_nos');

  If Nvl(v_Nos, '-') = '-' Then
    Json_Out := Zljsonout('δ����NO�����飡');
    Return;
  End If;

  Select /*+cardinality(B,10)*/
   f_List2str(Cast(Collect(b.Column_Value) As t_Strlist)), Nvl(Max(a.��������), 0)
  Into v_Nos_Out, n_����
  From ���˽����쳣��¼ A, Table(f_Str2list(v_Nos)) B
  Where a.�������� In (1, 2) And (a.Ԥ������ = b.Column_Value Or a.ҽ�ƿ����� = b.Column_Value) And
        Decode(n_����id, 0, 0, a.����id) = n_����id;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","bill_nos":"' || v_Nos_Out || '","occasion":' || n_���� || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkdepositerrorno;
/
Create Or Replace Procedure Zl_Patisvr_Checkidcardunique
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID�����֤�ż��ͬһ���ֻ֤�ܶ�Ӧһ����������
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_idcard           C   1  ���֤��
  --    pati_id              N   1  ����ID
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist              N   1   �Ƿ���ڣ�0-������ 1-���ڣ�
  ---------------------------------------------------------------------------
  n_����id   ������Ϣ.����id%Type;
  v_���֤�� ������Ϣ.���֤��%Type;
  n_Count    Number;
  j_Json     Pljson;
  j_Jsonin   Pljson;

  n_Isexist Number(1);
Begin

  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_���֤�� := j_Json.Get_String('pati_idcard');
  n_����id   := j_Json.Get_Number('pati_id');
  Select Count(1) Into n_Count From ������Ϣ A Where a.���֤�� = v_���֤�� And a.����id <> n_����id And Rownum < 2;
  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkidcardunique;
/
Create Or Replace Procedure Zl_Patisvr_Checkinsnoisexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ����ҽ�����Ƿ����
  --��Σ�Json_In:��ʽ
  --  input
  --    insurance_num        C   1  ҽ����
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist              N   1  1-����;0-������
  v_ҽ���� ������Ϣ.ҽ����%Type;
  n_Exist  Number(2);
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin

  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  v_ҽ���� := j_Json.Get_String('insurance_num');
  If Nvl(v_ҽ����, '-') = '-' Then
    Json_Out := Zljsonout('δ���벡��ҽ����');
    Return;
  End If;
  Select Count(1) Into n_Exist From ������Ϣ Where ҽ���� = v_ҽ����;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkinsnoisexist;

/
Create Or Replace Procedure Zl_Patisvr_Checkoutnoisexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ����������Ƿ��Ѿ���ʹ��
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id              N   1  ����ID:��ǰ�����Ĳ���
  --    outpatient_num       C   1  �����
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist              N   1  1-����;0-������
  ---------------------------------------------------------------------------
  n_����id ������Ϣ.����id%Type;
  n_����� ������Ϣ.�����%Type;
  n_Exist  Number(2);
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin

  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_����� := To_Number(j_Json.Get_String('outpatient_num'));
  Select Count(1) Into n_Exist From ������Ϣ Where ����� = n_����� And ����id <> n_����id;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Exist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkoutnoisexist;
/
Create Or Replace Procedure Zl_Patisvr_Checkpatirealname
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID,�жϸò����Ƿ������ʵ����֤
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id              N   1  ����ID
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist              N   1   �Ƿ���ڣ�0-������ 1-���ڣ�
  ---------------------------------------------------------------------------
  n_Count   Number;
  n_Isexist Number;
  j_Json    Pljson;
  j_Jsonin  Pljson;
  n_����id  Number;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  Select Count(1) Into n_Count From ����ʵ����Ϣ Where ����id = n_����id And Rownum < 2;
  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Isexist || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkpatirealname;
/
Create Or Replace Procedure Zl_Patisvr_Checkregisterinpati
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��Ժ�ǼǼ��
  --��Σ�Json_In:��ʽ
  --   input
  --      type              N 1 ��������  1-�����Ǽ�;2-�޸ĵǼ�
  --      pati_id           N 1 ����id
  --      pati_idcard       N 1 ���֤��
  --      isnew             N  1-�²���;0-���²���
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json        Pljson;
  j_Jsonin      Pljson;
  n_Type        Number(5);
  n_Pati_Id     Number(18);
  v_Pati_Idcard Varchar2(18);
  n_Isnew       Number(1);
  n_Count       Number(5);
  n_Uniqueid    Number(1);
  v_Msg         Varchar2(200);
Begin
  --�������
  j_Jsonin      := Pljson(Json_In);
  j_Json        := j_Jsonin.Get_Pljson('input');
  n_Type        := j_Json.Get_Number('type');
  n_Pati_Id     := j_Json.Get_Number('pati_id');
  v_Pati_Idcard := j_Json.Get_String('pati_idcard');
  n_Isnew       := j_Json.Get_Number('isnew');

  --�жϲ����Ƿ�����
  Select Count(����id) Into n_Count From ������Ϣ Where ����id = n_Pati_Id;

  If n_Count <> 0 Then
    Zl_������Ϣ_�������(n_Pati_Id);
  End If;

  --���֤�Ų����ڿ�,����ϵͳ�����ж��Ƿ�Ψһ��������
  If v_Pati_Idcard Is Not Null And ((n_Isnew = 1 And n_Type = 1) Or n_Type = 2) Then
    n_Uniqueid := Nvl(zl_GetSysParameter(279), 0);
    If n_Uniqueid = 1 Then
      Select Count(1) Into n_Count From ������Ϣ Where ���֤�� = v_Pati_Idcard And ����id <> Nvl(n_Pati_Id, 0);
      If n_Count <> 0 Then
        v_Msg    := '�Ѿ��������֤��Ϊ' || v_Pati_Idcard || '�Ĳ���,������¼����ͬ�����֤��!';
        Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Msg) || '"}}';
        Return;
      End If;
    End If;
  End If;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkregisterinpati;
/
Create Or Replace Procedure Zl_Patisvr_Checkreturncard
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�жϵ�ǰ���Ƿ������˿�������������˿���������ʾ���ݣ����򷵻�NULL
  --��Σ�Json_In:��ʽ
  -- input
  --   occasion             N 1 ���ϣ�1-ҽ�ƿ����ţ�2-����Һţ�
  --   pati_id              N 1 ��ǰ����id
  --   gvcard_type_id       N 1 �����ID
  --   gvcard_no            C 1 ҽ�ƿ���
  --����: Json_Out,��ʽ����
  --  output
  --    code               C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist            N 1 �Ƿ���ڣ�1-����;0-������

  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_����id   ����ҽ�ƿ���Ϣ.����id%Type;
  n_�����id ����ҽ�ƿ���Ϣ.�����id%Type;
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_����     Number(2);
  n_ģ��     Zlparameters.ģ��%Type;
  v_Msg      Varchar2(3000);

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  n_�����id := j_Json.Get_Number('gvcard_type_id');
  v_����     := j_Json.Get_String('gvcard_no');
  n_����     := j_Json.Get_Number('occasion');
  If Nvl(n_����, 0) = 0 Then
    Json_Out := Zljsonout('δ���볡�ϣ����飡');
    Return;
  End If;
  If n_���� = 1 Then
    n_ģ�� := 1107;
  Else
    n_ģ�� := 1111;
  End If;
  v_Msg := Zl1_Ex_Refundcard_Check(n_ģ��, n_����id, n_�����id, v_����);

  If Nvl(v_Msg, '-') = '-' Then
    Json_Out := Zljsonout('�ɹ�', 1);
  Else
    Json_Out := Zljsonout(v_Msg);
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Checkreturncard;
/
Create Or Replace Procedure Zl_Patisvr_Chkcardchangevalid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ҽ�ƿ��䶯ǰ�ĺϷ��Լ��
  --��Σ�Json_In:��ʽ
  --  input
  --    oper_state          N  1  ����״̬:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
  --    cardtype_id         N  1  �����ID
  --    cardno              C  1  ���ţ������������������ȵĿ��Ż�����������ԭʼ����
  --    new_cardno          C     �¿���:����ʱ���¿���
  --    pati_id             N  1  ����ID
  --    err_id              N     �쳣ҵ��ID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json   PLJson;
  j_Jsonin PLJson;

  n_����״̬ Number(3);
  n_�����id Number(18);
  v_����     Varchar2(100);
  v_�¿���   Varchar2(100);
  n_����id   Number;
  v_Ӧ����Ϣ Varchar2(32767);
  n_Ӧ����   Number(5);
  n_�쳣id   Number(18);
Begin
  --�������
  j_Jsonin := PLJson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_����״̬ := j_Json.Get_Number('oper_state');
  n_�����id := j_Json.Get_Number('cardtype_id');
  v_����     := j_Json.Get_String('cardno');
  v_�¿���   := j_Json.Get_String('new_cardno');
  n_����id   := j_Json.Get_Number('pati_id');
  n_�쳣id   := Nvl(j_Json.Get_Number('err_id'), 0);

  Zl_ҽ�ƿ��䶯_Insert_Check(n_����״̬, n_�����id, v_����, v_�¿���, n_����id, 0, n_Ӧ����, v_Ӧ����Ϣ,n_�쳣id);

  Json_Out := zlJsonOut(v_Ӧ����Ϣ, n_Ӧ����);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Patisvr_Chkcardchangevalid;
/

Create Or Replace Procedure Zl_Patisvr_Confirmcardchange
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:ҽ�ƿ��䶯ȷ��
  --��Σ�Json_In:��ʽ
  --  input
  --    change_id           N  1  �䶯id
  --    pati_id             N  1  ����ID
  --    cardtype_id         N  1 �����ID
  --    card_no             C  1 ҽ�ƿ���
  --    card_notes          C  1 �䶯ԭ��
  --    card_pwd            C  1 ����
  --    card_use_endtime    C  1  ��ֹʹ��ʱ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  n_�䶯id Number(18);
  n_����id Number(18);

  n_�����id ����ҽ�ƿ��䶯.�����id%Type;
  v_����     ����ҽ�ƿ��䶯.����%Type;
  v_�䶯ԭ�� ����ҽ�ƿ��䶯.�䶯ԭ��%Type;
  v_����     ����ҽ�ƿ��䶯.ԭ����%Type;
  d_��ֹʱ�� ����ҽ�ƿ��䶯.��ֹʹ��ʱ��%Type;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_�䶯id := j_Json.Get_Number('change_id');
  n_����id := j_Json.Get_Number('pati_id');

  n_�����id := j_Json.Get_Number('cardtype_id');
  v_����     := j_Json.Get_String('card_no');
  v_�䶯ԭ�� := j_Json.Get_String('card_notes');
  v_����     := j_Json.Get_String('card_pwd');
  d_��ֹʱ�� := To_Date(j_Json.Get_String('card_use_endtime'), 'yyyy-mm-dd hh24:mi:ss');
  If Nvl(n_�����id, 0) = 0 Then
    n_�����id := Null;
  End If;

  Zl_����ҽ�ƿ��䶯_Confirm(n_�䶯id, n_����id, n_�����id, v_����, v_�䶯ԭ��, v_����, d_��ֹʱ��);

  Json_Out := Zljsonout('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Confirmcardchange;
/
Create Or Replace Procedure Zl_Patisvr_Delcardchangeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:ɾ��ҽ�ƿ��䶯��Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --  change_id             N 1 �䶯id
  --  cardtype_id           C 1 �����id
  --  cardno                C 1 ����
  --  pati_id               N 1 ����ID

  --����: Json_Out,��ʽ����
  --  output
  --    code                N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_�䶯id   Number(18);
  n_�����id Number(18);
  v_����     Varchar2(100);
  n_����id   Number(18);

Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_�䶯id   := j_Json.Get_Number('change_id');
  n_�����id := j_Json.Get_Number('cardtype_id');
  v_����     := j_Json.Get_String('cardno');
  n_����id   := j_Json.Get_Number('pati_id');

  Zl_ҽ�ƿ��䶯��¼_Delete(n_�䶯id, n_�����id, v_����, n_����id);

  Json_Out := Zljsonout('�ɹ�', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Delcardchangeinfo;
/
Create Or Replace Procedure Zl_Patisvr_Deletepatiinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:ɾ��ָ������
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N  1  ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  n_����id ������Ϣ.����id%Type;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');

  Zl_������Ϣ_Delete_s(n_����id);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Deletepatiinfo;
/
Create Or Replace Procedure Zl_Patisvr_Deletepatiphoto
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:ɾ��ָ��������Ƭ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N  1  ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  n_����id ������Ƭ.����id%Type;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  Zl_������Ƭ_Delete(n_����id);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Deletepatiphoto;

/
Create Or Replace Procedure Zl_Patisvr_Getblacklistbycons
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  --------------------------------------------------------------------------------------------------
  --����:ʵ����֤ǰ�ļ��  
  --��� JSOM��ʽ
  --input
  --  pati_id        N 1 ����id
  --  operat_type    C 1 ��Ϊ���
  --���� JSON��ʽ
  --output
  --  code            N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  black_list[]
  --     pati_id         N 1 ����id
  --     sign            C 1 ������Ϣ
  --------------------------------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_����id   Number;
  v_ִ����� Varchar2(200);
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  v_ִ����� := j_Json.Get_String('operat_type');

  For c_���� In (Select ����id, ������Ϣ
               From ���˲�����¼
               Where ��Ϊ��� = v_ִ����� And ((����id = n_����id And Nvl(n_����id, 0) <> 0) Or Nvl(n_����id, 0) = 0)) Loop
  
    v_Jtmp := v_Jtmp || ',{"pati_id":' || c_����.����id;
    v_Jtmp := v_Jtmp || ',"sign":"' || Zljsonstr(c_����.������Ϣ) || '"';
    v_Jtmp := v_Jtmp || '}';
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  
  End Loop;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","black_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","black_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getblacklistbycons;
/
Create Or Replace Procedure Zl_Patisvr_Getblackregnos
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���ID,��ȡ����������ĹҺŵ���
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id  N  1  ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    last_time         C  1  ���벻����¼�����һ��ʱ�䣺yyyy-mm-dd hh24:mi:ss
  --    regnos            C  1  ����������ĹҺŵ���,����ö��ŷ���

  ---------------------------------------------------------------------------
  j_Json         Pljson;
  j_Jsonin       Pljson;
  n_����id       ���˲�����¼.����id%Type;
  d_��������     Date;
  v_���ԤԼʱ�� Varchar2(30);
  v_Nos          Varchar2(32680);

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    d_�������� := Trunc(Sysdate) - 1 / 24 / 60 / 60; -- ȱʡ����ͷһ�������
    For c_ԤԼ In (Select Distinct a.������Ϣ
                 From ���˲�����¼ A
                 Where ��Ϊ��� = 'ԤԼ�Һ�' And a.����ʱ�� >= Trunc(d_��������) And a.����ʱ�� <= d_��������) Loop
    
      v_Nos := Nvl(v_Nos, '') || ',' || c_ԤԼ.������Ϣ;
    End Loop;
  Else
    --��Ҫ�������ʷ����,���ܻ�����
    Select To_Char(Nvl(Max(����ʱ��), To_Date('2000-01-01', 'yyyy-mm-dd')), 'yyyy-mm-dd hh24:mi:ss')
    Into v_���ԤԼʱ��
    From ���˲�����¼ A
    Where ����id = n_����id And (��Ϊ��� = 'ԤԼ����' Or (����ԭ�� = 'ԤԼʧԼ��������,�Զ����������' And ��Ϊ��� = '����'));
  
    For c_ԤԼ In (Select Distinct ������Ϣ From ���˲�����¼ Where ����id = n_����id And ��Ϊ��� = 'ԤԼ�Һ�') Loop
      v_Nos := Nvl(v_Nos, '') || ',' || c_ԤԼ.������Ϣ;
    End Loop;
  
  End If;
  If v_Nos Is Not Null Then
    v_Nos := Substr(v_Nos, 2);
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","last_time":"' || v_���ԤԼʱ�� || '","regnos":"' || Zljsonstr(v_Nos) ||
              '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getblackregnos;
/
Create Or Replace Procedure Zl_Patisvr_Getcardlastchange
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡָ����ҽ�ƿ������һ�α䶯��Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      cardtype_id       N 1 �����ID
  --      card_no             C 1 ����
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    change_type       N   1   ���һ�εı䶯����
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;

  n_����id   ����ҽ�ƿ���Ϣ.����id%Type;
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_�����id ����ҽ�ƿ���Ϣ.�����id%Type;
  n_�䶯���� ����ҽ�ƿ��䶯.�䶯���%Type;

Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  v_����     := j_Json.Get_String('card_no');
  n_�����id := j_Json.Get_Number('cardtype_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡����Ϣ�����飡');
    Return;
  End If;

  If Nvl(n_�����id, 0) = 0 Or Nvl(v_����, '-') = '-' Then
    Json_Out := Zljsonout('δ���뿨���򿨺ţ����飡');
    Return;
  End If;

  Select Max(�䶯���)
  Into n_�䶯����
  From (With ҽ�ƿ��䶯 As (Select ����id, ID, �䶯���, �䶯ʱ��
                       From ����ҽ�ƿ��䶯 Bd
                       Where Bd.���� = v_���� And �����id = n_�����id And ����id = n_����id)
         Select a.�䶯���
         From ҽ�ƿ��䶯 A, (Select Max(�䶯ʱ��) As �䶯ʱ�� From ҽ�ƿ��䶯 C) B
         Where a.�䶯ʱ�� = b.�䶯ʱ��) A;


  Json_Out := '{"output":{"code":1,"message":"�ɹ�","change_type":' || Nvl(n_�䶯����, 0) || '}}';

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Getcardlastchange;

/
Create Or Replace Procedure Zl_Patisvr_Getcardtypes
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡҽ�ƿ��������
  --��Σ�Json_In:��ʽ
  --    input
  --      cardtype_id          N 0 �����id:NULL��ʾ���������ID����
  --      query_type           N 1 ��ѯ����:0-������Ϣ;1-������Ϣ(����:id,���룬����,���ų���,ǰ׺�ı�,�Ƿ�����,���㷽ʽ,�Ƿ�ȫ��,�Ƿ�����)
  --      cert_cardtype        N 0 ֻ��ȡ֤����Ϊ������ҽ�ƿ����:1-ֻ��ȡ�Ƿ�֤��=1��ҽ�ƿ�;0-ȫ����ȡ
  --      dffective_cardtype   N 0 ֻ��ȡ��Ч�Ŀ����:1-ֻ��ȡ��Ч�Ŀ����;0-ȫ����ȡ
  --      cardtype_name        C 0 �����ƣ����뿨���ƻ��ض���Ŀ���ƽ��й��ˣ�Ĭ�ϴ���
  --����: Json_Out,��ʽ����
  --  output
  --    code                   N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    type_list[]            C   1   ֧�ֵĿ�����б�
  --        cardtype_id        N   1   ID
  --        cardtype_code      C   1   ����
  --        cardtype_name      C   1   ����
  --        cardtype_stname    C   1   ����
  --        prefix_text        C   1   ǰ׺�ı�
  --        cardno_len         N   1   ���ų���
  --        default            N   1   ȱʡ��־
  --        fixed              N   1   �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
  --        strict             N   1   �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
  --        self_make          N   1   �Ƿ�����:1-�ǵ�;0-����
  --        exist_account      N   1   �Ƿ�����ʻ�:1-�����ʻ�;0-�������˻�
  --        allow_return_cash  N   1   �Ƿ�����:1-����;0-������
  --        must_all_return    N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
  --        component          C   1   ����
  --        memo               C   1   ��ע
  --        spec_item          C   1   �ض���Ŀ
  --        blnc_mode          C   1   ���㷽ʽ
  --        blnc_nature        N   1   ��������
  --        cardno_pwdtxt      C   1   ��������:���Ŵӵڼ�λ���ڼ�λ��ʾ����,��ʽΪ:S-N:S��ʾ�ӵڼ�λ��ʼ,���ڼ�λ����.����:3-10,��ʾ��3λ��10λ������*��ʾ:12********3323��Ҫ����Ӧ��ͬ����ҽ�ƿ�
  --        allow_repeat_use   N   1   �Ƿ��ظ�ʹ��:1-����;0-������
  --        enabled            N   1   �Ƿ�����:1-������;0-δ����
  --        pwd_len            N   1   ���볤��
  --        pwd_len_limit      N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��



  --        pwd_rule           N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
  --        allow_vaguefind    N   1   �Ƿ�ģ������:1-֧��ģ������;0-��֧��
  --        pwd_require        N   1   ������������:0-������;1-������,����;2-�������ֹ;ȱʡΪ������
  --        default_pwd        N   1   �Ƿ�ȱʡ����:1-�����֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
  --        allow_makecard     N   1   �Ƿ��ƿ�:1-��;0-��
  --        allow_sendcard     N   1   �Ƿ񷢿�:1-��;0-��
  --        allow_writcard     N   1   �Ƿ�д��:1-��;0-��
  --        insurance_type     N   1   ����
  --        insurance_name     C   1   ��������
  --        sendcard_nature    N   1   ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0
  --        allow_transfer     N   1   �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
  --        readcard_nature    C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
  --        keyboard_mode      N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
  --        advsend_buildqrcode N   1   �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
  --        holding_pay         N   1   �Ƿ�ֿ�����:1-��;0-��
  --        cert_cardtype       N   1   �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
  --        verfycard           N   1   �Ƿ��˿��鿨
  --        sendcard_sign       N   1   ��������:0��NULL-����ʱ�����ű���ﵽ���ų���;1-����ʱ��������С�ڵ��ڿ��ų���,����ʱ��С�ڿ��ų���ʱ������ʾ����Ա;2-����ʱ��������С�ڵ��ڿ��ų���,С��ʱ����ʾ����Ա��
  --        enterkey_enabled    N   1   �豸�Ƿ����ûس�:ҽ�ƿ���Ӧ��ˢ���豸�Ƿ������˻س�����������˻س����򿨺ų���Ĭ������һλ�����λس�



  --        def_return_cash     N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
  --        balalone            N   1   �Ƿ��������:1-��������;0-�Ƕ�������
  --        discern_rule        N   1   ����ʶ�����:1-ȫ��ת��Ϊ��д;0-�����ִ�Сд
  --        def_valid_time      C   1   ȱʡ��Чʱ��:NULLʱ����ʾ������;�ǿ�ʱ����ʽΪ:ʱ��+��λ(�죬��),���磺3��,3��
  --        scanpay             N   1   �Ƿ�֧��ɨ�븶:�Ƿ�֧��ɨ�븶,֧��ʱ������á�zlReadQRCode������
  ---------------------------------------------------------------------------

  j_Json       Pljson;
  j_Jsonin     Pljson;
  v_Jvals      Varchar2(32767);
  n_�����id   ҽ�ƿ����.Id%Type;
  n_��ѯ����   Number(2);
  n_�Ƿ�֤��   Number(2);
  n_�Ƿ���Ч�� Number(2);
  v_������     Varchar2(2000);
Begin
  --�������
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_�����id   := j_Json.Get_Number('cardtype_id');
  n_��ѯ����   := j_Json.Get_Number('query_type');
  n_�Ƿ�֤��   := j_Json.Get_String('cert_cardtype');
  n_�Ƿ���Ч�� := j_Json.Get_Number('dffective_cardtype');
  v_������     := j_Json.Get_String('cardtype_name');

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","type_list":[';
  v_Jvals  := Null;

  For c_����� In (Select a.Id, a.����, a.����, a.����, a.ǰ׺�ı�, a.���ų���, a.ȱʡ��־, a.�Ƿ�̶�, a.�Ƿ��ϸ����, a.�Ƿ�����, a.�Ƿ�����ʻ�, a.�Ƿ�����, a.�Ƿ�ȫ��,
                       a.����, a.��ע, a.�ض���Ŀ, a.���㷽ʽ, a.��������, a.�Ƿ��ظ�ʹ��, a.�Ƿ�����, a.���볤��, a.���볤������, a.�������, a.�Ƿ�ģ������, a.������������,
                       a.�Ƿ�ȱʡ����, a.�Ƿ��ƿ�, a.�Ƿ񷢿�, a.�Ƿ�д��, a.����, a.��������, a.�Ƿ�ת�ʼ�����, a.��������, a.���̿��Ʒ�ʽ, a.���͵��ýӿ�, a.�Ƿ�ֿ�����,
                       a.�Ƿ�֤��, a.�Ƿ��˿��鿨, a.��������, a.�豸�Ƿ����ûس�, a.�Ƿ�ȱʡ����, a.�Ƿ��������, a.����ʶ�����, a.ȱʡ��Чʱ��, a.�Ƿ�֧��ɨ�븶,
                       b.���� As ��������, c.���� As ��������
                From ҽ�ƿ���� A, ���㷽ʽ��b, ������� C
                Where a.���㷽ʽ = b.����(+) And Decode(Nvl(n_�����id, 0), 0, 0, a.Id) = Nvl(n_�����id, 0) And
                      Decode(Nvl(n_�Ƿ�֤��, 0), 0, 0, Nvl(a.�Ƿ�֤��, 0)) = Nvl(n_�Ƿ�֤��, 0) And
                      Decode(Nvl(n_�Ƿ���Ч��, 0), 0, 0, Nvl(a.�Ƿ�����, 0)) = Nvl(n_�Ƿ���Ч��, 0) And Nvl(a.����, 0) = c.���(+) And
                      (v_������ Is Null Or a.���� = v_������ Or a.�ض���Ŀ = v_������)) Loop
  
    v_Jvals := v_Jvals || ',{"cardtype_id":' || c_�����.Id;
    v_Jvals := v_Jvals || ',"cardtype_code":"' || c_�����.���� || '"';
    v_Jvals := v_Jvals || ',"cardtype_name":"' || c_�����.���� || '"';
    v_Jvals := v_Jvals || ',"cardtype_stname":"' || c_�����.���� || '"';
    v_Jvals := v_Jvals || ',"prefix_text":"' || c_�����.ǰ׺�ı� || '"';
    v_Jvals := v_Jvals || ',"cardno_len":' || Nvl(c_�����.���ų���, 0);
    v_Jvals := v_Jvals || ',"default":' || Nvl(c_�����.ȱʡ��־, 0);
    v_Jvals := v_Jvals || ',"fixed":' || Nvl(c_�����.�Ƿ�̶�, 0);
    v_Jvals := v_Jvals || ',"strict":' || Nvl(c_�����.�Ƿ��ϸ����, 0);
    v_Jvals := v_Jvals || ',"self_make":' || Nvl(c_�����.�Ƿ�����, 0);
    v_Jvals := v_Jvals || ',"exist_account":' || Nvl(c_�����.�Ƿ�����ʻ�, 0);
    v_Jvals := v_Jvals || ',"allow_return_cash":' || Nvl(c_�����.�Ƿ�����, 0);
    v_Jvals := v_Jvals || ',"must_all_return":' || Nvl(c_�����.�Ƿ�ȫ��, 0);
    v_Jvals := v_Jvals || ',"component":"' || c_�����.���� || '"';
    v_Jvals := v_Jvals || ',"memo":"' || c_�����.��ע || '"';
    v_Jvals := v_Jvals || ',"spec_item":"' || c_�����.�ض���Ŀ || '"';
    v_Jvals := v_Jvals || ',"blnc_mode":"' || c_�����.���㷽ʽ || '"';
    v_Jvals := v_Jvals || ',"blnc_nature":' || Nvl(c_�����.��������, 0);
    v_Jvals := v_Jvals || ',"cardno_pwdtxt":"' || c_�����.�������� || '"';
    v_Jvals := v_Jvals || ',"allow_repeat_use":' || Nvl(c_�����.�Ƿ��ظ�ʹ��, 0);
    v_Jvals := v_Jvals || ',"enabled":' || Nvl(c_�����.�Ƿ�����, 0);
  
    If Nvl(n_��ѯ����, 0) = 0 Then
    
      --��ʾ����
      v_Jvals := v_Jvals || ',"pwd_len":' || Nvl(c_�����.���볤��, 0);
      v_Jvals := v_Jvals || ',"pwd_len_limit":' || Nvl(c_�����.���볤������, 0);
      v_Jvals := v_Jvals || ',"pwd_rule":' || Nvl(c_�����.�������, 0);
      v_Jvals := v_Jvals || ',"allow_vaguefind":' || Nvl(c_�����.�Ƿ�ģ������, 0);
      v_Jvals := v_Jvals || ',"pwd_require":' || Nvl(c_�����.������������, 0);
      v_Jvals := v_Jvals || ',"default_pwd":' || Nvl(c_�����.�Ƿ�ȱʡ����, 0);
      v_Jvals := v_Jvals || ',"allow_makecard":' || Nvl(c_�����.�Ƿ��ƿ�, 0);
      v_Jvals := v_Jvals || ',"allow_sendcard":' || Nvl(c_�����.�Ƿ񷢿�, 0);
      v_Jvals := v_Jvals || ',"allow_writecard":' || Nvl(c_�����.�Ƿ�д��, 0);
      v_Jvals := v_Jvals || ',"insurance_type":' || Nvl(c_�����.����, 0);
      v_Jvals := v_Jvals || ',"insurance_name":"' || c_�����.�������� || '"';
      v_Jvals := v_Jvals || ',"sendcard_nature":' || Nvl(c_�����.��������, 0);
      v_Jvals := v_Jvals || ',"allow_transfer":' || Nvl(c_�����.�Ƿ�ת�ʼ�����, 0);
      v_Jvals := v_Jvals || ',"readcard_nature":"' || Nvl(c_�����.��������, '1000') || '"';
      v_Jvals := v_Jvals || ',"keyboard_mode":' || Nvl(c_�����.���̿��Ʒ�ʽ, 0);
      v_Jvals := v_Jvals || ',"advsend_buildqrcode":' || Nvl(c_�����.���͵��ýӿ�, 0);
      v_Jvals := v_Jvals || ',"holding_pay":' || Nvl(c_�����.�Ƿ�ֿ�����, 0);
      v_Jvals := v_Jvals || ',"cert_cardtype":' || Nvl(c_�����.�Ƿ�֤��, 0);
      v_Jvals := v_Jvals || ',"verfycard":' || Nvl(c_�����.�Ƿ��˿��鿨, 0);
      v_Jvals := v_Jvals || ',"sendcard_sign":' || Nvl(c_�����.��������, 0);
      v_Jvals := v_Jvals || ',"enterkey_enabled":' || Nvl(c_�����.�豸�Ƿ����ûس�, 0);
      v_Jvals := v_Jvals || ',"def_return_cash":' || Nvl(c_�����.�Ƿ�ȱʡ����, 0);
      v_Jvals := v_Jvals || ',"balalone":' || Nvl(c_�����.�Ƿ��������, 0);
      v_Jvals := v_Jvals || ',"discern_rule":' || Nvl(c_�����.����ʶ�����, 0);
      v_Jvals := v_Jvals || ',"def_valid_time":"' || c_�����.ȱʡ��Чʱ�� || '"';
      v_Jvals := v_Jvals || ',"scanpay":' || Nvl(c_�����.�Ƿ�֧��ɨ�븶, 0);
    End If;
  
    v_Jvals := v_Jvals || '}';
    If Length(v_Jvals) > 30000 Then
      Json_Out := Json_Out || Substr(v_Jvals, 2);
      v_Jvals  := Null;
    End If;
  End Loop;
  If v_Jvals Is Not Null Then
    v_Jvals  := v_Jvals || ']}}';
    Json_Out := Json_Out || Substr(v_Jvals, 2);
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getcardtypes;
/
Create Or Replace Procedure Zl_Patisvr_Getcommunityinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡ���˵�������Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --  query_type        N    1 �������� ��1-ͨ������id���������ƻ�ȡ�����ţ�2-��ȡ������ҽ����[���أ�community_code+insurance_num]
  --  pati_id           N    1 ����id
  --  community_id      N    1 ����id

  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ

  --    community_code      C   1 �����š�query_type=2��
  --    insurance_num       C   1 ҽ���š�query_type=2��

  --    community_list[]������Ϣ�б�֧�ֶ����[����]��query_type=1��
  --       community_code    C   1 ������
  -------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
  v_List   Varchar2(32767);

  n_Type   Number(18);
  n_����id Number(18);
  n_����id Number(18);
  v_ҽ���� Varchar2(4000);
  v_������ Varchar2(4000);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_Type   := j_Json.Get_Number('query_type');
  n_����id := j_Json.Get_Number('pati_id');

  If n_Type = 2 Then
    n_����id := j_Json.Get_String('community_id');
    Select Max(a.ҽ����) Into v_ҽ���� From ������Ϣ A Where a.����id = n_����id;
    Select Max(a.������) Into v_������ From ����������Ϣ A Where a.����id = n_����id And a.���� = n_����id;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","community_code":"' || Zljsonstr(v_������) || '","insurance_num":"' ||
                Zljsonstr(v_ҽ����) || '"}}';
    Return;
  End If;

  If n_Type = 1 Then
  
    n_����id := j_Json.Get_String('community_num');
  
    For c_������Ϣ In (Select a.������ From ����������Ϣ A Where a.����id = n_����id And ���� = n_����id) Loop
    
      v_List := v_List || ',{"community_code":"' || Zljsonstr(c_������Ϣ.������) || '"}';
    
    End Loop;
  
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","community_list":[' || Substr(v_List, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getcommunityinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getcustompatiinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:�������֤�źͻ�����Ϣ����ȡ���˵���ϸ��Ϣ(�û��Զ��巵��)
  --��Σ�Json_In:��ʽ
  --    input
  --      occasion          N 1 ����
  --      pati_idcard       C 1 ���֤��
  --      pati_name         C   ����
  --      pati_sex          C 1 �Ա�
  --      query_type        N   ��ѯ���ͣ�0-��ѯ������Ϣ��1-ֻ��ѯ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_ids          N   1   ����ids.��ѯ����Ϊ1ʱ����
  --    pati_count        N   1   ������Ϣ����.��ѯ����Ϊ0ʱ����
  --    pati_list         C       ������Ϣ�б�.��ѯ����Ϊ0ʱ����
  --      pati_id         N 1 ����id
  --      pati_pageid     N 1 ��ҳid��������Ϣ.��ҳID
  --      pati_name       C 1 ����
  --      pati_sex        C 1 �Ա�
  --      pati_age        C 1 ����
  --      pati_birthdate  D 1 ��������
  --      pati_nation     C 1 ����
  --      pati_idcard     C 1 ���֤��
  --      pati_education  C   ѧ��
  --      pati_identity   C   ���
  --      pati_marital_cstatus  C   ����״��
  --      pat_home_addr   C   ��ͥ��ַ
  --      pati_area       C   ����
  --      pati_birthplace C   �����ص�
  --      pati_emp_name   C   ������λ����
  --      outpatient_num  C   �����
  --      inpatient_num   C   סԺ��
  --      insurance_num   C   ҽ����
  --      phone_number    C   ��ϵ�绰(��ϵ�˵绰���ֻ��ţ���ͥ�绰����ȡһ)
  --      pati_bed        C   ��ǰ����
  --      pati_type       C   ��������(��ͨ��ҽ��������)
  --      out_date        D   ��Ժ����
  --      create_time     C   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  j_Jsonput  Pljson;
  j_Json     Pljson;
  n_����     Number(18);
  v_���֤�� ������Ϣ.���֤��%Type;
  v_����     ������Ϣ.����%Type;
  v_�Ա�     ������Ϣ.�Ա�%Type;
  n_��ѯ���� Number(1);
  v_����ids  Varchar2(32767);
  v_List     Varchar2(32767);
  n_Count    Number(10);

Begin
  --�������
  j_Jsonput  := Pljson(Json_In);
  j_Json     := j_Jsonput.Get_Pljson('input');
  n_����     := j_Json.Get_Number('module');
  v_���֤�� := j_Json.Get_String('pati_idcard');
  v_����     := j_Json.Get_String('pati_name');
  v_�Ա�     := j_Json.Get_String('pati_sex');
  n_Count    := 0;
  n_��ѯ���� := j_Json.Get_Number('query_type');
  If v_���֤�� Is Null And v_���� Is Null Then
    Json_Out := Zljsonout('δ�������֤�ź�����������');
    Return;
  End If;
  v_����ids := Zl_Custom_Patiids_Get(n_����, v_���֤��, v_����, v_�Ա�);
  If v_����ids Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_count":0,"pati_list":[]}}';
    Return;
  End If;
  --n_��ѯ���ͣ�0-��ѯ������Ϣ��1-ֻ��ѯ����id
  If Nvl(n_��ѯ����, 0) = 0 Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[';
    For c_������Ϣ In (Select /*+cardinality(B,10)*/
                   Distinct a.����id As ID, a.��ҳid, a.����id, a.����, a.�Ա�, a.����, a.��������, a.����, a.���֤��, a.ѧ��, a.���, a.����״��,
                            a.��ͥ��ַ, a.����, a.�����ص�, a.�����, a.סԺ��, a.ҽ����, Nvl(a.�ֻ���, a.��ͥ�绰) As ��ϵ�绰, a.������λ, a.��ǰ����, a.��������,
                            a.��Ժʱ��, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��
                   From ������Ϣ A, Table(f_Str2list(v_����ids)) B
                   Where a.����id = b.Column_Value
                   Order By ����, �Ա�, ����) Loop
      n_Count := n_Count + 1;
      Zljsonputvalue(v_List, 'pati_id', c_������Ϣ.����id, 1, 1);
      Zljsonputvalue(v_List, 'pati_pageid', c_������Ϣ.��ҳid, 1);
      Zljsonputvalue(v_List, 'pati_name', c_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_sex', c_������Ϣ.�Ա�);
      Zljsonputvalue(v_List, 'pati_age', c_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_birthdate', To_Char(c_������Ϣ.��������, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'pati_nation', c_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_idcard', c_������Ϣ.���֤��);
      Zljsonputvalue(v_List, 'pati_education', c_������Ϣ.ѧ��);
      Zljsonputvalue(v_List, 'pati_identity', c_������Ϣ.���);
      Zljsonputvalue(v_List, 'pati_marital_cstatus', c_������Ϣ.����״��);
      Zljsonputvalue(v_List, 'pat_home_addr', c_������Ϣ.��ͥ��ַ);
      Zljsonputvalue(v_List, 'pati_area', c_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_birthplace', c_������Ϣ.�����ص�);
      Zljsonputvalue(v_List, 'pati_emp_name', c_������Ϣ.������λ);
      Zljsonputvalue(v_List, 'outpatient_num', c_������Ϣ.�����, 0);
      Zljsonputvalue(v_List, 'inpatient_num', c_������Ϣ.סԺ��, 0);
      Zljsonputvalue(v_List, 'insurance_num', c_������Ϣ.ҽ����);
      Zljsonputvalue(v_List, 'phone_number', c_������Ϣ.��ϵ�绰);
      Zljsonputvalue(v_List, 'pati_bed', c_������Ϣ.��ǰ����);
      Zljsonputvalue(v_List, 'pati_type', c_������Ϣ.��������);
      Zljsonputvalue(v_List, 'out_date', To_Char(c_������Ϣ.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'create_time', To_Char(c_������Ϣ.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || '],"pati_count":' || n_Count || '}}';
    End If;
    Return;
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_ids":"' || v_����ids || '"}}';
    Return;
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getcustompatiinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getinputitemlength
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ŀ��ʵ�ʴ�С
  --��Σ�Json_In:��ʽ
  --    input
  --    item_list[]
  --      table_name  C 1 ����
  --      column_name C 1 ����,����ö���
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[] C
  --      table_name  C 1 ����
  --      column_name C 1 �б�
  --      column_size N 1 ����

  ---------------------------------------------------------------------------
  j_Jsonin   Pljson;
  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonlist Pljson_List := Pljson_List();

  v_���� Varchar2(100);
  v_�ֶ� Varchar2(32767);
  v_Jtmp Varchar2(32767);
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('item_list');

  If j_Jsonlist Is Null Then
    Json_Out := Zljsonout('δ������Ҫ��ѯ�ı���Ϣ������!');
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[';
  For I In 1 .. j_Jsonlist.Count Loop
    o_Json := Pljson();
    o_Json := Pljson(j_Jsonlist.Get(I));
    v_���� := o_Json.Get_String('table_name');
    v_�ֶ� := o_Json.Get_String('column_name');
    For c_����Ϣ In (Select Column_Name As ����, Max(Data_Length) As ����
                  From User_Tab_Columns
                  Where Table_Name = v_���� And Instr(',' || v_�ֶ� || ',', ',' || Column_Name || ',') > 0
                  Group By Column_Name) Loop
      v_Jtmp := v_Jtmp || ',{"table_name":"' || Zljsonstr(v_����) || '"';
      v_Jtmp := v_Jtmp || ',"column_name":"' || Zljsonstr(c_����Ϣ.����) || '"';
      v_Jtmp := v_Jtmp || ',"column_size":' || Zljsonstr(c_����Ϣ.����, 1);
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getinputitemlength;
/
Create Or Replace Procedure Zl_Patisvr_Getinsureaccbalance
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ�����ʻ���Ϣ��ҽ���ʻ��� ���ò�������)
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id               N  1  ����ID
  --  insurance_type        N  0   ����
  --����      json
  --output
  --  code                  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message               C  1  Ӧ����Ϣ��
  --  pati_type             C  0  ���ò�������
  --  insure_srpls_chrg     N  0  ҽ���˻����
  ---------------------------------------------------------------------------
  j_Input  Pljson;
  j_Jsonin Pljson;

  n_����id   Number(18);
  n_����     Number(18);
  v_���ò��� Varchar2(200);
  n_�ʻ���� ҽ�����˵���.�ʻ����%Type;

Begin
  j_Jsonin := Pljson(Json_In);
  j_Input  := j_Jsonin.Get_Pljson('input');
  n_����id := j_Input.Get_Number('pati_id');
  n_����   := j_Input.Get_Number('insurance_type');

  Select Decode(Nvl(n_����, 0), 0, Max(����), n_����), Max(Zl_Patiwarnscheme(����id))
  Into n_����, v_���ò���
  From ������Ϣ
  Where ����id = n_����id;

  Select Nvl(Max(e.�ʻ����), 0)
  Into n_�ʻ����
  From ҽ�����˹����� D, ҽ�����˵��� E
  Where d.����id = n_����id And d.���� = Nvl(n_����, 0) And d.���� = e.����(+) And d.ҽ���� = e.ҽ����(+) And d.��־(+) = 1;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_type":"' || Zljsonstr(v_���ò���) || '","insure_srpls_chrg":' ||
              Zljsonstr(n_�ʻ����, 1) || '}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getinsureaccbalance;
/
Create Or Replace Procedure Zl_Patisvr_Getinsureinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���id����ȡ���˵ı�����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id             N   1 ����id
  --    insure_type         N   1 ����
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    insure_type         C   1 ����
  --    insure_name         C   1 ��������
  --    insure_no           C   1 ҽ����
  --    card_no             C   1 ����
  --    pati_create_time    C   1 ���˵ĵǼ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    insure_pwd          C   1 ҽ������
  --    dz_type_id          N   1 ����id
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_����id   ������Ϣ.����id%Type;
  n_����     ������Ϣ.����%Type;
  v_�������� �������.����%Type;
  v_ҽ����   ҽ�����˵���.ҽ����%Type;
  v_����     ҽ�����˵���.����%Type;
  d_�Ǽ�ʱ�� ҽ�����˵���.����ʱ��%Type;
  v_����     ҽ�����˵���.����%Type;
  n_����id   ҽ�����˵���.����id%Type;
  n_Type     Number(1);
  v_Jtmp     Varchar2(32767);
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  n_����   := j_Json.Get_Number('insure_type');
  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('ʧ�ܣ�δ���벡��id��');
    Return;
  End If;
  If Nvl(n_����, 0) = 0 Then
    Select Max(����) Into n_���� From ������Ϣ Where ����id = n_����id;
  Else
    n_Type := 1;
    Select Max(b.����), Max(a.ҽ����), Max(a.����ʱ��), Max(a.����), Max(a.����), Max(a.����id)
    Into v_��������, v_ҽ����, d_�Ǽ�ʱ��, v_����, v_����, n_����id
    From �����ʻ� A, ������� B
    Where a.���� = b.��� And a.����id = n_����id And a.���� = n_����;
  End If;

  If Nvl(n_Type, 0) <> 0 Then
  
    v_Jtmp := v_Jtmp || ',"insure_type":' || Nvl(n_���� || '', 'null');
    v_Jtmp := v_Jtmp || ',"insure_name":"' || Zljsonstr(v_��������) || '"';
    v_Jtmp := v_Jtmp || ',"insure_no":"' || Zljsonstr(v_ҽ����) || '"';
    v_Jtmp := v_Jtmp || ',"card_no":"' || Zljsonstr(v_����) || '"';
    v_Jtmp := v_Jtmp || ',"pati_create_time":"' || To_Char(d_�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
    v_Jtmp := v_Jtmp || ',"insure_pwd":"' || Zljsonstr(v_����) || '"';
    v_Jtmp := v_Jtmp || ',"dz_type_id":' || Nvl(n_����id || '', 'null');
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"' || v_Jtmp || '}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","insure_type":' || Nvl(n_���� || '', 'null') || '}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output": {"code": 0,"message": "' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getinsureinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getlastblackinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------
  --����:��ȡ���һ�εĲ�����¼����Ϣ
  --��� JSOM��ʽ
  --input
  --  pati_id        N 1 ����id
  --���� JSON��ʽ
  --output
  --  code            N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  pati_id         N 1 ����id
  --  last_time       C 1 ���ԤԼʱ��
  --------------------------------------------------------------------------------------------------
  j_Json         Pljson;
  j_Jsonin       Pljson;
  n_����id       Number;
  v_���ԤԼʱ�� Varchar2(30);
Begin
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  Begin
    Select Max(����ʱ��) As ���ԤԼʱ��
    Into v_���ԤԼʱ��
    From ���˲�����¼
    Where ����id = n_����id And (��Ϊ��� = 'ԤԼ����' Or (����ԭ�� = 'ԤԼʧԼ��������,�Զ����������' And ��Ϊ��� = '����'));
  Exception
    When Others Then
      Null;
  End;
  Json_Out := '{"output":{"pati_id":' || n_����id || ',"last_time":"' || To_Char(v_���ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss') ||
              '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getlastblackinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getnextid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡָ��������Ӧ������(���淶������������Ϊ��������_id��)����һ��ֵ
  --��Σ�Json_In:��ʽ
  --input
  --  table_name    C  1 ����
  --  col_name      C  1 �ֶ���  �������Ʋ�һ����ID�������¼ID
  -- ����:
  --  output
  --  next_id      N   1  ����
  -------------------------------------------

  v_Table     Varchar2(500);
  v_Col       Varchar2(500);
  n_Nextid    Number;
  j_Json      Pljson;
  j_Jsoninput Pljson;

Begin
  --�������
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  v_Table     := j_Json.Get_String('table_name');
  v_Col       := Nvl(j_Json.Get_String('col_name'), 'ID');
  Execute Immediate 'select ' || v_Table || '_' || v_Col || '.nextval from dual'
    Into n_Nextid;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","next_id":' || n_Nextid || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Getnextid;
/
Create Or Replace Procedure Zl_Patisvr_Getnextno
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ܣ������ض���������µĺ���
  --��Σ�Json_In:��ʽ
  --  input
  --    item_num            N   1   ��Ŀ���
  --    dept_id             N   0   ����ID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    next_no             C   1   ��һ������
  ---------------------------------------------------------------------------
  j_Json      Pljson;
  j_Jsoninput Pljson;
  v_No        Varchar2(64);
  n_���      Number(10);
  n_����id    Number(18);
Begin
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  n_���      := j_Json.Get_Number('item_num');
  n_����id    := j_Json.Get_Number('dept_id');

  Select Zl_Pati_Nextno(n_���, n_����id) Into v_No From Dual;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","next_no":"' || v_No || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Getnextno;

/
Create Or Replace Procedure Zl_Patisvr_Getpatallergicdrugs
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ϣ�Ĺ���ҩ����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id             N   1 ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    drug_list[]         C       ����ҩ���б�
  --      medicinal_id      N   1   ����ҩƷID
  --      medicinal_name    C   1   ����ҩ������
  --      allergy_info      C   1   ��ÿҩ�ﷴӦ
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  n_����id ���˹���ҩ��.����id%Type;
  v_List   Varchar2(32767);
  v_Jtmp   Varchar2(32767);
  c_Jtmp   Clob;
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id������!');
    Return;
  End If;

  For r_������¼ In (Select Distinct ����ҩ��id, ����ҩ��, ������Ӧ From ���˹���ҩ�� Where ����id = n_����id) Loop
  
    v_Jtmp := v_Jtmp || ',{"medicinal_id":' || Nvl(r_������¼.����ҩ��id || '', 'null');
    v_Jtmp := v_Jtmp || ',"medicinal_name":"' || Zljsonstr(r_������¼.����ҩ��) || '"';
    v_Jtmp := v_Jtmp || ',"allergy_info":"' || Zljsonstr(r_������¼.������Ӧ) || '"';
    v_Jtmp := v_Jtmp || '}';
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","drug_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","drug_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatallergicdrugs;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiaddrssinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID,��ȡ���˵ĵ�ַ��Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id              N   1  ����ID
  --    pati_pageid          N   0  ��ҳid
  --    addr_type            N   0  ��ַ���:1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ;Ϊ0ʱ��ʾ��ѯ�������͵ĵ�ַ��Ϣ
  --    addr_types           C   0  ��ַ���s:1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ;������,��","�ָ�.�����˸ýڵ�ʱ��addr_type��Ч
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    addr_list[]          C       ��ַ�б���Ϣ
  --      pat_addr_type      C   1   ��ַ���
  --      pat_addr_state     C   1   ��ַ_ʡ
  --      pat_addr_city      C   1   ��ַ_��
  --      pat_addr_county    C   1   ��ַ_��
  --      pat_addr_township  C   1   ��ַ_��
  --      pat_addr_other     C   1   ��ַ_����
  --      pat_region_code    C   1   ��������
  ---------------------------------------------------------------------------
  n_����id    ������Ϣ.����id%Type;
  n_��ҳid    ������Ϣ.��ҳid%Type;
  n_��ַ���  ���˵�ַ��Ϣ.��ַ���%Type;
  v_��ַ���s Varchar2(3000);
  v_List      Varchar2(32767);
  j_Json      Pljson;
  j_Jsonin    Pljson;
Begin

  --�������
  j_Jsonin    := Pljson(Json_In);
  j_Json      := j_Jsonin.Get_Pljson('input');
  n_����id    := j_Json.Get_Number('pati_id');
  n_��ҳid    := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  n_��ַ���  := Nvl(j_Json.Get_Number('addr_type'), 0);
  v_��ַ���s := Nvl(j_Json.Get_String('addr_type'), '-');
  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��ID');
    Return;
  End If;
  If v_��ַ���s <> '-' Then
    For r_��ַ In (Select ��ַ���, ʡ, ��, ��, ����, ����, ��������
                 From ���˵�ַ��Ϣ
                 Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(n_��ҳid, 0) And
                       Instr(',' || v_��ַ���s || ',', ',' || ��ַ��� || ',') > 0) Loop
      --      pat_addr_type      C   1 ��ַ���
      --      pat_addr_state     C   1 ��ַ_ʡ
      --      pat_addr_city      C   1 ��ַ_��
      --      pat_addr_county    C   1 ��ַ_��
      --      pat_addr_township  C   1 ��ַ_��
      --      pat_addr_other     C   1 ��ַ_����
      --      pat_region_code    C   1 ��������
      Zljsonputvalue(v_List, 'pat_addr_type', r_��ַ.��ַ���, 0, 1);
      Zljsonputvalue(v_List, 'pat_addr_state', r_��ַ.ʡ);
      Zljsonputvalue(v_List, 'pat_addr_city', r_��ַ.��);
      Zljsonputvalue(v_List, 'pat_addr_county', r_��ַ.��);
      Zljsonputvalue(v_List, 'pat_addr_township', r_��ַ.����);
      Zljsonputvalue(v_List, 'pat_addr_other', r_��ַ.����);
      Zljsonputvalue(v_List, 'pat_region_code', r_��ַ.��������, 0, 2);
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","addr_list":[' || v_List || ']}}';
    Return;
  End If;
  For r_��ַ In (Select ��ַ���, ʡ, ��, ��, ����, ����, ��������
               From ���˵�ַ��Ϣ
               Where ����id = n_����id And Nvl(��ҳid, 0) = Nvl(n_��ҳid, 0) And
                     ((��ַ��� = n_��ַ��� And Nvl(n_��ַ���, 0) <> 0) Or Nvl(n_��ַ���, 0) = 0)) Loop
    --      pat_addr_type      C   1 ��ַ���
    --      pat_addr_state     C   1 ��ַ_ʡ
    --      pat_addr_city      C   1 ��ַ_��
    --      pat_addr_county    C   1 ��ַ_��
    --      pat_addr_township  C   1 ��ַ_��
    --      pat_addr_other     C   1 ��ַ_����
    --      pat_region_code    C   1 ��������
    Zljsonputvalue(v_List, 'pat_addr_type', r_��ַ.��ַ���, 0, 1);
    Zljsonputvalue(v_List, 'pat_addr_state', r_��ַ.ʡ);
    Zljsonputvalue(v_List, 'pat_addr_city', r_��ַ.��);
    Zljsonputvalue(v_List, 'pat_addr_county', r_��ַ.��);
    Zljsonputvalue(v_List, 'pat_addr_township', r_��ַ.����);
    Zljsonputvalue(v_List, 'pat_addr_other', r_��ַ.����);
    Zljsonputvalue(v_List, 'pat_region_code', r_��ַ.��������, 0, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","addr_list":[' || v_List || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiaddrssinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiblackinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���˺�������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id
  --      occasion          C 1 Ӧ�ó���:ԤԼ���Һţ����ʣ���Ժ����Ժ
  --      appt_mode_name    C 0 ԤԼ��ʽ

  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    tip_mode          N   1   ���Ʒ�ʽ��1-��ֹ;2-��ʾ(��ѯ��)
  --    tip_message       C   1   ��ʾ����Ϣ
  ---------------------------------------------------------------------------

  j_Json        Pljson;
  j_Jsonin      Pljson;
  n_����id      ���˲�����¼.����id%Type;
  v_Ӧ�ó���    ������Ϊ����.Ӧ�ó���%Type;
  v_ԤԼ��ʽ    ������Ϊ����.ԤԼ��ʽ%Type;
  n_���Ʒ�ʽ    Number(1);
  v_Message     Varchar2(30000);
  v_Black_Infor Varchar2(32767);
  v_Tmp         Varchar2(30000);

Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  v_Ӧ�ó��� := j_Json.Get_String('occasion');
  v_ԤԼ��ʽ := j_Json.Get_String('appt_mode_name');

  v_Black_Infor := Zl_Fun_Getblacklistinfor(n_����id, v_Ӧ�ó���, v_ԤԼ��ʽ);

  If Nvl(v_Black_Infor, '-') <> '-' Then
    v_Tmp      := Substr(v_Black_Infor, 1, 1);
    n_���Ʒ�ʽ := To_Number(v_Tmp);
    v_Message  := Substr(v_Black_Infor, 3);
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","tip_mode":' || Nvl(n_���Ʒ�ʽ, 0) || ',"tip_message":"' ||
              Zljsonstr(v_Message) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiblackinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpaticardinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_ids            C  1  ����ids,�����","�ָ�
  --      cardtype_ids        C  1  �����IDs,����ö��ŷ���
  --      card_no             C  0  ����
  --      card_name           C  0  ҽ�ƿ�����
  --      cert_cardtype       N  1  ֻ��ȡ֤����Ϊ������ҽ�ƿ����:1-ֻ��ȡ�Ƿ�֤��=1��ҽ�ƿ�;0-ȫ����ȡ
  --      query_type          N  1  ��ѯ��������:0-ֻ��ȡ����ID,1-ֻ��ȡ�����ID;2-�������˻�����Ϣ;3-����
  --      dffective_cardtype  N  0  ֻ��ȡ��Ч�Ŀ����:1-ֻ��ȡ��Ч�Ŀ����;0-ȫ����ȡ

  -- ���Σ�json
  --  output
  --  code                    N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                 C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  card_list[]             C     ���˿���Ϣ�б�
  --    pati_id               N  1  ����id
  --    pati_name             C  1  ����
  --    pati_sex              C  1  �Ա�
  --    pati_age              C  1  ����
  --    pati_birthdate        C  1  �������ڣ�yyyy-mm-dd hh24:mi:ss
  --    outpatient_num        C  1  �����
  --    pati_idcard           C  1  ���֤��
  --    cardtype_id           N  1  �����ID
  --    card_no               C  1  ����
  --    card_qrcode           C  1  ��ά��
  --    card_passwod          C  1  ����
  --    cardtype_name         C  1  ���������
  --    cardtype_cardlen      N  1  ���ų���
  --    card_statu            N  1  ״̬:0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
  --    loscard_creator       C  1  ��ʧ��
  --    loscard_time          C  1  ��ʧʱ��:yyyy-mm-dd hh24:mi:ss
  --    loscard_mode          C  1  ��ʧ��ʽ
  --    loscard_days          N  1  ��ʧ����
  --    sendcard_oper         C  1  ������
  --    end_time              C  1  ��ֹʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Json    Pljson;
  j_Jsonin  Pljson;

  j_Json_Tmp   Pljson;
  v_����ids    Varchar2(3000);
  v_List       Varchar2(32767);
  v_Tmp        Varchar2(32767);
  v_�����ids  Varchar2(32680);
  v_����       Varchar2(1000);
  v_������     Varchar2(3000);
  n_��ѯ����   Number(2);
  n_�Ƿ�֤��   Number(2);
  n_�Ƿ���Ч�� Number(2);
  n_�����id   ����ҽ�ƿ���Ϣ.�����id%Type;

  Cursor c_���˻�����Ϣ Is
    Select a.����id, a.�����id, a.���������, a.����, a.���ų���, a.����, a.״̬, To_Char(a.��ʧʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʧʱ��, a.��ʧ��ʽ,
           a.��ʧ��, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.������,
           To_Char(a.��ֹʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ֹʹ��ʱ��, a.��ά��, b.����, b.�Ա�, b.����,
           To_Char(b.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, b.���֤��, b.�����, a.��Ч����
    From (Select a.����id, a.�����id, a.����, a.����, a.״̬, a.��ʧʱ��, a.��ʧ��ʽ, a.��ʧ��, a.��������, a.������, a.��ֹʹ��ʱ��, a.��ά��, q.���� As ���������,
                  q.���ų���, m.��Ч����
           From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
           Where a.��ʧ��ʽ = m.����(+) And a.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And a.����id = 0) A, ������Ϣ B
    Where a.����id = b.����id And Rownum < 1;
  r_���� c_���˻�����Ϣ%RowType;

  Type Ty_������Ϣ Is Ref Cursor;
  c_������Ϣ Ty_������Ϣ; --��̬�α����

Begin
  --�������
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  v_����ids    := j_Json.Get_String('pati_ids');
  v_�����ids  := j_Json.Get_String('cardtype_ids');
  n_��ѯ����   := j_Json.Get_Number('query_type');
  n_�Ƿ�֤��   := j_Json.Get_Number('cert_cardtype');
  v_����       := j_Json.Get_String('card_no');
  n_�Ƿ���Ч�� := j_Json.Get_Number(' dffective_cardtype');
  v_������     := j_Json.Get_String('card_name');

  If Nvl(v_����ids, '-') = '-' And v_���� Is Null Then
    v_Err_Msg := 'δ������Ч�Ĳ�ѯ���������ܻ�ȡ���������еĿ����!';
    Json_Out  := Zljsonout(v_Err_Msg);
    Return;
  End If;

  If Not v_������ Is Null Then
    Select Nvl(Max(ID), 0) Into n_�����id From ҽ�ƿ���� Where ���� = v_������;
  End If;

  If Nvl(v_����ids, '-') <> '-' Then
    Open c_������Ϣ For
      Select a.����id, a.�����id, a.���������, a.����, a.���ų���, a.����, a.״̬, To_Char(a.��ʧʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʧʱ��,
             a.��ʧ��ʽ, a.��ʧ��, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.������,
             To_Char(a.��ֹʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ֹʹ��ʱ��, a.��ά��, b.����, b.�Ա�, b.����,
             To_Char(b.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, b.���֤��, b.�����, a.��Ч����
      From (Select a.����id, a.�����id, a.����, a.����, a.״̬, a.��ʧʱ��, a.��ʧ��ʽ, a.��ʧ��, a.��������, a.������, a.��ֹʹ��ʱ��, a.��ά��,
                    q.���� As ���������, q.���ų���, m.��Ч����,
                    Case
                       When Nvl(a.״̬, 0) = 1 And Nvl(n_�Ƿ���Ч��, 0) = 1 And
                            (Nvl(m.��Ч����, 0) = 0 Or Nvl(a.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.��Ч����, 0) > Sysdate) Then
                        1
                       Else
                        Nvl(a.״̬, 0)
                     End As ״̬1
             From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
             Where a.��ʧ��ʽ = m.����(+) And a.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And
                   a.����id In (Select /*+cardinality(B,10) */
                               Column_Value As ����id
                              From Table(f_Str2list(v_����ids)) B) And
                   (v_�����ids Is Null Or Instr(',' || v_�����ids || ',', ',' || a.�����id || ',') > 0) And
                   Decode(Nvl(n_�Ƿ�֤��, 0), 0, 0, Nvl(q.�Ƿ�֤��, 0)) = Nvl(n_�Ƿ�֤��, 0) And
                   (v_���� Is Null Or a.���� = Nvl(v_����, '-'))) A, ������Ϣ B
      Where a.����id = b.����id And a.״̬1 = 0;
  Elsif v_���� Is Not Null Then
  
    Open c_������Ϣ For
      Select a.����id, a.�����id, a.���������, a.����, a.���ų���, a.����, a.״̬, To_Char(a.��ʧʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʧʱ��,
             a.��ʧ��ʽ, a.��ʧ��, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.������,
             To_Char(a.��ֹʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ֹʹ��ʱ��, a.��ά��, b.����, b.�Ա�, b.����,
             To_Char(b.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, b.���֤��, b.�����, a.��Ч����
      From (Select a.����id, a.�����id, a.����, a.����, a.״̬, a.��ʧʱ��, a.��ʧ��ʽ, a.��ʧ��, a.��������, a.������, a.��ֹʹ��ʱ��, a.��ά��,
                    q.���� As ���������, q.���ų���, m.��Ч����,
                    Case
                       When Nvl(a.״̬, 0) = 1 And Nvl(n_�Ƿ���Ч��, 0) = 1 And
                            (Nvl(m.��Ч����, 0) = 0 Or Nvl(a.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.��Ч����, 0) > Sysdate) Then
                        1
                       Else
                        Nvl(a.״̬, 0)
                     End As ״̬1
             From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
             Where a.��ʧ��ʽ = m.����(+) And a.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And a.���� = v_���� And
                   (v_�����ids Is Null Or Instr(',' || v_�����ids || ',', ',' || a.�����id || ',') > 0) And
                   (a.�����id = n_�����id Or n_�����id = 0) And Decode(Nvl(n_�Ƿ�֤��, 0), 0, 0, Nvl(q.�Ƿ�֤��, 0)) = Nvl(n_�Ƿ�֤��, 0)) A,
           ������Ϣ B
      Where a.����id = b.����id And a.״̬1 = 0;
  
  Else
    v_Err_Msg := 'δ������Ч�Ĳ�ѯ���������ܻ�ȡ������Ϣ!';
    Json_Out  := Zljsonout(v_Err_Msg);
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","card_list":[';

  Loop
    Fetch c_������Ϣ
      Into r_����;
    Exit When c_������Ϣ%NotFound;
  
    j_Json_Tmp := Pljson();
    --1.ȡ������Ϣ
    --0-ֻ��ȡ����ID,1-ֻ��ȡ�����ID;2-�������˻�����Ϣ;3-����
    If Nvl(n_��ѯ����, 0) <> 1 Then
      v_Tmp := v_Tmp || ',{"pati_id":' || Nvl(r_����.����id, 0);
      If Nvl(n_��ѯ����, 0) = 0 Then
        v_Tmp := v_Tmp || '}';
      End If;
      If Nvl(n_��ѯ����, 0) <> 0 Then
        v_Tmp := v_Tmp || ',"pati_name":"' || Nvl(r_����.����, '') || '"';
        v_Tmp := v_Tmp || ',"pati_sex":"' || Nvl(r_����.�Ա�, '') || '"';
        v_Tmp := v_Tmp || ',"pati_age":"' || Nvl(r_����.����, '') || '"';
        v_Tmp := v_Tmp || ',"pati_birthdate":"' || Nvl(r_����.��������, '') || '"';
        v_Tmp := v_Tmp || ',"outpatient_num":"' || Zljsonstr(r_����.�����) || '"';
        v_Tmp := v_Tmp || ',"pati_idcard":"' || Nvl(r_����.���֤��, '') || '"';
        v_Tmp := v_Tmp || ',"cardtype_id":' || Nvl(r_����.�����id, 0);
        v_Tmp := v_Tmp || ',"card_no":"' || Nvl(r_����.����, '') || '"';
        v_Tmp := v_Tmp || ',"card_qrcode":"' || Nvl(r_����.��ά��, '') || '"';
        v_Tmp := v_Tmp || ',"card_passwod":"' || Nvl(r_����.����, '') || '"';
        v_Tmp := v_Tmp || ',"cardtype_name":"' || Nvl(r_����.���������, '') || '"';
        If Nvl(n_��ѯ����, 0) = 2 Then
          v_Tmp := v_Tmp || '}';
        End If;
        If Nvl(n_��ѯ����, 0) <> 2 Then
          v_Tmp := v_Tmp || ',"cardtype_cardlen":' || Nvl(r_����.���ų���, 0);
          v_Tmp := v_Tmp || ',"card_statu":' || Nvl(r_����.״̬, 0);
          v_Tmp := v_Tmp || ',"loscard_creator":"' || Nvl(r_����.��ʧ��, '') || '"';
          v_Tmp := v_Tmp || ',"loscard_time":"' || Nvl(r_����.��ʧʱ��, '') || '"';
          v_Tmp := v_Tmp || ',"loscard_days":' || Nvl(r_����.��Ч����, 0);
          v_Tmp := v_Tmp || ',"loscard_mode":"' || Nvl(r_����.��ʧ��ʽ, '') || '"';
          v_Tmp := v_Tmp || ',"sendcard_oper":"' || Nvl(r_����.������, '') || '"';
          v_Tmp := v_Tmp || ',"end_time":"' || Nvl(r_����.��ֹʹ��ʱ��, '') || '"}';
        End If;
      End If;
      If Length(v_Tmp) > 20000 Then
        Json_Out := Json_Out || v_Tmp;
        v_Tmp    := ',';
      End If;
    Else
      v_List := v_List || ',{"cardtype_id":' || Nvl(r_����.�����id, 0) || '}';
    End If;
  End Loop;
  If Nvl(n_��ѯ����, 0) = 1 Then
    v_List   := Substr(v_List, 2);
    Json_Out := Json_Out || v_List || ']}}';
  Else
    If v_Tmp = ',' Then
      Json_Out := Json_Out || ']}}';
    Else
      v_Tmp    := Substr(v_Tmp, 2);
      Json_Out := Json_Out || v_Tmp || ']}}';
    End If;
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpaticardinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpaticardno
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���id��ȱʡ������ȡ���˵Ŀ���Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_ids            C 1  ����ids,��������Զ��ŷָ�
  --   card_type_id        N 1 �����id
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    card_list[]
  --       vcard_no       C 1 ���￨��
  --       pati_id        N 1 ����id
  ---------------------------------------------------------------------------
  l_����id   t_Strlist := t_Strlist();
  c_����ids  Clob;
  n_�����id Number;
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_List     Varchar2(32767);
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  c_����ids  := j_Json.Get_Clob('pati_ids');
  n_�����id := j_Json.Get_Number('card_type_id');

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","card_list":[';
  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      l_����id.Extend;
      l_����id(l_����id.Count) := c_����ids;
      c_����ids := Null;
    Else
      l_����id.Extend;
      l_����id(l_����id.Count) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
    End If;
  End Loop;
  For I In 1 .. l_����id.Count Loop
    For R In (Select f_List2str(Cast(Collect(g.����) As t_Strlist)) As ����, g.����id
              From ����ҽ�ƿ���Ϣ G, ҽ�ƿ���� H, (Select Column_Value As ����id From Table(f_Num2list(l_����id(I)))) A
              Where g.����id = a.����id And g.�����id = h.Id And g.״̬ = 0 And h.Id = n_�����id
              Group By g.����id) Loop
      Zljsonputvalue(v_List, 'pati_id', r.����id, 1, 1);
      Zljsonputvalue(v_List, 'vcard_no', r.����, 0, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  End Loop;
  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpaticardno;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiextendinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ������Ϣ�ӱ�
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id             N 1 ����id
  --    info_names          C 1 ��Ϣ��������ö���
  --    visit_id            N 0 ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    slave_list[]        C     ������Ϣ�ӱ��б�
  --     info_name          C 1   ��Ϣ��
  --     info_value         C 1   ��Ϣֵ
  --     visit_id           N 1   ����id
  ---------------------------------------------------------------------------

  j_Json    Pljson;
  j_Jsonin  Pljson;
  v_List    Varchar2(32767);
  n_����id  ������Ϣ�ӱ�.����id%Type;
  v_��Ϣ��s Varchar2(32680);
  n_����id  ������Ϣ�ӱ�.����id%Type;
Begin

  --�������
  j_Jsonin  := Pljson(Json_In);
  j_Json    := j_Jsonin.Get_Pljson('input');
  n_����id  := j_Json.Get_Number('pati_id');
  v_��Ϣ��s := j_Json.Get_String('info_names');
  n_����id  := j_Json.Get_Number('visit_id');
  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('ʧ�ܣ�δ���벡��id��');
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","slave_list":[';
  If Nvl(v_��Ϣ��s, '-') <> '-' Then
    For r_��Ϣ�ӱ� In (Select a.��Ϣ��, a.��Ϣֵ, a.����id
                   From ������Ϣ�ӱ� A, Table(f_Str2list(v_��Ϣ��s)) B
                   Where a.����id = n_����id And a.��Ϣ�� = b.Column_Value And Nvl(a.����id, 0) = Nvl(n_����id, 0)) Loop
      Zljsonputvalue(v_List, 'info_name', r_��Ϣ�ӱ�.��Ϣ��, 0, 1);
      Zljsonputvalue(v_List, 'info_value', r_��Ϣ�ӱ�.��Ϣֵ);
      Zljsonputvalue(v_List, 'visit_id', r_��Ϣ�ӱ�.����id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  Else
    For r_��Ϣ�ӱ� In (Select Upper(��Ϣ��) ��Ϣ��, ��Ϣֵ, ����id
                   From ������Ϣ�ӱ�
                   Where ����id = n_����id And (����id = n_����id Or ����id Is Null)
                   Order By Nvl(����id, 999999999)) Loop
      Zljsonputvalue(v_List, 'info_name', r_��Ϣ�ӱ�.��Ϣ��, 0, 1);
      Zljsonputvalue(v_List, 'info_value', r_��Ϣ�ӱ�.��Ϣֵ);
      Zljsonputvalue(v_List, 'visit_id', r_��Ϣ�ӱ�.����id, 1, 2);
      If Length(v_List) > 20000 Then
        Json_Out := Json_Out || v_List;
        v_List   := ',';
      End If;
    End Loop;
  End If;
  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiextendinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatifamilymember
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID����ȡ�ò��˵ļ�����Ա��Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id              N   1  ����ID
  --    query_type           N   1  ��ѯ���ͣ�0-ֻ���ؼ�����Ա����id��1-��ѯ������Ա�Ļ�����Ϣ

  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    family_list[]        C       ������Ա:���˼���
  --      pati_id            N   1   ����ID:����id
  --      pati_relation      C   1   ��ϵ
  --      pati_name          C   1   ����
  --      pati_sex           C   1   �Ա�
  --      pati_age           C   1   ����
  --      pati_birthdate     C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
  --      pati_nation        C   1   ����
  --      pati_idcard        C   1   ���֤��
  --      family_id          N   1   ����id
  --      visit_cardno       C   1   ���￨��
  --      state              N   1   ״̬
  ---------------------------------------------------------------------------

  n_��ѯ���� Number(1);
  n_����id   ������Ϣ.����id%Type;
  v_List     Varchar2(32767);
  j_Json     Pljson;  
  j_Jsonin   Pljson;
Begin

  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ѯ���� := Nvl(j_Json.Get_Number('query_type'), 0);

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ���벡��ID"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","family_list":[';

  For r_������Ϣ In (Select /*+cardinality(B,10)*/
                  b.����id, a.���￨��, a.����id, b.��ϵ, a.����, a.�Ա�, a.����, a.��������, a.����, a.���֤��, 1 As ״̬
                 From ������Ϣ A, ���˼��� B
                 Where a.����id = b.����id And b.����id = n_����id And
                       (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss'))) Loop
    --      pati_id            N   1   ����ID:����id
    --      pati_relation      C   1   ��ϵ
    --      pati_name          C   1   ����
    --      pati_sex           C   1   �Ա�
    --      pati_age           C   1   ����
    --      pati_birthdate     C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
    --      pati_nation        C   1   ����
    --      pati_idcard        C   1   ���֤��
  
    If n_��ѯ���� = 1 Then
      Zljsonputvalue(v_List, 'pati_id', r_������Ϣ.����id, 1, 1);
      Zljsonputvalue(v_List, 'pati_relation', r_������Ϣ.��ϵ);
      Zljsonputvalue(v_List, 'pati_name', r_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_sex', r_������Ϣ.�Ա�);
      Zljsonputvalue(v_List, 'pati_age', r_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_birthdate', To_Char(r_������Ϣ.��������, 'yyyy-mm-dd hh24:mi:ss'));
      Zljsonputvalue(v_List, 'pati_nation', r_������Ϣ.����);
      Zljsonputvalue(v_List, 'pati_idcard', r_������Ϣ.���֤��);
      Zljsonputvalue(v_List, 'family_id', r_������Ϣ.����id, 1);
      Zljsonputvalue(v_List, 'visit_cardno', r_������Ϣ.���￨��);
      Zljsonputvalue(v_List, 'state', r_������Ϣ.״̬, 1, 2);
    Else
      v_List := v_List || ',{"pati_id":' || Nvl(r_������Ϣ.����id, 0) || '}';
    End If;
  End Loop;
  If n_��ѯ���� = 1 Then
    Json_Out := Json_Out || v_List || ']}}';
    Return;
  Else
    v_List   := Substr(v_List, 2);
    Json_Out := Json_Out || v_List || ']}}';
    Return;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatifamilymember;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ָ��������ȡ������Ϣ�Ĳ���ID
  --��Σ�Json_In:��ʽ
  --  input
  --     card_find             C
  --         cardtype_id       N  1  ҽ�ƿ����ID:=0ʱ����ʾģ������
  --         card_no           C  1  ����
  --         qrcode            C     ��ά��
  --         is_check_usetime  N  1  �Ƿ���ʹ��ʱ��:1-���;0-�����
  --         is_check_stop     N  1  �Ƿ���ͣ�û��ʧ:1-���;0-�����
  --     comminuty_find        C
  --        comminuty_num      N  1  �������
  --        comminuty_code     C     ������
  --     other_cons_find       C
  --        find_name          C  1  ���ҵ�����
  --        find_text          C  1  ���ҵ��ı�
  --        pati_id            N     �д˽ڵ�ʱ���������ſ��˲���ID
  --     is_stop               N     ����ͣ�� 0-����ͣ�õ� 1-��ͣ�õ�

  --����: Json_Out,��ʽ����
  --  output
  --    code                        N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                     C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  pati_list[]                   C  1  �����б�,ģ������ʱ�����ܴ��ڶ��
  --         cardtype_id            N  1  �����ID
  --         pati_id                N  1  ����ID:δ�ҵ�ʱҲ�ɹ�������0
  --         card_pwd               C  1  ����
  --         pati_pageid            N  1  ��ҳID
  --         enduse_time            C  1  ��ֹʹ��ʱ��:yyyy-mm-dd hh24mi:ss
  --         card_status            N  1  ��ǰ��״̬��0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��;3-ʧЧ��������ҽ�ƿ���Ϣ.��ֹʹ��ʱ�䵽��ʱ���ظ�״̬����������ʹ�ã�

  ---------------------------------------------------------------------------
  j_Json           Pljson;
  j_Jsonin         Pljson;
  j_Tmp            Pljson;
  n_����id         ������Ϣ.����id%Type;
  n_�����id       ����ҽ�ƿ���Ϣ.�����id%Type;
  v_����           ����ҽ�ƿ���Ϣ.����%Type;
  n_��ҳid         ������Ϣ.��ҳid%Type;
  v_��ά��         Varchar2(500);
  n_����           Number(5);
  v_������         Varchar2(500);
  v_��������       Varchar2(50);
  v_����ֵ         Varchar2(500);
  n_Find           Number(2);
  v_Err_Msg        Varchar2(500);
  n_�����Ч       Number(2);
  n_���ͣ�ü���ʧ Number(2);
  n_ͣ��           Number;
  n_�ſ�����id     Number;
  v_List           Varchar2(32767);
  --��װʧ��ʱ���ص�����
  Function Get_Err_Message
  (
    Message_In    Varchar2,
    ��ǰ��״̬_In ����ҽ�ƿ���Ϣ.״̬%Type := 0
  ) Return Varchar2 Is
    j_Out Varchar2(32767);
  Begin
    j_Out := '{"output":{"code":0,"message":"' || Zljsonstr(Message_In) || '","card_status":' || ��ǰ��״̬_In || '}}';
    Return j_Out;
  End Get_Err_Message;

  --��װ�ɹ���Ϣ
  Function Get_Succes_Message
  (
    �����id_In     ҽ�ƿ����.Id%Type,
    ����id_In       ������Ϣ.����id%Type,
    ��ҳid_In       ������Ϣ.��ҳid%Type,
    ����_In         ����ҽ�ƿ���Ϣ.����%Type := Null,
    ��ֹʹ��ʱ��_In Varchar2 := Null,
    ��ǰ��״̬_In   ����ҽ�ƿ���Ϣ.״̬%Type := Null
  ) Return Varchar2 Is
    j_Out  Varchar2(32767);
    v_List Varchar2(32767);
  Begin
    v_List := '';
    If Nvl(����id_In, 0) <> 0 Then
      v_List := '{"cardtype_id":' || Nvl(�����id_In, 0) || ',';
      v_List := v_List || '"pati_id":' || Nvl(����id_In, 0) || ',';
      v_List := v_List || '"pati_pageid":' || Nvl(��ҳid_In, 0) || ',';
      v_List := v_List || '"card_pwd":"' || Nvl(����_In, '') || '",';
      v_List := v_List || '"enduse_time":"' || Nvl(��ֹʹ��ʱ��_In, '') || '",';
      v_List := v_List || '"card_status":' || Nvl(��ǰ��״̬_In, 0) || '}';
    End If;
  
    j_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[' || v_List || ']}}';
    Return j_Out;
  End Get_Succes_Message;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_ͣ��   := j_Json.Get_Number('is_stop');
  --1.��ҽ�ƿ�����Ϣ��ѯ
  If j_Json.Exist('card_find') Then
    --              cardtype_id       N  1  ҽ�ƿ����ID:=0ʱ����ʾģ������
    --              card_no           C  1  ����
    --              qrcode            C     ��ά��
    --              is_check_usetime  N  1  �Ƿ���ʹ��ʱ��:1-���;0-�����
    --              is_check_stop     N  1  �Ƿ���ͣ�û��ʧ:1-���;0-�����
    j_Tmp            := Pljson();
    j_Tmp            := j_Json.Get_Pljson('card_find');
    n_�����id       := j_Tmp.Get_Number('cardtype_id');
    v_����           := j_Tmp.Get_String('card_no');
    v_��ά��         := j_Tmp.Get_String('qrcode');
    n_�����Ч       := j_Tmp.Get_Number('is_check_usetime');
    n_���ͣ�ü���ʧ := j_Tmp.Get_Number('is_check_stop');
  
    --1.1 �������id����
    If n_�����id <> 0 Then
      --1.1.1��ҽ�ƿ����ID������
      If v_���� Is Not Null Then
      
        For c_���� In (Select a.�����id, a.����id, ����, a.״̬,
                            Nvl(��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(b.��Ч����, 0) As ��ʧʱ��, Sysdate As ��ǰʱ��,
                            a.��ֹʹ��ʱ��, c.��ҳid
                     From ����ҽ�ƿ���Ϣ A, ������Ϣ C, ҽ�ƿ���ʧ��ʽ B
                     Where a.����id = c.����id And a.�����id = n_�����id And a.���� = v_���� And a.��ʧ��ʽ = b.����(+) And c.ͣ��ʱ�� Is Null)
        
         Loop
        
          If c_����.��ֹʹ��ʱ�� Is Not Null And Nvl(n_�����Ч, 0) = 1 Then
            If c_����.��ֹʹ��ʱ�� <= c_����.��ǰʱ�� Then
              v_Err_Msg := '����Ϊ' || v_���� || '��ʧЧ';
              Json_Out  := Get_Err_Message(v_Err_Msg, 3);
              Return;
            End If;
          End If;
        
          --0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
          If Nvl(c_����.״̬, 0) = 1 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
            --��ʧ���
            If Nvl(c_����.��ʧʱ��, c_����.��ǰʱ�� - 1) < c_����.��ǰʱ�� Then
              v_Err_Msg := '����Ϊ' || v_���� || '�ѹ�ʧ!';
              Json_Out  := Get_Err_Message(v_Err_Msg, c_����.״̬);
              Return;
            End If;
          End If;
        
          If Nvl(c_����.״̬, 0) = 2 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
            --ͣ�ü��
            v_Err_Msg := '����Ϊ' || v_���� || '��ͣ��!';
            Json_Out  := Get_Err_Message(v_Err_Msg, c_����.״̬);
            Return;
          End If;
          Json_Out := Get_Succes_Message(n_�����id, c_����.����id, c_����.��ҳid, c_����.����,
                                         To_Char(c_����.��ֹʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss'), Nvl(c_����.״̬, 0));
          Return;
        End Loop;
      End If;
    
      --1.1.2 ����ά���ѯ
      For c_���� In (Select a.�����id, a.����id, ����, a.״̬,
                          Nvl(��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(b.��Ч����, 0) As ��ʧʱ��, Sysdate As ��ǰʱ��,
                          a.��ֹʹ��ʱ��, c.��ҳid
                   From ����ҽ�ƿ���Ϣ A, ������Ϣ C, ҽ�ƿ���ʧ��ʽ B
                   Where a.����id = c.����id And a.�����id = n_�����id And a.��ά�� = v_��ά�� And a.��ʧ��ʽ = b.����(+) And c.ͣ��ʱ�� Is Null)
      
       Loop
      
        If c_����.��ֹʹ��ʱ�� Is Not Null And Nvl(n_�����Ч, 0) = 1 Then
          If c_����.��ֹʹ��ʱ�� <= c_����.��ǰʱ�� Then
            v_Err_Msg := '����Ϊ' || v_���� || '��ʧЧ';
            Json_Out  := Get_Err_Message(v_Err_Msg, 3);
            Return;
          End If;
        End If;
      
        --0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
        If Nvl(c_����.״̬, 0) = 1 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
          --��ʧ���
          If Nvl(c_����.��ʧʱ��, c_����.��ǰʱ�� - 1) < c_����.��ǰʱ�� Then
            v_Err_Msg := '����Ϊ' || v_���� || '�ѹ�ʧ!';
            Json_Out  := Get_Err_Message(v_Err_Msg, c_����.״̬);
            Return;
          End If;
        End If;
      
        If Nvl(c_����.״̬, 0) = 2 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
          --ͣ�ü��
          v_Err_Msg := '����Ϊ' || v_���� || '��ͣ��!';
          Json_Out  := Get_Err_Message(v_Err_Msg, c_����.״̬);
          Return;
        End If;
        Json_Out := Get_Succes_Message(n_�����id, c_����.����id, c_����.��ҳid, c_����.����,
                                       To_Char(c_����.��ֹʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss'), Nvl(c_����.״̬, 0));
        Return;
      
      End Loop;
    
      Json_Out := Get_Succes_Message(Null, Null, Null);
      Return;
    
    End If;
  
    --1.2 .ģ��ģ��
    --1.2.1 ������ģ������
    If v_���� Is Not Null Then
    
      v_Err_Msg := Null;
      For c_���� In (Select a.����id, a.�����id, a.����, a.����, a.״̬, a.��ʧʱ��, a.��ʧ��ʽ, a.��ʧ��, a.��������, a.������, a.��ֹʹ��ʱ��, a.��ά��,
                          d.��ҳid
                   From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���ʧ��ʽ B, ҽ�ƿ���� C, ������Ϣ D
                   Where a.�����id = c.Id And Nvl(c.�Ƿ�ģ������, 0) = 1 And a.����id = d.����id And a.���� = v_���� And
                         a.��ʧ��ʽ = b.����(+) And Nvl(c.�Ƿ�����, 0) = 1 And d.ͣ��ʱ�� Is Null
                   Order By a.״̬) Loop
        n_Find := 1;
        If c_����.��ֹʹ��ʱ�� Is Not Null And Nvl(n_�����Ч, 0) = 1 Then
          If c_����.��ֹʹ��ʱ�� <= Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '����Ϊ' || v_���� || '��ʧЧ';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        --0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
        If Nvl(c_����.״̬, 0) = 1 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
          --��ʧ���
          If Nvl(c_����.��ʧʱ��, Sysdate - 1) < Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '����Ϊ' || v_���� || '�ѹ�ʧ!';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        If Nvl(c_����.״̬, 0) = 2 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
          --ͣ�ü��
          If v_Err_Msg Is Null Then
            v_Err_Msg := '����Ϊ' || v_���� || '��ͣ��!';
          End If;
          n_Find := 0;
        End If;
      
        If n_Find = 1 Then
          Zljsonputvalue(v_List, 'cardtype_id', Nvl(c_����.�����id, 0), 1, 1);
          Zljsonputvalue(v_List, 'pati_id', Nvl(c_����.����id, 0), 1);
          Zljsonputvalue(v_List, 'pati_pageid', Nvl(c_����.��ҳid, 0), 1);
          Zljsonputvalue(v_List, 'card_pwd', Nvl(c_����.����, ''));
          Zljsonputvalue(v_List, 'enduse_time', Nvl(To_Char(c_����.��ֹʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss'), ''));
          Zljsonputvalue(v_List, 'card_status', Nvl(c_����.״̬, 0), 1, 2);
        End If;
      End Loop;
    
      If v_List Is Null Then
        Json_Out := Get_Err_Message(v_Err_Msg);
        Return;
      End If;
    
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[' || v_List || ']}}';
      Return;
    End If;
  
    --1.2.2.����ά����к�������
    If v_��ά�� Is Not Null Then
    
      v_Err_Msg := Null;
      For c_���� In (Select a.����id, a.�����id, a.����, a.����, a.״̬, a.��ʧʱ��, a.��ʧ��ʽ, a.��ʧ��, a.��������, a.������, a.��ֹʹ��ʱ��, a.��ά��,
                          d.��ҳid
                   From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���ʧ��ʽ B, ҽ�ƿ���� C, ������Ϣ D
                   Where a.�����id = c.Id And Nvl(c.�Ƿ�ģ������, 0) = 1 And a.����id = d.����id And a.��ά�� = v_��ά�� And
                         a.��ʧ��ʽ = b.����(+) And Nvl(c.�Ƿ�����, 0) = 1 And d.ͣ��ʱ�� Is Null
                   Order By a.״̬) Loop
        n_Find := 1;
        If c_����.��ֹʹ��ʱ�� Is Not Null And Nvl(n_�����Ч, 0) = 1 Then
          If c_����.��ֹʹ��ʱ�� <= Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '����Ϊ' || v_���� || '��ʧЧ';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        --0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
        If Nvl(c_����.״̬, 0) = 1 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
          --��ʧ���
          If Nvl(c_����.��ʧʱ��, Sysdate - 1) < Sysdate Then
            If v_Err_Msg Is Null Then
              v_Err_Msg := '����Ϊ' || v_���� || '�ѹ�ʧ!';
            End If;
            n_Find := 0;
          End If;
        End If;
      
        If Nvl(c_����.״̬, 0) = 2 And Nvl(n_���ͣ�ü���ʧ, 0) = 1 Then
          --ͣ�ü��
          If v_Err_Msg Is Null Then
            v_Err_Msg := '����Ϊ' || v_���� || '��ͣ��!';
          End If;
          n_Find := 0;
        End If;
      
        If n_Find = 1 Then
          Zljsonputvalue(v_List, 'cardtype_id', Nvl(c_����.�����id, 0), 1, 1);
          Zljsonputvalue(v_List, 'pati_id', Nvl(c_����.����id, 0), 1);
          Zljsonputvalue(v_List, 'pati_pageid', Nvl(c_����.��ҳid, 0), 1);
          Zljsonputvalue(v_List, 'card_pwd', Nvl(c_����.����, ''));
          Zljsonputvalue(v_List, 'enduse_time', Nvl(To_Char(c_����.��ֹʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss'), ''));
          Zljsonputvalue(v_List, 'card_status', Nvl(c_����.״̬, 0), 1, 2);
        End If;
      End Loop;
    
      If v_List Is Null Then
        Json_Out := Get_Err_Message(v_Err_Msg);
        Return;
      End If;
    
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[' || v_List || ']}}';
      Return;
    End If;
  
    Return;
  
    v_Err_Msg := 'δ����ҽ�ƿ���Ϣ����������';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;

  --2.�������������Ҳ���
  If j_Json.Exist('comminuty_find') Then
    --            comminuty_num       N  1  �������
    --            comminuty_code      C     ������
    j_Tmp    := Pljson();
    j_Tmp    := j_Json.Get_Pljson('comminuty_find');
    n_����   := j_Tmp.Get_Number('comminuty_num');
    v_������ := j_Tmp.Get_String('comminuty_code');
  
    If Nvl(n_����, 0) = 0 Or Nvl(v_������, '-') = '-' Then
      v_Err_Msg := 'δ����������Ϣ����������';
      Json_Out  := Get_Err_Message(v_Err_Msg);
      Return;
    End If;
  
    Select Max(a.����id), Max(b.��ҳid)
    Into n_����id, n_��ҳid
    From ����������Ϣ A, ������Ϣ B
    Where a.����id = b.����id And a.���� = n_���� And a.������ = v_������ And
          (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And b.ͣ��ʱ�� Is Null));
  
    Json_Out := Get_Succes_Message(0, n_����id, n_��ҳid, '');
    Return;
  
  End If;

  --3.��������ʽ����
  If Not j_Json.Exist('other_cons_find') Then
    v_Err_Msg := 'δ������Ϣ��ѯ����������';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  j_Tmp        := Pljson();
  j_Tmp        := j_Json.Get_Pljson('other_cons_find');
  v_��������   := j_Tmp.Get_String('find_name');
  v_����ֵ     := j_Tmp.Get_String('find_text');
  n_�ſ�����id := j_Tmp.Get_Number('pati_id');
  If Nvl(v_��������, '-') = '-' Or Nvl(v_����ֵ, '-') = '-' Then
    v_Err_Msg := 'δ������Ϣ��ѯ����������';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  v_�������� := Replace(v_��������, ' ', '');
  If v_�������� = 'IC��' Or v_�������� = 'IC����' Then
    Select Max(����id), Max(��ҳid)
    Into n_����id, n_��ҳid
    From ������Ϣ
    Where Ic���� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
  Elsif v_�������� = '���֤' Or v_�������� = '���֤��' Or v_�������� = '�������֤' Then
    Select Max(����id), Max(��ҳid)
    Into n_����id, n_��ҳid
    From ������Ϣ
    Where ���֤�� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
  Elsif v_�������� = 'ҽ����' Or v_�������� = 'ҽ��֤��' Then
    --ҽ����֧��ģ�����ң�������ҽ��
    If Instr(v_����ֵ, '%') > 0 Then
      Select Max(����id), Max(��ҳid)
      Into n_����id, n_��ҳid
      From ������Ϣ
      Where ҽ���� Like v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
    Else
      Select Max(����id), Max(��ҳid)
      Into n_����id, n_��ҳid
      From ������Ϣ
      Where ҽ���� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
    End If;
  Elsif v_�������� = '�ֻ���' Then
    Select Max(����id), Max(��ҳid)
    Into n_����id, n_��ҳid
    From ������Ϣ
    Where �ֻ��� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null)) And
          (Nvl(n_�ſ�����id, 0) = 0 Or ����id <> Nvl(n_�ſ�����id, 0));
  Elsif v_�������� = '�����' Then
    Select Max(����id), Max(��ҳid)
    Into n_����id, n_��ҳid
    From ������Ϣ
    Where ����� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null)) And
          (Nvl(n_�ſ�����id, 0) = 0 Or ����id <> Nvl(n_�ſ�����id, 0));
  Elsif v_�������� = 'סԺ��' Then
    Select Max(����id), Max(��ҳid)
    Into n_����id, n_��ҳid
    From ������Ϣ
    Where סԺ�� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null)) And
          (Nvl(n_�ſ�����id, 0) = 0 Or ����id <> Nvl(n_�ſ�����id, 0));
  Elsif Upper(v_��������) = Upper('����ID') Then
    Select Max(����id), Max(��ҳid) Into n_����id, n_��ҳid From ������Ϣ Where ����id = v_����ֵ;
  Elsif v_�������� = '������' Then
    Select Max(����id), Max(��ҳid)
    Into n_����id, n_��ҳid
    From ������Ϣ
    Where ������ = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
  Elsif v_�������� = '���￨��' Then
    Select Max(����id), Max(��ҳid)
    Into n_����id, n_��ҳid
    From ������Ϣ
    Where ���￨�� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
  Elsif v_�������� = '����' Then
    If Instr(v_����ֵ, '%') > 0 Then
      Select Max(����id), Max(��ҳid)
      Into n_����id, n_��ҳid
      From ������Ϣ
      Where ���� Like v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
    Else
      Select Max(����id), Max(��ҳid)
      Into n_����id, n_��ҳid
      From ������Ϣ
      Where ���� = v_����ֵ And (Nvl(n_ͣ��, 0) = 0 Or (Nvl(n_ͣ��, 0) = 1 And ͣ��ʱ�� Is Null));
    End If;
  Else
    --��֧�ֵķ�ʽ
    v_Err_Msg := '��֧��' || v_�������� || '��ʽ���Ҳ���!';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  If Nvl(n_����id, 0) <> 0 Then
    Json_Out := Get_Succes_Message(0, n_����id, n_��ҳid, '');
    Return;
  End If;

  Json_Out := Get_Succes_Message(0, n_����id, n_��ҳid, '');

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiid;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiidsbyrange
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ָ��������ȡ������Ϣ�Ĳ���ID
  --��Σ�Json_In:��ʽ
  --    input
  --      query_condition C 1 ��ѯ����
  --      ctt_unit_id     N 1 ��ͬ��λID����ѯָ����ͬ��λ�����ﲡ��
  --����: Json_Out,��ʽ����
  --  output
  --    code              N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_ids          C  1  ����IDs������ƴ��
  ---------------------------------------------------------------------------
  j_Json    Pljson;
  j_Jsonin  Pljson;
  v_����ֵ  Varchar2(3000);
  v_Temp    Varchar2(500);
  v_����ids Varchar2(32767);

  n_��ͬ��λid ������Ϣ.��ͬ��λid%Type;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  If j_Json.Exist('query_condition') Then
    v_����ֵ := j_Json.Get_String('query_condition');
    If v_����ֵ Is Null Then
      Json_Out := Zljsonout('δ�����ѯ���������飡');
      Return;
    End If;
  
    Select LTrim(v_����ֵ, '0123456789') Into v_Temp From Dual;
    If v_Temp Is Null Then
      Select f_List2str(Cast(Collect(To_Char(a.����id)) As t_Strlist))
      Into v_����ids
      From ������Ϣ A
      Where a.����� = To_Number(v_����ֵ) Or a.���￨�� = v_����ֵ Or a.���֤�� = v_����ֵ Or a.Ic���� = v_����ֵ;
    Else
      Select f_List2str(Cast(Collect(To_Char(a.����id)) As t_Strlist))
      Into v_����ids
      From ������Ϣ A
      Where a.���￨�� = v_����ֵ Or a.���֤�� = v_����ֵ Or a.Ic���� = v_����ֵ;
    End If;
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_ids":"' || v_����ids || '"}}';
    Return;
  End If;

  --����ͬ��λID��ȡ
  If j_Json.Exist('ctt_unit_id') Then
    n_��ͬ��λid := j_Json.Get_Number('ctt_unit_id');
    If Nvl(n_��ͬ��λid, 0) = 0 Then
      Json_Out := Zljsonout('δ�����ͬ��λID�����飡');
      Return;
    End If;
    v_����ids := Null;
  
    For c_���� In (Select Distinct a.����id From ������Ϣ A Where a.��ͬ��λid = n_��ͬ��λid And a.��ǰ����id Is Null) Loop
      v_����ids := v_����ids || ',' || c_����.����id;
    End Loop;
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_ids":"' || Substr(v_����ids, 2) || '"}}';
    Return;
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiidsbyrange;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiinfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����id  ����ID<>0ʱ����ѯ�б��е�������Ч
  --      query_type        N 1 ��ѯ����:�磺0-����;1-����+��ϵ��;2-����
  --      query_card        N 1 �Ƿ��������Ϣ:1-����ҽ�ƿ�;0-������ҽ�ƿ�
  --      query_family      N 1 �Ƿ��������:1-����������Ϣ��0-������������Ϣ
  --      query_drug        N 1 �Ƿ��������ҩ��:1-������0-������
  --      query_immune      N 1 �Ƿ����������:1-����;0-������
  --      query_insurance_pwd C  �Ƿ����ҽ������:1-����;0-������
  --      query_cons_list   C 1 ��ѯ����:����ѡ��һ���������в�ѯ����And��ϵ),ֻ��һ��
  --        pati_ids        C   ����IDs:����ö���
  --        pati_name       C   ����:���Դ�%�ֺű������ƥ��
  --        outpatient_num  C   �����
  --        inpatient_num   C   סԺ��
  --        pati_idcard     C   ���֤��
  --        contacts_idcard C   ��ϵ�����֤��
  --        cardtype_id     N   ҽ�ƿ����ID
  --        medc_card_name  N   ҽ�ƿ�����
  --        card_no         C   ����
  --        qrcode          C   ��ά��
  --        iccard_no       C   Ic����
  --        visit_card      C   ���￨��
  --        insurance_num   C   ҽ����
  --        qrspt_statu     C   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
  --        phone_number    C   �ֻ���
  --        pati_bed        C   ��ǰ����
  --        dept_id         N   ��ǰ����ID
  --        search_days     N   �д˽ڵ�ʱ��ָ�������������Ҳ���(��������ģ������)
  --����      json
  --output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_list[]                 ������Ϣ�б�
  --    pati_id             N   1   ����id
  --    pati_pageid         N   1   ��ҳid��������Ϣ.��ҳID
  --    pati_name           C   1   ����
  --    pati_sex            C   1   �Ա�
  --    pati_age            C   1   ����
  --    pati_birthdate      C   1   �������ڣ�yyyy-mm-dd hh24:mi:ss
  --    fee_category        C   1   �ѱ�
  --    outpatient_num      C   1   �����
  --    inpatient_num       C   1   סԺ��
  --    mdlpay_mode_name    C   1   ҽ�Ƹ��ʽ����
  --    mdlpay_mode_code    C   1   ҽ�Ƹ��ʽ����
  --    pati_nation         C   1   ����
  --    insurance_num       C   1   ҽ����
  --    pati_idcard         C   1   ���֤��
  --    vcard_no            C   1   ���￨��
  --    iccard_no           C   1   Ic����
  --    health_num          C   1   ������
  --    inp_times           N   1   סԺ����
  --    pati_education      C   1   ѧ��
  --    ocpt_name           C   1   ְҵ
  --    pati_identity       C   1   ���
  --    ntvplc_name         C   1   ����
  --    country_name        C   1   ����
  --    pati_marital_cstatus    C   1   ����״��
  --    pat_home_addr           C   1   ��ͥ��ַ
  --    pat_home_phno           C   1   ��ͥ�绰
  --    pat_home_postcode   C   1   ��ͥ��ַ�ʱ�
  --    pati_area           C   1   ����
  --    pati_birthplace     C   1   �����ص�
  --    pat_hous_addr       C   1   ���ڵ�ַ
  --    pat_hous_postcode   C   1   ���ڵ�ַ�ʱ�
  --    emp_name            C   1   ������λ����
  --    emp_phno            C   1   ��λ�绰
  --    emp_postcode        C   1   ��λ�ʱ�
  --    emp_bank_name       C   1   ��λ������
  --    emp_bank_accnum     C   1   ��λ�ʺ�
  --    emp_addr             C   1   ��λ��ַ
  --    ctt_unit_id         N   1   ��ͬ��λID
  --    phone_number        C   1   �ֻ���
  --    pati_bed            C   1   ��ǰ����
  --    pati_type           C   1   ��������(��ͨ��ҽ��������)
  --    insurance_type      C   1   ����
  --    insurance_name      C   1   ��������
  --    pati_wardarea_id    N   1   ��ǰ����id
  --    pati_wardarea_name  C   1   ��ǰ��������
  --    pati_dept_id        N   1   ��ǰ����id
  --    pati_dept_name      C   1   ��ǰ��������
  --    adta_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
  --    adtd_time           C   1   ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
  --    contacts_name       C   1   ��ϵ������
  --    contacts_relation   C   1   ��ϵ�˹�ϵ
  --    contacts_idcard     C   1   ��ϵ�����֤��
  --    contacts_addr       C   1   ��ϵ�˵�ַ
  --    contacts_phno       C   1   ��ϵ�˵绰
  --    pat_grdn_name       C   1   �໤��
  --    cert_no_other       C   1   ����֤��
  --    is_inhspt            C   1   �Ƿ���Ժ:1-��Ժ ;0-����Ժ
  --    pati_show_color      N   1   ������ʾ��ɫ
  --    visit_room           C   1   ��������
  --    visit_statu          N   1   ����״̬
  --    visit_time           C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --    create_time          C   1   �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    pati_email           C   1   email
  --    pati_qq              C   1   qq
  --    card_captcha         C   1  ����֤��
  --    insurance_pwd        C       ҽ������
  --    family_list[]        C   1   ������Ա:���˼���() query_family=1����
  --        family_id        N   1   ����id  query_family=1
  --        family_relation  C   1   ��ϵ
  --    drug_list[]          C   1   ����ҩ���б�    query_drug=1ʱ����
  --        pat_algc_cadn_id N   1   ����ҩƷID
  --        pat_algc_cadn    C   1   ����ҩ������
  --        allergy_info     C   1   ��ÿҩ�ﷴӦ
  --    immune_list[]        C   1   ���������б�    query_immune=1ʱ����
  --        vaccinate_time   C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --        vaccinate_name   C   1   ��������
  --    card_list[]          C   1   ����ҽ�ƿ���Ϣ�б�(��������д����˿����ID�ģ��򷵻ظÿ����Ŀ���Ϣ)  query_card=1ʱ����
  --        cardtype_id      N   1   ҽ�ƿ����ID
  --        card_no          C   1   ����
  --        card_pwd         C   1   ����
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Json    Pljson;
  o_Json    Pljson;
  j_Jsonin  Pljson;

  n_�����id   ҽ�ƿ����.Id%Type;
  v_ҽ�ƿ����� ҽ�ƿ����.����%Type;
  n_����id     ����ҽ�ƿ���Ϣ.����id%Type;
  n_�����     ������Ϣ.�����%Type;
  n_סԺ��     ������Ϣ.סԺ��%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_��ɫ       ��������.��ɫ%Type;

  v_����           Varchar2(100);
  v_����           Varchar2(1000);
  v_���￨��       ������Ϣ.���￨��%Type;
  v_��ά��         Varchar2(1000);
  v_�ֻ���         Varchar2(50);
  v_��ϵ�����֤�� Varchar2(50);
  v_ҽ����         Varchar2(30);
  v_����           Varchar2(30);
  v_Ic����         Varchar2(100);
  n_��ѯ����       Number(2);
  n_��������       Number(10);
  n_��ѯסԺ״̬   Number(2);
  n_����id         ������Ϣ.��ǰ����id%Type;

  n_�Ƿ��������Ϣ   Number(2);
  n_�Ƿ��������     Number(2);
  n_�Ƿ��������ҩ�� Number(2);
  n_�Ƿ����������Ϣ Number(2);
  n_����ҽ������     Number(2);

  l_����ids  t_Strlist;
  c_����ids  Clob;
  P          Number;
  v_ҽ������ ҽ�����˵���.����%Type;

  Cursor c_���˻�����Ϣ Is
    Select ����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
           b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
           a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����,
           a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����,
           a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������,
           d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
    From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
    Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And Rownum < 1;
  r_���� c_���˻�����Ϣ%RowType;

  Type Ty_������Ϣ Is Ref Cursor;
  c_������Ϣ Ty_������Ϣ; --��̬�α����

  v_Json Varchar2(32767);
  v_Temp Varchar2(32767);

  n_Firstitem    Number;
  n_Firstsubitem Number;
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ѯ���� := j_Json.Get_Number('query_type');

  n_�Ƿ��������Ϣ   := j_Json.Get_Number('query_card');
  n_�Ƿ��������     := j_Json.Get_Number('query_family');
  n_�Ƿ��������ҩ�� := j_Json.Get_Number('query_drug');
  n_�Ƿ����������Ϣ := j_Json.Get_Number('query_immune');
  n_����ҽ������     := j_Json.Get_Number('query_insurance_pwd');

  o_Json := j_Json.Get_Pljson('query_cons_list');
  If Nvl(n_����id, 0) = 0 And o_Json Is Null Then
    v_Err_Msg := 'δ������Ч�Ĳ�ѯ���������ܻ�ȡ������Ϣ!';
    Json_Out  := '{"output":{"code":0,"message":"' || Zljsonstr(v_Err_Msg) || '"}}';
  
    Return;
  End If;

  If o_Json Is Not Null Then
    Begin
      c_����ids := o_Json.Get_Clob('pati_ids');
    Exception
      When Others Then
        c_����ids := Null;
    End;
    If Not c_����ids Is Null Then
      l_����ids := t_Strlist();
      c_����ids := c_����ids || ',';
      Loop
        P := Instr(c_����ids, ',');
        Exit When(Nvl(P, 0) = 0);
      
        l_����ids.Extend;
        l_����ids(l_����ids.Count) := (Substr(c_����ids, 1, P - 1));
        c_����ids := Substr(c_����ids, P + 1);
      End Loop;
    End If;
    v_����           := o_Json.Get_String('pati_name');
    n_�����         := To_Number(o_Json.Get_String('outpatient_num'));
    n_סԺ��         := To_Number(o_Json.Get_String('inpatient_num'));
    v_���֤��       := o_Json.Get_String('pati_idcard');
    v_��ϵ�����֤�� := o_Json.Get_String('contacts_idcard');
    n_�����id       := o_Json.Get_Number('cardtype_id');
    v_ҽ�ƿ�����     := o_Json.Get_String('medc_card_name');
    v_����           := o_Json.Get_String('card_no');
    v_��ά��         := o_Json.Get_String('qrcode');
    n_��ѯסԺ״̬   := o_Json.Get_Number('qrspt_statu');
    v_�ֻ���         := o_Json.Get_String('phone_number');
    v_Ic����         := o_Json.Get_String('iccard_no');
    v_���￨��       := o_Json.Get_String('visit_card');
    v_ҽ����         := o_Json.Get_String('insurance_num');
    v_����           := o_Json.Get_String('pati_bed');
    n_����id         := o_Json.Get_Number('dept_id');
    n_��������       := o_Json.Get_Number('search_days');
  End If;

  If Nvl(n_����id, 0) <> 0 Then
    --������IDΪ��Ҫ��ѯ�������в�ѯ
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
             b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
             a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��,
             a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���,
             c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.����id = n_����id And
            a.ͣ��ʱ�� Is Null;
  
  Elsif n_����� <> 0 Then
    --������Ų�ѯ
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
             b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
             a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��,
             a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���,
             c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.����� = n_����� And
            Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    --0-������;1-��Ժ ;2-���Ｐ��Ժ
  Elsif n_סԺ�� <> 0 Then
    --������Ų�ѯ
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
             b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
             a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��,
             a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���,
             c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.סԺ�� = n_סԺ�� And
            Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    --0-������;1-��Ժ ;2-���Ｐ��Ժ
  Elsif v_Ic���� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
             b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
             a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��,
             a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���,
             c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.Ic���� = v_Ic���� And
            Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    --0-������;1-��Ժ ;2-���Ｐ��Ժ
  Elsif v_���֤�� Is Not Null Then
    Select Max(����id)
    Into n_����id
    From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���� B
    Where Nvl(a.״̬, 0) = 0 And a.�����id = b.Id And b.���� = '�������֤' And b.�Ƿ����� = 1 And a.���� = v_���֤�� And
          Nvl(a.��ֹʹ��ʱ��, Sysdate + 1) > Sysdate;
    If Nvl(n_����id, 0) <> 0 Then
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�,
               a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰,
               a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�,
               a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��,
               a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��,
               a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.���֤�� = v_���֤�� And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    Else
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�,
               a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰,
               a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�,
               a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��,
               a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��,
               a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.���֤�� = v_���֤�� And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    End If;
    --0-������;1-��Ժ ;2-���Ｐ��Ժ
  Elsif v_ҽ���� Is Not Null Then
    If Instr(v_ҽ����, '%') > 0 Then
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�,
               a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰,
               a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�,
               a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��,
               a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��,
               a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.ҽ���� Like v_ҽ���� And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    Else
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�,
               a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰,
               a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�,
               a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��,
               a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��,
               a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.ҽ���� = v_ҽ���� And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    End If;
  Elsif v_���� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
             b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
             a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��,
             a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���,
             c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And
            a.��ǰ���� Like '%' || v_���� || '%' And Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And
            a.ͣ��ʱ�� Is Null;
    --0-������;1-��Ժ ;2-���Ｐ��Ժ
  Elsif l_����ids Is Not Null Then
    Open c_������Ϣ For
      Select /*+cardinality(Q,10)*/
       a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
       b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ,
       a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id,
       a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����,
       a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������,
       a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F, Table(l_����ids) Q
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And
            a.����id = q.Column_Value And Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And
            a.ͣ��ʱ�� Is Null;
  Elsif v_�ֻ��� Is Not Null Then
    Open c_������Ϣ For
      Select /*+cardinality(c,10)*/
       a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
       b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ,
       a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id,
       a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����,
       a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������,
       a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.�ֻ��� = v_�ֻ��� And
            Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
  Elsif v_��ϵ�����֤�� Is Not Null Then
    Open c_������Ϣ For
      Select /*+cardinality(c,10)*/
       a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
       b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ,
       a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id,
       a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����,
       a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������,
       a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And
            a.��ϵ�����֤�� = v_��ϵ�����֤�� And Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And
            a.ͣ��ʱ�� Is Null;
  Elsif v_���￨�� Is Not Null Then
    Open c_������Ϣ For
      Select /*+cardinality(c,10)*/
       a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
       b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ,
       a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id,
       a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����,
       a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������,
       a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.���￨�� = v_���￨�� And
            Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
  
  Elsif v_���� Is Not Null Then
    --����������
    If Instr(v_����, '%') > 0 Then
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�,
               a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰,
               a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�,
               a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��,
               a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��,
               a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.���� Like v_���� And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null And
              Decode(Nvl(n_��������, '0'), '0', '0', Nvl(a.����ʱ��, a.�Ǽ�ʱ��) || '') >=
              Decode(Nvl(n_��������, '0'), '0', '0', Trunc(Sysdate - Nvl(n_��������, 0)) || '');
    
    Else
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�,
               a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰,
               a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�,
               a.����ʱ��, a.����״̬, a.��������, a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��,
               a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��,
               a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.���� = v_���� And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    End If;
  Elsif v_���� Is Not Null Or v_��ά�� Is Not Null Then
    If Nvl(n_�����id, 0) = 0 Then
      Select Max(ID) Into n_�����id From ҽ�ƿ���� Where ���� = v_ҽ�ƿ�����;
    End If;
    If Nvl(n_�����id, 0) <> 0 And v_���� Is Not Null Then
      Open c_������Ϣ For
        Select /*+cardinality(c,10)*/
         a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
         b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
         a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����,
         a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����,
         a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������,
         d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F,
             (Select Distinct ����id
               From (Select j.����id,
                             Case
                                When Nvl(j.״̬, 0) = 1 And
                                     (Nvl(m.��Ч����, 0) = 0 Or Nvl(j.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.��Ч����, 0) > Sysdate) Then
                                 1
                                Else
                                 Nvl(j.״̬, 0)
                              End As ״̬
                      From ����ҽ�ƿ���Ϣ J, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
                      Where j.��ʧ��ʽ = m.����(+) And j.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And j.�����id = n_�����id And
                            j.���� = v_���� And Sysdate < Nvl(j.��ֹʹ��ʱ��, Sysdate + 1))
               Where ״̬ = 0) Q
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.����id = q.����id And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
      --0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
    Elsif Nvl(n_�����id, 0) <> 0 And v_��ά�� Is Not Null Then
      Open c_������Ϣ For
        Select /*+cardinality(c,10)*/
         a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
         b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
         a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����,
         a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����,
         a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������,
         d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F,
             (Select Distinct ����id
               From (Select j.����id,
                             Case
                                When Nvl(j.״̬, 0) = 1 And
                                     (Nvl(m.��Ч����, 0) = 0 Or Nvl(j.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.��Ч����, 0) > Sysdate) Then
                                 1
                                Else
                                 Nvl(j.״̬, 0)
                              End As ״̬
                      From ����ҽ�ƿ���Ϣ J, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
                      Where j.��ʧ��ʽ = m.����(+) And j.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And j.�����id = n_�����id And
                            j.��ά�� = v_��ά�� And Sysdate < Nvl(j.��ֹʹ��ʱ��, Sysdate + 1))
               Where ״̬ = 0) Q
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.����id = q.����id And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
      --0-������Ч��;1-�ѹ�ʧ; 2-����ͣ��
    Elsif v_���� Is Not Null Then
      Open c_������Ϣ For
        Select /*+cardinality(c,10)*/
         a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
         b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
         a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����,
         a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����,
         a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������,
         d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F,
             (Select Distinct ����id
               From (Select j.����id,
                             Case
                                When Nvl(j.״̬, 0) = 1 And
                                     (Nvl(m.��Ч����, 0) = 0 Or Nvl(j.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.��Ч����, 0) > Sysdate) Then
                                 1
                                Else
                                 Nvl(j.״̬, 0)
                              End As ״̬
                      From ����ҽ�ƿ���Ϣ J, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
                      Where j.��ʧ��ʽ = m.����(+) And j.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And j.���� = v_���� And
                            Sysdate < Nvl(j.��ֹʹ��ʱ��, Sysdate + 1))
               Where ״̬ = 0) Q
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.����id = q.����id And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    
    Else
      Open c_������Ϣ For
        Select /*+cardinality(c,10)*/
         a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
         b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
         a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������, a.סԺ����,
         a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��, a.�໤��, a.��ѯ����,
         a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���, c.���� As ��ǰ��������,
         d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F,
             
             (Select Distinct ����id
               From (Select j.����id,
                             Case
                               When Nvl(j.״̬, 0) = 1 And
                                    (Nvl(m.��Ч����, 0) = 0 Or
                                     Nvl(j.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.��Ч����, 0) > Sysdate) Then
                                1
                               Else
                                Nvl(j.״̬, 0)
                             End As ״̬
                      From ����ҽ�ƿ���Ϣ J, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
                      Where j.��ʧ��ʽ = m.����(+) And j.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And j.��ά�� = v_��ά�� And
                            Sysdate < Nvl(j.��ֹʹ��ʱ��, Sysdate + 1))
               Where ״̬ = 0) Q
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And a.����id = q.����id And
              Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.ͣ��ʱ�� Is Null;
    End If;
  Elsif Nvl(n_����id, 0) <> 0 Then
    --����ǰ���Ҳ���
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����,
             b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������, a.��ϵ�˹�ϵ,
             a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.��ͬ��λid, a.������λ As ������λ����, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.����ʱ��, a.����״̬, a.��������,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, a.��Ժʱ��, a.��Ժʱ��, a.Ic����, a.������, a.����, a.�Ǽ�ʱ��, a.ͣ��ʱ��, a.��ǰ����, a.ҽ����, a.����֤��,
             a.�໤��, a.��ѯ����, a.��Ժ, a.����, a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, a.Email, a.Qq, a.��ϵ�����֤��, a.��������, a.��ҳid, a.�ֻ���,
             c.���� As ��ǰ��������, d.���� As ��ǰ��������, a.��λ��ַ As ������λ��ַ, f.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ������� F
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.���� = f.���(+) And
            Decode(Nvl(n_��ѯסԺ״̬, 0), 2, 2, Nvl(a.��Ժ, 0)) = Nvl(n_��ѯסԺ״̬, 0) And a.��ǰ����id = n_����id And a.ͣ��ʱ�� Is Null;
  Else
    v_Err_Msg := 'δ������Ч�Ĳ�ѯ���������ܻ�ȡ������Ϣ!';
    Json_Out  := '{"output":{"code":0,"message":"' || Zljsonstr(v_Err_Msg) || '"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[';

  v_Json      := '';
  n_Firstitem := 1;
  Loop
    Fetch c_������Ϣ
      Into r_����;
    Exit When c_������Ϣ %NotFound;
    If Nvl(n_Firstitem, 0) = 0 Then
      v_Json := v_Json || ',';
    Else
      n_Firstitem := 0;
    End If;
  
    v_Json := v_Json || '{';
    --1.ȡ������Ϣ
    v_Json := v_Json || '"pati_id":' || Nvl(r_����.����id, 0);
    v_Json := v_Json || ',"pati_pageid":' || Nvl(r_����.��ҳid, 0);
    v_Json := v_Json || ',"pati_name":"' || Zljsonstr(r_����.����) || '"';
    v_Json := v_Json || ',"pati_sex":"' || Zljsonstr(r_����.�Ա�) || '"';
    v_Json := v_Json || ',"pati_age":"' || Zljsonstr(r_����.����) || '"';
  
    v_Json := v_Json || ',"pati_birthdate":"' || Zljsonstr(To_Char(r_����.��������, 'yyyy-mm-dd hh24:mi:ss')) || '"';
    v_Json := v_Json || ',"fee_category":"' || Zljsonstr(r_����.�ѱ�) || '"';
    If Nvl(r_����.�����, 0) = 0 Then
      v_Json := v_Json || ',"outpatient_num":null';
    Else
      v_Json := v_Json || ',"outpatient_num":"' || r_����.����� || '"';
    End If;
    If Nvl(r_����.סԺ��, 0) = 0 Then
      v_Json := v_Json || ',"inpatient_num":null';
    Else
      v_Json := v_Json || ',"inpatient_num":"' || r_����.סԺ�� || '"';
    End If;
    v_Json := v_Json || ',"pati_nation":"' || Zljsonstr(r_����.����) || '"';
    v_Json := v_Json || ',"mdlpay_mode_name":"' || Zljsonstr(r_����.ҽ�Ƹ��ʽ����) || '"';
    v_Json := v_Json || ',"mdlpay_mode_code":"' || Zljsonstr(r_����.ҽ�Ƹ��ʽ����) || '"';
    v_Json := v_Json || ',"insurance_num":"' || Zljsonstr(r_����.ҽ����) || '"';
    v_Json := v_Json || ',"pati_idcard":"' || Zljsonstr(r_����.���֤��) || '"';
    v_Json := v_Json || ',"vcard_no":"' || Zljsonstr(r_����.���￨��) || '"';
  
    v_Json := v_Json || ',"iccard_no":"' || Zljsonstr(r_����.Ic����) || '"';
    v_Json := v_Json || ',"inp_times":' || Nvl(r_����.סԺ����, 0);
    v_Json := v_Json || ',"pati_education":"' || Zljsonstr(r_����.ѧ��) || '"';
    v_Json := v_Json || ',"ocpt_name":"' || Zljsonstr(r_����.ְҵ) || '"';
    v_Json := v_Json || ',"pati_marital_cstatus":"' || Zljsonstr(r_����.����״��) || '"';
  
    v_Json := v_Json || ',"phone_number":"' || Zljsonstr(r_����.�ֻ���) || '"';
    v_Json := v_Json || ',"pati_bed":"' || Zljsonstr(r_����.��ǰ����) || '"';
    v_Json := v_Json || ',"pati_birthplace":"' || Zljsonstr(r_����.�����ص�) || '"';
    v_Json := v_Json || ',"pat_home_addr":"' || Zljsonstr(r_����.��ͥ��ַ) || '"';
    v_Json := v_Json || ',"pat_home_phno":"' || Zljsonstr(r_����.��ͥ�绰) || '"';
  
    v_Json := v_Json || ',"insurance_type":' || Nvl(r_����.����, 0);
    v_Json := v_Json || ',"insurance_name":"' || Zljsonstr(r_����.��������) || '"';
    v_Json := v_Json || ',"is_inhspt":' || Nvl(r_����.��Ժ, 0);
    v_Json := v_Json || ',"pati_type":"' || Zljsonstr(r_����.��������) || '"';
  
    --2.��ѯ������Ϣ+��ϵ����Ϣ
    If Nvl(n_��ѯ����, 0) >= 1 Then
      --��ѯ����:�磺0-����;1-����+��ϵ��;2-����
      v_Json := v_Json || ',"contacts_name":"' || Zljsonstr(r_����.��ϵ������) || '"';
      v_Json := v_Json || ',"contacts_relation":"' || Zljsonstr(r_����.��ϵ�˹�ϵ) || '"';
      v_Json := v_Json || ',"contacts_idcard":"' || Zljsonstr(r_����.��ϵ�����֤��) || '"';
      v_Json := v_Json || ',"contacts_addr":"' || Zljsonstr(r_����.��ϵ�˵�ַ) || '"';
      v_Json := v_Json || ',"contacts_phno":"' || Zljsonstr(r_����.��ϵ�˵绰) || '"';
    End If;
  
    --3.��ѯ��������:�磺0-����;1-����+��ϵ��;2-����
    If Nvl(n_��ѯ����, 0) > 1 Then
      v_Json := v_Json || ',"pati_wardarea_id":' || Nvl(r_����.��ǰ����id, 0);
      v_Json := v_Json || ',"pati_wardarea_name":"' || Zljsonstr(r_����.��ǰ��������) || '"';
      v_Json := v_Json || ',"pati_dept_id":' || Nvl(r_����.��ǰ����id, 0);
      v_Json := v_Json || ',"pati_dept_name":"' || Zljsonstr(r_����.��ǰ��������) || '"';
      v_Json := v_Json || ',"health_num":"' || Zljsonstr(r_����.������) || '"';
    
      v_Json := v_Json || ',"pati_identity":"' || Zljsonstr(r_����.���) || '"';
      v_Json := v_Json || ',"ntvplc_name":"' || Zljsonstr(r_����.����) || '"';
      v_Json := v_Json || ',"country_name":"' || Zljsonstr(r_����.����) || '"';
      v_Json := v_Json || ',"pat_home_postcode":"' || Zljsonstr(r_����.��ͥ��ַ�ʱ�) || '"';
      v_Json := v_Json || ',"pati_area":"' || Zljsonstr(r_����.����) || '"';
    
      v_Json := v_Json || ',"pat_hous_addr":"' || Zljsonstr(r_����.���ڵ�ַ) || '"';
      v_Json := v_Json || ',"pat_hous_postcode":"' || Zljsonstr(r_����.���ڵ�ַ�ʱ�) || '"';
      v_Json := v_Json || ',"emp_addr":"' || Zljsonstr(r_����.������λ��ַ) || '"';
      v_Json := v_Json || ',"emp_name":"' || Zljsonstr(r_����.������λ����) || '"';
      v_Json := v_Json || ',"emp_phno":"' || Zljsonstr(r_����.��λ�绰) || '"';
    
      v_Json := v_Json || ',"emp_postcode":"' || Zljsonstr(r_����.��λ�ʱ�) || '"';
      v_Json := v_Json || ',"emp_bank_name":"' || Zljsonstr(r_����.��λ������) || '"';
      v_Json := v_Json || ',"emp_bank_accnum":"' || Zljsonstr(r_����.��λ�ʺ�) || '"';
      v_Json := v_Json || ',"ctt_unit_id":' || Nvl(r_����.��ͬ��λid, 0);
      v_Json := v_Json || ',"adta_time":"' || Zljsonstr(To_Char(r_����.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
    
      v_Json := v_Json || ',"adtd_time":"' || Zljsonstr(To_Char(r_����.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
      v_Json := v_Json || ',"pat_grdn_name":"' || Zljsonstr(r_����.�໤��) || '"';
      v_Json := v_Json || ',"cert_no_other":"' || Zljsonstr(r_����.����֤��) || '"';
      v_Json := v_Json || ',"pati_email":"' || Zljsonstr(r_����.Email) || '"';
      v_Json := v_Json || ',"pati_qq":"' || Zljsonstr(r_����.Qq) || '"';
      v_Json := v_Json || ',"card_captcha":"' || Zljsonstr(r_����.����֤��) || '"';
    
      n_��ɫ := Null;
      If r_����.�������� Is Not Null Then
        Select Max(��ɫ) Into n_��ɫ From �������� Where ���� = r_����.��������;
      End If;
      v_Json := v_Json || ',"pati_show_color":' || Nvl(n_��ɫ, 0);
      v_Json := v_Json || ',"visit_room":"' || Zljsonstr(r_����.��������) || '"';
      v_Json := v_Json || ',"visit_statu":' || Nvl(r_����.����״̬, 0);
      v_Json := v_Json || ',"visit_time":"' || Zljsonstr(To_Char(r_����.����ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
      v_Json := v_Json || ',"create_time":"' || Zljsonstr(To_Char(r_����.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss')) || '"';
    
      If Nvl(r_����.����, 0) <> 0 And Nvl(n_����ҽ������, 0) = 1 Then
        Select Max(d.����)
        Into v_ҽ������
        From ҽ�����˵��� D, ҽ�����˹����� E
        Where e.����id = r_����.����id And e.���� = r_����.���� And e.ҽ���� = r_����.ҽ���� And e.��־ = 1 And e.ҽ���� = d.ҽ����(+) And
              e.���� = d.����(+) And e.���� = d.����(+);
      Else
        v_ҽ������ := '';
      End If;
      v_Json := v_Json || ',"insurance_pwd":"' || Zljsonstr(v_ҽ������) || '"';
    End If;
  
    --��ȡ������ϵ
    If Nvl(n_�Ƿ��������, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_���� In (Select ����id, ��ϵ From ���˼��� Where ����id = r_����.����id And Nvl(����ʱ��, Sysdate) <= Sysdate) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"family_id":' || c_����.����id;
        v_Temp := v_Temp || ',"family_relation":"' || Zljsonstr(c_����.��ϵ) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"family_list":[' || v_Temp || ']';
    End If;
  
    --��ȡ����ҩ��
    If Nvl(n_�Ƿ��������ҩ��, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_����ҩ�� In (Select ����ҩ��id, ����ҩ��, ������Ӧ From ���˹���ҩ�� Where ����id = r_����.����id) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"pat_algc_cadn_id":' || Zljsonstr(c_����ҩ��.����ҩ��id, 1);
        v_Temp := v_Temp || ',"pat_algc_cadn":"' || Zljsonstr(c_����ҩ��.����ҩ��) || '"';
        v_Temp := v_Temp || ',"allergy_info":"' || Zljsonstr(c_����ҩ��.������Ӧ) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"drug_list":[' || v_Temp || ']';
    End If;
  
    -- ��ȡ����������Ϣ
    If Nvl(n_�Ƿ����������Ϣ, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_���߼�¼ In (Select To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������
                     From �������߼�¼
                     Where ����id = r_����.����id) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"vaccinate_time":"' || Zljsonstr(c_���߼�¼.����ʱ��) || '"';
        v_Temp := v_Temp || ',"vaccinate_name":"' || Zljsonstr(c_���߼�¼.��������) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"immune_list":[' || v_Temp || ']';
    End If;
  
    --��ȡ����Ϣ
    If Nvl(n_�Ƿ��������Ϣ, 0) = 1 Then
      v_Temp         := '';
      n_Firstsubitem := 1;
      For c_ҽ�ƿ� In (Select Distinct �����id, ����, ����
                    From (Select j.�����id, j.����, j.����,
                                  Case
                                    When Nvl(j.״̬, 0) = 1 And
                                         (Nvl(m.��Ч����, 0) = 0 Or
                                          Nvl(j.��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(m.��Ч����, 0) > Sysdate) Then
                                     1
                                    Else
                                     Nvl(j.״̬, 0)
                                  End As ״̬
                           From ����ҽ�ƿ���Ϣ J, ҽ�ƿ���ʧ��ʽ M, ҽ�ƿ���� Q
                           Where j.��ʧ��ʽ = m.����(+) And j.�����id = q.Id And Nvl(q.�Ƿ�����, 0) = 1 And j.����id = r_����.����id And
                                 Decode(Nvl(n_�����id, 0), 0, j.�����id) = Nvl(n_�����id, 0) And
                                 Sysdate < Nvl(j.��ֹʹ��ʱ��, Sysdate + 1))
                    Where ״̬ = 0) Loop
        If Nvl(n_Firstsubitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstsubitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"cardtype_id":' || c_ҽ�ƿ�.�����id;
        v_Temp := v_Temp || ',"card_no":' || Zljsonstr(c_ҽ�ƿ�.����) || '"';
        v_Temp := v_Temp || ',"card_pwd":' || Zljsonstr(c_ҽ�ƿ�.����) || '"';
        v_Temp := v_Temp || '}';
      End Loop;
      v_Json := v_Json || ',"card_list":[' || v_Temp || ']';
    End If;
    v_Json := v_Json || '}';
  
    If Length(v_Json) > 20000 Then
      Json_Out := Json_Out || v_Json;
      v_Json   := '';
    End If;
  End Loop;
  Close c_������Ϣ;
  Json_Out := Json_Out || v_Json || ']}}';

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiinfsbyrange
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      query_type          N 1 0����ѯ������Ϣ��1����ѯ������Ϣ+��չ��Ϣ
  --      pati_ids            C   ����IDs:����ö���
  --      pati_name           C   ����:���Դ�%�ֺű������ƥ��
  --      pati_sex            C   �Ա�
  --      pati_age            C   ����
  --      birthdate_start     C   ��ʼ��������
  --      birthdate_end       C   ��ֹ��������
  --      outpatient_num      C   �����
  --      pati_idcard         C   ���֤��
  --      fee_category        C   �ѱ�
  --      pati_area           C   ����
  --      insurance_num       C   ҽ����
  --      vcard_no            C   ���￨��
  --      iccard_no           C   Ic����
  --      wardarea_ids        C   ����ids������ö���
  --      qurey_max           N   ��ѯ������¼����Ϊ0��NULLʱ��ʾ������
  --      qrspt_statu         N   ��ѯסԺ״̬:0-������;1-��Ժ ;2-���Ｐ��Ժ
  --      visit_star_time     C   ���￪ʼʱ��:yyyy-mm-dd hh24:mi:ss
  --      visit_end_time      C   �������ʱ��:yyyy-mm-dd hh24:mi:ss
  --      create_start_time   C   ��ʼ�Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --      create_end_time     C   ��ֹ�Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --      occasion            N   ����:����Zl_Custom_Patiids_Get(�������֤���ز���id)����ʱ�贫��
  --      only_ctorg_pati     N   ֻ��ѯ��Լ��λ�Ĳ���
  --      ctt_unit_id         N   ��ͬ��λid,ֻ��ѯ��Լ��λ�Ĳ���ʱ��Ч
  --      default_cardtype_id N   ȱʡ�����id
  --      dept_ids            C   ����ids:����ö��ŷָ�
  --      mdlpay_mode_name    C   ҽ�Ƹ��ʽ
  --      phone_number        C   �ֻ���
  --      is_stop             N   �Ƿ���ʾͣ��
  --      pati_similar        C   ��������
  --        pati_name         C 1 ����
  --        pati_sex          C 1 �Ա�
  --        country_name      C 1 ����
  --        pati_nation       C 1 ����
  --        pati_birthdate    C 1 �������ڣ�yyyy-mm-dd hh24:mi:ss
  --        pati_idcard       C 1 ���֤��
  --����      json
  --output
  -- code                     N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  -- message                  C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -- pati_list[]                  ������Ϣ�б�
  --   pati_id                N 1 ����id
  --   pati_pageid            N 1 ��ҳid��������Ϣ.��ҳID
  --   pati_name              C 1 ����
  --   pati_sex               C 1 �Ա�
  --   pati_age               C 1 ����
  --   pati_birthdate         C 1 �������ڣ�yyyy-mm-dd hh24:mi:ss
  --   pati_birthplace        C 1 �����ص�
  --   fee_category           C 1 �ѱ�
  --   outpatient_num         C 1 �����
  --   inpatient_num          C 1 סԺ��
  --   inp_times              N 1 סԺ����
  --   pati_nation            C 1 ����
  --   pati_idcard            C 1 ���֤��
  --   vcard_no               C 1 ���￨��
  --   phone_number           C 1 �ֻ���
  --   pat_home_phno          C 1 ��ͥ�绰
  --   pati_education         C 1 ѧ��
  --   ocpt_name              C 1 ְҵ
  --   pati_identity          C 1 ���
  --   country_name           C 1 ����
  --   pat_home_addr          C 1 ��ͥ��ַ
  --   pati_area              C 1 ����
  --   emp_name               C 1 ������λ����
  --   pati_bed               C 1 ��ǰ����
  --   is_inhspt              N 1 �Ƿ���Ժ��1-��Ժ��0-����Ժ
  --   pati_type              C 1 ��������(��ͨ��ҽ��������)
  --   insurance_type         C 1 ����
  --   insurance_type_name    C 1 ��������
  --   pati_wardarea_id       N 1 ��ǰ����id
  --   pati_wardarea_name     C 1 ��ǰ��������
  --   pati_dept_id           N 1 ��ǰ����id
  --   pati_dept_name         C 1 ��ǰ��������
  --   adta_time              C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
  --   adtd_time              C 1 ��Ժʱ��:yyyy-mm-dd hh24:mi:ss
  --   create_time            C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --   medc_card_no           C   ҽ�ƿ��ţ�����νڵ�default_cardtype_id��Ϊ��ʱ���ŷ���
  --   visit_time             C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --   ctt_unit_id            N   ��ͬ��λid
  --   mdlpay_mode_name       C 1 ҽ�Ƹ��ʽ����
  --   mdlpay_mode_code       C 1 ҽ�Ƹ��ʽ����
  --   stop_time              C  ͣ��ʱ��
  --   insurance_num          C  ҽ����
  --   emp_addr               C  ��λ��ַ
  --   contacts               C   ��ϵ����Ϣ�ڵ�
  --     name                 C 1 ��ϵ������
  --     phone                C 1 ��ϵ�˵绰
  ---------------------------------------------------------------------------
  v_Err_Msg      Varchar2(500);
  j_Jsonin       Pljson;
  j_Json         Pljson;
  j_Json_Similar Pljson;
  c_����ids      Clob;
  v_List         Varchar2(32767);
  v_Listtmp      Varchar2(32767);
  n_��ѯ����     Number(1);
  v_����         Varchar2(200);
  v_�Ա�         Varchar2(50);
  v_����ids      Varchar2(3000);
  d_��ʼ�������� Date;
  d_��ֹ�������� Date;
  d_��������     Date;
  n_�����       Number(18);
  v_���֤��     Varchar2(50);
  v_�ѱ�         Varchar2(50);

  v_����         Varchar2(100);
  v_���￨��     Varchar2(200);
  v_Ic����       Varchar2(200);
  d_���￪ʼʱ�� Date;
  d_�������ʱ�� Date;
  n_Like         Number(2);
  n_Max          Number(10);
  d_��ʼ�Ǽ�ʱ�� Date;
  d_�����Ǽ�ʱ�� Date;
  v_����ids      Varchar2(32680);
  n_��ѯסԺ״̬ Number(2);
  v_ҽ����       Varchar2(200);
  n_����         Number(20);
  n_����Լ��λ   Number(1);
  n_ȱʡ�����id ����ҽ�ƿ���Ϣ.�����id%Type;
  l_����id       t_Strlist := t_Strlist();
  v_����ids      Varchar2(32680);
  v_ҽ�Ƹ��ʽ Varchar2(100);
  v_�ֻ���       Varchar2(100);
  v_����         ������Ϣ.����%Type;
  n_��ͬ��λid   ������Ϣ.��ͬ��λid%Type;
  n_�Ƿ�ͣ��     Number;
  v_����         ������Ϣ.����%Type;
  v_����         ������Ϣ.����%Type;

  Cursor c_���˻�����Ϣ Is
    Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
           a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
           a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
           To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
           a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
           To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
           To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
    From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
    Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.ͣ��ʱ�� Is Null And
          a.����id = f.����id(+) And a.���� = x.���(+) And Rownum < 1;
  r_���� c_���˻�����Ϣ%RowType;

  Type Ty_������Ϣ Is Ref Cursor;
  c_������Ϣ Ty_������Ϣ; --��̬�α����

  --��װʧ��ʱ���ص�����
  Function Get_Err_Message(Message_In Varchar2) Return Varchar2 Is
    j_Out Varchar2(32767);
  Begin
    j_Out := '{"output":{"code":0,"message":"' || Zljsonstr(Message_In) || '"}}';
    Return j_Out;
  End Get_Err_Message;

Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_��ѯ���� := Nvl(j_Json.Get_Number('query_type'), 0);
  Begin
    c_����ids := j_Json.Get_Clob('pati_ids');
  Exception
    When Others Then
      c_����ids := Null;
  End;

  d_��ʼ�������� := To_Date(j_Json.Get_String('birthdate_start'), 'YYYY-MM-DD hh24:mi:ss');
  d_��ֹ�������� := To_Date(j_Json.Get_String('birthdate_end'), 'YYYY-MM-DD hh24:mi:ss');
  n_�����       := To_Number(j_Json.Get_String('outpatient_num'));
  v_���֤��     := j_Json.Get_String('pati_idcard');
  v_�ѱ�         := j_Json.Get_String('fee_category');
  v_�Ա�         := j_Json.Get_String('pati_sex');
  v_����         := j_Json.Get_String('pati_area');
  v_���￨��     := j_Json.Get_String('vcard_no');
  v_Ic����       := j_Json.Get_String('iccard_no');
  d_���￪ʼʱ�� := To_Date(j_Json.Get_String('visit_start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_�������ʱ�� := To_Date(j_Json.Get_String('visit_end_time'), 'yyyy-mm-dd hh24:mi:ss');

  d_��ʼ�Ǽ�ʱ�� := To_Date(j_Json.Get_String('create_start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_�����Ǽ�ʱ�� := To_Date(j_Json.Get_String('create_end_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_����         := j_Json.Get_String('pati_name');
  v_����ids      := j_Json.Get_String('wardarea_ids');
  n_��ѯסԺ״̬ := Nvl(j_Json.Get_Number('qrspt_statu'), 0);
  n_Max          := j_Json.Get_Number('qurey_Max');
  v_ҽ����       := j_Json.Get_String('insurance_num');
  n_����         := j_Json.Get_Number('occasion');
  n_ȱʡ�����id := j_Json.Get_Number('default_cardtype_id');
  n_����Լ��λ   := Nvl(j_Json.Get_Number('only_ctorg_pati'), 0);
  n_��ͬ��λid   := j_Json.Get_Number('ctt_unit_id');
  v_����ids      := j_Json.Get_Number('dept_ids');
  v_ҽ�Ƹ��ʽ := j_Json.Get_String('mdlpay_mode_name');
  v_�ֻ���       := j_Json.Get_String('phone_number');
  v_����         := j_Json.Get_String('pati_age');
  n_�Ƿ�ͣ��     := j_Json.Get_Number('is_stop');
  ---���Ʋ��˲�ѯ����
  --      pati_similar        C   ��������
  --        pati_name         C 1 ����
  --        pati_sex          C 1 �Ա�
  --        country_name      C 1 ����
  --        pati_nation       C 1 ����
  --        pati_birthdate    C 1 �������ڣ�yyyy-mm-dd hh24:mi:ss
  --        pati_idcard       C 1 ���֤��
  j_Json_Similar := j_Json.Get_Pljson('pati_similar');
  If Not j_Json_Similar Is Null Then
    v_����     := j_Json_Similar.Get_String('pati_name');
    v_�Ա�     := j_Json_Similar.Get_String('pati_sex');
    v_����     := j_Json_Similar.Get_String('country_name');
    v_����     := j_Json_Similar.Get_String('pati_nation');
    d_�������� := To_Date(j_Json_Similar.Get_String('birthdate_start'), 'YYYY-MM-DD hh24:mi:ss');
    v_���֤�� := j_Json_Similar.Get_String('pati_idcard');
  End If;

  If v_����ids Is Not Null Then
    v_����ids := ',' || v_����ids || ',';
  End If;

  n_Like := 0;
  If Instr(Nvl(v_����, '-'), '%') > 0 Then
    n_Like := 1;
  End If;

  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      l_����id.Extend;
      l_����id(l_����id.Count) := c_����ids;
      c_����ids := Null;
    Else
      l_����id.Extend;
      l_����id(l_����id.Count) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
    End If;
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[';
  If l_����id.Count <> 0 Then
    --0-������;1-��Ժ ;2-���Ｐ��Ժ
    For I In 1 .. l_����id.Count Loop
      For r_���� In (With c_���� As
                      (Select Column_Value As ����id From Table(f_Num2list(l_����id(I))))
                     Select /*+cardinality(B,10)*/
                      a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����,
                      a.סԺ��, a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��,
                      a.��ͥ��ַ, a.������λ, a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
                      To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����,
                      To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������,
                      d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��,
                      a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid, To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������,
                      a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
                     From ������Ϣ A, ҽ�Ƹ��ʽ B, c_���� B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
                     Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And
                           a.���� = x.���(+) And ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And
                           (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And a.����id = f.����id(+) And f.�����id(+) = n_ȱʡ�����id And
                           a.����id = b.����id And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
                           Nvl(d_��ֹ��������, Sysdate + 1 - 1 / 24 / 60 / 60) And
                           Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
                           Nvl(d_�����Ǽ�ʱ��, (Sysdate) + 1 - 1 / 24 / 60 / 60) And
                           Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
                           Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
                           (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
                           Decode(Nvl(n_�����, 0), 0, 0, a.�����) = Nvl(n_�����, 0) And
                           Decode(Nvl(v_���֤��, '-'), '-', '-', a.���֤��) = Nvl(v_���֤��, '-') And
                           Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
                           Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
                           Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
                           Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
                           Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
                           Decode(Nvl(v_���￨��, '-'), '-', '-', a.���￨��) = Nvl(v_���￨��, '-') And
                           Decode(Nvl(v_Ic����, '-'), '-', '-', a.Ic����) = Nvl(v_Ic����, '-') And
                           Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
                           Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
                           (v_����ids Is Null Or a.��ǰ����id Is Null Or
                           a.��ǰ����id In (Select ����id
                                         From �������Ҷ�Ӧ
                                         Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
                           (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or
                           Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
                           (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And
                           (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null))) Loop
      
        Zljsonputvalue(v_List, 'pati_id', r_����.����id, 1, 1);
        Zljsonputvalue(v_List, 'pati_pageid', r_����.��ҳid, 1);
        Zljsonputvalue(v_List, 'pati_name', r_����.����);
        Zljsonputvalue(v_List, 'pati_sex', r_����.�Ա�);
        Zljsonputvalue(v_List, 'pati_age', r_����.����);
        Zljsonputvalue(v_List, 'pati_birthdate', r_����.��������);
        Zljsonputvalue(v_List, 'pati_birthplace', r_����.�����ص�);
        Zljsonputvalue(v_List, 'fee_category', r_����.�ѱ�);
        Zljsonputvalue(v_List, 'outpatient_num', r_����.�����, 0);
        Zljsonputvalue(v_List, 'inpatient_num', r_����.סԺ��, 0);
        Zljsonputvalue(v_List, 'inp_times', r_����.סԺ����, 1);
      
        Zljsonputvalue(v_List, 'pati_nation', r_����.����);
        Zljsonputvalue(v_List, 'pati_idcard', r_����.���֤��);
        Zljsonputvalue(v_List, 'vcard_no', r_����.���￨��);
      
        Zljsonputvalue(v_List, 'pati_education', r_����.ѧ��);
        Zljsonputvalue(v_List, 'ocpt_name', r_����.ְҵ);
      
        Zljsonputvalue(v_List, 'pati_identity', r_����.���);
        Zljsonputvalue(v_List, 'country_name', r_����.����);
        Zljsonputvalue(v_List, 'pat_home_addr', r_����.��ͥ��ַ);
        Zljsonputvalue(v_List, 'pati_area', r_����.����);
        Zljsonputvalue(v_List, 'emp_name', r_����.������λ����);
        Zljsonputvalue(v_List, 'emp_addr', r_����.��λ��ַ);
      
        Zljsonputvalue(v_List, 'is_inhspt', r_����.��Ժ, 1);
        Zljsonputvalue(v_List, 'pati_bed', r_����.��ǰ����);
        Zljsonputvalue(v_List, 'pati_type', r_����.��������);
        Zljsonputvalue(v_List, 'insurance_type', r_����.����, 1);
        Zljsonputvalue(v_List, 'insurance_type_name', r_����.��������);
        Zljsonputvalue(v_List, 'pati_wardarea_id', r_����.��ǰ����id, 1);
        Zljsonputvalue(v_List, 'pati_wardarea_name', r_����.��ǰ��������);
        Zljsonputvalue(v_List, 'pati_dept_id', r_����.��ǰ����id, 1);
        Zljsonputvalue(v_List, 'pati_dept_name', r_����.��ǰ��������);
      
        Zljsonputvalue(v_List, 'adta_time', r_����.��Ժʱ��);
        Zljsonputvalue(v_List, 'adtd_time', r_����.��Ժʱ��);
        Zljsonputvalue(v_List, 'create_time', r_����.�Ǽ�ʱ��);
        Zljsonputvalue(v_List, 'phone_number', r_����.�ֻ���);
        Zljsonputvalue(v_List, 'pat_home_phno', r_����.��ͥ�绰);
        If n_��ѯ���� = 1 Then
          Zljsonputvalue(v_List, 'stop_time', r_����.ͣ��ʱ��);
        Else
          Zljsonputvalue(v_List, 'stop_time', r_����.ͣ��ʱ��, 0, 2);
        End If;
        If n_��ѯ���� = 1 Then
          Zljsonputvalue(v_List, 'medc_card_no', r_����.����);
          Zljsonputvalue(v_List, 'visit_time', r_����.����ʱ��);
          Zljsonputvalue(v_List, 'ctt_unit_id', r_����.��Լ��λid);
          Zljsonputvalue(v_List, 'mdlpay_mode_name', r_����.ҽ�Ƹ��ʽ����);
          Zljsonputvalue(v_List, 'mdlpay_mode_code', r_����.ҽ�Ƹ��ʽ����);
          Zljsonputvalue(v_List, 'insurance_num', r_����.ҽ����, 0);
          v_Listtmp := '"contacts":{"name":"' || Nvl(r_����.��ϵ������, '') || '","phone":"' || Nvl(r_����.��ϵ�˵绰, '') || '"}}';
          v_List    := v_List || ',' || v_Listtmp;
        End If;
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
    End Loop;
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
  
    Return;
  
  Elsif v_�ֻ��� Is Not Null Then
    --���ֻ��Ų�ѯ
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
            (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And f.�����id(+) = n_ȱʡ�����id And
            Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
            Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
            Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.����� = n_����� And
            Decode(Nvl(v_���֤��, '-'), '-', '-', a.���֤��) = Nvl(v_���֤��, '-') And
            Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
            Decode(Nvl(v_���￨��, '-'), '-', '-', a.���￨��) = Nvl(v_���￨��, '-') And
            Decode(Nvl(v_Ic����, '-'), '-', '-', a.Ic����) = Nvl(v_Ic����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And a.�ֻ��� = v_�ֻ��� And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0));
  Elsif n_����� <> 0 Then
    --������Ų�ѯ
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And
            a.����id = f.����id(+) And (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And f.�����id(+) = n_ȱʡ�����id And
            a.���� = x.���(+) And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
            Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
            Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.����� = n_����� And
            Decode(Nvl(v_���֤��, '-'), '-', '-', a.���֤��) = Nvl(v_���֤��, '-') And
            Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
            Decode(Nvl(v_���￨��, '-'), '-', '-', a.���￨��) = Nvl(v_���￨��, '-') And
            Decode(Nvl(v_Ic����, '-'), '-', '-', a.Ic����) = Nvl(v_Ic����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0));
  Elsif v_���֤�� Is Not Null Then
    If Nvl(n_����, 0) <> 0 Then
      v_����ids := Zl_Custom_Patiids_Get(Nvl(n_����, 0), v_���֤��);
    End If;
    If v_����ids Is Null Then
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
               a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
               a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
               To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
               To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And
              a.���� = x.���(+) And (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And
              ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
              f.�����id(+) = n_ȱʡ�����id And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
              Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
              Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
              Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.���֤�� = v_���֤�� And
              Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
              Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
              Decode(Nvl(v_���￨��, '-'), '-', '-', a.���￨��) = Nvl(v_���￨��, '-') And
              Decode(Nvl(v_Ic����, '-'), '-', '-', a.Ic����) = Nvl(v_Ic����, '-') And
              Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
              Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
              (v_����ids Is Null Or a.��ǰ����id Is Null Or
              a.��ǰ����id In
              (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
              (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
              (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    Else
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
               a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
               a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
               To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
               To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And
              a.���� = x.���(+) And (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And
              ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
              f.�����id(+) = n_ȱʡ�����id And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
              Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
              Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
              Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
              Instr(v_����ids, ',' || a.����id || ',') > 0 And Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
              Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
              Decode(Nvl(v_���￨��, '-'), '-', '-', a.���￨��) = Nvl(v_���￨��, '-') And
              Decode(Nvl(v_Ic����, '-'), '-', '-', a.Ic����) = Nvl(v_Ic����, '-') And
              Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
              Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
              (v_����ids Is Null Or a.��ǰ����id Is Null Or
              a.��ǰ����id In
              (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
              (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
              (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    End If;
  Elsif v_���￨�� Is Not Null Then
  
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And
            a.����id = f.����id(+) And f.�����id(+) = n_ȱʡ�����id And
            Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
            Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
            Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.���￨�� = v_���￨�� And
            Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
            Decode(Nvl(v_Ic����, '-'), '-', '-', a.Ic����) = Nvl(v_Ic����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
            (v_����ids Is Null Or a.��ǰ����id Is Null Or
            a.��ǰ����id In (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  
  Elsif v_Ic���� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
            f.�����id(+) = n_ȱʡ�����id And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
            Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
            Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And a.Ic���� = v_Ic���� And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And
            Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
            (v_����ids Is Null Or a.��ǰ����id Is Null Or
            a.��ǰ����id In (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif v_ҽ���� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
            f.�����id(+) = n_ȱʡ�����id And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
            Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
            Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
            (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And a.ҽ���� = v_ҽ���� And
            Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
            (v_����ids Is Null Or a.��ǰ����id Is Null Or
            a.��ǰ����id In (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif v_���� Is Not Null Then
  
    v_����ids := Zl_Custom_Patiids_Get(Nvl(n_����, 0), Null, v_����, v_�Ա�);
  
    If v_����ids Is Null Then
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
               a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
               a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
               To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
               To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And
              a.���� = x.���(+) And ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
              f.�����id(+) = n_ȱʡ�����id And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
              Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
              Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
              Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              ((n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
              (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
              Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
              Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
              (v_����ids Is Null Or a.��ǰ����id Is Null Or
              a.��ǰ����id In
              (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
              (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
              (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    Else
      Open c_������Ϣ For
        Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
               a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
               a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
               To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
               To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
        From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
        Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And
              a.���� = x.���(+) And (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And
              ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
              f.�����id(+) = n_ȱʡ�����id And Decode(d_��ʼ��������, Null, Sysdate, a.��������) Between Nvl(d_��ʼ��������, Sysdate) And
              Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
              Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
              Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
              (v_���� Is Null Or (n_Like = 1 And a.���� Like v_����) Or (n_Like = 0 And a.���� = v_����)) And
              Instr(v_����ids, ',' || a.����id || ',') > 0 And Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
              Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
              Decode(Nvl(v_ҽ����, '-'), '-', '-', a.ҽ����) = Nvl(v_ҽ����, '-') And
              Decode(Nvl(v_���￨��, '-'), '-', '-', a.���￨��) = Nvl(v_���￨��, '-') And
              Decode(Nvl(v_Ic����, '-'), '-', '-', a.Ic����) = Nvl(v_Ic����, '-') And
              Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
              Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
              (v_����ids Is Null Or a.��ǰ����id Is Null Or
              a.��ǰ����id In
              (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
              (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
              (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
    End If;
  
  Elsif d_��ʼ�������� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
            f.�����id(+) = n_ȱʡ�����id And a.�������� Between d_��ʼ�������� And Nvl(d_��ֹ��������, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_��ʼ�Ǽ�ʱ��, Null, Sysdate, a.�Ǽ�ʱ��) Between Nvl(d_��ʼ�Ǽ�ʱ��, Sysdate) And
            Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
            (v_����ids Is Null Or a.��ǰ����id Is Null Or
            a.��ǰ����id In (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And
            (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif d_��ʼ�Ǽ�ʱ�� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
            f.�����id(+) = n_ȱʡ�����id And a.�Ǽ�ʱ�� Between d_��ʼ�Ǽ�ʱ�� And Nvl(d_�����Ǽ�ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And
            Decode(d_���￪ʼʱ��, Null, Sysdate, Nvl(a.����ʱ��, a.�Ǽ�ʱ��)) Between Nvl(d_���￪ʼʱ��, Sysdate) And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
            (v_����ids Is Null Or a.��ǰ����id Is Null Or
            a.��ǰ����id In (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And
            (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif d_���￪ʼʱ�� Is Not Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
            f.�����id(+) = n_ȱʡ�����id And Nvl(a.����ʱ��, a.�Ǽ�ʱ��) Between d_���￪ʼʱ�� And
            Nvl(d_�������ʱ��, (Sysdate + 1 - 1 / 24 / 60 / 60)) And Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
            (v_����ids Is Null Or a.��ǰ����id Is Null Or
            a.��ǰ����id In (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (n_����Լ��λ = 0 Or a.��ͬ��λid Is Not Null) And (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And
            (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  
  Elsif Nvl(n_����Լ��λ, 0) = 1 Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, b.���� As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, f.����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, e.Id As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ҽ�Ƹ��ʽ B, ���ű� C, ���ű� D, ��Լ��λ E, ����ҽ�ƿ���Ϣ F, ������� X
      Where a.ҽ�Ƹ��ʽ = b.����(+) And a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.���� = x.���(+) And
            ((a.ͣ��ʱ�� Is Null And Nvl(n_�Ƿ�ͣ��, 0) = 0) Or Nvl(n_�Ƿ�ͣ��, 0) = 1) And a.����id = f.����id(+) And
            f.�����id(+) = n_ȱʡ�����id And Decode(Nvl(v_�ѱ�, '-'), '-', '-', a.�ѱ�) = Nvl(v_�ѱ�, '-') And
            Decode(Nvl(v_�Ա�, '-'), '-', '-', a.�Ա�) = Nvl(v_�Ա�, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_����, '-'), '-', '-', a.����) = Nvl(v_����, '-') And
            Decode(Nvl(v_ҽ�Ƹ��ʽ, '-'), '-', '-', a.ҽ�Ƹ��ʽ) = Nvl(v_ҽ�Ƹ��ʽ, '-') And
            Decode(Nvl(v_�ֻ���, '-'), '-', '-', a.�ֻ���) = Nvl(v_�ֻ���, '-') And
            (v_����ids Is Null Or a.��ǰ����id Is Null Or
            a.��ǰ����id In (Select ����id From �������Ҷ�Ӧ Where Instr(',' || v_����ids || ',', ',' || ����id || ',') > 0)) And
            (v_����ids Is Null Or n_��ѯסԺ״̬ = 2 And a.��ǰ����id Is Null Or Instr(v_����ids, ',' || a.��ǰ����id || ',') > 0) And
            (Nvl(n_��ͬ��λid, 0) = 0 And a.��ͬ��λid Is Not Null Or a.��ͬ��λid = n_��ͬ��λid) And
            (n_��ѯסԺ״̬ = 2 Or Nvl(a.��Ժ, 0) = Nvl(n_��ѯסԺ״̬, 0)) And (Nvl(n_Max, 0) = 0 Or Rownum < = Nvl(n_Max, Null));
  Elsif Not j_Json_Similar Is Null Then
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.�����ص�, a.���֤��, a.�����, a.סԺ��,
             a.���￨��, a.�ѱ�, a.ҽ�Ƹ��ʽ As ҽ�Ƹ��ʽ����, Null As ҽ�Ƹ��ʽ����, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ,
             a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
             a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, Null As ��ǰ��������, Null As ��ǰ��������, a.������λ As ������λ����, Null ����,
             To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�ֻ���, a.��ͥ�绰, Null As ��Լ��λid,
             To_Char(a.ͣ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ͣ��ʱ��, a.ҽ����, a.��ϵ������, a.��ϵ�˵绰, a.��λ��ַ, x.���� As ��������
      From ������Ϣ A, ������� X
      Where a.ͣ��ʱ�� Is Null And a.���� = x.���(+) And
            ((a.���� = v_���� And a.�Ա� = v_�Ա� And a.�������� = d_�������� And a.���� = v_���� And a.���� = v_����) Or a.���֤�� = v_���֤��)
      Order By Nvl(Nvl(a.����ʱ��, a.��Ժʱ��), a.�Ǽ�ʱ��) Desc;
  Else
    v_Err_Msg := 'δ������Ч�Ĳ�ѯ���������ܻ�ȡ������Ϣ!';
    Json_Out  := Get_Err_Message(v_Err_Msg);
    Return;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[';
  Loop
    Fetch c_������Ϣ
      Into r_����;
    Exit When c_������Ϣ%NotFound;
    Zljsonputvalue(v_List, 'pati_id', r_����.����id, 1, 1);
    Zljsonputvalue(v_List, 'pati_pageid', r_����.��ҳid, 1);
  
    Zljsonputvalue(v_List, 'pati_name', r_����.����);
    Zljsonputvalue(v_List, 'pati_sex', r_����.�Ա�);
  
    Zljsonputvalue(v_List, 'pati_age', r_����.����);
    Zljsonputvalue(v_List, 'pati_birthdate', r_����.��������);
    Zljsonputvalue(v_List, 'pati_birthplace', r_����.�����ص�);
    Zljsonputvalue(v_List, 'fee_category', r_����.�ѱ�);
  
    Zljsonputvalue(v_List, 'outpatient_num', r_����.�����, 0);
    Zljsonputvalue(v_List, 'inpatient_num', r_����.סԺ��, 0);
    Zljsonputvalue(v_List, 'inp_times', r_����.סԺ����, 1);
  
    Zljsonputvalue(v_List, 'pati_nation', r_����.����);
    Zljsonputvalue(v_List, 'pati_idcard', r_����.���֤��);
    Zljsonputvalue(v_List, 'vcard_no', r_����.���￨��);
  
    Zljsonputvalue(v_List, 'pati_education', r_����.ѧ��);
    Zljsonputvalue(v_List, 'ocpt_name', r_����.ְҵ);
  
    Zljsonputvalue(v_List, 'pati_identity', r_����.���);
    Zljsonputvalue(v_List, 'country_name', r_����.����);
    Zljsonputvalue(v_List, 'pat_home_addr', r_����.��ͥ��ַ);
    Zljsonputvalue(v_List, 'pati_area', r_����.����);
    Zljsonputvalue(v_List, 'emp_name', r_����.������λ����);
    Zljsonputvalue(v_List, 'emp_addr', r_����.��λ��ַ);
  
    Zljsonputvalue(v_List, 'is_inhspt', r_����.��Ժ, 1);
    Zljsonputvalue(v_List, 'pati_bed', r_����.��ǰ����);
    Zljsonputvalue(v_List, 'pati_type', r_����.��������);
    Zljsonputvalue(v_List, 'insurance_type', r_����.����, 1);
    Zljsonputvalue(v_List, 'insurance_type_name', r_����.��������);
    Zljsonputvalue(v_List, 'pati_wardarea_id', r_����.��ǰ����id, 1);
    Zljsonputvalue(v_List, 'pati_wardarea_name', r_����.��ǰ��������);
    Zljsonputvalue(v_List, 'pati_dept_id', r_����.��ǰ����id, 1);
    Zljsonputvalue(v_List, 'pati_dept_name', r_����.��ǰ��������);
  
    Zljsonputvalue(v_List, 'adta_time', r_����.��Ժʱ��);
    Zljsonputvalue(v_List, 'adtd_time', r_����.��Ժʱ��);
    Zljsonputvalue(v_List, 'create_time', r_����.�Ǽ�ʱ��);
    Zljsonputvalue(v_List, 'phone_number', r_����.�ֻ���);
    Zljsonputvalue(v_List, 'pat_home_phno', r_����.��ͥ�绰);
    If n_��ѯ���� = 1 Then
      Zljsonputvalue(v_List, 'stop_time', r_����.ͣ��ʱ��);
    Else
      Zljsonputvalue(v_List, 'stop_time', r_����.ͣ��ʱ��, 0, 2);
    End If;
    If n_��ѯ���� = 1 Then
      Zljsonputvalue(v_List, 'medc_card_no', r_����.����);
      Zljsonputvalue(v_List, 'visit_time', r_����.����ʱ��);
      Zljsonputvalue(v_List, 'ctt_unit_id', r_����.��Լ��λid, 1);
      Zljsonputvalue(v_List, 'mdlpay_mode_name', r_����.ҽ�Ƹ��ʽ����);
      Zljsonputvalue(v_List, 'mdlpay_mode_code', r_����.ҽ�Ƹ��ʽ����);
      Zljsonputvalue(v_List, 'insurance_num', r_����.ҽ����, 0);
      v_Listtmp := '"contacts":{"name":"' || Nvl(r_����.��ϵ������, '') || '","phone":"' || Nvl(r_����.��ϵ�˵绰, '') || '"}}';
      v_List    := v_List || ',' || v_Listtmp;
    End If;
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  End Loop;
  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiinfsbyrange;
/
Create Or Replace Procedure Zl_Patisvr_Getpatimergeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���id��ȡ���˵����кϲ���Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id             N 1 ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    merge_list[]        C       �ϲ���Ϣ�б�
  --      info_old          C   1   ԭ��Ϣ
  --      merge_reason      C   1   �ϲ�ԭ��
  --      operator_name     C   1   ����Ա
  --      merge_time        C   1   �ϲ�ʱ��:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  n_����id ���˲�����¼.����id%Type;
  v_Jtmp   Varchar2(32767);
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('ʧ�ܣ�δ���벡��id');
    Return;
  End If;
  For r_�ϲ� In (Select ԭ��Ϣ, �ϲ�ԭ��, ����Ա����, �ϲ�ʱ��
               From ���˺ϲ���¼
               Where ����id = n_����id
               Order By �ϲ�ʱ�� Desc) Loop
  
    v_Jtmp := v_Jtmp || ',{"info_old":"' || Zljsonstr(r_�ϲ�.ԭ��Ϣ) || '"';
    v_Jtmp := v_Jtmp || ',"merge_reason":"' || Zljsonstr(r_�ϲ�.�ϲ�ԭ��) || '"';
    v_Jtmp := v_Jtmp || ',"operator_name":"' || Zljsonstr(r_�ϲ�.����Ա����) || '"';
    v_Jtmp := v_Jtmp || ',"merge_time":"' || To_Char(r_�ϲ�.�ϲ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
  
    v_Jtmp := v_Jtmp || '}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","merge_list":[' || Substr(v_Jtmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatimergeinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatimmuneinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ϣ��������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id           N   1 ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    immune_list[]     C       ���������б�
  --      vaccinate_time    C   1   ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --      vaccinate_name    C   1   ��������
  ---------------------------------------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  n_����id �������߼�¼.����id%Type;
  v_Jtmp   Varchar2(32767);
  c_Jtmp   Clob;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ���벡��id������!');
    Return;
  End If;
  For r_���߼�¼ In (Select Distinct ����ʱ��, �������� From �������߼�¼ Where ����id = n_����id) Loop
    v_Jtmp := v_Jtmp || ',{"vaccinate_time":"' || To_Char(r_���߼�¼.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
    v_Jtmp := v_Jtmp || ',"vaccinate_name":"' || Zljsonstr(r_���߼�¼.��������) || '"';
    v_Jtmp := v_Jtmp || '}';
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","immune_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","immune_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatimmuneinfo;
/
Create Or Replace Procedure Zl_Patisvr_Getpatiphoto
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ������Ƭ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id           N 1 ����ID
  --����      json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     pati_photo         C 1 ����:base64
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_����id   ������Ƭ.����id%Type;
  b_������Ƭ ������Ƭ.��Ƭ%Type;
  v_Clob     Clob;
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ���벡��id�����ܱ��没����Ƭ!"}}';
    Return;
  End If;

  Begin
    Select ��Ƭ Into b_������Ƭ From ������Ƭ Where ����id = n_����id;
    v_Clob := Zltools.Zlbase64.Encode(b_������Ƭ);
    v_Clob := Replace(v_Clob, Chr(13), '');
    v_Clob := Replace(v_Clob, Chr(10), '');
  Exception
    When Others Then
      v_Clob := Null;
  End;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_photo":"' || v_Clob || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatiphoto;
/
Create Or Replace Procedure Zl_Patisvr_Getpatirelate
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡָ������֮�ص����Ĳ���ID�������в�������ǰ����Ĳ���ID
  --��Σ�Json_In:��ʽ
  --input
  --  query_type        N    1 �������� ��1-ͨ�����֤��ѯ��������ID,2-ͨ��������ݹ������ѯ�����Ĳ���id
  --  pati_id           N    1 ����id
  --  pati_idcard       C    1 ���֤��

  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_ids            C   1 ����id,����ƴ��
  -------------------------------------------

  j_Json   Pljson;
  j_Jsonin Pljson;
  v_List   Varchar2(32767);

  n_Type     Number(18);
  n_����id   Number(18);
  v_���֤�� Varchar2(50);
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_Type     := j_Json.Get_Number('query_type');
  n_����id   := j_Json.Get_Number('pati_id');
  v_���֤�� := j_Json.Get_String('pati_idcard');

  If n_Type = 1 Then
    For c_���� In (Select a.����id From ������Ϣ A Where a.����id <> n_����id And a.���֤�� = v_���֤��) Loop
      v_List := v_List || ',' || c_����.����id;
    End Loop;
  Elsif n_Type = 2 Then
    For c_���� In (Select b.����id
                 From ������ݹ��� A, ������ݹ��� B, ������Ϣ C
                 Where a.����id = b.����id And b.����id = c.����id And a.����id = n_����id And b.����id + 0 <> n_����id And
                       (Nvl(c.���֤��, '-') <> v_���֤�� Or v_���֤�� Is Null)) Loop
      v_List := v_List || ',' || c_����.����id;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_ids":"' || Substr(v_List, 2) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getpatirelate;
/
Create Or Replace Procedure Zl_Patisvr_Getvisitpatis
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���ﲡ����Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_ids          C   ����IDs:����ö���
  --      vcard_no          C   ���￨��

  --����      json
  --output
  -- code                   N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  -- message                C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -- pati_list[]                ������Ϣ�б�
  --   pati_id              N 1 ����id
  --   pati_name            C 1 ����
  --   pati_sex             C 1 �Ա�
  --   pati_age             C 1 ����
  --   pati_birthdate       C 1 �������ڣ�yyyy-mm-dd hh24:mi:ss
  --   fee_category         C 1 �ѱ�
  --   outpatient_num       C 1 �����
  --   pati_nation          C 1 ����
  --   pati_idcard          C 1 ���֤��
  --   vcard_no             C 1 ���￨��
  --   pati_education       C 1 ѧ��
  --   ocpt_name            C 1 ְҵ
  --   pati_identity        C 1 ���
  --   country_name         C 1 ����
  --   pat_home_addr        C 1 ��ͥ��ַ
  --   pati_area            C 1 ����
  --   emp_name             C 1 ������λ����
  --   pati_type            C 1 ��������(��ͨ��ҽ��������)
  --   insurance_type       C 1 ����
  --   create_time          C 1 �Ǽ�ʱ��
  --   pati_dept_id         N 1 ��ǰ����id
  --   pati_dept_name       C 1 ��ǰ��������
  --   iccard_no            C 1 Ic����
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Json    Pljson;
  j_Jsonin  Pljson;

  v_List Varchar2(32767);

  c_����ids Clob;

  v_���￨�� Varchar2(200);

  l_����id t_Strlist := t_Strlist();

  Cursor c_���˻�����Ϣ Is
    Select a.����id, a.����, a.�Ա�, a.����, a.��������, a.���֤��, a.�����, a.סԺ��, a.���￨��, a.�ѱ�, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��,
           a.��ͥ��ַ, a.������λ, a.סԺ����, a.��ǰ����id, a.��ǰ����id, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
           To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
           a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������, d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, a.Ic����
    From ������Ϣ A, ���ű� C, ���ű� D, ��Լ��λ E
    Where a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.ͣ��ʱ�� Is Null And Rownum < 1;
  r_���� c_���˻�����Ϣ%RowType;

  Type Ty_������Ϣ Is Ref Cursor;
  c_������Ϣ Ty_������Ϣ; --��̬�α����
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  Begin
    c_����ids := j_Json.Get_Clob('pati_ids');
  Exception
    When Others Then
      c_����ids := Null;
  End;
  v_���￨�� := j_Json.Get_String('vcard_no');
  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      l_����id.Extend;
      l_����id(l_����id.Count) := c_����ids;
      c_����ids := Null;
    Else
      l_����id.Extend;
      l_����id(l_����id.Count) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
    End If;
  End Loop;

  If l_����id.Count <> 0 Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[';
    For I In 1 .. l_����id.Count Loop
      For c_���˲�ѯ In (With c_���� As
                        (Select Column_Value As ����id From Table(f_Num2list(l_����id(I))))
                       Select /*+cardinality(B,10)*/
                        a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.���֤��, a.�����, a.סԺ��,
                        a.���￨��, a.�ѱ�, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ, a.סԺ����, a.��ǰ����id, a.��ǰ����id,
                        To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��,
                        To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����,
                        To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������,
                        d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, a.Ic����
                       From ������Ϣ A, c_���� B, ���ű� C, ���ű� D, ��Լ��λ E
                       Where a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.ͣ��ʱ�� Is Null And
                             a.����id = b.����id) Loop
        Zljsonputvalue(v_List, 'pati_id', c_���˲�ѯ.����id, 1, 1);
        Zljsonputvalue(v_List, 'pati_name', c_���˲�ѯ.����, 0);
        Zljsonputvalue(v_List, 'pati_sex', Nvl(c_���˲�ѯ.�Ա�, ''), 0);
        Zljsonputvalue(v_List, 'pati_age', Nvl(c_���˲�ѯ.����, ''), 0);
        Zljsonputvalue(v_List, 'pati_birthdate', Nvl(c_���˲�ѯ.��������, ''), 0);
        Zljsonputvalue(v_List, 'fee_category', Nvl(c_���˲�ѯ.�ѱ�, ''), 0);
        Zljsonputvalue(v_List, 'outpatient_num', c_���˲�ѯ.�����, 0);
        Zljsonputvalue(v_List, 'pati_nation', Nvl(c_���˲�ѯ.����, ''), 0);
        Zljsonputvalue(v_List, 'pati_idcard', Nvl(c_���˲�ѯ.���֤��, ''), 0);
        Zljsonputvalue(v_List, 'vcard_no', Nvl(c_���˲�ѯ.���￨��, ''), 0);
        Zljsonputvalue(v_List, 'pati_education', Nvl(c_���˲�ѯ.ѧ��, ''), 0);
        Zljsonputvalue(v_List, 'ocpt_name', Nvl(c_���˲�ѯ.ְҵ, ''), 0);
        Zljsonputvalue(v_List, 'pati_identity', Nvl(c_���˲�ѯ.���, ''), 0);
        Zljsonputvalue(v_List, 'country_name', Nvl(c_���˲�ѯ.����, ''), 0);
        Zljsonputvalue(v_List, 'pat_home_addr', Nvl(c_���˲�ѯ.��ͥ��ַ, ''), 0);
        Zljsonputvalue(v_List, 'pati_area', Nvl(c_���˲�ѯ.����, ''), 0);
        Zljsonputvalue(v_List, 'emp_name', Nvl(c_���˲�ѯ.������λ����, ''), 0);
        Zljsonputvalue(v_List, 'pati_type', Nvl(c_���˲�ѯ.��������, ''), 0);
        Zljsonputvalue(v_List, 'insurance_type', Nvl(c_���˲�ѯ.����, ''), 0);
        Zljsonputvalue(v_List, 'pati_dept_id', c_���˲�ѯ.��ǰ����id, 1);
        Zljsonputvalue(v_List, 'pati_dept_name', Nvl(c_���˲�ѯ.��ǰ��������, ''), 0);
        Zljsonputvalue(v_List, 'create_time', Nvl(c_���˲�ѯ.�Ǽ�ʱ��, ''), 0);
        Zljsonputvalue(v_List, 'iccard_no', Nvl(c_���˲�ѯ.Ic����, ''), 0, 2);
        If Length(v_List) > 20000 Then
          Json_Out := Json_Out || v_List;
          v_List   := ',';
        End If;
      End Loop;
    End Loop;
  
    If v_List <> ',' Then
      Json_Out := Json_Out || v_List || ']}}';
    Else
      Json_Out := Json_Out || ']}}';
    End If;
    Return;
  Elsif v_���￨�� Is Not Null Then
  
    Open c_������Ϣ For
      Select a.����id, a.����, a.�Ա�, a.����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, a.���֤��, a.�����, a.סԺ��, a.���￨��,
             a.�ѱ�, a.���, a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.��ͥ��ַ, a.������λ, a.סԺ����, a.��ǰ����id, a.��ǰ����id,
             To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, To_Char(a.��Ժʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Ժʱ��, a.����,
             To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.��Ժ, a.��ǰ����, a.��������, a.��ҳid, c.���� As ��ǰ��������,
             d.���� As ��ǰ��������, Nvl(e.����, a.������λ) As ������λ����, a.Ic����
      From ������Ϣ A, ���ű� C, ���ű� D, ��Լ��λ E
      Where a.��ǰ����id = c.Id(+) And a.��ǰ����id = d.Id(+) And a.��ͬ��λid = e.Id(+) And a.ͣ��ʱ�� Is Null And a.���￨�� = v_���￨��;
  Else
    v_Err_Msg := 'δ������Ч�Ĳ�ѯ���������ܻ�ȡ������Ϣ!';
    Json_Out  := Zljsonout(v_Err_Msg);
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[';
  Loop
    Fetch c_������Ϣ
      Into r_����;
    Exit When c_������Ϣ%NotFound;
  
    Zljsonputvalue(v_List, 'pati_id', r_����.����id, 1, 1);
    Zljsonputvalue(v_List, 'pati_name', r_����.����, 0);
    Zljsonputvalue(v_List, 'pati_sex', Nvl(r_����.�Ա�, ''), 0);
    Zljsonputvalue(v_List, 'pati_age', Nvl(r_����.����, ''), 0);
    Zljsonputvalue(v_List, 'pati_birthdate', Nvl(r_����.��������, ''), 0);
    Zljsonputvalue(v_List, 'fee_category', Nvl(r_����.�ѱ�, ''), 0);
    Zljsonputvalue(v_List, 'outpatient_num', r_����.�����, 0);
    Zljsonputvalue(v_List, 'pati_nation', Nvl(r_����.����, ''), 0);
    Zljsonputvalue(v_List, 'pati_idcard', Nvl(r_����.���֤��, ''), 0);
    Zljsonputvalue(v_List, 'vcard_no', Nvl(r_����.���￨��, ''), 0);
    Zljsonputvalue(v_List, 'pati_education', Nvl(r_����.ѧ��, ''), 0);
    Zljsonputvalue(v_List, 'ocpt_name', Nvl(r_����.ְҵ, ''), 0);
    Zljsonputvalue(v_List, 'pati_identity', Nvl(r_����.���, ''), 0);
    Zljsonputvalue(v_List, 'country_name', Nvl(r_����.����, ''), 0);
    Zljsonputvalue(v_List, 'pat_home_addr', Nvl(r_����.��ͥ��ַ, ''), 0);
    Zljsonputvalue(v_List, 'pati_area', Nvl(r_����.����, ''), 0);
    Zljsonputvalue(v_List, 'emp_name', Nvl(r_����.������λ����, ''), 0);
    Zljsonputvalue(v_List, 'pati_type', Nvl(r_����.��������, ''), 0);
    Zljsonputvalue(v_List, 'insurance_type', Nvl(r_����.����, ''), 0);
    Zljsonputvalue(v_List, 'pati_dept_id', r_����.��ǰ����id, 1);
    Zljsonputvalue(v_List, 'pati_dept_name', Nvl(r_����.��ǰ��������, ''), 0);
    Zljsonputvalue(v_List, 'create_time', Nvl(r_����.�Ǽ�ʱ��, ''), 0);
    Zljsonputvalue(v_List, 'iccard_no', Nvl(r_����.Ic����, ''), 0, 2);
    If Length(v_List) > 20000 Then
      Json_Out := Json_Out || v_List;
      v_List   := ',';
    End If;
  
  End Loop;

  If v_List <> ',' Then
    Json_Out := Json_Out || v_List || ']}}';
  Else
    Json_Out := Json_Out || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Getvisitpatis;
/
Create Or Replace Procedure Zl_Patisvr_Lockcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���鲡���Ƿ�����
  --��� JSON��ʽ
  --input
  --  pati_id     N 1 ����id
  --���Σ�JSON��ʽ
  --output
  --  code        N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --  message     C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  v_����      Varchar2(2550);
  v_����      Number;
  j_Json      Pljson;
  j_Jsoninput Pljson;
  n_����id    Number;
Begin
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  n_����id    := j_Json.Get_Number('pati_id');
  Select ����, ���� Into v_����, v_���� From ������Ϣ Where ����id = n_����id;
  If Nvl(v_����, 0) = 1 Then
    Json_Out := '{"output":{"code":0,"message":"���ˡ�' || v_���� || '����ǰ�ѱ���������������κβ�������ȴ�һ��ʱ������ԡ�"}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Lockcheck;
/
Create Or Replace Procedure Zl_Patisvr_Newpatiarchives
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------
  --���ܣ��²��˽���
  --���      json
  --  input
  --    pati_id               N  1  ����id
  --    pati_name             C  1  ����
  --    pati_sex              C  1  �Ա�
  --    pati_age              C  1  ����
  --    pati_birthdate        C  1  ��������:yyyy-mm-dd hh24:mi:ss
  --    pati_idcard           C  1  ���֤��
  --    pati_type             C  1  ��������(��ͨ��ҽ��������)
  --    outpatient_num        C  1  �����
  --    vcard_no              C  1  ���￨��
  --    vcard_pwd             C  1  ����֤��
  --    fee_category          C  1  �ѱ�
  --    mdlpay_mode_name      C  1  ҽ�Ƹ��ʽ����
  --    native_place          C  1  ����
  --    country_name          C  1  ����
  --    nation_name           C  1  ����
  --    mari_status           C  1  ����״��
  --    edu_name              C  1  ѧ��
  --    ocpt_name             C  1  ְҵ
  --    pati_identity         C  1  ���
  --    emp_name              C  1  ������λ
  --    emp_postcode          C  1  ��λ�ʱ�
  --    emp_phno              C  1  ��λ�绰
  --    emp_bank_name       C   1   ��λ������
  --    emp_bank_accnum     C   1   ��λ�ʺ�
  --    ctt_unit_id           N  1  ��ͬ��λid
  --    pat_home_addr         C  1  ��ͥ��ַ
  --    pat_home_phno         C  1  ��ͥ�绰
  --    pat_home_postcode     C  1  ��ͥ��ַ�ʱ�
  --    region                C  1  ����
  --    pat_baddr             C  1  �����ص�
  --    pat_hous_addr         C  1  ���ڵ�ַ
  --    pat_hous_postcode     C  1  ���ڵ�ַ�ʱ�
  --    pat_grdn_name         C  1  �໤��
  --    phone_number          C  1  �ֻ���
  --    insurance_num         C  1  ҽ����
  --    iccard_no             C  1  Ic����
  --    create_time           C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    operator_name         C  1  ����Ա����
  --    idcard_sign           N     ���֤ǩԼ
  --    idcard_sign_pwd       C     ǩԼ����
  --    insurance_type      N   1   ����
  --    cert_no_other       C   1   ����֤��
  --    contacts              C     ������ϵ����Ϣ�ڵ�
  --      name                C  1  ��ϵ������
  --      idcard              C  1  ��ϵ�����֤��
  --      phone               C  1  ��ϵ�˵绰
  --      relation            C  1  ��ϵ�˹�ϵ
  --      address             C     ��ϵ�˵�ַ
  --    community_info        C     ������Ϣ�ڵ�
  --      num                 N  1  �������
  --      code                C  1  ��������
  --      oper_type           N  1  ������������
  --    visit_info            C     ������Ϣ�ڵ�
  --      statu               N     ���µľ���״̬
  --      room                C     ���µľ�������
  --      time                C     ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --    addr_list[]           C     ��ַ��Ϣ�б�
  --      oper_fun            N  1  ��������:1-����,�޸�   2-ɾ��
  --      type                C  1  ��ַ���
  --      state               C  1  ��ַ_ʡ
  --      city                C  1  ��ַ_��
  --      county              C  1  ��ַ_��
  --      township            C  1  ��ַ_��
  --      other               C  1  ��ַ_����
  --      code                C  1  ��������
  --    ext_list[]            C     ������Ϣ�����б�
  --      info_name           C  1  ��Ϣ��
  --      upd_info_value      N  1  �޸ĵ���Ϣֵ
  --    cert_list[]                 ֤���б�(��Ҫ�ǵ��ɰ󿨴���)
  --      cert_name           C  1  ֤������
  --      cert_no             C  1  ֤�ź���
  --    allergic_drugs_list[]       ���˹���ҩ���б�:������ʱ������ɾ������ҩ�����ķ�ʽ
  --      pat_algc_cadn_id    N  1  ����ҩƷID
  --      pat_algc_cadn       C  1  ����ҩ������
  --      allergy_info        C  1  ��ÿҩ�ﷴӦ
  --    immune_list[]         C     ���������б�
  --      vaccinate_time      C  1  ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --      vaccinate_name      C  1  ��������
  --    card_property_list[]  C     ҽ�ƿ������б�
  --      cardtype_id         N  1  ҽ�ƿ����ID
  --      card_no             C  1  ����
  --      info_name           C  1  ��Ϣ��
  --      info_value          N  1  ��Ϣֵ

  --����      json
  --  output
  --    code  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message C 1Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ
  -----------------------------------------------------------------------------------------------------
  n_����id         ������Ϣ.����id%Type;
  v_����           ������Ϣ.����%Type;
  v_���֤��       ������Ϣ.���֤��%Type;
  v_��������       ������Ϣ.��������%Type;
  v_����           ������Ϣ.����%Type;
  v_�Ա�           ������Ϣ.�Ա�%Type;
  d_��������       ������Ϣ.��������%Type;
  v_���䵥λ       Varchar2(20);
  v_�ֻ���         ������Ϣ.�ֻ���%Type;
  v_��ͥ�绰       ������Ϣ.��ͥ�绰%Type;
  n_�����         ������Ϣ.�����%Type;
  v_�ѱ�           ������Ϣ.�ѱ�%Type;
  v_ҽ�Ƹ��ʽ   ������Ϣ.ҽ�Ƹ��ʽ%Type;
  v_����           ������Ϣ.����%Type;
  v_����           ������Ϣ.����%Type;
  v_����           ������Ϣ.����%Type;
  v_����           ������Ϣ.����״��%Type;
  v_ְҵ           ������Ϣ.ְҵ%Type;
  v_ѧ��           ������Ϣ.ѧ��%Type;
  v_������λ       ������Ϣ.������λ%Type;
  n_��ͬ��λid     ������Ϣ.��ͬ��λid%Type;
  v_��λ�绰       ������Ϣ.��λ�绰%Type;
  v_��λ�ʱ�       ������Ϣ.��λ�ʱ�%Type;
  v_��λ������     ������Ϣ.��λ������%Type;
  v_��λ�ʺ�       ������Ϣ.��λ�ʺ�%Type;
  v_��ͥ��ַ       ������Ϣ.��ͥ��ַ%Type;
  v_��ͥ��ַ�ʱ�   ������Ϣ.��ͥ��ַ�ʱ�%Type;
  v_���ڵ�ַ       ������Ϣ.���ڵ�ַ%Type;
  v_���ڵ�ַ�ʱ�   ������Ϣ.���ڵ�ַ�ʱ�%Type;
  d_�Ǽ�ʱ��       ������Ϣ.�Ǽ�ʱ��%Type;
  v_ҽ����         ������Ϣ.ҽ����%Type;
  v_����           ������Ϣ.����%Type;
  v_��ϵ�����֤�� ������Ϣ.��ϵ�����֤��%Type;
  v_��ϵ������     ������Ϣ.��ϵ������%Type;
  v_��ϵ�˵绰     ������Ϣ.��ϵ�˵绰%Type;
  v_��ϵ�˹�ϵ     ������Ϣ.��ϵ�˹�ϵ%Type;
  v_��ϵ�˵�ַ     ������Ϣ.��ϵ�˵�ַ%Type;
  v_�໤��         ������Ϣ.�໤��%Type;
  v_�����ص�       ������Ϣ.�����ص�%Type;
  v_���           ������Ϣ.���%Type;
  v_����Ա����     ����ҽ�ƿ���Ϣ.������%Type;
  n_����id         ����������Ϣ.����%Type;
  v_��������       ����������Ϣ.������%Type;
  n_��������       ����������Ϣ.��������%Type;
  n_��ַ���       ���˵�ַ��Ϣ.��ַ���%Type;
  v_��ַ_ʡ        ���˵�ַ��Ϣ.ʡ%Type;
  v_��ַ_��        ���˵�ַ��Ϣ.��%Type;
  v_��ַ_��        ���˵�ַ��Ϣ.��%Type;
  v_��ַ_��        ���˵�ַ��Ϣ.����%Type;
  v_��ַ_����      ���˵�ַ��Ϣ.����%Type;
  v_��������       ���˵�ַ��Ϣ.��������%Type;
  v_��Ϣ��         ������Ϣ�ӱ�.��Ϣ��%Type;
  v_��Ϣֵ         ������Ϣ�ӱ�.��Ϣֵ%Type;
  v_������         ҽ�ƿ����.����%Type;
  n_�����id       ҽ�ƿ����.Id%Type;
  v_����           ����ҽ�ƿ���Ϣ.����%Type;
  n_���ų���       ҽ�ƿ����.���ų���%Type;
  v_����           ҽ�ƿ����.����%Type;
  v_������         ����ҽ�ƿ���Ϣ.����%Type;
  v_�䶯ԭ��       ����ҽ�ƿ��䶯.�䶯ԭ��%Type;
  d_��ֹʹ��ʱ��   ����ҽ�ƿ���Ϣ.��ֹʹ��ʱ��%Type;
  n_����ҩƷid     ���˹���ҩ��.����ҩ��id%Type;
  v_����ҩ������   ���˹���ҩ��.����ҩ��%Type;
  v_��ÿҩ�ﷴӦ   ���˹���ҩ��.������Ӧ%Type;
  d_����ʱ��       �������߼�¼.����ʱ��%Type;
  v_��������       �������߼�¼.��������%Type;
  v_���￨��       ������Ϣ.���￨��%Type;
  v_����֤��       ������Ϣ.����֤��%Type;
  v_Ic����         ������Ϣ.Ic����%Type;
  n_��������       Number(2);
  n_Ψһ���֤     Number(2);
  n_����״̬       ������Ϣ.����״̬%Type;
  v_��������       ������Ϣ.��������%Type;
  d_����ʱ��       ������Ϣ.����ʱ��%Type;
  n_����           ������Ϣ.����%Type;
  v_����֤��       ������Ϣ.����֤��%Type;
  n_�ֵ         Number(10);
  n_���ֵ         Number(10);

  n_Count    Number(2);
  j_Input    Pljson;
  o_Json     Pljson;
  o_Json1    Pljson;
  j_Jsonlist Pljson_List := Pljson_List();
  j_Jsonin   Pljson;
Begin
  j_Jsonin := Pljson(Json_In);
  j_Input  := j_Jsonin.Get_Pljson('input');
  --    pati_id               N  1  ����id
  --    pati_name             C  1  ����
  --    pati_sex              C  1  �Ա�
  --    pati_age              C  1  ����
  --    pati_birthdate        C  1  ��������:yyyy-mm-dd hh24:mi:ss
  --    pati_type             C   1   ��������(��ͨ��ҽ��������)
  n_����id   := j_Input.Get_Number('pati_id');
  v_����     := j_Input.Get_String('pati_name');
  v_�Ա�     := j_Input.Get_String('pati_sex');
  v_����     := j_Input.Get_String('pati_age');
  d_�������� := To_Date(j_Input.Get_String('pati_birthdate'), 'YYYY-MM-DD hh24:mi:ss');
  v_�������� := j_Input.Get_String('pati_type');
  --    pati_idcard           C  1  ���֤��
  --    outpno                N  1  �����
  --    vcard_no              C  1  ���￨��
  --    vcard_pwd             C  1  ����֤��
  --    fee_category          C  1  �ѱ�
  --    mdlpay_mode_name      C  1  ҽ�Ƹ��ʽ����
  v_���֤��     := j_Input.Get_String('pati_idcard');
  n_�����       := To_Number(j_Input.Get_String('outpatient_num'));
  v_���￨��     := j_Input.Get_String('vcard_no');
  v_����֤��     := j_Input.Get_String('vcard_pwd');
  v_�ѱ�         := j_Input.Get_String('fee_category');
  v_ҽ�Ƹ��ʽ := j_Input.Get_String('mdlpay_mode_name');

  --    native_place          C  1  ����
  --    country_name          C  1  ����
  --    nation_name           C  1  ����
  --    mari_status           C  1  ����״��
  --    ocpt_name             C  1  ְҵ
  --    edu_name              C  1  ѧ��
  --    pati_identity         C  1  ���
  v_���� := j_Input.Get_String('native_place');
  v_���� := j_Input.Get_String('country_name');
  v_���� := j_Input.Get_String('nation_name');
  v_���� := j_Input.Get_String('mari_status');
  v_ְҵ := j_Input.Get_String('ocpt_name');
  v_��� := j_Input.Get_String('pati_identity');
  v_ѧ�� := j_Input.Get_String('edu_name');
  --    emp_name              C  1  ������λ
  --    emp_postcode          C  1  ��λ�ʱ�
  --    emp_phno              C  1  ��λ�绰
  --    emp_bank_name       C   1   ��λ������
  --    emp_bank_accnum     C   1   ��λ�ʺ�
  --    ctt_unit_id           N  1  ��ͬ��λid
  --    pat_home_addr         C  1  ��ͥ��ַ
  --    pat_home_phno         C  1  ��ͥ�绰
  --    pat_home_postcode     C  1  ��ͥ��ַ�ʱ�
  v_������λ     := j_Input.Get_String('emp_name');
  v_��λ�绰     := j_Input.Get_String('emp_phno');
  v_��λ�ʱ�     := j_Input.Get_String('emp_postcode');
  v_��λ������   := j_Input.Get_String('emp_bank_name');
  v_��λ�ʺ�     := j_Input.Get_String('emp_bank_accnum');
  n_��ͬ��λid   := j_Input.Get_Number('ctt_unit_id');
  v_��ͥ��ַ     := j_Input.Get_String('pat_home_addr');
  v_��ͥ�绰     := j_Input.Get_String('pat_home_phno');
  v_��ͥ��ַ�ʱ� := j_Input.Get_String('pat_home_postcode');

  --    region                C  1  ����
  --    pat_baddr             C  1  �����ص�
  --    pat_hous_addr         C  1  ���ڵ�ַ
  --    pat_hous_postcode     C  1  ���ڵ�ַ�ʱ�

  v_����         := j_Input.Get_String('region');
  v_�����ص�     := j_Input.Get_String('pat_baddr');
  v_���ڵ�ַ     := j_Input.Get_String('pat_hous_addr');
  v_���ڵ�ַ�ʱ� := j_Input.Get_String('pat_hous_postcode');

  --    pat_grdn_name         C  1  �໤��
  --    phone_number          C  1  �ֻ���
  --    insurance_num         C  1  ҽ����
  --    iccard_no             C  1  Ic����
  --    insurance_type        N  1  ����
  --    create_time           C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    operator_name         C  1  ����Ա����
  --    idcard_sign           N     ���֤ǩԼ
  --    idcard_sign_pwd       C     ǩԼ����

  v_�໤��   := j_Input.Get_String('pat_grdn_name');
  v_�ֻ���   := j_Input.Get_String('phone_number');
  v_ҽ����   := j_Input.Get_String('insurance_num');
  v_Ic����   := j_Input.Get_String('iccard_no');
  d_�Ǽ�ʱ�� := To_Date(j_Input.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;
  v_����Ա���� := j_Input.Get_String('operator_name');
  --    insurance_type      N   1   ����
  --    cert_no_other       C   1   ����֤��
  n_����     := j_Input.Get_Number('insurance_type');
  v_����֤�� := j_Input.Get_String('cert_no_other');

  If v_���֤�� Is Not Null Then
    n_Ψһ���֤ := Nvl(zl_GetSysParameter(279), 0);
    If n_Ψһ���֤ = 1 Then
      --������֤Ψһ��
      Select Count(1) Into n_Count From ������Ϣ Where ���֤�� = v_���֤�� And ����id <> n_����id;
      If n_Count <> 0 Then
        Json_Out := Zljsonout('�Ѿ��������֤��Ϊ' || v_���֤�� || '�Ĳ���,������¼����ͬ�����֤�ţ�');
        Return;
      End If;
    End If;
  End If;

  If d_�������� Is Null And v_���� Is Not Null Then
    --�����������������
    v_���䵥λ := Substr(v_����, Length(v_����), 1);
    If Instr('��,��,��', v_���䵥λ) <= 0 Then
      v_���䵥λ := Null;
    Else
      v_���� := Replace(v_����, v_���䵥λ, '');
    End If;
    Begin
      v_���� := To_Number(v_����);
    Exception
      When Others Then
        v_���� := Null;
    End;
    If v_���� Is Not Null And v_���䵥λ Is Not Null Then
      Select Decode(v_���䵥λ, '��', Add_Months(Sysdate, -12 * v_����), '��', Add_Months(Sysdate, -1 * v_����), '��',
                     Sysdate - v_����)
      Into d_��������
      From Dual;
    End If;
  End If;

  If v_�ֻ��� Is Null And v_��ͥ�绰 Is Not Null Then
    Select Count(1)
    Into n_Count
    From �ֻ��ų��úŶα�
    Where Length(v_��ͥ�绰) = ���볤�� And v_��ͥ�绰 Like �Ŷ� || '%';
    If n_Count <> 0 Then
      v_�ֻ��� := v_��ͥ�绰;
    End If;
  End If;

  --��ϵ����Ϣ
  --    contacts              C     ������ϵ����Ϣ�ڵ�
  --      name                C  1  ��ϵ������
  --      idcard              C  1  ��ϵ�����֤��
  --      phone               C  1  ��ϵ�˵绰
  --      relation            C  1  ��ϵ�˹�ϵ
  --      address             C     ��ϵ�˵�ַ
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('contacts');
  If o_Json1 Is Not Null Then
    v_��ϵ������     := o_Json1.Get_String('name');
    v_��ϵ�����֤�� := o_Json1.Get_String('idcard');
    v_��ϵ�˵绰     := o_Json1.Get_String('phone');
    v_��ϵ�˹�ϵ     := o_Json1.Get_String('relation');
    v_��ϵ�˵�ַ     := o_Json1.Get_String('address');
  End If;

  --������Ϣ
  --    visit_info            C     ������Ϣ�ڵ�
  --      visit_statu         N     ���µľ���״̬
  --      visit_room          C     ���µľ�������
  --      visit_time          C     ����ʱ��:yyyy-mm-dd hh24:mi:ss
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('visit_info');
  If o_Json1 Is Not Null Then
    n_����״̬ := o_Json1.Get_Number('visit_statu');
    v_�������� := o_Json1.Get_String('visit_room');
    d_����ʱ�� := To_Date(o_Json1.Get_String('visit_time'), 'yyyy-mm-dd hh24:mi:ss');
  End If;

  --�²�����Ϣ
  Insert Into ������Ϣ
    (����id, �����, ����, �Ա�, ����, ��������, �ѱ�, ҽ�Ƹ��ʽ, ����, ����, ����, ����״��, ְҵ, ѧ��, ��������, ���֤��, ������λ, ��λ������, ��λ�ʺ�, ��ͬ��λid, ��λ�绰,
     ��λ�ʱ�, ��ͥ��ַ, ��ͥ�绰, ��ͥ��ַ�ʱ�, ���ڵ�ַ, ���ڵ�ַ�ʱ�, �Ǽ�ʱ��, ҽ����, ����, ��ϵ�����֤��, ��ϵ������, ��ϵ�˵绰, ��ϵ�˹�ϵ, ��ϵ�˵�ַ, �໤��, �����ص�, �ֻ���, ���,
     ���￨��, ����֤��, Ic����, ����״̬, ����ʱ��, ��������, ����, ����֤��)
  Values
    (n_����id, n_�����, v_����, v_�Ա�, v_����, d_��������, v_�ѱ�, v_ҽ�Ƹ��ʽ, v_����, v_����, v_����, v_����, v_ְҵ, v_ѧ��, v_��������, v_���֤��,
     v_������λ, v_��λ������, v_��λ�ʺ�, Decode(n_��ͬ��λid, 0, Null, n_��ͬ��λid), v_��λ�绰, v_��λ�ʱ�, v_��ͥ��ַ, v_��ͥ�绰, v_��ͥ��ַ�ʱ�, v_���ڵ�ַ,
     v_���ڵ�ַ�ʱ�, d_�Ǽ�ʱ��, v_ҽ����, v_����, v_��ϵ�����֤��, v_��ϵ������, v_��ϵ�˵绰, v_��ϵ�˹�ϵ, v_��ϵ�˵�ַ, v_�໤��, v_�����ص�, v_�ֻ���, v_���, v_���￨��,
     v_����֤��, v_Ic����, n_����״̬, d_����ʱ��, v_��������, Decode(n_����, 0, Null, n_����), v_����֤��);

  --������Ϣ
  --    community_info        C     ������Ϣ�ڵ�
  --      num                 N  1  �������
  --      code                C  1  ��������
  --      oper_type           N  1  ������������
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('community_info');
  If o_Json1 Is Not Null Then
    n_����id   := o_Json1.Get_Number('num');
    v_�������� := o_Json1.Get_String('code');
    n_�������� := o_Json1.Get_Number('oper_type');
  End If;
  --����������
  If n_����id <> 0 And v_�������� Is Not Null Then
    Zl_����������Ϣ_Insert(n_����id, n_����id, v_��������, n_��������, d_�Ǽ�ʱ��);
  End If;

  --���µ�ַ��Ϣ
  --    addr_list[]           C     ��ַ��Ϣ�б�
  --      oper_fun            N  1  ��������:1-����,�޸�   2-ɾ��
  --      type                C  1  ��ַ���
  --      state               C  1  ��ַ_ʡ
  --      city                C  1  ��ַ_��
  --      county              C  1  ��ַ_��
  --      township            C  1  ��ַ_��
  --      other               C  1  ��ַ_����
  --      code                C  1  ��������
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('addr_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json     := Pljson();
      o_Json     := Pljson(j_Jsonlist.Get(I));
      n_��������  := o_Json.Get_Number('oper_fun');
      n_��ַ���  := o_Json.Get_Number('type');
      v_��ַ_ʡ   := o_Json.Get_String('state');
      v_��ַ_��   := o_Json.Get_String('city');
      v_��ַ_��   := o_Json.Get_String('county');
      v_��ַ_��   := o_Json.Get_String('township');
      v_��ַ_���� := o_Json.Get_String('other');
      v_��������  := o_Json.Get_String('code');
    
      Zl_���˵�ַ��Ϣ_Update_s(n_��������, n_����id, Null, n_��ַ���, v_��ַ_ʡ, v_��ַ_��, v_��ַ_��, v_��ַ_��, v_��ַ_����, v_��������);
    End Loop;
  End If;

  --���²��˴�����Ϣ
  --    ext_list[]            C     ������Ϣ�����б�
  --      info_name           C  1  ��Ϣ��
  --      upd_info_value      N  1  �޸ĵ���Ϣֵ
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('ext_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      v_��Ϣ�� := o_Json.Get_String('info_name');
      v_��Ϣֵ := o_Json.Get_String('upd_info_value');
    
      If v_��Ϣ�� Is Not Null And v_��Ϣֵ Is Not Null Then
        Zl_������Ϣ�ӱ�_Update(n_����id, v_��Ϣ��, v_��Ϣֵ);
      End If;
    End Loop;
  End If;

  --����֤������
  --    cert_list[]                 ֤���б�(��Ҫ�ǵ��ɰ󿨴���)
  --      cert_name           C  1  ֤������
  --      cert_no             C  1  ֤�ź���
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('cert_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json   := Pljson();
      o_Json   := Pljson(j_Jsonlist.Get(I));
      v_������ := o_Json.Get_String('cert_name');
      v_����   := o_Json.Get_String('cert_no');
    
      If v_������ Is Not Null Then
        If v_���� Is Not Null Then
          --��鿨���Ƿ�����ʹ��
          Select Count(1)
          Into n_Count
          From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���� B
          Where a.�����id = b.Id And b.���� = v_������ And b.�Ƿ�֤�� = 1 And a.���� = v_���� And a.����id <> n_����id;
          If n_Count <> 0 Then
            Json_Out := Zljsonout(v_������ || ':' || v_���� || '���ڱ�����ʹ��,���飡');
            Return;
          End If;
        
          --�����ڵľ���������Ҫ������������
          Select Nvl(Max(ID), 0), Nvl(Max(���ų���), 0), Max(����), Max(LPad(����, 10)), Max(Length(����))
          Into n_�����id, n_���ų���, v_����, n_���ֵ, n_�ֵ
          From ҽ�ƿ����
          Where ���� = v_������;
        
          Select Max(����), Max(LPad(����, 10)), Max(Length(����)) Into v_����, n_���ֵ, n_�ֵ From ҽ�ƿ����;
        
          If v_���� Is Null Then
            Select LPad(1, 10, '0') Into v_���� From Dual;
          Else
            n_���ֵ := n_���ֵ + 1;
            Select LPad(n_���ֵ, n_�ֵ, '0') Into v_���� From Dual;
          End If;
          If n_�����id = 0 Then
            --����
            Select ҽ�ƿ����_Id.Nextval Into n_�����id From Dual;
            Zl_ҽ�ƿ����_Update(n_�����id, v_����, v_������, Substr(v_������, 1, 1), Null, Length(v_����), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_����, 0, Null, 1, Null, 1, 10, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          Elsif Length(v_����) > n_���ų��� Then
            --�޸ĳ���
            Zl_ҽ�ƿ����_Update(n_�����id, v_����, v_������, Substr(v_������, 1, 1), Null, Length(v_����), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_����, 0, Null, 1, Null, 1, 10, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          End If;
        End If;
      
        --��������˿���Ϣ
        n_Count := 0;
        For c_֤�� In (Select a.�����id, a.����
                     From ����ҽ�ƿ���Ϣ A
                     Where a.�����id = n_�����id And a.����id = n_����id) Loop
          If c_֤��.���� = Nvl(v_����, '_') Then
            n_Count := 1;
          Else
            Zl_ҽ�ƿ��䶯_Insert_s(14, n_����id, c_֤��.�����id, Null, c_֤��.����, '֤����ȡ����', Null, v_����Ա����, d_�Ǽ�ʱ��);
          End If;
        End Loop;
        --�������˿���Ϣ
        If n_Count = 0 And v_���� Is Not Null Then
          Zl_ҽ�ƿ��䶯_Insert_s(11, n_����id, n_�����id, Null, v_����, '֤������', Null, v_����Ա����, d_�Ǽ�ʱ��);
        End If;
      End If;
    End Loop;
  End If;

  --���¹�������
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('allergic_drugs_list');
  If j_Jsonlist Is Not Null Then
    --������м�¼
    Zl_���˹���ҩ��_Delete(n_����id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json         := Pljson();
      o_Json         := Pljson(j_Jsonlist.Get(I));
      n_����ҩƷid   := o_Json.Get_Number('pat_algc_cadn_id');
      v_����ҩ������ := o_Json.Get_String('pat_algc_cadn');
      v_��ÿҩ�ﷴӦ := o_Json.Get_String('allergy_info');
    
      If v_����ҩ������ Is Not Null Then
        Zl_���˹���ҩ��_Update(n_����id, n_����ҩƷid, v_����ҩ������, v_��ÿҩ�ﷴӦ);
      End If;
    End Loop;
  End If;

  --�������߼�¼
  --    immune_list[]         C     ���������б�
  --      vaccinate_time      C  1  ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --      vaccinate_name      C  1  ��������
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('immune_list');
  If j_Jsonlist Is Not Null Then
    --������м�¼
    Zl_�������߼�¼_Delete(n_����id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json     := Pljson();
      o_Json     := Pljson(j_Jsonlist.Get(I));
      d_����ʱ�� := To_Date(o_Json.Get_String('vaccinate_time'), 'YYYY-MM-DD hh24:mi:ss');
      v_�������� := o_Json.Get_String('vaccinate_name');
    
      If v_�������� Is Not Null Then
        Zl_�������߼�¼_Update(n_����id, d_����ʱ��, v_��������);
      End If;
    End Loop;
  End If;

  --����ҽ�ƿ�����
  --    card_property_list[]  C     ҽ�ƿ������б�
  --      cardtype_id         N  1  ҽ�ƿ����ID
  --      card_no             C  1  ����
  --      info_name           C  1  ��Ϣ��
  --      info_value          N  1  ��Ϣֵ
  j_Jsonlist := Pljson_List();
  j_Jsonlist := j_Input.Get_Pljson_List('card_property_list');
  If j_Jsonlist Is Not Null Then
    --������м�¼
    Zl_�������߼�¼_Delete(n_����id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json     := Pljson();
      o_Json     := Pljson(j_Jsonlist.Get(I));
      n_�����id := o_Json.Get_Number('cardtype_id');
      v_����     := o_Json.Get_String('card_no');
      v_��Ϣ��   := o_Json.Get_String('info_name');
      v_��Ϣֵ   := o_Json.Get_String('info_value');
    
      Zl_����ҽ�ƿ�����_Update(n_����id, n_�����id, v_����, v_��Ϣ��, v_��Ϣֵ);
    End Loop;
  End If;

  --ǩԼ��Ϣ
  --    sign_info             C   ǩԼ��Ϣ
  --      card_type_id        N 1 �����ID
  --      card_no             C 1 ����
  --      card_pwd            C   ������
  --      qrcode              C   ��ά��
  --      card_notes          C   �䶯ԭ��
  --      card_use_endtime    C   ��ֹʹ��ʱ��
  o_Json1 := Pljson();
  o_Json1 := j_Input.Get_Pljson('sign_info');
  If o_Json1 Is Not Null Then
    n_�����id     := o_Json1.Get_Number('card_type_id');
    v_����         := o_Json1.Get_String('card_no');
    v_������       := o_Json1.Get_String('card_pwd');
    v_�䶯ԭ��     := o_Json1.Get_String('card_notes');
    d_��ֹʹ��ʱ�� := To_Date(o_Json1.Get_String('card_use_endtime'), 'YYYY-MM-DD hh24:mi:ss');
    --ǩԼ
    Select Count(1) Into n_Count From ҽ�ƿ���� Where ID = n_�����id;
    If n_Count = 1 Then
      Select Count(1) Into n_Count From ����ҽ�ƿ���Ϣ Where ���� = v_���� And �����id = n_�����id;
      If n_Count = 0 Then
        Zl_ҽ�ƿ��䶯_Insert_s(11, n_����id, n_�����id, '', v_����, v_�䶯ԭ��, v_������, v_����Ա����, d_�Ǽ�ʱ��, Null, d_��ֹʹ��ʱ��);
      End If;
    End If;
  End If;

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Newpatiarchives;
/

Create Or Replace Procedure Zl_Patisvr_Patiagecheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ----------------------------------
  --���ܣ���������������
  --���:Json��ʽ
  --input
  --       pati_age         C 1 ����
  --       pati_birthdate   C 1 ����
  --       calcdate         C 1 ��������
  --����:json��ʽ
  --output
  --       code             N 1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --       message          C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --       error_info       C 1 ������Ϣ
  -----------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_����     Varchar2(20);
  d_�������� Date;
  d_�������� Date;
  v_Info     Varchar2(32767);
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_����     := j_Json.Get_String('pati_age');
  d_�������� := To_Date(j_Json.Get_String('pati_birthdate'), 'YYYY-MM-DD HH24:MI:SS');
  d_�������� := To_Date(j_Json.Get_String('calcdate'), 'YYYY-MM-DD HH24:MI:SS');
  v_Info     := Zl_Age_Check(v_����, d_��������, d_��������);
  Json_Out   := '{"output":{"code":1,"message":"�ɹ�","error_info":"' || Zljsonstr(v_Info, 0) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Patiagecheck;
/
Create Or Replace Procedure Zl_Patisvr_Patiidcardcheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ----------------------------------
  --���ܣ����֤�ż��
  --���:Json��ʽ
  --input
  --    pati_idcard           C 1 ���֤��
  --    calcdate              C 1 ��������
  --����:json��ʽ
  --output
  --    code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    info                  C 1 ����ͨ�����ص���Ϣ
  -----------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_���֤�� Varchar2(20);
  d_�������� Date;
  v_Info     Varchar2(32767);
Begin
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_���֤�� := j_Json.Get_String('pati_idcard');
  d_�������� := To_Date(j_Json.Get_String('calcdate'), 'YYYY-MM-DD HH24:MI:SS');
  v_Info     := Zl_Fun_Checkidcard(v_���֤��, d_��������);
  Json_Out   := '{"output":{"code":1,"message":"�ɹ�","info":"' || Zljsonstr(v_Info) || '"}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Patiidcardcheck;
/
Create Or Replace Procedure Zl_Patisvr_Patirealnamecheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------
  --����:ʵ����֤ǰ�ļ��
  --��� JSOM��ʽ
  --input
  --  opr_fun               N 1  ���� 0-����ʵ����Ϣʱ���  1-�޸�ʵ����Ϣʱ���
  --  real_id               N 1  ʵ��id  opr_fun=1 ʱ����
  --  pati_name             C 1 ����
  --  pati_sex              C 1 �Ա�
  --  pati_age              C 1 ����
  --  pati_birthdate        C 1 ��������
  --  pati_idcard           C 1 ���֤��
  --  owner                 N 1 ������
  --  grdn_name             C 1 ����������
  --  grdn_sex              C 1 �������Ա�
  --  grdn_birthdate        C 1 �����˳�������
  --  grdn_idcard           C 1 ���������֤��
  --  grdn_relation         C 1 �����˹�ϵ
  --  papers_info           C 1 ֤����Ϣƴ��
  --���� JSON��ʽ
  --output
  --  code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  real_id               N 1 ʵ��id
  --  pati_id               N 1 ����id
  --  new_pati              N 1 �Ƿ��²���
  --  pati_age              C 1 ����
  --  pati_name             C 1 ����
  --  pati_sex              C 1 �Ա�
  --  pati_birthdate        C 1 ��������
  --------------------------------------------------------------------------------------------------
  Type r_֤����Ϣ Is Record(
    ֤����Ϣ Varchar2(4000));

  Type t_֤����Ϣ Is Table Of r_֤����Ϣ;
  Rs_Sql֤����Ϣ t_֤����Ϣ := t_֤����Ϣ();

  Type r_֤�� Is Record(
    ʵ��id   ����ʵ��֤��.ʵ��id%Type,
    ֤��id   ����ʵ��֤��.Id%Type,
    ֤������ ����ʵ��֤��.֤������%Type,
    ֤������ ����ʵ��֤��.֤������%Type,
    ֤����ע ����ʵ��֤��.��ע%Type,
    ������   ����ʵ��֤��.������%Type,
    ���     Number(1));

  Type t_֤�� Is Table Of r_֤��;
  Rs_Sql֤�� t_֤�� := t_֤��();

  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_����     ������Ϣ.����%Type;
  v_�Ա�     ������Ϣ.�Ա�%Type;
  v_����     ������Ϣ.����%Type; --����ǰ������
  v_�������� ������Ϣ.��������%Type;

  v_���֤��       ����ʵ����Ϣ.���֤��%Type;
  n_������         ����ʵ��֤��.������%Type;
  v_����������     ����ʵ����Ϣ.����������%Type;
  v_���������֤�� ����ʵ����Ϣ.���������֤��%Type;
  v_�������Ա�     ����ʵ����Ϣ.�������Ա�%Type;
  v_�����˳������� ����ʵ����Ϣ.�����˳�������%Type;
  v_�����˹�ϵ     ����ʵ����Ϣ.�����˹�ϵ%Type;

  n_New    Number(1);
  n_Id     Number(18);
  n_����id Number(18);
  n_ʵ��id Number(18);
  n_Count  Number(5);

  v_֤��������  Varchar2(200);
  v_֤������    Varchar2(400);
  v_֤����Ϣ_In Varchar2(4000);
  v_֤����Ϣ    Varchar2(4000);
  n_���        Number;
  n_Realid      Number;

  t_Key   t_Strlist;
  v_Error Varchar2(200);
Begin
  j_Jsonin         := Pljson(Json_In);
  j_Json           := j_Jsonin.Get_Pljson('input');
  n_���           := j_Json.Get_Number('opr_fun');
  n_Realid         := j_Json.Get_Number('real_id');
  v_����           := j_Json.Get_String('pati_name');
  v_�Ա�           := j_Json.Get_String('pati_sex');
  v_����           := j_Json.Get_String('pati_age');
  v_��������       := To_Date(j_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
  v_���֤��       := j_Json.Get_String('pati_idcard');
  n_������         := j_Json.Get_Number('owner');
  v_����������     := j_Json.Get_String('grdn_name');
  v_���������֤�� := j_Json.Get_String('grdn_idcard');
  v_�������Ա�     := j_Json.Get_String('grdn_sex');
  v_�����˳������� := To_Date(j_Json.Get_String('grdn_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
  v_�����˹�ϵ     := j_Json.Get_String('grdn_relation');
  v_֤����Ϣ_In    := j_Json.Get_String('papers_info');
  --����¼��Ϣ,����¼�벡�˵��������Ա𡢳�������
  If v_���� Is Null Then
    v_Error := '����¼�벡��������';
  Elsif v_�Ա� Is Null Then
    v_Error := '����¼�벡���Ա�';
  Elsif v_�������� Is Null Then
    v_Error := '����¼�벡�˳������ڣ�';
  End If;
  If v_Error Is Not Null Then
    Json_Out := Zljsonout(v_Error, 0);
    Return;
  End If;
  --��ȡ֤����Ϣ
  For X In (Select Column_Value As ֤����Ϣ From Table(f_Str2list(v_֤����Ϣ_In, ','))) Loop
    Rs_Sql֤����Ϣ.Extend;
    Rs_Sql֤����Ϣ(Rs_Sql֤����Ϣ.Count).֤����Ϣ := x.֤����Ϣ;
  End Loop;

  For I In 1 .. Rs_Sql֤����Ϣ.Count Loop
    Rs_Sql֤��.Extend;
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Rs_Sql֤����Ϣ(I).֤����Ϣ, '-'));
    Rs_Sql֤��(Rs_Sql֤��.Count).ʵ��id := t_Key(1);
  
    Rs_Sql֤��(Rs_Sql֤��.Count).֤��id := t_Key(2);
  
    Rs_Sql֤��(Rs_Sql֤��.Count).֤������ := t_Key(3);
  
    Rs_Sql֤��(Rs_Sql֤��.Count).֤������ := t_Key(4);
  
    Rs_Sql֤��(Rs_Sql֤��.Count).֤����ע := t_Key(5);
  
    Rs_Sql֤��(Rs_Sql֤��.Count).������ := To_Number(t_Key(6));
  
    Rs_Sql֤��(Rs_Sql֤��.Count).��� := I;
  End Loop;

  For N In 1 .. Rs_Sql֤��.Count Loop
    v_֤�������� := v_֤�������� || Rs_Sql֤��(N).������;
    v_֤������   := v_֤������ || Rs_Sql֤��(N).֤������;
    v_֤����Ϣ   := v_֤����Ϣ || ',' || Rs_Sql֤��(N).֤������ || ',' || Rs_Sql֤��(N).֤������ || ',' || Rs_Sql֤��(N).������;
  End Loop;
  If (v_���������� Is Null And v_���������֤�� Is Null) And (v_���������� Is Null And v_֤������ Is Null) Then
    --��û��������������Ϣ������£��������֤�ź�����֤���ű���¼��һ��
    If v_���֤�� Is Null And v_֤������ Is Null Then
      v_Error := '����¼�벡�����֤�Ż�������֤�����룡';
    End If;
  Else
    --¼������������Ϣ�����¼��v_�������Ա������˳������ڡ������˹�ϵ
    If (Not v_���������� Is Null And Not v_���������֤�� Is Null) Or (n_������ = 2 And Not v_���������� Is Null And Not v_֤������ Is Null) Then
      If v_�������Ա� Is Null Then
        v_Error := '����¼���������Ա�';
      Elsif v_�����˳������� Is Null Then
        v_Error := '����¼�������˳������ڣ�';
      Elsif v_�����˹�ϵ Is Null Then
        v_Error := '����¼�������˹�ϵ��';
      End If;
    End If;
  End If;
  If v_Error Is Not Null Then
    Json_Out := Zljsonout(v_Error, 0);
    Return;
  End If;
  If Nvl(n_���, 0) = 0 Then
    --��ѯ�Ƿ����ظ��Ĳ���ʵ����Ϣ
    --��һ�����������+���֤��
    If Nvl(v_���֤��, '|') <> '|' Then
      Select Count(1) Into n_Count From ����ʵ����Ϣ Where ���֤�� = v_���֤��;
      If n_Count > 0 Then
        v_Error := '�Ѿ��������֤��Ϊ��' || v_���֤�� || '����ʵ����֤��Ϣ,���飡';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --�ڶ������������+����֤������+����֤������(���˵�)
    If n_������ = 1 Then
      Select Count(1)
      Into n_Count
      From ����ʵ����Ϣ A, ����ʵ��֤�� B
      Where a.ʵ��id = b.ʵ��id And a.���� = v_���� And
            Instr(v_֤����Ϣ || ',', ',' || b.֤������ || ',' || b.֤������ || ',' || b.������ || ',') > 0;
      If n_Count > 0 Then
        v_Error := '�ò����Ѿ�������Ч��ʵ����֤��Ϣ������Ҫ�ٴ���֤��';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --���������������+����������+���������֤��
    Select Count(1)
    Into n_Count
    From ����ʵ����Ϣ
    Where ���� = v_���� And ���������� = v_���������� And ���������֤�� = v_���������֤��;
    If n_Count > 0 Then
      v_Error := '�ò����Ѿ�������Ч��ʵ����֤��Ϣ������Ҫ�ٴ���֤��';
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --���������������+����������+����֤������+����֤������(�����˵�)
    If n_������ = 2 Then
      Select Count(1)
      Into n_Count
      From ����ʵ����Ϣ A, ����ʵ��֤�� B
      Where a.ʵ��id = b.ʵ��id And a.���� = v_���� And a.���������� = v_���������� And
            Instr(v_֤����Ϣ || ',', ',' || b.֤������ || ',' || b.֤������ || ',' || b.������ || ',') > 0;
      If n_Count > 0 Then
        v_Error := '�ò����Ѿ�������Ч��ʵ����֤��Ϣ������Ҫ�ٴ���֤��';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    Select ����ʵ����Ϣ_ʵ��id.Nextval Into n_ʵ��id From Dual;
    --�½�ָ�����˵�ʵ����֤��Ϣ
    If v_���֤�� Is Null Then
      n_New := 1;
    Else
      Select Max(����id) As ����id
      Into n_Id
      From (Select Nvl(Nvl(����ʱ��, ��Ժʱ��), �Ǽ�ʱ��) As ʱ��, ����id
             From ������Ϣ
             Where ���� = v_���� And ���֤�� = v_���֤��
             Order By ʱ�� Desc)
      Where Rownum = 1;
      If n_Id Is Null Then
        n_New := 1;
      Else
        n_New := 0;
      End If;
    End If;
  
    If n_New = 1 Then
      Select ������Ϣ_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := n_Id;
    End If;
  Else
  
    --��ѯ�Ƿ����ظ��Ĳ���ʵ����Ϣ
    --��һ����������֤��
    If Nvl(v_���֤��, '|') <> '|' Then
      Select Count(1)
      Into n_Count
      From ����ʵ����Ϣ
      Where ���� = v_���� And ���֤�� = v_���֤�� And ʵ��id <> n_Realid;
      If n_Count > 0 Then
        v_Error := '�Ѿ��������֤��Ϊ��' || v_���֤�� || '����ʵ����֤��Ϣ,���飡';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --�ڶ������������+����֤������+����֤������(���˵�)
    If n_������ = 1 Then
      Select Count(1)
      Into n_Count
      From ����ʵ����Ϣ A, ����ʵ��֤�� B
      Where a.ʵ��id = b.ʵ��id And a.���� = v_���� And a.ʵ��id <> n_Realid And
            Instr(v_֤����Ϣ || ',', ',' || b.֤������ || ',' || b.֤������ || ',' || b.������ || ',') > 0;
      If n_Count > 0 Then
        v_Error := '�ò����Ѿ�������Ч��ʵ����֤��Ϣ������Ҫ�ٴ���֤��';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --���������������+����������+���������֤��
    Select Count(1)
    Into n_Count
    From ����ʵ����Ϣ
    Where ���� = v_���� And ���������� = v_���������� And ���������֤�� = v_���������֤�� And ʵ��id <> n_Realid;
    If n_Count > 0 Then
      v_Error := '�ò����Ѿ�������Ч��ʵ����֤��Ϣ������Ҫ�ٴ���֤��';
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
    n_Count := 0;
    --���������������+����������+����֤������+����֤������(�����˵�)
    If n_������ = 2 Then
      Select Count(1)
      Into n_Count
      From ����ʵ����Ϣ A, ����ʵ��֤�� B
      Where a.ʵ��id = b.ʵ��id And a.���� = v_���� And a.���������� = v_���������� And a.ʵ��id <> n_Realid And
            Instr(v_֤����Ϣ || ',', ',' || b.֤������ || ',' || b.֤������ || ',' || b.������ || ',') > 0;
      If n_Count > 0 Then
        v_Error := '�ò����Ѿ�������Ч��ʵ����֤��Ϣ������Ҫ�ٴ���֤��';
      End If;
    End If;
    If v_Error Is Not Null Then
      Json_Out := Zljsonout(v_Error, 0);
      Return;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"';
  If Nvl(n_���, 0) = 0 Then
    Json_Out := Json_Out || ',"real_id":' || n_ʵ��id || ',"pati_id":' || n_����id || ',"new_pati":' || n_New || '}}';
  Else
    Json_Out := Json_Out || '}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Patirealnamecheck;
/
Create Or Replace Procedure Zl_Patisvr_Phonenumberexist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ�����ֻ����Ƿ��Ѿ���ʹ��
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id              N   1  ����ID:��ǰ�����Ĳ���
  --    phone_number         C   1  �ֻ���
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exist                N   1   1-����;0-������
  ---------------------------------------------------------------------------
  n_����id ������Ϣ.����id%Type;
  v_�ֻ��� ������Ϣ.�ֻ���%Type;
  n_Exist  Number(1);
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin

  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');
  v_�ֻ��� := j_Json.Get_String('phone_number');

  Select Count(1) Into n_Exist From ������Ϣ Where �ֻ��� = v_�ֻ��� And ����id <> n_����id And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exist":' || Nvl(n_Exist, 0) || '}}';
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Phonenumberexist;
/
Create Or Replace Procedure Zl_Patisvr_Recalcage
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����㲡������
  --��Σ�Json_In:��ʽ
  --input
  --    pati_ids                   C   1  ����IDs,����ö��ŷ���(���˻���ͬʱ����������������)
  --����: Json_Out,��ʽ����
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  v_Pati_Ids Varchar2(2000);
  v_Age      Varchar2(20);
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  v_Pati_Ids := j_Json.Get_String('pati_ids');
  If Nvl(v_Pati_Ids, '_') = '_' Then
    Json_Out := Zljsonout('δ���벡��Id,���飡');
    Return;
  End If;
  For R In (Select /*+cardinality(a,10)*/
             Column_Value As ����id
            From Table(f_Num2list(v_Pati_Ids)) A) Loop
  
    v_Age := Zl_Age_Calc(r.����id);
    Update ������Ϣ Set ���� = v_Age Where ����id = r.����id;
  End Loop;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Patisvr_Recalcage;
/
CREATE OR REPLACE Procedure Zl_Patisvr_Savebadrecord
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���˲�����¼���ݱ���
  --��Σ�Json_In:��ʽ
  --    input
  --      badrec_list[]          C   ����Ĳ�����¼�б�
  --        pati_id            N 1 ����id
  --        behavior_category  C 1 ��Ϊ���:��ԤԼ�Һ�
  --        happen_time        C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --        add_time           C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --        add_Reason         C 1 ����ԭ����ԤԼ����
  --        add_memo           C 1 ����˵��
  --        additional_info    C 1 ������Ϣ
  --        creator            C 1 �Ǽ���
  --����: Json_Out,��ʽ����
  --  output
  --    code                   N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  o_Json     Pljson;
  j_Jsonin   Pljson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_����id     ���˲�����¼.����id%Type;
  v_��Ϊ���   ���˲�����¼.��Ϊ���%Type;
  d_����ʱ��   ���˲�����¼.����ʱ��%Type;
  d_����ʱ��   ���˲�����¼.����ʱ��%Type;
  v_����ԭ��   ���˲�����¼.����ԭ��%Type;
  v_����˵��   ���˲�����¼.����˵��%Type;
  v_������Ϣ   ���˲�����¼.������Ϣ%Type;
  v_�Ǽ���     ���˲�����¼.�Ǽ���%Type;
  v_����Ա���� ���˲�����¼.�Ǽ���%Type;

Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  j_Jsonlist := j_Json.Get_Pljson_List('badrec_list');

  If j_Jsonlist Is Null Then
    Json_Out := Zljsonout('δ������Ҫ���治����¼���ݣ����ܱ���');
    Return;
  End If;

  For I In 1 .. j_Jsonlist.Count Loop
    o_Json     := Pljson();
    o_Json     := Pljson(j_Jsonlist.Get(I));
    n_����id   := o_Json.Get_Number('pati_id');
    v_��Ϊ��� := o_Json.Get_String('behavior_category');
    d_����ʱ�� := To_Date(o_Json.Get_String('add_time'), 'yyyy-mm-dd hh24:mi:ss');
    d_����ʱ�� := To_Date(o_Json.Get_String('happen_time'), 'yyyy-mm-dd hh24:mi:ss');
    v_����ԭ�� := o_Json.Get_String('add_reason');
    v_����˵�� := o_Json.Get_String('add_memo');
    v_������Ϣ := o_Json.Get_String('additional_info');
    v_�Ǽ���   := o_Json.Get_String('creator');

    If Nvl(n_����id, 0) = 0 Then
      Json_Out := Zljsonout('����ȷ��������Ϣ������!');
      Return;
    End If;
    If v_�Ǽ��� Is Null And v_����Ա���� Is Null Then
      v_����Ա���� := zl_UserName;
    End If;

    Insert Into ���˲�����¼
      (ID, ��Ϊ���, ����id, ����ʱ��, ����ʱ��, ����ԭ��, ����˵��, ������Ϣ, �Ǽ���)
      Select ���˲�����¼_Id.Nextval, v_��Ϊ���, n_����id, d_����ʱ��, Nvl(d_����ʱ��, Sysdate), v_����ԭ��, v_����˵��, v_������Ϣ,
             Nvl(v_�Ǽ���, v_����Ա����)
      From Dual;

  End Loop;

  Json_Out := Zljsonout('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Savebadrecord;
/

Create Or Replace Procedure Zl_Patisvr_Savemedccard
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --���ܣ��Բ��˵�ҽ�ƿ����š��󶨿�����������ز�������ҽ�ƿ��䶯���������ݽ��б���
  --��Σ�json��ʽ
  --input
  --   oper_state            N  1  ����״̬::0��NULL������¼;1-�����쳣����;2-ֻ�����䶯��¼
  --   oper_type             N  1 ��������:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����);5-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ),7-��ֹʱ�����
  --   change_id             N  1  �䶯ID
  --   pati_id               N  1 ����id
  --   card_type_id          N  1 �����ID
  --   card_no_old           C  1 ԭ����
  --   card_no               C  1 ҽ�ƿ���
  --   card_notes            C  1 �䶯ԭ��
  --   card_pwd              C  1 ����
  --   iccard_no             C  1 IC����
  --   loss_mode             C  1 ��ʧ��ʽ
  --   qrcode                C  1 ��ά��
  --   card_use_endtime      C  1 ��ֹʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
  --   operator_time         C  1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --   operator_name         C  1 ����Ա����
  --   card_price            N  1 ����
  --   fee_no                C  1 ���õ���

  --���Σ�json��ʽ
  --Json_Out
  --   code                  N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C  1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------
  n_����״̬     Number(2);
  n_��������     Number(2);
  n_����id       ����ҽ�ƿ���Ϣ.����id%Type;
  n_�����id     ����ҽ�ƿ���Ϣ.�����id%Type;
  v_ԭ����       ����ҽ�ƿ���Ϣ.����%Type;
  v_ҽ�ƿ���     ����ҽ�ƿ���Ϣ.����%Type;
  v_�䶯ԭ��     ����ҽ�ƿ��䶯.�䶯ԭ��%Type;
  v_����         ������Ϣ.����֤��%Type;
  v_����Ա����   ����ҽ�ƿ��䶯.����Ա����%Type;
  d_����ʱ��     Date;
  v_Ic����       ������Ϣ.Ic����%Type := Null;
  v_��ʧ��ʽ     ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null;
  d_��ֹʹ��ʱ�� Date;
  v_��ά��       ����ҽ�ƿ���Ϣ.��ά��%Type;
  n_����         ����ҽ�ƿ��䶯.����%Type;
  v_���õ�       ����ҽ�ƿ��䶯.���õ���%Type;
  n_�䶯id       ����ҽ�ƿ��䶯.Id%Type;

  j_Json   Pljson;
  j_Jsonin Pljson;
Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_�������� := j_Json.Get_Number('oper_type');
  n_����id   := j_Json.Get_Number('pati_id');
  n_�����id := j_Json.Get_Number('card_type_id');
  If Nvl(n_����id, 0) = 0 Or Nvl(n_�����id, 0) = 0 Then
    Json_Out := Zljsonout('ʧ�ܣ�δ���벡��ID�����id');
    Return;
  End If;

  n_����״̬     := Nvl(j_Json.Get_Number('oper_state'), 0);
  v_ԭ����       := j_Json.Get_String('card_no_old');
  v_ҽ�ƿ���     := j_Json.Get_String('card_no');
  v_�䶯ԭ��     := j_Json.Get_String('card_notes');
  v_����         := j_Json.Get_String('card_pwd');
  v_Ic����       := j_Json.Get_String('iccard_no');
  v_��ʧ��ʽ     := j_Json.Get_String('loss_mode');
  v_��ά��       := j_Json.Get_String('qrcode');
  d_��ֹʹ��ʱ�� := To_Date(j_Json.Get_String('card_use_endtime'), 'yyyy-mm-dd hh24:mi:ss');
  d_����ʱ��     := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_����Ա����   := j_Json.Get_String('operator_name');
  n_����         := j_Json.Get_Number('card_price');
  v_���õ�       := j_Json.Get_String('fee_no');
  n_�䶯id       := j_Json.Get_Number('change_id');

  Zl_ҽ�ƿ��䶯_Insert_s(n_��������, n_����id, n_�����id, v_ԭ����, v_ҽ�ƿ���, v_�䶯ԭ��, v_����, v_����Ա����, d_����ʱ��, v_��ʧ��ʽ, d_��ֹʹ��ʱ��, n_����, Null,
                    v_���õ�, Null, n_�䶯id, n_����״̬);

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Savemedccard;
/
Create Or Replace Procedure Zl_Patisvr_Savepatiphoto
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ָ��������Ƭ
  --��Σ�Json_In:��ʽ
  --   input
  --      pati_id           N 1 ����ID
  --      pati_photo        C 1 ����:base64

  --����      json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_����id   ������Ƭ.����id%Type;
  b_������Ƭ ������Ƭ.��Ƭ%Type;
  c_������Ƭ Clob;
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ���벡��id�����ܱ��没����Ƭ!"}}';
    Return;
  End If;

  c_������Ƭ := j_Json.Get_Clob('pati_photo');
  b_������Ƭ := Zltools.Zlbase64.Decode(c_������Ƭ);

  Update ������Ƭ Set ��Ƭ = b_������Ƭ Where ����id = n_����id;
  If Sql%RowCount = 0 Then
    Insert Into ������Ƭ (����id, ��Ƭ) Values (n_����id, b_������Ƭ);
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Savepatiphoto;
/
Create Or Replace Procedure Zl_Patisvr_Updatecardtype
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�޸Ļ�����ҽ�ƿ����
  --��Σ�Json_In:��ʽ
  --    input
  --      cardtype_id         N  1  ID
  --      cardtype_code       C  1  ����
  --      cardtype_name       C  1  ����
  --      cardtype_stname     C  1  ����
  --      prefix_text         C  1  ǰ׺�ı�
  --      cardno_len          N  1  ���ų���
  --      default             N  1  ȱʡ��־
  --      fixed               N  1  �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
  --      strict              N  1  �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
  --      self_make           N  1  �Ƿ�����:1-�ǵ�;0-����
  --      exsit_account       N  1  �Ƿ�����ʻ�:1-�����ʻ�;0-�������˻�
  --      allow_return_cash   N  1  �Ƿ�����:1-����;0-������
  --      must_all_return     N  1  �Ƿ�ȫ��:1-����ȫ��;0-��������
  --      component           C  1  ����
  --      memo                C  1  ��ע
  --      spec_item           C  1  �ض���Ŀ
  --      blnc_mode           C  1  ���㷽ʽ
  --      cardno_pwdtxt       C  1  ��������:���Ŵӵڼ�λ���ڼ�λ��ʾ����,��ʽΪ:S-N:S��ʾ�ӵڼ�λ��ʼ,���ڼ�λ����.����:3-10;��ʾ��3λ��10λ������*��ʾ:12********3323��Ҫ����Ӧ��ͬ����ҽ�ƿ�
  --      allow_repeat_use    N  1  �Ƿ��ظ�ʹ��:1-����;0-������
  --      enabled             N  1  �Ƿ�����:1-������;0-δ����
  --      pwd_len             N  1  ���볤��
  --      pwd_len_limit       N  1  ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
  --      pwd_rule            N  1  �������:��-���ֺ��ַ����;1-��Ϊ�������
  --      allow_vaguefind     N  1  �Ƿ�ģ������:1-֧��ģ������;0-��֧��
  --      pwd_require         N  1  ������������:0-������;1-������,����;2-�������ֹ;ȱʡΪ������
  --      default_pwd         N  1  �Ƿ�ȱʡ����:1-�����֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
  --      allow_makecard      N  1  �Ƿ��ƿ�:1-��;0-��
  --      allow_sendcard      N  1  �Ƿ񷢿�:1-��;0-��
  --      allow_writcard      N  1  �Ƿ�д��:1-��;0-��
  --      insurance_type      N  1  ����
  --      sendcard_nature     N  1  ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0
  --      allow_transfer      N  1  �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
  --      readcard_nature     C  1  ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
  --      keyboard_mode       N  1  ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
  --      advsend_buildqrcode N  1  �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
  --      holding_pay         N  1  �Ƿ�ֿ�����:1-��;0-��
  --      cert_cardtype       N  1  �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
  --      verfycard           N  1  �Ƿ��˿��鿨
  --      sendcard_sign       N  1  ��������:0��NULL-����ʱ�����ű���ﵽ���ų���;1-����ʱ��������С�ڵ��ڿ��ų���,����ʱ��С�ڿ��ų���ʱ������ʾ����Ա;2-����ʱ��������С�ڵ��ڿ��ų���,С��ʱ����ʾ����Ա��
  --      enterkey_enabled    N  1  �豸�Ƿ����ûس�:ҽ�ƿ���Ӧ��ˢ���豸�Ƿ������˻س�����������˻س����򿨺ų���Ĭ������һλ�����λس�


  --      def_return_cash     N  1  �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
  --      balalone            N  1  �Ƿ��������:1-��������;0-�Ƕ�������
  --      discern_rule        N  1  ����ʶ�����:1-ȫ��ת��Ϊ��д;0-�����ִ�Сд
  --      def_valid_time      C  1  ȱʡ��Чʱ��:NULLʱ����ʾ������;�ǿ�ʱ����ʽΪ:ʱ��+��λ(�죬��),���磺3��,3��
  --      scanpay             N  1  �Ƿ�֧��ɨ�븶:�Ƿ�֧��ɨ�븶,֧��ʱ������á�zlReadQRCode������
  --����: Json_Out,��ʽ����
  --   output
  --      code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --      message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json             Pljson;
  j_Jsonin           Pljson;
  n_Id               ҽ�ƿ����.Id%Type;
  v_����             ҽ�ƿ����.����%Type;
  v_����             ҽ�ƿ����.����%Type;
  v_����             ҽ�ƿ����.����%Type;
  v_ǰ׺�ı�         ҽ�ƿ����.ǰ׺�ı�%Type;
  n_���ų���         ҽ�ƿ����.���ų���%Type;
  n_ȱʡ��־         ҽ�ƿ����.ȱʡ��־%Type;
  n_�Ƿ�̶�         ҽ�ƿ����.�Ƿ�̶�%Type;
  n_�Ƿ��ϸ����     ҽ�ƿ����.�Ƿ��ϸ����%Type;
  n_�Ƿ�����         ҽ�ƿ����.�Ƿ�����%Type;
  n_�Ƿ�����ʻ�     ҽ�ƿ����.�Ƿ�����ʻ�%Type;
  n_�Ƿ�ȫ��         ҽ�ƿ����.�Ƿ�ȫ��%Type;
  v_����             ҽ�ƿ����.����%Type;
  v_��ע             ҽ�ƿ����.��ע%Type;
  v_�ض���Ŀ         ҽ�ƿ����.�ض���Ŀ%Type;
  v_���㷽ʽ         ҽ�ƿ����.���㷽ʽ%Type;
  n_�Ƿ�����         ҽ�ƿ����.�Ƿ�����%Type;
  v_��������         ҽ�ƿ����.��������%Type;
  n_�Ƿ��ظ�ʹ��     ҽ�ƿ����.�Ƿ��ظ�ʹ��%Type;
  n_���볤��         ҽ�ƿ����.���볤��%Type;
  n_���볤������     ҽ�ƿ����.���볤������%Type;
  n_�������         ҽ�ƿ����.�������%Type;
  n_�Ƿ�����         ҽ�ƿ����.�Ƿ�����%Type;
  n_������ʽ         Integer := 0;
  n_�Ƿ�ģ������     ҽ�ƿ����.�Ƿ�ģ������%Type := 0;
  n_������������     ҽ�ƿ����.������������%Type := 0;
  n_�Ƿ�ȱʡ����     ҽ�ƿ����.�Ƿ�ȱʡ����%Type := 0;
  n_�Ƿ��ƿ�         ҽ�ƿ����.�Ƿ��ƿ�%Type := 0;
  n_�Ƿ񷢿�         ҽ�ƿ����.�Ƿ񷢿�%Type := 0;
  n_�Ƿ�д��         ҽ�ƿ����.�Ƿ�д��%Type := 0;
  n_����             ҽ�ƿ����.����%Type := 0;
  n_��������         ҽ�ƿ����.��������%Type := 0;
  n_�Ƿ�ת�ʼ�����   ҽ�ƿ����.�Ƿ�ת�ʼ�����%Type := 0;
  v_��������         ҽ�ƿ����.��������%Type := '1000';
  n_���̿��Ʒ�ʽ     ҽ�ƿ����.���̿��Ʒ�ʽ%Type := 0;
  n_�Ƿ�֤��         ҽ�ƿ����.�Ƿ�֤��%Type := 0;
  n_�Ƿ�ֿ�����     ҽ�ƿ����.�Ƿ�ֿ�����%Type := 0;
  n_���͵��ýӿ�     ҽ�ƿ����.���͵��ýӿ�%Type := 0;
  n_�Ƿ��˿��鿨     ҽ�ƿ����.�Ƿ��˿��鿨%Type := 0;
  n_�豸�Ƿ����ûس� ҽ�ƿ����.�豸�Ƿ����ûس�%Type := 0;
  n_�������ſ���     ҽ�ƿ����.��������%Type := 0;
  n_�Ƿ�ȱʡ����     ҽ�ƿ����.�Ƿ�ȱʡ����%Type := 0;
  n_�Ƿ��������     ҽ�ƿ����.�Ƿ��������%Type := 0;
  d_ȱʡ��Чʱ��     ҽ�ƿ����.ȱʡ��Чʱ��%Type := Null;
  n_����ʶ�����     ҽ�ƿ����.����ʶ�����%Type := 0;
  n_�Ƿ�֧��ɨ�븶   ҽ�ƿ����.�Ƿ�֧��ɨ�븶%Type := 0;

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  --      cardtype_id         N  1 ID
  --      cardtype_code       C  1 ����
  --      cardtype_name       C  1 ����
  --      cardtype_stname     C  1 ����
  --      prefix_text         C  1 ǰ׺�ı�
  n_Id       := j_Json.Get_Number('cardtype_id');
  v_����     := j_Json.Get_String('cardtype_code');
  v_����     := j_Json.Get_String('cardtype_name');
  v_����     := j_Json.Get_String('cardtype_stname');
  v_ǰ׺�ı� := j_Json.Get_String('prefix_text');
  --      cardno_len          N  1 ���ų���
  --      default             N  1 ȱʡ��־
  --      fixed               N  1 �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
  --      strict              N  1 �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
  --      self_make           N  1 �Ƿ�����:1-�ǵ�;0-����
  n_���ų���     := j_Json.Get_Number('cardno_len');
  n_ȱʡ��־     := j_Json.Get_Number('default');
  n_�Ƿ�̶�     := j_Json.Get_Number('fixed');
  n_�Ƿ��ϸ���� := j_Json.Get_Number('strict');
  n_�Ƿ�����     := j_Json.Get_Number('self_make');
  --      exsit_account          N  1  �Ƿ�����ʻ�:1-�����ʻ�;0-�������˻�
  --      allow_return_cash      N  1  �Ƿ�����:1-����;0-������
  --      must_all_return        N  1  �Ƿ�ȫ��:1-����ȫ��;0-��������
  --      component              C  1  ����
  --      memo                   C  1  ��ע
  n_�Ƿ�����ʻ� := j_Json.Get_Number('exsit_account');
  n_�Ƿ�����     := j_Json.Get_Number('allow_return_cash');
  n_�Ƿ�ȫ��     := j_Json.Get_Number('must_all_return');
  v_����         := j_Json.Get_String('component');
  v_��ע         := j_Json.Get_String('memo');
  --      spec_item           C  1  �ض���Ŀ
  --      blnc_mode           C  1  ���㷽ʽ
  --      cardno_pwdtxt       C  1  ��������:���Ŵӵڼ�λ���ڼ�λ��ʾ����,��ʽΪ:S-N:S��ʾ�ӵڼ�λ��ʼ,���ڼ�λ����.����:3-10;��ʾ��3λ��10λ������*��ʾ:12********3323��Ҫ����Ӧ��ͬ����ҽ�ƿ�
  --      allow_repeat_use    N  1  �Ƿ��ظ�ʹ��:1-����;0-������
  --      enabled             N  1  �Ƿ�����:1-������;0-δ����
  v_�ض���Ŀ     := j_Json.Get_String('spec_item');
  v_���㷽ʽ     := j_Json.Get_String('blnc_mode');
  v_��������     := j_Json.Get_String('cardno_pwdtxt');
  n_�Ƿ�����     := j_Json.Get_Number('enabled');
  n_�Ƿ��ظ�ʹ�� := j_Json.Get_Number('allow_repeat_use');
  --      pwd_len             N  1  ���볤��
  --      pwd_len_limit       N  1  ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
  --      pwd_rule            N  1  �������:��-���ֺ��ַ����;1-��Ϊ�������
  --      allow_vaguefind     N  1  �Ƿ�ģ������:1-֧��ģ������;0-��֧��
  --      pwd_require         N  1  ������������:0-������;1-������,����;2-�������ֹ;ȱʡΪ������

  n_���볤��     := j_Json.Get_Number('pwd_len');
  n_���볤������ := j_Json.Get_Number('pwd_len_limit');
  n_�������     := j_Json.Get_Number('pwd_rule');
  n_�Ƿ�ģ������ := j_Json.Get_Number('allow_vaguefind');
  n_������������ := j_Json.Get_Number('pwd_require');
  --      default_pwd            N  1  �Ƿ�ȱʡ����:1-�����֤��N(�����볤��Ϊ׼)λ��Ϊȱʡ����;0-��ȱʡ����
  --      allow_makecard         N  1  �Ƿ��ƿ�:1-��;0-��
  --      allow_sendcard         N  1  �Ƿ񷢿�:1-��;0-��
  --      allow_writcard         N  1  �Ƿ�д��:1-��;0-��
  --      insurance_type         N  1  ����
  n_�Ƿ�ȱʡ���� := j_Json.Get_Number('default_pwd');
  n_�Ƿ��ƿ�     := j_Json.Get_Number('allow_makecard');
  n_�Ƿ񷢿�     := j_Json.Get_Number('allow_sendcard');
  n_�Ƿ�д��     := j_Json.Get_Number('allow_writcard');
  n_����         := j_Json.Get_Number('insurance_type');
  --      sendcard_nature     N  1  ��������:0-������;1-ͬһ����ֻ�ܷ�һ�ſ�;2-ͬһ�����������ſ���������ʾ;ȱʡΪ0
  --      allow_transfer      N  1  �Ƿ�ת�ʼ�����:1-֧��ת�ʼ�����;0-��֧��
  --      readcard_nature     C  1  ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
  --      keyboard_mode       N  1  ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
  --      advsend_buildqrcode N  1  �Ƿ�ҽ�����͵����������ɽӿ�:1-���͵������ɶ�ά��ӿ�;0-������
  --      holding_pay         N  1  �Ƿ�ֿ�����:1-��;0-��
  --      cert_cardtype       N  1  �Ƿ�֤�����͵�ҽ�ƿ�:0-���ǣ�1-��
  --      verfycard           N  1  �Ƿ��˿��鿨
  n_��������         := j_Json.Get_Number('sendcard_nature');
  n_�Ƿ�ת�ʼ�����   := j_Json.Get_Number('allow_transfer');
  v_��������         := j_Json.Get_String('readcard_nature');
  n_���̿��Ʒ�ʽ     := j_Json.Get_Number('keyboard_mode');
  n_���͵��ýӿ�     := j_Json.Get_Number('advsend_buildqrcode');
  n_�Ƿ�ֿ�����     := j_Json.Get_Number('holding_pay');
  n_�Ƿ�֤��         := j_Json.Get_Number('cert_cardtype');
  n_�Ƿ��˿��鿨     := j_Json.Get_Number('verfycard');
  n_�豸�Ƿ����ûس� := j_Json.Get_Number('enterkey_enabled');
  n_�������ſ���     := j_Json.Get_Number('sendcard_sign');
  --      def_return_cash         N 1 �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
  --      balalone                N 1 �Ƿ��������:1-��������;0-�Ƕ�������
  --      discern_rule            N 1 ����ʶ�����:1-ȫ��ת��Ϊ��д;0-�����ִ�Сд
  --      def_valid_time          C 1 ȱʡ��Чʱ��:NULLʱ����ʾ������;�ǿ�ʱ����ʽΪ:ʱ��+��λ(�죬��),���磺3��,3��
  --      scanpay                 N 1 �Ƿ�֧��ɨ�븶:�Ƿ�֧��ɨ�븶,֧��ʱ������á�zlReadQRCode������
  n_�Ƿ�ȱʡ����   := j_Json.Get_Number('def_return_cash');
  n_�Ƿ��������   := j_Json.Get_Number('balalone');
  d_ȱʡ��Чʱ��   := j_Json.Get_Number('def_valid_time');
  n_����ʶ�����   := j_Json.Get_String('discern_rule');
  n_�Ƿ�֧��ɨ�븶 := j_Json.Get_Number('scanpay');

  Zl_ҽ�ƿ����_Update(n_Id, v_����, v_����, v_����, v_ǰ׺�ı�, n_���ų���, n_ȱʡ��־, n_�Ƿ�̶�, n_�Ƿ��ϸ����, n_�Ƿ�����, n_�Ƿ�����ʻ�, n_�Ƿ�ȫ��, v_����,
                  v_��ע, v_�ض���Ŀ, Null, v_���㷽ʽ, n_�Ƿ�����, v_��������, n_�Ƿ��ظ�ʹ��, n_���볤��, n_���볤������, n_�������, n_�Ƿ�����, n_������ʽ,
                  n_�Ƿ�ģ������, n_������������, n_�Ƿ�ȱʡ����, n_�Ƿ��ƿ�, n_�Ƿ񷢿�, n_�Ƿ�д��, n_����, n_��������, n_�Ƿ�ת�ʼ�����, v_��������, n_���̿��Ʒ�ʽ,
                  n_�Ƿ�֤��, n_�Ƿ�ֿ�����, n_���͵��ýӿ�, n_�Ƿ��˿��鿨, n_�豸�Ƿ����ûس�, n_�������ſ���, n_�Ƿ�ȱʡ����, n_�Ƿ��������, d_ȱʡ��Чʱ��, n_����ʶ�����,
                  n_�Ƿ�֧��ɨ�븶);

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatecardtype;
/
Create Or Replace Procedure Zl_Patisvr_Updateinpatistate
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����סԺ���˾���״̬
  --��Σ�Json_In:��ʽ
  --input
  --    pati_list[]              ����
  --      pati_id              N 1   ����id
  --      pati_pageid          N 1   ��ҳid
  --      outpatient_num       C 1   �����
  --      inpatient_num        C 1   סԺ��
  --      in_time              C 1   ��Ժʱ��
  --      adtd_time            C 1   ��Ժʱ��
  --      pati_deptid          N 1   ��ǰ����id
  --      wardarea_id          N 1   ��ǰ����id
  --      pati_bed             C 1   ��ǰ����
  --      inp_status           N 1   �Ƿ���Ժ��0/1
  --      inp_times            N 1   סԺ����
  --      inp_times_increment  N 1   =1ʱ-סԺ���������� =-1ʱ סԺ�����Լ�
  --      insurance_type       N 1  ����
  --      addr_list[]           C     ��ַ��Ϣ�б� 
  --        oper_fun            N  1  ��������:1-����,�޸�   2-ɾ�� 
  --        type                C  1  ��ַ��� 
  --        state               C  1  ��ַ_ʡ 
  --        city                C  1  ��ַ_�� 
  --        county              C  1  ��ַ_�� 
  --        township            C  1  ��ַ_�� 
  --        other               C  1  ��ַ_���� 
  --        code                C  1  ��������
  --����: Json_Out,��ʽ����
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json      Pljson;
  j_Temp      Pljson;
  j_Json_List Pljson_List;
  j_Addr_List Pljson_List;

  o_Json      Pljson;
  j_Jsoninput Pljson;

  n_����id ������Ϣ.����id%Type;

  n_��ҳid     ������Ϣ.��ҳid%Type;
  n_�����     ������Ϣ.�����%Type;
  n_סԺ��     ������Ϣ.סԺ��%Type;
  n_��ǰ����id ������Ϣ.��ǰ����id%Type;
  n_��ǰ����id ������Ϣ.��ǰ����id%Type;
  n_סԺ����   ������Ϣ.סԺ����%Type;
  n_��Ժ       ������Ϣ.��Ժ%Type;
  n_����       ������Ϣ.����%Type;
  v_��ǰ����   ������Ϣ.��ǰ����%Type;

  d_��Ժʱ�� ������Ϣ.��Ժʱ��%Type;
  d_��Ժʱ�� ������Ϣ.��Ժʱ��%Type;
  --���˵�ַ��Ϣ
  n_��ַ���  ���˵�ַ��Ϣ.��ַ���%Type;
  v_��ַ_ʡ   ���˵�ַ��Ϣ.ʡ%Type;
  v_��ַ_��   ���˵�ַ��Ϣ.��%Type;
  v_��ַ_��   ���˵�ַ��Ϣ.��%Type;
  v_��ַ_��   ���˵�ַ��Ϣ.����%Type;
  v_��ַ_���� ���˵�ַ��Ϣ.����%Type;
  v_��������  ���˵�ַ��Ϣ.��������%Type;

  n_��������     Number(2);
  n_��ҳid_b     Number(1);
  n_סԺ��_b     Number(1);
  n_�����_b     Number(1);
  n_��ǰ����id_b Number(1);
  n_��ǰ����id_b Number(1);
  n_סԺ����_b   Number(1);
  n_��Ժʱ��_b   Number(1);
  n_��Ժʱ��_b   Number(1);
  n_��Ժ_b       Number(1);
  n_��ǰ����_b   Number(1);
  n_����_b       Number(1);

Begin
  --�������
  j_Jsoninput := Pljson(Json_In);
  j_Json      := j_Jsoninput.Get_Pljson('input');
  j_Json_List := j_Json.Get_Pljson_List('pati_list');

  If j_Json_List Is Null Then
    Json_Out := Zljsonout('����ֵ����,���顣');
    Return;
  End If;
  For I In 1 .. j_Json_List.Count Loop
    j_Temp := Pljson(j_Json_List.Get(I));
  
    n_��ҳid_b     := Null;
    n_סԺ��_b     := Null;
    n_�����_b     := Null;
    n_��ǰ����id_b := Null;
    n_��ǰ����id_b := Null;
    n_סԺ����_b   := Null;
    n_��Ժʱ��_b   := Null;
    n_��Ժʱ��_b   := Null;
    n_��Ժ_b       := Null;
    n_��ǰ����_b   := Null;
    n_����_b       := Null;
    --����ID
    If j_Temp.Exist('pati_id') Then
      n_����id := j_Temp.Get_Number('pati_id');
    End If;
  
    --��ҳID
    If j_Temp.Exist('pati_pageid') Then
      n_��ҳid   := j_Temp.Get_Number('pati_pageid');
      n_��ҳid_b := 1;
    End If;
    --�����
    If j_Temp.Exist('outpatient_num') Then
      n_�����   := To_Number(j_Temp.Get_String('outpatient_num'));
      n_�����_b := 1;
    End If;
  
    --סԺ��
    If j_Temp.Exist('inpatient_num') Then
      n_סԺ��   := To_Number(j_Temp.Get_String('inpatient_num'));
      n_סԺ��_b := 1;
    End If;
    --��Ժʱ��
    If j_Temp.Exist('in_time') Then
      d_��Ժʱ��   := To_Date(j_Temp.Get_String('in_time'), 'yyyy-mm-dd hh24:mi:ss');
      n_��Ժʱ��_b := 1;
    End If;
    --��Ժʱ��
    If j_Temp.Exist('adtd_time') Then
      d_��Ժʱ��   := To_Date(j_Temp.Get_String('adtd_time'), 'yyyy-mm-dd hh24:mi:ss');
      n_��Ժʱ��_b := 1;
    End If;
    --��ǰ����ID
    If j_Temp.Exist('pati_deptid') Then
      n_��ǰ����id := j_Temp.Get_Number('pati_deptid');
    
      n_��ǰ����id_b := 1;
    End If;
    --��ǰ����ID
    If j_Temp.Exist('wardarea_id') Then
      n_��ǰ����id   := j_Temp.Get_Number('wardarea_id');
      n_��ǰ����id_b := 1;
    End If;
    --��ǰ����
    If j_Temp.Exist('pati_bed') Then
      v_��ǰ����   := j_Temp.Get_String('pati_bed');
      n_��ǰ����_b := 1;
    End If;
    --�Ƿ���Ժ
    If j_Temp.Exist('inp_status') Then
      n_��Ժ   := j_Temp.Get_Number('inp_status');
      n_��Ժ_b := 1;
    End If;
    --סԺ����
    If j_Temp.Exist('inp_times') Then
      n_סԺ����   := j_Temp.Get_Number('inp_times');
      n_סԺ����_b := 1;
    End If;
    --����
    If j_Temp.Exist('insurance_type') Then
      n_����   := j_Temp.Get_Number('insurance_type');
      n_����_b := 1;
    End If;
  
    Update ������Ϣ
    Set ��ҳid = Decode(n_��ҳid_b, 1, n_��ҳid, ��ҳid), ����� = Decode(n_�����_b, 1, n_�����, �����),
        סԺ�� = Decode(n_סԺ��_b, 1, n_סԺ��, סԺ��), ��Ժʱ�� = Decode(n_��Ժʱ��_b, 1, d_��Ժʱ��, ��Ժʱ��),
        ��Ժʱ�� = Decode(n_��Ժʱ��_b, 1, d_��Ժʱ��, ��Ժʱ��), ��ǰ����id = Decode(n_��ǰ����id_b, 1, n_��ǰ����id, ��ǰ����id),
        ��ǰ����id = Decode(n_��ǰ����id_b, 1, n_��ǰ����id, ��ǰ����id), ��ǰ���� = Decode(n_��ǰ����_b, 1, v_��ǰ����, ��ǰ����),
        ��Ժ = Decode(n_��Ժ_b, 1, n_��Ժ, ��Ժ), סԺ���� = Decode(n_סԺ����_b, 1, n_סԺ����, סԺ����), ���� = Decode(n_����_b, 1, n_����, ����)
    Where ����id = n_����id;
  
    --סԺ��������
    If j_Temp.Exist('inp_times_increment') Then
      n_סԺ���� := j_Temp.Get_Number('inp_times_increment');
      If Nvl(n_סԺ����, 0) = 1 Then
        Update ������Ϣ Set סԺ���� = Nvl(סԺ����, 0) + 1 Where ����id = n_����id;
      Elsif Nvl(n_סԺ����, 0) = -1 Then
        Update ������Ϣ Set סԺ���� = Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null) Where ����id = n_����id;
      End If;
    End If;
  
    --���µ�ַ��Ϣ 
    --    addr_list[]           C     ��ַ��Ϣ�б� 
    --      oper_fun            N  1  ��������:1-����,�޸�   2-ɾ�� 
    --      type                C  1  ��ַ��� 
    --      state               C  1  ��ַ_ʡ 
    --      city                C  1  ��ַ_�� 
    --      county              C  1  ��ַ_�� 
    --      township            C  1  ��ַ_�� 
    --      other               C  1  ��ַ_���� 
    --      code                C  1  �������� 
    If j_Temp.Exist('addr_list') Then
      j_Addr_List := Pljson_List();
      j_Addr_List := j_Temp.Get_Pljson_List('addr_list');
      If j_Addr_List Is Not Null Then
        For I In 1 .. j_Addr_List.Count Loop
          o_Json      := Pljson();
          o_Json      := Pljson(j_Addr_List.Get(I));
          n_��������  := o_Json.Get_Number('oper_fun');
          n_��ַ���  := o_Json.Get_Number('type');
          v_��ַ_ʡ   := o_Json.Get_String('state');
          v_��ַ_��   := o_Json.Get_String('city');
          v_��ַ_��   := o_Json.Get_String('county');
          v_��ַ_��   := o_Json.Get_String('township');
          v_��ַ_���� := o_Json.Get_String('other');
          v_��������  := o_Json.Get_String('code');
          Zl_���˵�ַ��Ϣ_Update_s(n_��������, n_����id, n_��ҳid, n_��ַ���, v_��ַ_ʡ, v_��ַ_��, v_��ַ_��, v_��ַ_��, v_��ַ_����, v_��������);
        End Loop;
      End If;
    End If;
  End Loop;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updateinpatistate;
/
Create Or Replace Procedure Zl_Patisvr_Updateoutpatistate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��������ﲡ�˾���״̬
  --      �����жϽ���Ǵ��ڣ�����������򲻸��£����߸���Ϊԭֵ��Ŀǰ��ʱδ�õ��������Ի������������չ
  --��Σ�Json_In:��ʽ
  --input
  --    pati_id            N 1 ����id
  --    pati_age           C 0 ����
  --    phone_number       C 0 �����ֻ���
  --    fee_category       C 0 �ѱ�
  --    visit_room         C 0 ���µľ�������
  --    visit_status       N 0 ���µľ���״̬
  --    visit_time         C 0 ���µľ���ʱ��
  --    outpatient_num     C 0 �����

  --����: Json_Out,��ʽ����
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  o_Json   Pljson;
  j_Jsonin Pljson;

  n_����id   ������Ϣ.����id%Type;
  v_����     ������Ϣ.����%Type;
  n_����״̬ ������Ϣ.����״̬%Type;
  v_�������� ������Ϣ.��������%Type;
  d_����ʱ�� ������Ϣ.����ʱ��%Type;
  v_�ѱ�     ������Ϣ.�ѱ�%Type;
  v_�ֻ���   ������Ϣ.�ֻ���%Type;
  n_�����   ������Ϣ.�����%Type;
  n_�ѱ����� �ѱ�.����%Type;

  n_�ѱ�_b     Number(1);
  n_����״̬_b Number(1);
  n_��������_b Number(1);
  n_����ʱ��_b Number(1);
  n_�ֻ���_b   Number(1);
  n_�����_b   Number(1);
  n_����_b     Number(1);
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  o_Json   := j_Jsonin.Get_Pljson('input');

  n_����id := o_Json.Get_Number('pati_id');

  If o_Json.Exist('phone_number') Then
    v_�ֻ���   := o_Json.Get_String('phone_number');
    n_�ֻ���_b := 1;
  End If;

  If o_Json.Exist('visit_status') Then
    n_����״̬   := o_Json.Get_Number('visit_status');
    n_����״̬_b := 1;
  End If;
  If o_Json.Exist('visit_room') Then
    v_��������   := o_Json.Get_String('visit_room');
    n_��������_b := 1;
  End If;
  If o_Json.Exist('visit_time') Then
    d_����ʱ��   := To_Date(o_Json.Get_String('visit_time'), 'yyyy-mm-dd hh24:mi:ss');
    n_����ʱ��_b := 1;
  End If;

  If o_Json.Exist('fee_category') Then
    v_�ѱ�   := o_Json.Get_String('fee_category');
    n_�ѱ�_b := 1;
  End If;

  If o_Json.Exist('outpatient_num') Then
    n_�����   := o_Json.Get_String('outpatient_num');
    n_�����_b := 1;
  End If;

  If o_Json.Exist('pati_age') Then
    v_����   := o_Json.Get_String('pati_age');
    n_����_b := 1;
  End If;

  If v_�ѱ� Is Not Null Then
    Select Max(����) Into n_�ѱ����� From �ѱ� Where ���� = v_�ѱ�; --2-��̬�ѱ𲻸���
    If n_�ѱ����� = 2 Then
      n_�ѱ�_b := 0;
    End If;
  End If;

  Update ������Ϣ
  Set �ѱ� = Decode(n_�ѱ�_b, 1, v_�ѱ�, �ѱ�), �ֻ��� = Decode(n_�ֻ���_b, 1, v_�ֻ���, �ֻ���), ����״̬ = Decode(n_����״̬_b, 1, n_����״̬, ����״̬),
      �������� = Decode(n_��������_b, 1, v_��������, ��������), ����ʱ�� = Decode(n_����ʱ��_b, 1, d_����ʱ��, ����ʱ��),
      ����� = Decode(n_�����_b, 1, n_�����, �����), ���� = Decode(n_����_b, 1, v_����, ����)
  Where ����id = n_����id;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updateoutpatistate;
/
Create Or Replace Procedure Zl_Patisvr_Updatepatiarchives
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------- 
  --���ܣ��޸Ĳ��˵�����Ϣ 
  --��Σ�Json_In:��ʽ 
  --input 
  --    oper_fun              N  1   0-Ҫ���²�����Ϣ�� 1-�����²�����Ϣ�� 
  --    is_realname_check     N  1  �Ƿ�ʵ�����:1-ʵ�����;0-����� 
  --    pati_id               N  1  ����id:�������� 
  --    pati_pageid           N  1  ��ҳID 
  --    pati_name_old         N     ��������(δ�޸�ǰ������):���磺�²��� 
  --    pati_name             N  1  �������� 
  --    pati_sex              C  1  �Ա� 
  --    pati_age              C  1  ���� 
  --    pati_type             C  1  ��������(��ͨ��ҽ��������) 
  --    pati_birthdate        C  1  ��������:yyyy-mm-dd hh24:mi:ss 
  --    phone_number          C  1  �ֻ��� 
  --    insurance_num         C  1  ҽ���� 
  --    pati_idcard           C  1  ���֤�� 
  --    outpatient_num        C  1  ����� 
  --    fee_category          C  1  �ѱ� 
  --    mdlpay_mode_name      C  1  ҽ�Ƹ��ʽ���� 
  --    country_name          C  1  ���� 
  --    native_place          C  1  ���� 
  --    nation_name           C  1  ���� 
  --    mari_status           C  1  ����״�� 
  --    ocpt_name             C  1  ְҵ 
  --    edu_name              C  1  ѧ�� 
  --    pati_identity         C  1  ��� 
  --    insurance_type        N  1  ���� 
  --    emp_name              C  1  ������λ 
  --    emp_postcode          C  1  ��λ�ʱ� 
  --    emp_phno              C  1  ��λ�绰 
  --    emp_bank_name         C   1   ��λ������ 
  --    emp_bank_accnum       C   1   ��λ�ʺ� 
  --    ctt_unit_id           N  1  ��ͬ��λid 
  --    pat_home_addr         C  1  ��ͥ��ַ 
  --    pat_home_phno         C  1  ��ͥ�绰 
  --    pat_home_postcode     C  1  ��ͥ��ַ�ʱ� 
  --    region                C  1  ���� 
  --    pat_baddr             C  1  �����ص� 
  --    pat_hous_addr         C  1  ���ڵ�ַ 
  --    pat_hous_postcode     C  1  ���ڵ�ַ�ʱ� 
  --    pat_grdn_name         C  1  �໤�� 
  --    vcard_no              C  1  ���￨�� 
  --    vcard_pwd             C  1  ����֤�� 
  --    iccard_no             C  1  Ic���� 
  --    create_time           C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss 
  --    operator_name         C  1  ����Ա���� 
  --    cardno_clear          N     ������￨��Ϣ 
  --    pati_wardarea_id      N     ��ǰ����id 
  --    pati_bed              C     ��ǰ���� 
  --    idcard_sign           N     ���֤ǩԼ 
  --    idcard_sign_pwd       C     ǩԼ���� 
  --    cert_no_other         C  1 ����֤�� 
  --    qq                    C      qq 
  --    email                 C      email 
  --    emp_addr              C     ��λ��ַ 
  --    contacts              C     ������ϵ����Ϣ�ڵ� 
  --      name                C  1  ��ϵ������ 
  --      idcard              C  1  ��ϵ�����֤�� 
  --      phone               C  1  ��ϵ�˵绰 
  --      relation            C  1  ��ϵ�˹�ϵ 
  --      address             C     ��ϵ�˵�ַ 
  --    community_info        C     ������Ϣ�ڵ� 
  --      num                 N  1  ������� 
  --      code                C  1  �������� 
  --      oper_type           N  1  ������������ 
  --    visit_info            C     ������Ϣ�ڵ� 
  --      status              N     ���µľ���״̬ 
  --      room                C     ���µľ������� 
  --      time                C     ����ʱ��:yyyy-mm-dd hh24:mi:ss 
  --    addr_list[]           C     ��ַ��Ϣ�б� 
  --      oper_fun            N  1  ��������:1-����,�޸�   2-ɾ�� 
  --      type                C  1  ��ַ��� 
  --      state               C  1  ��ַ_ʡ 
  --      city                C  1  ��ַ_�� 
  --      county              C  1  ��ַ_�� 
  --      township            C  1  ��ַ_�� 
  --      other               C  1  ��ַ_���� 
  --      code                C  1  �������� 
  --      visit_or_in         N  1  �Ƿ���ھ������סԺ��Ϣ 
  --    ext_list[]            C     ������Ϣ�����б� 
  --      info_name           C  1  ��Ϣ�� 
  --      upd_info_value      N  1  �޸ĵ���Ϣֵ 
  --      visit_id            N     ����id_In 
  --    cert_list[]                 ֤���б�(��Ҫ�ǵ��ɰ󿨴���) 
  --      cert_name           C  1  ֤������ 
  --      cert_no             C  1  ֤�ź��� 
  --    oper_allergic_drugs N  1  ����ҩ��  0-ɾ������� 1-������¼���� 
  --    allergic_drugs_list[]       ���˹���ҩ���б�:������ʱ������ɾ������ҩ�����ķ�ʽ 
  --      oper_type            N  1  0-���� 1-ɾ�� 
  --      pat_algc_cadn_id    N  1  ����ҩƷID 
  --      pat_algc_cadn       C  1  ����ҩ������ 
  --      allergy_info        C  1  ��ÿҩ�ﷴӦ 
  --      allergic_drugs      C  1  ����ҩƷID:����ҩ������ƴ�� 
  --    immune_list[]         C     ���������б� 
  --      vaccinate_time      C  1  ����ʱ��:yyyy-mm-dd hh24:mi:ss 
  --      vaccinate_name      C  1  �������� 
  --    card_property_list[]  C     ҽ�ƿ������б� 
  --      cardtype_id         N  1  ҽ�ƿ����ID 
  --      card_no             C  1  ���� 
  --      info_name           C  1  ��Ϣ�� 
  --      info_value          N  1  ��Ϣֵ 
  --      item_list[]         ���²�����Ϣĳһ���ֶε�ֵ 
  --      item_name           C  1  �ֶ��� 
  --      item_value          C  1   �ֶ�ֵ 

  --����: Json_Out,��ʽ���� 
  --  output 
  --    code                  N 1   Ӧ���룺0-ʧ�ܣ�1-�ɹ� 
  --    message               C 1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ 
  --------------------------------------------------------------------------- 

  j_Json         Pljson;
  j_Jsonin       Pljson;
  j_Jsonlist     Pljson_List := Pljson_List();
  o_Json         Pljson;
  o_Json1        Pljson;
  n_ʵ�����     Number(1);
  n_����id       ������Ϣ.����id%Type;
  n_��ҳid       ������Ϣ.��ҳid%Type;
  v_����_Old     ������Ϣ.����%Type;
  v_����         ������Ϣ.����%Type;
  v_���֤��     ������Ϣ.���֤��%Type;
  v_����         ������Ϣ.����%Type;
  v_�Ա�         ������Ϣ.�Ա�%Type;
  d_��������     ������Ϣ.��������%Type;
  v_�ֻ���       ������Ϣ.�ֻ���%Type;
  v_��ͥ�绰     ������Ϣ.��ͥ�绰%Type;
  n_�����       ������Ϣ.�����%Type;
  v_�ѱ�         ������Ϣ.�ѱ�%Type;
  v_ҽ�Ƹ��ʽ ������Ϣ.ҽ�Ƹ��ʽ%Type;
  v_����         ������Ϣ.����%Type;
  v_����         ������Ϣ.����%Type;
  v_����         ������Ϣ.����%Type;
  v_����         ������Ϣ.����״��%Type;
  v_ְҵ         ������Ϣ.ְҵ%Type;
  v_ѧ��         ������Ϣ.ѧ��%Type;
  v_������λ     ������Ϣ.������λ%Type;
  n_��ͬ��λid   ������Ϣ.��ͬ��λid%Type;
  v_��λ�绰     ������Ϣ.��λ�绰%Type;
  v_��λ�ʱ�     ������Ϣ.��λ�ʱ�%Type;
  v_��ͥ��ַ     ������Ϣ.��ͥ��ַ%Type;
  v_��ͥ��ַ�ʱ� ������Ϣ.��ͥ��ַ�ʱ�%Type;
  v_���ڵ�ַ     ������Ϣ.���ڵ�ַ%Type;
  v_���ڵ�ַ�ʱ� ������Ϣ.���ڵ�ַ�ʱ�%Type;
  d_�Ǽ�ʱ��     ������Ϣ.�Ǽ�ʱ��%Type;
  v_ҽ����       ������Ϣ.ҽ����%Type;
  v_����         ������Ϣ.����%Type;
  v_�໤��       ������Ϣ.�໤��%Type;
  v_�����ص�     ������Ϣ.�����ص�%Type;
  v_���         ������Ϣ.���%Type;
  v_����Ա����   ����ҽ�ƿ���Ϣ.������%Type;
  n_����id       ����������Ϣ.����%Type;
  v_��������     ����������Ϣ.������%Type;
  n_��������     ����������Ϣ.��������%Type;
  v_������       ҽ�ƿ����.����%Type;
  n_�����id     ҽ�ƿ����.Id%Type;
  v_����         ����ҽ�ƿ���Ϣ.����%Type;
  n_���ų���     ҽ�ƿ����.���ų���%Type;
  v_����         ҽ�ƿ����.����%Type;
  v_������       ����ҽ�ƿ���Ϣ.����%Type;
  v_�䶯ԭ��     ����ҽ�ƿ��䶯.�䶯ԭ��%Type;
  d_��ֹʹ��ʱ�� ����ҽ�ƿ���Ϣ.��ֹʹ��ʱ��%Type;
  n_����ҩƷid   ���˹���ҩ��.����ҩ��id%Type;
  v_����ҩ������ ���˹���ҩ��.����ҩ��%Type;
  v_��ÿҩ�ﷴӦ ���˹���ҩ��.������Ӧ%Type;
  d_����ʱ��     �������߼�¼.����ʱ��%Type;
  v_��������     �������߼�¼.��������%Type;
  v_���￨��     ������Ϣ.���￨��%Type;
  v_����֤��     ������Ϣ.����֤��%Type;
  v_Ic����       ������Ϣ.Ic����%Type;
  v_��ǰ����     ������Ϣ.��ǰ����%Type;
  n_��ǰ����id   ������Ϣ.��ǰ����id %Type;
  n_����״̬     ������Ϣ.����״̬%Type;
  v_��������     ������Ϣ.��������%Type;
  d_����ʱ��     ������Ϣ.����ʱ��%Type;
  v_����֤��     ������Ϣ.����֤��%Type;
  v_��λ�ʺ�     ������Ϣ.��λ�ʺ�%Type;
  v_��λ������   ������Ϣ.��λ������%Type;
  v_��������     ������Ϣ.��������%Type;
  v_Qq           ������Ϣ.Qq%Type;
  v_Email        ������Ϣ.Email%Type;
  n_�Ƿ����     Number;
  v_��λ��ַ     ������Ϣ.��λ��ַ%Type;

  n_������￨��Ϣ Number(1);
  n_Count          Number(10);
  --��ϵ�� 
  v_��ϵ������     ������Ϣ.��ϵ������%Type;
  v_��ϵ�˹�ϵ     ������Ϣ.��ϵ�˹�ϵ%Type;
  v_��ϵ�˵绰     ������Ϣ.��ϵ�˵绰%Type;
  v_��ϵ�����֤�� ������Ϣ.��ϵ�����֤��%Type;
  v_��ϵ�˵�ַ     ������Ϣ.��ϵ�˵�ַ%Type;
  --������Ϣ�ӱ� 
  v_��Ϣ�� ������Ϣ�ӱ�.��Ϣ��%Type;
  v_��Ϣֵ ������Ϣ�ӱ�.��Ϣֵ%Type;
  --���˵�ַ��Ϣ 
  n_��������       Number(3);
  n_��ַ����       ���˵�ַ��Ϣ.��ַ���%Type;
  v_ʡ             ���˵�ַ��Ϣ.ʡ%Type;
  v_��             ���˵�ַ��Ϣ.��%Type;
  v_��             ���˵�ַ��Ϣ.��%Type;
  v_����           ���˵�ַ��Ϣ.����%Type;
  v_����           ���˵�ַ��Ϣ.����%Type;
  v_��������       ���˵�ַ��Ϣ.��������%Type;
  n_����           ������Ϣ.����%Type;
  v_Msg            Varchar2(4000);
  v_Strtmpbefor    Varchar2(4000);
  v_�ֶ���         Varchar2(1000);
  v_�ֶ�ֵ         Varchar2(3682);
  v_Sql            Varchar2(3682);
  n_�ѱ�����       Number(1);
  n_����_b         Number(1); --�Ӻ�׺_b:Ϊ1ʱ��ʾ��Ӧ�ֶε�json�ڵ���ڣ�Ϊ0ʱ��ʾ��Ӧ�ֶε�json�ڵ㲻���� 
  n_�Ա�_b         Number(1);
  n_����_b         Number(1);
  n_��������_b     Number(1);
  n_�����_b       Number(1);
  n_�ѱ�_b         Number(1);
  n_ҽ����_b       Number(1);
  n_����_b         Number(1);
  n_ҽ�Ƹ��ʽ_b Number(1);
  n_����_b         Number(1);
  n_�ֻ���_b       Number(1);
  n_��ǰ����_b     Number(1);
  n_��ǰ����_b     Number(1);
  n_���֤��_b     Number(1);
  n_����_b         Number(1);
  n_����_b         Number(1);
  n_����״��_b     Number(1);
  n_�����ص�_b     Number(1);
  n_ѧ��_b         Number(1);
  n_ְҵ_b         Number(1);
  n_����_b         Number(1);
  n_������λ_b     Number(1);
  n_��ͬ��λid_b   Number(1);
  n_��λ�绰_b     Number(1);
  n_��λ�ʱ�_b     Number(1);
  n_��ͥ��ַ_b     Number(1);
  n_��ͥ�绰_b     Number(1);
  n_��ͥ��ַ�ʱ�_b Number(1);
  n_���ڵ�ַ_b     Number(1);
  n_���ڵ�ַ�ʱ�_b Number(1);
  n_��ϵ��_b       Number(1);
  n_����_b         Number(1);
  n_���_b         Number(1);
  n_�໤��_b       Number(1);
  n_���￨��_b     Number(1);
  n_����֤��_b     Number(1);
  n_Ic����_b       Number(1);
  n_����֤��_b     Number(1);
  n_��������_b     Number(1);
  n_��λ������_b   Number(1);
  n_��λ�ʺ�_b     Number(1);
  n_Qq_b           Number(1);
  n_Email_b        Number(1);
  n_����id         ������Ϣ�ӱ�.����id%Type;
  n_��λ��ַ_b     Number(1);

  n_�ֵ       Number(10);
  n_���ֵ       Number(10);
  n_����         Number(1);
  n_����ҩ����� Number(1);
  n_��ʽ         Number(1);
  c_����ҩ��     Clob;
  l_����ҩ��     t_Strlist := t_Strlist();

Begin
  --������� 
  j_Jsonin := Pljson(Json_In);
  If j_Jsonin Is Null Then
    Json_Out := Zljsonout('δ�����κ���Ϣ������');
    Return;
  Else
    o_Json := j_Jsonin.Get_Pljson('input');
  End If;
  --    is_realname_check     N  1  �Ƿ�ʵ�����:1-ʵ�����;0-����� 
  --    pati_id               N  1  ����id:�������� 
  --    pati_pageid           N  1  ��ҳID 
  --    pati_name_old         N     ��������(δ�޸�ǰ������):���磺�²��� 
  --    pati_name             N  1  �������� 
  --    pati_sex              C  1  �Ա� 
  --    pati_age              C  1  ���� 
  --    pati_birthdate        C  1  ��������:yyyy-mm-dd hh24:mi:ss 
  n_����id       := o_Json.Get_Number('pati_id');
  n_��ҳid       := o_Json.Get_Number('pati_pageid');
  n_����         := o_Json.Get_Number('oper_fun');
  n_����ҩ����� := o_Json.Get_Number('oper_allergic_drugs');
  If Nvl(n_����, 0) = 0 Then
    n_ʵ����� := o_Json.Get_Number('is_realname_check');
  
    If Nvl(n_����id, 0) = 0 Then
      Json_Out := Zljsonout('δ���벡��id�����ܱ���');
      Return;
    End If;
    v_����_Old := o_Json.Get_String('pati_name_old');
  
    If o_Json.Exist('pati_name') Then
      v_����   := o_Json.Get_String('pati_name');
      n_����_b := 1;
    End If;
  
    If o_Json.Exist('pati_sex') Then
      v_�Ա�   := o_Json.Get_String('pati_sex');
      n_�Ա�_b := 1;
    End If;
  
    If o_Json.Exist('pati_age') Then
      v_����   := o_Json.Get_String('pati_age');
      n_����_b := 1;
    End If;
  
    If o_Json.Exist('pati_birthdate') Then
      d_��������   := To_Date(o_Json.Get_String('pati_birthdate'), 'yyyy-mm-dd hh24:mi:ss');
      n_��������_b := 1;
    End If;
  
    --    phone_number          C  1  �ֻ��� 
    --    insurance_num         C  1  ҽ���� 
    --    pati_idcard           C  1  ���֤�� 
    --    outpatient_num        C  1  ����� 
    --    fee_category          C  1  �ѱ� 
    --    mdlpay_mode_name      C  1  ҽ�Ƹ��ʽ���� 
    If o_Json.Exist('phone_number') Then
      v_�ֻ���   := o_Json.Get_String('phone_number');
      n_�ֻ���_b := 1;
    End If;
  
    If o_Json.Exist('insurance_num') Then
      v_ҽ����   := o_Json.Get_String('insurance_num');
      n_ҽ����_b := 1;
    End If;
  
    If o_Json.Exist('pati_idcard') Then
      v_���֤��   := o_Json.Get_String('pati_idcard');
      n_���֤��_b := 1;
    End If;
  
    If o_Json.Exist('outpatient_num') Then
      n_�����   := To_Number(o_Json.Get_String('outpatient_num'));
      n_�����_b := 1;
    End If;
  
    If o_Json.Exist('fee_category') Then
      v_�ѱ�   := o_Json.Get_String('fee_category');
      n_�ѱ�_b := 1;
    End If;
  
    If o_Json.Exist('mdlpay_mode_name') Then
      v_ҽ�Ƹ��ʽ   := o_Json.Get_String('mdlpay_mode_name');
      n_ҽ�Ƹ��ʽ_b := 1;
    End If;
  
    --    country_name          C  1  ���� 
    --    native_place          C  1  ���� 
    --    nation_name           C  1  ���� 
    --    mari_status           C  1  ����״�� 
    --    ocpt_name             C  1  ְҵ 
    --    edu_name              C  1  ѧ�� 
    --    pati_identity         C  1  ��� 
  
    If o_Json.Exist('country_name') Then
      v_����   := o_Json.Get_String('country_name');
      n_����_b := 1;
    End If;
  
    If o_Json.Exist('native_place') Then
      v_����   := o_Json.Get_String('native_place');
      n_����_b := 1;
    End If;
  
    If o_Json.Exist('nation_name') Then
      v_����   := o_Json.Get_String('nation_name');
      n_����_b := 1;
    End If;
  
    If o_Json.Exist('mari_status') Then
      v_����       := o_Json.Get_String('mari_status');
      n_����״��_b := 1;
    End If;
  
    If o_Json.Exist('ocpt_name') Then
      v_ְҵ   := o_Json.Get_String('ocpt_name');
      n_ְҵ_b := 1;
    End If;
  
    If o_Json.Exist('edu_name') Then
      v_ѧ��   := o_Json.Get_String('edu_name');
      n_ѧ��_b := 1;
    End If;
  
    If o_Json.Exist('pati_identity') Then
      v_���   := o_Json.Get_String('pati_identity');
      n_���_b := 1;
    End If;
  
    --    insurance_type        N  1  ���� 
    --    emp_name              C  1  ������λ 
    --    emp_postcode          C  1  ��λ�ʱ� 
    --    emp_phno              C  1  ��λ�绰 
    --    ctt_unit_id           N  1  ��ͬ��λid 
    --    pat_home_addr         C  1  ��ͥ��ַ 
    --    pat_home_phno         C  1  ��ͥ�绰 
    --    pat_home_postcode     C  1  ��ͥ��ַ�ʱ� 
    If o_Json.Exist('insurance_type') Then
      n_����   := o_Json.Get_Number('insurance_type');
      n_����_b := 1;
    End If;
  
    If o_Json.Exist('emp_name') Then
      v_������λ   := o_Json.Get_String('emp_name');
      n_������λ_b := 1;
    End If;
  
    If o_Json.Exist('emp_postcode') Then
      v_��λ�ʱ�   := o_Json.Get_String('emp_postcode');
      n_��λ�ʱ�_b := 1;
    End If;
  
    If o_Json.Exist('emp_phno') Then
      v_��λ�绰   := o_Json.Get_String('emp_phno');
      n_��λ�绰_b := 1;
    End If;
  
    If o_Json.Exist('ctt_unit_id') Then
      n_��ͬ��λid := o_Json.Get_Number('ctt_unit_id');
      If n_��ͬ��λid > 0 Then
        n_��ͬ��λid_b := 1;
      End If;
    End If;
  
    If o_Json.Exist('pat_home_addr') Then
      v_��ͥ��ַ   := o_Json.Get_String('pat_home_addr');
      n_��ͥ��ַ_b := 1;
    End If;
  
    If o_Json.Exist('pat_home_phno') Then
      v_��ͥ�绰   := o_Json.Get_String('pat_home_phno');
      n_��ͥ�绰_b := 1;
    End If;
  
    If o_Json.Exist('pat_home_postcode') Then
      v_��ͥ��ַ�ʱ�   := o_Json.Get_String('pat_home_postcode');
      n_��ͥ��ַ�ʱ�_b := 1;
    End If;
  
    --    region                C  1  ���� 
    --    pat_baddr             C  1  �����ص� 
    --    pat_hous_addr         C  1  ���ڵ�ַ 
    --    pat_hous_postcode     C  1  ���ڵ�ַ�ʱ� 
    --    pat_grdn_name         C  1  �໤�� 
    If o_Json.Exist('region') Then
      v_����   := o_Json.Get_String('region');
      n_����_b := 1;
    End If;
  
    If o_Json.Exist('pat_baddr') Then
      v_�����ص�   := o_Json.Get_String('pat_baddr');
      n_�����ص�_b := 1;
    End If;
  
    If o_Json.Exist('pat_hous_addr') Then
      v_���ڵ�ַ   := o_Json.Get_String('pat_hous_addr');
      n_���ڵ�ַ_b := 1;
    End If;
  
    If o_Json.Exist('pat_hous_postcode') Then
      v_���ڵ�ַ�ʱ�   := o_Json.Get_String('pat_hous_postcode');
      n_���ڵ�ַ�ʱ�_b := 1;
    End If;
  
    If o_Json.Exist('pat_grdn_name') Then
      v_�໤��   := o_Json.Get_String('pat_grdn_name');
      n_�໤��_b := 1;
    End If;
  
    --    vcard_no              C  1  ���￨�� 
    --    vcard_pwd             C  1  ����֤�� 
    --    iccard_no             C  1  Ic���� 
    --    create_time           C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss 
    --    operator_name         C  1  ����Ա���� 
    --    pati_wardarea_id      N     ��ǰ����id 
    --    pati_bed              C     ��ǰ���� 
  
    If o_Json.Exist('vcard_no') Then
      v_���￨��   := o_Json.Get_String('vcard_no');
      n_���￨��_b := 1;
    End If;
  
    If o_Json.Exist('vcard_pwd') Then
      v_����֤��   := o_Json.Get_String('vcard_pwd');
      n_����֤��_b := 1;
    End If;
  
    If o_Json.Exist('iccard_no') Then
      v_Ic����   := o_Json.Get_String('iccard_no');
      n_Ic����_b := 1;
    End If;
  
    If o_Json.Exist('create_time') Then
      d_�Ǽ�ʱ�� := To_Date(o_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss');
    End If;
    If d_�Ǽ�ʱ�� Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    End If;
  
    If o_Json.Exist('operator_name') Then
      v_����Ա���� := o_Json.Get_String('operator_name');
    End If;
  
    If o_Json.Exist('pati_wardarea_id') Then
      n_��ǰ����id := o_Json.Get_Number('pati_wardarea_id');
      n_��ǰ����_b := 1;
    End If;
  
    If o_Json.Exist('pati_bed') Then
      v_��ǰ����   := o_Json.Get_String('pati_bed');
      n_��ǰ����_b := 1;
    End If;
  
    --    cert_no_other       C   1   ����֤�� 
    --    pati_type           C   1   ��������(��ͨ��ҽ��������) 
    --    emp_bank_name       C   1   ��λ������ 
    --    emp_bank_accnum     C   1   ��λ�ʺ� 
  
    If o_Json.Exist('cert_no_other') Then
      v_����֤��   := o_Json.Get_String('cert_no_other');
      n_����֤��_b := 1;
    End If;
  
    If o_Json.Exist('pati_type') Then
      v_��������   := o_Json.Get_String('pati_type');
      n_��������_b := 1;
    End If;
  
    If o_Json.Exist('qq') Then
      v_Qq   := o_Json.Get_String('qq');
      n_Qq_b := 1;
    End If;
  
    If o_Json.Exist('email') Then
      v_Email   := o_Json.Get_String('email');
      n_Email_b := 1;
    End If;
  
    If o_Json.Exist('emp_bank_name') Then
      v_��λ������   := o_Json.Get_String('emp_bank_name');
      n_��λ������_b := 1;
    End If;
    If o_Json.Exist('emp_bank_accnum') Then
      v_��λ�ʺ�   := o_Json.Get_String('emp_bank_accnum');
      n_��λ�ʺ�_b := 1;
    End If;
  
    If o_Json.Exist('emp_addr') Then
      v_��λ��ַ   := o_Json.Get_String('emp_addr');
      n_��λ��ַ_b := 1;
    End If;
  
    --    contacts              C     ������ϵ����Ϣ�ڵ� 
    --      name                C  1  ��ϵ������ 
    --      idcard              C  1  ��ϵ�����֤�� 
    --      phone               C  1  ��ϵ�˵绰 
    --      relation            C  1  ��ϵ�˹�ϵ 
    --      address             C     ��ϵ�˵�ַ 
    o_Json1 := Pljson();
    o_Json1 := o_Json.Get_Pljson('contacts');
    If Not o_Json1 Is Null Then
      v_��ϵ������     := o_Json1.Get_String('name');
      v_��ϵ�˹�ϵ     := o_Json1.Get_String('relation');
      v_��ϵ�����֤�� := o_Json1.Get_String('idcard');
      v_��ϵ�˵绰     := o_Json1.Get_String('phone');
      v_��ϵ�˵�ַ     := o_Json1.Get_String('address');
      n_��ϵ��_b       := 1;
    End If;
  
    --        visit_info          ��������Ǽ���Ϣ 
    --          status        N 1 ���µľ���״̬ 
    --          room          C 1 ���µľ������� 
    --          time          C 1 ���µľ���ʱ�� 
    o_Json1 := Pljson();
    o_Json1 := o_Json.Get_Pljson('visit_info');
    If o_Json1 Is Not Null Then
      n_����״̬ := o_Json1.Get_Number('status');
      v_�������� := o_Json1.Get_String('room');
      d_����ʱ�� := To_Date(o_Json1.Get_String('time'), 'yyyy-mm-dd hh24:mi:ss');
      n_����_b   := 1;
    End If;
  
    If Nvl(n_ʵ�����, 0) = 1 Then
      Select Zl_Fun_Checkidentify(1, n_����id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    End If;
    If v_�ѱ� Is Not Null Then
      Select Max(����) Into n_�ѱ����� From �ѱ� Where ���� = v_�ѱ�; --2-��̬�ѱ𲻸��� 
      If n_�ѱ����� = 2 Then
        n_�ѱ�_b := 0;
      End If;
    End If;
  
    Update ������Ϣ
    Set ���� = Decode(n_����_b, 1, v_����, ����), �Ա� = Decode(n_�Ա�_b, 1, v_�Ա�, �Ա�), ���� = Decode(n_����_b, 1, v_����, ����),
        �������� = Decode(n_��������_b, 1, d_��������, ��������), ����� = Decode(n_�����_b, 1, n_�����, �����), �ѱ� = Decode(n_�ѱ�_b, 1, v_�ѱ�, �ѱ�),
        ҽ���� = Decode(n_ҽ����_b, 1, v_ҽ����, ҽ����), ���� = Decode(n_����_b, 1, n_����, ����),
        ҽ�Ƹ��ʽ = Decode(n_ҽ�Ƹ��ʽ_b, 1, v_ҽ�Ƹ��ʽ, ҽ�Ƹ��ʽ), �ֻ��� = Decode(n_�ֻ���_b, 1, v_�ֻ���, �ֻ���),
        ���֤�� = Decode(n_���֤��_b, 1, v_���֤��, ���֤��), �����ص� = Decode(n_�����ص�_b, 1, v_�����ص�, �����ص�),
        ����״�� = Decode(n_����״��_b, 1, v_����, ����״��), ���� = Decode(n_����_b, 1, v_����, ����), ѧ�� = Decode(n_ѧ��_b, 1, v_ѧ��, ѧ��),
        ְҵ = Decode(n_ְҵ_b, 1, v_ְҵ, ְҵ), ���� = Decode(n_����_b, 1, v_����, ����), ���� = Decode(n_����_b, 1, v_����, ����),
        ������λ = Decode(n_������λ_b, 1, v_������λ, ������λ), ��ͬ��λid = Decode(n_��ͬ��λid_b, 1, n_��ͬ��λid, ��ͬ��λid),
        ��λ�绰 = Decode(n_��λ�绰_b, 1, v_��λ�绰, ��λ�绰), ��λ�ʱ� = Decode(n_��λ�ʱ�_b, 1, v_��λ�ʱ�, ��λ�ʱ�),
        ��ͥ��ַ = Decode(n_��ͥ��ַ_b, 1, v_��ͥ��ַ, ��ͥ��ַ), ��ͥ�绰 = Decode(n_��ͥ�绰_b, 1, v_��ͥ�绰, ��ͥ�绰),
        ��ͥ��ַ�ʱ� = Decode(n_��ͥ��ַ�ʱ�_b, 1, v_��ͥ��ַ�ʱ�, ��ͥ��ַ�ʱ�), ���ڵ�ַ = Decode(n_���ڵ�ַ_b, 1, v_���ڵ�ַ, ���ڵ�ַ),
        ���ڵ�ַ�ʱ� = Decode(n_���ڵ�ַ�ʱ�_b, 1, v_���ڵ�ַ�ʱ�, ���ڵ�ַ�ʱ�), ��ϵ������ = Decode(n_��ϵ��_b, 1, v_��ϵ������, ��ϵ������),
        ��ϵ�˹�ϵ = Decode(n_��ϵ��_b, 1, v_��ϵ�˹�ϵ, ��ϵ�˹�ϵ), ��ϵ�����֤�� = Decode(n_��ϵ��_b, 1, v_��ϵ�����֤��, ��ϵ�����֤��),
        ��ϵ�˵绰 = Decode(n_��ϵ��_b, 1, v_��ϵ�˵绰, ��ϵ�˵绰), ��ϵ�˵�ַ = Decode(n_��ϵ��_b, 1, v_��ϵ�˵�ַ, ��ϵ�˵�ַ),
        ��ǰ���� = Decode(n_��ǰ����_b, 1, v_��ǰ����, ��ǰ����), ��ǰ����id = Decode(n_��ǰ����_b, 1, n_��ǰ����id, ��ǰ����id),
        ����״̬ = Decode(n_����_b, 1, n_����״̬, ����״̬), �������� = Decode(n_����_b, 1, v_��������, ��������),
        ����ʱ�� = Decode(n_����_b, 1, d_����ʱ��, ����ʱ��), ���� = Decode(n_����_b, 1, v_����, ����), ��� = Decode(n_���_b, 1, v_���, ���),
        �໤�� = Decode(n_�໤��_b, 1, v_�໤��, �໤��), ���￨�� = Decode(n_���￨��_b, 1, v_���￨��, ���￨��),
        ����֤�� = Decode(n_����֤��_b, 1, v_����֤��, ����֤��), Ic���� = Decode(n_Ic����_b, 1, v_Ic����, Ic����),
        ����֤�� = Decode(n_����֤��_b, 1, v_����֤��, ����֤��), �������� = Decode(n_��������_b, 1, v_��������, ��������),
        ��λ������ = Decode(n_��λ������_b, 1, v_��λ������, ��λ������), ��λ�ʺ� = Decode(n_��λ�ʺ�_b, 1, v_��λ�ʺ�, ��λ�ʺ�),
        Qq = Decode(n_Qq_b, 1, v_Qq, Qq), Email = Decode(n_Email_b, 1, v_Email, Email),
        ��λ��ַ = Decode(n_��λ��ַ_b, 1, v_��λ��ַ, ��λ��ַ)
    Where ����id = n_����id And Decode(n_��ҳid, Null, 0, ��ҳid) = Decode(n_��ҳid, Null, 0, n_��ҳid) And
          Decode(v_����_Old, Null, '-', ����) = Decode(v_����_Old, Null, '-', v_����_Old);
  
    n_������￨��Ϣ := o_Json.Get_Number('cardno_clear');
    If Nvl(n_������￨��Ϣ, 0) = 1 Then
      Update ������Ϣ Set ���￨�� = Null, ����֤�� = Null, Ic���� = Null Where ����id = n_����id;
    End If;
  
    If Nvl(n_ʵ�����, 0) = 1 Then
      Select Zl_Fun_Checkidentify(1, n_����id, v_Strtmpbefor) Into v_Msg From Dual;
    End If;
  End If;
  --������Ϣ 
  --    community_info        C     ������Ϣ�ڵ� 
  --      num                 N  1  ������� 
  --      code                C  1  �������� 
  --      oper_type           N  1  ������������ 
  o_Json1 := Pljson();
  o_Json1 := o_Json.Get_Pljson('community_info');
  If o_Json1 Is Not Null Then
    n_����id   := o_Json1.Get_Number('num');
    v_�������� := o_Json1.Get_String('code');
    n_�������� := o_Json1.Get_Number('oper_type');
    --���������� 
    If n_����id <> 0 And v_�������� Is Not Null Then
      Zl_����������Ϣ_Insert(n_����id, n_����id, v_��������, n_��������, d_�Ǽ�ʱ��);
    End If;
  End If;

  --���µ�ַ��Ϣ 
  --    addr_list[]           C     ��ַ��Ϣ�б� 
  --      oper_fun            N  1  ��������:1-����,�޸�   2-ɾ�� 
  --      type                C  1  ��ַ��� 
  --      state               C  1  ��ַ_ʡ 
  --      city                C  1  ��ַ_�� 
  --      county              C  1  ��ַ_�� 
  --      township            C  1  ��ַ_�� 
  --      other               C  1  ��ַ_���� 
  --      code                C  1  �������� 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('addr_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1    := Pljson();
      o_Json1    := Pljson(j_Jsonlist.Get(I));
      n_�������� := o_Json1.Get_Number('oper_fun');
      n_��ַ���� := o_Json1.Get_Number('type');
      v_ʡ       := o_Json1.Get_String('state');
      v_��       := o_Json1.Get_String('city');
      v_��       := o_Json1.Get_String('county');
      v_����     := o_Json1.Get_String('township');
      v_����     := o_Json1.Get_String('other');
      v_�������� := o_Json1.Get_String('code');
      n_�Ƿ���� := o_Json1.Get_Number('visit_or_in');
    
      Zl_���˵�ַ��Ϣ_Update_s(n_��������, n_����id, n_��ҳid, n_��ַ����, v_ʡ, v_��, v_��, v_����, v_����, v_��������, n_�Ƿ����);
    End Loop;
  End If;
  --      item_list[]         ���²�����Ϣĳһ���ֶε�ֵ 
  --      item_name           C  1  �ֶ��� 
  --      item_value          C  1   �ֶ�ֵ 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('item_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1  := Pljson();
      o_Json1  := Pljson(j_Jsonlist.Get(I));
      v_�ֶ��� := o_Json1.Get_String('item_name');
      v_�ֶ�ֵ := o_Json1.Get_String('item_value');
      If Nvl(n_����id, 0) <> 0 And Nvl(v_�ֶ���, '-') <> '-' Then
        If Nvl(v_�ֶ�ֵ, '-') = 'Null' Then
          v_Sql := 'Update ������Ϣ Set ' || v_�ֶ��� || '=Null Where ����ID=:1';
          Execute Immediate v_Sql
            Using n_����id;
        Else
          v_Sql := 'Update ������Ϣ Set ' || v_�ֶ��� || '=:1 Where ����ID=:2';
          Execute Immediate v_Sql
            Using v_�ֶ�ֵ, n_����id;
        End If;
      End If;
    End Loop;
  End If;
  --���²��˴�����Ϣ 
  --    ext_list[]            C     ������Ϣ�����б� 
  --      info_name           C  1  ��Ϣ�� 
  --      upd_info_value      N  1  �޸ĵ���Ϣֵ 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('ext_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1  := Pljson();
      o_Json1  := Pljson(j_Jsonlist.Get(I));
      v_��Ϣ�� := o_Json1.Get_String('info_name');
      v_��Ϣֵ := o_Json1.Get_String('upd_info_value');
      n_����id := o_Json1.Get_Number('visit_id');
      If v_��Ϣ�� Is Not Null Then
        Zl_������Ϣ�ӱ�_Update(n_����id, v_��Ϣ��, v_��Ϣֵ, n_����id);
      End If;
    End Loop;
  End If;

  --����֤������ 
  --    cert_list[]                 ֤���б�(��Ҫ�ǵ��ɰ󿨴���) 
  --      cert_name           C  1  ֤������ 
  --      cert_no             C  1  ֤�ź��� 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('cert_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1  := Pljson();
      o_Json1  := Pljson(j_Jsonlist.Get(I));
      v_������ := o_Json1.Get_String('cert_name');
      v_����   := o_Json1.Get_String('cert_no');
    
      If v_������ Is Not Null Then
        If v_���� Is Not Null Then
          --��鿨���Ƿ�����ʹ�� 
          Select Count(1)
          Into n_Count
          From ����ҽ�ƿ���Ϣ A, ҽ�ƿ���� B
          Where a.�����id = b.Id And b.���� = v_������ And b.�Ƿ�֤�� = 1 And a.���� = v_���� And a.����id <> n_����id;
          If n_Count <> 0 Then
            Json_Out := Zljsonout(v_������ || ':' || v_���� || '���ڱ�����ʹ��,���飡');
            Return;
          End If;
        
          --�����ڵľ���������Ҫ������������ 
          Select Nvl(Max(ID), 0), Nvl(Max(���ų���), 0), Max(����), Max(LPad(����, 10)), Max(Length(����))
          Into n_�����id, n_���ų���, v_����, n_���ֵ, n_�ֵ
          From ҽ�ƿ����
          Where ���� = v_������;
        
          Select Max(����), Max(LPad(����, 10)), Max(Length(����)) Into v_����, n_���ֵ, n_�ֵ From ҽ�ƿ����;
        
          If v_���� Is Null Then
            Select LPad(1, 10, '0') Into v_���� From Dual;
          Else
            n_���ֵ := n_���ֵ + 1;
            Select LPad(n_���ֵ, n_�ֵ, '0') Into v_���� From Dual;
          End If;
        
          If n_�����id = 0 Then
            --���� 
            Select ҽ�ƿ����_Id.Nextval Into n_�����id From Dual;
          
            Zl_ҽ�ƿ����_Update(n_�����id, v_����, v_������, Substr(v_������, 1, 1), Null, Length(v_����), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_����, 0, Null, 1, Null, 1, 10, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          Elsif Length(v_����) > n_���ų��� Then
            --�޸ĳ��� 
            Zl_ҽ�ƿ����_Update(n_�����id, v_����, v_������, Substr(v_������, 1, 1), Null, Length(v_����), 0, 1, 0, 0, 0, 0, Null, Null,
                            v_����, 0, Null, 1, Null, 1, 10, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1000, 0, 1);
          End If;
        End If;
      
        --��������˿���Ϣ 
        n_Count := 0;
        For c_֤�� In (Select a.�����id, a.����
                     From ����ҽ�ƿ���Ϣ A
                     Where a.�����id = n_�����id And a.����id = n_����id) Loop
          If c_֤��.���� = Nvl(v_����, '_') Then
            n_Count := 1;
          Else
            Zl_ҽ�ƿ��䶯_Insert_s(14, n_����id, c_֤��.�����id, Null, c_֤��.����, '֤����ȡ����', Null, v_����Ա����, d_�Ǽ�ʱ��);
          End If;
        End Loop;
        --�������˿���Ϣ 
        If n_Count = 0 And v_���� Is Not Null Then
          Zl_ҽ�ƿ��䶯_Insert_s(11, n_����id, n_�����id, Null, v_����, '֤������', Null, v_����Ա����, d_�Ǽ�ʱ��);
        End If;
      End If;
    End Loop;
  End If;

  --���¹������� 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('allergic_drugs_list');
  If j_Jsonlist Is Not Null Then
    --������м�¼ 
    If Nvl(n_����ҩ�����, 0) = 0 Then
      Zl_���˹���ҩ��_Delete(n_����id);
    End If;
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1        := Pljson();
      o_Json1        := Pljson(j_Jsonlist.Get(I));
      n_��ʽ         := o_Json1.Get_Number('oper_type');
      n_����ҩƷid   := o_Json1.Get_Number('pat_algc_cadn_id');
      v_����ҩ������ := o_Json1.Get_String('pat_algc_cadn');
      v_��ÿҩ�ﷴӦ := o_Json1.Get_String('allergy_info');
      If o_Json1.Get_Clob('allergic_drugs') Is Not Null Then
        c_����ҩ�� := o_Json1.Get_Clob('allergic_drugs');
      End If;
      If Nvl(n_��ʽ, 0) = 0 Then
        If v_����ҩ������ Is Not Null Then
          If n_����ҩƷid = 0 Then
            n_����ҩƷid := Null;
          End If;
          Zl_���˹���ҩ��_Update(n_����id, n_����ҩƷid, v_����ҩ������, v_��ÿҩ�ﷴӦ);
        End If;
      End If;
      If Nvl(n_��ʽ, 0) = 1 Then
        While c_����ҩ�� Is Not Null Loop
          If Length(c_����ҩ��) <= 4000 Then
            l_����ҩ��.Extend;
            l_����ҩ��(l_����ҩ��.Count) := c_����ҩ��;
            c_����ҩ�� := Null;
          Else
            l_����ҩ��.Extend;
            l_����ҩ��(l_����ҩ��.Count) := Substr(c_����ҩ��, 1, Instr(c_����ҩ��, ',', 3980) - 1);
            c_����ҩ�� := Substr(c_����ҩ��, Instr(c_����ҩ��, ',', 3980) + 1);
          End If;
        End Loop;
        For I In 1 .. l_����ҩ��.Count Loop
          Delete From ���˹���ҩ��
          Where ����id = n_����id And
                (����ҩ��id, ����ҩ��) Not In
                (Select Distinct C1 As ����ҩ��id, C2 As ����ҩ�� From Table(f_Str2list2(l_����ҩ��(I), ',')));
        End Loop;
        If l_����ҩ��.Count = 0 Then
          Delete From ���˹���ҩ�� Where ����id = n_����id;
        End If;
      End If;
    End Loop;
  End If;

  --�������߼�¼ 
  --    immune_list[]         C     ���������б� 
  --      vaccinate_time      C  1  ����ʱ��:yyyy-mm-dd hh24:mi:ss 
  --      vaccinate_name      C  1  �������� 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('immune_list');
  If j_Jsonlist Is Not Null Then
    --������м�¼ 
    Zl_�������߼�¼_Delete(n_����id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1    := Pljson();
      o_Json1    := Pljson(j_Jsonlist.Get(I));
      d_����ʱ�� := To_Date(o_Json1.Get_String('vaccinate_time'), 'YYYY-MM-DD hh24:mi:ss');
      v_�������� := o_Json1.Get_String('vaccinate_name');
    
      If v_�������� Is Not Null Then
        Zl_�������߼�¼_Update(n_����id, d_����ʱ��, v_��������);
      End If;
    End Loop;
  End If;

  --����ҽ�ƿ����� 
  --    card_property_list[]  C     ҽ�ƿ������б� 
  --      cardtype_id         N  1  ҽ�ƿ����ID 
  --      card_no             C  1  ���� 
  --      info_name           C  1  ��Ϣ�� 
  --      info_value          N  1  ��Ϣֵ 
  j_Jsonlist := Pljson_List();
  j_Jsonlist := o_Json.Get_Pljson_List('card_property_list');
  If j_Jsonlist Is Not Null Then
    --������м�¼ 
    Zl_�������߼�¼_Delete(n_����id);
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json1    := Pljson();
      o_Json1    := Pljson(j_Jsonlist.Get(I));
      n_�����id := o_Json1.Get_Number('cardtype_id');
      v_����     := o_Json1.Get_String('card_no');
      v_��Ϣ��   := o_Json1.Get_String('info_name');
      v_��Ϣֵ   := o_Json1.Get_String('info_value');
    
      Zl_����ҽ�ƿ�����_Update(n_����id, n_�����id, v_����, v_��Ϣ��, v_��Ϣֵ);
    End Loop;
  End If;

  --ǩԼ��Ϣ 
  --    sign_info             C   ǩԼ��Ϣ 
  --      card_type_id        N 1 �����ID 
  --      card_no             C 1 ���� 
  --      card_pwd            C   ������ 
  --      qrcode              C   ��ά�� 
  --      card_notes          C   �䶯ԭ�� 
  --      card_use_endtime    C   ��ֹʹ��ʱ�� 
  o_Json1 := Pljson();
  o_Json1 := o_Json.Get_Pljson('sign_info');
  If o_Json1 Is Not Null Then
    n_�����id     := o_Json1.Get_Number('card_type_id');
    v_����         := o_Json1.Get_String('card_no');
    v_������       := o_Json1.Get_String('card_pwd');
    v_�䶯ԭ��     := o_Json1.Get_String('card_notes');
    d_��ֹʹ��ʱ�� := To_Date(o_Json1.Get_String('card_use_endtime'), 'YYYY-MM-DD hh24:mi:ss');
    --ǩԼ 
    Select Count(1) Into n_Count From ҽ�ƿ���� Where ID = n_�����id;
    If n_Count = 1 Then
      Select Count(1) Into n_Count From ����ҽ�ƿ���Ϣ Where ���� = v_���� And �����id = n_�����id;
      If n_Count = 0 Then
        Zl_ҽ�ƿ��䶯_Insert_s(11, n_����id, n_�����id, '', v_����, v_�䶯ԭ��, v_������, v_����Ա����, d_�Ǽ�ʱ��, Null, d_��ֹʹ��ʱ��);
      End If;
    End If;
  End If;
  b_Message.Zlhis_Patient_016(n_����id);

  Json_Out := Zljsonout('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatepatiarchives;
/
Create Or Replace Procedure Zl_Patisvr_Updatepatibaseinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���²��˻�����Ϣ
  --��Σ�JSON��ʽ
  --input
  --  pati_id               N 1 ����id
  --  visit_id              N 1 ����id
  --  model                 N 1 ģ��
  --  pati_name_n           C 1 ����
  --  pati_sex_n            C 1 �Ա�
  --  pati_age_n            C 1 ����
  --  pati_birthdate_n      C 1 ��������
  --  occasion              N 1 ���� 1-����;2-סԺ
  --  pati_name_o           C 1 ����
  --  pati_sex_o            C 1 �Ա�
  --  pati_age_o            C 1 ����
  --  pati_birthdate_o      C 1 ��������
  --  explain               C 1 ˵��
  --���Σ�JSON��ʽ
  --output
  --   code                 N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message              C  1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Json        Pljson;
  j_Jsonin      Pljson;
  v_Username    ��Ա��.����%Type;
  d_�䶯ʱ��    ������Ϣ�䶯.�䶯ʱ��%Type;
  v_˵��        ������Ϣ�䶯.˵��%Type;
  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
  n_����id      ������Ϣ�䶯.����id%Type;
  v_ģ��        ������Ϣ�䶯.�䶯ģ��%Type;
  v_����_n      ������Ϣ.����%Type;
  v_�Ա�_n      ������Ϣ.�Ա�%Type;
  v_����_n      ������Ϣ.����%Type;
  d_��������_n  ������Ϣ.��������%Type;
  v_����_o      ������Ϣ.����%Type;
  v_�Ա�_o      ������Ϣ.�Ա�%Type;
  v_����_o      ������Ϣ.����%Type;
  d_��������_o  ������Ϣ.��������%Type;
Begin
  j_Jsonin     := Pljson(Json_In);
  j_Json       := j_Jsonin.Get_Pljson('input');
  n_����id     := j_Json.Get_Number('pati_id');
  v_ģ��       := j_Json.Get_String('model');
  v_����_n     := j_Json.Get_String('pati_name_n');
  v_�Ա�_n     := j_Json.Get_String('pati_sex_n');
  v_����_n     := j_Json.Get_String('pati_age_n');
  d_��������_n := To_Date(j_Json.Get_String('pati_birthdate_n'), 'yyyy-mm-dd hh24:mi:ss');
  v_����_o     := j_Json.Get_String('pati_name_o');
  v_�Ա�_o     := j_Json.Get_String('pati_sex_o');
  v_����_o     := j_Json.Get_String('pati_age_o');
  d_��������_o := To_Date(j_Json.Get_String('pati_birthdate_o'), 'yyyy-mm-dd hh24:mi:ss');
  v_˵��       := j_Json.Get_String('explain');
  v_Username   := zl_UserName;
  --3����첿��
  --��첿�ֲ������ӹ���,�����ϵͳ�ľ����¼�����ϵͳ�Լ�����,���Դ˴�����n_����id�޷������ϵͳ����������
  --���ϵͳ�ṩ�������޸���ڡ�
  --4��PACS����
  Select Zl_Fun_Checkidentify(0, n_����id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
  Update ������Ϣ
  Set ���� = v_����_n, �Ա� = v_�Ա�_n, ���� = v_����_n, �������� = d_��������_n
  Where ����id = n_����id;
  Select Zl_Fun_Checkidentify(1, n_����id, v_Strtmpbefor) Into v_Msg From Dual;

  d_�䶯ʱ�� := Sysdate;
  If Nvl(v_����_n, '_') <> Nvl(v_����_o, '_') Then
    Insert Into ������Ϣ�䶯
      (����id, �䶯��Ŀ, ԭ��Ϣ, ����Ϣ, �䶯ʱ��, �䶯��, �䶯ģ��, ˵��)
    Values
      (n_����id, '����', v_����_o, v_����_n, d_�䶯ʱ��, v_Username, v_ģ��, v_˵��);
  End If;
  If Nvl(v_�Ա�_n, '_') <> Nvl(v_�Ա�_o, '_') Then
    Insert Into ������Ϣ�䶯
      (����id, �䶯��Ŀ, ԭ��Ϣ, ����Ϣ, �䶯ʱ��, �䶯��, �䶯ģ��, ˵��)
    Values
      (n_����id, '�Ա�', v_�Ա�_o, v_�Ա�_n, d_�䶯ʱ��, v_Username, v_ģ��, v_˵��);
  End If;
  If Nvl(v_����_n, '_') <> Nvl(v_����_o, '_') Then
    Insert Into ������Ϣ�䶯
      (����id, �䶯��Ŀ, ԭ��Ϣ, ����Ϣ, �䶯ʱ��, �䶯��, �䶯ģ��, ˵��)
    Values
      (n_����id, '����', v_����_o, v_����_n, d_�䶯ʱ��, v_Username, v_ģ��, v_˵��);
  End If;
  If Nvl(d_��������_n, Sysdate) <> Nvl(d_��������_o, Sysdate) Then
    Insert Into ������Ϣ�䶯
      (����id, �䶯��Ŀ, ԭ��Ϣ, ����Ϣ, �䶯ʱ��, �䶯��, �䶯ģ��, ˵��)
    Values
      (n_����id, '��������', To_Char(d_��������_o, 'YYYY-MM-DD hh24:mi'), To_Char(d_��������_n, 'YYYY-MM-DD hh24:mi'), d_�䶯ʱ��,
       v_Username, v_ģ��, v_˵��);
  End If;
  b_Message.Zlhis_Patient_016(n_����id);
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatepatibaseinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updatepatirelate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ�����˵���Ϣ���й���
  --��Σ�Json_In:��ʽ
  --  input
  --    oper_fun            N   1 ��������:0-���ӹ���;1-ȡ������;2-���¹���ID;3-��Ժ�Ǽ��Զ�����
  --    relate_id           N     ����ID
  --    relate_pati_ids     C   1 ��Ҫ�����Ĳ���ids:����ö���
  --    operator_name       C   1 ����Ա����
  --    operator_time       C   1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ----------------------------------------------------------------------------

  v_����ids  Varchar2(32680);
  n_�������� Number(1);
  n_����id   ������ݹ���.����id%Type;
  v_����Ա   ������ݹ���.������Ա%Type;
  d_����ʱ�� ������ݹ���.����ʱ��%Type;
  j_Json     Pljson;
  j_Jsonin   Pljson;
Begin

  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_�������� := j_Json.Get_Number('oper_fun');
  n_����id   := j_Json.Get_Number('relate_id');
  v_����ids  := j_Json.Get_String('relate_pati_ids');
  v_����Ա   := j_Json.Get_String('operator_name');
  d_����ʱ�� := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  Zl_������ݹ���_Update(n_��������, n_����id, v_����ids, v_����Ա, d_����ʱ��);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updatepatirelate;
/
Create Or Replace Procedure Zl_Patisvr_Updateproxy
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���������Ϣ����
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id               N 1 ����id_In
  --  visit_id              N 1 ����ID
  --  pati_idcard           C 1 ���֤��
  --  proxy_name            C 1 ����������
  --  proxy_idno            C 1 ���������֤��
  --  proxy_sex             C 1 �������Ա�
  --  pati_age              C 1 ����������
  --  proxy_phno            C 1 �����˵绰
  --  reason                C 1 ��ҩ����
  --����: Json_Out,��ʽ����
  --output
  --  code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --  message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  j_Json   Pljson;
  j_Jsonin Pljson;
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');
  Zl_��������Ϣ_Insert(j_Json.Get_Number('pati_id'), j_Json.Get_String('pati_idcard'), j_Json.Get_String('proxy_name'),
                  j_Json.Get_String('proxy_idno'), j_Json.Get_Number('visit_id'), j_Json.Get_String('proxy_sex'),
                  j_Json.Get_String('pati_age'), j_Json.Get_String('proxy_phno'), j_Json.Get_String('reason'));
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updateproxy;
/
Create Or Replace Procedure Zl_Patisvr_Updcommunityinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���������ҽ��վ���������˽��в��������֤ʱʹ��
  --���:Json_In:��ʽ
  --input
  --    pati_id               N  1 ����ID
  --    community_num         N  1  �������
  --    community_code        C  1  ��������
  --    community_oper_type   N  1  ������������
  --    visit_time            C  1  ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --    pati_name             C  1  ����
  --    pati_sex              C  1  �Ա�
  --    pati_age              C  1  ����
  --    pati_birthdate        C  1  ��������:yyyy-mm-dd hh24:mi:ss
  --    pat_baddr             C  1  �����ص�
  --    pati_idcard           C  1  ���֤��
  --    nation_name           C  1  ����
  --    country_name          C  1  ����
  --    mari_name             C  1  ����״��
  --    ocpt_name             C  1  ְҵ
  --    pat_home_addr         C  1  ��ͥ��ַ
  --    pat_home_phno         C  1  ��ͥ�绰
  --    pat_home_postcode     C  1  ��ͥ��ַ�ʱ�
  --    emp_name              C  1  ������λ
  --    emp_phno              C  1  ��λ�绰
  --    emp_postcode          C  1  ��λ�ʱ�
  --    contacts_name         C  1  ��ϵ������
  --    contacts_relation     C  1  ��ϵ�˹�ϵ
  --    ontacts_phno          C  1  ��ϵ�˵绰
  --    ontacts_addr          C  1  ��ϵ�˵�ַ
  --    pat_hous_addr         C  1  ���ڵ�ַ
  --    pat_hous_postcode     C  1  ���ڵ�ַ�ʱ�

  -- ����:
  --  output
  --    code                            N 1 Ӧ����:0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ:ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------

  n_����id       Number;
  n_�������     Number;
  v_��������     Varchar2(20);
  n_������������ Number;
  d_����ʱ��     Date;
  v_����         ������Ϣ.����%Type;
  v_�Ա�         ������Ϣ.�Ա�%Type;
  v_����         ������Ϣ.����%Type;
  d_��������     ������Ϣ.��������%Type;
  v_�����ص�     ������Ϣ.�����ص�%Type;
  v_���֤��     ������Ϣ.���֤��%Type;
  v_����         ������Ϣ.����%Type;
  v_����         ������Ϣ.����%Type;
  v_����״��     ������Ϣ.����״��%Type;
  v_ְҵ         ������Ϣ.ְҵ%Type;
  v_��ͥ��ַ     ������Ϣ.��ͥ��ַ%Type;
  v_��ͥ�绰     ������Ϣ.��ͥ�绰%Type;
  v_��ͥ��ַ�ʱ� ������Ϣ.��ͥ��ַ�ʱ�%Type;
  v_������λ     ������Ϣ.������λ%Type;
  v_��λ�绰     ������Ϣ.��λ�绰%Type;
  v_��λ�ʱ�     ������Ϣ.��λ�ʱ�%Type;
  v_��ϵ������   ������Ϣ.��ϵ������%Type;
  v_��ϵ�˹�ϵ   ������Ϣ.��ϵ�˹�ϵ%Type;
  v_��ϵ�˵绰   ������Ϣ.��ϵ�˵绰%Type;
  v_��ϵ�˵�ַ   ������Ϣ.��ϵ�˵�ַ%Type;
  v_���ڵ�ַ     ������Ϣ.���ڵ�ַ%Type;
  v_���ڵ�ַ�ʱ� ������Ϣ.���ڵ�ַ�ʱ�%Type;

  j_Json   Pljson;
  j_Jsonin Pljson;

  v_Strtmpbefor Varchar2(4000);
  v_Msg         Varchar2(4000);
Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  n_����id       := j_Json.Get_Number('pati_id');
  n_�������     := j_Json.Get_Number('community_num');
  v_��������     := j_Json.Get_String('community_code');
  n_������������ := j_Json.Get_Number('community_oper_type');
  d_����ʱ��     := To_Date(j_Json.Get_String('visit_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_����         := j_Json.Get_String('pati_name');
  v_�Ա�         := j_Json.Get_String('pati_sex');
  v_����         := j_Json.Get_String('pati_age');
  d_��������     := To_Date(j_Json.Get_String('pati_birthdate'), 'YYYY-MM-DD hh24:mi:ss');
  v_�����ص�     := j_Json.Get_String('pat_baddr');
  v_���֤��     := j_Json.Get_String('pati_idcard');
  v_����         := j_Json.Get_String('nation_name');
  v_����         := j_Json.Get_String('country_name');
  v_����״��     := j_Json.Get_String('mari_name');
  v_ְҵ         := j_Json.Get_String('ocpt_name');
  v_��ͥ��ַ     := j_Json.Get_String('pat_home_addr');
  v_��ͥ�绰     := j_Json.Get_String('pat_home_phno');
  v_��ͥ��ַ�ʱ� := j_Json.Get_String('pat_home_postcode');
  v_������λ     := j_Json.Get_String('emp_name');
  v_��λ�绰     := j_Json.Get_String('emp_phno');
  v_��λ�ʱ�     := j_Json.Get_String('emp_postcode');
  v_��ϵ������   := j_Json.Get_String('contacts_name');
  v_��ϵ�˹�ϵ   := j_Json.Get_String('contacts_relation');
  v_��ϵ�˵绰   := j_Json.Get_String('ontacts_phno');
  v_��ϵ�˵�ַ   := j_Json.Get_String('ontacts_addr');
  v_���ڵ�ַ     := j_Json.Get_String('pat_hous_addr');
  v_���ڵ�ַ�ʱ� := j_Json.Get_String('pat_hous_postcode');

  If d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;

  Zl_����������Ϣ_Insert(n_����id, n_�������, v_��������, n_������������, d_����ʱ��);

  Select Zl_Fun_Checkidentify(0, n_����id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
  Update ������Ϣ
  Set ���� = Decode(v_����, Null, ����, v_����), �Ա� = Decode(v_�Ա�, Null, �Ա�, v_�Ա�), ���� = Decode(v_����, Null, ����, v_����),
      �������� = Decode(d_��������, Null, ��������, d_��������), �����ص� = Decode(v_�����ص�, Null, �����ص�, v_�����ص�),
      ���֤�� = Decode(v_���֤��, Null, ���֤��, v_���֤��), ���� = Decode(v_����, Null, ����, v_����), ���� = Decode(v_����, Null, ����, v_����),
      ����״�� = Decode(v_����״��, Null, ����״��, v_����״��), ְҵ = Decode(v_ְҵ, Null, ְҵ, v_ְҵ),
      ��ͥ��ַ = Decode(v_��ͥ��ַ, Null, ��ͥ��ַ, v_��ͥ��ַ), ��ͥ�绰 = Decode(v_��ͥ�绰, Null, ��ͥ�绰, v_��ͥ�绰),
      ��ͥ��ַ�ʱ� = Decode(v_��ͥ��ַ�ʱ�, Null, ��ͥ��ַ�ʱ�, v_��ͥ��ַ�ʱ�), ������λ = Decode(v_������λ, Null, ������λ, v_������λ),
      ��λ�绰 = Decode(v_��λ�绰, Null, ��λ�绰, v_��λ�绰), ��λ�ʱ� = Decode(v_��λ�ʱ�, Null, ��λ�ʱ�, v_��λ�ʱ�),
      ��ϵ������ = Decode(v_��ϵ������, Null, ��ϵ������, v_��ϵ������), ��ϵ�˹�ϵ = Decode(v_��ϵ������, Null, ��ϵ�˹�ϵ, v_��ϵ�˹�ϵ),
      ��ϵ�˵绰 = Decode(v_��ϵ������, Null, ��ϵ�˵绰, v_��ϵ�˵绰), ��ϵ�˵�ַ = Decode(v_��ϵ������, Null, ��ϵ�˵�ַ, v_��ϵ�˵�ַ),
      ���ڵ�ַ = Decode(v_���ڵ�ַ, Null, ���ڵ�ַ, v_���ڵ�ַ), ���ڵ�ַ�ʱ� = Decode(v_���ڵ�ַ�ʱ�, Null, ���ڵ�ַ�ʱ�, v_���ڵ�ַ�ʱ�)
  Where ����id = n_����id;
  Select Zl_Fun_Checkidentify(1, n_����id, v_Strtmpbefor) Into v_Msg From Dual;

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updcommunityinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updpatiaddressinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�޸Ĳ�����ҳ�ӱ������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      oper_fun          N 1 ��������:1-����,�޸�   2-ɾ��
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳId
  --      pat_addr_type     C 1 ��ַ���
  --      pat_addr_state    C 1 ��ַ_ʡ
  --      pat_addr_city     C 1 ��ַ_��
  --      pat_addr_county   C 1 ��ַ_��
  --      pat_addr_township C 1 ��ַ_��
  --      pat_addr_other    C 1 ��ַ_����
  --      pat_region_code   C 1 ��������

  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Json     Pljson;
  j_Jsonin   Pljson;
  n_����id   ������Ϣ.����id%Type;
  n_��ҳid   ������Ϣ.��ҳid%Type;
  n_����     Number(3);
  v_��ַ��� ���˵�ַ��Ϣ.��ַ���%Type;
  v_ʡ       ���˵�ַ��Ϣ.ʡ%Type;
  v_��       ���˵�ַ��Ϣ.��%Type;
  v_��       ���˵�ַ��Ϣ.��%Type;
  v_����     ���˵�ַ��Ϣ.����%Type;
  v_����     ���˵�ַ��Ϣ.����%Type;
  v_�������� ���˵�ַ��Ϣ.��������%Type;

Begin
  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����     := j_Json.Get_Number('oper_fun');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  v_��ַ��� := j_Json.Get_String('pat_addr_type');

  v_ʡ       := j_Json.Get_String('pat_addr_state');
  v_��       := j_Json.Get_String('pat_addr_city');
  v_��       := j_Json.Get_String('pat_addr_county');
  v_����     := j_Json.Get_String('pat_addr_township');
  v_����     := j_Json.Get_String('pat_addr_other');
  v_�������� := j_Json.Get_String('pat_region_code');

  Zl_���˵�ַ��Ϣ_Update_s(n_����, n_����id, n_��ҳid, v_��ַ���, v_ʡ, v_��, v_��, v_����, v_����, v_��������);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updpatiaddressinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updpatiallerinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����ڸ��²�����Ϣ�Ĺ���ҩ����Ϣ
  --���:Json_In:��ʽ
  --input
  --    aller_list    ���˹���ҩ���б�
  --      execute_type   N 1 ִ�з�ʽ 1-ɾ�� 2-�������߸���
  --      pati_id        N 1 ����id
  --      drug_id        N 1 ҩ��id
  --      drug_name      C 1 ҩ����
  --      aller_reflex   C 1 ������Ӧ

  -- ����:
  --  output
  --    code                            N 1 Ӧ����:0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ:ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------

  n_����id   Number;
  n_Type     Number;
  n_ҩ��id   ���˹���ҩ��.����ҩ��id%Type;
  v_����ҩ�� ���˹���ҩ��.����ҩ��%Type;
  v_������Ӧ ���˹���ҩ��.������Ӧ%Type;
  o_Json     Pljson;

  j_Json           Pljson;
  j_Jsonin         Pljson;
  j_Json_Allerlist Pljson_List := Pljson_List();

Begin
  --�������
  j_Jsonin := Pljson(Json_In);
  j_Json   := j_Jsonin.Get_Pljson('input');

  j_Json_Allerlist := j_Json.Get_Pljson_List('aller_list');
  For I In 1 .. j_Json_Allerlist.Count Loop
    o_Json     := Pljson();
    o_Json     := Pljson(j_Json_Allerlist.Get(I));
    n_Type     := o_Json.Get_Number('execute_type');
    n_����id   := o_Json.Get_Number('pati_id');
    n_ҩ��id   := o_Json.Get_Number('drug_id');
    v_����ҩ�� := o_Json.Get_String('drug_name');
    v_������Ӧ := o_Json.Get_String('aller_reflex');
    If Nvl(n_Type, 0) = 1 Then
      --���û�й����ļ�¼��ɾ����ҩƷ�Ĺ�����¼
      Delete From ���˹���ҩ�� A Where a.����id = n_����id And a.����ҩ�� = v_����ҩ�� And a.����ҩ��id = n_ҩ��id;
    Else
      Update ���˹���ҩ��
      Set ������Ӧ = v_������Ӧ, ����ҩ��id = n_ҩ��id
      Where ����id = n_����id And ����ҩ�� = v_����ҩ��;
      If Sql%RowCount = 0 Then
        Insert Into ���˹���ҩ��
          (����id, ����ҩ��id, ����ҩ��, ������Ӧ)
        Values
          (n_����id, n_ҩ��id, v_����ҩ��, v_������Ӧ);
      End If;
    End If;
  End Loop;

  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updpatiallerinfo;
/
Create Or Replace Procedure Zl_Patisvr_Updpatifamilyinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����²��˼�����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --  opr_fun      N  1  ������ʽ  --1-����,2-����,3-��ɾ��
  --  pati_id      N  1  ����ID
  --  family_id    N  1  ����id
  --  reg_name     C  1  �Ǽ���
  --  reg_time     C  1  �Ǽ�ʱ��
  --  relation     C  0  ��ϵ
  --  cancel_name  C  0  ������
  --  cancel_time  C  0  ����ʱ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  n_����     Number;
  n_����id   ���˼���.����id%Type;
  n_����id   ���˼���.����id%Type;
  v_�Ǽ���   ���˼���.�Ǽ���%Type;
  d_�Ǽ�ʱ�� ���˼���.�Ǽ�ʱ��%Type;
  v_��ϵ     ���˼���.��ϵ%Type;
  v_������   ���˼���.������%Type;
  v_����ʱ�� ���˼���.����ʱ��%Type;
  j_Json     Pljson;
  j_Jsonin   Pljson;
Begin

  --�������
  j_Jsonin   := Pljson(Json_In);
  j_Json     := j_Jsonin.Get_Pljson('input');
  n_����     := j_Json.Get_Number('opr_fun');
  n_����id   := j_Json.Get_Number('pati_id');
  n_����id   := j_Json.Get_Number('family_id');
  v_�Ǽ���   := j_Json.Get_String('reg_name');
  d_�Ǽ�ʱ�� := To_Date(j_Json.Get_String('reg_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_��ϵ     := j_Json.Get_String('relation');
  v_������   := j_Json.Get_String('cancel_name');
  v_����ʱ�� := To_Date(j_Json.Get_String('cancel_time'), 'yyyy-mm-dd hh24:mi:ss');
  Zl_���˼���_Update(n_����, n_����id, n_����id, v_�Ǽ���, d_�Ǽ�ʱ��, v_��ϵ, v_������, v_����ʱ��);
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Patisvr_Updpatifamilyinfo;
/