Create Or Replace Procedure Zl_Exsesvr_Addeinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ӵ���Ʊ����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    balance_id          N  1  ����ID
  --    balance_delid       N     �˿�ID:�˿�ߺ�Ʊʱ��Ч��Ŀǰֻ��Ԥ������Ч,��д�����˿�Ԥ��ID
  --    einvoice_id         N  1  ����Ʊ��ID
  --    operator_code       C  1  ����Ա���
  --    operator_name       C  1  ����Ա����
  --    happen_time         C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    pati_info           C     ������Ϣ
  --      pati_id           N  1  ����ID
  --      pati_pageid       N     ��ҳID
  --      pati_name         C  1  ����
  --      pati_sex          C  1  �Ա�
  --      pati_age          C  1  ����
  --      outpatient_num    C  1  �����
  --      inpatient_num     C  1  סԺ��
  --    einvoce_info        C     ����Ʊ����Ϣ
  --      invoice_type      N  1  Ʊ�֣�1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
  --      placeCode         C  1  ��Ʊ�����
  --      inv_total         N  1  ��Ʊ���
  --      inv_oldid         N     ԭƱ��ID
  --      sys_source        C  1  ϵͳ��Դ
  --      demo              C  1  ��ע
  --      einvoice_code     C  1  ����Ʊ�ݴ���
  --      einvoice_no       C  1  ����Ʊ�ݺ���
  --      einvoice_random   C  1  ����У����
  --      voucher_code      C  1  Ԥ����ƾ֤����
  --      voucher_no        C  1  Ԥ����ƾ֤����
  --      voucher_random    C  1  Ԥ����ƾ֤У����
  --      create_time       C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url       C  1  ����Ʊ��H5ҳ��URL
  --      picture_neturl    C  1  ����Ʊ������H5ҳ��URL
  --      qrcode            C  1  ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --    --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_Id         ����Ʊ��ʹ�ü�¼.Id%Type;
  n_Ʊ��       ����Ʊ��ʹ�ü�¼.Ʊ��%Type;
  n_��¼״̬   ����Ʊ��ʹ�ü�¼.��¼״̬%Type;
  n_����id     ����Ʊ��ʹ�ü�¼.����id%Type;
  n_����id     ����Ʊ��ʹ�ü�¼.����id%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_�Ա�       ����Ʊ��ʹ�ü�¼.�Ա�%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  n_�����     ����Ʊ��ʹ�ü�¼.�����%Type;
  n_סԺ��     ����Ʊ��ʹ�ü�¼.סԺ��%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_������     ����Ʊ��ʹ�ü�¼.������%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤������ ����Ʊ��ʹ�ü�¼.ƾ֤������%Type;
  n_Ʊ�ݽ��   ����Ʊ��ʹ�ü�¼.Ʊ�ݽ��%Type;
  v_����ʱ��   ����Ʊ��ʹ�ü�¼.����ʱ��%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  c_��ά��     Clob;
  n_ԭƱ��id   ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type;
  n_�˿�id     ����Ʊ��ʹ�ü�¼.�˿�id%Type;
  v_��ע       ����Ʊ��ʹ�ü�¼.��ע%Type;
  v_��Ʊ��     ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
  v_ϵͳ��Դ   ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type;
  v_����Ա��� ����Ʊ��ʹ�ü�¼.����Ա���%Type;
  v_����Ա���� ����Ʊ��ʹ�ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type;

  n_��¼״̬ Number(2);
  j_Input    PLJson;
  j_Json     PLJson;
  j_Temp     PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := j_Json.Get_Number('balance_id');
  n_�˿�id     := j_Json.Get_Number('balance_delid');
  n_Id         := j_Json.Get_Number('einvoice_id');
  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  d_�Ǽ�ʱ��   := Nvl(To_Date(j_Json.Get_String('happen_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  --��ȡ������Ϣ

  If Not j_Json.Exist('pati_info') Then
  
    Json_Out := zlJsonOut('�޲�����Ϣ���������ӵ���Ʊ�ݡ�');
    Return;
  End If;

  j_Temp   := j_Json.Get_Pljson('pati_info');
  n_����id := j_Temp.Get_Number('pati_id');
  --n_��ҳid := j_Temp.Get_Number('pati_pageid');

  v_����   := j_Temp.Get_String('pati_name');
  v_�Ա�   := j_Temp.Get_String('pati_sex');
  v_����   := j_Temp.Get_String('pati_age');
  n_����� := j_Temp.Get_Number('outpatient_num');
  n_סԺ�� := j_Temp.Get_Number('inpatient_num');

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('�޵���Ʊ����Ϣ,�������ӵ���Ʊ�ݡ�');
    Return;
  End If;

  --��ȡ����Ʊ����Ϣ
  j_Temp := PLJson();
  j_Temp := j_Json.Get_Pljson('einvoce_info');
  --Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  n_Ʊ��       := Nvl(j_Temp.Get_Number('invoice_type'), 1);
  v_��Ʊ��     := j_Temp.Get_String('placeCode');
  n_ԭƱ��id   := j_Temp.Get_Number('inv_oldid');
  v_ϵͳ��Դ   := j_Temp.Get_String('sys_source');
  v_��ע       := j_Temp.Get_String('demo');
  v_����       := j_Temp.Get_String('einvoice_code');
  v_����       := j_Temp.Get_String('einvoice_no');
  v_������     := j_Temp.Get_String('einvoice_random');
  v_ƾ֤����   := j_Temp.Get_String('voucher_code');
  v_ƾ֤����   := j_Temp.Get_String('voucher_no');
  v_ƾ֤������ := j_Temp.Get_String('voucher_random');
  n_Ʊ�ݽ��   := j_Temp.Get_Number('inv_total');
  v_����ʱ��   := j_Temp.Get_String('create_time');
  v_Url����    := j_Temp.Get_String('picture_url');
  v_Url����    := j_Temp.Get_String('picture_neturl');
  c_��ά��     := j_Temp.Get_Clob('qrcode');

  --���ӵ���Ʊ����Ϣ
  Zl_����Ʊ��ʹ�ü�¼_Insert(n_Id, n_Ʊ��, n_����id, n_����id, v_����, v_�Ա�, v_����, n_�����, n_סԺ��, n_Ʊ�ݽ��, v_��Ʊ��, v_ϵͳ��Դ, v_����ʱ��, v_��ע,
                     v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_ԭƱ��id, n_�˿�id, v_����, v_����, v_������, v_ƾ֤����, v_ƾ֤����, v_ƾ֤������, v_Url����,
                     v_Url����);
  --���¶�ά��
  Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (n_Id, c_��ά��);
  Json_Out := zlJsonOut('�ɹ�', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addeinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Deleinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ɾ������Ʊ����Ϣ
  --��Σ�Json_In:��ʽ
  -- input      
  --  einvoice_id  N  1  ����Ʊ��ID
  --  operator_code  C  1  ����Ա���
  --  operator_name  C  1  ����Ա����
  --  create_time  C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --  einvoce_info  C    ����Ʊ����Ϣ
  --    placeCode  C  1  ��Ʊ�����
  --    sys_source  C  1  ϵͳ��Դ
  --    demo  C  1  ��ע
  --    inv_oldid  N    ԭƱ��ID
  --    einvoice_code  C  1  ����Ʊ�ݴ���
  --    einvoice_no  C  1  ����Ʊ�ݺ���
  --    einvoice_random  C  1  ����У����
  --    voucher_code  C  1  Ԥ����ƾ֤����
  --    voucher_no  C  1  Ԥ����ƾ֤����
  --    voucher_random  C  1  Ԥ����ƾ֤У����
  --    happen_time  C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --    picture_url  C  1  ����Ʊ��H5ҳ��URL
  --    picture_neturl  C  1  ����Ʊ������H5ҳ��URL
  --    qrcode  C  1  ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --    --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_Id         ����Ʊ��ʹ�ü�¼.Id%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_������     ����Ʊ��ʹ�ü�¼.������%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤������ ����Ʊ��ʹ�ü�¼.ƾ֤������%Type;
  v_����ʱ��   ����Ʊ��ʹ�ü�¼.����ʱ��%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  c_��ά��     Clob;
  n_ԭƱ��id   ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type;
  v_��ע       ����Ʊ��ʹ�ü�¼.��ע%Type;
  v_��Ʊ��     ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
  v_ϵͳ��Դ   ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type;
  v_����Ա��� ����Ʊ��ʹ�ü�¼.����Ա���%Type;
  v_����Ա���� ����Ʊ��ʹ�ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type;

  j_Input PLJson;
  j_Json  PLJson;
  j_Temp  PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Id         := j_Json.Get_Number('einvoice_id');
  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  d_�Ǽ�ʱ��   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('einvoce_info') Then
  
    Json_Out := zlJsonOut('�޵���Ʊ����Ϣ,�������ӵ���Ʊ�ݡ�');
    Return;
  End If;

  --��ȡ����Ʊ����Ϣ
  j_Temp     := PLJson();
  j_Temp     := j_Json.Get_Pljson('einvoce_info');
  v_��Ʊ��   := j_Temp.Get_String('placeCode');
  v_ϵͳ��Դ := j_Temp.Get_String('sys_source');
  v_��ע     := j_Temp.Get_String('demo');
  n_ԭƱ��id := j_Temp.Get_Number('inv_oldid');

  v_����       := j_Temp.Get_String('einvoice_code');
  v_����       := j_Temp.Get_String('einvoice_no');
  v_������     := j_Temp.Get_String('einvoice_random');
  v_ƾ֤����   := j_Temp.Get_String('voucher_code');
  v_ƾ֤����   := j_Temp.Get_String('voucher_no');
  v_ƾ֤������ := j_Temp.Get_String('voucher_random');
  v_����ʱ��   := j_Temp.Get_String('happen_time');
  v_Url����    := j_Temp.Get_String('picture_url');
  v_Url����    := j_Temp.Get_String('picture_neturl');
  c_��ά��     := j_Temp.Get_Clob('qrcode');

  Zl_����Ʊ��ʹ�ü�¼_Delete(n_Id, v_��Ʊ��, v_ϵͳ��Դ, v_����ʱ��, v_��ע, v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_ԭƱ��id, v_����, v_����, v_������, v_ƾ֤����,
                     v_ƾ֤����, v_ƾ֤������, v_Url����, v_Url����);
  --���¶�ά��
  Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (n_Id, c_��ά��);
  Json_Out := zlJsonOut('�ɹ�', 1);
  Return;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deleinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Savepaperinvoice
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ֽ��Ʊ��ʹ����Ϣ
  --��Σ�Json_In:��ʽ
  --   input
  --    oper_mode           N  1  ������ʽ:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  --    einvoice_id         N  1  ����Ʊ��ID
  --    operator_code       C  1  ����Ա���
  --    operator_name       C  1  ����Ա����
  --    create_time         C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info       C     ֽ��Ʊ����Ϣ:���ڶ���ʱ���밴����˳���ϴ�(�������ݴ���)
  --      inv_occasion      N  1  Ӧ�ó���:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  --      invoice_type      N  1  Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      inv_red           N     �Ƿ��Ʊ:1-��Ʊ;0-�Ǻ�Ʊ
  --      invoice_no        C  1  ��Ʊ��
  --      inv_total         N  1  ��Ʊ���
  --      recv_id           N     ����id
  --    einvoce_info        C     ����Ʊ����Ϣ
  --      placeCode         C  1  ��Ʊ�����
  --      sys_source        C  1  ϵͳ��Դ
  --      demo              C  1  ��ע
  --      einvoice_id       N  1  ����Ʊ��ID(����)
  --      inv_oldid         N     ԭƱ��ID
  --      einvoice_code     C  1  ����Ʊ�ݴ���
  --      einvoice_no       C  1  ����Ʊ�ݺ���
  --      einvoice_random   C  1  ����У����
  --      voucher_code      C  1  Ԥ����ƾ֤����
  --      voucher_no        C  1  Ԥ����ƾ֤����
  --      voucher_random    C  1  Ԥ����ƾ֤У����
  --      happen_time       C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url       C  1  ����Ʊ��H5ҳ��URL
  --      picture_neturl    C  1  ����Ʊ������H5ҳ��URL
  --      qrcode            C  1  ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG

  --    --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_Id         ����Ʊ��ʹ�ü�¼.Id%Type;
  n_����id     ����Ʊ��ʹ�ü�¼.Id%Type;
  v_����Ա��� ����Ʊ��ʹ�ü�¼.����Ա���%Type;
  v_����Ա���� ����Ʊ��ʹ�ü�¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ����Ʊ��ʹ�ü�¼.�Ǽ�ʱ��%Type;
  n_����id     ����Ʊ��ʹ�ü�¼.����id%Type;
  v_��Ʊ��     Ʊ��ʹ����ϸ.����%Type;
  n_��Ʊ���   Ʊ��ʹ����ϸ.Ʊ�ݽ��%Type;
  n_����id     Ʊ��ʹ����ϸ.����id%Type;
  n_������ʽ   Number(2);
  n_Ӧ�ó���   Number(2);
  n_Ʊ��       Number(2);
  n_�Ƿ��Ʊ   Number(2);
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_����       ����Ʊ��ʹ�ü�¼.����%Type;
  v_������     ����Ʊ��ʹ�ü�¼.������%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤����   ����Ʊ��ʹ�ü�¼.ƾ֤����%Type;
  v_ƾ֤������ ����Ʊ��ʹ�ü�¼.ƾ֤������%Type;
  v_����ʱ��   ����Ʊ��ʹ�ü�¼.����ʱ��%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  v_Url����    ����Ʊ��ʹ�ü�¼.Url����%Type;
  c_��ά��     Clob;
  n_ԭƱ��id   ����Ʊ��ʹ�ü�¼.ԭƱ��id%Type;
  v_��ע       ����Ʊ��ʹ�ü�¼.��ע%Type;
  v_��Ʊ��     ����Ʊ��ʹ�ü�¼.��Ʊ��%Type;
  v_ϵͳ��Դ   ����Ʊ��ʹ�ü�¼.ϵͳ��Դ%Type;
  j_Input      PLJson;
  j_Json       PLJson;
  j_Temp       PLJson;
  j_Temp1      PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --������ʽ:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  n_������ʽ := j_Json.Get_Number('oper_mode');
  n_Id       := j_Json.Get_Number('einvoice_id');
  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  d_�Ǽ�ʱ��   := Nvl(To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss'), Sysdate);

  If Not j_Json.Exist('paperinv_info') Then
    Json_Out := zlJsonOut('��ֽ��Ʊ����Ϣ��');
    Return;
  End If;
  Select Max(����id) Into n_����id From ����Ʊ��ʹ�ü�¼ Where ID = n_Id;
  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('����ĵ���Ʊ����Ч!');
    Return;
  End If;

  j_Temp := j_Json.Get_Pljson('paperinv_info');
  --Ӧ�ó���:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  n_Ӧ�ó��� := j_Temp.Get_Number('inv_occasion');
  --Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  n_Ʊ�� := j_Temp.Get_Number('invoice_type');
  --�Ƿ��Ʊ:1-��Ʊ;0-�Ǻ�Ʊ
  n_�Ƿ��Ʊ := Nvl(j_Temp.Get_Number('inv_red'), 0);
  v_��Ʊ��   := j_Temp.Get_String('invoice_no');
  n_��Ʊ��� := j_Temp.Get_Number('inv_total');
  n_����id   := j_Temp.Get_Number('recv_id');

  --ֽ��Ʊ�ݴ���
  --������ʽ_In:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  Zl_ֽ��Ʊ��ʹ��_Update(n_Ӧ�ó���, n_Ʊ��, n_����id, n_Id, v_��Ʊ��, n_��Ʊ���, n_����id, v_����Ա����, d_�Ǽ�ʱ��, n_������ʽ, 0, n_�Ƿ��Ʊ);

  If j_Json.Exist('einvoce_info') Then
    --��ȡ����Ʊ����Ϣ
    j_Temp1      := PLJson();
    j_Temp1      := j_Json.Get_Pljson('einvoce_info');
    v_��Ʊ��     := j_Temp1.Get_String('placeCode');
    v_ϵͳ��Դ   := j_Temp1.Get_String('sys_source');
    v_��ע       := j_Temp1.Get_String('demo');
    n_ԭƱ��id   := j_Temp1.Get_Number('inv_oldid');
    n_����id     := j_Temp1.Get_Number('einvoice_id');
    v_����       := j_Temp1.Get_String('einvoice_code');
    v_����       := j_Temp1.Get_String('einvoice_no');
    v_������     := j_Temp1.Get_String('einvoice_random');
    v_ƾ֤����   := j_Temp1.Get_String('voucher_code');
    v_ƾ֤����   := j_Temp1.Get_String('voucher_no');
    v_ƾ֤������ := j_Temp1.Get_String('voucher_random');
    v_����ʱ��   := j_Temp1.Get_String('happen_time');
    v_Url����    := j_Temp1.Get_String('picture_url');
    v_Url����    := j_Temp1.Get_String('picture_neturl');
    c_��ά��     := j_Temp1.Get_Clob('qrcode');
  
    Zl_����Ʊ��ʹ�ü�¼_Delete(n_����id, v_��Ʊ��, v_ϵͳ��Դ, v_����ʱ��, v_��ע, v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, n_ԭƱ��id, v_����, v_����, v_������,
                       v_ƾ֤����, v_ƾ֤����, v_ƾ֤������, v_Url����, v_Url����);
    --���¶�ά��
    Insert Into ����Ʊ�ݶ�ά�� (ʹ�ü�¼id, ��ά��) Values (n_����id, c_��ά��);
    Json_Out := zlJsonOut('�ɹ�', 1);
    Return;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Savepaperinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Getstarteinvoices
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���õ���Ʊ��ҵ��
  --��Σ�Json_In:NULL
  --     
  --����: Json_Out,��ʽ����
  --output      
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  data[]      ����վ���б�
  --    occasion  N  1  ����:1-�շ�,2-Ԥ��,3-����,4-�Һ�
  --    client_name  C  1  վ����
  ---------------------------------------------------------------------------

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --�������
  For c_����� In (Select ����, վ�� From ����Ʊ��վ�����) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'occasion', c_�����.����, 1, 1);
    zlJsonPutValue(v_Output, 'client_name', c_�����.վ��, 0, 2);
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","data":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getstarteinvoices;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoicecode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ���ȡ��Ʊ����
  --��Σ�json��ʽ
  --input
  --   operator_id    N  1  ����ԱID
  --   ssite          C  1  �ͻ���
  --���Σ�json��ʽ
  --Json_Out
  --  code            C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message         C  1  Ӧ����Ϣ�� �ɹ�ʱ���ش���No��[����] ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  einvoice_code   C  1  ��Ʊ�����
  --  is_exist        N  1  Ʊ�ݿ�Ʊ������Ƿ��������:1-����;0-������
  -------------------------------------------------------------------------------------------------
  n_����Աid   Ʊ�ݿ�Ʊ�����.��Աid%Type;
  v_�ͻ���     Ʊ�ݿ�Ʊ�����.�ͻ���%Type;
  v_��Ʊ����� ����Ʊ�ݿ�Ʊ��.����%Type;
  j_Input      PLJson;
  j_Json       PLJson;
  n_Count      Number(2);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����Աid := j_Json.Get_Number('operator_id');
  v_�ͻ���   := j_Json.Get_String('ssite');

  Select Count(1) Into n_Count From Ʊ�ݿ�Ʊ����� Where Rownum < 2;
  If n_Count = 0 Then
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"","is_exist":0}}';
    Return;
  End If;

  --���շ�Ա+�ͻ��˶���
  For r_��Ʊ�� In (Select b.���� As ��Ʊ�����
                From Ʊ�ݿ�Ʊ����� A, ����Ʊ�ݿ�Ʊ�� B
                Where a.��Ʊ��id = b.Id And Nvl(b.����ʱ��, Sysdate + 1) >= Sysdate And a.��Աid = n_����Աid And a.�ͻ��� = v_�ͻ���) Loop
    v_��Ʊ����� := r_��Ʊ��.��Ʊ�����;
    Json_Out     := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Nvl(v_��Ʊ�����, '') || '","is_exist":1}}';
    Return;
  End Loop;

  --���շ�Ա����
  For r_��Ʊ�� In (Select Nvl(a.��Աid, 0) As ��Աid, Nvl(a.�ͻ���, '-') As �ͻ���, b.���� As ��Ʊ�����
                From Ʊ�ݿ�Ʊ����� A, ����Ʊ�ݿ�Ʊ�� B
                Where a.��Ʊ��id = b.Id And Nvl(b.����ʱ��, Sysdate + 1) >= Sysdate And a.��Աid = n_����Աid And
                      Nvl(a.�ͻ���, '-') = '-') Loop
    v_��Ʊ����� := r_��Ʊ��.��Ʊ�����;
    Json_Out     := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Nvl(v_��Ʊ�����, '') || '","is_exist":1}}';
    Return;
  End Loop;

  --���ͻ��˶���
  For r_��Ʊ�� In (Select Nvl(a.��Աid, 0) As ��Աid, Nvl(a.�ͻ���, '-') As �ͻ���, b.���� As ��Ʊ�����
                From Ʊ�ݿ�Ʊ����� A, ����Ʊ�ݿ�Ʊ�� B
                Where a.��Ʊ��id = b.Id And Nvl(b.����ʱ��, Sysdate + 1) >= Sysdate And Nvl(a.��Աid, 0) = 0 And a.�ͻ��� = v_�ͻ���) Loop
    v_��Ʊ����� := r_��Ʊ��.��Ʊ�����;
    Json_Out     := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Nvl(v_��Ʊ�����, '') || '","is_exist":1}}';
    Return;
  End Loop;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","einvoice_code":"' || Null || '","is_exist":1}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicecode;
/

Create Or Replace Procedure Zl_Exsesvr_Geteinvoicedata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݽ���ID,��ȡ��Ч�ĵ���Ʊ��ID
  --��Σ�Json_In:
  --input
  --  fun_oper            N 1 �������ͣ�0-����Ʊ�ֺͽ���id��ȡ����Ʊ��ID��1-���ݵ���Ʊ��ID��ȡ �Ƿ񻻿���ֽ�ʷ�Ʊ�š�����id
  --  blnc_id             N   ����ID(����Ʊ��ʹ�ü�¼.����id)
  --  inv_type            N   Ʊ��:1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-ҽ�Ʒ��� 
  --  einvoice_id         N   ����Ʊ��ID
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  einvoice_id         N    ��Ч�ĵ���Ʊ��ID(��������=0ʱ����)
  --  blnc_id             N    ����ID(��������=1ʱ����)
  --  is_turn             N    �Ƿ񻻿�(��������=1ʱ����)
  --  inv_no              N    ֽ��Ʊ��(��������=1ʱ����)
  ---------------------------------------------------------------------------
  j_Input          PLJson;
  j_Json           PLJson;
  n_��������       Number(2);
  n_Ʊ��           ����Ʊ��ʹ�ü�¼.Ʊ��%Type;
  n_�Ƿ񻻿�       ����Ʊ��ʹ�ü�¼.�Ƿ񻻿�%Type;
  v_ֽ�ʷ�Ʊ��     ����Ʊ��ʹ�ü�¼.ֽ�ʷ�Ʊ��%Type;
  n_����id         ����Ʊ��ʹ�ü�¼.����id%Type;
  n_����id_Out     ����Ʊ��ʹ�ü�¼.����id%Type;
  n_����Ʊ��id     ����Ʊ��ʹ�ü�¼.Id%Type;
  n_����Ʊ��id_Out ����Ʊ��ʹ�ü�¼.Id%Type;
Begin
  --�������

  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_�������� := Nvl(Pljson_Ext.Get_Number(j_Json, 'fun_oper'), 0);

  If n_�������� = 0 Then
    --����Ʊ�ֺͽ���id��ȡ����Ʊ��ID
    n_����id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
    n_Ʊ��   := Pljson_Ext.Get_Number(j_Json, 'inv_type');
    If (Nvl(n_����id, 0) = 0 Or Nvl(n_Ʊ��, 0) = 0) Then
      Json_Out := '{"output":{"code":0,"message": "ʧ��,����Ľ���id�򳡺�Ϊ0"}}';
      Return;
    End If;
  
    Select Max(ID)
    Into n_����Ʊ��id_Out
    From ����Ʊ��ʹ�ü�¼
    Where ����id = n_����id And Ʊ�� = n_Ʊ�� And ��¼״̬ = 1 And Nvl(ԭƱ��id, 0) = 0;
  
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","einvoice_id":' || Nvl(n_����Ʊ��id_Out, 0) || '}}';
  Else
    --���ݵ���Ʊ��ID��ȡ �Ƿ񻻿���ֽ�ʷ�Ʊ�š�����id
    n_����Ʊ��id := Pljson_Ext.Get_Number(j_Json, 'einvoice_id');
    If Nvl(n_����Ʊ��id, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message": "ʧ��,����ĵ���Ʊ��IDΪ0"}}';
      Return;
    End If;
  
    Select Max(�Ƿ񻻿�), Max(ֽ�ʷ�Ʊ��), Max(����id)
    Into n_�Ƿ񻻿�, v_ֽ�ʷ�Ʊ��, n_����id_Out
    From ����Ʊ��ʹ�ü�¼
    Where ID = n_����Ʊ��id;
  
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","blnc_id":' || Nvl(n_����id_Out, 0) || ',"is_turn":' ||
                Nvl(n_�Ƿ񻻿�, 0) || ',"inv_no":"' || v_ֽ�ʷ�Ʊ�� || '"}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicedata;
/
Create Or Replace Procedure Zl_Exsesvr_Checkiseinvoice
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݴ����Ʊ�ֺͽ���ID,��鵱ǰ�����Ƿ��������е���Ʊ��
  --��Σ�Json_In:
  --input
  --  blnc_id             N 1 ����ID(����Ʊ��ʹ�ü�¼.id)
  --  inv_type            N 1 Ʊ��:1-�շ�, 2-Ԥ��, 3-����, 4-�Һ�;5-ҽ�Ʒ��� 
  --����: Json_Out,��ʽ����
  --output      
  --  code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  is_einvoice         N  1 �Ƿ����õ���Ʊ��:1-����;0:δ����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Ʊ��     ����Ʊ��ʹ�ü�¼.Ʊ��%Type;
  n_����id   ����Ʊ��ʹ�ü�¼.Id%Type;
  n_Einvoice Number(2);
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := Pljson_Ext.Get_Number(j_Json, 'blnc_id');
  n_Ʊ��   := Pljson_Ext.Get_Number(j_Json, 'inv_type');

  If Nvl(n_����id, 0) = 0 Or Nvl(n_Ʊ��, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,����Ľ���id�򳡺�Ϊ0"}}';
    Return;
  End If;

  If n_Ʊ�� = 2 Then
    --Ԥ����¼
    Select Max(Ԥ������Ʊ��) Into n_Einvoice From ����Ԥ����¼ Where Mod(��¼����, 10) = 1 And ID = n_����id;
  Else
    Select Max(�Ƿ����Ʊ��) Into n_Einvoice From ����Ԥ����¼ Where ����id = n_����id;
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","is_einvoice":' || Nvl(n_Einvoice, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkiseinvoice;
/
Create Or Replace Procedure Zl_Exsesvr_Geteinvoiceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ����Ʊ����Ϣ
  --��Σ�Json_In:
  --input
  --err_id              N 1 �쳣ID
  --����: Json_Out,��ʽ����
  --output      
  --code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --message             C  1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --��¼��־=0 ʱ ����----------------------------------------
  --data
  --  input          
  --    balance_id        N  1  ����ID
  --    balance_delid     N     �˿�ID:�˿�ߺ�Ʊʱ��Ч��Ŀǰֻ��Ԥ������Ч,��д�����˿�Ԥ��ID
  --    einvoice_id       N  1  ����Ʊ��ID
  --    operator_code     C  1  ����Ա���
  --    operator_name     C  1  ����Ա����
  --    create_time       C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    pati_info         C     ������Ϣ
  --      pati_id         N  1  ����ID
  --      pati_pageid     N     ��ҳID
  --      pati_name       C  1  ����
  --      pati_sex        C  1  �Ա�
  --      pati_age        C  1  ����
  --      outpatient_num  C  1  �����
  --      inpatient_num   C  1  סԺ��
  --    einvoce_info      C     ����Ʊ����Ϣ
  --      invoice_type    N  1  Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      placeCode       C  1  ��Ʊ�����
  --      inv_total       N  1  ��Ʊ���
  --      inv_oldid       N    ԭƱ��ID
  --      sys_source      C  1  ϵͳ��Դ
  --      demo            C  1  ��ע
  --      einvoice_code   C  1  ����Ʊ�ݴ���
  --      einvoice_no     C  1  ����Ʊ�ݺ���
  --      einvoice_random C  1  ����У����
  --      voucher_code    C  1  Ԥ����ƾ֤����
  --      voucher_no      C  1  Ԥ����ƾ֤����
  --      voucher_random  C  1  Ԥ����ƾ֤У����
  --      happen_time     C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url     C  1  ����Ʊ��H5ҳ��URL
  --      picture_neturl  C  1  ����Ʊ������H5ҳ��URL
  --      qrcode          C  1  ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --��¼��־=1 ʱ ����-----------------------------------------
  --data
  --  input         
  --    einvoice_id       N 1 ����Ʊ��ID
  --    operator_code     C 1 ����Ա���
  --    operator_name     C 1 ����Ա����
  --    create_time       C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    einvoce_info      C   ����Ʊ����Ϣ
  --      placeCode       C 1 ��Ʊ�����
  --      sys_source      C 1 ϵͳ��Դ
  --      demo            C 1 ��ע
  --      inv_oldid       N   ԭƱ��ID
  --      einvoice_code   C 1 ����Ʊ�ݴ���
  --      einvoice_no     C 1 ����Ʊ�ݺ���
  --      einvoice_random C 1 ����У����
  --      voucher_code    C 1 Ԥ����ƾ֤����
  --      voucher_no      C 1 Ԥ����ƾ֤����
  --      voucher_random  C 1 Ԥ����ƾ֤У����
  --      happen_time     C 1 ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url     C 1 ����Ʊ��H5ҳ��URL
  --      picture_neturl  C 1 ����Ʊ������H5ҳ��URL
  --      qrcode          C 1 ����Ʊ�ݶ�ά��ͼƬ����:��ֵ��Base64���룬����ʱ��ҪBase64����,ͼƬ��ʽΪ:PNG
  --��¼��־=2,3 ʱ ����-------------------------------------------
  --data
  --  input         
  --    oper_mode         N 1 ������ʽ:0-����;1-���»���;2-����Ʊ��;3-����Ʊ��
  --    einvoice_id       N 1 ����Ʊ��ID
  --    operator_code     C 1 ����Ա���
  --    operator_name     C 1 ����Ա����
  --    create_time       C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --    paperinv_info     C   ֽ��Ʊ����Ϣ:���ڶ���ʱ���밴����˳���ϴ�(�������ݴ���)
  --      inv_occasion    N 1 Ӧ�ó���:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  --      invoice_type    N 1 Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      inv_red         N   �Ƿ��Ʊ:1-��Ʊ;0-�Ǻ�Ʊ
  --      invoice_no      C 1 ��Ʊ��
  --      inv_total       N 1 ��Ʊ���
  --      recv_id         N   ����id
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  c_Output Clob;
  n_�쳣id ����Ʊ���쳣��¼.Id%Type;

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�쳣id := Pljson_Ext.Get_Number(j_Json, 'err_id');

  If Nvl(n_�쳣id, 0) = 0 Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,������쳣idΪ0"}}';
    Return;
  End If;

  Begin
    Select Ʊ����Ϣ Into c_Output From ����Ʊ���쳣��¼ Where ID = n_�쳣id;
  Exception
    When Others Then
      Json_Out := '{"output":{"code":0,"message": "ʧ��,���ݴ�����쳣idδ�ҵ�����"}}';
      Return;
  End;

  If c_Output Is Null Then
    Json_Out := '{"output":{"code":0,"message": "ʧ��,���ݴ�����쳣idδ�ҵ�����"}}';
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message": "�ɹ�","data":' || c_Output || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoiceinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Drugwriteoff_Check
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ�ҩƷ�����ķ���������˼��
  --��Σ�Json_In:��ʽ
  --input      ҩƷ��������ǰ���
  --  part_ban_writeoffs    N  1  ��ֹ��������:0-����;1-������������(�����ŵ��ݵĲ��ֻ�ĳ�ʵĲ���)
  --  fee_origin            N  1  ������Դ:1-���2-סԺ
  --  rcpdtl_list[]               ���������б�
  --    oper_type           N  1  ��������:0-���ͨ�� 1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
  --    rcpdtl_id           N  1  ������ϸID(����ID)
  --    request_time        D     ����ʱ��
  --    request_type        N     �������ȱʡΪ1
  --    quantity            N  1  ����������Ϊ���nullʱ,������ID��������ֱ������
  --    sended_num          N  1  �ѷ�����
  --  pati_list[]                 ������Ϣ
  --    pati_id             N     ����ID,ΪNULL��0ʱ����ʾ���ŵ���
  --    fee_audit_status    N     ������˱�־:0���-δ���;1-����˻�ʼ���;2-������,��Ͻ���Ȩ��
  --    si_inp_status       N     סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
  --    catalog_date        C     ������Ŀ���ڣ�yyyy-mm-dd hh24:mi:ss
  --����: Json_Out,��ʽ����
  --output
  --   code                          C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  tip_list[]  C  1  ��ʾ�б�:��Ҫ�ǿ��ܴ��ڶ����ʾѯ�ʷ�ʽ���������б�,��ֹʱ������һ����Ϣ
  --    tip_mode  C  1  ���Ʒ�ʽ:1-��ʾѯ��;2-��ֹ
  --    tip_message  C  1  ��ʾ��Ϣ
  ---------------------------------------------------------------------------
  j_Input        PLJson;
  j_Json         PLJson;
  j_List         Pljson_List;
  j_Temp         PLJson;
  n_��ֹ�������� Number(2);
  n_���ʷ�ʽ     Number(2);
  n_��������     Number(2);
  n_�������     Number(2);
  d_����ʱ��     Date;
  n_��������     ���˷�������.����%Type;
  n_�ѷ�����     ���˷�������.����%Type;
  n_��������     ���˷�������.����%Type;
  n_��˲���id   ���˷�������.��˲���id%Type;
  n_״̬         Number(2);
  n_����id       ���˷�������.����id%Type;
  n_Find         Number(2);
  n_�ѽᵥ�ݲ��� Number(3);
  v_Err_Msg      Varchar2(1000);

  l_Writeoffs  t_NumList2 := t_NumList2(); --���ñ�����������
  l_Excutes    t_NumList2 := t_NumList2(); --ҩƷ�ѷ�����
  v_Patilist   Varchar2(32767);
  n_������Դ   Number;
  v_Json_In    Varchar2(32767);
  v_Itemlist   Varchar2(32767);
  v_Excutelist Varchar2(32767);

  v_�ѽ���� Varchar2(32767);
  n_Code     Number(2);
  Cursor c_������Ϣ Is
    Select Distinct /*+cardinality(b,10)*/ a.Id As ����id, a.�շ����, a.No, ���, a.���� As ��������, a.���� As �ѷ�����
    From סԺ���ü�¼ A
    Where a.Id = 0;

  r_������Ϣ c_������Ϣ%RowType;

  Type Ty_������Ϣ Is Ref Cursor;
  c_���ʷ�����Ϣ Ty_������Ϣ; --��̬�α����

  v_No ������ü�¼.No%Type;
Begin

  --ȡjson�ڵ��ֵ��Ҳ�Ǹ�json��
  j_Input        := PLJson(Json_In);
  j_Json         := j_Input.Get_Pljson('input');
  n_��ֹ�������� := Nvl(j_Json.Get_Number('part_ban_writeoffs'), 0); --��ֹ��������
  n_������Դ     := Nvl(j_Json.Get_Number('fee_origin'), 1);
  n_���ʷ�ʽ     := 1; --ҩ��ʹ�ã�ֻ��1:���ʷ�ʽ��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������

  If Not j_Json.Exist('rcpdtl_list') Then
    Json_Out := zlJsonOut('δ���뱾����Ҫ���ʵ�ҩƷ����������', 0);
    Return;
  End If;

  --0-���� 1-��ʾ 2-��ֹ
  n_�ѽᵥ�ݲ��� := To_Number(Nvl(zl_GetSysParameter('�ѽ��ʵ��ݲ���'), '0'));

  If n_�ѽᵥ�ݲ��� = 1 Then
    n_�ѽᵥ�ݲ��� := 2;
  Elsif n_�ѽᵥ�ݲ��� = 2 Then
    n_�ѽᵥ�ݲ��� := 1;
  End If;

  --������ؼ��
  j_List     := Pljson_List();
  j_List     := j_Json.Get_Pljson_List('pati_list');
  v_Patilist := j_List.To_Char();
  v_Patilist := ',"pati_list":' || v_Patilist;

  j_List := Pljson_List();
  j_List := j_Json.Get_Pljson_List('rcpdtl_list');
  For J In 1 .. j_List.Count Loop
    j_Temp     := PLJson();
    j_Temp     := PLJson(j_List.Get(J));
    n_�������� := Nvl(j_Temp.Get_Number('oper_type'), 0);
    n_����id   := Nvl(j_Temp.Get_Number('rcpdtl_id'), 0);
    n_������� := Nvl(j_Temp.Get_Number('request_type'), 1);
    n_�������� := Nvl(j_Temp.Get_Number('quantity'), 0);
    n_�ѷ����� := Nvl(j_Temp.Get_Number('sended_num'), 0);
    d_����ʱ�� := To_Date(j_Temp.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  
    --��������:0-���ͨ�� 1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
    If d_����ʱ�� Is Not Null Then
      Begin
        --״̬:0-����,1-���ͨ��,2-���δͨ��
        Select ״̬, ����, ��˲���id
        Into n_״̬, n_��������, n_��˲���id
        From ���˷������� A
        Where ����id = n_����id And ����ʱ�� = d_����ʱ��;
      Exception
        When Others Then
          n_״̬ := -1;
      End;
    End If;
    If Nvl(n_��������, 0) In (2, 3) Then
      --������˾ܾ� :
      --ȡ���ܾ�:��Ҫ��ɾ���Ѿ��ܾ�������
      If Nvl(n_״̬, 0) = -1 Or d_����ʱ�� Is Null Then
        If d_����ʱ�� Is Null Then
          Json_Out := zlJsonOut('δ����ָ���������������ݣ�����!', 0);
        Else
          Json_Out := zlJsonOut('δ��������ʱ��Ϊ' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '�ķ��������¼������!', 0);
        End If;
        Return;
      End If;
    
      If n_״̬ <> 2 Then
        Begin
          If Nvl(n_������Դ, 0) = 1 Then
            Select '���ݺ�:' || a.No || '�е�' || a.��� || '��(' || b.���� || ')��' || Decode(a.�շ����, '4', '����', 'ҩƷ') ||
                    ',��������˾ܾ��ļ�¼,���ܱ�����ȡ����'
            Into v_Err_Msg
            From ������ü�¼ A, �շ���ĿĿ¼ B
            Where a.Id = n_����id And a.�շ�ϸĿid = b.Id(+);
          Else
            Select '���ݺ�:' || a.No || '�е�' || a.��� || '��(' || b.���� || ')��' || Decode(a.�շ����, '4', '����', 'ҩƷ') ||
                    ',��������˾ܾ��ļ�¼,���ܱ�����ȡ����'
            Into v_Err_Msg
            From סԺ���ü�¼ A, �շ���ĿĿ¼ B
            Where a.Id = n_����id And a.�շ�ϸĿid = b.Id(+);
          End If;
        Exception
          When Others Then
            v_Err_Msg := Null;
        End;
        If v_Err_Msg Is Null Then
          Json_Out := zlJsonOut('δ�ҵ�����ID=' || n_����id || '�ķ��ü�¼�����鴫��ķ���ID�Ƿ���ȷ!', 0);
          Return;
        End If;
        v_Err_Msg := '{"tip_mode":2,"tip_message":"' || zlJsonStr(v_Err_Msg) || '"}';
        Json_Out  := '{"output":{"code":1,"message":"�ɹ�","tip_list":[' || v_Err_Msg || ']}}';
        Return;
      End If;
    End If;
    --��������:0-���ͨ�� 1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
    If Nvl(n_��������, 0) In (0, 2) Then
      --n_���ʷ�ʽ:0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
      If Nvl(n_���ʷ�ʽ, 0) = 1 Then
        If Nvl(n_��������, 0) < Nvl(n_��������, 0) Then
          Begin
            If Nvl(n_������Դ, 0) = 1 Then
              Select '���ݺ�:' || a.No || '�е�' || a.��� || '��(' || b.���� || ')��' || Decode(a.�շ����, '4', '����', 'ҩƷ') ||
                      '�ı�����������(' || Nvl(n_��������, 0) || ')�����˱�����������(' || Nvl(n_��������, 0) || ')��'
              Into v_Err_Msg
              From ������ü�¼ A, �շ���ĿĿ¼ B
              Where a.Id = n_����id And a.�շ�ϸĿid = b.Id(+);
            Else
              Select '���ݺ�:' || a.No || '�е�' || a.��� || '��(' || b.���� || ')��' || Decode(a.�շ����, '4', '����', 'ҩƷ') ||
                      '�ı�����������(' || Nvl(n_��������, 0) || ')�����˱�����������(' || Nvl(n_��������, 0) || ')��'
              Into v_Err_Msg
              From סԺ���ü�¼ A, �շ���ĿĿ¼ B
              Where a.Id = n_����id And a.�շ�ϸĿid = b.Id(+);
            End If;
          Exception
            When Others Then
              v_Err_Msg := Null;
          End;
          If v_Err_Msg Is Null Then
            Json_Out := zlJsonOut('δ�ҵ�����ID=' || n_����id || '�ķ��ü�¼�����鴫��ķ���ID�Ƿ���ȷ!', 0);
            Return;
          End If;
          v_Err_Msg := '{"tip_mode":2,"tip_message":"' || zlJsonStr(v_Err_Msg) || '"}';
          Json_Out  := '{"output":{"code":1,"message":"�ɹ�","tip_list":[' || v_Err_Msg || ']}}';
          Return;
        End If;
      End If;
    
      n_Find := 0;
      For I In 1 .. l_Writeoffs.Count Loop
        If l_Writeoffs(I).C1 = n_����id Then
          l_Writeoffs(I).C2 := n_�������� + l_Writeoffs(I).C2;
          l_Excutes(I).C2 := Nvl(n_�ѷ�����, 0) + l_Excutes(I).C2;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If Nvl(n_Find, 0) = 0 Then
        l_Writeoffs.Extend;
        l_Writeoffs(l_Writeoffs.Count) := t_NumObj2(n_����id, n_��������);
        l_Excutes.Extend;
        l_Excutes(l_Excutes.Count) := t_NumObj2(n_����id, n_�ѷ�����);
      End If;
    
    End If;
  End Loop;

  If l_Writeoffs.Count = 0 Then
    --��Ҫ��ȡ������
    Json_Out := zlJsonOut('�ɹ�', 1);
    Return;
  End If;

  If Nvl(n_������Դ, 0) = 1 Or n_������Դ Is Null Then
    --�������
    Open c_���ʷ�����Ϣ For
      Select Distinct /*+cardinality(b,10)*/ a.Id As ����id, a.�շ����, a.No, a.���, b.C2 As ��������, c.C2 As �ѷ�����
      From ������ü�¼ A, Table(l_Writeoffs) B, Table(l_Excutes) C
      Where a.Id = b.C1 And a.Id = c.C1(+)
      Order By a.No, a.���;
  Else
    Open c_���ʷ�����Ϣ For
      Select Distinct /*+cardinality(b,10)*/ a.Id As ����id, a.�շ����, a.No, a.���, b.C2 As ��������, c.C2 As �ѷ�����
      From סԺ���ü�¼ A, Table(l_Writeoffs) B, Table(l_Excutes) C
      Where a.Id = b.C1 And a.Id = c.C1(+)
      Order By a.No, a.���;
  End If;

  v_No         := Null;
  v_Json_In    := Null;
  v_Itemlist   := Null;
  v_Excutelist := Null;
  Loop
    Fetch c_���ʷ�����Ϣ
      Into r_������Ϣ;
    Exit When c_���ʷ�����Ϣ%NotFound;
  
    If Nvl(v_No, '.') <> r_������Ϣ.No Then
    
      If v_Itemlist Is Not Null Then
      
        v_Itemlist := ',"item_list":[' || v_Itemlist || ']';
        If Not v_Excutelist Is Null Then
          --�����ѷ������б�
          v_Excutelist := ',"excute_list":[' || v_Excutelist || ']';
        End If;
      
        v_Json_In := '"fee_no":"' || v_No || '"';
        v_Json_In := v_Json_In || ',"fee_bill_type":2';
        v_Json_In := v_Json_In || ',"balance_ban_writeoffs":' || Nvl(n_�ѽᵥ�ݲ���, 0);
        v_Json_In := v_Json_In || ',"part_ban_writeoffs":' || Nvl(n_��ֹ��������, 0);
        v_Json_In := v_Json_In || ',"oper_type":' || Nvl(n_���ʷ�ʽ, 0);
      
        --���������б�
        v_Json_In := v_Json_In || v_Itemlist;
        --�����ѷ������б�
        v_Json_In := v_Json_In || Nvl(v_Excutelist, '');
      
        v_Json_In := v_Json_In || Nvl(v_Patilist, '');
        v_Json_In := '{"input":{' || v_Json_In || '}}';
      
        If Nvl(n_������Դ, 1) = 1 Then
          Zl_������ʼ�¼_Delete_Check(v_Json_In, Json_Out);
        Else
          Zl_סԺ���ʼ�¼_Delete_Check(v_Json_In, Json_Out);
        End If;
        --
        j_Input := PLJson(Json_Out);
        j_Json  := j_Input.Get_Pljson('output');
      
        n_Code := j_Json.Get_Number('code');
        If n_Code = 0 Then
          Json_Out := '{"output":{"code":1,"message":"�ɹ�","tip_list":[{"tip_mode":2,"tip_message":"' ||
                      zlJsonStr(j_Json.Get_String('message')) || '"}]}}';
          Return;
        End If;
        If j_Json.Exist('balance_serials') Then
          If v_�ѽ���� Is Not Null Then
            v_�ѽ���� := v_�ѽ���� || Chr(13);
          End If;
          v_�ѽ���� := v_�ѽ���� || v_No || ':' || j_Json.Get_String('balance_serials');
        End If;
      End If;
      v_No         := r_������Ϣ.No;
      v_Itemlist   := Null;
      v_Excutelist := Null;
    End If;
  
    --���ļ�ҩƷ���
    If Instr(',4,5,6,7,', ',' || r_������Ϣ.�շ���� || ',') = 0 Then
      v_Err_Msg := '�ڵ���:' || v_No || '�еĵ�' || r_������Ϣ.��� || '���д��ڷ�ҩƷ�����ĵ��շ���Ŀ';
      Json_Out  := '{"output":{"code":1,"message":"�ɹ�","tip_list":[{"tip_mode":2,"tip_message":"' ||
                   zlJsonStr(v_Err_Msg) || '"}]}}';
    
      Return;
    End If;
    --������ϸ����
    If v_Itemlist Is Not Null Then
      v_Itemlist := v_Itemlist || ',';
    End If;
    v_Itemlist := Nvl(v_Itemlist, '') || '{"serial_num":' || Nvl(r_������Ϣ.���, 0);
    v_Itemlist := v_Itemlist || ',"quantity":' || zlJsonStr(r_������Ϣ.��������, 1) || '}';
  
    --�����ѷ�Ϊ����
    If v_Excutelist Is Not Null Then
      v_Excutelist := v_Excutelist || ',';
    End If;
    v_Excutelist := Nvl(v_Excutelist, '') || '{"fee_id":' || Nvl(r_������Ϣ.����id, 0);
    v_Excutelist := v_Excutelist || ',"sended_num":' || zlJsonStr(r_������Ϣ.�ѷ�����, 1) || '}';
  End Loop;

  If v_Itemlist Is Not Null Then
  
    v_Itemlist := ',"item_list":[' || v_Itemlist || ']';
    If Not v_Excutelist Is Null Then
      --�����ѷ������б�
      v_Excutelist := ',"excute_list":[' || v_Excutelist || ']';
    End If;
  
    v_Json_In := '"fee_no":"' || v_No || '"';
    v_Json_In := v_Json_In || ',"fee_bill_type":2';
    v_Json_In := v_Json_In || ',"balance_ban_writeoffs":' || Nvl(n_�ѽᵥ�ݲ���, 0);
    v_Json_In := v_Json_In || ',"part_ban_writeoffs":' || Nvl(n_��ֹ��������, 0);
    v_Json_In := v_Json_In || ',"oper_type":' || Nvl(n_���ʷ�ʽ, 0);
  
    --���������б�
    v_Json_In := v_Json_In || v_Itemlist;
    --�����ѷ������б�
    v_Json_In := v_Json_In || Nvl(v_Excutelist, '');
    v_Json_In := v_Json_In || Nvl(v_Patilist, '');
    v_Json_In := '{"input":{' || v_Json_In || '}}';
    If Nvl(n_������Դ, 1) = 1 Then
      Zl_������ʼ�¼_Delete_Check(v_Json_In, Json_Out);
    Else
      Zl_סԺ���ʼ�¼_Delete_Check(v_Json_In, Json_Out);
    End If;
    j_Input := PLJson(Json_Out);
    j_Json  := j_Input.Get_Pljson('output');
  
    n_Code := j_Json.Get_Number('code');
    If n_Code = 0 Then
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","tip_list":[{"tip_mode":2,"tip_message":"' ||
                  zlJsonStr(j_Json.Get_String('message')) || '"}]}}';
      Return;
    End If;
    If j_Json.Exist('balance_serials') Then
      If v_�ѽ���� Is Not Null Then
        v_�ѽ���� := v_�ѽ���� || Chr(13);
      End If;
      v_�ѽ���� := v_�ѽ���� || v_No || ':' || j_Json.Get_String('balance_serials');
    End If;
  End If;

  If v_�ѽ���� Is Not Null Then
    --���أ�ѯ�ʷ�ʽ
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","tip_list":[{"tip_mode":1,"tip_message":"' ||
                zlJsonStr('���µ����Ѿ�����:' || Chr(13) || v_�ѽ����) || '"}]}}';
    Return;
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Drugwriteoff_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Drugwriteoff
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ�ҩƷ�����ķ�������(�������ͨ��������ܾ���ȡ���ܾ�)
  --��Σ�Json_In:��ʽ
  --input     
  --  fee_origin            N  1  ������Դ��1-���2-סԺ��
  --  operator_code         C  1  ����Ա����
  --  operator_name         C  1  ����Ա���� 
  --  operator_time         C     ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --  rcpdtl_list                 [����]ÿ��������ϸ��Ϣ
  --    rcpdtl_id           N  1  ������ϸid(����id)
  --    request_time        D  1  ����ʱ��
  --    oper_type           N  1  ��������:0-���ͨ��;1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
  --    request_type        N  1  �������Ĭ�ϴ�1��
  --    quantity            N  1  ��������
  --    sended_num          N  1  �ѷ�����

  --����: Json_Out,��ʽ����
  --output
  --   code                          C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  Cursor c_������Ϣ Is
    Select Distinct /*+cardinality(b,10)*/ a.No, ���, a.�շ����, a.���� As ʣ������, a.���� As ��������, a.���� As �ѷ�����
    From סԺ���ü�¼ A
    Where a.Id = 0;

  r_������Ϣ c_������Ϣ%RowType;

  Type Ty_������Ϣ Is Ref Cursor;
  c_���ʷ�����Ϣ Ty_������Ϣ; --��̬�α����

  j_Input PLJson;
  j_Json  PLJson;
  j_List  Pljson_List;
  j_Temp  PLJson;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  n_�������� Number(2);
  n_������� Number(2);
  d_����ʱ�� Date;
  n_Find     Number(2);

  l_Writeoffs t_NumList2 := t_NumList2(); --���ñ�����������
  l_Excutes   t_NumList2 := t_NumList2(); --ҩƷ�ѷ�����

  n_���ʷ�ʽ Number(2);
  v_���     Varchar2(32767);
  n_Temp     Number(2);
  n_������Դ Number(2);

  v_No         ������ü�¼.No%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  n_��������   ���˷�������.����%Type;
  n_�ѷ�����   ���˷�������.����%Type;
  n_����id     ���˷�������.����id%Type;
  d_����ʱ��   Date;
  n_ִ��״̬   Number(2);
  n_Count      Number(2);
Begin

  --ȡjson�ڵ��ֵ��Ҳ�Ǹ�json��
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  n_������Դ   := Nvl(j_Json.Get_Number('fee_origin'), 1);

  If j_Json.Exist('operator_time') Then
    d_����ʱ�� := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_����ʱ�� := Sysdate;
  End If;
  n_���ʷ�ʽ := 1; --ҩ��ʹ�ã�ֻ��1:���ʷ�ʽ��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������

  If Not j_Json.Exist('rcpdtl_list') Then
    Json_Out := zlJsonOut('δ���뱾����Ҫ���ʵ�ҩƷ����������', 0);
    Return;
  End If;

  j_List := Pljson_List();
  j_List := j_Json.Get_Pljson_List('rcpdtl_list');
  For J In 1 .. j_List.Count Loop
    j_Temp := PLJson();
    j_Temp := PLJson(j_List.Get(J));
  
    n_�������� := Nvl(j_Temp.Get_Number('oper_type'), 0);
    n_����id   := Nvl(j_Temp.Get_Number('rcpdtl_id'), 0);
    d_����ʱ�� := To_Date(j_Temp.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
    n_������� := Nvl(j_Temp.Get_Number('request_type'), 1);
  
    n_�������� := Nvl(j_Temp.Get_Number('quantity'), 0);
    n_�ѷ����� := Nvl(j_Temp.Get_Number('sended_num'), 0);
  
    If Nvl(n_��������, 0) In (2, 3) Then
      --��������:0-���ͨ��;1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
      If Nvl(n_��������, 0) = 3 Then
        n_Temp := 1;
      Else
        n_Temp := 0;
      End If;
      -- n_Temp:0-��˾ܾ� 1-ȡ���ܾ�
      Zl_���˷�������_Cancel_s(n_����id, d_����ʱ��, v_����Ա����, d_����ʱ��, n_Temp, n_�������);
    Else
      If Nvl(n_��������, 0) = 1 Then
        --��������:0-���ͨ��; 1-��˲�ͨ��
        n_Temp := 2;
      Elsif Nvl(n_��������, 0) = 0 Then
        n_Temp := 1;
      Else
        v_Err_Msg := '����״̬�������(����ID=' || n_����id || ')��ֻ��Ϊ����״̬:0-���ͨ��;1-��˲�ͨ��;2-��˾ܾ�;3-ȡ���ܾ�';
        Raise Err_Item;
      End If;
      Select Count(1)
      Into n_Count
      From ���˷�������
      Where ����id = n_����id And ����ʱ�� = d_����ʱ�� And ������� = n_������� And ״̬ = n_Temp;
      If n_Count <> 0 Then
        v_Err_Msg := '�õ���(����ID=' || n_����id || ')����ˣ���ֹ�������';
        Raise Err_Item;
      End If;
      Zl_���˷�������_Audit_s(n_����id, d_����ʱ��, v_����Ա����, d_����ʱ��, n_Temp, n_�������);
    End If;
  
    --��������:0-���ͨ��;1-��˲�ͨ�� 2-��˾ܾ� 3-ȡ���ܾ�;
    If Nvl(n_��������, 0) In (0, 2) Then
      --�������ʴ���
      n_Find := 0;
      For I In 1 .. l_Writeoffs.Count Loop
        If l_Writeoffs(I).C1 = n_����id Then
          l_Writeoffs(I).C2 := n_�������� + l_Writeoffs(I).C2;
          l_Excutes(I).C2 := n_�ѷ�����;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If Nvl(n_Find, 0) = 0 Then
        l_Writeoffs.Extend;
        l_Writeoffs(l_Writeoffs.Count) := t_NumObj2(n_����id, n_��������);
        l_Excutes.Extend;
        l_Excutes(l_Excutes.Count) := t_NumObj2(n_����id, n_�ѷ�����);
      End If;
    End If;
  End Loop;

  If l_Writeoffs.Count = 0 Then
    Json_Out := zlJsonOut('�ɹ�', 1);
    Return;
  
  End If;
  If Nvl(n_������Դ, 0) = 1 Or n_������Դ Is Null Then
    --�������
    Open c_���ʷ�����Ϣ For
      Select a.No, a.���, a.�շ����, Sum(Nvl(a.����, 1) * Nvl(a.����, 0)) As ʣ������, Max(b.��������) As ��������, Max(b.�ѷ�����) As �ѷ�����
      From ������ü�¼ A,
           (Select Distinct /*+cardinality(b,10)*/ a.No, a.���, b.C2 As ��������, c.C2 As �ѷ�����
             From ������ü�¼ A, Table(l_Writeoffs) B, Table(l_Excutes) C
             Where a.Id = b.C1 And a.Id = c.C1(+)) B
      Where a.No = b.No And a.��� = b.��� And a.��¼���� = 2 And �۸񸸺� Is Null
      Group By a.No, a.���, a.�շ����
      Order By a.No, a.���;
  Else
    Open c_���ʷ�����Ϣ For
      Select a.No, a.���, a.�շ����, Sum(Nvl(a.����, 1) * Nvl(a.����, 0)) As ʣ������, Max(b.��������) As ��������, Max(b.�ѷ�����) As �ѷ�����
      From סԺ���ü�¼ A,
           (Select Distinct /*+cardinality(b,10)*/ a.No, a.���, b.C2 As ��������, c.C2 As �ѷ�����
             From סԺ���ü�¼ A, Table(l_Writeoffs) B, Table(l_Excutes) C
             Where a.Id = b.C1 And a.Id = c.C1(+)) B
      Where a.No = b.No And a.��� = b.��� And a.��¼���� = 2 And �۸񸸺� Is Null
      Group By a.No, a.���, a.�շ����
      Order By a.No, a.���;
  End If;

  v_No   := Null;
  v_��� := Null;
  Loop
    Fetch c_���ʷ�����Ϣ
      Into r_������Ϣ;
    Exit When c_���ʷ�����Ϣ%NotFound;
  
    If Nvl(v_No, '.') <> r_������Ϣ.No Then
      If v_��� Is Not Null Then
        If n_������Դ = 1 Then
          Zl_������ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, d_����ʱ��);
        Elsif n_������Դ = 2 Then
          Zl_סԺ���ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, 2, 1, d_����ʱ��);
        End If;
      End If;
      v_No   := r_������Ϣ.No;
      v_��� := Null;
    End If;
  
    n_ִ��״̬ := 0;
    If Nvl(r_������Ϣ.�ѷ�����, 0) <> 0 Then
      n_ִ��״̬ := 2;
    End If;
    If Nvl(r_������Ϣ.�ѷ�����, 0) = Nvl(r_������Ϣ.ʣ������, 0) - Nvl(r_������Ϣ.��������, 0) Then
      n_ִ��״̬ := 1;
    End If;
    If Instr(',4,5,6,7,', ',' || r_������Ϣ.�շ���� || ',') = 0 Then
      v_Err_Msg := '�ڵ���:' || v_No || '�еĵ�' || r_������Ϣ.��� || '���д��ڷ�ҩƷ�����ĵ��շ���Ŀ';
      Raise Err_Item;
    End If;
    If v_��� Is Not Null Then
      v_��� := v_��� || ',';
    End If;
    v_��� := v_��� || r_������Ϣ.��� || ':' || r_������Ϣ.�������� || ':' || n_ִ��״̬;
    --���1:����1:ִ��״̬1,���2:����2:ִ��״̬2,...���n:����n:ִ��״̬n  ��:"1:2:1,2:10:1,3:2:1"
  End Loop;
  If v_��� Is Not Null Then
    If n_������Դ = 1 Then
      Zl_������ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, d_����ʱ��);
    Elsif n_������Դ = 2 Then
      Zl_סԺ���ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, 2, n_���ʷ�ʽ, d_����ʱ��);
    End If;
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Json_Out := zlJsonOut(v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Drugwriteoff;
/
Create Or Replace Procedure Zl_Exsesvr_Updateexeinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ����·���ִ�в��š�ִ���ˡ���ҩ���ڼ�ִ��״̬����Ϣ
  --��Σ�Json_In:��ʽ
  --input     
  --  fee_origin            N  1  ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
  --  operator_code         C     ����Ա����
  --  operator_name         C     ����Ա����
  --  operator_time         C     ����ʱ��
  --  item_list                   ���б����ִ�������Ϣ�������б�ʱͬʱ��Ҫ����fee_origin
  --    fee_id              N     ����id,������ʱ�Է��õ��ݺš�ҽ��id���շ�ϸĿidΪ׼         
  --    fee_no              C     ���õ��ݺ�
  --    advice_id           N     ҽ��id(�Ѿ���ȷ�����շѵ����Ǽ��ʵ��ˣ����Բ����ٴ��뵥������)
  --    fee_item_id         N     �շ�ϸĿid     
  --                              ע�⣺fee_id��(fee_no��advice_id��fee_item_id)�ش�����һ������.
  --    exe_nums            N  1  ��ִ������:Ϊ0��ʾ��δִ��
  --    exe_people          C     ִ����:����ִ�л���ȫִ��ʱ����Ҫ���룬������ʱ����operator_nameΪ׼
  --    exe_time            D     ִ��ʱ��:yyyy-mm-dd hh24:mi:ss,:����ִ�л���ȫִ��ʱ����Ҫ���룬������ʱ����"create_time"Ϊ׼
  --    pharmacy_window     C     ��ҩ����:ҩƷ��������Ч,�޴˽ӵ㣬������·�ҩ����
  --  deptchange_list       C  1  ִ�п��ұ����Ϣ�б�
  --    fee_id              N  1  ����id
  --    exe_old_deptid      N     ԭִ�п���ID 
  --    exe_deptid          N  1  ִ�в���id
  --  delrcp_list           C     [����]�Զ�����ʱ,��Ҫͬ������
  --    rcp_no              C  1  ����no
  --    serial_nums         C  1  ��ʽ: ���1:����:ִ��״̬1,���2:����2:ִ��״̬2,...
  --    operator_status     N     ����״̬��0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
  --����: Json_Out,��ʽ����
  --output
  --   code                 C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message              C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  Cursor c_���� Is
    Select a.No, a.���, Mod(a.��¼����, 10) As ��¼����, a.�շ����, a.���� As ʣ������
    From סԺ���ü�¼ A
    Where a.Id = 0;

  r_������Ϣ c_����%RowType;

  Cursor c_���ұ�� Is
    Select a.����id, a.No, Nvl(a.�۸񸸺�, a.���) As ���, Mod(a.��¼����, 10) As ��¼����, a.��ҳid, a.���˲���id, a.���˿���id, a.��������id,
           a.ִ�в���id, a.������Ŀid, a.�����־, Sum(Nvl(a.ʵ�ս��, 0) - Nvl(a.���ʽ��, 0)) As δ����, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��
    From סԺ���ü�¼ A
    Where ID = 0
    Order By a.�շ�ϸĿid;

  r_���ұ�� c_���ұ��%RowType;

  Type Ty_������Ϣ Is Ref Cursor;
  c_������Ϣ Ty_������Ϣ; --��̬�α����

  j_Input PLJson;
  j_Json  PLJson;
  j_List  Pljson_List;
  j_Temp  PLJson;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  v_��� Varchar2(32767);

  v_No         ������ü�¼.No%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_ִ����     ������ü�¼.ִ����%Type;
  d_ִ��ʱ��   ������ü�¼.ִ��ʱ��%Type;
  n_ִ������   ������ü�¼.����%Type;
  n_����ֵ1    ������ü�¼.���ʽ��%Type;
  n_����ֵ     ������ü�¼.���ʽ��%Type;
  n_����id     ���˷�������.����id%Type;
  v_���ݺ�     ������ü�¼.No%Type;
  n_ҽ��id     ������ü�¼.ҽ�����%Type;
  n_�շ�ϸĿid ������ü�¼.�շ�ϸĿid%Type;

  v_��ҩ���� ������ü�¼.��ҩ����%Type;

  n_ԭִ�в���id ������ü�¼.ִ�в���id%Type;

  n_��ǰִ�в���id ������ü�¼.ִ�в���id%Type;

  d_����ʱ��     Date;
  n_������Դ     Number(2);
  n_ִ��״̬     Number(2);
  n_���ӱ�־     Number(2);
  n_����״̬     Number(2);
  n_����ִ����Ϣ Number(2);
Begin

  --ȡjson�ڵ��ֵ��Ҳ�Ǹ�json��
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  n_������Դ   := Nvl(j_Json.Get_Number('fee_origin'), 2);
  If j_Json.Exist('operator_time') Then
    d_����ʱ�� := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_����ʱ�� := Sysdate;
  End If;

  If j_Json.Exist('item_list') Then
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('item_list');
    For J In 1 .. j_List.Count Loop
      j_Temp := PLJson();
      j_Temp := PLJson(j_List.Get(J));
    
      n_����id     := Nvl(j_Temp.Get_Number('fee_id'), 0);
      n_ִ������   := Nvl(j_Temp.Get_Number('exe_nums'), 0);
      v_���ݺ�     := Nvl(j_Temp.Get_String('fee_no'), '-');
      n_ҽ��id     := Nvl(j_Temp.Get_Number('advice_id'), 0);
      n_�շ�ϸĿid := Nvl(j_Temp.Get_Number('fee_item_id'), 0);
    
      If n_����id = 0 Then
        If v_���ݺ� = '-' Or n_ҽ��id = 0 Or n_�շ�ϸĿid = 0 Then
          v_Err_Msg := '��νڵ�fee_id��(fee_no��advice_id��fee_item_id)������һ����Ϊ�գ�����!';
          Raise Err_Item;
        End If;
      End If;
    
      If j_Temp.Exist('exe_people') Then
        v_ִ����       := j_Temp.Get_String('exe_people');
        d_ִ��ʱ��     := To_Date(j_Temp.Get_String('exe_time'), 'yyyy-mm-dd hh24:mi:ss');
        n_����ִ����Ϣ := 1;
      End If;
    
      v_��ҩ���� := Null;
      If j_Temp.Exist('pharmacy_window') Then
        --����ڵ㣬�򰴸ýڵ����
        v_��ҩ���� := Nvl(j_Temp.Get_String('pharmacy_window'), ' ');
      End If;
    
      If Nvl(n_������Դ, 1) = 1 Then
        --�������
        If n_����id <> 0 Then
          Open c_������Ϣ For
            Select Distinct /*+cardinality(b,10)*/ a.No, a.���, Mod(a.��¼����, 10) As ��¼����, a.�շ����,
                            Sum(Nvl(a.����, 1) * a.����) As ʣ������
            From ������ü�¼ A, (Select NO, ���, Mod(��¼����, 10) As ��¼���� From ������ü�¼ Where ID = n_����id) B
            Where a.No = b.No And a.��� = b.��� And Mod(a.��¼����, 10) = b.��¼���� And a.��¼���� <> 12
            Group By a.No, a.���, Mod(a.��¼����, 10), a.�շ����;
        Else
          Open c_������Ϣ For
            Select a.No, a.���, Mod(a.��¼����, 10) As ��¼����, a.�շ����, Sum(Nvl(a.����, 1) * a.����) As ʣ������
            From ������ü�¼ A
            Where a.No = v_���ݺ� And a.ҽ����� + 0 = n_ҽ��id And a.�շ�ϸĿid + 0 = n_�շ�ϸĿid And a.�۸񸸺� Is Null And a.��¼���� <> 12
            Group By a.No, a.���, Mod(a.��¼����, 10), a.�շ����;
        End If;
      Else
        --סԺ����
        If n_����id <> 0 Then
          Open c_������Ϣ For
            Select Distinct /*+cardinality(b,10)*/ a.No, a.���, Mod(a.��¼����, 10) As ��¼����, a.�շ����,
                            Sum(Nvl(a.����, 1) * a.����) As ʣ������
            From סԺ���ü�¼ A, (Select NO, ���, Mod(��¼����, 10) As ��¼���� From סԺ���ü�¼ Where ID = n_����id) B
            Where a.No = b.No And a.��� = b.��� And a.��¼���� = b.��¼����
            Group By a.No, a.���, Mod(a.��¼����, 10), a.�շ����;
        Else
          Open c_������Ϣ For
            Select a.No, a.���, Mod(a.��¼����, 10) As ��¼����, a.�շ����, Sum(Nvl(a.����, 1) * a.����) As ʣ������
            From סԺ���ü�¼ A
            Where a.No = v_���ݺ� And a.ҽ����� + 0 = n_ҽ��id And a.�շ�ϸĿid + 0 = n_�շ�ϸĿid And a.�۸񸸺� Is Null And a.��¼���� <> 12
            Group By a.No, a.���, Mod(a.��¼����, 10), a.�շ����;
        End If;
      End If;
    
      Fetch c_������Ϣ
        Into r_������Ϣ;
    
      If c_������Ϣ%NotFound Then
        If n_����id <> 0 Then
          v_Err_Msg := 'δ�ҵ���Ӧ�ķ��ü�¼(����ID=' || n_����id || ')';
        Else
          v_Err_Msg := 'δ�ҵ���Ӧ�ķ��ü�¼(���ݺ�=' || v_���ݺ� || ')';
        End If;
        Raise Err_Item;
      End If;
    
      n_ִ��״̬ := 0;
      If Nvl(n_ִ������, 0) <> 0 Then
        n_ִ��״̬ := 2;
        If Nvl(r_������Ϣ.ʣ������, 0) = Nvl(n_ִ������, 0) Then
          n_ִ��״̬ := 1;
        End If;
        If Abs(Nvl(r_������Ϣ.ʣ������, 0)) < Abs(Nvl(n_ִ������, 0)) Then
          v_Err_Msg := '���ݺ�Ϊ' || r_������Ϣ.No || '�еĵ�' || r_������Ϣ.��� || '�е�ʣ��������С����ִ������������!';
          Raise Err_Item;
        End If;
      End If;
    
      n_���ӱ�־ := Null;
      If Instr(',5,6,7,', ',' || r_������Ϣ.�շ���� || ',') > 0 Then
        n_���ӱ�־ := 0;
        If Nvl(v_ִ����, '-') <> '-' Then
          n_���ӱ�־ := 1;
        End If;
      End If;
    
      If Nvl(n_������Դ, 1) = 1 Then
        Update ������ü�¼
        Set ִ��״̬ = n_ִ��״̬, ���ӱ�־ = Nvl(n_���ӱ�־, ���ӱ�־), ִ���� = Decode(n_����ִ����Ϣ, 1, v_ִ����, ִ����),
            ִ��ʱ�� = Decode(n_����ִ����Ϣ, 1, d_ִ��ʱ��, ִ��ʱ��), ��ҩ���� = LTrim(Nvl(v_��ҩ����, ��ҩ����))
        Where NO = r_������Ϣ.No And Nvl(�۸񸸺�, ���) = r_������Ϣ.��� And ��¼״̬ In (0, 1, 3) And ��¼���� = r_������Ϣ.��¼����;
      Else
        Update סԺ���ü�¼
        Set ִ��״̬ = n_ִ��״̬, ���ӱ�־ = Nvl(n_���ӱ�־, ���ӱ�־), ִ���� = Decode(n_����ִ����Ϣ, 1, v_ִ����, ִ����),
            ִ��ʱ�� = Decode(n_����ִ����Ϣ, 1, d_ִ��ʱ��, ִ��ʱ��)
        Where NO = r_������Ϣ.No And Nvl(�۸񸸺�, ���) = r_������Ϣ.��� And ��¼״̬ In (0, 1, 3) And ��¼���� = r_������Ϣ.��¼����;
      End If;
    
    End Loop;
    Close c_������Ϣ;
  End If;

  --ִ�п��ұ��
  --1 �����·���ִ�в��ż�δ�ӷ���
  If j_Json.Exist('deptchange_list') Then
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('deptchange_list');
    For J In 1 .. j_List.Count Loop
      j_Temp           := PLJson();
      j_Temp           := PLJson(j_List.Get(J));
      n_����id         := Nvl(j_Temp.Get_Number('fee_id'), 0);
      n_ԭִ�в���id   := Nvl(j_Temp.Get_Number('exe_old_deptid'), 0);
      n_��ǰִ�в���id := j_Temp.Get_Number('exe_deptid');
      If Nvl(n_��ǰִ�в���id, 0) = 0 Then
        v_Err_Msg := '������ĵ�ִ�п���Ϊ0������!';
        Raise Err_Item;
      End If;
    
      If Nvl(n_������Դ, 1) = 1 Then
        Open c_������Ϣ For
          Select /*+cardinality(b,10)*/
           a.����id, a.No, Nvl(a.�۸񸸺�, a.���) As ���, Mod(a.��¼����, 10) As ��¼����, 0 ��ҳid, 0 ���˲���id, a.���˿���id, a.��������id,
           a.ִ�в���id, a.������Ŀid, a.�����־, Sum(Nvl(a.ʵ�ս��, 0) - Nvl(a.���ʽ��, 0)) As δ����, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��
          From ������ü�¼ A,
               (Select NO, ���, Mod(��¼����, 10) As ��¼����
                 From ������ü�¼
                 Where ID = n_����id And (Nvl(ִ�в���id, 0) = n_ԭִ�в���id Or ִ�в���id Is Null) And Nvl(ִ�в���id, 0) <> n_��ǰִ�в���id) B
          Where a.No = b.No And a.��� = b.��� And Mod(a.��¼����, 10) = b.��¼����
          Group By a.����id, a.No, Nvl(a.�۸񸸺�, a.���), Mod(a.��¼����, 10), a.���˿���id, a.��������id, a.ִ�в���id, a.������Ŀid, a.�����־
          Order By a.No, Nvl(a.�۸񸸺�, a.���);
      Else
        Open c_������Ϣ For
          Select /*+cardinality(b,10)*/
           a.����id, a.No, Mod(a.��¼����, 10) As ��¼����, a.��ҳid, a.���˲���id, a.���˿���id, a.��������id, a.ִ�в���id, a.������Ŀid, a.�����־,
           Sum(Nvl(a.ʵ�ս��, 0) - Nvl(a.���ʽ��, 0)) As δ����, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��
          From סԺ���ü�¼ A,
               (Select NO, ���, Mod(��¼����, 10) As ��¼����
                 From סԺ���ü�¼
                 Where ID = n_����id And (Nvl(ִ�в���id, 0) = n_ԭִ�в���id Or ִ�в���id Is Null) And Nvl(ִ�в���id, 0) <> n_��ǰִ�в���id) B
          Where a.No = b.No And a.��� = b.��� And Mod(a.��¼����, 10) = b.��¼����
          Group By a.����id, a.No, Nvl(a.�۸񸸺�, a.���), Mod(a.��¼����, 10), a.��ҳid, a.���˲���id, a.���˿���id, a.��������id, a.ִ�в���id,
                   a.������Ŀid, a.�����־
          Order By a.No, Nvl(a.�۸񸸺�, a.���);
      End If;
      Loop
        Fetch c_������Ϣ
          Into r_���ұ��;
        Exit When c_������Ϣ%NotFound;
      
        If r_���ұ��.��¼���� = 2 Then
        
          If Trunc(Sysdate) > Trunc(r_���ұ��.�Ǽ�ʱ��) Then
            --���˷��û��ܿ����Ѿ����㣬��ˣ���������
            v_Err_Msg := '���ݺ�Ϊ' || r_������Ϣ.No || '�ļ��ʵ��ǵ���ļ��ʵ��ݣ���ֹ����ִ�п���!';
            Raise Err_Item;
          End If;
        
          --��ԭ�ⷿ��δ�����
          Update ����δ�����
          Set ��� = Nvl(���, 0) - Nvl(r_���ұ��.δ����, 0)
          Where ����id = r_���ұ��.����id And Nvl(��ҳid, 0) = Nvl(r_���ұ��.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_���ұ��.���˲���id, 0) And
                Nvl(���˿���id, 0) = Nvl(r_���ұ��.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_���ұ��.��������id, 0) And
                Nvl(ִ�в���id, 0) = Nvl(r_���ұ��.ִ�в���id, 0) And ������Ŀid + 0 = r_���ұ��.������Ŀid And ��Դ;�� + 0 = r_���ұ��.�����־
          Returning ��� Into n_����ֵ1;
        
          If Sql%RowCount <> 0 Then
            --�����ֿⷿ��δ�����
            Update ����δ�����
            Set ��� = Nvl(���, 0) + Nvl(r_���ұ��.δ����, 0)
            Where ����id = r_���ұ��.����id And Nvl(��ҳid, 0) = Nvl(r_���ұ��.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_���ұ��.���˲���id, 0) And
                  Nvl(���˿���id, 0) = Nvl(r_���ұ��.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_���ұ��.��������id, 0) And
                  Nvl(ִ�в���id, 0) = n_��ǰִ�в���id And ������Ŀid + 0 = r_���ұ��.������Ŀid And ��Դ;�� + 0 = r_���ұ��.�����־
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_���ұ��.����id, Decode(r_���ұ��.��ҳid, 0, Null, r_���ұ��.��ҳid), Decode(r_���ұ��.���˲���id, 0, Null, r_���ұ��.���˲���id),
                 Decode(r_���ұ��.���˿���id, 0, Null, r_���ұ��.���˿���id), Decode(r_���ұ��.��������id, 0, Null, r_���ұ��.��������id),
                 Decode(n_��ǰִ�в���id, 0, Null, n_��ǰִ�в���id), Decode(r_���ұ��.������Ŀid, 0, Null, r_���ұ��.������Ŀid), r_���ұ��.�����־,
                 Nvl(r_���ұ��.δ����, 0));
              n_����ֵ := Nvl(r_���ұ��.δ����, 0);
            End If;
          End If;
        
          If n_����ֵ = 0 Or n_����ֵ1 = 0 Then
            Delete From ����δ����� Where ����id = r_���ұ��.����id And Nvl(���, 0) = 0;
          End If;
        
        End If;
      
        If Nvl(n_������Դ, 1) = 1 Then
          Update ������ü�¼
          Set ִ�в���id = n_��ǰִ�в���id
          Where NO = r_���ұ��.No And Mod(��¼����, 10) = r_���ұ��.��¼���� And Nvl(�۸񸸺�, ���) = r_���ұ��.���;
        Else
        
          Update סԺ���ü�¼
          Set ִ�в���id = n_��ǰִ�в���id
          Where NO = r_���ұ��.No And Mod(��¼����, 10) = r_���ұ��.��¼���� And Nvl(�۸񸸺�, ���) = r_���ұ��.���;
        End If;
      
      End Loop;
      Close c_������Ϣ;
    End Loop;
  End If;

  --��������:ȡ����Һʱ����Ҫͬ������(ȡ����ҩʱ���ӵ���Ŀ)
  If j_Json.Exist('delrcp_list') Then
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('delrcp_list');
    For J In 1 .. j_List.Count Loop
      j_Temp     := PLJson(j_List.Get(J));
      v_No       := j_Temp.Get_String('rcp_no');
      v_���     := j_Temp.Get_String('serial_nums');
      n_����״̬ := j_Temp.Get_Number('operator_status');
      If n_����״̬ Is Null Then
        n_����״̬ := 0;
      End If;
      If Nvl(n_������Դ, 1) = 1 Then
        Zl_������ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, d_����ʱ��);
      Elsif n_������Դ = 2 Then
        Zl_סԺ���ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, 2, n_����״̬, d_����ʱ��);
      End If;
    End Loop;
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateexeinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Updatedepositinvinf
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����Ԥ��Ʊ����Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      fun_oper          N 1 ��������:1-����;2-�ش�3-����;4-��Ʊ��ӡ
  --      deposit_no        C 1 Ԥ������
  --      recv_id           N 1 ����id
  --      inv_no            C 1 ��ǰ��Ʊ�Ż�ʼʹ�÷�Ʊ��
  --      inv_usenums       N 1 ��Ʊʹ������
  --      use_time          C 1 Ʊ��ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
  --      inv_user          C 1 ��Ʊʹ����
  --����: Json_Out,��ʽ����
  --   output
  --     code               C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� �ɹ�ʱ���سɹ���Ϣ ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_��������     Number(2);
  v_Ԥ������     Varchar2(20);
  n_����id       Number(18);
  n_��Ʊʹ������ Number(18);
  v_Ʊ�ݺ�     Ʊ��ʹ����ϸ.����%Type;
  v_ʹ����       Ʊ��ʹ����ϸ.ʹ����%Type;
  d_ʹ��ʱ��     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type;
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin
  --�������
  
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��������     := j_Json.Get_Number('fun_oper');
  v_Ԥ������     := j_Json.Get_String('deposit_no');
  n_����id       := j_Json.Get_Number('recv_id');
  v_Ʊ�ݺ�       := j_Json.Get_String('inv_no');
  n_��Ʊʹ������ := j_Json.Get_Number('inv_usenums');
  v_ʹ����       := j_Json.Get_String('inv_user');
  d_ʹ��ʱ��     := To_Date(j_Json.Get_String('use_time'), 'yyyy-mm-dd hh24:mi:ss');
  If d_ʹ��ʱ�� Is Null Then
    d_ʹ��ʱ�� := Sysdate;
  End If;
  If v_Ԥ������ Is Null Then
    Json_Out := zlJsonOut('δ������Ҫ��ӡ�ĵ�����Ϣ!');
    Return;
  End If;
  If n_��������=1 Then
     zl_����Ԥ��Ʊ��_Insert(v_Ԥ������,v_Ʊ�ݺ�,n_����id,v_ʹ����,d_ʹ��ʱ��,n_��Ʊʹ������);
  Elsif n_��������=2 Then
     zl_����Ԥ����¼_RePrint(v_Ԥ������,v_Ʊ�ݺ�,n_����id,v_ʹ����);
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatedepositinvinf;
/

Create Or Replace Procedure Zl_Exsesvr_Getexsespec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ����ù���Ƿ���������ü�¼
  --input   ���ݲ���id����Ƿ���������ü�¼
  --  item_id       N   1   �շ�ϸĿid
  --output
  --  code          C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message       C   1   Ӧ����Ϣ��
  --  item_id       N   1   �շ�ϸĿid
  -------------------------------------------------------------------------------------------------
  n_�շ�ϸĿid Number(18);
  j_Input      PLJson;
  j_Json       PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_�շ�ϸĿid := Nvl(j_Json.Get_Number('item_id'), 0);

  Select Nvl(Max(�շ�ϸĿid), 0)
  Into n_�շ�ϸĿid
  From (Select �շ�ϸĿid
         From ������ü�¼
         Where �շ�ϸĿid = n_�շ�ϸĿid And Rownum < 2
         Union All
         Select �շ�ϸĿid
         From סԺ���ü�¼
         Where �շ�ϸĿid = n_�շ�ϸĿid And Rownum < 2)
  Where Rownum < 2;
  If Nvl(n_�շ�ϸĿid, 0) = 0 Then
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","item_id":null}}';
  Else
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","item_id":' || n_�շ�ϸĿid || '}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getexsespec;
/


Create Or Replace Procedure Zl_Exsesvr_Delbill_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ������ָ�����н�������
  --��Σ�Json_In:��ʽ
  --  input
  --      fee_no                  C   1   ���õ��ݺ�
  --      fee_bill_type           N   1   ��������:2-������ʵ�,3-�Զ����ʵ�
  --      balance_ban_writeoffs   N   1   �ѽ��ֹ����:����ѽ��ʵ��ݽ�ֹ����,����ҽ�����ʵĵ��ݡ�����ԭʼ��������ֻȡδ���ʲ���
  --      part_ban_writeoffs      N   1   ��ֹ��������:1-������0-����
  --      fee_origin              N   1   ������Դ��1-������ʣ�2-סԺ���ʣ�
  --      item_list[]             ���������б�
  --          serial_num          N   1   ���
  --          quantity            N   1   ��������(Ϊ��ʱ�������ֱ������)
  --      excute_list[]           ������ִ���б�(ҩƷ�����ķ���),��ʹ��ִ����Ϊ0ҲҪ����
  --          fee_id              N   1   ����ID
  --          sended_num          N   1   �ѷ�����
  --      advice_excute_list[]    ������ִ���б�(ҽ������),��ʹ��ִ����Ϊ0ҲҪ����
  --          advice_id           N   1   ҽ��ID
  --          fee_item_id         N   1   �շ�ϸĿID
  --          execute_num         N   1   ��ִ����
  --      pati_list[]             ������Ϣ���������Щ���˵ķ���
  --          pati_id             N   1   ����ID
  --          fee_audit_status    N   1   ������˱�־:0���-δ���;1-����˻�ʼ���(��ϲ���:������˷�ʽ������);2-������,��Ͻ���Ȩ��[��ֹδ��˲��˽���]���й������
  --          si_inp_status       N   1   סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
  --          catalog_date        C   0   ������Ŀ���ڣ�yyyy-mm-dd hh24:mi:ss
  --����: Json_Out,��ʽ����
  --  output
  --      code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --      message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --      item_list[]                         ���������б�
  --          serial_num          N   1   ���
  --          quantity            N   1   ��������
  --          execute_tag         N   1   ִ��״̬��0-δִ��;1-��ִ��;2-����ִ��

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_������Դ Number(1);
Begin

  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_������Դ := Nvl(j_Json.Get_Number('fee_origin'), 0);

  If n_������Դ = 1 Or n_������Դ Is Null Then
    Zl_������ʼ�¼_Delete_Check(Json_In, Json_Out);
  Else
    Zl_סԺ���ʼ�¼_Delete_Check(Json_In, Json_Out);
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Delbill_Check;
/


Create Or Replace Procedure Zl_Exsesvr_Cancelacc_Reaudit
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --����֮ǰ�Ѿܾ������˼�¼
  ---------------------------------------------------------------------------
  --input      ����֮ǰ�Ѿܾ������˼�¼
  --  rcpdtl_id  N  1  ������ϸid(����id)
  --  request_time  D  1  ����ʱ��
  --  audit_operator  C  1  �����
  --  fee_audit_time  D  1  ���ʱ��
  --  oper_type  N  1  ��������:0-��˾ܾ� 1-ȡ���ܾ�
  --  auto_stuff_return  N  1  �Զ�����
  --  request_type  N    �������
  ---------------------------------------------------------------------------
  j_Json  PLJson;
  j_Input PLJson;

  n_Id       Number(18);
  d_����ʱ�� Date;
  v_�����   Varchar2(20);
  d_���ʱ�� Date;
  n_�������� Number;
  n_������� Number(2);
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Id       := j_Json.Get_Number('rcpdtl_id');
  d_����ʱ�� := To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  v_�����   := j_Json.Get_String('audit_operator');
  d_���ʱ�� := To_Date(j_Json.Get_String('fee_audit_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_�������� := j_Json.Get_Number('oper_type');
  --Int�Զ����� := j_Json.Get_Number('auto_stuff_return');
  n_������� := j_Json.Get_Number('request_type');

  Zl_���˷�������_Cancel_s(n_Id, d_����ʱ��, v_�����, d_���ʱ��, n_��������, n_�������);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Cancelacc_Reaudit;
/


Create Or Replace Procedure Zl_Exsesvr_Getrequestcancel
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --��ѯ���������¼
  --���      json
  --  input      ��ѯ�Ƿ�������������¼
  --    query_type          N 1 ��ѯ��ʽ:0-���ݷ���ID��ѯ,1-���ݲ��˱䶯��¼��ѯ(ת��������)
  --    rcpdtl_id           C 0 ������ϸid,[����]��[1,2,3],��ѯ��ʽ=0ʱ��Ч
  --    request_type        N 0 �������,��ѯ��ʽ=0ʱ��Ч
  --    cancel_status       N 1 ����״̬,��ѯ��ʽ=0ʱ��Ч
  --    change_id_old       N 0 ԭ�����ı䶯��¼��ID,��ѯ��ʽ=1ʱ��Ч
  --    change_id_new       N 0 Ŀ�겡���ı䶯��¼��ID,��ѯ��ʽ=1ʱ��Ч
  --����      json
  -- output
  --   code     C  1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message  C  1   Ӧ����Ϣ��
  --   fee_cancel_list      [����]����������ÿ���������ʼ�¼
  --     rcpdtl_id          N    ������ϸid(����id)
  --     apply_type         N    �������:��ҩƷ��������Ч:0-δִ��;1-��ִ��;��ҩƷ�����Ĺ̶���Ϊ0
  --     apply_time         N    ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --     aplnt_name         N    ������
  --     apply_dept_id      N    ���벿��id
  --     apply_dept_name    N    ���벿������
  --     audit_dept_id      N    ��˲���id;
  --     audit_dept_name    N    ��˲�������
  --     bill_no            N    ���õ��ݺ�
  --     item_id            N    �շ�ϸĿid
  --     item_name          N    �շ���Ŀ����
  --     quantity           N    ����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  j_Jsonlist Pljson_List;

  n_��ѯ���� Number(2);

  n_������� Number(1);
  n_״̬     Number(1);
  l_Feelist  t_NumList := t_NumList();

  n_ԭ�䶯id   ���ñ䶯��¼.ԭ�䶯id%Type;
  n_Ŀ��䶯id ���ñ䶯��¼.Ŀ��䶯id%Type;

  n_Firstitem Number(1);
  v_Temp      Varchar2(32767);
  c_Temp      Clob;
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ���� := j_Json.Get_Number('query_type');

  n_Firstitem := 1;
  v_Temp      := '{"output":{"code":1,"message":"�ɹ�","fee_cancel_list":[';
  If Nvl(n_��ѯ����, 0) = 0 Then
    n_������� := j_Json.Get_Number('request_type');
    n_״̬     := j_Json.Get_Number('cancel_status');
    j_Jsonlist := j_Json.Get_Pljson_List('rcpdtl_id');
    If j_Jsonlist Is Not Null Then
      --������id��ѯ
      For I In 1 .. j_Jsonlist.Count Loop
        l_Feelist.Extend;
        l_Feelist(l_Feelist.Count) := j_Jsonlist.Get_Number(I);
      End Loop;
    
      For r_���� In (Select /*+cardinality(b,10)*/
                    a.����id
                   From ���˷������� A, Table(l_Feelist) B
                   Where a.����id = b.Column_Value And a.������� = n_������� And a.״̬ = n_״̬) Loop
      
        If Nvl(n_Firstitem, 0) = 0 Then
          v_Temp := v_Temp || ',';
        Else
          n_Firstitem := 0;
        End If;
      
        v_Temp := v_Temp || '{';
        v_Temp := v_Temp || '"rcpdtl_id":' || r_����.����id;
        v_Temp := v_Temp || '}';
      
        If Length(v_Temp) > 30000 Then
          c_Temp := c_Temp || To_Clob(v_Temp);
          v_Temp := '';
        End If;
      End Loop;
    End If;
  
  Elsif Nvl(n_��ѯ����, 0) = 1 Then
    n_ԭ�䶯id   := j_Json.Get_Number('change_id_old');
    n_Ŀ��䶯id := j_Json.Get_Number('change_id_new');
  
    For r_���� In (Select a.�������, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.������, a.���벿��id, e.���� As ���벿��, a.��˲���id,
                        f.���� As ��˲���, b.No, a.�շ�ϸĿid, c.���� As �շ���Ŀ, Sum(a.����) As ����
                 From ���˷������� A, ���ñ䶯��¼ B, �շ���ĿĿ¼ C, ���ű� E, ���ű� F
                 Where a.����id = b.����id And a.�շ�ϸĿid = c.Id And a.���벿��id = e.Id And a.��˲���id = f.Id And b.ԭ�䶯id = n_ԭ�䶯id And
                       b.Ŀ��䶯id = n_Ŀ��䶯id And b.״̬ = 2 And a.״̬ In (0, 2)
                 Group By a.�������, a.����ʱ��, a.������, a.���벿��id, e.����, a.��˲���id, f.����, b.No, a.�շ�ϸĿid, c.����
                 Order By NO, �շ�ϸĿid) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
    
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"apply_type":' || r_����.�������;
      v_Temp := v_Temp || ',"apply_time":"' || zlJsonStr(r_����.����ʱ��) || '"';
      v_Temp := v_Temp || ',"aplnt_name":"' || zlJsonStr(r_����.������) || '"';
      v_Temp := v_Temp || ',"apply_dept_id":' || r_����.���벿��id;
      v_Temp := v_Temp || ',"apply_dept_name":"' || r_����.���벿�� || '"';
      v_Temp := v_Temp || ',"audit_dept_id":' || r_����.��˲���id;
      v_Temp := v_Temp || ',"audit_dept_name":"' || r_����.��˲��� || '"';
      v_Temp := v_Temp || ',"bill_no":"' || r_����.No || '"';
      v_Temp := v_Temp || ',"item_id":' || r_����.�շ�ϸĿid;
      v_Temp := v_Temp || ',"item_name":"' || zlJsonStr(r_����.�շ���Ŀ) || '"';
      v_Temp := v_Temp || ',"quantity":' || zlJsonStr(r_����.����, 1);
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 30000 Then
        c_Temp := c_Temp || To_Clob(v_Temp);
        v_Temp := '';
      End If;
    End Loop;
  End If;
  v_Temp := v_Temp || ']}}';

  If c_Temp Is Not Null Then
    Json_Out := c_Temp || To_Clob(v_Temp);
  Else
    Json_Out := v_Temp;
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrequestcancel;
/


Create Or Replace Procedure Zl_Exsesvr_Getremainmoney
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) Is
  --��ȡ���˷������
  ---------------------------------------------------------------------------
  --input      ��ȡ���˷������
  --  pati_id                 N  1  ����ID
  --  pati_pageid             N  1  ��ҳID
  --  insure_account_balance  N  1  ҽ���˻����
  --  query_type              N  0  ��ѯ��ʽ��0-�������˲�ѯ��1-������ѯ������2-������ѯ������������ò�����Ϣ
  --  pati_ids                C  0  ������ѯ���˹ؼ���Ϣƴ�������ָ�ʽ��1-����ID1:��ҳID1,����ID2:��ҳID2,....��2-����ID1,����ID2,....
  --  fee_source              N  1  ������Դ��0-�����֣�1-���2-סԺ����ѯ��ʽ=1�ҽ�������ID��ѯʱ��Ч
  --output
  --  code                    C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                 C  1  Ӧ����Ϣ
  --  remain_money            N     ʣ���
  --  guarantee_money         N     ������
  --  expected_money          N     Ԥ�����
  --  prepay_money            N  0  Ԥ�����
  --  nobalance_money         N  0  δ����ý��
  --  item_list[]����������������Ϣʱ�ŷ��أ����б���Բ�����
  --       pati_id            N 1 ����id
  --       pati_pageid        N 1 ��ҳid
  --       prepay_money       N 0 Ԥ�����
  --       nobalance_money    N 0 δ����ý��
  --       remain_money       N 1 ʣ���
  --       guarantee_money    N 1 ������
  --       pati_type          C 1 ���ò���
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

  n_����id   �������.����id%Type;
  n_��ҳid   ����ģ�����.��ҳid%Type;
  n_������   �������.Ԥ�����%Type;
  n_�ʻ���� �������.Ԥ�����%Type;
  n_ʣ���   �������.Ԥ�����%Type;
  n_Ԥ����� �������.Ԥ�����%Type;
  n_Ԥ����� �������.Ԥ�����%Type;
  n_������� �������.Ԥ�����%Type;

  n_Find     Number(2);
  n_��ѯ��ʽ Number(2);
  l_����ids  t_StrList;
  c_����ids  Clob;
  n_������Դ Number(2);
Begin
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_��ѯ��ʽ := j_Json.Get_Number('query_type');

  --�������˲�ѯ
  If Nvl(n_��ѯ��ʽ, 0) = 0 Then
    n_����id   := j_Json.Get_Number('pati_id');
    n_��ҳid   := j_Json.Get_Number('pati_pageid');
    n_�ʻ���� := j_Json.Get_Number('insure_account_balance');
    n_������   := Zl_Patientsurety(n_����id, n_��ҳid);
  
    n_Find := 0;
    If n_��ҳid > 0 Then
      Select (Nvl(Sum(a.Ԥ�����), 0) - Nvl(Sum(a.�������), 0) + Nvl(Sum(a.Ԥ�����), 0)) As ʣ���, Nvl(Sum(a.Ԥ�����), 0) As Ԥ�����,
             Nvl(Sum(a.Ԥ�����), 0) As Ԥ�����, Nvl(Sum(a.�������), 0) As �������
      Into n_ʣ���, n_Ԥ�����, n_Ԥ�����, n_�������
      From (Select ����id, Ԥ�����, �������, 0 As Ԥ�����
             From �������
             Where ���� = 1 And ���� = 2 And ����id = n_����id
             Union All
             Select a.����id, 0, 0, Sum(���)
             From ����ģ����� A
             Where a.����id = n_����id And a.��ҳid = n_��ҳid
             Group By a.����id) A;
    
      n_Find := 1;
      v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(n_Ԥ�����, 1);
      v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(n_�������, 1);
      v_Jtmp := v_Jtmp || ',"expected_money":' || zlJsonStr(n_Ԥ�����, 1);
      v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(n_������, 1);
      v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(n_ʣ���, 1);
    Else
      Select (Nvl(Sum(Ԥ�����), 0) - Nvl(Sum(�������), 0) + Nvl(n_�ʻ����, 0)) As ʣ���, Nvl(Sum(Ԥ�����), 0) As Ԥ�����,
             Nvl(Sum(�������), 0) As �������
      Into n_ʣ���, n_Ԥ�����, n_�������
      From �������
      Where ���� = 1 And ���� = 1 And ����id = n_����id;
    
      n_Find := 1;
      v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(n_Ԥ�����, 1);
      v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(n_�������, 1);
      v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(n_������, 1);
      v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(n_ʣ���, 1);
    End If;
  
    If n_Find = 0 Then
      If n_��ҳid > 0 Then
        n_�ʻ���� := 0;
      End If;
      v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(n_�ʻ����, 1);
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"' || v_Jtmp || '}}';
    Return;
  End If;

  --������ѯ
  v_Jtmp     := Null;
  c_����ids  := j_Json.Get_Clob('pati_ids');
  n_������Դ := Nvl(j_Json.Get_Number('fee_source'), 0);
  l_����ids  := t_StrList();
  While c_����ids Is Not Null Loop
    If Length(c_����ids) <= 4000 Then
      l_����ids.Extend;
      l_����ids(l_����ids.Count) := c_����ids;
      c_����ids := Null;
    Else
      l_����ids.Extend;
      l_����ids(l_����ids.Count) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
      c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
    End If;
  End Loop;

  If n_��ѯ��ʽ = 1 Then
    For I In 1 .. l_����ids.Count Loop
      If Instr(l_����ids(I), ':') = 0 Then
        --��ʽ��2-����ID1,����ID2,....
        For R In (Select /*+cardinality(b,10)*/
                   a.����id, Nvl(Sum(a.Ԥ�����), 0) As Ԥ�����, Nvl(Sum(a.�������), 0) As �������,
                   Nvl(Sum(a.Ԥ�����), 0) - Nvl(Sum(a.�������), 0) As ʣ���
                  From ������� A, Table(f_Num2List(l_����ids(I))) B
                  Where a.����id = b.Column_Value And a.���� = 1 And Decode(n_������Դ, 0, 0, a.����) = n_������Դ
                  Group By a.����id) Loop
        
          v_Jtmp := v_Jtmp || ',{"pati_id":' || r.����id;
          v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(r.ʣ���, 1);
          v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(r.Ԥ�����, 1);
          v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(r.�������, 1);
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
      
      Else
        --��ʽ��1-����ID1:��ҳID1,����ID2:��ҳID2,....
        For R In (Select a.����id, a.��ҳid, Nvl(Sum(a.Ԥ�����), 0) As Ԥ�����, Nvl(Sum(a.�������), 0) As �������,
                         (Nvl(Sum(a.Ԥ�����), 0) - Nvl(Sum(a.�������), 0) + Nvl(Sum(a.Ԥ�����), 0)) As ʣ���
                  From (Select n.��ҳid, a.����id, a.Ԥ�����, a.�������, 0 As Ԥ�����
                         From ������� A,
                              (Select /*+cardinality(f,10)*/
                                 To_Number(f.C1) As ����id, To_Number(f.C2) As ��ҳid
                                From Table(f_Num2List2(l_����ids(I))) F) N
                         Where a.���� = 1 And a.���� = 2 And a.����id = n.����id
                         Union All
                         Select a.��ҳid, a.����id, 0, 0, Sum(a.���)
                         From ����ģ����� A,
                              (Select /*+cardinality(f,10)*/
                                 To_Number(f.C1) As ����id, To_Number(f.C2) As ��ҳid
                                From Table(f_Num2List2(l_����ids(I))) F) N
                         Where a.����id = n.����id And a.��ҳid = n.��ҳid
                         Group By a.����id, a.��ҳid) A
                  Group By a.����id, a.��ҳid
                  Having Nvl(Sum(a.Ԥ�����), 0) - Nvl(Sum(a.�������), 0) + Nvl(Sum(a.Ԥ�����), 0) <> 0) Loop
        
          v_Jtmp := v_Jtmp || ',{"pati_id":' || r.����id;
          v_Jtmp := v_Jtmp || ',"pati_pageid":' || r.��ҳid;
          v_Jtmp := v_Jtmp || ',"remain_money":' || zlJsonStr(r.ʣ���, 1);
          v_Jtmp := v_Jtmp || ',"prepay_money":' || zlJsonStr(r.Ԥ�����, 1);
          v_Jtmp := v_Jtmp || ',"nobalance_money":' || zlJsonStr(r.�������, 1);
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
      
      End If;
    End Loop;
  End If;

  If n_��ѯ��ʽ = 2 Then
    For I In 1 .. l_����ids.Count Loop
      For R In (Select n.����id, n.��ҳid, Zl_Patiwarnscheme(n.����id) As ���ò���, a.������
                From (Select a.����id, a.��ҳid, Sum(a.������) As ������
                       From ���˵�����¼ A,
                            (Select /*+cardinality(f,10)*/
                               To_Number(f.C1) As ����id, To_Number(f.C2) As ��ҳid
                              From Table(f_Num2List2(l_����ids(I))) F) N
                       Where a.����id = n.����id And Nvl(a.��ҳid, 0) = n.��ҳid And (a.����ʱ�� Is Null Or a.����ʱ�� > Sysdate) And
                             a.ɾ����־ = 1
                       Group By a.����id, a.��ҳid) A,
                     (Select /*+cardinality(f,10)*/
                        To_Number(f.C1) As ����id, To_Number(f.C2) As ��ҳid
                       From Table(f_Num2List2(l_����ids(I))) F) N
                Where n.����id = a.����id(+) And n.��ҳid = a.��ҳid(+)) Loop
      
        v_Jtmp := v_Jtmp || ',{"pati_id":' || r.����id;
        v_Jtmp := v_Jtmp || ',"pati_pageid":' || r.��ҳid;
        v_Jtmp := v_Jtmp || ',"guarantee_money":' || zlJsonStr(r.������, 1);
        v_Jtmp := v_Jtmp || ',"pati_type":"' || zlJsonStr(r.���ò���) || '"';
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
    End Loop;
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getremainmoney;
/
 
Create Or Replace Procedure Zl_Exsesvr_Getnextno
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
  --    quantity            N   0   ����no�ŵĸ��������ֻȡһ���òβ����򶼴�0 
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    next_no             C   1   ��һ������,quantity>1 ʱ����ʾȡ�������,���ʱ�ö��ŷ���
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_No     Varchar2(64);
  n_���   Number(10);
  n_����id Number(18);
  n_����   Number;
  v_Nos    Varchar2(32767);
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_���   := j_Json.Get_Number('item_num');
  n_����id := j_Json.Get_Number('dept_id');
  n_����   := j_Json.Get_Number('quantity');

  If Nvl(n_����, 0) > 1 Then
    For I In 1 .. n_���� Loop
      Select Zl_Exse_Nextno(n_���, n_����id) Into v_No From Dual;
      v_Nos := v_Nos || ',' || v_No;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","next_no":"' || Substr(v_Nos, 2) || '"}}';
  Else
    Select Zl_Exse_Nextno(n_���, n_����id) Into v_No From Dual;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","next_no":"' || v_No || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnextno;
/


Create Or Replace Procedure Zl_Exsesvr_Cancelacc_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --�������ʼ�¼�˲�
  ---------------------------------------------------------------------------
  --input     �������ʼ�¼�˲�
  --  check_people  C 1 �˲���
  --  check_time    D 1 �˲�ʱ��
  --  request_type  N   �������0-δִ��;1-��ִ��;��ҩƷ�����Ĺ̶���Ϊ0
  --  rcpdtl_list     [����]ÿ��������ϸ��Ϣ
  --    rcpdtl_id     N 1 ������ϸid(����id)
  --    request_time  D 1 ����ʱ��
  --output
  --  code         C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message      C 1 Ӧ����Ϣ��
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  j_List     Pljson_List;
  j_Temp     PLJson;
  n_Id       Number(18);
  d_����ʱ�� Date;
  v_�˲���   Varchar2(20);
  d_�˲����� Date;
  n_������� Number(2); --0-δִ��;1-��ִ��;��ҩƷ�����Ĺ̶���Ϊ0
Begin
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  v_�˲���   := j_Json.Get_String('check_people');
  d_�˲����� := To_Date(j_Json.Get_String('check_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_������� := j_Json.Get_Number('request_type');
  If n_������� Is Null Then
    n_������� := 1;
  End If;

  j_List := j_Json.Get_Pljson_List('rcpdtl_list');
  For J In 1 .. j_List.Count Loop
    j_Temp     := PLJson();
    j_Temp     := PLJson(j_List.Get(J));
    n_Id       := j_Temp.Get_Number('rcpdtl_id');
    d_����ʱ�� := To_Date(j_Temp.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  
    Zl_���˷�������_Check(n_Id, d_����ʱ��, v_�˲���, d_�˲�����, n_�������);
  End Loop;

  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Cancelacc_Check;
/
 
Create Or Replace Procedure Zl_Exsesvr_Setsendwin
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����÷���ҩƷ���ݵķ�ҩ����
  --��Σ�Json_In:��ʽ
  --  input
  --    pharmacy_id              N   1  �ⷿid
  --    pharmacy_window_old      C   1  �ɷ�ҩ����
  --    pharmacy_window_new      C   1  �·�ҩ����
  --    bill_list[]
  --      billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
  --      rcp_no                 C   1 ����No
  --����: Json_Out,��ʽ����
  --  output
  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ------------------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json      PLJson;
  j_Bill_List Pljson_List;
  n_�ⷿid    Number(18);
  v_�ɴ���    Varchar2(50);
  v_�´���    Varchar2(50);
  n_����      Number(1);
  v_No        Varchar2(20);
  n_Count     Number;
Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�ⷿid    := j_Json.Get_Number('pharmacy_id');
  v_�ɴ���    := j_Json.Get_String('pharmacy_window_old');
  v_�´���    := j_Json.Get_String('pharmacy_window_new');
  j_Bill_List := j_Json.Get_Pljson_List('bill_list');
  n_Count     := j_Bill_List.Count;

  If n_Count > 0 Then
    For I In 1 .. n_Count Loop
      o_Json := PLJson(j_Bill_List.Get(I));
      n_���� := o_Json.Get_Number('billtype');
      v_No   := o_Json.Get_String('rcp_no');
    
      Update ������ü�¼
      Set ��ҩ���� = v_�´���
      Where ִ�в���id = n_�ⷿid And ��¼���� = n_���� And NO = v_No And ��ҩ���� = v_�ɴ���;
    
      Update סԺ���ü�¼
      Set ��ҩ���� = v_�´���
      Where ִ�в���id = n_�ⷿid And ��¼���� = n_���� And NO = v_No And ��ҩ���� = v_�ɴ���;
    End Loop;
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Setsendwin;
/

Create Or Replace Procedure Zl_Exsesvr_Getnobyinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ���Ʊ�ݺŷ�ҩ����ҩ��ͨ��¼�뷢Ʊ�Ż�ȡ��Ӧ��ҩƷ����NO
  --��Σ�json��ʽ
  --Input
  --   invc_no  C  1  Ʊ�ݺ�
  --���Σ�json��ʽ
  --Json_Out
  --  code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  "Ӧ����Ϣ�� �ɹ�ʱ���ش���No��[����] ʧ��ʱ���ؾ���Ĵ�����Ϣ"
  --  rcp_nos  C  1 �������ݺţ�����ö��ŷָ� 
  -------------------------------------------------------------------------------------------------
  v_Ʊ�ݺ� Ʊ��ʹ����ϸ.����%Type;

  v_Tmp   Varchar2(32767);
  j_Input PLJson;
  j_Json  PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Ʊ�ݺ� := j_Json.Get_String('invc_no');
  For v_ҩƷ����no In (Select Distinct a.No
                   From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B
                   Where a.Id = b.��ӡid And a.�������� = 1 And b.Ʊ�� = 1 And b.���� = v_Ʊ�ݺ�) Loop
  
    v_Tmp := Nvl(v_Tmp, '') || ',' || v_ҩƷ����no.No;
  End Loop;
  If v_Tmp Is Not Null Then
    v_Tmp := Substr(v_Tmp, 2);
  End If;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�","rcp_nos":"' || Nvl(v_Tmp, '') || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnobyinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Getchargeoffinfo
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����������ѯ����������Ϣ
  --���      json
  --  input      ����������ѯ����������Ϣ
  --    audit_dept_id       N    ��˲���ID(ҩ��)
  --    request_begin_time  D    ���뿪ʼʱ��
  --    request_end_time    D    �������ʱ��
  --    audit_begin_time    D    ��˿�ʼʱ��
  --    audit_end_time      D    ��˽���ʱ��
  --    cancel_status       N  1 ״̬
  --    request_dept_id     N    ���벿��ID
  --    request_operator    C    ������
  --    pati_id             N    ����ID
  --    cancel_condition    C    ��������
  --    cancel_check        N    �˲飨ѡ�����������������Ҫ�˲顿ʱ���룬0-δ�˲� 1-�Ѻ˲飩
  --    rcpdtl_id          C     ������ϸid,[����]��[1,2,3]
  --    request_dept_ids   C     ���벿��id��������������ѯ
  --    item_ids           C     �շ�ϸĿid��,����������ѯ
  --    request_type       N     �������-1-������;0-δִ��;1-��ִ��
  --����      json
  -- output
  --   code     C  1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message  C  1   Ӧ����Ϣ��
  --   fee_cancel_list      [����]����������ÿ���������ʼ�¼
  --     rcpdtl_id          N    ������ϸid(����id)
  --     request_type       N    �������
  --     item_id            N    �շ�ϸĿid
  --     request_dept_id    N    ���벿��id
  --     request_dept       C    ���벿��
  --     audit_dept_id      N    ��˲���id
  --     quantity           N    ����
  --     request_operator   C    ������
  --     request_time       D    ����ʱ��
  --     auditor            C    �����
  --     audit_time         D    ���ʱ��
  --     cancel_status      N    ״̬
  --     cancel_reason      C    ����ԭ��
  --     checker            C    �˲���
  --     price_retail       N    ���ۼ�
  --     advice_id          N    ҽ��id
  --     pati_id            N    ����ID
  --     pati_name          C    ��������
  --     inpatient_num      C    סԺ��
  --     pati_pageid        N    ��ҳid
  ---------------------------------------------------------------------------

  v_Sql Varchar2(4000);

  j_Input Pljson;
  j_Json  Pljson;

  j_Jsonlist Pljson_List := Pljson_List();

  n_��˲���id   Number(18);
  v_���뿪ʼʱ�� Varchar2(50);
  v_�������ʱ�� Varchar2(50);
  v_��˿�ʼʱ�� Varchar2(50);
  v_��˽���ʱ�� Varchar2(50);
  n_״̬         Number(1);
  n_���벿��id   Number(18);
  v_������       Varchar2(20);
  n_����id       Number(18);
  v_��������     Varchar2(32767); --����ʱ��,����id|����ʱ��,����id...
  n_�˲�         Number(1); --״̬=0ʱʹ��
  v_���벿��ids  Varchar2(32767);
  v_�շ���Ŀids  Varchar2(32767);
  n_�������     Number(2);

  n_Count   Number := 0;
  l_Feelist t_Numlist := t_Numlist();

  v_Output Varchar2(32767);
  c_Output Clob;

  Type t_������Ϣ Is Ref Cursor;
  c_������Ϣ t_������Ϣ; --��̬�α����

  Cursor c_������Ϣ Is
    Select a.����id, a.�������, a.�շ�ϸĿid, a.���벿��id, c.����, a.��˲���id, a.����, a.������, a.����ʱ��, a.�����, a.���ʱ��, a.״̬, a.����ԭ��, a.�˲���,
           b.��׼����, b.ҽ�����, b.����id, b.����, b.��ʶ��, b.��ҳid
    From ���˷������� A, סԺ���ü�¼ B, ���ű� C
    Where a.����id = b.Id And a.���벿��id = c.Id And a.������� = 1 And a.��˲���id = 0 And a.����id = 0 And a.״̬ = 0 And
          a.����� Is Null;
  r_���� c_������Ϣ%RowType;

Begin
  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��˲���id   := j_Json.Get_Number('audit_dept_id');
  v_���뿪ʼʱ�� := j_Json.Get_String('request_begin_time');
  v_�������ʱ�� := j_Json.Get_String('request_end_time');
  v_��˿�ʼʱ�� := j_Json.Get_String('audit_begin_time');
  v_��˽���ʱ�� := j_Json.Get_String('audit_end_time');
  n_״̬         := j_Json.Get_Number('cancel_status');
  n_���벿��id   := j_Json.Get_Number('request_dept_id');
  v_������       := j_Json.Get_String('request_operator');
  n_����id       := j_Json.Get_Number('pati_id');
  v_��������     := j_Json.Get_String('cancel_condition');
  n_�˲�         := j_Json.Get_Number('cancel_check');
  v_���벿��ids  := j_Json.Get_String('request_dept_ids');
  v_�շ���Ŀids  := j_Json.Get_String('item_ids');
  n_�������     := Nvl(j_Json.Get_Number('request_type'), 1);

  j_Jsonlist := j_Json.Get_Pljson_List('rcpdtl_id');

  v_Output := Null;
  If j_Jsonlist Is Not Null Then
    --������id��ѯ
    n_Count := j_Jsonlist.Count;
  
    For I In 1 .. n_Count Loop
      l_Feelist.Extend;
      l_Feelist(l_Feelist.Count) := j_Jsonlist.Get_Number(I);
    End Loop;
  
    For c_�������� In (Select a.����id, a.�������, a.�շ�ϸĿid, a.���벿��id, c.����, a.��˲���id, a.����, a.������, a.����ʱ��, a.�����, a.���ʱ��, a.״̬,
                          a.����ԭ��, a.�˲���, b.��׼����, b.ҽ�����, b.����id, b.����, b.��ʶ��, b.��ҳid
                   From ���˷������� A, סԺ���ü�¼ B, ���ű� C
                   Where a.����id = b.Id And a.���벿��id = c.Id And a.������� = Decode(n_�������, -1, a.�������, n_�������) And
                         a.״̬ = n_״̬ And a.��˲���id = Nvl(n_��˲���id, a.��˲���id) And
                         a.����id In (Select /*+cardinality(j,10)*/
                                     Column_Value
                                    From Table(l_Feelist) J)) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      Zljsonputvalue(v_Output, 'rcpdtl_id', c_��������.����id, 1, 1);
      Zljsonputvalue(v_Output, 'request_type', c_��������.�������, 1);
      Zljsonputvalue(v_Output, 'item_id', c_��������.�շ�ϸĿid, 1);
      Zljsonputvalue(v_Output, 'request_dept_id', c_��������.���벿��id, 1);
      Zljsonputvalue(v_Output, 'request_dept', c_��������.����, 0);
      Zljsonputvalue(v_Output, 'audit_dept_id', c_��������.��˲���id, 1);
      Zljsonputvalue(v_Output, 'quantity', c_��������.����, 1);
      Zljsonputvalue(v_Output, 'request_operator', c_��������.������, 0);
      Zljsonputvalue(v_Output, 'request_time', To_Char(c_��������.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'auditor', c_��������.�����, 0);
      Zljsonputvalue(v_Output, 'audit_time', To_Char(c_��������.���ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'cancel_status', c_��������.״̬, 1);
      Zljsonputvalue(v_Output, 'cancel_reason', c_��������.����ԭ��, 0);
      Zljsonputvalue(v_Output, 'checker', c_��������.�˲���, 0);
      Zljsonputvalue(v_Output, 'price_retail', c_��������.��׼����, 1);
      Zljsonputvalue(v_Output, 'advice_id', c_��������.ҽ�����, 1);
      Zljsonputvalue(v_Output, 'pati_id', c_��������.����id, 1);
      Zljsonputvalue(v_Output, 'pati_name', c_��������.����, 0);
      Zljsonputvalue(v_Output, 'inpatient_num', c_��������.��ʶ��, 0);
      Zljsonputvalue(v_Output, 'pati_pageid', c_��������.��ҳid, 1, 2);
    
    End Loop;
  Else
  
    v_Sql := Nvl(v_Sql, '') || '   Select A.����id, A.�������, A.�շ�ϸĿid, A.���벿��id, C.����, A.��˲���id, ' ||
             '   A.����, A.������, A.����ʱ�� , A.����� , A.���ʱ�� , A.״̬ , A.����ԭ�� ,' ||
             '   A.�˲��� , B.��׼���� , B.ҽ����� , B.����id , B.���� , B.��ʶ��,b.��ҳid ';
    v_Sql := v_Sql || '   From ���˷������� A, סԺ���ü�¼ B,���ű� C';
    If v_�������� Is Not Null Then
      v_Sql := v_Sql || '   ,Table(f_Str2list2(''' || v_�������� || ''', ''| '', '','')) T';
    End If;
    v_Sql := v_Sql || '   Where A.����id = B.Id And A.���벿��id = C.Id';
  
    If Nvl(n_��˲���id, 0) <> 0 Then
      v_Sql := v_Sql || ' And A.��˲���id = ' || n_��˲���id;
    End If;
  
    If n_������� <> -1 Then
      v_Sql := v_Sql || ' And A.�������=' || n_�������;
    End If;
  
    If n_״̬ = 0 Then
      v_Sql := v_Sql || '   And  A.״̬ = 0 And A.����� Is Null';
      If n_�˲� Is Not Null Then
        If n_�˲� = 0 Then
          v_Sql := v_Sql || '   And A.�˲��� Is Null ';
        Else
          v_Sql := v_Sql || '   And A.�˲��� Is Not Null ';
        End If;
      End If;
    
      If v_���뿪ʼʱ�� Is Not Null Then
        v_Sql := v_Sql || '   And A.����ʱ�� Between to_date(''' || v_���뿪ʼʱ�� ||
                 ''',''yyyy-mm-dd hh24:mi:ss'') And to_date(''' || v_�������ʱ�� || ''',''yyyy-mm-dd hh24:mi:ss'') ';
      End If;
    Else
      v_Sql := v_Sql || '   And  A.״̬ <> 0 And A.����� Is Not Null ';
      If v_��˿�ʼʱ�� Is Not Null Then
        v_Sql := v_Sql || '   And A.���ʱ�� Between to_date(''' || v_��˿�ʼʱ�� ||
                 ''',''yyyy-mm-dd hh24:mi:ss'') And to_date(''' || v_��˽���ʱ�� || ''',''yyyy-mm-dd hh24:mi:ss'') ';
      End If;
    End If;
  
    If n_���벿��id Is Not Null Then
      v_Sql := v_Sql || '   And A.���벿��id = ' || n_���벿��id;
    End If;
  
    If v_���벿��ids Is Not Null Then
      v_Sql := v_Sql || '   And Instr('',' || v_���벿��ids || ','', '','' || A.���벿��id || '','') > 0 ';
    End If;
  
    If v_�շ���Ŀids Is Not Null Then
      v_Sql := v_Sql || '   And Instr('',' || v_�շ���Ŀids || ','', '','' || A.�շ�ϸĿid || '','') > 0 ';
    End If;
  
    If v_������ Is Not Null Then
      v_Sql := v_Sql || '   And A.������ = ''' || v_������ || '''';
    End If;
  
    If n_����id Is Not Null Then
      v_Sql := v_Sql || '   And B.����ID = ' || n_����id;
    End If;
  
    If v_�������� Is Not Null Then
      v_Sql := v_Sql || '   And A.����ʱ�� = To_Date(t.C1,''yyyy-mm-dd hh24:mi:ss'') And B.����ID = t.C2';
    End If;
  
    Open c_������Ϣ For v_Sql;
    Loop
      Fetch c_������Ϣ
        Into r_����;
      Exit When c_������Ϣ %NotFound;
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
        v_Output := Null;
      End If;
    
      Zljsonputvalue(v_Output, 'rcpdtl_id', r_����.����id, 1, 1);
      Zljsonputvalue(v_Output, 'request_type', r_����.�������, 1);
      Zljsonputvalue(v_Output, 'item_id', r_����.�շ�ϸĿid, 1);
      Zljsonputvalue(v_Output, 'request_dept_id', r_����.���벿��id, 1);
      Zljsonputvalue(v_Output, 'request_dept', r_����.����, 0);
      Zljsonputvalue(v_Output, 'audit_dept_id', r_����.��˲���id, 1);
      Zljsonputvalue(v_Output, 'quantity', r_����.����, 1);
      Zljsonputvalue(v_Output, 'request_operator', r_����.������, 0);
      Zljsonputvalue(v_Output, 'request_time', To_Char(r_����.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'auditor', r_����.�����, 0);
      Zljsonputvalue(v_Output, 'audit_time', To_Char(r_����.���ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
      Zljsonputvalue(v_Output, 'cancel_status', r_����.״̬, 1);
      Zljsonputvalue(v_Output, 'cancel_reason', r_����.����ԭ��, 0);
      Zljsonputvalue(v_Output, 'checker', r_����.�˲���, 0);
      Zljsonputvalue(v_Output, 'price_retail', r_����.��׼����, 1);
      Zljsonputvalue(v_Output, 'advice_id', r_����.ҽ�����, 1);
      Zljsonputvalue(v_Output, 'pati_id', r_����.����id, 1);
      Zljsonputvalue(v_Output, 'pati_name', r_����.����, 0);
      Zljsonputvalue(v_Output, 'inpatient_num', r_����.��ʶ��, 0);
      Zljsonputvalue(v_Output, 'pati_pageid', r_����.��ҳid, 1, 2);
    End Loop;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := Null;
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","fee_cancel_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_cancel_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getchargeoffinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Getbilldetailinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ���ȡ��ҩƷ��ҩҵ����صķ�����Ϣ����Ҫ���ڽ�����ʾ
  --��Σ�json��ʽ
  --Input
  --   fee_ids    C     ����id��֧�ֶ��id����ʽ�� ����id,����id,��
  --   bill_nos   C     ����no,��¼���ʣ���ʽ: no,��¼����|,...
  --���Σ�json��ʽ
  --Json_Out
  --fee_list      [����]ÿ������ID��Ϣ
  --  bill_prop           N    ��¼����:1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�;4-�Һŵ�;5-���￨;6-Ԥ����
  --  bill_no             C    ���ݺ�
  --  fee_id              N    ������ϸid(����id)
  --  fee_num             N    ���
  --  iden_id             N    ��ʶ��
  --  pati_bed            C    ����
  --  fee_ampaid          N    ʵ�ս��
  --  packages_num        N    ����
  --  quantity            N    ����
  --  placer              C    ������
  --  operator_code       C    ����Ա���
  --  operator_name       C    ����Ա����
  --  create_time         D    �Ǽ�ʱ��
  --  happen_time         D    ����ʱ��
  --  rcp_type            N    �������(������NO��˵��1-��ҩ��2-��ҩ��3-���)
  --  fee_type            C    �ѱ�
  --  rec_status          N    ��¼״̬
  --  register_id         N    �Һ�id
  --  register_no         C    �Һ�NO
  --  register_time       D    �ҺŵǼ�ʱ��
  --  income_item_id      N    ������Ŀid
  --  fee_origin          N    ������Դ(1-������ã�2-סԺ����)
  --  bill_deptid         N    ��������id
  --  order_id            N    ҽ��ID
  --  fee_item_id         N    �շ�ϸĿid
  --  fee_status         N    ����״̬
  -------------------------------------------------------------------------------------------------
  v_Output Varchar2(32767);
  c_Output Clob;

  v_����id Varchar2(32767); --����id
  v_Nos    Varchar2(32767);

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  l_����id Collection_Type;
  l_No     Collection_Type;

  j_Input Pljson;
  j_Json  Pljson;
Begin
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  If j_Json.Exist('bill_nos') Then
    v_Nos := j_Json.Get_String('bill_nos');
    --�� v_Nos ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Str2list ��������
    While v_Nos Is Not Null Loop
      If Length(v_Nos) <= 4000 Then
        l_No(l_No.Count) := v_Nos;
        v_Nos := Null;
      Else
        l_No(l_No.Count) := Substr(v_Nos, 1, Instr(v_Nos, '|', 3980) - 1);
        v_Nos := Substr(v_Nos, Instr(v_Nos, '|', 3980) + 1);
      End If;
    End Loop;
  
    For I In 0 .. l_No.Count - 1 Loop
      For r_���� In (Select a.��¼����, a.No, a.Id, a.���, a.��ʶ��, '' As ����, a.ʵ�ս��, a.����, a.����, a.������, a.����Ա���, a.����Ա����, a.�Ǽ�ʱ��,
                          a.����ʱ��, Zl_Get�շ����(a.��¼����, a.No, a.ִ�в���id) As �շ����, a.�ѱ�, a.��¼״̬, a.������Ŀid, a.�Һ�id,
                          b.No As �Һ�no, b.�Ǽ�ʱ�� As �ҺŵǼ�ʱ��, 1 As ������Դ, a.��������id, a.ҽ�����, a.�շ�ϸĿid, a.����״̬
                   
                   From ������ü�¼ A, ���˹Һż�¼ B,
                        (Select /*+cardinality(J,10)*/
                           C1 As NO, C2 As ��¼����
                          From Table(f_Str2list2(l_No(I), '|', ',')) J) C
                   Where a.�Һ�id = b.Id(+) And a.No = c.No And (a.��¼���� = c.��¼���� Or Nvl(c.��¼����, 0) = 0)
                   Union All
                   Select a.��¼����, a.No, a.Id, a.���, a.��ʶ��, ����, a.ʵ�ս��, a.����, a.����, a.������, a.����Ա���, a.����Ա����, a.�Ǽ�ʱ��,
                          a.����ʱ��, Zl_Get�շ����(a.��¼����, a.No, a.ִ�в���id) As �շ����, a.�ѱ�, a.��¼״̬, a.������Ŀid, 0 As �Һ�id,
                          '' As �Һ�no, Null As �ҺŵǼ�ʱ��, 2 As ������Դ, a.��������id, a.ҽ�����, a.�շ�ϸĿid, a.����״̬
                   From סԺ���ü�¼ A,
                        (Select /*+cardinality(J,10)*/
                           C1 As NO, C2 As ��¼����
                          From Table(f_Str2list2(l_No(I), '|', ',')) J) C
                   Where a.No = c.No And (a.��¼���� = c.��¼���� Or Nvl(c.��¼����, 0) = 0)) Loop
      
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        Zljsonputvalue(v_Output, 'bill_prop', r_����.��¼����, 1, 1);
        Zljsonputvalue(v_Output, 'bill_no', r_����.No, 0);
        Zljsonputvalue(v_Output, 'fee_id', r_����.Id, 1);
        Zljsonputvalue(v_Output, 'fee_num', r_����.���, 1);
        Zljsonputvalue(v_Output, 'iden_id', r_����.��ʶ��, 1);
        Zljsonputvalue(v_Output, 'pati_bed', r_����.����, 0);
        Zljsonputvalue(v_Output, 'fee_ampaid', r_����.ʵ�ս��, 1);
        Zljsonputvalue(v_Output, 'packages_num', Nvl(r_����.����, 1), 1);
        Zljsonputvalue(v_Output, 'quantity', r_����.����, 1);
        Zljsonputvalue(v_Output, 'placer', r_����.������, 0);
        Zljsonputvalue(v_Output, 'operator_code', r_����.����Ա���, 0);
        Zljsonputvalue(v_Output, 'operator_name', r_����.����Ա����, 0);
        Zljsonputvalue(v_Output, 'create_time', To_Char(r_����.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'happen_time', To_Char(r_����.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'rcp_type', r_����.�շ����, 1);
        Zljsonputvalue(v_Output, 'fee_type', r_����.�ѱ�, 0);
        Zljsonputvalue(v_Output, 'rec_status', r_����.��¼״̬, 1);
        Zljsonputvalue(v_Output, 'register_id', r_����.�Һ�id, 1);
        Zljsonputvalue(v_Output, 'register_no', r_����.�Һ�no, 0);
        Zljsonputvalue(v_Output, 'register_time', To_Char(r_����.�ҺŵǼ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'income_item_id', r_����.������Ŀid, 1);
        Zljsonputvalue(v_Output, 'fee_origin', r_����.������Դ, 1);
        Zljsonputvalue(v_Output, 'bill_deptid', r_����.��������id, 1);
        Zljsonputvalue(v_Output, 'order_id', r_����.ҽ�����, 1);
        Zljsonputvalue(v_Output, 'fee_item_id', r_����.�շ�ϸĿid, 1);
        Zljsonputvalue(v_Output, 'fee_status', r_����.����״̬, 1, 2);
      
      End Loop;
    End Loop;
  
  Else
  
    v_����id := j_Json.Get_String('fee_ids');
    --�� v_����id ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Num2list ��������
    While v_����id Is Not Null Loop
      If Length(v_����id) <= 4000 Then
        l_����id(l_����id.Count) := v_����id;
        v_����id := Null;
      Else
        l_����id(l_����id.Count) := Substr(v_����id, 1, Instr(v_����id, ',', 3980) - 1);
        v_����id := Substr(v_����id, Instr(v_����id, ',', 3980) + 1);
      End If;
    End Loop;
  
    For I In 0 .. l_����id.Count - 1 Loop
      For r_���� In (Select /*+cardinality(J,10)*/
                    a.��¼����, a.No, a.Id, a.���, a.��ʶ��, '' As ����, a.ʵ�ս��, a.����, a.����, a.������, a.����Ա���, a.����Ա����, a.�Ǽ�ʱ��,
                    a.����ʱ��, Decode(a.�շ����, '7', 2, 1) As �շ����, a.�ѱ�, a.��¼״̬, a.������Ŀid, a.�Һ�id, b.No As �Һ�no,
                    b.�Ǽ�ʱ�� As �ҺŵǼ�ʱ��, 1 As ������Դ, a.��������id, a.ҽ�����, a.�շ�ϸĿid, a.����״̬
                   From ������ü�¼ A, ���˹Һż�¼ B, Table(f_Num2list(l_����id(I))) J
                   Where a.�Һ�id = b.Id(+) And a.Id = j.Column_Value
                   Union All
                   Select /*+cardinality(J,10)*/
                    a.��¼����, a.No, a.Id, a.���, a.��ʶ��, ����, a.ʵ�ս��, a.����, a.����, a.������, a.����Ա���, a.����Ա����, a.�Ǽ�ʱ��, a.����ʱ��,
                    Decode(a.�շ����, '7', 2, 1) As �շ����, a.�ѱ�, a.��¼״̬, a.������Ŀid, 0 As �Һ�id, '' As �Һ�no, Null As �ҺŵǼ�ʱ��,
                    2 As ������Դ, a.��������id, a.ҽ�����, a.�շ�ϸĿid, a.����״̬
                   From סԺ���ü�¼ A, Table(f_Num2list(l_����id(I))) J
                   Where a.Id = j.Column_Value) Loop
      
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        Zljsonputvalue(v_Output, 'bill_prop', r_����.��¼����, 1, 1);
        Zljsonputvalue(v_Output, 'bill_no', r_����.No, 0);
        Zljsonputvalue(v_Output, 'fee_id', r_����.Id, 1);
        Zljsonputvalue(v_Output, 'fee_num', r_����.���, 1);
        Zljsonputvalue(v_Output, 'iden_id', r_����.��ʶ��, 1);
        Zljsonputvalue(v_Output, 'pati_bed', r_����.����, 0);
        Zljsonputvalue(v_Output, 'fee_ampaid', r_����.ʵ�ս��, 1);
        Zljsonputvalue(v_Output, 'packages_num', Nvl(r_����.����, 1), 1);
        Zljsonputvalue(v_Output, 'quantity', r_����.����, 1);
        Zljsonputvalue(v_Output, 'placer', r_����.������, 0);
        Zljsonputvalue(v_Output, 'operator_code', r_����.����Ա���, 0);
        Zljsonputvalue(v_Output, 'operator_name', r_����.����Ա����, 0);
        Zljsonputvalue(v_Output, 'create_time', To_Char(r_����.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'happen_time', To_Char(r_����.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'rcp_type', r_����.�շ����, 1);
        Zljsonputvalue(v_Output, 'fee_type', r_����.�ѱ�, 0);
        Zljsonputvalue(v_Output, 'rec_status', r_����.��¼״̬, 1);
        Zljsonputvalue(v_Output, 'register_id', r_����.�Һ�id, 1);
        Zljsonputvalue(v_Output, 'register_no', r_����.�Һ�no, 0);
        Zljsonputvalue(v_Output, 'register_time', To_Char(r_����.�ҺŵǼ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0);
        Zljsonputvalue(v_Output, 'income_item_id', r_����.������Ŀid, 1);
        Zljsonputvalue(v_Output, 'fee_origin', r_����.������Դ, 1);
        Zljsonputvalue(v_Output, 'bill_deptid', r_����.��������id, 1);
        Zljsonputvalue(v_Output, 'order_id', r_����.ҽ�����, 1);
        Zljsonputvalue(v_Output, 'fee_item_id', r_����.�շ�ϸĿid, 1);
        Zljsonputvalue(v_Output, 'fee_status', r_����.����״̬, 1, 2);
      End Loop;
    End Loop;
  
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbilldetailinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ���ȡ��ҩƷ��ҩҵ����صķ�����Ϣ����Ҫ���ڽ�����ʾ
  --��Σ�json��ʽ
  --Input
  --   pharmacy_id���ⷿid
  --   fee_nos������no��֧�ֶ��no����ʽ�� ��¼����,no,��
  --���Σ�json��ʽ
  --Json_Out
  --  code          C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message       C   1   Ӧ����Ϣ��
  --  fee_list      C       [����]ÿ������NO��Ϣ
  --    fee_properties      N ��¼����
  --    bill_no             C ����no
  --    real_amount         N ʵ�ս��
  --    rcp_type            N �շ����(������NO��˵��1-��ҩ��2-��ҩ��3-���)
  --    iden_id             C ��ʶ��
  --    placer              C ������
  --    bill_deptid         N ��������id
  --    create_time         D �Ǽ�ʱ��
  --    pati_bed            C ��ǰ����
  --    operator_name       C ����Ա����
  -------------------------------------------------------------------------------------------------
  n_�ⷿid ������ü�¼.ִ�в���id%Type;
  v_����no Varchar2(32767);

  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_����id Collection_Type;
  I          Number;

  v_Output Varchar2(32767);
  c_Output Clob;
  j_Input  PLJson;
  j_Json   PLJson;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�ⷿid := j_Json.Get_Number('pharmacy_id');
  v_����no := j_Json.Get_String('fee_nos');
  I        := 0;
  While v_����no Is Not Null Loop
    If Length(v_����no) <= 4000 Then
      Col_����id(I) := v_����no;
      v_����no := Null;
    Else
      Col_����id(I) := Substr(v_����no, 1, Instr(v_����no, '|', 3980) - 1);
      v_����no := Substr(v_����no, Instr(v_����no, '|', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I := 0;

  For I In 0 .. Col_����id.Count - 1 Loop
  
    For v_������Ϣ In (Select /*+ rule*/
                    NO, ��¼����, Sum(ʵ�ս��) As ʵ�ս��, �շ����, Max(��ʶ��) As ��ʶ��, Max(������) As ������, Max(��������id) As ��������id,
                    Max(�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(����) As ����, Max(����Ա����) As ����Ա����
                   From (Select a.No, a.��¼����, a.ʵ�ս��, Zl_Get�շ����(a.��¼����, a.No, a.ִ�в���id) As �շ����, a.��ʶ��, a.������, a.��������id,
                                 Decode(Nvl(a.��¼״̬, 0), 2, To_Date(Null), a.�Ǽ�ʱ��) As �Ǽ�ʱ��, '' As ����,
                                 Decode(Nvl(a.��¼״̬, 0), 2, Null, a.����Ա����) As ����Ա����
                          From ������ü�¼ A,
                               (Select /*+cardinality(c,10)*/
                                  C1 As ��¼����, C2 As NO
                                 From Table(f_Str2List2(Col_����id(I), '|', ',')) C) C
                          Where a.ִ�в���id = Decode(Nvl(n_�ⷿid, 0), 0, a.ִ�в���id, n_�ⷿid) And Mod(a.��¼����, 10) = c.��¼���� And
                                a.No = c.No And a.��¼���� In (1, 2)
                          Union All
                          Select a.No, a.��¼����, a.ʵ�ս��, Zl_Get�շ����(a.��¼����, a.No, a.ִ�в���id) As �շ����,
                                 Decode(Nvl(�ಡ�˵�, 0), 1, -1 * Null, ��ʶ��) As ��ʶ��, a.������, a.��������id,
                                 Decode(Nvl(a.��¼״̬, 0), 2, To_Date(Null), a.�Ǽ�ʱ��) As �Ǽ�ʱ��,
                                 Decode(Nvl(�ಡ�˵�, 0), 1, '', ����) As ����, Decode(Nvl(a.��¼״̬, 0), 2, Null, a.����Ա����) As ����Ա����
                          From סԺ���ü�¼ A,
                               (Select /*+cardinality(c,10)*/
                                  C1 As ��¼����, C2 As NO
                                 From Table(f_Str2List2(Col_����id(I), '|', ',')) C) C
                          Where a.ִ�в���id = Decode(Nvl(n_�ⷿid, 0), 0, a.ִ�в���id, n_�ⷿid) And Mod(a.��¼����, 10) = c.��¼���� And
                                a.No = c.No And a.��¼���� In (1, 2))
                   Group By NO, ��¼����, �շ����) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
      zlJsonPutValue(v_Output, 'fee_properties', v_������Ϣ.��¼����, 1, 1);
      zlJsonPutValue(v_Output, 'bill_no', v_������Ϣ.No, 0, 0);
      zlJsonPutValue(v_Output, 'real_amount', v_������Ϣ.ʵ�ս��, 1, 0);
      zlJsonPutValue(v_Output, 'rcp_type', v_������Ϣ.�շ����, 1, 0);
      zlJsonPutValue(v_Output, 'iden_id', v_������Ϣ.��ʶ��, 0, 0);
      zlJsonPutValue(v_Output, 'placer', v_������Ϣ.������, 0, 0);
      zlJsonPutValue(v_Output, 'bill_deptid', v_������Ϣ.��������id, 1, 0);
      zlJsonPutValue(v_Output, 'create_time', To_Char(v_������Ϣ.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), 0, 0);
      zlJsonPutValue(v_Output, 'pati_bed', v_������Ϣ.����, 0, 0);
      zlJsonPutValue(v_Output, 'operator_name', v_������Ϣ.����Ա����, 0, 2);
    End Loop;
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Getinsureiteminfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡҽ�����������Ϣ�����ȡ���ƣ��Ƿ������˱���֧����Ŀ
  --��Σ�Json_In:��ʽ
  --  input
  --    insurance_type          N 1 ����
  --    fee_item_id             N 1 �շ�ϸĿID
  --                                ����Ϊ������ȡ���������
  --    fee_item_ids            C 0 �շ�ϸĿids��ȡһ���շ���Ŀ��ҽ����������
  --    insurance_types         C 0 ���ය��ƴ��
  --    query_type              N 0 ��ѯ��ʽ
  --                                   0-����fee_item_ids+insurance_typeȡһ���շ���Ŀ��ҽ����������
  --                                   1-����fee_item_ids+insurance_type���������˵�fee_item_ids ����һ�������˱���֧����Ŀids��
  --                                   2-����fee_item_ids+insurance_types���������˵�fee_item_ids��insurance_type�б�
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    insure_name             C 1 ���մ�������
  --    isexist                 N 1 �������˱���֧����Ŀ,0-δ���ã�1-������
  --    fee_item_ids            C 1 �����˱���֧����Ŀ���շ�ϸĿidƴ��
  --    item_list[]������ȡʱ�ŷ���
  --           fee_item_id      N 1 �շ�ϸĿID
  --           insure_name      C 1 ���մ�������
  --           insure_name_ex   C 1 ��������ٴ�����ѡ������
  --    pay_list[]�����˱���֧����Ŀ�б�query_type=2ʱ����
  --           insurance_type   N 1 ����
  --           fee_item_ids     C 1 �����˱���֧����Ŀids
  ---------------------------------------------------------------------------
  j_Json       Pljson;
  j_Input      Pljson;
  v_Tmp        Varchar2(32767);
  n_Tmp        Number;
  n_�շ�ϸĿid Number;
  n_����       Number;
  v_����s      Varchar2(32767);
  v_Vals       Clob;
  l_Vals       t_Strlist;
  n_��ѯ��ʽ   Number;
  v_Jtmp       Varchar2(32767);
  c_Jtmp       Clob;

Begin
  --�������
  j_Input    := Pljson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_����     := j_Json.Get_Number('insurance_type');
  n_��ѯ��ʽ := j_Json.Get_Number('query_type');

  If j_Json.Exist('fee_item_ids') Then
    v_Vals := j_Json.Get_Clob('fee_item_ids');
    l_Vals := t_Strlist();
    While v_Vals Is Not Null Loop
      If Length(v_Vals) <= 4000 Then
        l_Vals.Extend;
        l_Vals(l_Vals.Count) := v_Vals;
        v_Vals := Null;
      Else
        l_Vals.Extend;
        l_Vals(l_Vals.Count) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
        v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
      End If;
    End Loop;
  
    If 0 = Nvl(n_��ѯ��ʽ, 0) Then
      For Lp In 1 .. l_Vals.Count Loop
        For R In (Select m.�շ�ϸĿid, n.���� || Decode(m.���շ��õȼ�, Null, Null, '(' || m.���շ��õȼ� || ')') As ҽ������, n.����
                  From ����֧����Ŀ M, ����֧������ N
                  Where m.��Ŀ���� Is Not Null And m.����id = n.Id And m.���� = n_���� And
                        m.�շ�ϸĿid In (Select /*+cardinality(b,10)*/
                                      b.Column_Value As �շ�ϸĿid
                                     From Table(f_Num2list(l_Vals(Lp))) B)
                  Group By m.�շ�ϸĿid, n.����, m.���շ��õȼ�) Loop
        
          v_Jtmp := v_Jtmp || ',{"fee_item_id":' || r.�շ�ϸĿid;
          v_Jtmp := v_Jtmp || ',"insure_name":"' || Zljsonstr(r.����) || '"';
          v_Jtmp := v_Jtmp || ',"insure_name_ex":"' || Zljsonstr(r.ҽ������) || '"';
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
      End Loop;
    
      If c_Jtmp Is Null Then
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Jtmp, 2) || ']}}';
      Else
        c_Jtmp   := c_Jtmp || v_Jtmp;
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || c_Jtmp || ']}}';
      End If;
    
    Elsif 1 = n_��ѯ��ʽ Then
    
      For Lp In 1 .. l_Vals.Count Loop
        For R In (Select Distinct �շ�ϸĿid
                  From ����֧����Ŀ
                  Where ��Ŀ���� Is Not Null And
                        �շ�ϸĿid In (Select /*+cardinality(b,10)*/
                                    b.Column_Value As �շ�ϸĿid
                                   From Table(f_Num2list(l_Vals(Lp))) B) And ���� = n_����) Loop
          v_Jtmp := v_Jtmp || ',' || r.�շ�ϸĿid;
        
          If Length(v_Jtmp) > 32000 Then
            If c_Jtmp Is Null Then
              c_Jtmp := Substr(v_Jtmp, 2);
            Else
              c_Jtmp := c_Jtmp || v_Jtmp;
            End If;
            v_Jtmp := Null;
          End If;
        
        End Loop;
      End Loop;
      If c_Jtmp Is Null Then
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_item_ids":"' || Substr(v_Jtmp, 2) || '"}}';
      Else
        c_Jtmp   := c_Jtmp || v_Jtmp;
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_item_ids":"' || c_Jtmp || '"}}';
      End If;
    
    Elsif 2 = n_��ѯ��ʽ Then
    
      v_����s := j_Json.Get_String('insurance_types');
      v_Tmp   := Null;
      n_����  := Null;
    
      For Lp In 1 .. l_Vals.Count Loop
        For R In (Select a.����, a.�շ�ϸĿid
                  From ����֧����Ŀ A
                  Where a.��Ŀ���� Is Not Null And Nvl(a.����, 0) <> 0 And
                        a.�շ�ϸĿid In (Select /*+cardinality(b,10)*/
                                      b.Column_Value As �շ�ϸĿid
                                     From Table(f_Num2list(l_Vals(Lp))) B) And
                        a.���� In (Select /*+cardinality(x,10)*/
                                  x.Column_Value
                                 From Table(f_Num2list(v_����s)) X)
                  Group By a.����, a.�շ�ϸĿid
                  Order By a.����, a.�շ�ϸĿid) Loop
        
          If n_���� <> r.���� And n_���� Is Not Null Then
          
            v_Jtmp := v_Jtmp || ',{"insurance_type":' || n_����;
            v_Jtmp := v_Jtmp || ',"fee_item_ids":"' || Substr(v_Tmp, 2) || '"';
            v_Jtmp := v_Jtmp || '}';
          
            If Length(v_Jtmp) > 30000 Then
              If c_Jtmp Is Null Then
                c_Jtmp := Substr(v_Jtmp, 2);
              Else
                c_Jtmp := c_Jtmp || v_Jtmp;
              End If;
              v_Jtmp := Null;
            End If;
          
            v_Tmp := Null;
          End If;
          n_���� := r.����;
          v_Tmp  := v_Tmp || ',' || r.�շ�ϸĿid;
        End Loop;
      End Loop;
    
      --��ĩһ��
      If n_���� Is Not Null Then
        v_Jtmp := v_Jtmp || ',{"insurance_type":' || n_����;
        v_Jtmp := v_Jtmp || ',"fee_item_ids":"' || Substr(v_Tmp, 2) || '"';
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 30000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      End If;
      If c_Jtmp Is Null Then
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","pay_list":[' || Substr(v_Jtmp, 2) || ']}}';
      Else
        c_Jtmp   := c_Jtmp || v_Jtmp;
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","pay_list":[' || c_Jtmp || ']}}';
      End If;
    End If;
  Else
    n_�շ�ϸĿid := j_Json.Get_Number('fee_item_id');
    Select Max(n.����)
    Into v_Tmp
    From ����֧����Ŀ M, ����֧������ N
    Where m.��Ŀ���� Is Not Null And m.�շ�ϸĿid = n_�շ�ϸĿid And m.����id = n.Id And m.���� = n_����;
    Select Count(1) Into n_Tmp From ����֧����Ŀ Where �շ�ϸĿid = n_�շ�ϸĿid And ���� = n_����;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","insure_name":"' || v_Tmp || '","isexist":' || n_Tmp || '}}';
  End If;
Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Exsesvr_Getinsureiteminfo;
/
 
CREATE OR REPLACE Procedure Zl_Exsesvr_Delbill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ҽ�����ϣ�סԺҽ�����˷��ͣ�ɾ�����õ���
  --��Σ�Json_In:��ʽ
  --input
  --   operator_name         C  1 ����Ա���������ʵ�ɾ��ʱ���롿
  --   operator_code         C  1 ����Ա��š����ʵ�ɾ��ʱ���롿
  --   operator_time         C  1 ����ʱ��:yyyy-mm-dd hh:mi:ss�����ʵ�ɾ��ʱ���롿
  --   del_list  ֱ��ɾ���ĵ����б�
  --             fee_source          N 1 ������Դ:1-������ü�¼;2-סԺ���ü�¼
  --             fee_bill_type       N 1 ��¼���ʣ�1-�շѵ���2-���ʵ�
  --             fee_no              C 1 ���õ��ݺ�
  --             del_type            N   �˷ѷ�ʽ:0-����Ŵ�ɾ�����ã�1-������id��ɾ������;2-ȫ��
  --             serial_num          C   ��Ŵ�,query_type=0ʱ��Ч��
  --                                             ���ʵ���ʽ: ���1:����:ִ��״̬1,���2:����2:ִ��״̬2,...
  --                                                 ��ʽ˵����ִ��״̬:0-δִ��;1-��ȫִ��;2-����ִ��
  --                                             �շѵ���ʽ�����1,���2,���3...
  --
  --             exe_sta_nums        C   ��Ҫ��ȡ��ִ�е���Ŀ����ʽ:���1,���2,���3...
  --             fee_ids             C   ����id����query_type=1ʱ��Ч
  --                                               ���ʵ���ʽ: id1:����:ִ��״̬1,id2:����2:ִ��״̬2,...
  --                                                   ��ʽ˵����ִ��״̬:0-δִ��;1-��ȫִ��;2-����ִ��
  --                                               �շѵ���ʽ��id1,id2,id3...
  --             oper_status         N 1 ����״̬��סԺ���ʵ�ɾʱ�Ŵ��룬0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������);2-��ʾת��������
  --����: Json_Out,��ʽ����
  --output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  --˵�������ȷ��ΪClobԭ��
  --        1.�ھ���������ȡ����ҩʱɾ������ʱ��γ��ȿ��ܻᳬ��32767��
  --        2.�������Clob����һ�α�Varchar2ƽ���̶�ֻ��10ms���ҵ��ò�Ƶ��
  j_Input Pljson;
  j_Json  Pljson;

  j_List Pljson_List;
  j_Item Pljson;

  v_��� סԺ���ü�¼.����Ա���%Type;
  v_��Ա סԺ���ü�¼.����Ա����%Type;
  d_ʱ�� סԺ���ü�¼.�Ǽ�ʱ��%Type;

  n_��Դ Number(2);
  n_���� Number(2);

  v_No       סԺ���ü�¼.No%Type;
  v_������� Varchar2(32767);
  v_���ִ�� Varchar2(32767);
  v_����ID���� Varchar2(32767);
  v_����ID   varchar2(4000);
  v_��ǰ��� varchar2(4000);
  n_����״̬ Number;
  n_�˷ѷ�ʽ Number(2);

Begin
  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');
  v_���  := j_Json.Get_String('operator_code');
  v_��Ա  := j_Json.Get_String('operator_name');
  d_ʱ��  := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  j_List := j_Json.Get_Pljson_List('del_list');
  For J In 1 .. j_List.Count Loop
    j_Item     := Pljson();
    j_Item     := Pljson(j_List.Get(J));
    n_��Դ     := j_Item.Get_Number('fee_source');
    n_����     := j_Item.Get_Number('fee_bill_type');
    v_No       := j_Item.Get_String('fee_no');
    v_���ִ�� := j_Item.Get_String('exe_sta_nums');
    n_����״̬ := j_Item.Get_Number('oper_status');
    n_�˷ѷ�ʽ := j_Item.Get_Number('del_type');
    If Nvl(n_�˷ѷ�ʽ, 0) = 0 Then
      v_������� := j_Item.Get_String('serial_num');
    Elsif Nvl(n_�˷ѷ�ʽ, 0) = 2 Then
      --ȫ��,������no��
      v_������� := Null;
    Else
      --������ID��
      --������IDת��Ϊ���
      v_����id���� := j_Item.Get_String('fee_ids');
      If v_����id���� Is Null Then
        Json_Out := zlJsonOut('δ������Ҫ���ʵķ���id');
        Return;
      End If;
      If n_���� = 1 Then
        v_������� := v_����id����;
      Else
        --���ʵ�
        v_����ID���� := v_����ID���� || ',';
        While v_����ID���� Is Not Null Loop
          v_����ID := Substr(v_����ID����, 1, Instr(v_����ID����, ',', 3940) - 1);
          If n_��Դ = 1 Then
            Select /*+cardinality(b,10)*/
             f_List2str(Cast(Collect(a.��� || ':' || B.C2) As t_Strlist))
            Into v_��ǰ���
            From ������ü�¼ a, Table(f_Str2list2(v_����ID)) b
            Where a.Id = b.C1 And a.No = v_No And a.��¼���� = 2;
          Else
            Select /*+cardinality(b,10)*/
             f_List2str(Cast(Collect(a.��� || ':' || B.C2) As t_Strlist))
            Into v_��ǰ���
            From סԺ���ü�¼ a, Table(f_Str2list2(v_����ID)) b
            Where a.Id = b.C1 And a.No = v_No And a.��¼���� = 2;
          End If;

          v_������� := v_������� || ',' || v_��ǰ���;
          v_����ID���� := Substr(v_����ID����, Instr(v_����ID����, ',', 3940) + 1);
        End Loop;
        If v_������� Is Not Null Then
          v_������� := substr(v_�������, 2);
        End If;
      End If;
    End If;

    --��Ϊ�����˷ѹ����м����ִ��״̬����Ҫ����������ִ��״̬
    --��Ա����Զ�ִ����ɵ�ҽ��������ҽ�����͡���������ҽ��ʱ��Ҫ��ȡ��ִ����ɣ����˷�
    If v_���ִ�� Is Not Null Or Nvl(n_�˷ѷ�ʽ, 0) = 2 Then
      If n_��Դ = 2 Then
        Update סԺ���ü�¼
        Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
        Where NO = v_No And (Instr(',' || v_���ִ�� || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or Nvl(n_�˷ѷ�ʽ, 0) = 2) And
              Mod(��¼����, 10) = n_���� And ��¼״̬ In (0, 1, 3);
      Else
        Update ������ü�¼
        Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
        Where NO = v_No And (Instr(',' || v_���ִ�� || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or Nvl(n_�˷ѷ�ʽ, 0) = 2) And
              Mod(��¼����, 10) = n_���� And ��¼״̬ In (0, 1, 3);
      End If;
    End If;

    If n_��Դ = 1 Then
      --����
      If n_���� = 1 Then
        --���ﻮ��
        Zl_���ﻮ�ۼ�¼_Delete_s(v_No, v_�������, n_�˷ѷ�ʽ);
      Else
        --�������
        Zl_������ʼ�¼_Delete_s(v_No, v_�������, v_���, v_��Ա, d_ʱ��, 2);
      End If;
    Else
      --סԺ
      Zl_סԺ���ʼ�¼_Delete_s(v_No, v_�������, v_���, v_��Ա, 2, Nvl(n_����״̬, 0), d_ʱ��);
    End If;
  End Loop;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delbill;
/

Create Or Replace Procedure Zl_Exsesvr_Billverify
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����õ������
  --��Σ�Json_In:��ʽ
  --  input
  --    operator_time         C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --    operator_name         C 1 ����Ա����
  --    operator_code         C 1 ����Ա���
  --    item_list
  --        fee_source        N 1 ������Դ:1-����;2-סԺ
  --        fee_no            C 1 ���õ��ݺ�
  --        serial_nums       C 0 ��Ŵ���������ʾ���ŵ���
  --        pharmacy_window   C 0 ��ҩ���ڣ�������ԴΪ����ʱ���룬��ʽ���ⷿID1:��ҩ����1,�ⷿID2:��ҩ����2,....
  --        pati_id           N 0 ����id��������ԴΪסԺ�Ұ��������ʱ����(��Ҫ��Լ��ʱ�)
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  j_Item PLJson;
  j_List Pljson_List;

  v_��Ա Varchar2(300);
  v_��� Varchar2(300);
  d_ʱ�� Date;

  n_��Դ     Number(1); --1-����;2-סԺ
  v_No       ������ü�¼.No%Type;
  v_���     Varchar2(32767);
  v_��ҩ���� Varchar2(32767);
  n_����id   ������ü�¼.����id%Type;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_��Ա := j_Json.Get_String('operator_name');
  v_��� := j_Json.Get_String('operator_code');
  d_ʱ�� := To_Date(j_Json.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  j_List := j_Json.Get_Pljson_List('item_list');
  For I In 1 .. j_List.Count Loop
    j_Item     := PLJson();
    j_Item     := PLJson(j_List.Get(I));
    n_��Դ     := j_Item.Get_Number('fee_source');
    v_No       := j_Item.Get_String('fee_no');
    v_���     := j_Item.Get_String('serial_nums');
    v_��ҩ���� := j_Item.Get_String('pharmacy_window');
    n_����id   := j_Item.Get_Number('pati_id');
  
    If n_��Դ = 1 Then
      Zl_������ʼ�¼_Verify_s(v_No, v_���, v_��Ա, v_���, d_ʱ��, v_��ҩ����, 0);
    Else
      Zl_סԺ���ʼ�¼_Verify_s(v_No, v_���, v_��Ա, v_���, n_����id, d_ʱ��, 0);
    End If;
  End Loop;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Billverify;
/


  
Create Or Replace Procedure Zl_Exsesvr_Getnextid
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
  --  quantity      N  0 �������еĸ��������ֻȡһ���òβ����򶼴�0 
  -- ����:
  --  output
  --  next_id      C   1  ���У�quantity>1 ʱ�����ض����ţ��ö��ŷ���
  -------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Table  Varchar2(500);
  v_Col    Varchar2(500);
  n_Nextid Number;
  n_����   Number;
  v_Ids    Varchar2(32767);
  v_Sql    Varchar2(4000);
  --��̬�α�����
  Type Rs_Recordset Is Ref Cursor;
  c_Tmp Rs_Recordset;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Table := j_Json.Get_String('table_name');
  v_Col   := Nvl(j_Json.Get_String('col_name'), 'ID');
  n_����  := j_Json.Get_Number('quantity');

  If Nvl(n_����, 0) > 1 Then
    v_Sql := 'Select ' || v_Table || '_' || v_Col || '.Nextval as ���� From Dual Connect By Level <= :1';
    Open c_Tmp For v_Sql
      Using In n_����;
  
    v_Ids := Null;
    Loop
      Fetch c_Tmp
        Into n_Nextid;
      Exit When c_Tmp%NotFound;
      If c_Tmp%RowCount > 0 Then
        v_Ids := v_Ids || ',' || n_Nextid;
      End If;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","next_id":"' || Substr(v_Ids, 2) || '"}}';
    Return;
  End If;

  Execute Immediate 'select ' || v_Table || '_' || v_Col || '.nextval from dual'
    Into n_Nextid;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","next_id":"' || n_Nextid || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnextid;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Newbill_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ʲ���ʱ���Ƚ���������ݵĺϷ��Լ��
  --��Σ�Json_In:��ʽ
  --input
  --        pati_id             N 1  ����ID
  --        pati_pageid         N 1  ��ҳId
  --        pati_deptid         N 1  ���˿��� id
  --        pati_wardarea_id    N 1  ���˲���iD  
  --        pati_name           C 1  ��������
  --        fee_audit_status    N 1  ������˱�־:0���-δ���;1-����˻�ʼ���(��ϲ���:������˷�ʽ������);2-������,��Ͻ���Ȩ��[��ֹδ��˲��˽���]���й������
  --        si_inp_status       N 1  סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ

  --        dept_list[]��������ID�����ò���ID
  --                            plcdept_id          N 1  ��������ID
  --                            takedept_id         N 1  ���ò���ID���� ҩƷ����ҩ����id
  --����: Json_Out,��ʽ����
  --
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --        item_list[]                         ������
  --            pati_id             N   1    ����ID
  --            takedept_id             N   1   ���ò���ID
  ---------------------------------------------------------------------------
  j_Input   Pljson;
  j_Json    Pljson;
  j_Item    Pljson;
  j_List    Pljson_List := Pljson_List();
  v_Jtmp_In Varchar2(4000);
  v_Jpati   Varchar2(3000);
  n_����id  סԺ���ü�¼.����id%Type;
  n_��ҳid  סԺ���ü�¼.��ҳid%Type;
  v_����    סԺ���ü�¼.����%Type;

  n_���˿���id ���ű�.Id%Type;
  n_���˲���id ���ű�.Id%Type;
  n_��������id סԺ���ü�¼.��������id%Type;

  n_��ҩ����id ���ű�.Id%Type;

  n_������˱�־ Number(2);
  n_סԺ״̬     Number(2);
Begin

  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id       := j_Json.Get_Number('pati_id');
  n_��ҳid       := j_Json.Get_Number('pati_pageid');
  n_������˱�־ := j_Json.Get_Number('fee_audit_status');
  n_סԺ״̬     := j_Json.Get_Number('si_inp_status');
  n_���˿���id   := j_Json.Get_Number('pati_deptid');
  n_���˲���id   := j_Json.Get_Number('pati_wardarea_id');
  v_����         := j_Json.Get_String('pati_name');

  j_List := j_Json.Get_Pljson_List('dept_list');
  For I In 1 .. j_List.Count Loop
    j_Item := Pljson();
    j_Item := Pljson(j_List.Get(I));
  
    n_��������id := j_Item.Get_Number('plcdept_id');
    n_��ҩ����id := j_Item.Get_Number('takedept_id');
  
    v_Jpati := v_Jpati || ',{"pati_id":' || n_����id;
    v_Jpati := v_Jpati || ',"pati_pageid":' || n_��ҳid;
    v_Jpati := v_Jpati || ',"pati_deptid":' || Nvl(n_���˿���id, 0);
    v_Jpati := v_Jpati || ',"pati_wardarea_id":' || Nvl(n_���˲���id, 0);
    v_Jpati := v_Jpati || ',"plcdept_id":' || Nvl(n_��������id, 0);
    v_Jpati := v_Jpati || ',"takedept_id":' || Nvl(n_��ҩ����id, 0);
    v_Jpati := v_Jpati || ',"pati_name":"' || Zljsonstr(v_����) || '"';
    v_Jpati := v_Jpati || ',"fee_audit_status":' || Nvl(n_������˱�־, 0);
    v_Jpati := v_Jpati || ',"si_inp_status":' || Nvl(n_סԺ״̬, 0);
    v_Jpati := v_Jpati || '}';
  End Loop;
  v_Jtmp_In := '{"input":{"item_list":[' || Substr(v_Jpati, 2) || ']}}';

  Zl_סԺ���ʼ�¼_Insert_Check(v_Jtmp_In, Json_Out);
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Newbill_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Checkexcitemvalid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���˺�������Ϣ 
  --��Σ�Json_In:��ʽ
  --    input
  --      module  N  1  ģ���
  --      pati_id  N  1  ����id
  --      balance_mode  N  1  ����ģʽ
  --      fitem_type  C  1  �շ����:����ö� ��
  --      fitem_ids  C  1  �շ�ϸĿids:����ö���
  --      fee_nos  C    1 ���õ��ݺ�
  --
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    black_infor C 1 "��������Ϣ:�ò��˲��Ǻ��������ˣ�����NULl�����򷵻ظ�ʽ:���Ʒ�ʽ|��ʾ����Ϣ ;���Ʒ�ʽ��1-��ֹ;2-��ʾ(��ѯ��)"
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_����id   ������ü�¼.����id%Type;
  n_����ģʽ Number(10);
  v_�շ���� Varchar2(32680);
  v_Nos      Varchar2(32680);
  n_ģ���   Number(18);

  v_�շ�ϸĿids Varchar2(32680);
  v_Infor       Varchar2(32680);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_ģ���   := j_Json.Get_Number('module');
  n_����id   := j_Json.Get_Number('pati_id');
  n_����ģʽ := j_Json.Get_Number('balance_mode');

  v_�շ�ϸĿids := j_Json.Get_String('fitem_type');
  v_�շ����    := j_Json.Get_String('fitem_ids');

  v_Nos := j_Json.Get_String('fee_nos');

  v_Infor := Zl_Get_Excuteitem_Infor_s(n_ģ���, n_����id, n_����ģʽ, v_�շ����, v_Nos, v_�շ�ϸĿids);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","ctrl_infor":"' || zlJsonStr(v_Infor) || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkexcitemvalid;
/



CREATE OR REPLACE Procedure Zl_Exsesvr_Getnewblacklists
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���ID,��ȡ��Ҫ����ĺ���������
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id              N 1 ����id
  --      last_rgsappt_time    C 1 ���벻����¼�����һ��ʱ�䣺yyyy-mm-dd hh24:mi:ss
  --      operator_name        C 1 ����Ա����
  --      blackLst_regnos      C 1 ����������ĹҺŵ���,����ö��ŷ���
  --
  --����: Json_Out,��ʽ����
  --  output
  --    code                   N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    badrec_list            C   ����Ĳ�����¼�б�
  --      pati_id              N 1 ����id
  --      behavior_category    C 1 ��Ϊ���:��ԤԼ�Һ�
  --      happen_time          C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --      add_time             C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --      add_note             C 1 ����ԭ����ԤԼ����
  --      add_memo             C 1 ����˵��
  --      additional_info      C 1 ������Ϣ
  --      creator              C 1 �Ǽ���
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);
  c_Output Clob;

  n_����id     ���˹Һż�¼.����id%Type;
  v_����Ա���� Varchar2(20);
  v_���������� Varchar2(32680);
  d_��������   Date;

  n_Count        Number(18);
  n_ԤԼ����Ч�� Number(18);
  n_ԤԼ�˺�Ч�� Number(18);
  n_ԤԼ����Ч�� Number(18);
  d_���ԤԼʱ�� Date;
  v_Para         Varchar2(4000);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id       := j_Json.Get_Number('pati_id');
  d_���ԤԼʱ�� := To_Date(j_Json.Get_String('last_rgsappt_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_����Ա���� := j_Json.Get_String('operator_name');
  v_���������� := j_Json.Get_String('blackLst_regnos');
  v_���������� := ',' || Nvl(v_����������, '') || ',';

  If v_����Ա���� Is Null Then
    v_����Ա���� := zl_UserName;
  End If;

  --��ʽ:ԤԼδ�������|ԤԼ�������|ԤԼ�˺ſ���
  --ԤԼδ������ƣ�>0ԤԼ֮�󳬹���Чʱ��δ����ԤԼ��;<0��ʾԤԼ֮���ڳ����ӳٵ���Чʱ��δ����ԤԼ��
  --ԤԼ������ƣ�>0,ԤԼ֮�󳬹���Чʱ������δ������ΪˬԼ
  --ԤԼ�˺ſ���:>0,ԤԼ֮�󳬹���Чʱ��δ�����ҽ����˺���ΪˬԼ

  v_Para  := Nvl(zl_GetSysParameter('ԤԼ��Ч��������', '1111'), '0|0|0');
  n_Count := Instr(v_Para, '|');
  If n_Count = 0 Then
    n_ԤԼ����Ч�� := To_Number(Nvl(v_Para, '0'));
    v_Para         := Null;
  Else
    n_ԤԼ����Ч�� := To_Number(Substr(v_Para, 1, n_Count - 1));
    v_Para         := Substr(v_Para, n_Count + 1);
  End If;

  n_ԤԼ����Ч�� := 0;
  If v_Para Is Not Null Then
    n_Count := Instr(v_Para, '|');
    If n_Count = 0 Then
      n_ԤԼ����Ч�� := To_Number(Nvl(v_Para, '0'));
      v_Para         := Null;
    Else
      n_ԤԼ����Ч�� := To_Number(Substr(v_Para, 1, n_Count - 1));
      v_Para         := Substr(v_Para, n_Count + 1);
    End If;
  End If;

  n_ԤԼ�˺�Ч�� := 0;
  If v_Para Is Not Null Then
    n_Count := Instr(v_Para, '|');
    If n_Count = 0 Then
      n_ԤԼ�˺�Ч�� := To_Number(Nvl(v_Para, '0'));
      v_Para         := Null;
    Else
      n_ԤԼ�˺�Ч�� := To_Number(Substr(v_Para, 1, n_Count - 1));
    End If;
  End If;

  n_ԤԼ����Ч�� := -1 * Nvl(n_ԤԼ����Ч��, 0);
  n_ԤԼ����Ч�� := Nvl(n_ԤԼ����Ч��, 0);
  n_ԤԼ�˺�Ч�� := Nvl(n_ԤԼ�˺�Ч��, 0);

  If n_ԤԼ����Ч�� = 0 And n_ԤԼ����Ч�� = 0 And n_ԤԼ�˺�Ч�� = 0 Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","badrec_list":[' || '' || ']}}';

    Return;
  End If;

  If Nvl(n_����id, 0) = 0 Then
    d_�������� := Trunc(Sysdate) - 1 / 24 / 60 / 60; -- ȱʡ����ͷһ�������
    For c_ԤԼ In (Select Distinct a.No, a.����id, a.��¼����, a.��¼״̬, Nvl(a.ԤԼʱ��, a.����ʱ��) As ԤԼʱ��, c.���� As ��������, ִ����, ����ʱ��
                 From ���˹Һż�¼ A, ���ű� C
                 Where a.ִ�в���id = c.Id(+) And a.ԤԼ = 1 And
                       ((a.��¼���� = 2 And Nvl(a.��¼״̬, 0) = 1 And
                       ((a.ԤԼʱ�� + n_ԤԼ����Ч�� * (1 / 24 / 60)) <= Sysdate And n_ԤԼ����Ч�� <> 0)) Or
                       (a.��¼���� = 1 And Nvl(a.��¼״̬, 0) = 1 And
                       ((Nvl(a.ִ��ʱ��, Sysdate) - Nvl(a.ԤԼʱ��, Sysdate)) * 24 * 60 >= n_ԤԼ����Ч��) And n_ԤԼ����Ч�� <> 0) Or
                       (a.��¼״̬ = 2 And ((a.�Ǽ�ʱ�� - Nvl(a.����ʱ��, Sysdate)) * 24 * 60 >= n_ԤԼ�˺�Ч��) And n_ԤԼ�˺�Ч�� <> 0)) And
                       a.ԤԼʱ�� >= Trunc(d_��������) And a.ԤԼʱ�� <= d_�������� And Instr(v_����������, ',' || a.No || ',') = 0) Loop

      --ԤԼ����
      v_Para := '��' || To_Char(c_ԤԼ.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss');
      v_Para := v_Para || 'ԤԼ��"' || c_ԤԼ.�������� || '"����';

      If c_ԤԼ.ִ���� Is Not Null Then
        v_Para := v_Para || '��ҽ��Ϊ"' || c_ԤԼ.ִ���� || '"';
      End If;
      v_Para := v_Para || '(ԤԼ��:' || c_ԤԼ.No || Case
                  When c_ԤԼ.��¼״̬ = 2 Then
                   '�����˺�'
                  When c_ԤԼ.��¼���� = 1 Then
                   ' �������ڽ���'
                  Else
                   ''
                End || ')�ĺ�Դ��δ��ʱ���';

      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;

        v_Output := Null;
      End If;

      zlJsonPutValue(v_Output, 'pati_id', c_ԤԼ.����id, 1, 1);
      zlJsonPutValue(v_Output, 'behavior_category', 'ԤԼ�Һ�');
      zlJsonPutValue(v_Output, 'happen_time', To_Char(c_ԤԼ.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_time', To_Char(Sysdate, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_note', 'ԤԼ����');
      zlJsonPutValue(v_Output, 'add_memo', v_Para);
      zlJsonPutValue(v_Output, 'additional_info', c_ԤԼ.No);
      zlJsonPutValue(v_Output, 'creator', v_����Ա����, 0, 2);

    End Loop;
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;
  Else

    For c_ԤԼ In (Select Distinct a.No, a.����id, a.��¼����, a.��¼״̬, Nvl(a.ԤԼʱ��, a.����ʱ��) As ԤԼʱ��, c.���� As ��������, ִ����, ����ʱ��
                 From ���˹Һż�¼ A, ���ű� C
                 Where a.����id = n_����id And a.ִ�в���id = c.Id(+) And a.ԤԼ = 1 And
                       ((a.��¼���� = 2 And Nvl(a.��¼״̬, 0) = 1 And
                       ((a.ԤԼʱ�� + n_ԤԼ����Ч�� * (1 / 24 / 60)) <= Sysdate And n_ԤԼ����Ч�� <> 0)) Or
                       (a.��¼���� = 1 And Nvl(a.��¼״̬, 0) = 1 And
                       ((Nvl(a.ִ��ʱ��, Sysdate) - Nvl(a.ԤԼʱ��, Sysdate)) * 24 * 60 >= n_ԤԼ����Ч��) And n_ԤԼ����Ч�� <> 0) Or
                       (a.��¼״̬ = 2 And ((a.�Ǽ�ʱ�� - Nvl(a.����ʱ��, Sysdate)) * 24 * 60 >= n_ԤԼ�˺�Ч��) And n_ԤԼ�˺�Ч�� <> 0)) And
                       a.����ʱ�� + 0 >= Nvl(d_���ԤԼʱ��, To_Date('1990-01-01', 'YYYY-MM-DD')) And Instr(v_����������, ',' || a.No || ',') = 0

                 ) Loop

      --ԤԼ����
      v_Para := '��' || To_Char(c_ԤԼ.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss');
      v_Para := v_Para || 'ԤԼ��"' || c_ԤԼ.�������� || '"����';

      If c_ԤԼ.ִ���� Is Not Null Then
        v_Para := v_Para || '��ҽ��Ϊ"' || c_ԤԼ.ִ���� || '"';
      End If;
      v_Para := v_Para || '(ԤԼ��:' || c_ԤԼ.No || Case
                  When c_ԤԼ.��¼״̬ = 2 Then
                   ' ���������˺�'
                  When c_ԤԼ.��¼���� = 1 Then
                   ' �������ڽ���'
                  Else
                   ''
                End || ')�ĺ�Դ��δ��ʱ���';

      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;

        v_Output := Null;
      End If;
      zlJsonPutValue(v_Output, 'pati_id', Nvl(c_ԤԼ.����id, 0), 1, 1);
      zlJsonPutValue(v_Output, 'behavior_category', 'ԤԼ�Һ�');
      zlJsonPutValue(v_Output, 'happen_time', To_Char(c_ԤԼ.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_time', To_Char(Sysdate, 'yyyy-mm-dd hh24:mi:ss'));
      zlJsonPutValue(v_Output, 'add_note', 'ԤԼ����');
      zlJsonPutValue(v_Output, 'add_memo', v_Para);
      zlJsonPutValue(v_Output, 'additional_info', c_ԤԼ.No);
      zlJsonPutValue(v_Output, 'creator', v_����Ա����, 0, 2);

    End Loop;
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","badrec_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","badrec_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnewblacklists;
/

Create Or Replace Procedure Zl_Exsesvr_Updateregpatiinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�޸Ĳ��˹Һ����ݵĲ�����Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --        reg_no C 1 ���ݺ�
  --        pati_name C 1 ����
  --        pati_sex  C 1 �Ա�
  --        pati_age  C 1 ����
  --        outpatient_num  C   �����
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_No     ���˹Һż�¼.No%Type;
  v_����   ���˹Һż�¼.����%Type;
  v_�Ա�   ���˹Һż�¼.�Ա�%Type;
  v_����   ���˹Һż�¼.����%Type;
  n_����� ���˹Һż�¼.�����%Type;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_No := j_Json.Get_String('reg_no');

  v_����   := j_Json.Get_String('pati_name');
  v_�Ա�   := j_Json.Get_String('pati_sex');
  v_����   := j_Json.Get_String('pati_age');
  n_����� := To_Number(j_Json.Get_String('outpatient_num'));

  Zl_���˹ҺŻ�����Ϣ_Update(v_No, n_�����, v_����, v_�Ա�, v_����);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Updateregpatiinfo;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Getpatisurplusinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ����Ԥ�����ͷ������
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id            N   ����id
  --   pati_pageid        N   ��ҳid
  --   pati_ids           C   ����IDs,����ö��ŷ���
  --   use_type           N   0/null ���ز���Ԥ�����ͷ������  =1ʱ���ز���δ������б�

  --   ˵�������ݲ���id����ҳid��ѯ����ĳһ�ε�סԺ����������infee_surplus
  --         ���ݲ���ids����ѯ��Ӧ���˵����������Ϣ������surplus_list[]
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    infee_surplus       N 1 ��δ�����
  --    indpst_surplus      N 1 סԺԤ�����
  --    infee_surplusnew    N 1 ����δ�����
  --    surplus_list[]      C 1 ����б�
  --      pati_Id           N   ����ID
  --      outdpst_surplus   N 1 ����Ԥ�����
  --      indpst_surplus    N 1 סԺԤ�����
  --      outfee_surplus    N 1 ����������
  --      infee_surplus     N 1 סԺ�������
  --    unfinish_list[]     C 1 ����δ������б�
  --      pati_Id           N   ����ID
  --      page_Id           N   ��ҳID
  --      infee_surplusnew  N 1 ����δ�����
  ---------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  v_Output Varchar2(32767);
  c_Output Clob;

  n_����id ����δ�����.����id%Type;
  n_��ҳid ����δ�����.��ҳid%Type;
  v_Ids    Varchar2(3000);
  n_Type   Number;

  n_���η������ Number(16, 5);
  n_סԺ���     Number(16, 5);
  n_��סԺ����   Number(16, 5);

Begin

  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');
  n_Type   := Nvl(j_Json.Get_Number('use_type'), 0);

  v_Ids := j_Json.Get_String('pati_ids');

  If Nvl(v_Ids, '-') = '-' And Nvl(n_����id, 0) = 0 Then
    Json_Out := Zljsonout('δ�����κβ�ѯ����������!');
    Return;
  End If;

  If n_Type = 0 Then

    If Nvl(n_����id, 0) <> 0 Then
      If Nvl(n_��ҳid, 0) = 0 Then
        Json_Out := Zljsonout('δ������ҳid������!');
        Return;
      End If;
      Select Nvl(Sum(���), 0) Into n_���η������ From ����δ����� Where ����id = n_����id And ��ҳid = n_��ҳid;

      Select Sum(Decode(a.����, 2, a.Ԥ�����, 0)) As סԺ���, Sum(Decode(a.����, 2, a.�������, 0)) As סԺ����
      Into n_סԺ���, n_��סԺ����
      From ������� A
      Where a.����id = n_����id And a.���� = 1;

      Zljsonputvalue(v_Output, 'code', 1, 1, 1);
      Zljsonputvalue(v_Output, 'message', '�ɹ�');
      Zljsonputvalue(v_Output, 'infee_surplus', n_��סԺ����, 1);
      Zljsonputvalue(v_Output, 'indpst_surplus', n_סԺ���, 1);
      Zljsonputvalue(v_Output, 'infee_surplusnew', n_���η������, 1, 2);
      Json_Out := '{"output":' || v_Output || '}';

      Return;
    Else
      For r_������� In (Select a.����id, Sum(Decode(a.����, 1, a.Ԥ�����, 0)) As �������, Sum(Decode(a.����, 2, a.Ԥ�����, 0)) As סԺ���,
                            Sum(Decode(a.����, 1, a.�������, 0)) As �������, Sum(Decode(a.����, 2, a.�������, 0)) As סԺ����
                     From ������� A, Table(f_Num2list(v_Ids)) B
                     Where a.����id = b.Column_Value And a.���� = 1
                     Group By ����id) Loop
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;

        Zljsonputvalue(v_Output, 'pati_id', r_�������.����id, 1, 1);
        Zljsonputvalue(v_Output, 'outdpst_surplus', r_�������.�������, 1);
        Zljsonputvalue(v_Output, 'indpst_surplus', r_�������.סԺ���, 1);
        Zljsonputvalue(v_Output, 'outfee_surplus', r_�������.�������, 1);
        Zljsonputvalue(v_Output, 'infee_surplus', r_�������.סԺ����, 1, 2);
      End Loop;

      If Not c_Output Is Null And Not v_Output Is Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        v_Output := '';
      End If;

      If Not c_Output Is Null Then
        Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","surplus_list":[') || c_Output || To_Clob(']}}');
      Else
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","surplus_list":[' || v_Output || ']}}';
      End If;

      Return;
    End If;
  Elsif n_Type = 1 Then
    For r_δ����� In (Select a.����id, a.��ҳid, Sum(Nvl(a.���, 0)) As δ�����
                   From ����δ����� A, Table(f_Num2list(v_Ids)) B
                   Where a.����id = b.Column_Value And a.��ҳid Is Not Null
                   Group By a.����id, a.��ҳid
                   Order By a.����id, a.��ҳid) Loop
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
        v_Output := Null;
      End If;

      Zljsonputvalue(v_Output, 'pati_id', r_δ�����.����id, 1, 1);
      Zljsonputvalue(v_Output, 'page_id', r_δ�����.��ҳid, 1);
      Zljsonputvalue(v_Output, 'infee_surplusnew', r_δ�����.δ�����, 1, 2);
    End Loop;

    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;

    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","unfinish_list":[') || c_Output || To_Clob(']}}');
    Else
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","unfinish_list":[' || v_Output || ']}}';
    End If;

    Return;
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatisurplusinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Getconsumercardtype
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���ѿ����
  --��Σ�Json_In:��ʽ
  --    input
  --      enabled                N    �Ƿ�����:1-������;0-����
  --����: Json_Out,��ʽ����
  --  output
  --    code                      N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                   C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    type_list[]               C  1  ֧�ֵĿ�����б�
  --          cardtype_id       N  1  id
  --          cardtype_num      N  1  ���
  --          cardtype_name     C  1  ����
  --          cardtype_stname   C  1  ����
  --          prefix_text         C  1  ǰ׺�ı�
  --          cardno_len          N  1  ���ų���
  --          default             N  1  ȱʡ��־
  --          fixed               N  1  �Ƿ�̶�:1-��ϵͳ�̶�;0-����ϵͳ�̶�
  --          strict              N  1  �Ƿ��ϸ����:1-���ϸ����;0-�����ϸ����
  --          self_make           N  1  �Ƿ�����:1-�ǵ�;0-����
  --          allow_return_cash   N  1  �Ƿ�����:1-����;0-������
  --          must_all_return     N   1   �Ƿ�ȫ��:1-����ȫ��;0-��������
  --          specpati            N   1   �ض�����
  --          component           C   1   ����
  --          memo                C   1   ��ע
  --          blnc_mode           C   1   ���㷽ʽ
  --          blnc_nature         N   1   ��������
  --          pwdtxt           N   1   �Ƿ�����
  --          enabled             N   1   �Ƿ�����:1-������;0-δ����
  --          pwd_len             N   1   ���볤��
  --          pwd_len_limit       N   1   ���볤������:0-��������;1-�̶����볤��;-n��ʾ�����������ö��λ��������,�����ܳ������볤��
  --          pwd_rule            N   1   �������:��-���ֺ��ַ����;1-��Ϊ�������
  --          readcard_nature     C   1   ��������,ҽ�ƿ�������ʽ����һλΪ:�Ƿ�ˢ��;�ڶ�λΪ�Ƿ�ɨ��;����λ�Ƿ�Ӵ�ʽ����;����λ�Ƿ�ǽӴ�ʽ����������ˢ����'1000'
  --          keyboard_mode       N   1   ���̿��Ʒ�ʽ:��0-��ֹʹ�������;1-ʹ����������� ,2-ʹ���ַ������
  --          def_delcash         N   1   �Ƿ�ȱʡ����:��������ʱ,Ĭ���Ƿ�����
  ---------------------------------------------------------------------------
  v_Output Varchar2(32767);
  c_Output Clob;

  j_Input PLJson;
  j_Json  PLJson;

  n_�Ƿ����� Number(2);

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�Ƿ����� := j_Json.Get_Number('enabled');
  For c_���ѿ���� In (Select a.��� As ID, a.���, a.����, Nvl(a.ϵͳ, 0) As �Ƿ�̶�, a.���㷽ʽ, a.����, Nvl(a.����, 0) As �Ƿ�����,
                         Nvl(a.���ƿ�, 0) As �Ƿ�����, a.ǰ׺�ı�, a.���ų���, a.�Ƿ�����, a.�Ƿ�����, a.�Ƿ�ȫ��, a.���볤��, a.���볤������, a.�������, a.��������,
                         a.���̿��Ʒ�ʽ, a.�������, a.�Ƿ��ϸ����, a.�Ƿ��ض�����, a.�Ƿ�������, a.�Ƿ�������, a.�Ƿ���������˿�, a.Ӧ�ó���, 0 As ȱʡ��־,
                         Nvl(b.����, 0) As ��������, 0 As �Ƿ�ȱʡ����, '' As ��ע
                  From ���ѿ����Ŀ¼ A, ���㷽ʽ B
                  Where a.���㷽ʽ = b.����(+) And Decode(Nvl(n_�Ƿ�����, 0), 0, 0, Nvl(a.����, 0)) = Nvl(n_�Ƿ�����, 0)
                  
                  ) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'cardtype_id', c_���ѿ����.Id, 1, 1);
    zlJsonPutValue(v_Output, 'cardtype_num', c_���ѿ����.���);
    zlJsonPutValue(v_Output, 'cardtype_name', c_���ѿ����.����);
    zlJsonPutValue(v_Output, 'cardtype_stname', Substr(c_���ѿ����.����, 1, 1));
  
    zlJsonPutValue(v_Output, 'prefix_text', Nvl(c_���ѿ����.ǰ׺�ı�, ''));
    zlJsonPutValue(v_Output, 'cardno_len', Nvl(c_���ѿ����.���ų���, 0), 1);
    zlJsonPutValue(v_Output, 'default', Nvl(c_���ѿ����.ȱʡ��־, 0), 1);
  
    zlJsonPutValue(v_Output, 'fixed', Nvl(c_���ѿ����.�Ƿ�̶�, 0), 1);
    zlJsonPutValue(v_Output, 'strict', Nvl(c_���ѿ����.�Ƿ��ϸ����, 0), 1);
    zlJsonPutValue(v_Output, 'self_make', Nvl(c_���ѿ����.�Ƿ�����, 0), 1);
    zlJsonPutValue(v_Output, 'allow_return_cash', Nvl(c_���ѿ����.�Ƿ�����, 0), 1);
    zlJsonPutValue(v_Output, 'must_all_return', Nvl(c_���ѿ����.�Ƿ�ȫ��, 0), 1);
    zlJsonPutValue(v_Output, 'specpati', Nvl(c_���ѿ����.�Ƿ��ض�����, 0), 1);
  
    zlJsonPutValue(v_Output, 'component', Nvl(c_���ѿ����.����, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(c_���ѿ����.��ע, ''));
  
    zlJsonPutValue(v_Output, 'blnc_mode', Nvl(c_���ѿ����.���㷽ʽ, ''));
    zlJsonPutValue(v_Output, 'blnc_nature', Nvl(c_���ѿ����.��������, 0), 1);
  
    zlJsonPutValue(v_Output, 'pwdtxt', Nvl(c_���ѿ����.�Ƿ�����, 0), 1);
    zlJsonPutValue(v_Output, 'enabled', Nvl(c_���ѿ����.�Ƿ�����, 0), 1);
  
    zlJsonPutValue(v_Output, 'pwd_len', Nvl(c_���ѿ����.���볤��, 0), 1);
    zlJsonPutValue(v_Output, 'pwd_len_limit', Nvl(c_���ѿ����.���볤������, 0), 1);
    zlJsonPutValue(v_Output, 'pwd_rule', Nvl(c_���ѿ����.�������, 0), 1);
  
    zlJsonPutValue(v_Output, 'readcard_nature', Nvl(c_���ѿ����.��������, '1000'));
    zlJsonPutValue(v_Output, 'keyboard_mode', Nvl(c_���ѿ����.���̿��Ʒ�ʽ, 0), 1);
    zlJsonPutValue(v_Output, 'def_return_cash', Nvl(c_���ѿ����.�Ƿ�ȱʡ����, 0), 1, 2);
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","type_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","type_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getconsumercardtype;
/


Create Or Replace Procedure Zl_Exsesvr_Getpatisurety
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���˵�������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id            N 1 ����Id
  --     pati_pageid        N 0 ��ҳID
  --     pati_ids           C 0 ������ҳ�ؼ���Ϣƴ��������ID:��ҳID,....
  --     surety_prop        N 0 �Ƿ��ȡ�������ʣ�0-����ȡ��1-Ҫ��ȡ��Ŀǰ��֧�ֵ������ˣ�
  --     query_type         N 0 1-��ȡ���ڲ��˵�����¼�Ĳ���ID����ҳID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    guarantee_money     N       �������
  --    entsurety           C       ������
  --    surety_prop         N       �������ʣ�0-���ڵ�����1-��ʱ�ᱣ���������һ�ε�����¼������Ϊ׼
  --    item_list[]
  --       pati_id            N 1 ����id
  --       pati_pageid        N 1 ��ҳid
  --       guarantee_money    N 1 �������
  --       entsurety          C 1 ������
  --       surety_prop        N 1 ��������
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_��ҳid   ���˵�����¼.��ҳid%Type;
  n_����id   ���˵�����¼.����id%Type;
  v_������   ���˵�����¼.������%Type;
  n_������   ���˵�����¼.������%Type;
  n_�������� Number;
  n_��ѯ���� Number(3);
  l_����     t_StrList := t_StrList();
  v_����ids  Varchar2(32767);

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_��ѯ���� := j_Json.Get_Number('query_type');

  If j_Json.Exist('pati_ids') Then
    --û�иýڵ�ʱ,ִ�������������
    v_����ids := j_Json.Get_String('pati_ids');
  End If;

  If Nvl(n_����id, 0) = 0 And v_����ids Is Null Then
    Json_Out := zlJsonOut('δ���벡��id,���飡', 0);
    Return;
  End If;

  --��ȡ��ǰ��Ч�������Ч������¼��
  n_������ := 0;
  v_������ := Null;
  If n_����id <> 0 Then
    For r_�ᱣ��Ϣ In (Select ������, ������, ��������
                   From ���˵�����¼
                   Where ����id = n_����id And (��ҳid = n_��ҳid Or Nvl(n_��ҳid, 0) = 0) And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And
                         ɾ����־ = 1
                   Order By �Ǽ�ʱ�� Desc) Loop
      If n_�������� Is Null Then
        n_�������� := Nvl(r_�ᱣ��Ϣ.��������, 0);
      End If;
      n_������ := n_������ + r_�ᱣ��Ϣ.������;
      v_������ := v_������ || ',' || r_�ᱣ��Ϣ.������;
    End Loop;
    v_������ := Substr(v_������, 2, 100);
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","guarantee_money":' || zlJsonStr(Nvl(n_������, 0), 1) ||
                ',"entsurety":"' || zlJsonStr(v_������) || '","surety_prop":' || Nvl(n_��������, 0) || '}}';
    Return;
  
  End If;

  While v_����ids Is Not Null Loop
    If Length(v_����ids) <= 4000 Then
      l_����.Extend;
      l_����(l_����.Count) := v_����ids;
      v_����ids := Null;
    Else
      l_����.Extend;
      l_����(l_����.Count) := Substr(v_����ids, 1, Instr(v_����ids, ',', 3940) - 1);
      v_����ids := Substr(v_����ids, Instr(v_����ids, ',', 3940) + 1);
    End If;
  End Loop;

  v_Output := Null;
  For I In 1 .. l_����.Count Loop
    v_����ids := l_����(I);
    If Nvl(n_��ѯ����, 0) = 0 Then
    
      For R In (Select a.����id, a.��ҳid, ������, a.��������, Sum(a.������) As ������
                From ���˵�����¼ A,
                     (Select /*+cardinality(f,10)*/
                        To_Number(f.C1) As ����id, To_Number(f.C2) As ��ҳid
                       From Table(f_Str2List2(v_����ids)) F) N
                Where a.����id = n.����id And Nvl(a.��ҳid, 0) = n.��ҳid And (a.����ʱ�� Is Null Or a.����ʱ�� > Sysdate) And
                      a.ɾ����־ = 1
                Group By a.����id, a.��ҳid, ������, a.��������) Loop
      
        If Length(Nvl(v_Output, ' ')) > 32000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
        
          v_Output := Null;
        End If;
        zlJsonPutValue(v_Output, 'pati_id', r.����id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_pageid', r.��ҳid, 1, 0);
        zlJsonPutValue(v_Output, 'guarantee_money', r.������, 1, 0);
        zlJsonPutValue(v_Output, 'entsurety', r.������, 0, 0);
        zlJsonPutValue(v_Output, 'surety_prop', Nvl(r.��������, 0), 1, 2);
      End Loop;
    
    Else
    
      For R In (Select a.����id, a.��ҳid
                From ���˵�����¼ A,
                     (Select /*+cardinality(f,10)*/
                        To_Number(f.C1) As ����id, To_Number(f.C2) As ��ҳid
                       From Table(f_Str2List2(v_����ids)) F) N
                Where a.����id = n.����id And Nvl(a.��ҳid, 0) = n.��ҳid
                Group By a.����id, a.��ҳid) Loop
      
        If Length(Nvl(v_Output, ' ')) > 32700 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        zlJsonPutValue(v_Output, 'pati_id', r.����id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_pageid', r.��ҳid, 1, 2);
      End Loop;
    End If;
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatisurety;
/


Create Or Replace Procedure Zl_Exsesvr_Getpatisuretylist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���˵�������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id            N 1 ����Id
  --     pati_pageid        N 0 ��ҳID
  --     expidate           N 1 1-��ѯ��Ч�ĵ�����Ϣ;0-���е�����Ϣ
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]         ����
  --      type               C 1 ���
  --      guarantor          C 1 ������
  --      garnt_amount       C 1 ������
  --      garnt_prop         N 1 ��������
  --      garnt_reason       C 1 ����ԭ��
  --      create_time        C 1 �Ǽ�ʱ��
  --      due_time           C 1 ����ʱ��
  --      is_del             C 1 ɾ����־
  --      operator_name      C 1 ����Ա����
  --      operator_code      C 1 ����Ա���
  --      del_operator_name  C 1 ɾ������Ա����
  --      del_operator_code  C 1 ɾ������Ա���
  --      del_time           C 1 ɾ��ʱ��
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  n_��ҳid ���˵�����¼.��ҳid%Type;
  n_����id ���˵�����¼.����id%Type;
  n_��Ч�� Number(3);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');
  n_��Ч�� := j_Json.Get_Number('expidate');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('δ���벡��id,���飡');
    Return;
  End If;
 
  If Nvl(n_��Ч��, 0) = 0 Then
    v_Output := Null;
    For R In (Select Decode(��ҳid, Null, '����', '��' || ��ҳid || '��סԺ') ���, ������,
                     Decode(������, 999999999, '����', To_Char(������, '999999990.00')) As ������, ��������, ����ԭ��,
                     To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') �Ǽ�ʱ��, To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��,
                     Decode(ɾ����־, 1, '', -1, 'ɾ��', '') As ɾ����־, ����Ա����, ����Ա���, ɾ������Ա����, ɾ������Ա���, ɾ��ʱ��
              From ���˵�����¼
              Where ����id = n_����id And (��ҳid = n_��ҳid Or Nvl(n_��ҳid, 0) = 0)
              Order By �Ǽ�ʱ�� Desc) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'type', r.���, 0, 1);
      zlJsonPutValue(v_Output, 'guarantor', r.������);
      zlJsonPutValue(v_Output, 'garnt_amount', r.������);
      zlJsonPutValue(v_Output, 'garnt_prop', r.��������, 1);
      zlJsonPutValue(v_Output, 'garnt_reason', r.����ԭ��);
      zlJsonPutValue(v_Output, 'create_time', r.�Ǽ�ʱ��);
      zlJsonPutValue(v_Output, 'due_time', r.����ʱ��);
      zlJsonPutValue(v_Output, 'is_del', r.ɾ����־);
      zlJsonPutValue(v_Output, 'operator_name', r.����Ա����);
      zlJsonPutValue(v_Output, 'operator_code', r.����Ա���);
      zlJsonPutValue(v_Output, 'del_operator_name', r.ɾ������Ա����);
      zlJsonPutValue(v_Output, 'del_operator_code', r.ɾ������Ա���);
      zlJsonPutValue(v_Output, 'del_time', r.ɾ��ʱ��, 0, 2);
    End Loop;
  Else
    For R In (Select ������, Decode(������, 999999999, '����', To_Char(������, '999999990.00')) As ������, Nvl(��������, 0) As ��������, ����ԭ��,
                     To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') �Ǽ�ʱ��
              From ���˵�����¼
              Where ����id = n_����id And ��ҳid = n_��ҳid And (����ʱ�� Is Null Or ����ʱ�� > Sysdate) And ɾ����־ = 1) Loop
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
      zlJsonPutValue(v_Output, 'guarantor', r.������, 0, 1);
      zlJsonPutValue(v_Output, 'garnt_amount', r.������);
      zlJsonPutValue(v_Output, 'garnt_prop', r.��������, 1);
      zlJsonPutValue(v_Output, 'garnt_reason', r.����ԭ��);
      zlJsonPutValue(v_Output, 'create_time', r.�Ǽ�ʱ��);
      zlJsonPutValue(v_Output, 'due_time', r.����ʱ��, 0, 2);
    End Loop;
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatisuretylist;
/


Create Or Replace Procedure Zl_Exsesvr_Patisuretyexpire
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����²��˵�������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id            N 1 ����Id
  --     pati_pageid        N 1 ��ҳID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  n_��ҳid Number(5);
  n_����id Number(18);
  j_Input  PLJson;
  j_Json   PLJson;

Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('δ���벡��id,���飡');
    Return;
  End If;
  --����������¼��Ժ���Զ�ʧЧ 
  Update ���˵�����¼
  Set ����ʱ�� = Sysdate
  Where ����id = n_����id And ��ҳid = n_��ҳid And Nvl(����ʱ��, Sysdate + 1) > Sysdate And ɾ����־ = 1 And Nvl(��������, 0) = 0;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Patisuretyexpire;
/

Create Or Replace Procedure Zl_Exsesvr_Getwarnline
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���ʱ����ߣ����˲��������ſ�Ƿ�Ѳ���
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_scheme  C 1 ���ò���
  --     wardarea_id  N 1 ����id
  --     query_type   N 1 ��ѯ��ʽ
  --                     0-������ ����id / ���ò��� ���ң�����һ��ֵ
  --                     1-������id ���ң������б�
  --                     2-��ȡ���б���������
  --                     3-���ݲ���id�����ò��˲��ң����ر�������������ֵ��������־
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --      alarm_value       N 1 ����ֵ,-1����ʾδ��������
  --      item_list[]
  --        pati_scheme     C 1 ���ò���
  --        alarm_way       N 1 ��������
  --        alarm_value     N 1 ����ֵ
  --        alarm_one       C 1 ������־1
  --        alarm_two       C 1 ������־2
  --        alarm_three     C 1 ������־3
  --        wardarea_id     N 1 ����id
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_��ѯ��ʽ Number;
  v_���ò��� Varchar2(200);
  n_����id   Number(18);
  n_����ֵ   Number;
  v_Temp     Varchar2(32767);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ��ʽ := j_Json.Get_Number('query_type');
  v_���ò��� := j_Json.Get_String('pati_scheme');
  n_����id   := j_Json.Get_Number('wardarea_id');

  If Nvl(n_��ѯ��ʽ, 0) = 0 Then
    n_����ֵ := -1;
    For R In (Select ����ֵ
              From ���ʱ�����
              Where �������� = 1 And Nvl(����id, 0) = Nvl(n_����id, 0) And ����ֵ Is Not Null And ���ò��� = v_���ò���) Loop
      n_����ֵ := r.����ֵ;
      Exit;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","alarm_value":' || zlJsonStr(n_����ֵ, 1) || '}}';
  Elsif n_��ѯ��ʽ = 3 Then
    For r_������ In (Select ���ò���, Nvl(��������, 1) As ��������, ����ֵ, ������־1, ������־2, ������־3
                  From ���ʱ�����
                  Where Nvl(����id, 0) = Nvl(n_����id, 0) And ���ò��� = v_���ò���) Loop
      --ֻȡһ������
      v_Temp := v_Temp || '{"pati_scheme":"' || r_������.���ò��� || '"';
      v_Temp := v_Temp || ',"alarm_way":' || Nvl(r_������.��������, 0);
      v_Temp := v_Temp || ',"alarm_value":' || zlJsonStr(r_������.����ֵ, 1);
      v_Temp := v_Temp || ',"alarm_one":"' || r_������.������־1 || '"';
      v_Temp := v_Temp || ',"alarm_two":"' || r_������.������־2 || '"';
      v_Temp := v_Temp || ',"alarm_three":"' || r_������.������־3 || '"';
      v_Temp := v_Temp || '}';
      Exit;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Temp || ']}}';
  Elsif n_��ѯ��ʽ = 1 Then
    For r_������ In (Select ���ò���, Nvl(��������, 1) As ��������, ����ֵ, ������־1, ������־2, ������־3
                  From ���ʱ�����
                  Where ����id = Nvl(n_����id, 0)) Loop
      v_Temp := v_Temp || ',{"pati_scheme":"' || zlJsonStr(r_������.���ò���) || '"';
      v_Temp := v_Temp || ',"alarm_way":' || Nvl(r_������.��������, 0);
      v_Temp := v_Temp || ',"alarm_value":' || zlJsonStr(r_������.����ֵ, 1);
      v_Temp := v_Temp || ',"alarm_one":"' || r_������.������־1 || '"';
      v_Temp := v_Temp || ',"alarm_two":"' || r_������.������־2 || '"';
      v_Temp := v_Temp || ',"alarm_three":"' || r_������.������־3 || '"';
      v_Temp := v_Temp || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Temp, 2) || ']}}';
  Elsif n_��ѯ��ʽ = 2 Then
    For r_������ In (Select Nvl(����id, 0) As ����id, ���ò���, Nvl(��������, 1) As ��������, ����ֵ, ������־1, ������־2, ������־3
                  From ���ʱ�����) Loop
      v_Temp := v_Temp || ',{"wardarea_id":' || Nvl(r_������.����id || '', 'null');
      v_Temp := v_Temp || ',"pati_scheme":"' || zlJsonStr(r_������.���ò���) || '"';
      v_Temp := v_Temp || ',"alarm_way":' || Nvl(r_������.��������, 0);
      v_Temp := v_Temp || ',"alarm_value":' || zlJsonStr(r_������.����ֵ, 1);
      v_Temp := v_Temp || ',"alarm_one":"' || r_������.������־1 || '"';
      v_Temp := v_Temp || ',"alarm_two":"' || r_������.������־2 || '"';
      v_Temp := v_Temp || ',"alarm_three":"' || r_������.������־3 || '"';
      v_Temp := v_Temp || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || Substr(v_Temp, 2) || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getwarnline;
/

Create Or Replace Procedure Zl_Exsesvr_Newbill
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ﲡ��סԺ���˷���ҽ�����ɷ��õ���
  --��Σ�Json_In:��ʽ
  --input
  --  pati_list[] �����б���������ʱ�����޸ýڵ�
  --    billtype                                            N 1 ����,1-�շѵ���2-���ʵ�
  --    pati_source                                         N 1 ��Դ��1-���2-סԺ
  --    pati_id                                             N 1 ����id
  --    pati_pageid                                         N 1 ��ҳid
  --    baby_num                                            N 1 Ӥ����
  --    sgin_no                                             N 1 ��ʶ�ţ�����ţ�סԺ��
  --    bed_num                                             C 1 ����
  --    pati_name                                           C 1 ����
  --    pati_sex                                            C 1 �Ա�
  --    pati_age                                            C 1 ����
  --    fee_category                                        C 1 �ѱ�
  --    overtime_sign                                       N 1 �Ӱ��־
  --    pati_deptid                                         N 1 ���˿���id
  --    pati_wardarea_id                                    N 1 ���˲���id
  --    operator_name                                       C 1 ����Ա����
  --    operator_code                                       C 1 ����Ա���
  --    outpati_tag                                         N 1 �����־
  --    rgst_id                                             N 1 ����id
  --    emg_sign                                            N 1 �Ƿ���
  --    item_list[]  ��ϸ�б�
  --        fee_id                                        N 1 ����id
  --        fee_no                                        C 1 No
  --        serial_num                                    N 1 ���
  --        charge_tag                                    N 1 ����
  --        placer                                        C 1 ������
  --        plcdept_id                                    N 1 ��������id
  --        sub_serial_num                                N 1 ��������
  --        fitem_id                                      N 1 �շ�ϸĿid
  --        item_type                                     C 1 �շ����
  --        unit                                          C 1 ���㵥λ
  --        pharmacy_window                               C 1 ��ҩ����
  --        packages_num                                  N 1 ����
  --        send_num                                      N 1 ����
  --        ext_mark                                      N 1 ���ӱ�־
  --        exe_deptid                                    N 1 ִ�в���id
  --        price_ftrnum                                  N 1 �۸񸸺�
  --        income_item_id                                N 1 ������Ŀid
  --        receipt_name                                  C 1 �վݷ�Ŀ
  --        price                                         N 1 ��׼����
  --        fee_amrcvb                                    N 1 Ӧ�ս��
  --        fee_ampaib                                    N 1 ʵ�ս��
  --        happen_time                                   C 1 ����ʱ��
  --        create_time                                   C 1 �Ǽ�ʱ��
  --        memo                                          C 1 ����ժҪ
  --        order_id                                      N 1 ҽ�����
  --        baby_num                                      N 1 Ӥ����
  --        exe_properties                                N 1 ִ������
  --        decoction_method                              C 1 �巨
  --        morphology                                    C 1 ��ҩ��̬
  --        bakstuff_batch                                N 1 ����
  --        insurance                                     N 1 ������Ŀ��
  --        insure_id                                     N 1 ���մ���id
  --        insure_code                                   C 1 ���ձ���
  --        fee_type                                      C 1 ��������
  --        si_manp_money                                 N 1 ͳ����
  --        synchro                                       N 1 ����ͬ����־
  --        effective_time                                N 1 ��Ч
  --        receipt_issecret                              N 1 ����
  --        takedept_id                                   N 1 ��ҩ����id
  --        group_id                                      N 0 ҽ��С��id
  --        auto_finish                                   N 0 �Զ���ɣ���������Զ�����
  --����: Json_Out,��ʽ����
  --output
  --  code                                                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --  message                                             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;
  j_List  Pljson_List;

  n_����       סԺ���ü�¼.���ӱ�־%Type; --������Դ,1-������ü�¼,2-סԺ���ü�¼
  n_��Դ       סԺ���ü�¼.���ӱ�־%Type; --��������,1-�շ�,2-����
  n_����id     סԺ���ü�¼.����id%Type;
  n_��ҳid     סԺ���ü�¼.��ҳid%Type;
  n_Ӥ����     סԺ���ü�¼.Ӥ����%Type;
  n_��ʶ��     סԺ���ü�¼.��ʶ��%Type;
  v_����       סԺ���ü�¼.����%Type;
  v_����       סԺ���ü�¼.����%Type;
  v_�Ա�       סԺ���ü�¼.�Ա�%Type;
  v_����       סԺ���ü�¼.����%Type;
  v_�ѱ�       סԺ���ü�¼.�ѱ�%Type;
  n_�Ӱ��־   סԺ���ü�¼.�Ӱ��־%Type;
  n_���˿���id סԺ���ü�¼.���˿���id%Type;
  n_���˲���id סԺ���ü�¼.���˲���id%Type;
  v_����Ա���� סԺ���ü�¼.����Ա����%Type;
  v_����Ա��� סԺ���ü�¼.����Ա���%Type;
  n_�����־   סԺ���ü�¼.�����־%Type;
  n_����id     סԺ���ü�¼.����id%Type;
  n_�Ƿ���   סԺ���ü�¼.�Ƿ���%Type;

  n_Id         סԺ���ü�¼.Id%Type;
  v_No         סԺ���ü�¼.No%Type;
  n_���       סԺ���ü�¼.���%Type;
  n_����       סԺ���ü�¼.���ӱ�־%Type;
  v_������     סԺ���ü�¼.������%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  n_��������   סԺ���ü�¼.��������%Type;
  n_�շ�ϸĿid סԺ���ü�¼.�շ�ϸĿid%Type;
  v_�շ����   סԺ���ü�¼.�շ����%Type;
  v_���㵥λ   סԺ���ü�¼.���㵥λ%Type;
  v_��ҩ����   סԺ���ü�¼.��ҩ����%Type;
  n_����       סԺ���ü�¼.����%Type;
  n_����       סԺ���ü�¼.����%Type;
  n_���ӱ�־   סԺ���ü�¼.���ӱ�־%Type;
  n_ִ�в���id סԺ���ü�¼.ִ�в���id%Type;
  n_�۸񸸺�   סԺ���ü�¼.�۸񸸺�%Type;
  n_������Ŀid סԺ���ü�¼.������Ŀid%Type;
  v_�վݷ�Ŀ   סԺ���ü�¼.�վݷ�Ŀ%Type;
  n_��׼����   סԺ���ü�¼.��׼����%Type;
  n_Ӧ�ս��   סԺ���ü�¼.Ӧ�ս��%Type;
  n_ʵ�ս��   סԺ���ü�¼.ʵ�ս��%Type;
  d_����ʱ��   סԺ���ü�¼.����ʱ��%Type;
  d_�Ǽ�ʱ��   סԺ���ü�¼.�Ǽ�ʱ��%Type;
  v_ժҪ       סԺ���ü�¼.ժҪ%Type;
  n_ҽ�����   סԺ���ü�¼.ҽ�����%Type;
  n_ִ������   סԺ���ü�¼.���ӱ�־%Type;
  v_�巨       סԺ���ü�¼.����%Type;
  v_��ҩ��̬   סԺ���ü�¼.����%Type;
  n_����       סԺ���ü�¼.����%Type;
  n_������Ŀ�� סԺ���ü�¼.������Ŀ��%Type;
  n_���մ���id סԺ���ü�¼.���մ���id%Type;
  v_���ձ���   סԺ���ü�¼.���ձ���%Type;
  v_��������   סԺ���ü�¼.��������%Type;
  n_ͳ����   סԺ���ü�¼.ͳ����%Type;
  n_ͬ����־   ���˷����쳣��¼.ͬ����־%Type;
  n_ҽ����Ч   סԺ���ü�¼.ҽ����Ч%Type;
  n_�Ƿ���   סԺ���ü�¼.�Ƿ���%Type;
  n_��ҩ����id סԺ���ü�¼.��ҩ����id%Type;
  n_ҽ��С��id סԺ���ü�¼.ҽ��С��id%Type;
  n_�Զ������ Number;
  n_IӤ����    סԺ���ü�¼.Ӥ����%Type;

  v_ִ�����ids Varchar2(32767);

  n_Billcount Number;
  j_Patilist  Pljson_List;
Begin
  --������� 
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --�ж��Ƿ���� pati_list �ڵ�
  n_Billcount := 1;
  If j_Json.Exist('pati_list') Then
    j_Patilist  := j_Json.Get_Pljson_List('pati_list');
    n_Billcount := j_Patilist.Count;
  End If;

  For I In 1 .. n_Billcount Loop
    If j_Patilist Is Not Null Then
      j_Json := PLJson();
      j_Json := PLJson(j_Patilist.Get(I));
    End If;
  
    n_����       := j_Json.Get_Number('billtype');
    n_��Դ       := j_Json.Get_Number('pati_source');
    n_����id     := j_Json.Get_Number('pati_id');
    n_��ҳid     := j_Json.Get_Number('pati_pageid');
    n_Ӥ����     := j_Json.Get_Number('baby_num');
    n_��ʶ��     := j_Json.Get_Number('sgin_no');
    v_����       := j_Json.Get_String('bed_num');
    v_����       := j_Json.Get_String('pati_name');
    v_�Ա�       := j_Json.Get_String('pati_sex');
    v_����       := j_Json.Get_String('pati_age');
    v_�ѱ�       := j_Json.Get_String('fee_category');
    n_�Ӱ��־   := j_Json.Get_Number('overtime_sign');
    n_���˿���id := j_Json.Get_Number('pati_deptid');
    n_���˲���id := j_Json.Get_Number('pati_wardarea_id');
    v_����Ա���� := j_Json.Get_String('operator_name');
    v_����Ա��� := j_Json.Get_String('operator_code');
    n_�����־   := j_Json.Get_Number('outpati_tag');
    n_����id     := j_Json.Get_Number('rgst_id');
    n_�Ƿ���   := j_Json.Get_Number('emg_sign');
  
    j_List        := j_Json.Get_Pljson_List('item_list');
    v_ִ�����ids := Null;
    For I In 1 .. j_List.Count Loop
      j_Json       := PLJson();
      j_Json       := PLJson(j_List.Get(I));
      n_Id         := j_Json.Get_Number('fee_id');
      v_No         := j_Json.Get_String('fee_no');
      n_���       := j_Json.Get_Number('serial_num');
      n_����       := j_Json.Get_Number('charge_tag');
      v_������     := j_Json.Get_String('placer');
      n_��������id := j_Json.Get_Number('plcdept_id');
      n_��������   := j_Json.Get_Number('sub_serial_num');
      n_�շ�ϸĿid := j_Json.Get_Number('fitem_id');
      v_�շ����   := j_Json.Get_String('item_type');
      v_���㵥λ   := j_Json.Get_String('unit');
      v_��ҩ����   := j_Json.Get_String('pharmacy_window');
      n_����       := j_Json.Get_Number('packages_num');
      n_����       := j_Json.Get_Number('send_num');
      n_���ӱ�־   := j_Json.Get_Number('ext_mark');
      n_ִ�в���id := j_Json.Get_Number('exe_deptid');
      n_�۸񸸺�   := j_Json.Get_Number('price_ftrnum');
      n_������Ŀid := j_Json.Get_Number('income_item_id');
      v_�վݷ�Ŀ   := j_Json.Get_String('receipt_name');
      n_��׼����   := j_Json.Get_Number('price');
      n_Ӧ�ս��   := j_Json.Get_Number('fee_amrcvb');
      n_ʵ�ս��   := j_Json.Get_Number('fee_ampaib');
      d_����ʱ��   := To_Date(j_Json.Get_String('happen_time'), 'yyyy-mm-dd hh24:mi:ss');
      d_�Ǽ�ʱ��   := To_Date(j_Json.Get_String('create_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_ժҪ       := j_Json.Get_String('memo');
      n_ҽ�����   := j_Json.Get_Number('order_id');
      n_IӤ����    := j_Json.Get_Number('baby_num');
      n_ִ������   := j_Json.Get_Number('exe_properties');
      v_�巨       := j_Json.Get_String('decoction_method');
      v_��ҩ��̬   := j_Json.Get_String('morphology');
      n_����       := j_Json.Get_Number('bakstuff_batch');
      n_������Ŀ�� := j_Json.Get_Number('insurance');
      n_���մ���id := j_Json.Get_Number('insure_id');
      v_���ձ���   := j_Json.Get_String('insure_code');
      v_��������   := j_Json.Get_String('fee_type');
      n_ͳ����   := j_Json.Get_Number('si_manp_money');
      n_ͬ����־   := j_Json.Get_Number('synchro');
      n_ҽ����Ч   := j_Json.Get_Number('effective_time');
      n_�Ƿ���   := j_Json.Get_Number('receipt_issecret');
      n_��ҩ����id := j_Json.Get_Number('takedept_id');
      n_ҽ��С��id := j_Json.Get_Number('group_id');
      n_�Զ������ := j_Json.Get_Number('auto_finish');
    
      If n_IӤ���� Is Not Null Then
        n_Ӥ���� := n_IӤ����;
      End If;
    
      If n_��Դ = 1 And n_���� = 2 Then
        Zl_������ʼ�¼_Insert_s(v_No, n_���, n_����id, n_��ʶ��, v_����, v_�Ա�, v_����, v_�ѱ�, n_�Ӱ��־, n_Ӥ����, n_���˿���id, n_��������id, v_������,
                           n_��������, n_�շ�ϸĿid, v_�շ����, v_���㵥λ, n_����, n_����, n_���ӱ�־, n_ִ�в���id, n_�۸񸸺�, n_������Ŀid, v_�վݷ�Ŀ,
                           n_��׼����, n_Ӧ�ս��, n_ʵ�ս��, d_����ʱ��, d_�Ǽ�ʱ��, n_����, v_��ҩ����, v_����Ա���, v_����Ա����, n_Id, Null, v_ժҪ,
                           n_ҽ�����, n_�����־, v_��ҩ��̬, v_�巨, n_��ҳid, n_���˲���id, n_����, n_ͬ����־, n_����id, n_�Ƿ���, n_ҽ����Ч, n_�Ƿ���);
      Elsif n_��Դ = 1 And n_���� = 1 Then
        Zl_���ﻮ�ۼ�¼_Insert_s(v_No, n_���, n_����id, n_��ҳid, n_��ʶ��, Null, v_����, v_�Ա�, v_����, v_�ѱ�, n_�Ӱ��־, n_���˿���id, n_��������id,
                           v_������, n_��������, n_�շ�ϸĿid, v_�շ����, v_���㵥λ, v_��ҩ����, n_����, n_����, n_���ӱ�־, n_ִ�в���id, n_�۸񸸺�,
                           n_������Ŀid, v_�վݷ�Ŀ, n_��׼����, n_Ӧ�ս��, n_ʵ�ս��, d_����ʱ��, d_�Ǽ�ʱ��, v_����Ա����, n_Id, v_ժҪ, n_ҽ�����, v_�巨,
                           1, v_���ձ���, v_��������, n_������Ŀ��, n_���մ���id, v_��ҩ��̬, Null, n_���˲���id, n_����, n_ͬ����־, n_����id, n_�Ƿ���,
                           n_ҽ����Ч, n_�Ƿ���);
      Elsif n_��Դ = 2 And n_���� = 2 Then
        Zl_סԺ���ʼ�¼_Insert_s(v_No, n_���, n_����id, n_��ҳid, n_��ʶ��, v_����, v_�Ա�, v_����, v_����, v_�ѱ�, n_���˲���id, n_���˿���id, n_�Ӱ��־,
                           n_Ӥ����, n_��������id, v_������, n_��������, n_�շ�ϸĿid, v_�շ����, v_���㵥λ, n_������Ŀ��, n_���մ���id, v_���ձ���, n_����,
                           n_����, n_���ӱ�־, n_ִ�в���id, n_�۸񸸺�, n_������Ŀid, v_�վݷ�Ŀ, n_��׼����, n_Ӧ�ս��, n_ʵ�ս��, n_ͳ����, d_����ʱ��,
                           d_�Ǽ�ʱ��, n_����, v_����Ա���, v_����Ա����, n_Id, Null, Null, v_ժҪ, n_�Ƿ���, n_ҽ�����, Null, v_��������, Null,
                           v_��ҩ��̬, n_ҽ��С��id, v_�巨, n_ִ������, n_����, n_��ҩ����id, n_ͬ����־, n_ҽ����Ч, n_�Ƿ���);
      End If;
    
      If n_�Զ������ = 1 Then
        v_ִ�����ids := v_ִ�����ids || ',' || n_Id;
      End If;
    
    End Loop;
  
    If v_ִ�����ids Is Not Null Then
      v_ִ�����ids := Substr(v_ִ�����ids, 2);
      Select Sysdate Into d_�Ǽ�ʱ�� From Dual;
      If n_��Դ = 1 Then
        Update ������ü�¼
        Set ִ��״̬ = 1, ִ���� = v_����Ա����, ִ��ʱ�� = d_�Ǽ�ʱ��
        Where ID In (Select /*+cardinality(j,10)*/
                      j.Column_Value
                     From Table(f_Num2List(v_ִ�����ids)) J);
      Else
        Update סԺ���ü�¼
        Set ִ��״̬ = 1, ִ���� = v_����Ա����, ִ��ʱ�� = d_�Ǽ�ʱ��
        Where ID In (Select /*+cardinality(j,10)*/
                      j.Column_Value
                     From Table(f_Num2List(v_ִ�����ids)) J);
      End If;
    End If;
  End Loop;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Newbill;
/

Create Or Replace Procedure Zl_Exsesvr_Actualmoney
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -----------------------------------------------------------------------------
  --����:����Ӧ�ս��ѱ���д��ۼ���,����ʵ�ս�� 
  --���json
  --input     ����ʵ�ս��
  --          fee_category        C 1 �ѱ�
  --          fee_item_id         N 1 �շ�ϸĿid
  --          income_item_id      N 1 ������Ŀid
  --          fee_amrcvb          N 1 Ӧ�ս��
  --          quantity            N 1 ����
  --          price_cost          N 1 �ɱ���
  --          order_id            N 1 ҽ��id
  --          item_list[]�б�
  --                  fee_category        C 1 �ѱ�
  --                   fee_item_id         N 1 �շ�ϸĿid
  --                   income_item_id      N 1 ������Ŀid
  --                   fee_amrcvb          N 1 Ӧ�ս��
  --                   quantity            N 1 ����
  --                   price_cost          N 1 �ɱ���
  --                   order_id            N 1 ҽ��id
  --����json
  --output      
  --        code                  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --        message               C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --        fee_category          C 1 �ѱ�
  --        fee_ampaib            N 1 ʵ�ս�
  --        fee_ampaibs           C 1 ʵ�ս������ŷָ������Ϊ���б�ʱ����
  ---------------------------------------------------------------------------
  j_Input    Pljson;
  j_Json     Pljson;
  j_List     Pljson_List := Pljson_List();
  n_ʵ�ս�� ������ü�¼.ʵ�ս��%Type;
  v_�ѱ�     ������ü�¼.�ѱ�%Type;
  v_Tmp      Varchar2(1000);
  v_Jtmp     Varchar2(32767);
Begin
  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');
  If j_Json.Exist('item_list') Then
    j_List := j_Json.Get_Pljson_List('item_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Json := Pljson();
        j_Json := Pljson(j_List.Get(I));
        Select Zl_Actualmoney_s(j_Json.Get_String('fee_category'), j_Json.Get_Number('fee_item_id'),
                                 j_Json.Get_Number('income_item_id'), j_Json.Get_Number('fee_amrcvb'),
                                 j_Json.Get_Number('quantity'), j_Json.Get_Number('price_cost'),
                                 j_Json.Get_Number('order_id'))
        Into v_Tmp
        From Dual;
      
        Select To_Number(Substr(v_Tmp, Instr(v_Tmp, ':') + 1)) As ��� Into n_ʵ�ս�� From Dual;
      
        v_Jtmp := v_Jtmp || ',' || n_ʵ�ս��;
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ampaibs":"' || Substr(v_Jtmp, 2) || '"}}';
  Else
    Select Zl_Actualmoney_s(j_Json.Get_String('fee_category'), j_Json.Get_Number('fee_item_id'),
                             j_Json.Get_Number('income_item_id'), j_Json.Get_Number('fee_amrcvb'),
                             j_Json.Get_Number('quantity'), j_Json.Get_Number('price_cost'), j_Json.Get_Number('order_id'))
    Into v_Tmp
    From Dual;
    Select Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1) As �ѱ�, To_Number(Substr(v_Tmp, Instr(v_Tmp, ':') + 1)) As ���
    Into v_�ѱ�, n_ʵ�ս��
    From Dual;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_category":"' || v_�ѱ� || '","fee_ampaib":' || Nvl(n_ʵ�ս��, 0) || '}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Actualmoney;
/

Create Or Replace Procedure Zl_Exsesvr_Checkonetimefee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --���ܣ���鲡���Ƿ�ִ��һ���Է���
  --��Σ�Json_In,��ʽ����
  --  input
  --    pati_id            N 1 ����ID
  --    pati_pageid        N 1 ��ҳID
  --    pati_wardarea_id   N 1 ��Ժ����ID
  --    in_date            C 1 ��Ժ����  ��ʽ��YYYY-MM-DD HH:MM:SS
  --����: Json_Out,��ʽ����
  --  output
  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    is_exe             N 1 ִ�б��:0-��ִ��;1-ִ��
  --------------------------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Count            Number;
  n_Pati_Id          Number(18);
  n_Pati_Pageid      Number(18);
  n_Pati_Wardarea_Id Number(18);

  d_In_Date Date;
Begin

  --�������
  j_Input            := PLJson(Json_In);
  j_Json             := j_Input.Get_Pljson('input');
  n_Pati_Id          := j_Json.Get_Number('pati_id');
  n_Pati_Pageid      := j_Json.Get_Number('pati_pageid');
  n_Pati_Wardarea_Id := j_Json.Get_Number('pati_wardarea_id');

  d_In_Date := To_Date(j_Json.Get_String('in_date'), 'yyyy-mm-dd hh24:mi:ss');

  Select Count(*)
  Into n_Count
  From �Զ��Ƽ���Ŀ B
  Where b.����id = n_Pati_Wardarea_Id And b.�����־ = 8 And Nvl(b.��������, To_Date('3000-01-01', 'YYYY-MM-DD')) <= d_In_Date;
  If n_Count = 0 Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","is_exe":0}}';
    Return;
  End If;

  --���ò��˱���סԺ�Ƿ��Ѿ������
  Select Count(*)
  Into n_Count
  From סԺ���ü�¼
  Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 8;
  If n_Count > 0 Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","is_exe":0}}';
    Return;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","is_exe":1}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkonetimefee;
/
Create Or Replace Procedure Zl_Exsesvr_Calconetimefee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --���ܣ���סԺ���˼���һ���Է��á�
  --��Σ�Json_In,��ʽ����
  --  input
  --    pati_id            N 1 ����ID
  --    pati_pageid        N 1 ��ҳID
  --    inpatient_num      C 1 סԺ��
  --    pati_wardarea_id   N 1 ��Ժ����ID
  --    pati_deptid        N 1 ��Ժ����ID
  --    medical_team_id    N 1 ҽ��С��ID
  --    pati_name          C 1 ��������
  --    pati_Sex           C 1 �����Ա�
  --    pati_age           C 1 ��������
  --    pati_bed           C 1 ��Ժ����
  --    fee_category       C 1 �ѱ�
  --    in_date            C 1 ��Ժ����
  --    func_id            N 1 ����ID 0-����,1-ɾ��
  --    mdlpay_mode_name   C 1 ҽ�Ƹ��ʽ����
  --    operator_name      C 1 ����Ա����
  --    operator_code      C 1 ����Ա���
  --    operator_deptid    N 1 ����Ա����ID    
  --����: Json_Out,��ʽ����
  --  output
  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --------------------------------------------------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  n_Func_Id       Number(1);
  n_Pati_Id       Number(18);
  n_Pati_Pageid   Number(6);
  v_Operator_Name Varchar2(100);
  v_Operator_Code Varchar2(100);
Begin
  --�������
  j_Input   := Pljson(Json_In);
  j_Json    := j_Input.Get_Pljson('input');
  n_Func_Id := j_Json.Get_Number('func_id');

  If Nvl(n_Func_Id, 0) = 0 Then
    Zl_סԺһ�η���_Insert_s(Json_In, Json_Out);
  Else
    n_Pati_Id       := j_Json.Get_Number('pati_id');
    n_Pati_Pageid   := j_Json.Get_Number('pati_pageid');
    v_Operator_Code := j_Json.Get_String('operator_code');
    v_Operator_Name := j_Json.Get_String('operator_name');
    Zl_סԺһ�η���_Delete(n_Pati_Id, n_Pati_Pageid, v_Operator_Code, v_Operator_Name);
  End If;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Calconetimefee;
/

Create Or Replace Procedure Zl_Exsesvr_Checkmrbkfeeisdel
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��鲡�����Ƿ��Ѿ��˷�
  --��Σ�Json_In:��ʽ
  --   input
  --      fee_no            C 1 ���ݺ�
  --      pati_id           N 1 ����id
  --      fee_properties    N 1 ��¼����
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     isdel              N 1 �Ƿ����˷�:1-�Ѿ��˷�;0-δ�˷�
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Err_Msg Varchar2(500);

  v_No       סԺ���ü�¼.No%Type;
  n_����id   סԺ���ü�¼.����id%Type;
  n_��¼���� סԺ���ü�¼.��¼����%Type;
  n_Exist    Number(1);

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  v_No       := j_Json.Get_String('fee_no');
  n_��¼���� := j_Json.Get_Number('fee_properties');

  If Nvl(v_No, '-') = '-' Then
    v_Err_Msg := 'δ���뵥�ݺţ�����!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := 'δ���벡��id������!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If Nvl(n_��¼����, 0) = 0 Then
    v_Err_Msg := 'δ�����¼���ʣ�����!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If n_��¼���� = 4 Then
    Select Max(1)
    Into n_Exist
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 1 And ���ӱ�־ = 1 And ����id = n_����id And NO = v_No;
  Else
    Select Max(1)
    Into n_Exist
    From סԺ���ü�¼
    Where ��¼���� = 5 And ��¼״̬ = 1 And ���ӱ�־ = 8 And ����id = n_����id And NO = v_No;
  End If;
  If Nvl(n_Exist, 0) = 0 Then
    n_Exist := 1;
  Else
    n_Exist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isdel":' || n_Exist || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkmrbkfeeisdel;
/


Create Or Replace Procedure Zl_Exsesvr_Getdepositblncsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡԤ�����ݽ�����쳣״̬
  --��Σ�Json_In:��ʽ
  --   input
  --      deposit_no        C 1 Ԥ�����ݺ�
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     blance_sign        N 0-����״̬,1-�����쳣״̬,2-�˿��쳣״̬
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  v_Output  Varchar2(32767);

  j_Input PLJson;
  j_Json  PLJson;

  v_No    ����Ԥ����¼.No%Type;
  n_Count Number(1);
  n_State Number(1);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_No := j_Json.Get_String('deposit_no');

  If Nvl(v_No, '-') = '-' Then
    v_Err_Msg := 'δ���뵥�ݺţ�����!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;
  Select Count(1), Max(��¼״̬)
  Into n_Count, n_State
  From ����Ԥ����¼
  Where ��¼���� = 1 And NO = v_No And Nvl(У�Ա�־, 0) <> 0;
  If n_Count = 0 Then
    n_State := 0;
  Else
    If n_State = 0 Then
      n_State := 1;
    End If;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'isdelfee', n_State, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositblncsign;
/



Create Or Replace Procedure Zl_Exsesvr_Getcardfeeblncsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���˵��쳣��������
  --��Σ�Json_In:��ʽ
  --   input
  --      pati_id           N 1 ����id
  --      operator_name     C 1 ����Ա����
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     item_list[]
  --       balance_id       N 1 ����id
  --       operator_name    C 1 ����Ա����
  --       err_type         N 1 �쳣����:1-�����շ��쳣;2-�˷��쳣
  --       is_mrbk          N 1 �Ƿ�����:1-�ǲ�����;0-���ǲ�����
  --       create_time      C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --       rec_sign         N 1 ��¼״̬
  --       balance_num      N 1 �������
  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Input   PLJson;
  j_Json    PLJson;

  v_����Ա סԺ���ü�¼.����Ա����%Type;
  n_����id סԺ���ü�¼.����id%Type;
  v_Output Varchar2(32767);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  v_����Ա := j_Json.Get_String('operator_name');

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := 'δ���벡��id������!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If Nvl(v_����Ա, '-') = '-' Then
    v_Err_Msg := 'δ�������Ա����������!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  For r_�쳣���� In (Select Distinct a.No, a.����id, a.����Ա����, a.�쳣����, a.�Ǽ�ʱ��, b.�������, a.��¼״̬, Decode(a.���ӱ�־, 8, 1, 0) As ������
                 From (Select a.No, a.����id, a.����Ա����, 1 As �쳣����, a.�Ǽ�ʱ��, a.��¼״̬, a.���ӱ�־
                        From סԺ���ü�¼ A
                        Where Nvl(����״̬, 0) = 1 And ��¼���� = 5 And ����id = n_����id And ��¼״̬ = 1 And ����id Is Not Null
                        Union All
                        Select a.No, a.����id, a.����Ա����, 2 As �쳣����, a.�Ǽ�ʱ��, a.��¼״̬, 0 As ���ӱ�־
                        From סԺ���ü�¼ A
                        Where Nvl(����״̬, 0) = 1 And ��¼���� = 5 And ����id = n_����id And ��¼״̬ = 2 And Not Exists
                         (Select 1 From ����Ԥ����¼ Where ����id = a.����id And Nvl(У�Ա�־, 0) = 0)) A, ����Ԥ����¼ B
                 Where a.����id = b.����id
                 Order By Decode(a.����Ա����, v_����Ա, 0, 1), a.��¼״̬) Loop
  
    zlJsonPutValue(v_Output, 'balance_id', r_�쳣����.����id, 1, 1);
    zlJsonPutValue(v_Output, 'operator_name', r_�쳣����.����Ա����);
    zlJsonPutValue(v_Output, 'err_type', r_�쳣����.�쳣����, 1);
    zlJsonPutValue(v_Output, 'is_mrbk', r_�쳣����.������, 1);
    zlJsonPutValue(v_Output, 'create_time', r_�쳣����.�Ǽ�ʱ��);
    zlJsonPutValue(v_Output, 'rec_sign', r_�쳣����.��¼״̬, 1);
    zlJsonPutValue(v_Output, 'balance_num', r_�쳣����.�������, 1, 2);
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","surplus_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getcardfeeblncsign;
/


Create Or Replace Procedure Zl_Exsesvr_Getdepositdetail
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡָ�������µ�Ԥ����֧��ϸ����
  --��Σ�Json_In:��ʽ
  --   input
  --      pati_id           N 1 ����id
  --      begin_time        C 1 ��ʼʱ��:yyyy-mm-dd hh24:mi:ss
  --      end_time          C 1 ��ֹʱ��:yyyy-mm-dd hh24:mi:ss
  --      type              N 1 ���ͱ�־:0-����,1-����;2-סԺ
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     item_list[]
  --       business_type    C 1 ҵ������:�ڳ�����ֵ���շ��á����ʵ�
  --       happen_time      C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --       earlystage       N 1 �ڳ����
  --       recharge         N 1 ���ڳ�ֵ
  --       consume          N 1 ��������
  ---------------------------------------------------------------------------
  v_Err_Msg  Varchar2(500);
  j_Input    PLJson;
  j_Json     PLJson;
  n_����id   סԺ���ü�¼.����id%Type;
  d_��ʼʱ�� Date;
  d_��ֹʱ�� Date;
  n_����     Number(1);
  n_��ת��   Number(1);
  d_�ϴ����� Zldatamove.�ϴ�����%Type;

  v_Output Varchar2(32767);
  c_Output Clob;

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  d_��ʼʱ�� := To_Date(j_Json.Get_String('begin_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_��ֹʱ�� := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_����     := Nvl(j_Json.Get_Number('type'), 0);

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := 'δ���벡��id������!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  Select Max(�ϴ�����) Into d_�ϴ����� From zlDataMove Where ��� = 1 And ϵͳ = 100 And �ϴ����� Is Not Null;
  n_��ת�� := 0;
  If d_�ϴ����� Is Not Null Then
    If d_�ϴ����� > d_��ʼʱ�� Then
      n_��ת�� := 1;
    End If;
  End If;

  If n_��ת�� = 1 Then
    --������ʷ����
    For r_Ԥ����ϸ In (Select /*+ RULE */
                    ���, �տ�ʱ��, ҵ������, Sum(�ڳ����) As �ڳ����, Sum(���ڳ�ֵ) As ���ڳ�ֵ, Sum(��������) As ��������
                   From (With Ԥ�� As (Select ����id, �տ�ʱ��, 0 As ����, ����id, Nvl(���, 0) As ���, 0 As ��Ԥ��
                                     From ����Ԥ����¼ A
                                     Where a.�տ�ʱ�� >= d_��ʼʱ�� And a.��¼���� = 1 And
                                           ((a.��¼״̬ In (1, 3) And Nvl(a.У�Ա�־, 0) = 0) Or a.��¼״̬ = 2) And a.����id = n_����id And
                                           Nvl(a.Ԥ�����, 2) In (1, 2)
                                     Union All
                                     Select a.����id, b.�շ�ʱ�� As �տ�ʱ��, 2 As ����, b.Id As ����id, 0 As ���, Nvl(��Ԥ��, 0) As ��Ԥ��
                                     From ����Ԥ����¼ A, ���˽��ʼ�¼ B
                                     Where b.�շ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                           a.����id = b.Id And a.����id = n_����id And (Nvl(a.Ԥ�����, 2) = n_���� Or n_���� = 0)
                                     Union All
                                     Select ����id, �շ�ʱ��, 1 As ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ��
                                     From (Select a.����id, Min(b.�Ǽ�ʱ��) As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���,
                                                   Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                                            From ����Ԥ����¼ A, ������ü�¼ B
                                            Where b.�Ǽ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                                  Nvl(b.���ʷ���, 0) = 0 And a.����id = b.����id And b.����id = n_����id And
                                                  b.��¼���� In (1, 4) And Nvl(a.��Ԥ��, 0) <> 0 And
                                                  (Nvl(a.Ԥ�����, 2) = n_���� Or n_���� = 0)
                                            Group By a.����id, a.No, b.����id)
                                     Group By ����id, �շ�ʱ��, ����id
                                     Union All
                                     Select ����id, �շ�ʱ��, 1 As ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ��
                                     From (Select a.����id, Min(b.�Ǽ�ʱ��) As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���,
                                                   Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                                            From ����Ԥ����¼ A, סԺ���ü�¼ B
                                            Where b.�Ǽ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                                  a.����id = b.����id And b.����id = n_����id And b.��¼���� = 5 And Nvl(b.���ʷ���, 0) = 0 And
                                                  Nvl(a.��Ԥ��, 0) <> 0
                                            Group By a.����id, a.No, b.����id)
                                     Group By ����id, �շ�ʱ��, ����id
                                     Union All
                                     --��ʷԤ����¼
                                     Select ����id, �տ�ʱ��, 0 As ����, ����id, Nvl(���, 0) As ���, 0 As ��Ԥ��
                                     From H����Ԥ����¼ A
                                     Where a.�տ�ʱ�� >= d_��ʼʱ�� And a.��¼���� = 1 And
                                           ((a.��¼״̬ In (1, 3) And Nvl(a.У�Ա�־, 0) = 0) Or a.��¼״̬ = 2) And a.����id = n_����id And
                                           Nvl(a.Ԥ�����, 2) In (1, 2)
                                     Union All
                                     Select a.����id, b.�շ�ʱ�� As �տ�ʱ��, 2 As ����, b.Id As ����id, 0 As ���, Nvl(��Ԥ��, 0) As ��Ԥ��
                                     From H����Ԥ����¼ A, ���˽��ʼ�¼ B
                                     Where b.�շ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                           a.����id = b.Id And a.����id = n_����id And (Nvl(a.Ԥ�����, 2) = n_���� Or n_���� = 0)
                                     Union All
                                     Select ����id, �շ�ʱ��, 1 As ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ��
                                     From (Select a.����id, Min(b.�Ǽ�ʱ��) As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���,
                                                   Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                                            From H����Ԥ����¼ A, H������ü�¼ B
                                            Where b.�Ǽ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                                  Nvl(b.���ʷ���, 0) = 0 And a.����id = b.����id And b.����id = n_����id And
                                                  b.��¼���� In (1, 4) And Nvl(a.��Ԥ��, 0) <> 0 And
                                                  (Nvl(a.Ԥ�����, 2) = n_���� Or n_���� = 0)
                                            Group By a.����id, a.No, b.����id)
                                     Group By ����id, �շ�ʱ��, ����id
                                     Union All
                                     Select ����id, �շ�ʱ��, 1 As ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ��
                                     From (Select a.����id, Min(b.�Ǽ�ʱ��) As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���,
                                                   Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                                            From H����Ԥ����¼ A, HסԺ���ü�¼ B
                                            Where b.�Ǽ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                                  a.����id = b.����id And b.����id = n_����id And b.��¼���� = 5 And Nvl(b.���ʷ���, 0) = 0 And
                                                  Nvl(a.��Ԥ��, 0) <> 0
                                            Group By a.����id, a.No, b.����id)
                                     Group By ����id, �շ�ʱ��, ����id
                                     
                                     )
                          Select 0 As ���, '' As �տ�ʱ��, '�ڳ�' As ҵ������, Sum(Nvl(Ԥ�����, 0)) As �ڳ����, 0 As ���ڳ�ֵ, 0 As ��������
                          From ������� A
                          Where ����id = n_����id And ���� = 1 And (Nvl(a.����, 2) = n_���� Or n_���� = 0)
                          Union All
                          Select 0 As ���, '' As �տ�ʱ��, '�ڳ�' As ҵ������, -1 * Sum(Nvl(���, 0)) + Sum(Nvl(��Ԥ��, 0)) As �ڳ����, 0,
                                 0 As ��������
                          From Ԥ��
                          Where �տ�ʱ�� >= d_��ʼʱ��
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd')
                          Union All
                          Select 1 As ���, To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, '��ֵ' As ҵ������, 0 As �ڳ����,
                                 Sum(Nvl(���, 0)) As ��ֵ, 0 As ��������
                          From Ԥ��
                          Where �տ�ʱ�� Between d_��ʼʱ�� And d_��ֹʱ�� And Nvl(���, 0) > 0
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd')
                          Union All
                          Select 1 As ���, To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, Decode(����, 1, '�շ�', 2, '����', '����') As ҵ������,
                                 0 As �ڳ����, 0 As ��ֵ, Sum(Nvl(��Ԥ��, 0)) As ����
                          From Ԥ��
                          Where �տ�ʱ�� Between d_��ʼʱ�� And d_��ֹʱ�� Having Sum(Nvl(��Ԥ��, 0)) <> 0
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd'), Decode(����, 1, '�շ�', 2, '����', '����')
                          Union All
                          Select 2 As ���, To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, '��ֵ' As ҵ������, 0 As �ڳ����,
                                 Sum(Nvl(���, 0)) As ��ֵ, 0 As ��������
                          From Ԥ��
                          Where �տ�ʱ�� Between d_��ʼʱ�� And d_��ֹʱ�� And Nvl(���, 0) < 0
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd'))
                          Group By ���, �տ�ʱ��, ҵ������
                          Order By ���, �տ�ʱ��
                   ) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'business_type', r_Ԥ����ϸ.ҵ������, 0, 1);
      zlJsonPutValue(v_Output, 'happen_time', r_Ԥ����ϸ.�տ�ʱ��);
      zlJsonPutValue(v_Output, 'earlystage', r_Ԥ����ϸ.�ڳ����, 1);
      zlJsonPutValue(v_Output, 'recharge', r_Ԥ����ϸ.���ڳ�ֵ, 1);
      zlJsonPutValue(v_Output, 'consume', r_Ԥ����ϸ.��������, 1, 2);
    
    End Loop;
  
  Else
    For r_Ԥ����ϸ In (Select /*+ RULE */
                    ���, �տ�ʱ��, ҵ������, Sum(�ڳ����) As �ڳ����, Sum(���ڳ�ֵ) As ���ڳ�ֵ, Sum(��������) As ��������
                   From (With Ԥ�� As (Select ����id, �տ�ʱ��, 0 As ����, ����id, Nvl(���, 0) As ���, 0 As ��Ԥ��
                                     From ����Ԥ����¼ A
                                     Where a.�տ�ʱ�� >= d_��ʼʱ�� And a.��¼���� = 1 And
                                           ((a.��¼״̬ In (1, 3) And Nvl(a.У�Ա�־, 0) = 0) Or a.��¼״̬ = 2) And a.����id = n_����id And
                                           Nvl(a.Ԥ�����, 2) In (1, 2)
                                     Union All
                                     Select a.����id, b.�շ�ʱ�� As �տ�ʱ��, 2 As ����, b.Id As ����id, 0 As ���, Nvl(��Ԥ��, 0) As ��Ԥ��
                                     From ����Ԥ����¼ A, ���˽��ʼ�¼ B
                                     Where b.�շ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                           a.����id = b.Id And a.����id = n_����id And (Nvl(a.Ԥ�����, 2) = n_���� Or n_���� = 0)
                                     Union All
                                     Select ����id, �շ�ʱ��, 1 As ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ��
                                     From (Select a.����id, Min(b.�Ǽ�ʱ��) As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���,
                                                   Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                                            From ����Ԥ����¼ A, ������ü�¼ B
                                            Where b.�Ǽ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                                  Nvl(b.���ʷ���, 0) = 0 And a.����id = b.����id And b.����id = n_����id And
                                                  b.��¼���� In (1, 4) And Nvl(a.��Ԥ��, 0) <> 0 And
                                                  (Nvl(a.Ԥ�����, 2) = n_���� Or n_���� = 0)
                                            Group By a.����id, a.No, b.����id)
                                     Group By ����id, �շ�ʱ��, ����id
                                     Union All
                                     Select ����id, �շ�ʱ��, 1 As ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ��
                                     From (Select a.����id, Min(b.�Ǽ�ʱ��) As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���,
                                                   Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                                            From ����Ԥ����¼ A, סԺ���ü�¼ B
                                            Where b.�Ǽ�ʱ�� >= d_��ʼʱ�� And Mod(a.��¼����, 10) = 1 And Nvl(a.У�Ա�־, 0) = 0 And
                                                  a.����id = b.����id And b.����id = n_����id And b.��¼���� = 5 And Nvl(b.���ʷ���, 0) = 0 And
                                                  Nvl(a.��Ԥ��, 0) <> 0
                                            Group By a.����id, a.No, b.����id)
                                     Group By ����id, �շ�ʱ��, ����id)
                          Select 0 As ���, '' As �տ�ʱ��, '�ڳ�' As ҵ������, Sum(Nvl(Ԥ�����, 0)) As �ڳ����, 0 As ���ڳ�ֵ, 0 As ��������
                          From ������� A
                          Where ����id = n_����id And ���� = 1 And (Nvl(a.����, 2) = n_���� Or n_���� = 0)
                          Union All
                          Select 0 As ���, '' As �տ�ʱ��, '�ڳ�' As ҵ������, -1 * Sum(Nvl(���, 0)) + Sum(Nvl(��Ԥ��, 0)) As �ڳ����, 0,
                                 0 As ��������
                          From Ԥ��
                          Where �տ�ʱ�� >= d_��ʼʱ��
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd')
                          Union All
                          Select 1 As ���, To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, '��ֵ' As ҵ������, 0 As �ڳ����,
                                 Sum(Nvl(���, 0)) As ��ֵ, 0 As ��������
                          From Ԥ��
                          Where �տ�ʱ�� Between d_��ʼʱ�� And d_��ֹʱ�� And Nvl(���, 0) > 0
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd')
                          Union All
                          Select 1 As ���, To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, Decode(����, 1, '�շ�', 2, '����', '����') As ҵ������,
                                 0 As �ڳ����, 0 As ��ֵ, Sum(Nvl(��Ԥ��, 0)) As ����
                          From Ԥ��
                          Where �տ�ʱ�� Between d_��ʼʱ�� And d_��ֹʱ�� Having Sum(Nvl(��Ԥ��, 0)) <> 0
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd'), Decode(����, 1, '�շ�', 2, '����', '����')
                          Union All
                          Select 2 As ���, To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, '��ֵ' As ҵ������, 0 As �ڳ����,
                                 Sum(Nvl(���, 0)) As ��ֵ, 0 As ��������
                          From Ԥ��
                          Where �տ�ʱ�� Between d_��ʼʱ�� And d_��ֹʱ�� And Nvl(���, 0) < 0
                          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd'))
                          Group By ���, �տ�ʱ��, ҵ������
                          Order By ���, �տ�ʱ��
                   ) Loop
    
      --     item_list[]
      --       business_type    C 1 ҵ������:�ڳ�����ֵ���շ��á����ʵ�
      --       happen_time      C 1 ����ʱ��:yyyy-mm-dd hh24:mi:ss
      --       earlystage       N 1 �ڳ����
      --       recharge         N 1 ���ڳ�ֵ
      --       consume          N 1 ��������
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'business_type', r_Ԥ����ϸ.ҵ������, 0, 1);
      zlJsonPutValue(v_Output, 'happen_time', r_Ԥ����ϸ.�տ�ʱ��);
      zlJsonPutValue(v_Output, 'earlystage', r_Ԥ����ϸ.�ڳ����, 1);
      zlJsonPutValue(v_Output, 'recharge', r_Ԥ����ϸ.���ڳ�ֵ, 1);
      zlJsonPutValue(v_Output, 'consume', r_Ԥ����ϸ.��������, 1, 2);
    
    End Loop;
  
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositdetail;
/

Create Or Replace Procedure Zl_Exsesvr_Getdepositlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡָ�����˵�Ԥ���嵥��Ϣ
  --��Σ�Json_In:��ʽ
  --   input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳID
  --      type              N 1 Ԥ����� 1-����;2-סԺ
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     item_list[]        ����
  --       create_date      C 1 �տ�����
  --       bill_no          C 1 ���ݺ�
  --       dept_name        C 1 ��������
  --       money            N 1 ���
  --       blnc_mode        N 1 ���㷽ʽ
  --       operator_name    C 1 ����Ա����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output      Varchar2(32767);
  c_Output      Clob;
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Type        Number(3);

  v_Err_Msg Varchar2(500);
Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Type        := j_Json.Get_Number('type');

  If Nvl(n_Pati_Id, 0) = 0 Then
    v_Err_Msg := 'δ���벡��id������!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  
  End If;
  For R In (Select LTrim(To_Char(a.�տ�ʱ��, 'YYYY-MM-DD')) As ����, a.No, b.���� As ����, a.���, a.���㷽ʽ, a.����Ա����
            From ����Ԥ����¼ A, ���ű� B
            Where a.����id = b.Id(+) And a.��¼���� = 1 And a.����id = n_Pati_Id And
                  (a.��ҳid = n_Pati_Pageid Or Nvl(n_Pati_Pageid, 0) = 0) And a.Ԥ����� = n_Type
            Order By a.�տ�ʱ�� Desc) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'create_date', r.����, 0, 1);
    zlJsonPutValue(v_Output, 'bill_no', r.No);
    zlJsonPutValue(v_Output, 'dept_name', r.����);
    zlJsonPutValue(v_Output, 'money', r.���);
    zlJsonPutValue(v_Output, 'blnc_mode', r.���㷽ʽ);
    zlJsonPutValue(v_Output, 'operator_name', r.����Ա����, 0, 2);
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositlist;
/

Create Or Replace Procedure Zl_Exsesvr_Getwriteoffinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --����:��ȡ����סԺ��������������Ϣ
  --��Σ�Json_In,��ʽ����
  --  input
  --    pati_id            N 1 ����ID
  --    pati_pageid        N 1 ��ҳID
  --    request_time       C 1 ��������ʱ��

  --    type               N 1 0-����Ƿ����δ�������������;1-��ȡ����סԺ����δ��˵���������
  --����: Json_Out,��ʽ����
  --  output
  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexists           N 1   type=0��2 ʱ���� 1-����;0-������
  --    fee_list[]δ������ʵĵ�����Ϣ type=1ʱ����
  --      no               C  1   ���ݺ�
  --      fee_name         C  1   �շ���Ŀ����
  --      dept_name        C  1   ��������

  --------------------------------------------------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  n_Pati_Id     סԺ���ü�¼.����id%Type;
  n_Pati_Pageid סԺ���ü�¼.��ҳid%Type;
  n_Type        Number;
  n_Count       Number;
  Vjtmp         Varchar2(32767);
  d_����ʱ��    Date;
Begin

  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Type        := Nvl(j_Json.Get_Number('type'), 0);

  If n_Type = 0 Then
    Select Count(1)
    Into n_Count
    From סԺ���ü�¼ A, ���˷������� B
    Where a.����id = n_Pati_Id And a.��ҳid = n_Pati_Pageid And b.����id = a.Id And b.״̬ = 0;
    If n_Count > 1 Then
      n_Count := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","isexists":' || n_Count || '}}';
  Elsif n_Type = 1 Then
    Vjtmp := Null;
    For R In (Select Distinct a.No, d.���� As ��Ŀ, c.���� As ����
              From סԺ���ü�¼ A, ���˷������� B, ���ű� C, �շ���ĿĿ¼ D
              Where a.Id = b.����id And a.�շ�ϸĿid = d.Id And b.��˲���id = c.Id(+) And b.���ʱ�� Is Null And a.����id = n_Pati_Id And
                    Nvl(a.��ҳid, 0) = n_Pati_Pageid) Loop
    
      Vjtmp := Vjtmp || ',{';
      Vjtmp := Vjtmp || '"no":"' || r.No || '"';
      Vjtmp := Vjtmp || ',"fee_name":"' || Zljsonstr(r.��Ŀ) || '"';
      Vjtmp := Vjtmp || ',"dept_name":"' || Zljsonstr(r.����) || '"';
      Vjtmp := Vjtmp || '}';
    End Loop;
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","fee_list":[' || Substr(Vjtmp, 2) || ']}}';
  Elsif n_Type = 2 Then
    d_����ʱ�� := To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
    Select Count(1) Into n_Count From ���˷������� B Where b.����ʱ�� = d_����ʱ��;
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","isexists":' || n_Count || '}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getwriteoffinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Chargeissuccessed
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ܣ���鿨�ѻ�Ԥ�����Ƿ��Ѿ�����ɹ�
  --��Σ�Json_In:��ʽ
  --  input
  --    cardfee_no       C  1 ���Ѷ�Ӧ�ķ��õ��ݺ�
  --    deposit_no       C  1 Ԥ�����ݺ�
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    is_successed      N   1 �Ƿ�ɹ�:1-�ɹ�;0-���ɹ�
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  v_Card_No    סԺ���ü�¼.No%Type;
  v_Deposit_No ����Ԥ����¼.No%Type;
  n_Exist      Number(1);
  v_Output     Varchar2(32767);

Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Card_No    := j_Json.Get_String('cardfee_no');
  v_Deposit_No := j_Json.Get_String('deposit_no');

  If Nvl(v_Card_No, '-') = '-' Or Nvl(v_Deposit_No, '-') = '-' Then
    Json_Out := zlJsonOut('ʧ�ܣ����봫��Ԥ��NO�ͷ���NO��');
    Return;
  End If;

  Select Count(1)
  Into n_Exist
  From סԺ���ü�¼ A, ���˽��ʼ�¼ B
  Where a.����id = b.Id And a.��¼���� In (5, 15) And a.��¼״̬ = 1 And b.��¼״̬ = 1 And a.No = v_Card_No;

  If n_Exist = 0 Then
    Select Count(1)
    Into n_Exist
    From ����Ԥ����¼
    Where NO = v_Deposit_No And ��¼���� = 5 And ��¼״̬ In (1, 3) And У�Ա�־ In (0, 2);
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'is_successed', n_Exist, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Chargeissuccessed;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatifee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------------------------------------------------------------
  --���ܣ���ȡ���˷��������Ϣ
  --��Σ�Json_In��ʽ
  --  input
  --    pati_id            N 1 ����ID
  --    pati_pageid        N 1 ��ҳID
  --    fee_type           C 0 �շ����
  --    baby_num           N 0 Ӥ����
  --���Σ�json_out��ʽ
  --fee_list      [����]  ÿ�����ü�¼
  --  id               N    ����id
  --  no               N    ���ݺ�
  --  fee_id           N    �շ�ϸĿid
  -------------------------------------------------------------------------------------------------
  j_Input       PLJson;
  j_Json        PLJson;
  v_Tmp         Varchar2(32767);
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Baby_Num    Number(3);
  v_Fee_Type    סԺ���ü�¼.�շ����%Type;
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  v_Fee_Type    := j_Json.Get_String('fee_type');
  n_Baby_Num    := j_Json.Get_Number('baby_num');
  For R In (Select a.Id, a.No, a.�շ�ϸĿid
            From סԺ���ü�¼ A
            Where a.����id = n_Pati_Id And a.��ҳid = n_Pati_Pageid And (a.�շ���� = v_Fee_Type Or '��' = Nvl(v_Fee_Type, '��')) And
                  (Nvl(a.Ӥ����, 0) = n_Baby_Num Or - 1 = n_Baby_Num)) Loop
    v_Tmp := v_Tmp || ',' || '{"id":' || r.Id || ',"no": "' || r.No || '","fee_id":' || r.�շ�ϸĿid || '}';
  End Loop;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�","fee_list":[' || Substr(v_Tmp, 2) || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatifee;
/

Create Or Replace Procedure Zl_Exsesvr_Getreceiveinvoice
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡƱ��������Ϣ
  --��Σ�Json_In:��ʽ
  --    input
  --      oper_fun  N 1 0-��ȡƱ��������Ϣ 1-��ȡ��ȡָ��Ʊ�ֵĹ���Ʊ������
  --      recv_ids C 1 ����ids:Ʊ������id,����ö���
  --      inv_type  N 1 Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      use_mode  N 1 ʹ�÷�ʽ:1-���ã���Ʊ�ݽ����������Լ�ʹ�ã�2-���ã���Ʊ���ɶ����Ա��ͬʹ��
  --      use_type C 1 Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
  --      recvtr  C 1 ������
  --      min_nums  N 1 ��Ʊ��������
  --      nodeno  C  1  վ��
  --����: Json_Out,��ʽ����
  --    output
  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message C 1 "Ӧ����Ϣ�� �ɹ�ʱ���سɹ���Ϣ,ʧ��ʱ���ؾ���Ĵ�����Ϣ"
  --    item_list C
  --      recv_id N 1 ����ID
  --      use_mode  N 1 ʹ�÷�ʽ:1-���ã���Ʊ�ݽ����������Լ�ʹ�ã�2-���ã���Ʊ���ɶ����Ա��ͬʹ��
  --      use_type C 1 Ʊ��ʹ�����:1,4: Ʊ��ʹ�����.����;2Ԥ��:1-����Ԥ��;2-סԺԤ��;5:�洢����ҽ�ƿ����.ID
  --      prefix_text C 1 ǰ׺�ı�
  --      start_no  C 1 ��ʼ����
  --      end_no  C 1 ��ֹ����
  --      inv_no_cur  C 1 ��ǰ����
  --      surplus_num C 1 ʣ������
  --      create_time C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --      use_time  C 1 ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
  --      recvtr  C 1 ������
  --      use_typecode      C 1 ʹ��������
  --      use_typeid        N 1 ʹ�����id

  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  n_����id       Number(18);
  n_Ʊ��         Number(2);
  n_ʹ�÷�ʽ     Number(2);
  v_Ʊ��ʹ����� Varchar2(100);
  v_������       Varchar2(100);
  v_����ids      Varchar2(4000);
  n_������ʽ     Number;
  v_Node         Varchar2(100);

  n_�������� Number(18);

  Cursor c_Ʊ��������Ϣ Is
    Select ID, ǰ׺�ı�, ��ǰ����, ��ʼ����, ��ֹ����, ʣ������, ʹ�÷�ʽ, ʹ�����, ������, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
           To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��
    From Ʊ�����ü�¼
    Where Rownum < 1;

  Cursor c_Ʊ��������Ϣ Is
    Select a.Id, '' As ʹ��������, a.ʹ����� As ʹ�����id, a.������� As ʹ�����, a.������, a.�Ǽ�ʱ��, a.��ʼ����, a.��ֹ����, a.ʣ������
    From Ʊ�����ü�¼ A, ��Ա�� B
    Where a.ʹ�÷�ʽ = 2 And a.ʣ������ > 0 And a.������ = b.���� And Rownum < 1;

  r_Ʊ������ c_Ʊ��������Ϣ%RowType;

  Type Ty_Ʊ������ Is Ref Cursor;
  c_Ʊ������ Ty_Ʊ������; --��̬�α����

  r_Ʊ������ c_Ʊ��������Ϣ%RowType;

  Type Ty_Ʊ������ Is Ref Cursor;
  c_Ʊ������ Ty_Ʊ������; --��̬�α����

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����ids      := j_Json.Get_String('recv_ids');
  n_Ʊ��         := j_Json.Get_Number('inv_type');
  n_ʹ�÷�ʽ     := j_Json.Get_Number('use_mode');
  v_Ʊ��ʹ����� := j_Json.Get_String('use_type');
  v_������       := j_Json.Get_String('recvtr');
  n_��������     := j_Json.Get_Number('min_nums');
  n_������ʽ     := j_Json.Get_Number('oper_fun');
  v_Node         := j_Json.Get_String('nodeno');

  If Nvl(n_Ʊ��, 0) = 0 And v_����ids Is Null Then
    Json_Out := zlJsonOut('δ����Ʊ����Ϣ!');
    Return;
  End If;
  If v_����ids Is Not Null Then
    If Instr(v_����ids, ',') = 0 Then
      n_����id := To_Number(v_����ids);
    End If;
  End If;
  If Nvl(n_������ʽ, 0) = 0 Then
    If v_����ids Is Not Null Then
      --������IDΪ��Ҫ��ѯ�������в�ѯ
      If Nvl(n_����id, 0) <> 0 Then
        Open c_Ʊ������ For
          Select ID, ǰ׺�ı�, ��ǰ����, ��ʼ����, ��ֹ����, ʣ������, ʹ�÷�ʽ, ʹ�����, ������, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
                 To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��
          From Ʊ�����ü�¼
          Where ID = n_����id And (Nvl(n_Ʊ��, 0) = 0 Or Ʊ�� = n_Ʊ��) And ʣ������ > 0 And
                (Nvl(ʹ�����, 'LXH') = v_Ʊ��ʹ����� Or ʹ����� Is Null Or v_Ʊ��ʹ����� Is Null) And Nvl(ʣ������, 0) >= Nvl(n_��������, 0);
      Else
        Open c_Ʊ������ For
          Select ID, ǰ׺�ı�, ��ǰ����, ��ʼ����, ��ֹ����, ʣ������, ʹ�÷�ʽ, ʹ�����, ������, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
                 To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��
          From Ʊ�����ü�¼
          Where ID In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(f_Num2List(v_����ids)) B) And (Nvl(n_Ʊ��, 0) = 0 Or Ʊ�� = n_Ʊ��) And ʣ������ > 0 And
                (Nvl(ʹ�����, 'LXH') = v_Ʊ��ʹ����� Or ʹ����� Is Null Or v_Ʊ��ʹ����� Is Null) And Nvl(ʣ������, 0) >= Nvl(n_��������, 0)
          Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, ʹ����� Desc, ��ʼ����;
      End If;
    Else
      Open c_Ʊ������ For
        Select ID, ǰ׺�ı�, ��ǰ����, ��ʼ����, ��ֹ����, ʣ������, ʹ�÷�ʽ, ʹ�����, ������, To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               To_Char(ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʹ��ʱ��
        From Ʊ�����ü�¼
        Where Ʊ�� = n_Ʊ�� And ʹ�÷�ʽ = Nvl(n_ʹ�÷�ʽ, 0) And ʣ������ > 0 And ������ = v_������ And
              (Nvl(ʹ�����, 'LXH') = v_Ʊ��ʹ����� Or ʹ����� Is Null) And Nvl(ʣ������, 0) >= Nvl(n_��������, 0)
        Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, ʹ����� Desc, ��ʼ����;
    End If;
  
    Loop
      Fetch c_Ʊ������
        Into r_Ʊ������;
      Exit When c_Ʊ������%NotFound;
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'recv_id', r_Ʊ������.Id, 1, 1);
      zlJsonPutValue(v_Output, 'use_mode', r_Ʊ������.ʹ�÷�ʽ);
      zlJsonPutValue(v_Output, 'use_type', Nvl(r_Ʊ������.ʹ�����, ''));
      zlJsonPutValue(v_Output, 'prefix_text', Nvl(r_Ʊ������.ǰ׺�ı�, ''));
      zlJsonPutValue(v_Output, 'start_no', Nvl(r_Ʊ������.��ʼ����, ''));
      zlJsonPutValue(v_Output, 'end_no', Nvl(r_Ʊ������.��ֹ����, ''));
      zlJsonPutValue(v_Output, 'inv_no_cur', Nvl(r_Ʊ������.��ǰ����, ''));
      zlJsonPutValue(v_Output, 'surplus_num', r_Ʊ������.ʣ������, 1);
      zlJsonPutValue(v_Output, 'create_time', r_Ʊ������.�Ǽ�ʱ��);
      zlJsonPutValue(v_Output, 'use_time', r_Ʊ������.ʹ��ʱ��);
      zlJsonPutValue(v_Output, 'recvtr', r_Ʊ������.������, 0, 2);
    
    End Loop;
  
  Else
    If Nvl(n_Ʊ��, 0) = 1 Or Nvl(n_Ʊ��, 0) = 3 Then
      --�շѺͽ���
      Open c_Ʊ������ For
        Select a.Id, Nvl(m.����, ' ') As ʹ��������, Null ʹ�����id, a.ʹ�����, a.������, a.�Ǽ�ʱ��, a.��ʼ����, a.��ֹ����, a.ʣ������
        From Ʊ�����ü�¼ A, ��Ա�� B, Ʊ��ʹ����� M
        Where a.Ʊ�� = n_Ʊ�� And a.ʹ�÷�ʽ = 2 And a.ʣ������ > 0 And a.������ = b.���� And a.ʹ����� = m.����(+) And
              (b.վ�� = v_Node Or b.վ�� Is Null)
        Order By ʹ��������, ʣ������ Desc;
    Elsif Nvl(n_Ʊ��, 0) = 5 Then
      --���￨
      Open c_Ʊ������ For
        Select a.Id, Null As ʹ��������, a.ʹ����� As ʹ�����id, a.ʹ�����, a.������, a.�Ǽ�ʱ��, a.��ʼ����, a.��ֹ����, a.ʣ������
        From Ʊ�����ü�¼ A, ��Ա�� B
        Where a.Ʊ�� = n_Ʊ�� And a.ʹ�÷�ʽ = 2 And a.ʣ������ > 0 And a.������ = b.���� And (b.վ�� = v_Node Or b.վ�� Is Null)
        Order By ʹ��������, ʣ������ Desc;
    Elsif Nvl(n_Ʊ��, 0) = 2 Then
      --Ԥ��
      Open c_Ʊ������ For
        Select a.Id, Null ʹ��������, Null ʹ�����id, To_Number(Nvl(a.ʹ�����, '0')) As ʹ�����, a.������, a.�Ǽ�ʱ��, a.��ʼ����, a.��ֹ����,
               a.ʣ������
        From Ʊ�����ü�¼ A, ��Ա�� B
        Where a.Ʊ�� = n_Ʊ�� And a.ʹ�÷�ʽ = 2 And a.ʣ������ > 0 And a.������ = b.���� And (b.վ�� = v_Node Or b.վ�� Is Null)
        Order By ʹ�����, ʣ������ Desc;
    Else
      Open c_Ʊ������ For
        Select a.Id, Null ʹ��������, Null ʹ�����id, a.ʹ�����, a.������, a.�Ǽ�ʱ��, a.��ʼ����, a.��ֹ����, a.ʣ������
        From Ʊ�����ü�¼ A, ��Ա�� B
        Where a.Ʊ�� = n_Ʊ�� And a.ʹ�÷�ʽ = 2 And a.ʣ������ > 0 And a.������ = b.���� And (b.վ�� = v_Node Or b.վ�� Is Null)
        Order By ʹ�����, ʣ������ Desc;
    End If;
    Loop
      Fetch c_Ʊ������
        Into r_Ʊ������;
      Exit When c_Ʊ������%NotFound;
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'recv_id', r_Ʊ������.Id, 1, 1);
      zlJsonPutValue(v_Output, 'use_typecode', r_Ʊ������.ʹ��������);
      zlJsonPutValue(v_Output, 'use_typeid', r_Ʊ������.ʹ�����id, 1);
      zlJsonPutValue(v_Output, 'use_type', r_Ʊ������.ʹ�����);
      zlJsonPutValue(v_Output, 'recvtr', r_Ʊ������.������);
      zlJsonPutValue(v_Output, 'create_time', r_Ʊ������.�Ǽ�ʱ��);
      zlJsonPutValue(v_Output, 'start_no', r_Ʊ������.��ʼ����);
      zlJsonPutValue(v_Output, 'end_no', r_Ʊ������.��ֹ����);
      zlJsonPutValue(v_Output, 'surplus_num', r_Ʊ������.ʣ������, 1, 2);
    
    End Loop;
  
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getreceiveinvoice;
/



Create Or Replace Procedure Zl_Exsesvr_Getnextinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݵ�ǰ��Ʊ�ż�Ʊ��ʹ����ϸ����ȡһ����Ч�ķ�Ʊ��
  --��Σ�Json_In:��ʽ
  -- input
  --   recv_id N 1 ����id:Ʊ������id
  --   inv_no  C 1 ��Ʊ��

  --����: Json_Out,��ʽ����
  -- output
  --   code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message C 1 Ӧ����Ϣ��  �ɹ�ʱ���سɹ���Ϣ ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --   inv_no  C 1 ��һ����Ʊ��
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id     Number(18);
  v_��ǰ��Ʊ�� Varchar2(100);
  n_Count      Number(2);

  v_ǰ׺�ı� Ʊ�����ü�¼.ǰ׺�ı�%Type;
  v_��ʼ���� Ʊ�����ü�¼.��ʼ����%Type;
  v_��ֹ���� Ʊ�����ü�¼.��ֹ����%Type;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := j_Json.Get_Number('recv_id');
  v_��ǰ��Ʊ�� := j_Json.Get_String('inv_no');
  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('δ����Ʊ��������Ϣ!');
    Return;
  End If;
  Begin
    Select Upper(ǰ׺�ı�) As ǰ׺�ı�, Upper(��ʼ����), Upper(��ֹ����)
    Into v_ǰ׺�ı�, v_��ʼ����, v_��ֹ����
    From Ʊ�����ü�¼
    Where ID = n_����id;
    n_Count := 1;
  Exception
    When Others Then
      n_Count := 0;
  End;

  If n_Count <> 0 Then
  
    For c_Ʊ�� In (
                 
                 Select Upper(����) As ����
                 From Ʊ��ʹ����ϸ
                 Where ���� || '' >= v_��ǰ��Ʊ�� And ����id = n_����id
                 Order By ����) Loop
      If Substr(v_��ǰ��Ʊ��, 1, Length(Nvl(v_ǰ׺�ı�, ''))) <> Nvl(v_ǰ׺�ı�, '') Then
        v_��ǰ��Ʊ�� := '';
        Exit;
      End If;
      If Not (v_��ǰ��Ʊ�� >= v_��ʼ���� And v_��ǰ��Ʊ�� <= v_��ֹ����) Then
        v_��ǰ��Ʊ�� := '';
        Exit;
      End If;
    
      Select Nvl(Max(1), 0) Into n_Count From Ʊ��ʹ����ϸ Where ���� = v_��ǰ��Ʊ�� And ����id = n_����id;
      If n_Count = 0 Then
        Exit;
      End If;
      v_��ǰ��Ʊ�� := Zl_Incstr(v_��ǰ��Ʊ��);
    End Loop;
  Else
    v_��ǰ��Ʊ�� := Null;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","inv_no":"' || Nvl(v_��ǰ��Ʊ��, '') || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getnextinvoice;
/


Create Or Replace Procedure Zl_Exsesvr_Updatecardinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���Ѱ�����ҽ��Ʊ��ʹ��ʱ��������ҽ��Ʊ�ݵ���ظ��²���
  --��Σ�Json_In:��ʽ
  --    input
  --      fun_oper  N 1 ��������:1-������2-�˿���3-�ش�4-����5-����
  --      fee_nos C 1 ���õ���s:����ö���
  --      recv_id N 1 ����id
  --      inv_no  C 1 ��ǰ��Ʊ�Ż�ʼʹ�÷�Ʊ��
  --      inv_usenums N 1 ��Ʊʹ������
  --      use_time  C 1 Ʊ��ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
  --      inv_user  C 1 ��Ʊʹ����
  --����: Json_Out,��ʽ����
  --   output
  --     code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message C 1 Ӧ����Ϣ�� �ɹ�ʱ���سɹ���Ϣ ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     inv_outnos  C 1 ����ҽ��Ʊ��:ʹ�õ�����ҽ��Ʊ��,����ö��ŷ���
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_��������     Number(2);
  v_���õ���     Varchar2(20);
  n_����id       Number(18);
  v_��ʼ��Ʊ��   Varchar2(100);
  n_��Ʊʹ������ Number(18);

  Cursor c_Fact(n_����id Ʊ�����ü�¼.Id%Type) Is
    Select * From Ʊ�����ü�¼ Where ID = Nvl(n_����id, 0);
  r_Factrow c_Fact%RowType;

  v_�ջ�id     Ʊ�ݴ�ӡ����.Id%Type;
  v_Ʊ�ݺ�     Ʊ��ʹ����ϸ.����%Type;
  v_��ǰƱ�ݺ� Ʊ��ʹ����ϸ.����%Type;
  n_��ӡid     Ʊ�ݴ�ӡ����.Id%Type;

  n_Ʊ�ݽ�� Ʊ��ʹ����ϸ.Ʊ�ݽ��%Type;

  v_ʹ��Ʊ����Ϣ Varchar2(4000);
  v_ʹ����       Ʊ��ʹ����ϸ.ʹ����%Type;
  d_ʹ��ʱ��     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type;
  v_ʹ��ʱ��     Varchar2(30);
  v_Err_Msg      Varchar2(255);
  Err_Item Exception;

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��������     := j_Json.Get_Number('fun_oper');
  v_���õ���     := j_Json.Get_String('fee_no');
  n_����id       := j_Json.Get_Number('recv_id');
  v_Ʊ�ݺ�       := j_Json.Get_String('inv_no');
  n_��Ʊʹ������ := j_Json.Get_Number('inv_usenums');
  v_ʹ����       := j_Json.Get_String('inv_user');
  v_ʹ��ʱ��     := j_Json.Get_String('use_time');

  If v_ʹ��ʱ�� Is Null Then
    d_ʹ��ʱ�� := Sysdate;
  Else
    d_ʹ��ʱ�� := To_Date(v_ʹ��ʱ��, 'yyyy-mm-dd hh24:mi:ss');
  End If;

  If v_���õ��� Is Null Then
    Json_Out := zlJsonOut('δ������Ҫ��ӡ�ĵ�����Ϣ!');
    Return;
  End If;

  --��Ʊ�ݺ�ʱ,���ô���Ʊ��
  If v_Ʊ�ݺ� Is Null Then
    Return;
  End If;

  --�˿�
  If n_�������� = 2 Then
    Begin
      --�����һ�δ�ӡ��������ȡ
      Select ID
      Into v_�ջ�id
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And a.Ʊ�� = 1 And b.�������� = 5 And b.No = v_���õ���
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_�ջ�id Is Not Null Then
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
        Select Ʊ��ʹ����ϸ_Id.Nextval, 1, v_Ʊ�ݺ�, 2, 2, ����id, ��ӡid, d_ʹ��ʱ��, v_ʹ����
        From Ʊ��ʹ����ϸ
        Where ��ӡid = v_�ջ�id And Ʊ�� = 1 And ���� = 1;
    End If;
    Return;
  End If;

  --�ش��ջ�ԭʼƱ��
  If n_�������� = 3 Or n_�������� = 5 Then
    Begin
      --�����һ�δ�ӡ��������ȡ
      Select ID
      Into v_�ջ�id
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And a.Ʊ�� = 1 And b.�������� = 5 And b.No = v_���õ���
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If v_�ջ�id Is Not Null Then
      Begin
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, Decode(n_��������, 5, 2, 4), ����id, ��ӡid, d_ʹ��ʱ��, v_ʹ����, Ʊ�ݽ��
          
          From Ʊ��ʹ����ϸ
          Where ��ӡid = v_�ջ�id And Ʊ�� = 1 And ���� = 1;
      Exception
        When Others Then
          Null;
      End;
    End If;
  End If;

  --Ʊ�ݴ�ӡ���
  Select Nvl(Sum(ʵ�ս��), 0) Into n_Ʊ�ݽ�� From סԺ���ü�¼ Where ��¼���� = 5 And NO = v_���õ���;

  Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  --���ɵ��ݵ�Ʊ�ݴ�ӡ����
  Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 5, v_���õ���);

  --������Ʊ��
  If Nvl(n_����id, 0) <> 0 Then
    Open c_Fact(n_����id);
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '��Ч��Ʊ���������Σ��޷���ɹҺ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.ʣ������, 0) < n_��Ʊʹ������ Then
      v_Err_Msg := '��ǰ���ε�ʣ����������' || n_��Ʊʹ������ || '�ţ��޷���ɹҺ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;
  v_ʹ��Ʊ����Ϣ := Null;
  For I In 1 .. n_��Ʊʹ������ Loop
    --���Ʊ�ݷ�Χ�Ƿ���ȷ
    If Nvl(n_����id, 0) <> 0 Then
      If Not (Upper(v_Ʊ�ݺ�) >= Upper(r_Factrow.��ʼ����) And Upper(v_Ʊ�ݺ�) <= Upper(r_Factrow.��ֹ����) And
          Length(v_Ʊ�ݺ�) = Length(r_Factrow.��ֹ����)) Then
        v_Err_Msg := '�õ�����Ҫ��ӡ����Ʊ��,��Ʊ�ݺ�"' || v_Ʊ�ݺ� || '"����Ʊ�����õĺ��뷶Χ��';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --����Ʊ��
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����, Ʊ�ݽ��)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 1, v_Ʊ�ݺ�, 1, Decode(n_��������, 3, 3, 1), n_����id, n_��ӡid, d_ʹ��ʱ��, v_ʹ����, n_Ʊ�ݽ��);
  
    v_ʹ��Ʊ����Ϣ := Nvl(v_ʹ��Ʊ����Ϣ, '') || ',' || v_Ʊ�ݺ�;
    v_��ǰƱ�ݺ�   := v_Ʊ�ݺ�;
    --��һ��Ʊ�ݺ�
    v_Ʊ�ݺ� := Zl_Incstr(v_Ʊ�ݺ�);
  End Loop;

  If Not v_ʹ��Ʊ����Ϣ Is Null Then
    v_ʹ��Ʊ����Ϣ := Substr(v_ʹ��Ʊ����Ϣ, 2);
  
  End If;
  If Nvl(n_����id, 0) <> 0 Then
    Update Ʊ�����ü�¼
    Set ʹ��ʱ�� = d_ʹ��ʱ��, ��ǰ���� = v_��ǰƱ�ݺ�, ʣ������ = Nvl(ʣ������, 0) - n_��Ʊʹ������
    Where ID = n_����id;
    Close c_Fact;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","inv_outnos":"' || v_ʹ��Ʊ����Ϣ || '"}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatecardinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillopercontrols
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���ݲ�����������
  --��Σ�Json_In:��ʽ
  --  input       
  --    bill_type  N  1  ��������:1-�Һŵ���,2-�շѵ�,3-���۵�,4-�������,5-סԺ����,6-Ԥ����,7-���ʵ���,8-���￨
  --    operator_id  N  1  ��ԱID

  --����: Json_Out,��ʽ����
  --   output      
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --      is_exist  N  1  ���ڿ�������:1-����;0-������
  --    time_limit  N  1  0(NULL)-������,n-n����
  --    other_bill  N  1  �Ƿ�������������ݽ��в���
  --    uplimit_money  N  1  �������

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_�������� Number(2);
  n_��Աid   Number(18);
  n_Count    Number(5);

  n_ʱ������ ���ݲ�������.ʱ������%Type;
  n_���˵��� ���ݲ�������.���˵���%Type;
  n_������� ���ݲ�������.�������%Type;

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�������� := j_Json.Get_Number('bill_type');
  n_��Աid   := j_Json.Get_Number('operator_id');

  Select Max(1), Max(Nvl(ʱ������, 0)), Max(Nvl(���˵���, 0)), Max(Nvl(�������, 0))
  Into n_Count, n_ʱ������, n_���˵���, n_�������
  From ���ݲ�������
  Where ��Աid = n_��Աid And ���� = n_��������;

  Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '"';
  Json_Out := Json_Out || ',"is_exist":' || Nvl(n_Count, 0);
  Json_Out := Json_Out || ',"time_limit":' || Nvl(n_ʱ������, 0) || '';
  Json_Out := Json_Out || ',"other_bill":' || Nvl(n_���˵���, 0);
  Json_Out := Json_Out || ',"uplimit_money":' || Nvl(n_�������, 0);
  Json_Out := Json_Out || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillopercontrols;
/


Create Or Replace Procedure Zl_Exsesvr_Getfullno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�Զ����뵥�ݺ�
  --��Σ�Json_In:��ʽ
  --    input
  --      item_num  N 1 ��Ŀ���
  --      input_no  C 1 ����ĵ��ݺ�
  --      dept_id   N   ����ID

  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    full_no           C       �����ĵ��ݺ�
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  n_���   ������Ʊ�.��Ŀ���%Type;
  v_No     ������Ʊ�.������%Type;
  n_����id ���ű�.Id%Type := Null;
  v_No_Out ������Ʊ�.������%Type;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_���   := j_Json.Get_Number('item_num');
  v_No     := j_Json.Get_String('input_no');
  n_����id := j_Json.Get_Number('dept_id');
  If Nvl(n_���, 0) = 0 Then
    Json_Out := zlJsonOut('δ������ţ����飡');
    Return;
  End If;

  If Nvl(v_No, '-') = '-' Then
    Json_Out := zlJsonOut('δ����NO�����飡');
    Return;
  End If;
  v_No_Out := Fullno(n_���, v_No, n_����id);
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","full_no":"' || Nvl(v_No_Out, '') || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfullno;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatitotalmoney
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���ID,��ҳID��ҽ��id����ȡӦ�ա�ʵ���ܶ�
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_source N 1 ������Դ:0-����;1-����;2-סԺ
  --    pati_id N 1 ����ID
  --    visit_id  N   ����ID:סԺʱ��������ҳid,�����ݴ�NULL
  --    advice_ids  C   ҽ��ids:����ö��ŷ���
  --    today_fee N   �Ƿ��շ���:1-�ǵ�;0-������
  --    price_tag N   ���۱�־:0-������;1-�������۵�;2-��ͳ�ƻ��۵�
  --����: Json_Out,��ʽ����
  --  output
  --    code        C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message     C  1  Ӧ����Ϣ��
  --    fee_amrcvb  N  1  Ӧ�ս��
  --    fee_ampaib  N  1  ʵ�ս��
  ------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);

  n_����id   Number(18);
  n_����id   Number(18);
  n_������Դ Number(2);

  v_ҽ��ids  Varchar2(32767);
  n_���շ��� Number(2);
  n_���۱�־ Number(2);
  v_��¼״̬ Varchar2(10);
  n_ʵ�ս�� ������ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս�� ������ü�¼.Ӧ�ս�� %Type;

  Err_Item Exception;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_������Դ := Nvl(j_Json.Get_Number('pati_source'), 0);
  n_����id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_����id   := j_Json.Get_Number('visit_id');
  v_ҽ��ids  := j_Json.Get_String('advice_ids');
  n_���շ��� := Nvl(j_Json.Get_Number('today_fee'), 0);
  n_���۱�־ := Nvl(j_Json.Get_Number('only_price'), 0);

  --0-����;1-����;2-סԺ
  v_��¼״̬ := ',0,1,2,3,';
  If n_���۱�־ = 1 Then
    --0-������;1-�������۵�;2-��ͳ�ƻ��۵�
    v_��¼״̬ := ',1,2,3,';
  Elsif n_���۱�־ = 2 Then
    v_��¼״̬ := ',0,';
  End If;

  If v_ҽ��ids Is Not Null Then
    Select Sum(Ӧ�ս��), Sum(ʵ�ս��)
    Into n_ʵ�ս��, n_Ӧ�ս��
    From (With ҽ������ As (Select Distinct Column_Value As ҽ��id From Table(f_Num2List(v_ҽ��ids)))
           Select /*+cardinality(B,10) */
            Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.Ӧ�ս��) As Ӧ�ս��
           
           From סԺ���ü�¼ A, ҽ������ B
           Where a.ҽ����� = b.ҽ��id
           Union All
           Select /*+cardinality(B,10) */
            Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.Ӧ�ս��) As Ӧ�ս��
           
           From ������ü�¼ A, ҽ������ B
           Where a.ҽ����� = b.ҽ��id);
  
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '�ɹ�');
    zlJsonPutValue(v_Output, 'fee_ampaib', Nvl(n_ʵ�ս��, 0), 1);
    zlJsonPutValue(v_Output, 'fee_amrcvb', Nvl(n_Ӧ�ս��, 0), 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
    Return;
  
  End If;

  If Nvl(n_������Դ, 0) = 0 Or Nvl(n_������Դ, 0) = 1 Then
    --��ѯ���з��ü�����
    If Nvl(n_���շ���, 0) = 1 Then
      --�鵱�շ���
      Select Sum(ʵ�ս��), Sum(Ӧ�ս��)
      Into n_ʵ�ս��, n_Ӧ�ս��
      From (Select Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.Ӧ�ս��) As Ӧ�ս��
             From סԺ���ü�¼ A
             Where ����id = n_����id And (n_����id Is Null Or Nvl(��ҳid, 0) = Nvl(n_����id, 0)) And
                   Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And (n_������Դ = 0 Or n_������Դ = 1 And �����־ <> 2) And
                   ����ʱ�� >= Trunc(Sysdate) And ����ʱ�� <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And Nvl(a.���ʷ���, 0) = 1
             Union All
             Select Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.Ӧ�ս��) As Ӧ�ս��
             From ������ü�¼ A
             Where ����id = n_����id And (n_����id Is Null Or Nvl(��ҳid, 0) = Nvl(n_����id, 0)) And
                   Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And ����ʱ�� >= Trunc(Sysdate) And
                   ����ʱ�� <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And Nvl(a.���ʷ���, 0) = 1);
    
    Else
      Select Sum(ʵ�ս��), Sum(Ӧ�ս��)
      Into n_ʵ�ս��, n_Ӧ�ս��
      From (Select Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.Ӧ�ս��) As Ӧ�ս��
             From סԺ���ü�¼ A
             Where ����id = n_����id And (n_����id Is Null Or Nvl(��ҳid, 0) = Nvl(n_����id, 0)) And
                   Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And (n_������Դ = 0 Or n_������Դ = 1 And �����־ <> 2) And
                   Nvl(a.���ʷ���, 0) = 1
             Union All
             Select Sum(a.ʵ�ս��) As ʵ�ս��, Sum(a.Ӧ�ս��) As Ӧ�ս��
             From ������ü�¼ A
             Where ����id = n_����id And (n_����id Is Null Or Nvl(��ҳid, 0) = Nvl(n_����id, 0)) And
                   Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And Nvl(a.���ʷ���, 0) = 1);
    End If;
  Else
    --��סԺ 
    If Nvl(n_���շ���, 0) = 1 Then
      --�鵱�շ���
      Select Sum(ʵ�ս��), Sum(Ӧ�ս��)
      Into n_ʵ�ս��, n_Ӧ�ս��
      From סԺ���ü�¼ A
      Where ����id = n_����id And Nvl(a.���ʷ���, 0) = 1 And a.�����־ = 2 And (n_����id Is Null Or Nvl(��ҳid, 0) = Nvl(n_����id, 0)) And
            Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And ����ʱ�� >= Trunc(Sysdate) And
            ����ʱ�� <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And Nvl(a.���ʷ���, 0) = 1 And a.�����־ = 2;
    Else
      Select Sum(ʵ�ս��), Sum(Ӧ�ս��)
      Into n_ʵ�ս��, n_Ӧ�ս��
      From סԺ���ü�¼ A
      Where ����id = n_����id And Nvl(a.���ʷ���, 0) = 1 And a.�����־ = 2 And (n_����id Is Null Or Nvl(��ҳid, 0) = Nvl(n_����id, 0)) And
            Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0;
    End If;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'fee_ampaib', Nvl(n_ʵ�ս��, 0), 1);
  zlJsonPutValue(v_Output, 'fee_amrcvb', Nvl(n_Ӧ�ս��, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatitotalmoney;
/


Create Or Replace Procedure Zl_Exsesvr_Getbilltotalmoney
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ָ���ĵ��ݺţ���ȡӦ�ա�ʵ���ܶ�
  --��Σ�Json_In:��ʽ
  --  input       
  --    fee_origin  N  1  ����ҵԴ:1-����;2-סԺ
  --    bill_type  N  1  ��������:1-�շѵ�;3-�Զ����ʵ���2 -���ʼ�¼��4-�Һż�¼ ;5-���￨
  --    fee_no  C  1  ���õ���
  --    pati_id  N  1  ����id
  --    rec_status  C    ��¼״̬:���Զ��״̬,����:0,1

  --����: Json_Out,��ʽ����
  --  output      
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ��
  --    fee_amrcvb  N  1  Ӧ�ս��
  --    fee_ampaib  N  1  ʵ�ս��
  ---------------------------------------------------------------------------
  j_Input     PLJson;
  j_Json      PLJson;
  n_������Դ  Number(2);
  n_��������  Number(2);
  v_���ݺ�    Varchar2(100);
  n_����id    Number(18);
  v_��¼״̬s Varchar2(100);

  n_ʵ�ս�� ������ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս�� ������ü�¼.Ӧ�ս�� %Type;

  --��װʧ��ʱ���ص�����
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_������Դ  := j_Json.Get_Number('fee_origin');
  n_��������  := j_Json.Get_Number('bill_type');
  v_���ݺ�    := j_Json.Get_String('fee_no');
  n_����id    := j_Json.Get_Number('pati_id');
  v_��¼״̬s := j_Json.Get_String('rec_status');

  If v_��¼״̬s Is Null Then
    v_��¼״̬s := ',0,1,';
  Else
    v_��¼״̬s := ',' || v_��¼״̬s || ',';
  End If;

  If Nvl(n_������Դ, 0) <= 1 Then
    Select Sum(ʵ�ս��), Sum(Ӧ�ս��)
    Into n_ʵ�ս��, n_Ӧ�ս��
    From ������ü�¼
    Where NO = v_���ݺ� And ��¼���� = n_�������� And Instr(v_��¼״̬s, ',' || ��¼״̬ || ',') > 0 And
          (Nvl(n_����id, 0) = 0 Or ����id = Nvl(n_����id, 0));
  Else
    Select Sum(ʵ�ս��), Sum(Ӧ�ս��)
    Into n_ʵ�ս��, n_Ӧ�ս��
    From סԺ���ü�¼
    Where NO = v_���ݺ� And ��¼���� = n_�������� And Instr(v_��¼״̬s, ',' || ��¼״̬ || ',') > 0 And
          (Nvl(n_����id, 0) = 0 Or ����id = Nvl(n_����id, 0));
  End If;
  Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '"';
  Json_Out := Json_Out || ',"fee_ampaib":' || zlJsonStr(Nvl(n_ʵ�ս��, 0), 1);
  Json_Out := Json_Out || ',"fee_amrcvb":' || zlJsonStr(Nvl(n_Ӧ�ս��, 0), 1);
  Json_Out := Json_Out || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbilltotalmoney;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillinfobyno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݵ��ݺŻ�ȡ���ݵĵ�����Ϣ���磺�Ǽ��ˣ��Ǽ�ʱ��
  -- input     
  --   fee_origin  N 1 ������Դ:1-����;2-סԺ
  --   bill_type N 1 ��������:1-�շѵ�;3-�Զ����ʵ���2 -���ʼ�¼��4-�Һż�¼ ;5-���￨;-1-���ʵ�;-2-Ԥ����;-3-�������
  --   bill_no C 1 ���ݺ�
  --����: Json_Out,��ʽ����
  --  output     
  --   code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message C 1 Ӧ����Ϣ�� 
  --   operator_name C 1 ����Ա����
  --   create_time C 1 �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  --   pati_id N 1 ����id
  ---------------------------------------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);

  n_������Դ Number(2);
  n_�������� Number(2);
  v_���ݺ�   Varchar2(100);
  n_����id   Number(18);

  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_�Ǽ�ʱ��   Varchar2(30);

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_������Դ := j_Json.Get_Number('fee_origin');
  n_�������� := Nvl(j_Json.Get_Number('bill_type'), 0);
  v_���ݺ�   := j_Json.Get_String('bill_no');

  If n_�������� = -1 Then
    --1.����
    Select Max(����Ա����), To_Char(Max(�շ�ʱ��), 'yyyy-mm-dd hh24mi:ss'), Max(����id)
    Into v_����Ա����, v_�Ǽ�ʱ��, n_����id
    From ���˽��ʼ�¼
    Where NO = v_���ݺ� And ��¼״̬ In (1, 3);
  Elsif n_�������� = -2 Then
    --2.Ԥ��
    Select Max(����Ա����), To_Char(Max(�տ�ʱ��), 'yyyy-mm-dd hh24:mi:ss'), Max(����id)
    Into v_����Ա����, v_�Ǽ�ʱ��, n_����id
    From ����Ԥ����¼
    Where NO = v_���ݺ� And ��¼״̬ In (1, 3);
  Elsif n_�������� = -3 Then
    --3.�������
    Select Max(����Ա����), To_Char(Max(�Ǽ�ʱ��), 'yyyy-mm-dd hh24:mi:ss'), Max(����id)
    Into v_����Ա����, v_�Ǽ�ʱ��, n_����id
    From ���ò����¼
    Where NO = v_���ݺ� And ��¼״̬ In (1, 3);
  Else
    --4.�������
    Begin
      If Nvl(n_������Դ, 0) <= 1 Then
        Select Nvl(����Ա����, ������), To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), ����id
        Into v_����Ա����, v_�Ǽ�ʱ��, n_����id
        From ������ü�¼
        Where NO = v_���ݺ� And ��¼���� = n_�������� And ��¼״̬ In (0, 1, 3) And Rownum < 2;
      Else
        Select Nvl(����Ա����, ������), To_Char(�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss'), ����id
        Into v_����Ա����, v_�Ǽ�ʱ��, n_����id
        From סԺ���ü�¼
        Where NO = v_���ݺ� And ��¼���� = n_�������� And ��¼״̬ In (0, 1, 3) And Rownum < 2;
      End If;
    
    Exception
      When Others Then
        v_����Ա���� := Null;
        v_�Ǽ�ʱ��   := '';
        n_����id     := Null;
    End;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'operator_name', Nvl(v_����Ա����, ''));
  zlJsonPutValue(v_Output, 'create_time', Nvl(v_�Ǽ�ʱ��, ''));
  zlJsonPutValue(v_Output, 'pati_id', Nvl(n_����id, 0), 1, 2);

  v_Output := '{"output":' || v_Output || '}';
  Json_Out := v_Output;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillinfobyno;
/


Create Or Replace Procedure Zl_Exsesvr_Getpatiinvoiceclass
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡƱ��ʹ�����
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id              N 1 ����id
  --   pati_pageid          N 1 ��ҳid
  --   insure_type          N 1 ����

  --����: Json_Out,��ʽ����
  --  output
  --    code                C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    use_type            C  1  Ʊ��ʹ�����
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_����id   Number(18);
  n_��ҳid   Number(18);
  n_����     Number(18);
  v_ʹ����� Varchar2(4000);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_����     := j_Json.Get_Number('insure_type');
  v_ʹ����� := Zl_Billclass(n_����id, n_��ҳid, n_����);

  Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '","use_type":"' || Nvl(v_ʹ�����, '') || '"}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatiinvoiceclass;
/

Create Or Replace Procedure Zl_Exsesvr_Invoiceclassused
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�ж�ָ��Ʊ���Ƿ������˷�ʹ������ӡ
  --��Σ�Json_In:��ʽ
  --  input
  --    inv_type            N  1  Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨

  --����: Json_Out,��ʽ����
  --  output
  --    code                C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1  Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    is_start            N  1  �Ƿ�����:1-�����˵ģ�0-δ����

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Ʊ�� Number(5);
  n_���� Number(2);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Ʊ�� := j_Json.Get_Number('inv_type');

  Select Nvl(Max(1), 0)
  Into n_����
  From Ʊ�����ü�¼
  Where Ʊ�� = n_Ʊ�� And Nvl(ʹ�����, 'LXH') <> 'LXH' And Rownum < 2;

  Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '","is_start":' || Nvl(n_����, 0) || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Invoiceclassused;
/


Create Or Replace Procedure Zl_Exsesvr_Patimove_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ����˻���ǰ������ؼ��
  --��Σ�Json_In��ʽ
  --  input
  --    pati_id            N 1 ����ID
  --    pati_pageid        N 1 ��ҳID
  --    operator_time      C 0 ����ʱ��
  --����: Json_Out,��ʽ����
  --  output
  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------------------------------------------------------------
  j_Input         PLJson;
  j_Json          PLJson;
  v_Tmp           Varchar2(20);
  n_Pati_Id       Number(18);
  n_Pati_Pageid   Number(18);
  d_Operator_Time Date;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id       := j_Json.Get_Number('pati_id');
  n_Pati_Pageid   := j_Json.Get_Number('pati_pageid');
  v_Tmp           := j_Json.Get_String('operator_time');
  d_Operator_Time := To_Date(v_Tmp, 'YYYY-MM-DD HH24:MI:SS');

  For r_Fee In (Select NO
                From סԺ���ü�¼
                Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= d_Operator_Time And
                      �շ���� = 'J'
                Group By NO, ���, Mod(��¼����, 10)
                Having Sum(���ʽ��) <> 0) Loop
    Json_Out := zlJsonOut('�䶯ʱ��֮�������ѽ��ʵ��Զ����ʷ���,���ܽ��л���������');
    Return;
  End Loop;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Patimove_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Billinhistory
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ѯָ���ķ��õ����Ƿ���ں󱸱�ռ���
  --��Σ�Json_In:��ʽ
  --    input
  --      bill_no           C 1 ���ݺ�
  --      bill_type         C 1 ��������:1-�շѵ�,2-Ԥ����,3-���ʵ�,4-�Һŵ�,5-���￨����,6-���ʵ���;7-�Զ����ʵ�
  --      outpati_flag      N 1 �����־��1-���2-סԺ
  --����: Json_Out,��ʽ����
  --  output
  --    code              N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exits_history     C   1   ������ʷ�󱸱�:1-����;1-������
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_�������� Number(1);
  n_�����־ Number(1);
  v_���ݺ�   Varchar2(100);

  v_Output  Varchar2(32767);
  n_Nomoved Number(2);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ�   := j_Json.Get_String('bill_no');
  n_�������� := j_Json.Get_Number('bill_type');
  n_�����־ := j_Json.Get_Number('outpati_flag');

  If Nvl(v_���ݺ�, '-') = '-' Then
    Json_Out := zlJsonOut('δ���뵥�ݺţ����飡');
    Return;
  End If;

  If Nvl(n_��������, 0) = 0 Then
    Json_Out := zlJsonOut('δ���뵥�����ͣ����飡');
    Return;
  End If;

  If Nvl(n_�����־, 0) = 0 And Nvl(n_��������, 0) = 6 Then
    Json_Out := zlJsonOut('δ���������־�����飡');
    Return;
  End If;

  n_Nomoved := Zl_Fun_Checkinhistory(n_��������, v_���ݺ�, n_�����־);

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'exits_history', n_Nomoved, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Billinhistory;
/


Create Or Replace Procedure Zl_Exsesvr_Billisprintinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�ж�ָ���ĵ����Ƿ��ӡ�˷�Ʊ
  --��Σ�Json_In:��ʽ
  -- input     
  --    bill_no  C 1 ���ݺ�
  --    bill_type N 1 ��������:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
  --    inv_type  N 1 Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨

  --����: Json_Out,��ʽ����
  --  output      
  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    printed N 1 �Ƿ��ӡ:1-�Ѵ�ӡ;0-δ��ӡ
  --
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  v_���ݺ�   Varchar2(100);
  n_�������� Number(2);
  n_Ʊ��     Number(2);
  n_�Ƿ��ӡ Number(2);

  --��װʧ��ʱ���ص�����
  v_Output Varchar2(32767);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ�   := j_Json.Get_String('bill_no');
  n_�������� := j_Json.Get_Number('bill_type');
  n_Ʊ��     := j_Json.Get_Number('inv_type');

  Begin
    Select 1
    Into n_�Ƿ��ӡ
    From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
    Where a.��ӡid = b.Id And a.Ʊ�� = n_Ʊ�� And b.No = v_���ݺ� And b.�������� = n_�������� And Rownum < 2;
  
  Exception
    When Others Then
      n_�Ƿ��ӡ := 0;
  End;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'printed', n_�Ƿ��ӡ, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Billisprintinvoice;
/


Create Or Replace Procedure Zl_Exsesvr_Getcardfeeinfobyno
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݿ��ѵ��ݺţ���ȡ���Ѽ������Ѽ�Ԥ����������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    fee_no  C 1 ���ݺ�:���õ��ݺ�
  --    query_type  N   ��ѯ���ͣ�0-��ȡ��������:1-��ȡ���ϵ���;2-ʣ����õ���
  --    query_deposit N 1 �Ƿ����Ԥ��:1-����Ԥ����Ϣ��0-������Ԥ����Ϣ
  --����: Json_Out,��ʽ����
  -- output
  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message C 1 Ӧ����Ϣ��
  --    fee_list      [����]ÿ������ID��Ϣ
  --      fee_id  N 1 ����id
  --      fee_num N 1 ���
  --      pati_id N 1 ����id
  --      pati_name C 1 ����
  --      pati_sex  C 1 �Ա�
  --      pati_age  C 1 ����
  --      fee_category  C 1 �ѱ�
  --      item_id N 1 �շ���Ŀid
  --      income_item_id  N 1 ������Ŀid
  --      quantity  N 1 ����
  --      fee_amrcvb  N 1 Ӧ�ս��
  --      fee_ampaid  N 1 ʵ�ս��
  --      placer  C 1 ������
  --      operator_code C 1 ����Ա���
  --      operator_name C 1 ����Ա����
  --      create_time D 1 �Ǽ�ʱ��
  --      happen_time D 1 ����ʱ��
  --      rec_status  N 1 ��¼״̬
  --      mrbkfee_sign N 1 �Ƿ�����:1-�ǲ�����;0-���ǲ�����
  --      invoice_no  N 1 ��Ʊ��
  --      kpbooks_sign N 1 ���ʱ�־:1-�Ǽ���;0-����
  --      fee_status N 1 ����״̬:1-�쳣״̬;0-��������
  --      cardtype_id N 1 �����ID
  --      card_no C 1 ����
  --      sendcard_reg  N 1 �Ƿ�Һŷ���:1-�ǹҺ�ͬʱ����;0-�ǹҺ�ͬʱ����
  --    pricebill_info  C    �������ɻ��۷�����Ϣ
  --      fee_no  C    ���۵���
  --      cardfee_amrcvb  N 1 ����Ӧ�ս��
  --      cardfee_ampaid  N 1 ����ʵ�ս��
  --      mrbkfee_amrcvb N 1 ������Ӧ��
  --      mrbkfee_ampaid N 1 ������ʵ��
  --      charged_statu N 1 �շ�״̬:0-δ�շ�;1-���շ�;2-��ȫ��
  --    balance_list[]  C   ������Ϣ
  --      blnc_mode C 1 ���㷽ʽ����
  --      balance_id  N 1 ����ID�� ��ѯ���ϵĵ���ʱΪ����ID
  --      blnc_money  N 1 �����ܶ�
  --      pay_cardno  N 1 ֧������
  --      pay_swapno  C 1 ������ˮ��
  --      pay_swapmemo  C 1 ����˵��
  --      relation_id N 1 ��������id
  --      cardtype_id N 1 �����id
  --      consume_card  N 1 �Ƿ����ѿ�:1-��;0-����
  --      blnc_nature N 1 ��������:1-�ֽ���㷽ʽ,2-������ҽ������ , 8-���㿨���� ,9-����
  --      blnc_statu  N 1 ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
  --      consume_card_id N 1 ���ѿ�id
  --      blnc_no C 1 �������
  --      blnc_memo C 1 ժҪ
  --      original_id N 1 ԭ����ID:����ʱ����
  --      original_money N,1 ԭʼ���,��ʣ�����ʱ����

  --   deposit_info  C   Ԥ����Ϣ:query_deposit=1ʱ��Ч��ȱʡ����
  --      deposit_id  N 1 Ԥ��id
  --      deposit_no  C 1 Ԥ�����ݺ�
  --      deposit_money N 1 Ԥ�����
  --      blnc_mode C 1 ���㷽ʽ
  --      pay_cardno  N 1 ֧������
  --      pay_swapno  C 1 ������ˮ��
  --      pay_swapmemo  C 1 ����˵��
  --      relation_id N 1 ��������id
  --      cardtype_id N 1 �����id
  --      consume_card  N 1 �Ƿ����ѿ�:1-��;0-����
  --      blnc_mode C 1 ���㷽ʽ
  --      blnc_nature N 1 ��������
  --      blnc_statu  N 1 ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
  --      consume_card_id N 1 ���ѿ�id
  --      blnc_no C 1 �������
  --      blnc_memo C 1 ժҪ
  ---------------------------------------------------------------------------
  j_Input        PLJson;
  j_Json         PLJson;
  v_���ݺ�       Varchar2(100);
  n_��ѯ����     Number;
  v_��¼״̬     Varchar2(10);
  v_��Ʊ��       Ʊ��ʹ����ϸ.����%Type;
  v_���۵�       Varchar2(100);
  n_ʵ�ս��     ������ü�¼.ʵ�ս��%Type;
  n_ʣ����     ������ü�¼.ʵ�ս��%Type;
  n_��ѯԤ��     Number(2);
  n_����Ӧ��     ������ü�¼.ʵ�ս��%Type;
  n_����ʵ��     ������ü�¼.ʵ�ս��%Type;
  n_������Ӧ��   ������ü�¼.ʵ�ս��%Type;
  n_������ʵ��   ������ü�¼.ʵ�ս��%Type;
  n_�Һ�ͬ������ Number(2);
  n_����id       Number(18);
  n_ԭ����id     Number(18);
  n_Find         Number(2);
  n_Nomoved      Number(2);
  Cursor c_������Ϣ Is
    Select a.Id, a.No, a.��¼״̬, a.���, a.�ѱ�, a.����, a.�Ա�, a.����, a.����id, a.�շ�ϸĿid, a.������Ŀid, a.ʵ��Ʊ��, a.����,
           Decode(n_��ѯ����, 1, -1, 1) * a.Ӧ�ս�� As Ӧ�ս��, Decode(n_��ѯ����, 1, -1, 1) * a.ʵ�ս�� As ʵ�ս��, a.���ʷ���,
           Nvl(a.�Ӱ��־, 0) As �䶯���, a.����Ա����, a.����Ա���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
           To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(a.���ӱ�־, 0), 8, 1, 0) As ������, a.���� As �����id,
           a.����id, a.������, a.����״̬, a.ժҪ, a.ʵ��Ʊ�� As ����
    From סԺ���ü�¼ A
    Where a.��¼���� = 5 And NO = '-' And Rownum < 1;

  r_���� c_������Ϣ%RowType;

  Cursor c_������Ϣ Is
    Select a.No, a.���㷽ʽ, Nvl(a.��Ԥ��, 0) As ��Ԥ��, a.��������id, a.�����id, a.����, a.���㿨���, a.������ˮ��, a.����˵��, b.����, a.У�Ա�־,
           Max(c.���ѿ�id) As ���ѿ�id, Max(a.�������) As �������, Max(a.ժҪ) As ժҪ
    From ����Ԥ����¼ A, ���㷽ʽ B, ���˿������¼ C
    Where a.���㷽ʽ = ����(+) And a.����id = 1 And a.��¼���� = 1 And a.Id = c.����id(+);

  r_������Ϣ c_������Ϣ%RowType;

  Cursor c_Ԥ����Ϣ Is
    Select a.No, a.Id As Ԥ��id, Nvl(Sum(a.���), 0) As ���, Max(a.���㷽ʽ) As ���㷽ʽ, Nvl(Sum(a.��Ԥ��), 0) As ��Ԥ��,
           Max(a.��������id) As ��������id, Max(a.�����id) As �����id, Max(Decode(a.��¼����, 1, a.����, '')) As ����, Max(a.���㿨���) As ���㿨���,
           Max(Decode(a.��¼����, 1, a.������ˮ��, '')) As ������ˮ��, Max(Decode(a.��¼����, 1, a.����˵��, '')) ����˵��, Max(b.����) As ����,
           Max(Decode(a.��¼����, 1, a.У�Ա�־, 0)) As У�Ա�־, Max(c.���ѿ�id) As ���ѿ�id, Max(a.�������) As �������, Max(a.ժҪ) As ժҪ
    From ����Ԥ����¼ A, ���㷽ʽ B, ���˿������¼ C
    Where a.���㷽ʽ = ����(+) And a.��������id = 0 And Mod(a.��¼����, 10) = 1 And a.Id = c.����id(+)
    Group By NO;

  r_Ԥ����Ϣ c_Ԥ����Ϣ%RowType;

  Type t_��Ϣ Is Ref Cursor;

  c_��Ϣ t_��Ϣ; --��̬�α����

  n_�ܶ� סԺ���ü�¼.ʵ�ս��%Type;
  --��װʧ��ʱ���ص�����
  v_Priebill   Varchar2(32767);
  v_Balanceinf Varchar2(32767);
  v_Deposit    Varchar2(32767);
  v_Output     Varchar2(32767);

  c_Output Clob;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ�   := j_Json.Get_String('fee_no');
  n_��ѯ���� := Nvl(j_Json.Get_Number('query_type'), 0);

  n_��ѯԤ�� := Nvl(j_Json.Get_Number('query_deposit'), 1);

  v_��¼״̬ := ',1,3,';
  If Nvl(n_��ѯ����, 0) = 1 Then
    v_��¼״̬ := ',2,';
  Elsif n_��ѯ���� = 2 Then
    --ֻ��ѯʣ����õ���
    v_��¼״̬ := ',1,2,3,';
  End If;

  Select Nvl(Max(1), 0) Into n_Nomoved From HסԺ���ü�¼ A Where NO = v_���ݺ� And ��¼���� = 5 And Rownum <= 1;

  If Nvl(n_Nomoved, 0) = 0 Then
    Select Max(m.����)
    Into v_��Ʊ��
    From Ʊ�ݴ�ӡ���� N, Ʊ��ʹ����ϸ M
    Where n.�������� = 5 And n.Id = m.��ӡid And m.���� = 1 And m.Ʊ�� = 1 And
          m.ʹ��ʱ�� = (Select Max(M2.ʹ��ʱ��)
                    From Ʊ�ݴ�ӡ���� N2, Ʊ��ʹ����ϸ M2
                    Where M2.��ӡid = N2.Id And n.�������� = 5 And M2.Ʊ�� = 1 And N2.No = v_���ݺ�) And n.No = v_���ݺ�;
  Else
    Select Max(m.����)
    Into v_��Ʊ��
    From HƱ�ݴ�ӡ���� N, HƱ��ʹ����ϸ M
    Where n.�������� = 5 And n.Id = m.��ӡid And m.���� = 1 And m.Ʊ�� = 1 And
          m.ʹ��ʱ�� = (Select Max(M2.ʹ��ʱ��)
                    From HƱ�ݴ�ӡ���� N2, HƱ��ʹ����ϸ M2
                    Where M2.��ӡid = N2.Id And n.�������� = 5 And M2.Ʊ�� = 1 And N2.No = v_���ݺ�) And n.No = v_���ݺ�;
  End If;

  --�ȶ�ȡ����
  If Nvl(n_Nomoved, 0) = 0 Then
  
    Select Max(1)
    Into n_�Һ�ͬ������
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 1 And
          (����id, �Ǽ�ʱ��) In (Select ����id, �Ǽ�ʱ��
                           From סԺ���ü�¼
                           Where ��¼���� = 5 And NO = v_���ݺ� And ��¼״̬ In (0, 1, 3) And Rownum < 2);
  
    If Nvl(n_��ѯ����, 0) = 1 Then
      Select Max(����id) Into n_ԭ����id From סԺ���ü�¼ Where ��¼���� = 5 And ��¼״̬ In (1, 3) And NO = v_���ݺ�;
    End If;
  
    If Nvl(n_��ѯ����, 0) = 1 Then
      --����
      Open c_��Ϣ For
        Select a.Id, a.No, a.��¼״̬, a.���, a.�ѱ�, a.����, a.�Ա�, a.����, a.����id, a.�շ�ϸĿid, a.������Ŀid, a.ʵ��Ʊ��, a.����,
               Decode(n_��ѯ����, 1, -1, 1) * a.Ӧ�ս�� As Ӧ�ս��, Decode(n_��ѯ����, 1, -1, 1) * a.ʵ�ս�� As ʵ�ս��, a.���ʷ���,
               Nvl(a.�Ӱ��־, 0) As �䶯���, a.����Ա����, a.����Ա���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(a.���ӱ�־, 0), 8, 1, 0) As ������, a.���� As �����id,
               a.����id, a.������, a.����״̬, a.ժҪ, a.ʵ��Ʊ�� As ����
        From סԺ���ü�¼ A
        Where ����id In (Select Max(����id) As ����id
                       From סԺ���ü�¼
                       Where ��¼���� = 5 And ��¼״̬ = 2 And NO = v_���ݺ� And Nvl(���ӱ�־, 0) <> 8)
        Order By NO, ���;
    Elsif n_��ѯ���� = 2 Then
      --ʣ������
      Open c_��Ϣ For
        Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, a.No, Null As ��¼״̬, a.���, Max(a.�ѱ�) As �ѱ�, Max(a.����) As ����,
               Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, a.�շ�ϸĿid As �շ�ϸĿid, a.������Ŀid, Max(a.ʵ��Ʊ��) As ʵ��Ʊ��,
               Sum(a.����) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Max(a.���ʷ���) As ���ʷ���,
               Max(Nvl(a.�Ӱ��־, 0)) As �䶯���, Max(Decode(a.��¼״̬, 2, Null, a.����Ա����)) As ����Ա����,
               Max(Decode(a.��¼״̬, 2, Null, a.����Ա���)) As ����Ա���,
               To_Char(Max(Decode(a.��¼״̬, 2, To_Date(Null), a.�Ǽ�ʱ��)), 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               To_Char(Max(Decode(a.��¼״̬, 2, To_Date(Null), a.����ʱ��)), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��,
               Max(Decode(Nvl(a.���ӱ�־, 0), 8, 1, 0)) As ������, Max(a.����) As �����id,
               Max(Decode(a.��¼״̬, 2, Null, a.����id)) As ����id, Max(Decode(a.��¼״̬, 2, Null, a.������)) As ������,
               Max(a.����״̬) As ����״̬, Max(a.ժҪ) As ժҪ, Max(a.ʵ��Ʊ��) As ����
        From סԺ���ü�¼ A
        Where a.��¼���� = 5 And NO = v_���ݺ�
        Group By a.No, a.���, a.�շ�ϸĿid, a.������Ŀid
        Having Sum(a.����) <> 0
        Order By a.No, a.���;
    Else
    
      Open c_��Ϣ For
        Select a.Id, a.No, a.��¼״̬, a.���, a.�ѱ�, a.����, a.�Ա�, a.����, a.����id, a.�շ�ϸĿid, a.������Ŀid, a.ʵ��Ʊ��, a.����,
               Decode(n_��ѯ����, 1, -1, 1) * a.Ӧ�ս�� As Ӧ�ս��, Decode(n_��ѯ����, 1, -1, 1) * a.ʵ�ս�� As ʵ�ս��, a.���ʷ���,
               Nvl(a.�Ӱ��־, 0) As �䶯���, a.����Ա����, a.����Ա���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(a.���ӱ�־, 0), 8, 1, 0) As ������, a.���� As �����id,
               a.����id, a.������, a.����״̬, a.ժҪ, a.ʵ��Ʊ�� As ����
        From סԺ���ü�¼ A
        Where a.��¼���� = 5 And Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And NO = v_���ݺ�
        Order By NO, ���;
    End If;
  Else
  
    Select Max(1)
    Into n_�Һ�ͬ������
    From H������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 1 And
          (����id, �Ǽ�ʱ��) In (Select ����id, �Ǽ�ʱ��
                           From HסԺ���ü�¼
                           Where ��¼���� = 5 And NO = v_���ݺ� And Rownum < 2 And ��¼״̬ In (0, 1, 3));
  
    If Nvl(n_��ѯ����, 0) = 1 Then
      Select Max(����id)
      Into n_ԭ����id
      From HסԺ���ü�¼
      Where ��¼���� = 5 And ��¼״̬ In (1, 3) And NO = v_���ݺ�;
    
    End If;
  
    If Nvl(n_��ѯ����, 0) = 1 Then
      --����
      Open c_��Ϣ For
        Select a.Id, a.No, a.��¼״̬, a.���, a.�ѱ�, a.����, a.�Ա�, a.����, a.����id, a.�շ�ϸĿid, a.������Ŀid, a.ʵ��Ʊ��, a.����,
               Decode(n_��ѯ����, 1, -1, 1) * a.Ӧ�ս�� As Ӧ�ս��, Decode(n_��ѯ����, 1, -1, 1) * a.ʵ�ս�� As ʵ�ս��, a.���ʷ���,
               Nvl(a.�Ӱ��־, 0) As �䶯���, a.����Ա����, a.����Ա���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(a.���ӱ�־, 0), 8, 1, 0) As ������, a.���� As �����id,
               a.����id, a.������, a.����״̬, a.ժҪ, a.ʵ��Ʊ�� As ����
        From HסԺ���ü�¼ A
        Where ����id In (Select Max(����id) As ����id
                       From HסԺ���ü�¼
                       Where ��¼���� = 5 And ��¼״̬ = 2 And NO = v_���ݺ� And Nvl(���ӱ�־, 0) <> 8)
        Order By NO, ���;
    Elsif n_��ѯ���� = 2 Then
      --ʣ������
      Open c_��Ϣ For
        Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, a.No, Null As ��¼״̬, a.���, Max(a.�ѱ�) As �ѱ�, Max(a.����) As ����,
               Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.����id) As ����id, a.�շ�ϸĿid As �շ�ϸĿid, a.������Ŀid, Max(a.ʵ��Ʊ��) As ʵ��Ʊ��,
               Sum(a.����) As ����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Max(a.���ʷ���) As ���ʷ���,
               Max(Nvl(a.�Ӱ��־, 0)) As �䶯���, Max(Decode(a.��¼״̬, 2, Null, a.����Ա����)) As ����Ա����,
               Max(Decode(a.��¼״̬, 2, Null, a.����Ա���)) As ����Ա���,
               To_Char(Max(Decode(a.��¼״̬, 2, To_Date(Null), a.�Ǽ�ʱ��)), 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               To_Char(Max(Decode(a.��¼״̬, 2, To_Date(Null), a.����ʱ��)), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��,
               Max(Decode(Nvl(a.���ӱ�־, 0), 8, 1, 0)) As ������, Max(a.����) As �����id,
               Max(Decode(a.��¼״̬, 2, Null, a.����id)) As ����id, Max(Decode(a.��¼״̬, 2, Null, a.������)) As ������,
               Max(a.����״̬) As ����״̬, Max(a.ժҪ) As ժҪ, Max(a.ʵ��Ʊ��) As ����
        From HסԺ���ü�¼ A
        Where a.��¼���� = 5 And NO = v_���ݺ�
        Group By a.No, a.���, a.�շ�ϸĿid, a.������Ŀid
        Having Sum(a.����) <> 0
        Order By a.No, a.���;
    Else
      Open c_��Ϣ For
        Select a.Id, a.No, a.��¼״̬, a.���, a.�ѱ�, a.����, a.�Ա�, a.����, a.����id, a.�շ�ϸĿid, a.������Ŀid, a.ʵ��Ʊ��, a.����,
               Decode(n_��ѯ����, 1, -1, 1) * a.Ӧ�ս�� As Ӧ�ս��, Decode(n_��ѯ����, 1, -1, 1) * a.ʵ�ս�� As ʵ�ս��, a.���ʷ���,
               Nvl(a.�Ӱ��־, 0) As �䶯���, a.����Ա����, a.����Ա���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
               To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(a.���ӱ�־, 0), 8, 1, 0) As ������, a.���� As �����id,
               a.����id, a.������, a.����״̬, a.ժҪ, a.ʵ��Ʊ�� As ����
        From HסԺ���ü�¼ A
        Where a.��¼���� = 5 And Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And NO = v_���ݺ�
        Order By NO, ���;
    End If;
  End If;

  v_���۵� := Null;
  Loop
    Fetch c_��Ϣ
      Into r_����;
    Exit When c_��Ϣ%NotFound;
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
    --1.ȡ������Ϣ
    zlJsonPutValue(v_Output, 'fee_id', r_����.Id, 1, 1);
    zlJsonPutValue(v_Output, 'fee_num', r_����.���, 1);
  
    zlJsonPutValue(v_Output, 'pati_id', r_����.����id, 1);
  
    zlJsonPutValue(v_Output, 'pati_name', r_����.����);
    zlJsonPutValue(v_Output, 'pati_sex', Nvl(r_����.�Ա�, ''));
  
    zlJsonPutValue(v_Output, 'pati_age', Nvl(r_����.����, ''));
  
    zlJsonPutValue(v_Output, 'fee_category', Nvl(r_����.�ѱ�, ''));
  
    zlJsonPutValue(v_Output, 'item_id', r_����.�շ�ϸĿid, 1);
    zlJsonPutValue(v_Output, 'income_item_id', Nvl(r_����.������Ŀid, 0), 1);
  
    zlJsonPutValue(v_Output, 'quantity', Nvl(r_����.����, 0), 1);
  
    zlJsonPutValue(v_Output, 'fee_amrcvb', Nvl(r_����.Ӧ�ս��, 0), 1);
    zlJsonPutValue(v_Output, 'fee_ampaid', Nvl(r_����.ʵ�ս��, 0), 1);
  
    zlJsonPutValue(v_Output, 'placer', Nvl(r_����.������, ''));
    zlJsonPutValue(v_Output, 'operator_code', Nvl(r_����.����Ա���, ''));
    zlJsonPutValue(v_Output, 'operator_name', Nvl(r_����.����Ա����, ''));
  
    zlJsonPutValue(v_Output, 'create_time', Nvl(r_����.�Ǽ�ʱ��, ''));
    zlJsonPutValue(v_Output, 'happen_time', Nvl(r_����.����ʱ��, ''));
  
    zlJsonPutValue(v_Output, 'rec_status', Nvl(r_����.��¼״̬, 0), 1);
    zlJsonPutValue(v_Output, 'mrbkfee_sign', Nvl(r_����.������, 0), 1);
  
    zlJsonPutValue(v_Output, 'invoice_no', Nvl(v_��Ʊ��, ''));
  
    zlJsonPutValue(v_Output, 'kpbooks_sign', Nvl(r_����.���ʷ���, 0), 1);
    zlJsonPutValue(v_Output, 'fee_status', Nvl(r_����.����״̬, 0), 1);
  
    zlJsonPutValue(v_Output, 'cardtype_id', To_Number(Nvl(r_����.�����id, '0')), 1);
    zlJsonPutValue(v_Output, 'card_no', r_����.����);
  
    zlJsonPutValue(v_Output, 'sendcard_reg', Nvl(n_�Һ�ͬ������, 0), 1, 2);
  
    n_�ܶ� := Nvl(n_�ܶ�, 0) + Nvl(r_����.ʵ�ս��, 0);
  
    n_����id := Nvl(r_����.����id, 0);
    If Nvl(r_����.���ʷ���, 0) = 1 Then
      n_ԭ����id := Null;
    End If;
    If v_���۵� Is Null And r_����.ժҪ Is Not Null Then
      Select Max(No) Into v_���۵� From ������ü�¼ Where no = r_����.ժҪ And ��¼���� = 1 And ����ID = r_����.����id;
    End If;
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not v_���۵� Is Null Then
    If Nvl(n_Nomoved, 0) = 0 Then
      Select Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 0, 1) * Ӧ�ս��),
             Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 0, 1) * ʵ�ս��),
             Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 1, 0) * Ӧ�ս��),
             Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 1, 0) * ʵ�ս��), Sum(Decode(��¼״̬, 2, 0, 1) * ʵ�ս��),
             Nvl(Sum(ʵ�ս��), 0) - Nvl(Sum(���ʽ��), 0)
      Into n_����Ӧ��, n_����ʵ��, n_������Ӧ��, n_������ʵ��, n_ʵ�ս��, n_ʣ����
      From ������ü�¼ A
      Where Mod(��¼����, 10) = 1 And NO = v_���۵�;
    Else
      Select Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 0, 1) * Ӧ�ս��),
             Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 0, 1) * ʵ�ս��),
             Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 1, 0) * Ӧ�ս��),
             Sum(Decode(��¼״̬, 2, 0, 1) * Decode(a.���ӱ�־, 8, 1, 0) * ʵ�ս��), Sum(Decode(��¼״̬, 2, 0, 1) * ʵ�ս��),
             Nvl(Sum(ʵ�ս��), 0) - Nvl(Sum(���ʽ��), 0)
      Into n_����Ӧ��, n_����ʵ��, n_������Ӧ��, n_������ʵ��, n_ʵ�ս��, n_ʣ����
      From H������ü�¼ A
      Where Mod(��¼����, 10) = 1 And NO = v_���۵�;
    End If;
    zlJsonPutValue(v_Priebill, 'fee_no', v_���۵�, 0, 1);
    zlJsonPutValue(v_Priebill, 'cardfee_amrcvb', Nvl(n_����Ӧ��, 0));
    zlJsonPutValue(v_Priebill, 'cardfee_ampaid', Nvl(n_����ʵ��, 0));
    zlJsonPutValue(v_Priebill, 'mrbkfee_amrcvb', Nvl(n_������Ӧ��, 0));
    zlJsonPutValue(v_Priebill, 'mrbkfee_ampaid', Nvl(n_������ʵ��, 0));
  
    If Nvl(n_ʣ����, 0) = Nvl(n_ʵ�ս��, 0) Then
      --�շ�״̬:0-δ�շ�;1-���շ�;2-��ȫ��
      zlJsonPutValue(v_Priebill, 'charged_statu', 0, 1, 2);
    Elsif Nvl(n_ʣ����, 0) = 0 Then
      zlJsonPutValue(v_Priebill, 'charged_statu', 2, 1, 2);
    Else
      zlJsonPutValue(v_Priebill, 'charged_statu', 1, 1, 2);
    End If;
  End If;

  If v_Priebill Is Not Null Then
    v_Priebill := ',"pricebill_info":' || v_Priebill;
  End If;

  If Not c_Output Is Null Then
    c_Output := To_Clob(',"fee_list":[') || c_Output || To_Clob(']') || To_Clob(v_Priebill);
  Elsif Length(Nvl(v_Output, '') || Nvl(v_Priebill, '')) > 32767 Then
    c_Output := To_Clob(',"fee_list":[') || To_Clob(v_Output) || To_Clob(']') || To_Clob(v_Priebill);
    v_Output := '';
  Else
    v_Output := ',"fee_list":[' || v_Output || ']' || Nvl(v_Priebill, '');
  End If;

  If Nvl(n_����id, 0) = 0 Then
    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�"') || c_Output || '}}';
    Else
      Json_Out := '{"output":{"code":1,"message":"�ɹ�"' || v_Output || '}}';
    End If;
    Return;
  End If;

  Close c_��Ϣ;

  If Nvl(n_Nomoved, 0) = 0 Then
    Open c_��Ϣ For
      Select a.No, a.���㷽ʽ, Decode(n_��ѯ����, 1, -1, 1) * Nvl(a.��Ԥ��, 0) As ��Ԥ��, a.��������id, a.�����id, a.����, a.���㿨���, a.������ˮ��,
             a.����˵��, b.����, a.У�Ա�־, c.���ѿ�id, a.�������, a.ժҪ
      From ����Ԥ����¼ A, ���㷽ʽ B, ���˿������¼ C
      Where a.���㷽ʽ = ����(+) And a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And a.Id = c.����id(+);
  Else
    Open c_��Ϣ For
      Select a.No, a.���㷽ʽ, Decode(n_��ѯ����, 1, -1, 1) * Nvl(a.��Ԥ��, 0) As ��Ԥ��, a.��������id, a.�����id, a.����, a.���㿨���, a.������ˮ��,
             a.����˵��, b.����, a.У�Ա�־, c.���ѿ�id, a.�������, a.ժҪ
      From H����Ԥ����¼ A, ���㷽ʽ B, H���˿������¼ C
      Where a.���㷽ʽ = ����(+) And a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And a.Id = c.����id(+);
  End If;

  Loop
    Fetch c_��Ϣ
      Into r_������Ϣ;
    Exit When c_��Ϣ%NotFound;
  
    zlJsonPutValue(v_Balanceinf, 'blnc_mode', r_������Ϣ.���㷽ʽ, 0, 1);
    zlJsonPutValue(v_Balanceinf, 'balance_id', n_����id, 1);
    If n_��ѯ���� = 2 And Nvl(r_������Ϣ.����, 0) <> 9 Then
      --ֻ��һ������,��������ʱ
      zlJsonPutValue(v_Balanceinf, 'blnc_money', n_�ܶ�, 1);
    Else
      zlJsonPutValue(v_Balanceinf, 'blnc_money', Nvl(r_������Ϣ.��Ԥ��, 0), 1);
    End If;
  
    zlJsonPutValue(v_Balanceinf, 'pay_cardno', Nvl(r_������Ϣ.����, ''));
    zlJsonPutValue(v_Balanceinf, 'pay_swapno', Nvl(r_������Ϣ.������ˮ��, ''));
    zlJsonPutValue(v_Balanceinf, 'pay_swapmemo', Nvl(r_������Ϣ.����˵��, ''));
  
    zlJsonPutValue(v_Balanceinf, 'relation_id', Nvl(r_������Ϣ.��������id, 0), 1);
  
    If Nvl(r_������Ϣ.���㿨���, 0) <> 0 Then
      zlJsonPutValue(v_Balanceinf, 'cardtype_id', Nvl(r_������Ϣ.���㿨���, 0), 1);
      zlJsonPutValue(v_Balanceinf, 'consume_card', 1, 1);
    Else
      zlJsonPutValue(v_Balanceinf, 'cardtype_id', Nvl(r_������Ϣ.�����id, 0), 1);
      zlJsonPutValue(v_Balanceinf, 'consume_card', 0, 1);
    End If;
  
    zlJsonPutValue(v_Balanceinf, 'consume_card_id', To_Number(Nvl(r_������Ϣ.���ѿ�id, '0')), 1);
    zlJsonPutValue(v_Balanceinf, 'blnc_nature', Nvl(r_������Ϣ.����, 0), 1);
    zlJsonPutValue(v_Balanceinf, 'blnc_statu', Nvl(r_������Ϣ.У�Ա�־, 0), 1);
    zlJsonPutValue(v_Balanceinf, 'blnc_no', Nvl(r_������Ϣ.�������, ''));
    zlJsonPutValue(v_Balanceinf, 'blnc_memo', Nvl(r_������Ϣ.ժҪ, ''));
    If n_��ѯ���� = 2 And Nvl(r_������Ϣ.����, 0) <> 9 Then
      zlJsonPutValue(v_Balanceinf, 'original_money', Nvl(r_������Ϣ.��Ԥ��, 0), 1);
    Else
      zlJsonPutValue(v_Balanceinf, 'original_money', 0, 1);
    End If;
  
    zlJsonPutValue(v_Balanceinf, 'original_id', Nvl(n_ԭ����id, 0), 1, 2);
  
  End Loop;

  If v_Balanceinf Is Not Null Then
    v_Balanceinf := ',"balance_list":[' || v_Balanceinf || ']';
  
    If Not c_Output Is Null Then
      c_Output := c_Output || To_Clob(v_Balanceinf);
    
    Elsif Length(Nvl(v_Output, '') || Nvl(v_Balanceinf, '')) > 32767 Then
      c_Output := To_Clob(v_Output) || To_Clob(Nvl(v_Balanceinf, ''));
      v_Output := '';
    Else
      v_Output := Nvl(v_Output, '') || Nvl(v_Balanceinf, '');
    End If;
  
  End If;

  --3.Ԥ������
  If Nvl(n_��ѯԤ��, 0) = 1 Then
  
    Close c_��Ϣ;
  
    If Nvl(n_Nomoved, 0) = 0 Then
      Open c_��Ϣ For
        Select a.No, Max(a.Id) As Ԥ��id, Decode(n_��ѯ����, 1, -1, 1) * Nvl(Sum(a.���), 0) As ���, Max(a.���㷽ʽ) As ���㷽ʽ,
               Nvl(Sum(a.���), 0) As ��Ԥ��, Max(a.��������id) As ��������id, Max(a.�����id) As �����id,
               Max(Decode(a.��¼����, 1, a.����, '')) As ����, Max(a.���㿨���) As ���㿨���,
               Max(Decode(a.��¼����, 1, a.������ˮ��, '')) As ������ˮ��, Max(Decode(a.��¼����, 1, a.����˵��, '')) ����˵��, Max(b.����) As ����,
               Max(Decode(a.��¼����, 1, a.У�Ա�־, 0)) As У�Ա�־, Max(c.���ѿ�id) As ���ѿ�id, Max(a.�������) As �������, Max(a.ժҪ) As ժҪ
        From ����Ԥ����¼ A, ���㷽ʽ B, ���˿������¼ C
        Where a.���㷽ʽ = ����(+) And a.��������id In (Select ��������id From ����Ԥ����¼ Where ����id = n_����id) And Mod(a.��¼����, 10) = 1 And
              a.Id = c.����id(+) And (Nvl(n_��ѯ����, 0) = 1 And a.��¼״̬ = 2 Or Nvl(n_��ѯ����, 0) <> 1)
        Group By NO;
    Else
      Open c_��Ϣ For
        Select a.No, Max(a.Id) As Ԥ��id, Decode(n_��ѯ����, 1, -1, 1) * Nvl(Sum(a.���), 0) As ���, Max(a.���㷽ʽ) As ���㷽ʽ,
               Nvl(Sum(a.���), 0) As ��Ԥ��, Max(a.��������id) As ��������id, Max(a.�����id) As �����id,
               Max(Decode(a.��¼����, 1, a.����, '')) As ����, Max(a.���㿨���) As ���㿨���,
               Max(Decode(a.��¼����, 1, a.������ˮ��, '')) As ������ˮ��, Max(Decode(a.��¼����, 1, a.����˵��, '')) ����˵��, Max(b.����) As ����,
               Max(Decode(a.��¼����, 1, a.У�Ա�־, 0)) As У�Ա�־, Max(c.���ѿ�id) As ���ѿ�id, Max(a.�������) As �������, Max(a.ժҪ) As ժҪ
        From H����Ԥ����¼ A, ���㷽ʽ B, H���˿������¼ C
        Where a.���㷽ʽ = ����(+) And a.��������id In (Select ��������id From H����Ԥ����¼ Where ����id = n_����id) And Mod(a.��¼����, 10) = 1 And
              a.Id = c.����id(+) And (Nvl(n_��ѯ����, 0) = 1 And a.��¼״̬ = 2 Or Nvl(n_��ѯ����, 0) <> 1)
        Group By NO;
    End If;
    Loop
    
      Fetch c_��Ϣ
        Into r_Ԥ����Ϣ;
      Exit When c_��Ϣ%NotFound;
    
      --ֻ��һ��
    
      zlJsonPutValue(v_Deposit, 'deposit_id', r_Ԥ����Ϣ.Ԥ��id, 1, 1);
      zlJsonPutValue(v_Deposit, 'deposit_no', r_Ԥ����Ϣ.No);
      zlJsonPutValue(v_Deposit, 'deposit_money', Nvl(r_Ԥ����Ϣ.���, 0), 1);
    
      zlJsonPutValue(v_Deposit, 'blnc_mode', r_Ԥ����Ϣ.���㷽ʽ);
      zlJsonPutValue(v_Deposit, 'balance_id', n_����id, 1);
      zlJsonPutValue(v_Deposit, 'pay_cardno', Nvl(r_Ԥ����Ϣ.����, ''));
      zlJsonPutValue(v_Deposit, 'pay_swapno', Nvl(r_Ԥ����Ϣ.������ˮ��, ''));
      zlJsonPutValue(v_Deposit, 'pay_swapmemo', Nvl(r_Ԥ����Ϣ.����˵��, ''));
    
      zlJsonPutValue(v_Deposit, 'relation_id', Nvl(r_Ԥ����Ϣ.��������id, 0), 1);
    
      If Nvl(r_Ԥ����Ϣ.���㿨���, 0) <> 0 Then
        zlJsonPutValue(v_Deposit, 'cardtype_id', Nvl(r_Ԥ����Ϣ.���㿨���, 0), 1);
        zlJsonPutValue(v_Deposit, 'consume_card', 1, 1);
      Else
        zlJsonPutValue(v_Deposit, 'cardtype_id', Nvl(r_Ԥ����Ϣ.�����id, 0), 1);
        zlJsonPutValue(v_Deposit, 'consume_card', 0, 1);
      End If;
      zlJsonPutValue(v_Deposit, 'consume_card_id', To_Number(Nvl(r_������Ϣ.���ѿ�id, '0')), 1);
    
      zlJsonPutValue(v_Deposit, 'blnc_nature', Nvl(r_Ԥ����Ϣ.����, 0), 1);
      zlJsonPutValue(v_Deposit, 'blnc_statu', Nvl(r_Ԥ����Ϣ.У�Ա�־, 0), 1);
      zlJsonPutValue(v_Deposit, 'blnc_no', Nvl(r_������Ϣ.�������, ''));
      zlJsonPutValue(v_Deposit, 'blnc_memo', Nvl(r_������Ϣ.ժҪ, ''), 0, 2);
    
      n_Find := 1;
      Exit;
    End Loop;
  
    If v_Deposit Is Not Null Then
      v_Deposit := ',"deposit_info":' || v_Deposit;
    
      If Not c_Output Is Null Then
        c_Output := c_Output || To_Clob(v_Deposit);
      Elsif Length(Nvl(v_Output, '') || Nvl(v_Deposit, '')) > 32767 Then
        c_Output := To_Clob(v_Output) || To_Clob(Nvl(v_Deposit, ''));
        v_Output := '';
      
      Else
        v_Output := Nvl(v_Output, '') || Nvl(v_Deposit, '');
      End If;
    End If;
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�"') || c_Output || '}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"' || v_Output || '}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getcardfeeinfobyno;
/

Create Or Replace Procedure Zl_Exsesvr_Chkfeecategorydept
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ѱ����������п���,��ǰָ������
  --��Σ�Json_In:��ʽ
  --  input
  --    fee_category         N 1 �ѱ�
  --    pati_deptid         N 1 ���˿���ID
  --����: Json_Out,��ʽ����
  --  output
  --    code            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist          N 1 �ѱ��Ƿ���ڣ�0-�����ڣ�1-����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Count        Number(1);
  n_Pati_Deptid  Number(18);
  v_Fee_Category Varchar2(20);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Fee_Category := j_Json.Get_String('fee_category');
  n_Pati_Deptid  := j_Json.Get_Number('pati_deptid');
  Select Count(1)
  Into n_Count
  From Dual
  Where Not Exists (Select 1 From �ѱ����ÿ��� Where �ѱ� = v_Fee_Category) Or Exists
   (Select 1 From �ѱ����ÿ��� Where �ѱ� = v_Fee_Category And ����id = n_Pati_Deptid);
  If n_Count > 1 Then
    n_Count := 1;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Chkfeecategorydept;
/


Create Or Replace Procedure Zl_Exsesvr_Cardfeeisbalance
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ���鿨���Ƿ��Ѿ�����
  --���      json
  --input     
  --  cardfee_no            C 1 ���Ѷ�Ӧ�ķ��õ��ݺ�
  --  rdcardfee_sign        N 1 ��ȡ���ѱ�־:0-��ȡ����,1-������;2-���ѻ�������
  --����      json
  --output      
  --  code                      C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  isbalanced                N 1 �Ƿ��Ѿ�����:1-�ѽ����;0-δ����
  --  blnc_no                   C 1 ���ʵ��ݺ�
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  v_Output   Varchar2(32767);
  v_No       סԺ���ü�¼.No%Type;
  n_��־     Number(1);
  n_�ѽ���   Number(2);
  v_���ʵ��� ���˽��ʼ�¼.No%Type;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_No   := j_Json.Get_String('cardfee_no');
  n_��־ := Nvl(j_Json.Get_Number('rdcardfee_sign'), 0);

  If n_��־ = 0 Then
    Select Max(b.No), Count(1)
    Into v_���ʵ���, n_�ѽ���
    From סԺ���ü�¼ A, ���˽��ʼ�¼ B
    Where a.����id = b.Id And a.��¼���� In (5, 15) And a.��¼״̬ = 1 And b.��¼״̬ = 1 And a.No = v_No And Nvl(a.���ӱ�־, 0) <> 8;
  Elsif n_��־ = 1 Then
    Select Max(b.No), Count(1)
    Into v_���ʵ���, n_�ѽ���
    From סԺ���ü�¼ A, ���˽��ʼ�¼ B
    Where a.����id = b.Id And a.��¼���� In (5, 15) And a.��¼״̬ = 1 And b.��¼״̬ = 1 And a.No = v_No And Nvl(a.���ӱ�־, 0) = 8;
  Else
    Select Max(b.No), Count(1)
    Into v_���ʵ���, n_�ѽ���
    From סԺ���ü�¼ A, ���˽��ʼ�¼ B
    Where a.����id = b.Id And a.��¼���� In (5, 15) And a.��¼״̬ = 1 And b.��¼״̬ = 1 And a.No = v_No;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'isbalanced', n_�ѽ���, 1);
  zlJsonPutValue(v_Output, 'blnc_no', v_���ʵ���, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Cardfeeisbalance;
/


Create Or Replace Procedure Zl_Exsesvr_Recalcfee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����δ���������
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id            N 1  ����id
  --      pati_pageid        N 1  ��ҳID
  --      pati_nature        N 1 ��������:0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
  --      outfee             N 1  �Ƿ�����ѱ� 
  --      fee_type           C 1  �ѱ�
  --����: Json_Out,��ʽ����
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Outfee      Number;
  n_Pati_Nature Number;
  v_Feetype     Varchar2(1000);
Begin
  --�������
  j_Input       := PLJson(Json_In);
  j_Json        := j_Input.Get_Pljson('input');
  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Outfee      := j_Json.Get_Number('outfee');
  n_Pati_Nature := j_Json.Get_Number('pati_nature');
  v_Feetype     := j_Json.Get_String('fee_type');
  If Nvl(n_Outfee, 0) = 0 Then
    Zl_����δ�����_Recalc_s(n_Pati_Id, n_Pati_Pageid, n_Pati_Nature, v_Feetype);
  Else
    Zl_����δ���������_Recalc_s(n_Pati_Id, v_Feetype);
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Recalcfee;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Addcardfeeinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ӿ��Ѽ�Ԥ��������
  --��Σ�Json_In:��ʽ
  --  input
  --    oper_fun            C     ����״̬:0-������Ԥ����򿨷ѽɿ�;1-����Ϊδ��Ч��Ԥ������쳣�Ŀ���;2-����Ϊ���ʵ�;3-����Ϊ���۵�
  --    blnc_money          N  1  ���ν����ܼ�:Ԥ��+����
  --    balance_id          N     ����id
  --    pati_info           C     ������Ϣ
  --      pati_id           C  1  ����ID
  --      pati_pageid       N  1  ��ҳid
  --      pati_name         C  1  ��������
  --      pati_sex          C  1  �Ա�
  --      pati_age          C  1  ����
  --      outpno  N         1     �����
  --      mdlmode_name      C  1  ���ʽ����
  --      fee_category      C  1  �ѱ�
  --      insurance_type    N     ����
  --    card_info           C     ҽ�ƿ���Ϣ
  --      cardno            C  1  ��������
  --      cardtype_id       N  1  ���������ID
  --      send_mode         N  1  ������ʽ;0-����,1-����,2-����
  --      cardno_reusing    N  1  ��������:1-����;0-����������
  --      recv_id           N  1  ����id:����Id
  --      cardno_old        C  1  ԭ������:����ʱ����Ҫ����ԭ����
  --    deposit_info        C  1  Ԥ�����б�
  --      deposit_no        C  1  Ԥ�����ݺ�
  --      deposit_id        N     Ԥ��ID
  --      fact_no           C  1  ��Ʊ��
  --      deposit_type      N     Ԥ�����:1-����;2-סԺ
  --      pati_id           N  1  ����id
  --      pati_pageid       N  1  ��ҳid
  --      dept_id           N  1  �ɿ����id
  --      money             N  1  �ɿ���
  --      emp_name          C  1  �ɿλ
  --      emp_bank_name     C  1  ��λ������
  --      emp_bank_actno    C  1  �������˺�
  --      memo              C  1  ժҪ
  --      recv_id           N  1  Ʊ������id
  --      start_einv        N  1  �Ƿ����õ���Ʊ��:1-����;0-������
  --    cardfee_list[]      C  1  �����б�
  --      fee_no            C  1  ���õ��ݺ�
  --      serial_num        N  1  ���
  --      price_ftrnum      N  1  �۸񸸺�
  --      subde_ftrnum      N  1  ��������
  --      receipt_type      C  1  �շ����
  --      fitem_id          N  1  �շ�ϸĿid
  --      income_item_id    N  1  ������Ŀid
  --      price             N  1  ��׼����
  --      receipt_fee       C  1  �վݷ�Ŀ
  --      fee_amrcvb        N  1  Ӧ�ս��
  --      fee_ampaib        N  1  ʵ�ս��
  --      pati_deptid       N  1  ���˿���id
  --      pati_wardarea_id  N  1  ���˲���id
  --      exedept_id        N  1  ִ�в���id
  --      bill_deptid       N  1  ��������id
  --      mrbkfee_sign      N  1  �Ƿ�����:1-�ǲ�����;0-���ǲ�����
  --      insurance_code    C  1  ���ձ���
  --      insurance_type_id N  1  ���մ���id
  --      insurance_sign    N  1  ������Ŀ��:1-�Ǳ�����Ŀ;0-���Ǳ�����Ŀ
  --      si_manp_money     N  1  ͳ����
  --      memo              C  1  ժҪ
  --      overtime_flag     N  1  �Ӱ��־
  --      cardno            C  1  ��������
  --      cardtype_id       N  1  ���������ID
  --      send_mode         N  1  ������ʽ;0-����,1-����,2-����
  --    balance_info        C     ������Ϣ:Ŀǰֻ֧��һ�ֽ��㷽ʽ
  --      blnc_mode         C  1  ���㷽ʽ
  --      blnc_no           C  1  �������
  --      cardtype_id       C  1  �����id
  --      consumer_no       C  1  ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --      consume_card_id   N  1  ���ѿ�ID
  --      cardno            C  1  ֧������
  --      swapno            C  1  ������ˮ��
  --      swapmemo          C  1  ����˵��
  --      cprtion_unit      C  1  ������λ
  --      start_einv        N  1  �Ƿ����õ���Ʊ��:1-����;0-������
  --    operator_name       C  1  ����Ա����
  --    operator_code       C  1  ����Ա���
  --    create_time         C  1  �Ǽ�ʱ����տ�ʱ��:yyyy-mm-dd hh:mi:ss

  --����: Json_Out,��ʽ����
  --  output
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    cardfee_no  C  1  ���ѵķ��õ��ݺ�
  --    deposit_no  C  1  Ԥ�����ݺ�
  --    deposit_id  N 1 Ԥ��ID
  --    balance_id  N 1 ����ID
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  o_Json     PLJson;
  j_Billlist Pljson_List := Pljson_List();

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  --���ν�����Ϣ
  n_����״̬     Number(2);
  n_����id       ������ü�¼.����id%Type;
  n_���ν����ܼ� ������ü�¼.ʵ�ս��%Type;

  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_ԭ���ݺ�   ������ü�¼.No%Type;

  d_�Ǽ�ʱ�� ������ü�¼.�Ǽ�ʱ��%Type;
  n_�������� Number(2);
  n_�����ܶ� Number(16, 5);
  n_�Ƿ񻮼� Number(2);
  n_����     Number(5);
  n_���ѿ�id Number(18);
  n_�Ʒѷ�ʽ Number(18);
  n_����ֵ   ����Ԥ����¼.��Ԥ��%Type;
  v_ԭ����   Ʊ��ʹ����ϸ.����%Type;
  --������Ϣ��ض���

  n_����id       ������ü�¼.����id%Type;
  v_��������     ������ü�¼.����%Type;
  v_�Ա�         ������ü�¼.�Ա�%Type;
  v_����         ������ü�¼.����%Type;
  n_�����       Number(18);
  n_סԺ��       Number(18);
  v_���ʽ���� ҽ�Ƹ��ʽ.����%Type;
  v_�ѱ�         ������ü�¼.�ѱ�%Type;
  -- n_����         ���ս����¼.����%Type;
  n_������ҳid ����Ԥ����¼.��ҳid%Type;
  --������ض���
  v_���õ���     ������ü�¼.No%Type;
  v_���۵�       ������ü�¼.No%Type;
  n_���         ������ü�¼.���%Type;
  n_�۸񸸺�     ������ü�¼.�۸񸸺�%Type;
  n_��������     ������ü�¼.��������%Type;
  v_�շ����     ������ü�¼.�շ����%Type;
  n_�շ�ϸĿid   ������ü�¼.�շ�ϸĿid%Type;
  n_������Ŀid   ������ü�¼.������Ŀid%Type;
  n_��׼����     ������ü�¼.��׼����%Type;
  v_�վݷ�Ŀ     ������ü�¼.�վݷ�Ŀ%Type;
  n_Ӧ�ս��     ������ü�¼.Ӧ�ս��%Type;
  n_ʵ�ս��     ������ü�¼.ʵ�ս��%Type;
  n_���˿���id   ������ü�¼.���˿���id%Type;
  n_��������id   ������ü�¼.��������id%Type;
  n_�Ƿ�����   Number(2);
  n_�Ƿ����     Number(2);
  n_������ʽ     Number(2);
  v_���ձ���     ������ü�¼.���ձ���%Type;
  n_���մ���id   ������ü�¼.���մ���id%Type;
  n_������Ŀ��   ������ü�¼.������Ŀ��%Type;
  n_ͳ����     ������ü�¼.ͳ����%Type;
  v_����ժҪ     ������ü�¼.ժҪ%Type;
  v_��������     ����Ԥ����¼.����%Type;
  n_���������id ����Ԥ����¼.�����id%Type;
  n_��������id   Ʊ��ʹ����ϸ.����id%Type;

  n_���˲���id ������ü�¼.���˲���id%Type;
  v_���㵥λ   ������ü�¼.���㵥λ%Type;
  n_ִ�в���id ������ü�¼.ִ�в���id%Type;

  n_����״̬ ������ü�¼.����״̬%Type;
  n_У�Ա�־ ����Ԥ����¼.У�Ա�־%Type;
  n_���ӱ�־ Number(2);
  n_�Ӱ��־ סԺ���ü�¼.�Ӱ��־%Type;
  --֧����ʽ����
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_�������   ����Ԥ����¼.�������%Type;
  n_�����id   ����Ԥ����¼.�����id%Type;
  n_���㿨��� ����Ԥ����¼.���㿨���%Type;
  v_֧������   ����Ԥ����¼.����%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;
  v_������λ   ����Ԥ����¼.������λ%Type;
  n_��������id ����Ԥ����¼.��������id%Type;

  --Ԥ����ر�������
  n_Ԥ��id       ����Ԥ����¼.Id%Type;
  v_Ԥ������     ����Ԥ����¼.No%Type;
  v_��Ʊ��       Ʊ��ʹ����ϸ.����%Type;
  n_Ԥ�����     ����Ԥ����¼.Ԥ�����%Type;
  n_��ҳid       ����Ԥ����¼.��ҳid%Type;
  n_�ɿ����id   ����Ԥ����¼.����id%Type;
  n_�ɿ���     ����Ԥ����¼.���%Type;
  v_�ɿλ     ����Ԥ����¼.�ɿλ%Type;
  v_��λ������   ����Ԥ����¼.��λ������%Type;
  v_�������˺�   ����Ԥ����¼.��λ�ʺ�%Type;
  v_ժҪ         ����Ԥ����¼.ժҪ%Type;
  n_����id       Ʊ��ʹ����ϸ.����id%Type;
  n_��id         Number(18);
  n_Count        Number(10);
  n_Ԥ������Ʊ�� ����Ԥ����¼.Ԥ������Ʊ��%Type;
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  n_����Ԥ����� Number(1);
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����״̬     := Nvl(j_Json.Get_Number('oper_fun'), 0);
  n_���ν����ܼ� := j_Json.Get_Number('blnc_total');
  v_����Ա����   := j_Json.Get_String('operator_name');
  v_����Ա���   := j_Json.Get_String('operator_code');
  d_�Ǽ�ʱ��     := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_����id       := j_Json.Get_Number('balance_id');

  n_��id := Zl_Get��id(v_����Ա����);

  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  --1.��ȡ������Ϣ
  o_Json := j_Json.Get_Pljson('pati_info');
  If o_Json Is Null Then
    v_Err_Msg := '�����ڲ�����Ϣ���ݣ�����';
    Raise Err_Item;
  End If;

  n_����id     := o_Json.Get_Number('pati_id');
  n_������ҳid := o_Json.Get_Number('pati_pageid');
  v_��������   := o_Json.Get_String('pati_name');
  v_�Ա�       := o_Json.Get_String('pati_sex');
  v_����       := o_Json.Get_String('pati_age');
  n_�����     := To_Number(o_Json.Get_String('outpatient_num'));
  n_סԺ��     := To_Number(o_Json.Get_String('inpatient_num'));

  v_���ʽ���� := o_Json.Get_String('mdlpay_name');
  v_�ѱ�         := o_Json.Get_String('fee_category');
  --n_����         := o_Json.Get_Number('insurance_type');

  o_Json := PLJson();
  o_Json := j_Json.Get_Pljson('card_info');
  If o_Json Is Null Then
    v_Err_Msg := '�����ڷ�����󶨿�����Ϣ������!';
    Raise Err_Item;
  End If;

  v_��������     := o_Json.Get_String('cardno');
  n_���������id := o_Json.Get_Number('cardtype_id');
  n_������ʽ     := o_Json.Get_Number('send_mode');
  n_��������id   := o_Json.Get_Number('recv_id');
  n_��������     := o_Json.Get_Number('cardno_reusing');
  v_ԭ����       := o_Json.Get_String('cardno_old');
  --2.�������
  j_Billlist := j_Json.Get_Pljson_List('cardfee_list');
  If j_Billlist Is Null Then
    v_Err_Msg := '�����ڿ������漰�ķ�����Ϣ������';
    Raise Err_Item;
  End If;
  If j_Billlist.Count = 0 Then
    v_Err_Msg := '�����ڿ������漰�ķ�����Ϣ������';
    Raise Err_Item;
  End If;

  n_У�Ա�־ := Null;
  n_����״̬ := Null;
  n_�Ƿ񻮼� := 0;

  If Nvl(n_����״̬, 0) = 2 Then
    --2-����Ϊ���ʵ�
    n_����id   := Null;
    n_�Ƿ���� := 1;
  Elsif n_����״̬ = 3 Then
    --3.����Ϊ���۵�
    v_���۵�   := Nextno(13);
    n_�Ƿ񻮼� := 1;
  Elsif n_����״̬ = 1 Then
    --1-����Ϊδ��Ч��Ԥ������쳣�Ŀ���
    n_У�Ա�־ := 1;
    n_����״̬ := 1;

  Else
    --0-������Ԥ����򿨷ѽɿ�
    n_����Ԥ����� := 1;
  End If;

  If Nvl(n_����״̬, 0) <> 2 Then
    --���Ǽ���ʱ�������ڽ���id
    If Nvl(n_����id, 0) = 0 Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    End If;
  End If;

  For J In 1 .. j_Billlist.Count Loop
    o_Json       := PLJson(j_Billlist.Get(J));
    v_���õ���   := o_Json.Get_String('fee_no');
    n_���       := o_Json.Get_Number('serial_num');
    n_�۸񸸺�   := o_Json.Get_Number('price_ftrnum');
    n_��������   := o_Json.Get_Number('subde_ftrnum');
    v_�շ����   := o_Json.Get_String('receipt_type');
    n_�շ�ϸĿid := Nvl(o_Json.Get_Number('fitem_id'), 0);
    n_������Ŀid := Nvl(o_Json.Get_Number('income_item_id'), 0);
    n_��׼����   := Nvl(o_Json.Get_Number('price'), 0);
    v_�վݷ�Ŀ   := o_Json.Get_String('receipt_fee');
    n_Ӧ�ս��   := Nvl(o_Json.Get_Number('fee_amrcvb'), 0);
    n_ʵ�ս��   := Nvl(o_Json.Get_Number('fee_ampaib'), 0);
    n_���˿���id := Nvl(o_Json.Get_Number('pati_deptid'), 0);
    n_���˲���id := Nvl(o_Json.Get_Number('pati_wardarea_id'), 0);
    n_��������id := Nvl(o_Json.Get_Number('bill_deptid'), 0);
    n_ִ�в���id := Nvl(o_Json.Get_Number('exedept_id'), 0);
    n_�Ƿ����� := Nvl(o_Json.Get_Number('mrbkfee_sign'), 0);
    v_���ձ���   := o_Json.Get_String('insurance_code');
    n_���մ���id := o_Json.Get_Number('insurance_type_id');
    n_������Ŀ�� := o_Json.Get_Number('insurance_sign');
    n_ͳ����   := o_Json.Get_Number('si_manp_money');
    v_����ժҪ   := o_Json.Get_String('memo');
    n_�Ӱ��־   := o_Json.Get_Number('overtime_flag');

    If v_���õ��� Is Null Then
      v_Err_Msg := '������ָ���ķ��õ��ݺţ�����';
      Raise Err_Item;
    End If;
    If Nvl(n_���, 0) = 0 Then
      n_��� := 1;
    End If;
    If Nvl(n_�۸񸸺�, 0) = 0 Then
      n_�۸񸸺� := Null;
    End If;
    If Nvl(n_��������, 0) = 0 Then
      n_�������� := Null;
    End If;
    If n_�Ƿ񻮼� = 1 Then
      v_����ժҪ := v_���۵�;
    End If;
    Select Max(���㵥λ) Into v_���㵥λ From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;

    --�Ʒѷ�ʽ_In��0-�շ�;1-����;2-����
    If Nvl(n_�Ƿ񻮼�, 0) = 1 Then
      n_�Ʒѷ�ʽ := 1;
    Elsif Nvl(n_�Ƿ����, 0) = 1 Then
      n_�Ʒѷ�ʽ := 2;
    Else
      n_�Ʒѷ�ʽ := 0;
    End If;
    n_���ӱ�־ := n_������ʽ;

    If Nvl(n_�Ƿ�����, 0) = 1 Then
      n_���ӱ�־ := 8; --������
    End If;

    Zl_���˿���_Insert_s(n_�Ʒѷ�ʽ, v_���õ���, n_����id, n_������ҳid, n_�����, v_�ѱ�, v_��������, v_�Ա�, v_����, v_���ʽ����, n_���˲���id, n_���˿���id,
                     n_�շ�ϸĿid, v_�շ����, v_���㵥λ, n_������Ŀid, v_�վݷ�Ŀ, n_��׼����, n_Ӧ�ս��, n_ʵ�ս��, n_ִ�в���id, n_��������id, v_����Ա���,
                     v_����Ա����, n_�Ӱ��־, d_�Ǽ�ʱ��, n_���������id, v_��������, v_����ժҪ, n_����id, n_���ӱ�־, v_���۵�, n_���, n_����״̬);

    n_�����ܶ� := Nvl(n_�����ܶ�, 0) + Nvl(n_ʵ�ս��, 0);

  End Loop;

  --��Ҫ����ҽ�ƿ���Ʊ��ʹ��
  If n_�������� = 0 Then
    --��Ҫ����Ƿ����Ʊ��ʹ����ϸ��������ڣ��϶��ᷢ������
    If Nvl(n_��������id, 0) = 0 Then
      Select Nvl(Max(����), 0)
      Into n_����
      From Ʊ��ʹ����ϸ A
      Where a.Ʊ�� = 5 And a.���� = v_�������� And Nvl(a.����id, 0) = 0;

    Else
      Select Nvl(Max(����), 0)
      Into n_����
      From Ʊ��ʹ����ϸ A, Ʊ�����ü�¼ B
      Where a.Ʊ�� = 5 And a.���� = v_�������� And a.����id = n_��������id And a.����id = b.Id;
    End If;
    If n_���� <> 0 Then
      v_Err_Msg := '����:' || v_�������� || ' �Ѿ�ʹ�ã������ٽ��з�������,����!';
      Raise Err_Item;
    End If;
  End If;

  --������ʽ;0-����,1-����,2-����
  If Nvl(n_����״̬, 0) <> 1 Then
    --�䶯����_In=1-���� ;2-����;3-���� ;4-�˿�
    n_Count := Case
                 When Nvl(n_������ʽ, 0) = 0 Then
                  1
                 When Nvl(n_������ʽ, 0) = 1 Then
                  3
                 When Nvl(n_������ʽ, 0) = 2 Then
                  2
                 Else
                  4
               End;

    If n_Count = 2 Then
      --��Ҫ��ȡԭʼ��������
      Select Max(NO)
      Into v_ԭ���ݺ�
      From סԺ���ü�¼
      Where ��¼���� = 5 And ����id = n_����id And ʵ��Ʊ�� = v_ԭ���� And To_Number(Nvl(����, '0')) = Nvl(n_���������id, 0) And ���ӱ�־ <> 8;
      If v_ԭ���ݺ� Is Null Then

        v_Err_Msg := 'δ�ҵ�ԭʼ����:' || v_ԭ���� || '�ķ��õ��� �������ٽ��л�������,����!';
        Raise Err_Item;
      End If;

      --���ӱ�־:0-������1-������2-����
      Update סԺ���ü�¼
      Set ���ӱ�־ = Decode(Nvl(���ӱ�־, 0), 8, 8, n_������ʽ)
      Where ��¼���� = 5 And NO = v_ԭ���ݺ�;

      Zl_����ҽ�ƿ�Ʊ��_Update_s(n_Count, v_��������, v_����Ա����, d_�Ǽ�ʱ��, v_ԭ���ݺ�, n_��������id, v_ԭ����, n_��������);
    Else
      Zl_����ҽ�ƿ�Ʊ��_Update_s(n_Count, v_��������, v_����Ա����, d_�Ǽ�ʱ��, v_���õ���, n_��������id, Null, n_��������);
    End If;

  End If;

  If Nvl(n_�Ƿ����, 0) = 1 Then
    Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '","deposit_id":' || Nvl(n_Ԥ��id, 0) || ',"balance_id":' ||
                Nvl(n_����id, 0) || '}}';
    Return;
  End If;

  If Nvl(n_�Ƿ񻮼�, 0) = 1 Then
    Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '","deposit_id":' || Nvl(n_Ԥ��id, 0) || ',"balance_id":' ||
                Nvl(n_����id, 0) || '}}';
    Return;
  End If;

  --3.���������Ϣ
  o_Json := PLJson();
  o_Json := j_Json.Get_Pljson('balance_info');
  If Not o_Json Is Null Then
    v_���㷽ʽ     := o_Json.Get_String('blnc_mode');
    v_�������     := o_Json.Get_String('blnc_no');
    n_�����id     := o_Json.Get_Number('cardtype_id');
    n_���㿨���   := o_Json.Get_Number('consumer_no');
    v_֧������     := o_Json.Get_String('cardno');
    v_������ˮ��   := o_Json.Get_String('swapno');
    v_����˵��     := o_Json.Get_String('swapmemo');
    v_������λ     := o_Json.Get_String('cprtion_unit');
    n_���ѿ�id     := o_Json.Get_Number('consume_card_id');
    n_�Ƿ����Ʊ�� := o_Json.Get_Number('start_einv');

    If Nvl(n_����id, 0) = 0 Then
      v_Err_Msg := '����ID��ȡ����ȷ,����!';
      Raise Err_Item;
    End If;
    If Nvl(n_�����id, 0) = 0 Then
      n_�����id := Null;
    End If;
    If Nvl(n_���㿨���, 0) = 0 Then
      n_���㿨��� := Null;
    End If;
    If n_�Ƿ����Ʊ�� Is Null Then
      n_�Ƿ����Ʊ�� := Zl_Fun_Isstarteinvoice(5, 0);
    End If;

    Update ����Ԥ����¼
    Set ���㷽ʽ = v_���㷽ʽ, У�Ա�־ = n_У�Ա�־, �����id = n_�����id, ���㿨��� = n_���㿨���, ���� = v_֧������, ������ˮ�� = v_������ˮ��, ����˵�� = v_����˵��,
        ������� = v_�������, ժҪ = 'ҽ�ƿ�����', ������Ա = v_����Ա����, ����ʱ�� = d_�Ǽ�ʱ��, �������� = 5, �տ�ʱ�� = d_�Ǽ�ʱ��, ����Ա��� = v_����Ա���,
        ����Ա���� = v_����Ա����, ����id = n_����id, ��ҳid = Decode(n_������ҳid, 0, Null, n_������ҳid), ���� = v_��������, �Ա� = v_�Ա�, ���� = v_����,
        ����� = n_�����, סԺ�� = n_סԺ��, �Ƿ����Ʊ�� = Nvl(n_�Ƿ����Ʊ��, 0)
    Where ����id = Nvl(n_����id, 0) And ���㷽ʽ Is Null
    Returning ID, ��������id, ��Ԥ�� Into n_Ԥ��id, n_��������id, n_����ֵ;

    If Sql%NotFound Then

      --�����������
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      If Nvl(n_��������id, 0) = 0 Then
        n_��������id := n_Ԥ��id;

      End If;
      Insert Into ����Ԥ����¼
        (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �������,
         ժҪ, �ɿ���id, �����id, ����, ���㿨���, ������ˮ��, ����˵��, ������λ, ��������, ��������id, У�Ա�־, ������Ա, ����ʱ��, �Ƿ����Ʊ��)
      Values
        (n_Ԥ��id, v_���õ���, 5, 1, n_����id, Decode(n_������ҳid, 0, Null, n_������ҳid), v_��������, v_�Ա�, v_����, n_�����, n_סԺ��, v_���ʽ����,
         Decode(n_���˿���id, 0, Null, n_���˿���id), v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_�����ܶ�, n_����id, -1 * n_����id, 'ҽ�ƿ�����',
         n_��id, n_�����id, v_֧������, n_���㿨���, v_������ˮ��, v_����˵��, v_������λ, 5, n_��������id, n_У�Ա�־, v_����Ա����, d_�Ǽ�ʱ��,
         Nvl(n_�Ƿ����Ʊ��, 0));
    Elsif Nvl(n_����ֵ, 0) <> Nvl(n_�����ܶ�, 0) Then
      v_Err_Msg := '�������ȷ,����!';
      Raise Err_Item;

    End If;

    If Nvl(n_���㿨���, 0) <> 0 Then
      --���ѿ�
      Zl_���˿������¼_֧��(n_���㿨���, v_֧������, n_���ѿ�id, n_�����ܶ�, n_Ԥ��id, v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��);

    End If;
    If Nvl(n_У�Ա�־, 0) = 0 Then
      --���ʱ,��Ҫ������Ա�ɿ�����
      For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                   From ����Ԥ����¼ A
                   Where a.����id = Nvl(n_����id, 0) And Mod(a.��¼����, 10) <> 1
                   Group By ���㷽ʽ, ����Ա����) Loop

        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
        Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
        End If;
      End Loop;
    End If;
  End If;

  --4.��ȡԤ����Ϣ
  n_Ԥ��id := Null;
  o_Json   := PLJson();
  o_Json   := j_Json.Get_Pljson('deposit_info');
  If Not o_Json Is Null Then
    v_Ԥ������     := o_Json.Get_String('deposit_no');
    v_��Ʊ��       := o_Json.Get_String('fact_no');
    n_Ԥ�����     := Nvl(o_Json.Get_Number('deposit_type'), 2);
    n_��ҳid       := o_Json.Get_Number('pati_pageid');
    n_�ɿ����id   := o_Json.Get_Number('dept_id');
    n_�ɿ���     := o_Json.Get_Number('money');
    v_�ɿλ     := o_Json.Get_String('emp_name');
    v_��λ������   := o_Json.Get_String('emp_bank_name');
    v_�������˺�   := o_Json.Get_String('emp_bank_actno');
    v_ժҪ         := o_Json.Get_String('memo');
    n_����id       := o_Json.Get_Number('recv_id');
    n_Ԥ��id       := o_Json.Get_Number('deposit_id');
    n_Ԥ������Ʊ�� := o_Json.Get_Number('start_einv');

    If Nvl(n_Ԥ��id, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    End If;
    If Nvl(n_��������id, 0) = 0 Then
      n_��������id := n_Ԥ��id;
    End If;
    If n_Ԥ������Ʊ�� Is Null Then
      n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, 0, 1, n_Ԥ�����);
    End If;

    --����״̬_In:0-�������㣬1-����Ϊ�쳣���ݻ�δ��Ч�ĵ��ݣ�2-����쳣����
    Zl_����Ԥ����¼_Insert_s(n_Ԥ��id, v_Ԥ������, v_��Ʊ��, n_����id, n_��ҳid, v_��������, v_�Ա�, v_����, n_�����, n_סԺ��, v_���ʽ����, n_�ɿ����id,
                       n_�ɿ���, v_���㷽ʽ, v_�������, v_�ɿλ, v_��λ������, v_�������˺�, v_ժҪ, v_����Ա���, v_����Ա����, n_����id, n_Ԥ�����, n_�����id,
                       n_���㿨���, v_֧������, v_������ˮ��, v_����˵��, v_������λ, d_�Ǽ�ʱ��, n_����id, Null, Nvl(n_����Ԥ�����, 0), Nvl(n_����״̬, 0),
                       n_��������id, Null, Nvl(n_Ԥ������Ʊ��, 0));
    n_�����ܶ� := Nvl(n_�����ܶ�, 0) + Nvl(n_�ɿ���, 0);

  End If;
  If Nvl(n_�����ܶ�, 0) <> Nvl(n_���ν����ܼ�, 0) Then
    --���δ�����ܶ�������ܶһ��
    v_Err_Msg := '���ý������ȷ,����!';
    Raise Err_Item;
  End If;

  Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '","deposit_id":' || Nvl(n_Ԥ��id, 0) || ',"balance_id":' ||
              Nvl(n_����id, 0) || '}}';
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Addcardfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Delcardfeeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�Կ��Ѽ�Ԥ����������ʴ���
  --��Σ�Json_In:��ʽ
  --  input
  --    oper_fun  N 1 ����״̬:0-������Ԥ����򿨷ѵ��˿��¼;1-����Ϊ�쳣���˿��¼;2-�����쳣����;3-ɾ�������쳣��¼
  --    cardfee_no  C 1 ���Ѷ�Ӧ�ķ��õ��ݺ�
  --    deposit_no  C 1 Ԥ�����ݺ�
  --    cardfee_sign  N 1 �Ƿ��˿���:1-���˿���;0-���˿���
  --    mrbkfee_sign N 1 �Ƿ��˲�����:1-�˲�����;0-���˲�����
  --    operator_name C 1 ����Ա����
  --    operator_code C 1 ����Ա���
  --    del_time  C 1 �˷�ʱ��:yyyy-mm-dd hh:mi:ss
  --    balance_info  C   ֻ����һ������
  --      moeny N 1 �˿���
  --      blnc_mode C 1 ���㷽ʽ
  --      blnc_no C 1 �������
  --      memo  C 1 ժҪ
  --      cardtype_id N 1 �����id
  --      consumer_no N 1 ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --      consume_card_id N 1 ���ѿ�ID
  --      cardno  C 1 ����
  --      swapno  C 1 ������ˮ��
  --      swapmemo  C 1 ����˵��
  --      cprtion_unit  C 1 ������λ
  --      relation_id N 1 ��������ID
  --����: Json_Out,��ʽ����
  -- output
  --  code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  deposit_id  N 1 Ԥ��ID:���س�Ԥ��ID
  --  balance_id N 1 ����ID�����س���ID

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  n_����״̬ Number(2);
  v_���ѵ��� סԺ���ü�¼.No%Type;
  v_No       סԺ���ü�¼.No%Type;
  v_Ԥ������ ����Ԥ����¼.No%Type;
  v_���۵�   סԺ���ü�¼.No%Type;
  n_���ʷ��� סԺ���ü�¼.���ʷ���%Type;

  n_�Ƿ��˿���   Number(2);
  n_�Ƿ��˲����� Number(2);
  v_����Ա����   סԺ���ü�¼.����Ա����%Type;
  v_����Ա���   סԺ���ü�¼.����Ա���%Type;

  n_����id   סԺ���ü�¼.����id%Type;
  d_�˷�ʱ�� Date;

  n_�˿��� סԺ���ü�¼.ʵ�ս��%Type;

  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_�������   ����Ԥ����¼.�������%Type;
  n_�����id   ����Ԥ����¼.�����id%Type;
  n_���㿨��� ����Ԥ����¼.���㿨���%Type;
  v_����       ����Ԥ����¼.����%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;
  v_������λ   ����Ԥ����¼.������λ%Type;
  n_��������id ����Ԥ����¼.��������id%Type;

  n_�˿�ϼ� סԺ���ü�¼.ʵ�ս��%Type;

  n_����id   ����Ԥ����¼.����id%Type;
  n_ԭ����id ����Ԥ����¼.����id%Type;

  n_��id     ����Ԥ����¼.�ɿ���id%Type;
  n_Count    Number(18);
  n_����ֵ   סԺ���ü�¼.ʵ�ս��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  n_��ֵ��� ����Ԥ����¼.��Ԥ��%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_ԭԤ��id ����Ԥ����¼.Id%Type;

  n_�����       ����Ԥ����¼.�����%Type;
  n_סԺ��       ����Ԥ����¼.סԺ��%Type;
  v_���ʽ���� ����Ԥ����¼.���ʽ����%Type;
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  v_Output       Varchar2(32767);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����״̬ := Nvl(j_Json.Get_Number('oper_fun'), 0);

  v_���ѵ���     := j_Json.Get_String('cardfee_no');
  v_Ԥ������     := j_Json.Get_String('deposit_no');
  n_�Ƿ��˿���   := j_Json.Get_Number('cardfee_sign');
  n_�Ƿ��˲����� := j_Json.Get_Number('mrbkfee_sign');
  v_����Ա����   := j_Json.Get_String('operator_name');
  v_����Ա���   := j_Json.Get_String('operator_code');
  d_�˷�ʱ��     := To_Date(j_Json.Get_String('del_time'), 'YYYY-MM-DD hh24:mi:ss');

  If Nvl(n_����״̬, 0) = 3 Then
    --ɾ���쳣��¼
  
    Delete ����Ԥ����¼
    Where ����id In (Select ����id
                   From סԺ���ü�¼
                   Where NO = v_���ѵ��� And ��¼���� = 5 And ��¼״̬ = 1 And Nvl(����״̬, 0) = 1) And Mod(��¼����, 10) <> 1 And
          Nvl(У�Ա�־, 0) = 1;
  
    If Sql%NotFound Then
      v_Err_Msg := '���ݿ����򲢷�ԭ������ɾ�����Ѿ����㣬�������ٽ���ɾ������!';
      Raise Err_Item;
    End If;
  
    Delete סԺ���ü�¼ Where ��¼���� = 5 And ��¼״̬ = 1 And Nvl(����״̬, 0) = 1 And NO = v_���ѵ���;
  
    --ɾ��Ԥ����¼
    If v_Ԥ������ Is Not Null Then
      Delete ����Ԥ����¼ Where NO = v_Ԥ������ And ��¼���� = 1 And Nvl(У�Ա�־, 0) = 1;
      If Sql%NotFound Then
        v_Err_Msg := '���ݿ����򲢷�ԭ������ɾ�����Ѿ����㣬�������ٽ���ɾ������!';
        Raise Err_Item;
      End If;
    End If;
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '�ɹ�');
    zlJsonPutValue(v_Output, 'deposit_id', 0, 1);
    zlJsonPutValue(v_Output, 'balance_id', 0, 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
  
    Return;
  End If;
  n_��id := Zl_Get��id(v_����Ա����);

  If d_�˷�ʱ�� Is Null Then
    d_�˷�ʱ�� := Sysdate;
  End If;

  Select Max(NO), Nvl(Max(���ʷ���), 0), -1 * Sum(ʵ�ս��), Max(ժҪ), Max(����id), Max(����id)
  Into v_No, n_���ʷ���, n_�˿�ϼ�, v_���۵�, n_����id, n_ԭ����id
  From סԺ���ü�¼
  Where NO = v_���ѵ��� And ��¼���� = 5 And ��¼״̬ = 1 And
        (Nvl(n_�Ƿ��˿���, 0) = 1 And Nvl(���ӱ�־, 0) <> 8 Or Nvl(n_�Ƿ��˲�����, 0) = 1 And Nvl(���ӱ�־, 0) = 8);

  If v_No Is Null Then
    v_Err_Msg := '����Ϊ' || v_���ѵ��� || '������,���ܸõ����򲢷�ԭ���������ʻ��˷�,�������ٽ������ʻ��˷Ѵ���!';
    Raise Err_Item;
  End If;

  --1.�������ʷ��ü�¼
  n_����id := Null;
  If n_���ʷ��� = 0 Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  End If;

  Insert Into סԺ���ü�¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ��ʶ��, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����,
     �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��,
     �ɿ���id, ����, ժҪ, ����״̬)
    Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ��ʶ��, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid,
           ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, v_����Ա���,
           v_����Ա����, ����ʱ��, d_�˷�ʱ��, n_����id, Decode(n_����id, Null, Null, -���ʽ��), n_��id, ����, ժҪ,
           Decode(Nvl(n_����״̬, 0), 0, Null, 1)
    From סԺ���ü�¼
    Where NO = v_���ѵ��� And ��¼���� = 5 And ��¼״̬ = 1 And
          (Nvl(n_�Ƿ��˿���, 0) = 1 And Nvl(���ӱ�־, 0) <> 8 Or Nvl(n_�Ƿ��˲�����, 0) = 1 And Nvl(���ӱ�־, 0) = 8);

  Update סԺ���ü�¼
  Set ��¼״̬ = 3
  Where NO = v_���ѵ��� And ��¼���� = 5 And ��¼״̬ = 1 And
        (Nvl(n_�Ƿ��˿���, 0) = 1 And Nvl(���ӱ�־, 0) <> 8 Or Nvl(n_�Ƿ��˲�����, 0) = 1 And Nvl(���ӱ�־, 0) = 8);

  --���������۵���������ۻ�δ�շѣ�ֱ��ɾ��
  If Not v_���۵� Is Null Then
    Select Count(1)
    Into n_Count
    From ������ü�¼
    Where ����id = n_����id And ��¼���� = 1 And NO = v_���۵� And ��¼״̬ = 0;
  
    If n_Count <> 0 Then
      Zl_���ﻮ�ۼ�¼_Delete_s(v_���۵�);
    End If;
  End If;

  If Nvl(n_���ʷ���, 0) = 1 Then
    --���ʵ���Ҫ����������
    For c_�˷� In (Select a.No, ���, a.����id, a.��ҳid, a.�շ�ϸĿid, a.������Ŀid, a.���˲���id, a.��������id, a.ִ�в���id, a.���˿���id,
                        Nvl(a.ʵ�ս��, 0) As ʵ�ս��, Nvl(a.���ʽ��, 0) As ���ʽ��
                 From סԺ���ü�¼ A
                 Where NO = v_���ѵ��� And ��¼���� = 5 And ��¼״̬ = 2 And �Ǽ�ʱ�� = d_�˷�ʱ��) Loop
    
      Update �������
      Set ������� = Nvl(�������, 0) + c_�˷�.ʵ�ս��
      Where ���� = 1 And ����id = c_�˷�.����id And Nvl(����, 2) = Decode(Nvl(c_�˷�.��ҳid, 0), 0, 1, 2)
      Returning ������� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (c_�˷�.����id, 1, Decode(Nvl(c_�˷�.��ҳid, 0), 0, 1, 2), 0, c_�˷�.ʵ�ս��);
        n_����ֵ := c_�˷�.ʵ�ս��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete ������� Where ���� = 1 And ����id = c_�˷�.����id And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����'����δ�����'
      Update ����δ�����
      Set ��� = Nvl(���, 0) + c_�˷�.ʵ�ս��
      Where ����id = Nvl(c_�˷�.����id, 0) And Nvl(��ҳid, 0) = Nvl(c_�˷�.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(c_�˷�.���˲���id, 0) And
            Nvl(���˿���id, 0) = Nvl(c_�˷�.���˿���id, 0) And Nvl(��������id, 0) = Nvl(c_�˷�.��������id, 0) And
            Nvl(ִ�в���id, 0) = Nvl(c_�˷�.ִ�в���id, 0) And ������Ŀid + 0 = Nvl(c_�˷�.������Ŀid, 0) And ��Դ;�� = 3;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (c_�˷�.����id, Decode(c_�˷�.��ҳid, 0, Null, c_�˷�.��ҳid), Decode(c_�˷�.���˲���id, 0, Null, c_�˷�.���˲���id),
           Decode(c_�˷�.���˿���id, 0, Null, c_�˷�.���˿���id), Decode(c_�˷�.��������id, 0, Null, c_�˷�.��������id),
           Decode(c_�˷�.ִ�в���id, 0, Null, c_�˷�.ִ�в���id), c_�˷�.������Ŀid, 3, c_�˷�.ʵ�ս��);
      End If;
    End Loop;
  End If;
  If n_����״̬ = 2 Then
  
    --�����쳣
    If Nvl(n_���ʷ���, 0) <> 1 Then
    
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ժҪ, ���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���,
         ����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, ���㿨���, У�Ա�־, ��������id, ����ʱ��, ������Ա, ����id, ��Ԥ��,
         ��������, �Ƿ����Ʊ��)
      
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ժҪ, -1 * ���, ���㷽ʽ,
               �������, �տ�ʱ��, ����Ա���, ����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, ���㿨���, Null,
               ��������id, ����ʱ��, ������Ա, n_����id, -1 * ��Ԥ��, ��������, �Ƿ����Ʊ��
        From ����Ԥ����¼
        Where ����id = n_ԭ����id;
    
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ����id = n_ԭ����id And Mod(��¼����, 10) <> 1;
    End If;
  
    --�����쳣����
    If v_Ԥ������ Is Not Null Then
      --����_In:0-ɾ���쳣��ֵ���ݣ�1-ɾ���쳣�˿�ݣ�2-ɾ���쳣����˿��
      Zl_����Ԥ���쳣��¼_Delete(v_Ԥ������, 0);
    End If;
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '�ɹ�');
    zlJsonPutValue(v_Output, 'deposit_id', Nvl(n_Ԥ��id, 0), 1);
    zlJsonPutValue(v_Output, 'balance_id', Nvl(n_����id, 0), 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
  
    Return;
  End If;

  If n_����״̬ = 0 And Nvl(n_�Ƿ��˿���, 0) = 1 Then
    --����Ʊ�ݴ���
    Zl_����ҽ�ƿ�Ʊ��_Update_s(4, '', v_����Ա����, d_�˷�ʱ��, v_���ѵ���, Null, Null, Null);
  End If;
  --������Ϣ
  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Not Null Then
  
    n_�˿���   := o_Json.Get_Number('moeny');
    v_���㷽ʽ   := o_Json.Get_String('blnc_mode');
    v_�������   := o_Json.Get_String('blnc_no');
    n_�����id   := o_Json.Get_Number('cardtype_id');
    n_���㿨��� := o_Json.Get_Number('consumer_no');
    v_����       := o_Json.Get_String('cardno');
    v_������ˮ�� := o_Json.Get_String('swapno');
    v_����˵��   := o_Json.Get_String('swapmemo');
    v_������λ   := o_Json.Get_String('cprtion_unit');
    n_��������id := o_Json.Get_Number('relation_id');
    v_����ժҪ   := o_Json.Get_String('memo');
  
    If Nvl(n_�����id, 0) = 0 Then
      n_�����id := Null;
    End If;
    If Nvl(n_���㿨���, 0) = 0 Then
      n_���㿨��� := Null;
    End If;
  
    --1.�������ʷ���
    If n_���ʷ��� = 0 Then
      --�Ǽ��ʷ��ü����쳣״̬������Ҫ������������
      For c_���� In (Select NO, Max(����id) As ����id, Max(��ҳid) As ��ҳid, Max(����) As ����, Max(�Ա�) As �Ա�, Max(����) As ����,
                          Max(���˿���id) As ���˿���id, Sum(���ʽ��) As ���ʽ��
                   From סԺ���ü�¼
                   Where ����id = n_����id
                   Group By NO) Loop
        Select Max(�����), Max(סԺ��), Max(���ʽ����), Max(�Ƿ����Ʊ��)
        Into n_�����, n_סԺ��, v_���ʽ����, n_�Ƿ����Ʊ��
        From ����Ԥ����¼
        Where ����id In (Select ����id From סԺ���ü�¼ Where NO = c_����.No And ��¼���� = 5 And ��¼״̬ In (0, 1, 3));
      
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        If Nvl(n_��������id, 0) = 0 Then
          n_��������id := n_Ԥ��id;
        End If;
      
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ժҪ, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��,
           ����Ա���, ����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, ���㿨���, У�Ա�־, ��������id, ����ʱ��, ������Ա,
           ��������, �Ƿ����Ʊ��)
          Select n_Ԥ��id, c_����.No, '' As ʵ��Ʊ��, 5, 2, c_����.����id, c_����.��ҳid, c_����.����, c_����.�Ա�, c_����.����, n_�����, n_סԺ��,
                 v_���ʽ����, c_����.���˿���id, v_����ժҪ, n_����id, c_����.���ʽ��, v_���㷽ʽ, v_�������, d_�˷�ʱ��, v_����Ա���, v_����Ա����, '' As �ɿλ,
                 '' As ��λ������, '' As ��λ�ʺ�, n_��id, Null, n_�����id, v_����, v_������ˮ��, v_����˵��, v_������λ, n_���㿨���,
                 Decode(Nvl(n_����״̬, 0), 0, Null, 1), n_��������id, d_�˷�ʱ��, v_����Ա����, 5, n_�Ƿ����Ʊ��
          From Dual;
      
        If n_����״̬ = 0 Then
          If Nvl(n_���㿨���, 0) <> 0 Then
            Begin
              Select b.Id
              Into n_ԭԤ��id
              From ����Ԥ����¼ B
              Where b.����id = n_ԭ����id And b.���㿨��� = n_���㿨���;
            Exception
              When Others Then
                n_ԭԤ��id := -1;
            End;
          
            If n_ԭԤ��id = -1 Then
            
              v_Err_Msg := 'û�з���' || v_���㷽ʽ || '��ԭ�������ݣ�';
              Raise Err_Item;
            
            End If;
            Zl_���˿������¼_�˿�(n_���㿨���, v_����, Null, -1 * c_����.���ʽ��, n_ԭԤ��id, n_Ԥ��id, v_����Ա���, v_����Ա����, d_�˷�ʱ��);
          End If;
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + c_����.���ʽ��
          Where ���� = 1 And �տ�Ա = v_����Ա���� And ���㷽ʽ = v_���㷽ʽ
          Returning ��� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, v_���㷽ʽ, 1, c_����.���ʽ��);
            n_����ֵ := c_����.���ʽ��;
          
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
          End If;
        End If;
      End Loop;
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ����id = n_ԭ����id And ��¼���� = 5;
    End If;
    n_Ԥ��id := Null;
  
    --2.����Ԥ������
    If v_Ԥ������ Is Not Null Then
    
      --����Ԥ����¼
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    
      Select Sum(��Ԥ��), Max(Decode(��¼����, 1, ID, 0)), Max(����id)
      Into n_��Ԥ��, n_ԭԤ��id, n_����id
      From ����Ԥ����¼
      Where Mod(��¼����, 10) = 1 And NO = v_Ԥ������;
      If n_��Ԥ�� <> 0 Then
        v_Err_Msg := 'Ԥ�����Ѿ������������ݣ��������ٽ����˿����!';
        Raise Err_Item;
      End If;
    
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ID = n_ԭԤ��id;
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ժҪ, ���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���,
         ����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, ���㿨���, У�Ա�־, ��������id, ����ʱ��, ������Ա, ��������,
         ����id, Ԥ������Ʊ��)
        Select n_Ԥ��id, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, Nvl(v_����ժҪ, ժҪ) As ժҪ, -1 * ���,
               v_���㷽ʽ, v_�������, d_�˷�ʱ��, v_����Ա���, v_����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, n_��id, Ԥ�����, n_�����id, v_����, v_������ˮ��, v_����˵��,
               v_������λ, n_���㿨���, Decode(Nvl(n_����״̬, 0), 0, Null, 1), ��������id, d_�˷�ʱ��, v_����Ա����, ��������,
               Decode(Nvl(n_����id, 0), 0, Null, n_����id) As ����id, Ԥ������Ʊ��
        From ����Ԥ����¼
        Where ID = n_ԭԤ��id;
    
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ID = n_ԭԤ��id;
    
      If Nvl(n_����״̬, 0) = 0 Then
        --ʹ�����ѿ����������Ԥ�����˲�����������
        For c_Ԥ�� In (Select ID, NO, ���, ���㷽ʽ, ����id, Ԥ����� From ����Ԥ����¼ Where ID = n_Ԥ��id) Loop
          --��Ҫ�������
          n_��ֵ��� := Nvl(c_Ԥ��.���, 0);
        
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.���, 0)
          Where ���� = 1 And ����id = Nvl(c_Ԥ��.����id, 0) And Nvl(����, 2) = Nvl(c_Ԥ��.Ԥ�����, 2)
          Returning Ԥ����� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, ����, Ԥ�����, �������)
            Values
              (c_Ԥ��.����id, 1, Nvl(c_Ԥ��.Ԥ�����, 2), c_Ԥ��.���, 0);
            n_����ֵ := c_Ԥ��.���;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
          End If;
        
          --Ԥ���������
          Update Ԥ���������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + c_Ԥ��.���
          Where Ԥ��id = n_ԭԤ��id
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into Ԥ���������
              (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
            Values
              (n_ԭԤ��id, c_Ԥ��.����id, Nvl(c_Ԥ��.Ԥ�����, 2), c_Ԥ��.���);
            n_����ֵ := c_Ԥ��.���;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From Ԥ���������
            Where Ԥ��id = n_ԭԤ��id And Nvl(Ԥ�����, 2) = Nvl(c_Ԥ��.Ԥ�����, 2) And Nvl(Ԥ�����, 0) = 0;
          End If;
        
          --��Ҫ������Ա�ɿ����
        
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + c_Ԥ��.���
          Where ���� = 1 And �տ�Ա = v_����Ա���� And ���㷽ʽ = c_Ԥ��.���㷽ʽ
          Returning ��� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, c_Ԥ��.���㷽ʽ, 1, c_Ԥ��.���);
            n_����ֵ := c_Ԥ��.���;
          
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ And Nvl(���, 0) = 0;
          End If;
        
        End Loop;
      Else
        Select Sum(���) Into n_��ֵ��� From ����Ԥ����¼ Where ID = n_Ԥ��id;
      End If;
    
    End If;
  
    If Nvl(n_�˿�ϼ�, 0) + Nvl(n_��ֵ���, 0) <> Nvl(n_�˿���, 0) Then
      v_Err_Msg := '��ǰ�˿���(' || Trim(To_Char(Nvl(n_�˿���, 0), '9999999999.999')) || ')�뱾�����ʽ��(' ||
                   Trim(To_Char(Nvl(n_�˿�ϼ�, 0) + Nvl(n_��ֵ���, 0), '99999999999.99')) || ')����ȷ�����飡';
      Raise Err_Item;
    End If;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'deposit_id', Nvl(n_Ԥ��id, 0), 1);
  zlJsonPutValue(v_Output, 'balance_id', Nvl(n_����id, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delcardfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Delcardfeecheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --����:����˿��������������Ƿ�Ϸ�
  --���      json
  --input
  --  cardfee_no  C 1 ���ѵ���
  --  deposit_no  C 1 Ԥ�����ݺ�
  --  reretruned  N 1 �Ƿ��쳣����:1-���쳣����;0-���쳣����
  --  delfee_sign N 1 �˷ѱ�־��0-���˿���;1-���˲�����;2-�����Ѽ�����
  --  balance_info  C   �˿ʽ
  --    delmoney  N 1 �����˿���
  --    pay_mode  C 1 ���㷽ʽ
  --    cardtype_id N 1 �����id
  --    consumer_no N 1 ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --    must_allreturn  N 1 �Ƿ�ȫ��:1-����ȫ��;0-��������
  --����      json
  --output
  --  code                      C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  tip_list[]  C 1 ��ʾ�б�:��Ҫ�ǿ��ܴ��ڶ����ʾѯ�ʷ�ʽ���������б�,��ֹʱ������һ����Ϣ
  --    tip_mode  C 1 ���Ʒ�ʽ:1-��ʾѯ��;2-��ֹ
  --    tip_message C 1 ��ʾ��Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json     PLJson;
  v_���ѵ��� סԺ���ü�¼.No%Type;
  -- v_Ԥ�����ݺ�   ����Ԥ����¼.No%Type;
  n_�Ƿ��쳣���� Number(2);
  n_�˷ѱ�־     Number(2);
  n_���˿���     Number(2);
  n_����id       Number(18);
  n_����id       Number(18);
  v_No           סԺ���ü�¼.No%Type;
  n_����״̬     Number(5);

  n_�����˿��� Number(16, 5);
  -- v_���㷽ʽ     Varchar2(100);
  n_�����id   Number(18);
  n_���㿨��� Number(18);
  n_�Ƿ�ȫ��   Number(18);
  n_Count      Number(18);
  n_���ʷ���   Number(2);
  n_��Ԥ��     Number(16, 5);

  v_Err_Msg    Varchar2(1000);
  n_��������id Number(16, 5);
  n_ʵ�ս��   סԺ���ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս��   סԺ���ü�¼.ʵ�ս��%Type;
  n_���ʽ��   סԺ���ü�¼.ʵ�ս��%Type;
  n_��ֵ���   ����Ԥ����¼.���%Type;
  Function Get_Success_Message
  (
    Tip_Mod_In     Integer,
    Tip_Message_In Varchar2
  ) Return Clob Is
  
  Begin
    Return '{"output":{"code":1,"message":"�ɹ�"' || ',"tip_list":[{"tip_mode":' || Nvl(Tip_Mod_In, 0) || ',"tip_Message":"' || Zltools.Zljsonstr(Tip_Message_In) || '"}]}}';
  End Get_Success_Message;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ѵ��� := j_Json.Get_String('cardfee_no');
  --v_Ԥ�����ݺ�   := j_Json.Get_String('deposit_no');
  n_�Ƿ��쳣���� := j_Json.Get_Number('reretruned');
  n_�˷ѱ�־     := Nvl(j_Json.Get_Number('delfee_sign'), 0);

  If v_���ѵ��� Is Null Then
    v_Err_Msg := 'δ������Ч�ķ��õ��ݺ�';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  Select Max(���ʷ���), Max(����״̬), Max(����id)
  Into n_���ʷ���, n_����״̬, n_����id
  From סԺ���ü�¼
  Where NO = v_���ѵ��� And ��¼���� = 5 And ��¼״̬ In (0, 1, 3) And Rownum < 2;

  If Nvl(n_�Ƿ��쳣����, 0) = 1 Then
    If Nvl(n_���ʷ���, 0) <> 1 Then
      --0-���˿���;1-���˲�����;2-�����Ѽ�����
      If Nvl(n_�˷ѱ�־, 0) = 0 Then
        Select Max(����id)
        Into n_����id
        From סԺ���ü�¼
        Where NO = v_���ѵ��� And ��¼���� = 5 And Nvl(���ӱ�־, 0) <> 8 And ��¼״̬ = 2;
      Elsif Nvl(n_�˷ѱ�־, 0) = 1 Then
        Select Max(����id)
        Into n_����id
        From סԺ���ü�¼
        Where NO = v_���ѵ��� And ��¼���� = 5 And Nvl(���ӱ�־, 0) = 8 And ��¼״̬ = 2;
      Else
        Select Max(����id) Into n_����id From סԺ���ü�¼ Where NO = v_���ѵ��� And ��¼���� = 5 And ��¼״̬ = 2;
      End If;
    
      Select Max(1) Into n_Count From ����Ԥ����¼ Where ����id = n_����id And Nvl(У�Ա�־, 0) <> 0;
      If n_Count = 0 Then
        v_Err_Msg := '�õ��ݿ����ѱ��������ϻ�����';
        Json_Out  := zlJsonOut(v_Err_Msg);
        Return;
      End If;
    End If;
    Json_Out := Get_Success_Message(0, '');
    Return;
  
  End If;

  If Nvl(n_����״̬, 0) = 1 Then
    --��ǰΪ�쳣����
    v_Err_Msg := '���ݡ�' || v_���ѵ��� || '��Ϊ�쳣����';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Not Null Then
    n_�����˿��� := o_Json.Get_Number('delmoney');
    --v_���㷽ʽ     := o_Json.Get_String('pay_mode');
    n_�����id   := o_Json.Get_Number('cardtype_id');
    n_���㿨��� := o_Json.Get_Number('consumer_no');
    n_�Ƿ�ȫ��   := Nvl(o_Json.Get_Number('must_allreturn'), 0);
  
    If Nvl(n_�Ƿ�ȫ��, 0) = 1 Then
      --����ȫ�ˣ�����Ҫ������н����Ƿ�ȫ��
      Select Max(��������id), Max(Decode(��¼����, 1, 1, 0))
      Into n_��������id, n_Count
      From ����Ԥ����¼
      Where ����id = n_����id;
    
      --����Ƿ�����
      If Nvl(n_Count, 0) = 1 Then
        Select Sum(��Ԥ��), Sum(Decode(��¼����, 1, 1, 0) * ���)
        Into n_��Ԥ��, n_��ֵ���
        From ����Ԥ����¼
        Where ��������id = n_��������id And Mod(��¼����, 10) = 1;
        If Nvl(n_��Ԥ��, 0) <> 0 Then
          v_Err_Msg := '������ֵ����Ѿ��������ѣ���ǰ�����ֱ���ȫ��';
          Json_Out  := zlJsonOut(v_Err_Msg);
          Return;
        End If;
      End If;
    
      Select Sum(��Ԥ��) + Nvl(n_��ֵ���, 0) Into n_��Ԥ�� From ����Ԥ����¼ Where ����id = n_����id;
    
      If Nvl(n_��Ԥ��, 0) <> Nvl(n_�����˿���, 0) Then
        v_Err_Msg := '���ν���(' || Nvl(n_��Ԥ��, 0) || ')�뵱ǰ�˿���(' || Nvl(n_�����˿���, 0) || ')����������ȫ��';
        Json_Out  := zlJsonOut(v_Err_Msg);
        Return;
      End If;
    
      For c_Ԥ�� In (Select �����id, ���㿨��� From ����Ԥ����¼ Where ����id = n_����id) Loop
        If Nvl(n_�����id, 0) <> Nvl(c_Ԥ��.�����id, 0) And n_�����id <> 0 Then
          v_Err_Msg := '����ʹ�����������������˿�';
          Json_Out  := zlJsonOut(v_Err_Msg);
          Return;
        Elsif Nvl(c_Ԥ��.���㿨���, 0) <> 0 And Nvl(n_���㿨���, 0) <> Nvl(c_Ԥ��.���㿨���, 0) Then
          v_Err_Msg := '����ʹ�����������������˿�';
          Json_Out  := zlJsonOut(v_Err_Msg);
          Return;
        End If;
      End Loop;
    
    End If;
  End If;

  If Nvl(n_���ʷ���, 0) = 1 Then
    --�ѽ��ʵ��ݲ�������
    n_���˿��� := To_Number(Nvl(zl_GetSysParameter('�ѽ��ʵ��ݲ���'), '0'));
    --0-���� 1-��ʾ 2-��ֹ7
    If Nvl(n_���˿���, 0) <> 0 Then
      --�˷ѱ�־��0-���˿���;1-���˲�����;2-�����Ѽ�����
      Select Max(NO), Nvl(Sum(ʵ�ս��), 0), Nvl(Sum(Ӧ�ս��), 0), Nvl(Sum(���ʽ��), 0), Nvl(Max(����id), 0)
      Into v_No, n_ʵ�ս��, n_Ӧ�ս��, n_���ʽ��, n_����id
      From סԺ���ü�¼
      Where Mod(��¼����, 10) = 5 And ���ʷ��� = 1 And NO = v_���ѵ��� And
            ((n_�˷ѱ�־ = 0 And Nvl(���ӱ�־, 0) <> 8) Or (n_�˷ѱ�־ = 1 And Nvl(���ӱ�־, 0) = 8) Or n_�˷ѱ�־ = 2);
    
      If v_No Is Not Null Then
        If (n_ʵ�ս�� - n_���ʽ�� = 0 And n_Ӧ�ս�� <> 0) Or (n_ʵ�ս�� = 0 And n_���ʽ�� = 0 And n_����id <> 0) Then
          --�϶�����
          v_Err_Msg := '���ʵ���' || v_No || '���Ѿ�����';
          Json_Out  := Get_Success_Message(n_���˿���, v_Err_Msg);
          Return;
        End If;
      End If;
    End If;
  Else
    --�˷ѱ�־��0-���˿���;1-���˲�����;2-�����Ѽ�����
    Select Count(1)
    Into n_Count
    From סԺ���ü�¼
    Where NO = v_���ѵ��� And ��¼״̬ = 1 And ��¼���� = 5 And
          ((n_�˷ѱ�־ = 0 And Nvl(���ӱ�־, 0) <> 8) Or (n_�˷ѱ�־ = 1 And Nvl(���ӱ�־, 0) = 8) Or n_�˷ѱ�־ = 2);
    If Nvl(n_Count, 0) = 0 Then
      If n_�˷ѱ�־ = 1 Then
        Json_Out := Get_Success_Message(2, '��ǰ�������ѱ�������Ա�˷�');
      Else
        Json_Out := Get_Success_Message(2, '��ǰ�����ѱ�������Ա�˷�');
      End If;
      Return;
    End If;
  End If;
  Json_Out := Get_Success_Message(0, '');
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Delcardfeecheck;
/


Create Or Replace Procedure Zl_Exsesvr_Checkcardnoisused
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��鿨���Ƿ�ʹ�ã�ʹ�ú󷵻�����ID
  --��Σ�Json_In:��ʽ
  --   input      
  --    cardtype_id  N  1  �����id
  --    cardno  C  1  ����
  --����: Json_Out,��ʽ����
  --  output      
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexsit  N  1  �Ƿ����:1-����;0-������
  --    recv_id  N  1  ����id

  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_�����id Number(18);
  v_����     Varchar2(100);
  v_Output   Varchar2(32767);

  n_����id Ʊ�����ü�¼.Id%Type;

  n_Exist Number(2);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�����id := j_Json.Get_Number('cardtype_id');
  v_����     := j_Json.Get_String('cardno');

  Select Max(b.����id), Max(1)
  Into n_����id, n_Exist
  From Ʊ�����ü�¼ A, Ʊ��ʹ����ϸ B
  Where a.Id = b.����id And a.Ʊ�� = 5 And (Nvl(a.ʹ�����, 'LXH') = To_Char(n_�����id) Or a.ʹ����� Is Null) And b.���� = v_����;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'isexist', Nvl(n_Exist, 0), 1);
  zlJsonPutValue(v_Output, 'isexist', Nvl(n_����id, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkcardnoisused;
/

Create Or Replace Procedure Zl_Exsesvr_Updcardfeeblncinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�����ķѼ�����ͬ����Ԥ���������Ϣ(��Ʊ��Ԥ������Ϣ)
  --��Σ�Json_In:��ʽ
  --input
  --   oper_fun  N  1  ����״̬:0-��ɽ���;1-֧���ӿڵ���ǰ����;2-֧���ӿڵ��ú�����
  --   pati_id N 1 ����id
  --   fee_no  C 1 ���õ��ţ��������漰�ķ��õ���
  --   balance_id  N 1 ����ID
  --   operator_name C 1 ����Ա����
  --   operator_code C 1 ����Ա���
  --   create_time C 1 ����ʱ��:yyyy-mm-dd hh:mi:ss
  --   completioned N 1 ��ɱ�־: 1-��ɽ���;0-δ��ɽ���  ,δ���뱾�ӵ㣬Ĭ��Ϊ��ɽ���
  --   fee_einvoice  N  1  ���ѻ������Ƿ����õ���Ʊ��:1-����;0-������
  --   sendcard_info     ������Ϣ
  --     send_mode N 1 ������ʽ;0-����,1-����,2-����
  --     cardtype_id C 1 �����id
  --     cardno  C 1 ����:���η��Ż�󶨻򲹿��Ŀ���
  --     recv_id N 1 ����id:Ʊ������ID(����)
  --     cardno_reusing  N 1 ��������:1-���������ظ�ʹ����;0-�������ظ�ʹ��
  --     cardno_old  C 1 ԭ������:����ʱ����Ҫ����ԭ����
  --   balance_info  C   ������Ϣ
  --     deposit_no  C   Ԥ������
  --     deposit_id  N   Ԥ��ID
  --     deposit_einvoice      Ԥ�����õ���Ʊ��:1-����;0-������
  --     pay_mode  C 1 ���㷽ʽ
  --     blnc_no C 1 �������
  --     cardtype_id N 1 �����id
  --     consumer_no N 1 ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --     cardno  C 1 ����
  --     swapno  C 1 ������ˮ��
  --     swapmemo  C 1 ����˵��
  --     memo  C 1 ժҪ
  --     cprtion_unit  C 1 ������λ
  --     other_list[]  C 1 ����������Ϣ
  --       swap_name C 1 ��������
  --       swap_note C 1 ��������
  --����: Json_Out,��ʽ����
  -- output
  --   code                  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json     PLJson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_������ʽ   Number(2);
  n_����id     Number(18);
  v_Ԥ������   ����Ԥ����¼.No%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  v_���õ���   ������ü�¼.No%Type;
  n_����id     ������ü�¼.����id%Type;
  n_����id     Ʊ�����ü�¼.Id%Type;
  n_��������   Number(2);
  v_ԭ����     Varchar2(100);
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_�������   ����Ԥ����¼.�������%Type;
  n_�����id   ����Ԥ����¼.�����id%Type;
  n_���㿨��� ����Ԥ����¼.���㿨���%Type;
  v_֧������   ����Ԥ����¼.����%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;
  v_ժҪ       ����Ԥ����¼.ժҪ%Type;
  v_������λ   ����Ԥ����¼.������λ%Type;
  n_��������id ����Ԥ����¼.��������id%Type;
  v_��������   �������㽻��.������Ŀ%Type;
  v_��������   �������㽻��.��������%Type;

  n_Ԥ�����     ����Ԥ����¼.���%Type;
  n_������     ����Ԥ����¼.���%Type;
  v_����         Varchar2(100);
  v_����Ա����   ����Ԥ����¼.����Ա����%Type;
  v_����Ա���   ����Ԥ����¼.����Ա���%Type;
  d_�Ǽ�ʱ��     ����Ԥ����¼.�տ�ʱ��%Type;
  n_�����       ����Ԥ����¼.�����%Type;
  n_סԺ��       ����Ԥ����¼.סԺ��%Type;
  v_���ʽ���� ����Ԥ����¼.���ʽ����%Type;
  n_Count        Number(18);
  n_����ֵ       Number(16, 5);
  n_��id         Number(18);
  v_����ժҪ     ����Ԥ����¼.ժҪ%Type;
  n_����״̬     Number(5);
  v_ԭ���ݺ�     סԺ���ü�¼.No%Type;
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  n_Ԥ������Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  n_Temp         Number(2);
  Err_Item Exception;
  v_Err_Msg Varchar2(500);
  l_Ԥ��id  t_NumList := t_NumList();

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���õ���     := j_Json.Get_String('fee_no');
  n_����id       := j_Json.Get_Number('balance_id');
  v_����Ա����   := j_Json.Get_String('operator_name');
  v_����Ա���   := j_Json.Get_String('operator_code');
  d_�Ǽ�ʱ��     := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_����id       := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_����״̬     := Nvl(j_Json.Get_Number('oper_fun'), 1);
  n_�Ƿ����Ʊ�� := Nvl(j_Json.Get_Number('fee_einvoice'), 0);

  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  If Nvl(n_����id, 0) = 0 Then
  
    v_Err_Msg := '����ȷ��������Ϣ�����飡';
    Raise Err_Item;
  End If;
  If v_���õ��� Is Null Then
    v_Err_Msg := '����ȷ�����õ�����Ϣ�����飡';
    Raise Err_Item;
  End If;

  --��ȡ������Ϣ
  o_Json := j_Json.Get_Pljson('sendcard_info');
  If o_Json Is Not Null Then
    n_������ʽ := Nvl(o_Json.Get_Number('send_mode'), 0); --������ʽ;0-����,1-����,2-���� ,3-�˿�
    n_�����id := o_Json.Get_Number('cardtype_id');
    n_�������� := Nvl(o_Json.Get_Number('cardno_reusing'), 0);
    v_����     := o_Json.Get_String('cardno');
    v_ԭ����   := o_Json.Get_String('cardno_old');
    n_����id   := o_Json.Get_Number('recv_id');
  
    --Ʊ�ݴ���
    If Nvl(n_����״̬, 0) = 0 Then
      --���ʱ����Ҫ����Ʊ��
      If v_���� Is Null Then
        v_Err_Msg := '����ȷ�������Ŀ�����Ϣ�����飡';
        Raise Err_Item;
      End If;
    
      --1-���� ;2-����;3-���� ;4-�˿�
      n_Count := Case
                   When n_������ʽ = 0 Then
                    1
                   When n_������ʽ = 1 Then
                    3
                   When n_������ʽ = 2 Then
                    2
                   Else
                    4
                 End;
    
      If n_������ʽ = 2 Then
        --��Ҫ��ȡԭʼ��������
        Select Max(NO)
        Into v_ԭ���ݺ�
        From סԺ���ü�¼
        Where ��¼���� = 5 And ����id = n_����id And ʵ��Ʊ�� = v_ԭ���� And To_Number(Nvl(����, '0')) = Nvl(n_�����id, 0) And ���ӱ�־ <> 8;
        If v_ԭ���ݺ� Is Null Then
          v_Err_Msg := 'δ�ҵ�ԭʼ����:' || v_ԭ���� || '�ķ��õ��� �������ٽ��л�������,����!';
          Raise Err_Item;
        End If;
        --���ӱ�־:0-������1-������2-����
        Update סԺ���ü�¼
        Set ���ӱ�־ = Decode(Nvl(���ӱ�־, 0), 8, 8, n_������ʽ)
        Where ��¼���� = 5 And NO = v_ԭ���ݺ�;
      Else
        v_ԭ���ݺ� := v_���õ���;
      End If;
    
      Zl_����ҽ�ƿ�Ʊ��_Update_s(n_Count, v_����, v_����Ա����, d_�Ǽ�ʱ��, v_ԭ���ݺ�, n_����id, v_ԭ����, n_��������);
    End If;
  
    --�ǻ�������Ҫ����
    --����ǻ�����v_���õ��Ų�ΪNULLʱ����ʾֻ�ղ����ѣ����Բ�����ʵ��Ʊ�ż�����
    If Nvl(n_����id, 0) = 0 Then
      Update סԺ���ü�¼
      Set ����Ա���� = v_����Ա����, ����Ա��� = v_����Ա���, �Ǽ�ʱ�� = d_�Ǽ�ʱ��, ʵ��Ʊ�� = Nvl(v_����, ʵ��Ʊ��)
      Where Nvl(����״̬, 0) = 1 And NO = v_���õ��� And ��¼���� = 5 And ��¼״̬ In (1, 3);
    Else
      Update סԺ���ü�¼
      Set ����Ա���� = v_����Ա����, ����Ա��� = v_����Ա���, �Ǽ�ʱ�� = d_�Ǽ�ʱ��, ʵ��Ʊ�� = Nvl(v_����, ʵ��Ʊ��)
      Where Nvl(����״̬, 0) = 1 And ����id = Nvl(n_����id, 0);
    End If;
  
    If Sql%NotFound Then
      v_Err_Msg := '�����򲢷�ԭ�����������ջ����ˣ����飡';
      Raise Err_Item;
    End If;
  Elsif Nvl(n_����״̬, 0) = 0 Then
    --��ɱ�־=1ʱ����Ҫ����Ʊ��
    Select Count(1) Into n_Count From סԺ���ü�¼ Where ����id = Nvl(n_����id, 0) And ���ӱ�־ <> 8;
    If n_Count <> 0 Then
      v_Err_Msg := 'δ���뷢��������Ϣ������';
      Raise Err_Item;
    End If;
  End If;

  o_Json := PLJson();

  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Not Null Then
    n_Ԥ������Ʊ�� := Nvl(o_Json.Get_Number('deposit_einvoice'), 0);
    v_Ԥ������     := o_Json.Get_String('deposit_no');
    n_Ԥ��id       := o_Json.Get_Number('deposit_id');
  
    v_���㷽ʽ   := o_Json.Get_String('pay_mode');
    v_�������   := o_Json.Get_String('blnc_no');
    n_�����id   := o_Json.Get_Number('cardtype_id');
    n_���㿨��� := o_Json.Get_Number('consumer_no');
    v_֧������   := o_Json.Get_String('cardno');
    v_������ˮ�� := o_Json.Get_String('swapno');
    v_����˵��   := o_Json.Get_String('swapmemo');
    v_ժҪ       := o_Json.Get_String('memo');
    v_������λ   := o_Json.Get_String('cprtion_unit');
  
    If Nvl(n_�����id, 0) = 0 Then
      n_�����id := Null;
    End If;
    If Nvl(n_���㿨���, 0) = 0 Then
      n_���㿨��� := Null;
    End If;
  
    If v_Ԥ������ Is Not Null Then
    
      Update ����Ԥ����¼
      Set ���㷽ʽ = Nvl(v_���㷽ʽ, ���㷽ʽ), ������� = v_�������, �����id = n_�����id, ���㿨��� = Decode(Nvl(n_���㿨���, 0), 0, Null, ���㿨���),
          ���� = v_֧������, ������ˮ�� = v_������ˮ��, ����˵�� = v_����˵��, ժҪ = Nvl(v_ժҪ, ժҪ), ������λ = Nvl(v_������λ, ������λ), �տ�ʱ�� = d_�Ǽ�ʱ��,
          ����Ա���� = v_����Ա����, ����Ա��� = v_����Ա���, У�Ա�־ = Nvl(n_����״̬, У�Ա�־), Ԥ������Ʊ�� = Decode(��¼����, 5, Null, Nvl(n_Ԥ������Ʊ��, 0))
      Where ID = n_Ԥ��id And (��¼״̬ = 0 Or ��¼״̬ = 2);
    
      If Sql%NotFound Then
        v_Err_Msg := 'δ�ҵ����ݺ�Ϊ' || v_Ԥ������ || '��Ԥ������ ��';
        Raise Err_Item;
      End If;
    End If;
  
    If v_���õ��� Is Not Null Then
    
      Update ����Ԥ����¼
      Set ���㷽ʽ = Nvl(v_���㷽ʽ, ���㷽ʽ), ������� = v_�������, �����id = n_�����id, ���㿨��� = Decode(Nvl(n_���㿨���, 0), 0, Null, ���㿨���),
          ���� = v_֧������, ������ˮ�� = v_������ˮ��, ����˵�� = v_����˵��, ժҪ = Nvl(v_ժҪ, ժҪ), ������λ = Nvl(v_������λ, ������λ),
          У�Ա�־ = Nvl(n_����״̬, У�Ա�־), �տ�ʱ�� = d_�Ǽ�ʱ��, ����Ա���� = v_����Ա����, ����Ա��� = v_����Ա���,
          ��������id = Decode(Nvl(��������id, 0), 0, ID, ��������id), �Ƿ����Ʊ�� = Decode(��¼����, 5, Nvl(n_�Ƿ����Ʊ��, 0), Null)
      Where ����id = n_����id;
    
      If Sql%NotFound Then
      
        Select Max(��������id), Max(ժҪ), Max(�����), Max(סԺ��), Max(���ʽ����), Max(�Ƿ����Ʊ��)
        Into n_��������id, v_����ժҪ, n_�����, n_סԺ��, v_���ʽ����, n_�Ƿ����Ʊ��
        From ����Ԥ����¼
        Where ����id In (Select ����id From סԺ���ü�¼ Where NO = v_���õ��� And ��¼���� = 5 And ��¼״̬ In (0, 1, 3));
      
        n_��id := Zl_Get��id(v_����Ա����);
      
        For c_���� In (Select NO, Max(����id) As ����id, Max(��ҳid) As ��ҳid, Max(����) As ����, Max(�Ա�) As �Ա�, Max(����) As ����,
                            Max(���˿���id) As ���˿���id, Sum(���ʽ��) As ���ʽ��
                     From סԺ���ü�¼
                     Where ����id = n_����id
                     Group By NO) Loop
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, ժҪ, ���, ����id, �������, ��Ԥ��, ���㷽ʽ,
             �������, �տ�ʱ��, ����Ա���, ����Ա����, �ɿλ, ��λ������, ��λ�ʺ�, �ɿ���id, Ԥ�����, �����id, ����, ������ˮ��, ����˵��, ������λ, ���㿨���, У�Ա�־, ��������id,
             ����ʱ��, ������Ա, �Ƿ����Ʊ��)
          
            Select ����Ԥ����¼_Id.Nextval, v_���õ���, '' As ʵ��Ʊ��, 5, 2, c_����.����id, c_����.��ҳid, c_����.����, c_����.�Ա�, c_����.����, n_�����,
                   n_סԺ��, v_���ʽ����, c_����.���˿���id, v_����ժҪ, Null, n_����id, -1 * n_����id, c_����.���ʽ��, v_���㷽ʽ, v_�������, d_�Ǽ�ʱ��,
                   v_����Ա���, v_����Ա����, '' As �ɿλ, '' As ��λ������, '' As ��λ�ʺ�, n_��id, Null,
                   Decode(Nvl(n_�����id, 0), 0, Null, n_�����id), v_����, v_������ˮ��, v_����˵��, v_������λ,
                   Decode(Nvl(n_���㿨���, 0), 0, Null, n_���㿨���), n_����״̬, n_��������id, d_�Ǽ�ʱ��, v_����Ա����, n_�Ƿ����Ʊ��
            From Dual;
        End Loop;
      End If;
    
    End If;
  
    j_Jsonlist := o_Json.Get_Pljson_List('other_list');
    If Not j_Jsonlist Is Null Then
      --��ɾ����������
    
      If Nvl(n_Ԥ��id, 0) <> 0 Then
        l_Ԥ��id.Extend;
        l_Ԥ��id(l_Ԥ��id.Count) := n_Ԥ��id;
      End If;
      For c_Ԥ�� In (Select a.Id
                   From ����Ԥ����¼ A
                   Where a.�����id = Nvl(n_�����id, 0) And a.����id = Nvl(n_����id, 0) And a.����id = Nvl(n_����id, 0) And
                         a.Id <> Nvl(n_Ԥ��id, 0)) Loop
        l_Ԥ��id.Extend;
        l_Ԥ��id(l_Ԥ��id.Count) := c_Ԥ��.Id;
      End Loop;
    
      Forall I In 1 .. l_Ԥ��id.Count
        Delete �������㽻�� Where ����id = l_Ԥ��id(I);
    
      For J In 1 .. j_Jsonlist.Count Loop
        o_Json     := PLJson();
        o_Json     := PLJson(j_Jsonlist.Get(J));
        v_�������� := o_Json.Get_String('swap_name');
        v_�������� := o_Json.Get_String('swap_note');
      
        --�ٲ���
        Forall I In 1 .. l_Ԥ��id.Count
          Insert Into �������㽻��
            (����id, ������Ŀ, ��������, ԭԤ��id, ����)
            Select l_Ԥ��id(I) As Ԥ��id, v_��������, v_��������, -1 * Null, -1 * Null As ���� From Dual;
      End Loop;
    End If;
  
  End If;

  If Nvl(n_����״̬, 0) <> 0 Then
  
    Json_Out := zlJsonOut('�ɹ�', 1);
  
    Return;
  End If;
  -----------------------------------
  --��ɴ���

  Select Sum(Ԥ�����), Sum(���ʽ��)
  Into n_Ԥ�����, n_������
  From (Select Sum(��Ԥ��) As Ԥ�����, 0 As ���ʽ��
         From ����Ԥ����¼
         Where ����id = n_����id And ����id = Nvl(n_����id, 0)
         Union All
         Select 0 As ������, Sum(���ʽ��) From סԺ���ü�¼ Where ����id = n_����id And ����id = Nvl(n_����id, 0));

  If Nvl(n_Ԥ�����, 0) <> Nvl(n_������, 0) Then
    v_Err_Msg := '���ѽ���ϼ�(' || n_Ԥ����� || '����úϼ�(' || n_������ || ')��һ��,���ܼ���������';
    Raise Err_Item;
  End If;

  If Nvl(n_Ԥ��id, 0) <> 0 Then
    --1.Ԥ������
    --����ʱ����ԭԤ������Ʊ��IDΪ׼
    Begin
      Select Nvl(Ԥ������Ʊ��, 0)
      Into n_Count
      From ����Ԥ����¼
      Where ��¼״̬ In (1, 3) And NO = v_Ԥ������ And ��¼���� = 1 And ID + 0 <> n_Ԥ��id;
    Exception
      When Others Then
        n_Count := Null;
    End;
    If n_Count Is Not Null Then
      n_Ԥ������Ʊ�� := n_Count;
    End If;
  
    Update ����Ԥ����¼
    Set ��¼״̬ = Decode(��¼״̬, 0, 1, ��¼״̬), У�Ա�־ = Null, Ԥ������Ʊ�� = n_Ԥ������Ʊ��
    Where ID = n_Ԥ��id And ����id = n_����id And ��¼���� = 1 And (Nvl(��¼״̬, 0) = 0 Or Nvl(��¼״̬, 0) = 2);
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ�Ԥ�����ݺ�(' || v_Ԥ������ || ')�Ľ������ݣ������򲢷�ԭ�������տ�����ϣ����飡';
      Raise Err_Item;
    End If;
  
    For c_Ԥ�� In (Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, a.����id, Max(a.Ԥ�����) As Ԥ�����,
                        Sum(Decode(a.Id, n_Ԥ��id, a.���, 0)) As ���, Max(Decode(a.Id, n_Ԥ��id, b.����, -1)) As ����
                 From ����Ԥ����¼ A, ���㷽ʽ B
                 Where a.���㷽ʽ = b.����(+) And a.��¼���� = 1 And a.No = v_Ԥ������ And ����id = Nvl(n_����id, 0)
                 Group By a.����id) Loop
    
      Update Ԥ���������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.���, 0)
      Where Ԥ��id = c_Ԥ��.Id Return Nvl(Ԥ�����, 0) Into n_����ֵ;
      If Sql%NotFound Then
        Insert Into Ԥ���������
          (Ԥ��id, ����id, Ԥ�����, Ԥ�����)
        Values
          (c_Ԥ��.Id, c_Ԥ��.����id, c_Ԥ��.Ԥ�����, c_Ԥ��.���);
        n_����ֵ := c_Ԥ��.���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete Ԥ��������� Where Ԥ��id = c_Ԥ��.Id And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      If Nvl(c_Ԥ��.����, 1) <> 5 Then
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.���, 0)
        Where ���� = 1 And ����id = c_Ԥ��.����id And Nvl(����, 0) = Nvl(c_Ԥ��.Ԥ�����, 0)
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, ����, Ԥ�����, �������)
          Values
            (c_Ԥ��.����id, 1, Nvl(c_Ԥ��.Ԥ�����, 0), Nvl(c_Ԥ��.���, 0), 0);
          n_����ֵ := Nvl(c_Ԥ��.���, 0);
        End If;
        If Nvl(Nvl(c_Ԥ��.���, 0), 0) = 0 Then
          Delete From �������
          Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      End If;
    End Loop;
  End If;

  Update סԺ���ü�¼
  Set ����״̬ = Null
  Where ����id = Nvl(n_����id, 0) And ����id = n_����id And Nvl(����״̬, 0) = 1;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ����ѷ��õ���(' || v_���õ��� || ')�������򲢷�ԭ�������տ�����ϣ����飡';
    Raise Err_Item;
  End If;

  Select Max(�Ƿ����Ʊ��), Count(*)
  Into n_Temp, n_Count
  From ����Ԥ����¼
  Where ����id In (Select ����id From סԺ���ü�¼ Where NO = v_���õ��� And ��¼���� = 5 And ��¼״̬ In (0, 1, 3)) And
        ����id + 0 <> Nvl(n_����id, 0);

  If Nvl(n_Count, 0) <> 0 Then
    n_�Ƿ����Ʊ�� := n_Temp;
  End If;

  Update ����Ԥ����¼
  Set У�Ա�־ = Null, �Ƿ����Ʊ�� = Decode(��¼����, 5, Nvl(n_�Ƿ����Ʊ��, 0), Null)
  Where ����id = Nvl(n_����id, 0) And Nvl(��¼����, 10) <> 1 And ����id = n_����id;
  If Sql%NotFound Then
    v_Err_Msg := 'δ�ҵ����ѷ���(����Ϊ' || v_���õ��� || ')�Ľ�����Ϣ�������򲢷�ԭ�������տ�����ϣ����飡';
    Raise Err_Item;
  End If;

  --2.������Ա�ɿ����
  For c_Ԥ�� In (Select ���㷽ʽ, ���
               From ����Ԥ����¼
               Where ID = Nvl(n_Ԥ��id, 0) And ��¼���� = 1 And ����id = Nvl(n_����id, 0)
               Union All
               Select ���㷽ʽ, ��Ԥ�� From ����Ԥ����¼ Where ����id = Nvl(n_����id, 0) And ����id = Nvl(n_����id, 0)) Loop
    If Nvl(c_Ԥ��.���, 0) <> 0 Then
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + c_Ԥ��.���
      Where ���� = 1 And �տ�Ա = v_����Ա���� And ���㷽ʽ = c_Ԥ��.���㷽ʽ
      Returning ��� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (v_����Ա����, c_Ԥ��.���㷽ʽ, 1, c_Ԥ��.���);
        n_����ֵ := c_Ԥ��.���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = c_Ԥ��.���㷽ʽ And Nvl(���, 0) = 0;
      End If;
    End If;
  End Loop;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updcardfeeblncinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Geteinvoicesinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ------------------------------------------------------------------------------------------------- 
  --���ܣ���ȡ����Ʊ����Ϣ
  --��Σ�json��ʽ 
  --  input  
  --    query_type  N    ��ѯ��Χ:0-����;1-ֻ��ѯ��Ч�ĵ���Ʊ��;2-��ѯԭʼ����Ʊ����Ϣ
  --    occasion  N  1  ���㳡��:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨,6-����ҽ������
  --    fee_nos  C    query_type=2ʱ��Ч:���ݺ�:���㳡��=2ʱ��ΪԤ��NO, ����idδ���룬�ýڵ�ش�
  --    balance_id  N    ����ID�����㳡��=2ʱ��ΪԤ��ID
  --    read_oldbill  N  1  �Ƿ�ֻ��ȡԭʼ���ݵĵ���Ʊ��:1-��;2-��
  --    invoice_type  N    Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --���Σ�json��ʽ 
  --  output 
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data  C    ����Ʊ����Ϣ����
  --    pati_info  C    ������Ϣ
  --      pati_id  N  1  ����ID
  --      pati_pageid  N    ��ҳID
  --      pati_name  C  1  ����
  --      pati_sex  C  1  �Ա�
  --      pati_age  C  1  ����
  --      outpatient_num  C  1  �����
  --      inpatient_num  C  1  סԺ��
  --    einvoice_info  C    ����Ʊ����Ϣ:query_type=2ʱ����
  --      einv_id  N  1  ����Ʊ��ID
  --      paper_nos  C  1  δ���յ�ֽ�ʷ�Ʊ��Ϣ,����ö��ŷ���
  --    einvoice_list[]  C    ����Ʊ���б�,query_type in (0,1)ʱ����
  --      einv_id  N  1  ����Ʊ��ID
  --      invoice_type  N  1  Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --      rec_state  N  1  ��¼״̬
  --      placeCode  C  1  ��Ʊ�����
  --      inv_total  N  1  ��Ʊ���
  --      inv_oldid  N    ԭƱ��ID
  --      sys_source  C  1  ϵͳ��Դ
  --      demo  C  1  ��ע
  --      einvoice_code  C  1  ����Ʊ�ݴ���
  --      einvoice_no  C  1  ����Ʊ�ݺ���
  --      einvoice_random  C  1  ����У����
  --      voucher_code  C  1  Ԥ����ƾ֤����
  --      voucher_no  C  1  Ԥ����ƾ֤����
  --      voucher_random  C  1  Ԥ����ƾ֤У����
  --      happen_time  C  1  ����Ʊ������ʱ��:yyyymmddhh24miss
  --      picture_url  C  1  ����Ʊ��H5ҳ��URL
  --      picture_neturl  C  1  ����Ʊ������H5ҳ��URL
  --      tran_paper  N  1  �Ƿ񻻿�ֽ�ʷ�Ʊ
  --      trans_paperno  C  1  ������ֽ�ʷ�Ʊ��
  --      trans_printid  N  1  �����Ĵ�ӡid
  --    operator_code  C  1  ����Ա���
  --      operator_name  C  1  ����Ա����
  --    create_time  C  1  �Ǽ�ʱ��:yyyy-mm-dd hh24:mi:ss
  ------------------------------------------------------------------------------------------------- 
  j_Input PLJson;
  j_Json  PLJson;

  n_��ѯ���� Number(2);
  n_����     Number(2);
  --n_���ض�ά��   Number(2);
  n_����id       ������ü�¼.����id%Type;
  v_Nos          Varchar2(32767);
  n_Ʊ��         Number(2);
  v_Temp         Varchar2(32767);
  v_Output       Varchar2(32767);
  c_Output       Clob;
  n_����ԭʼ���� Number(2);

  Cursor c_����Ʊ����Ϣ Is(
    Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
           Max(c.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
    From סԺ���ü�¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
    Where a.No = '-' And a.��¼״̬ In (1, 3) And a.����id = b.����id And a.����id = c.����id(+) And c.Ʊ��(+) = 5 And c.��¼״̬(+) = 1
    Group By c.Id);

  r_����Ʊ����Ϣ c_����Ʊ����Ϣ%RowType;

  Type Ty_Einvoce Is Ref Cursor;
  c_Einvoice Ty_Einvoce; --��̬�α����

  v_Pati Varchar2(32767);

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  --0-����;1-ֻ��ѯ��Ч�ĵ���Ʊ��;2-��ѯԭʼ����Ʊ����Ϣ
  n_��ѯ����     := Nvl(j_Json.Get_Number('query_type'), 0);
  n_����         := Nvl(j_Json.Get_Number('occasion'), 0);
  n_Ʊ��         := j_Json.Get_Number('invoice_type');
  n_����ԭʼ���� := j_Json.Get_Number('read_oldbill');

  --n_���ض�ά�� := Nvl(j_Json.Get_Number('return_qrcode'), 0);

  n_����id := j_Json.Get_Number('balance_id');
  v_Nos    := j_Json.Get_String('fee_nos');

  If Nvl(n_����id, 0) = 0 And v_Nos Is Null Then
    Json_Out := zlJsonOut('δ������Ҫ��ѯ����id�����õ���!');
    Return;
  End If;

  If Nvl(n_��ѯ����, 0) = 2 Then
    --2-��ѯԭʼ����Ʊ����Ϣ
    --n_����:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨,6-����ҽ������
    If n_���� = 1 Or n_���� = 4 Then
      --�շѻ�Һ�
      If Nvl(n_����id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From ������ü�¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No In (Select Distinct NO From ������ü�¼ Where ����id = n_����id And Mod(��¼����, 10) = 1) And a.��¼���� = n_���� And
                a.��¼״̬ In (1, 3) And a.����id = b.����id And a.����id = c.����id(+) And c.Ʊ��(+) = 1 And c.��¼״̬(+) = 1
          
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From ������ü�¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And a.��¼״̬ In (1, 3) And a.��¼���� = n_���� And
                a.����id = b.����id And a.����id = c.����id(+) And c.Ʊ��(+) = 1 And c.��¼״̬(+) = 1
          
          Group By c.Id;
      
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From ������ü�¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No = v_Nos And a.��¼״̬ In (1, 3) And a.��¼���� = n_���� And a.����id = b.����id And a.����id = c.����id(+) And
                c.Ʊ��(+) = 1 And c.��¼״̬(+) = 1
          
          Group By c.Id;
      End If;
    Elsif n_���� = 2 Then
      --Ԥ��
      If Nvl(n_����id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(b.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(b.No) As NO
          From ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where b.No In (Select Distinct NO From ����Ԥ����¼ Where ID = n_����id And ��¼���� = 1) And b.��¼���� = 1 And
                b.��¼״̬ In (1, 3) And b.Id = c.����id(+) And c.Ʊ��(+) = 2 And c.��¼״̬(+) = 1
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(b.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(b.No) As NO
          From ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where b.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And b.��¼״̬ In (1, 3) And b.��¼���� = 1 And
                b.Id = c.����id(+) And c.Ʊ��(+) = 2 And c.��¼״̬(+) = 1
          Group By c.Id;
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(b.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(b.No) As NO
          From ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where b.No = v_Nos And b.��¼״̬ In (1, 3) And b.��¼���� = 1 And b.Id = c.����id(+) And c.Ʊ��(+) = 2 And c.��¼״̬(+) = 1
          Group By c.Id;
      End If;
    
    Elsif n_���� = 5 Then
      --���￨
      If Nvl(n_����id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From סԺ���ü�¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No In (Select Distinct NO From סԺ���ü�¼ Where ����id = n_����id And Mod(��¼����, 10) = 5) And a.��¼���� = 5 And
                a.��¼״̬ In (1, 3) And a.����id = b.����id And a.����id = c.����id(+) And c.Ʊ��(+) = 5 And c.��¼״̬(+) = 1
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From סԺ���ü�¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And a.��¼״̬ In (1, 3) And a.��¼���� = 5 And
                a.����id = b.����id And a.����id = c.����id(+) And c.Ʊ��(+) = 5 And c.��¼״̬(+) = 1
          Group By c.Id;
      
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From סԺ���ü�¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No = v_Nos And a.��¼״̬ In (1, 3) And a.����id = b.����id And a.��¼���� = 5 And a.����id = c.����id(+) And
                c.Ʊ��(+) = 5 And c.��¼״̬(+) = 1
          Group By c.Id;
      End If;
    Elsif n_���� = 6 Then
      --���ò����¼
      If Nvl(n_����id, 0) <> 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, 0 As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From ���ò����¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No In (Select Distinct NO From ���ò����¼ Where ����id = n_����id And Mod(��¼����, 10) = 1) And a.��¼���� = 1 And
                a.��¼״̬ In (1, 3) And a.����id = b.����id And a.����id = c.����id(+) And c.Ʊ��(+) = 1 And c.��¼״̬(+) = 1
          
          Group By c.Id;
      Elsif Instr(v_Nos, ',') > 0 Then
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, 0 As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From ���ò����¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And a.��¼״̬ In (1, 3) And a.��¼���� = 1 And
                a.����id = b.����id And a.����id = c.����id(+) And c.Ʊ��(+) = 1 And c.��¼״̬(+) = 1
          Group By c.Id;
      Else
        Open c_Einvoice For
          Select c.Id, Max(b.�Ƿ����Ʊ��), Max(c.Id) As ����Ʊ��id, Max(c.����) As ����, Max(c.�Ա�) As �Ա�, Max(c.����) As ����,
                 Max(c.����id) As ����id, 0 As ��ҳid, Max(c.�����) As �����, Max(c.סԺ��) As סԺ��, Max(a.No) As NO
          From ���ò����¼ A, ����Ԥ����¼ B, ����Ʊ��ʹ�ü�¼ C
          Where a.No = v_Nos And a.��¼״̬ In (1, 3) And a.��¼���� = 1 And a.����id = b.����id And a.����id = c.����id(+) And
                c.Ʊ��(+) = 1 And c.��¼״̬(+) = 1
          Group By c.Id;
      End If;
    Else
      Json_Out := zlJsonOut('���Ͻڵ㴫��ֵ����!');
      Return;
    End If;
    Fetch c_Einvoice
      Into r_����Ʊ����Ϣ;
    If c_Einvoice %NotFound Then
      Close c_Einvoice;
      If Nvl(n_����id, 0) = 0 Then
        Json_Out := zlJsonOut('δ�ҵ�ԭʼ����(NO=' || v_Nos || ')�ĵ���Ʊ�ݣ�����!');
      Else
        Json_Out := zlJsonOut('δ�ҵ�ԭʼ����(����id=' || n_����id || ')�ĵ���Ʊ�ݣ�����!');
      End If;
      Return;
    End If;
  
    --��������:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨
    --���㳡��:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨,6-����ҽ������
    v_Temp := Null;
    For c_Ʊ�� In (Select Distinct ����
                 From Ʊ��ʹ����ϸ
                 Where Ʊ�� = n_Ʊ�� And ��ӡid In (Select ID As ��ӡid
                                              From (Select b.Id
                                                     From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
                                                     Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And
                                                           b.�������� = Decode(Nvl(n_����, 0), 6, 1, n_����) And b.No = r_����Ʊ����Ϣ.No
                                                     Order By a.ʹ��ʱ�� Desc)
                                              Where Rownum < 2)) Loop
      If v_Temp Is Null Then
        v_Temp := c_Ʊ��.����;
      Else
        v_Temp := v_Temp || ',' || c_Ʊ��.����;
      End If;
    End Loop;
    v_Output := v_Output || '{"pati_id":' || zlJsonStr(r_����Ʊ����Ϣ.����id, 1);
    v_Output := v_Output || ',"pati_pageid":' || zlJsonStr(r_����Ʊ����Ϣ.��ҳid, 1);
    v_Output := v_Output || ',"pati_name":"' || zlJsonStr(r_����Ʊ����Ϣ.����) || '"';
    v_Output := v_Output || ',"pati_sex":"' || zlJsonStr(r_����Ʊ����Ϣ.�Ա�) || '"';
    v_Output := v_Output || ',"pati_age":"' || zlJsonStr(r_����Ʊ����Ϣ.����) || '"';
    v_Output := v_Output || ',"outpatient_num":"' || zlJsonStr(r_����Ʊ����Ϣ.�����) || '"';
    v_Output := v_Output || ',"inpatient_num":"' || zlJsonStr(r_����Ʊ����Ϣ.סԺ��) || '"';
    v_Output := v_Output || '}';
  
    v_Output := '"pati_info":' || v_Output;
    --����Ʊ����Ϣ
    v_Output := v_Output || ',"einvoice_info":';
    v_Output := v_Output || '{"einv_id":' || zlJsonStr(r_����Ʊ����Ϣ.Id, 1);
    v_Output := v_Output || ',"paper_nos":"' || zlJsonStr(v_Temp) || '"}';
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":{' || v_Output || '}}}';
    Return;
  End If;

  --����Ʊ����Ϣ
  v_Output := Null;
  v_Pati   := Null;
  For c_����Ʊ�� In (
                 
                 Select ID, Ʊ��, ��¼״̬, ����id, ����id, 0 As ��ҳid, ����, �Ա�, ����, �����, סԺ��, ����, ����, ������, ƾ֤����, ƾ֤����, ƾ֤������, Ʊ�ݽ��,
                         ����ʱ��, ԭƱ��id, �Ƿ񻻿�, ֽ�ʷ�Ʊ��, ��ӡid, ��ע, ����Ա���, ����Ա����, �Ǽ�ʱ��, ��Ʊ��, ϵͳ��Դ, Url����, Url����
                 From ����Ʊ��ʹ�ü�¼
                 Where ����id = n_����id And Ʊ�� = n_Ʊ�� And ((Nvl(n_����ԭʼ����, 0) = 1 And �˿�id Is Null) Or Nvl(n_����ԭʼ����, 0) = 0) And
                       ((Nvl(n_��ѯ����, 0) = 1 And ��¼״̬ = 1) Or Nvl(n_��ѯ����, 0) = 0)
                 Order By �Ǽ�ʱ�� Desc)
  
   Loop
  
    If v_Pati Is Null Then
      v_Pati := v_Pati || '{"pati_id":' || zlJsonStr(c_����Ʊ��.����id, 1);
      v_Pati := v_Pati || ',"pati_pageid":' || zlJsonStr(c_����Ʊ��.��ҳid, 1);
      v_Pati := v_Pati || ',"pati_name":"' || zlJsonStr(c_����Ʊ��.����) || '"';
      v_Pati := v_Pati || ',"pati_sex":"' || zlJsonStr(c_����Ʊ��.�Ա�) || '"';
      v_Pati := v_Pati || ',"pati_age":"' || zlJsonStr(c_����Ʊ��.����) || '"';
      v_Pati := v_Pati || ',"outpatient_num":"' || zlJsonStr(c_����Ʊ��.�����) || '"';
      v_Pati := v_Pati || ',"inpatient_num":":' || zlJsonStr(c_����Ʊ��.סԺ��) || '"';
      v_Pati := v_Pati || '}';
    End If;
  
    If v_Output Is Not Null Then
      v_Output := v_Output || ',';
    End If;
    v_Output := v_Output || '{"einv_id":' || zlJsonStr(c_����Ʊ��.Id, 1);
    v_Output := v_Output || ',"invoice_type":' || zlJsonStr(c_����Ʊ��.Ʊ��, 1);
    v_Output := v_Output || ',"rec_state":' || zlJsonStr(c_����Ʊ��.��¼״̬, 1);
    v_Output := v_Output || ',"placeCode":"' || zlJsonStr(c_����Ʊ��.��Ʊ��) || '"';
    v_Output := v_Output || ',"inv_total":' || zlJsonStr(c_����Ʊ��.Ʊ�ݽ��, 1);
    v_Output := v_Output || ',"inv_oldid":' || zlJsonStr(c_����Ʊ��.ԭƱ��id, 1);
    v_Output := v_Output || ',"sys_source":"' || zlJsonStr(c_����Ʊ��.ϵͳ��Դ) || '"';
    v_Output := v_Output || ',"demo":"' || zlJsonStr(c_����Ʊ��.��ע) || '"';
    v_Output := v_Output || ',"einvoice_code":"' || zlJsonStr(c_����Ʊ��.����) || '"';
    v_Output := v_Output || ',"einvoice_no":"' || zlJsonStr(c_����Ʊ��.����) || '"';
    v_Output := v_Output || ',"einvoice_random":"' || zlJsonStr(c_����Ʊ��.������) || '"';
    v_Output := v_Output || ',"voucher_code":"' || zlJsonStr(c_����Ʊ��.ƾ֤����) || '"';
    v_Output := v_Output || ',"voucher_no":"' || zlJsonStr(c_����Ʊ��.ƾ֤����) || '"';
    v_Output := v_Output || ',"voucher_random":"' || zlJsonStr(c_����Ʊ��.ƾ֤������) || '"';
    v_Output := v_Output || ',"happen_time":"' || zlJsonStr(c_����Ʊ��.����ʱ��) || '"';
    v_Output := v_Output || ',"picture_url":"' || zlJsonStr(c_����Ʊ��.Url����) || '"';
    v_Output := v_Output || ',"picture_neturl":"' || zlJsonStr(c_����Ʊ��.Url����) || '"';
  
    v_Output := v_Output || ',"tran_paper":' || zlJsonStr(Nvl(c_����Ʊ��.�Ƿ񻻿�, 0), 1);
    v_Output := v_Output || ',"trans_paperno":"' || zlJsonStr(c_����Ʊ��.ֽ�ʷ�Ʊ��) || '"';
    v_Output := v_Output || ',"trans_printid":' || zlJsonStr(Nvl(c_����Ʊ��.��ӡid, 0), 1);
    v_Output := v_Output || ',"operator_code":"' || zlJsonStr(c_����Ʊ��.����Ա���) || '"';
    v_Output := v_Output || ',"operator_name":"' || zlJsonStr(c_����Ʊ��.����Ա����) || '"';
    v_Output := v_Output || ',"create_time":"' || To_Char(c_����Ʊ��.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
  
    v_Output := v_Output || '}';
    If Length(v_Output) > 30000 Then
      If c_Output Is Null Then
        c_Output := Substr(v_Output, 2);
      Else
        c_Output := c_Output || v_Output;
      End If;
      v_Output := Null;
    End If;
  End Loop;

  If v_Pati Is Null Then
    v_Pati := v_Pati || '{"pati_id":' || zlJsonStr(0, 1);
    v_Pati := v_Pati || ',"pati_pageid":' || zlJsonStr(0, 1);
    v_Pati := v_Pati || ',"pati_name":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"pati_sex":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"pati_age":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"outpatient_num":"' || zlJsonStr('') || '"';
    v_Pati := v_Pati || ',"inpatient_num":":' || zlJsonStr('') || '"';
    v_Pati := v_Pati || '}';
  End If;
  v_Pati := '"pati_info":' || v_Pati;
  If Not c_Output Is Null And Not v_Output Is Null Then
  
    c_Output := To_Clob(',"einvoice_list":[') || c_Output || ',' || To_Clob(v_Output || ']');
    c_Output := To_Clob(v_Pati) || c_Output;
    v_Output := '';
  Elsif Not c_Output Is Null And v_Output Is Null Then
    c_Output := To_Clob(',"einvoice_list":[') || c_Output || To_Clob(']');
    c_Output := To_Clob(v_Pati) || c_Output;
    v_Output := '';
  Else
    If Length(v_Pati || ',"einvoice_list":[' || v_Output || ']') <= 30000 Then
      v_Output := v_Pati || ',"einvoice_list":[' || v_Output || ']';
    Else
      c_Output := To_Clob(',"einvoice_list":[') || To_Clob(v_Output || ']');
      c_Output := To_Clob(v_Pati) || c_Output;
      v_Output := '';
    End If;
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","data":{') || c_Output || To_Clob('}}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":{' || v_Output || '}}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Geteinvoicesinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Checkpativisitstate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����²��˾���״̬���
  --��Σ�Json_In:��ʽ
  --input
  --  reg_no             C  1 �Һŵ�
  --  exe_status         N  1 ִ��״̬ 0:���Ϊ����.-1:���Ϊ������

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    msg_mode                        N 1 �����Ϣ��ʾģʽ 0-��ֹ 1-ѯ��
  --    msg_text                        C 1 �����ʾ����
  -------------------------------------------
  v_�Һŵ�   Varchar2(50);
  n_ִ��״̬ Number;

  v_ժҪ     Varchar2(4000);
  v_����no   Varchar2(4000);
  n_Count    Number;
  n_��ʾģʽ Number;
  v_��ʾ���� Varchar2(4000);
  v_Output   Varchar2(1000);

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�   := j_Json.Get_String('reg_no');
  n_ִ��״̬ := Nvl(j_Json.Get_Number('exe_status'), 0);

  --��ȡ�ҺŻ��۵���Ϣ
  Select Max(ժҪ) Into v_ժҪ From ������ü�¼ Where NO = v_�Һŵ� And ��¼���� = 4 And ��¼״̬ = 1 And Rownum < 2;

  If v_ժҪ Is Not Null And Nvl(Instr(v_ժҪ || '', '����:'), 0) <> 0 Then
    --��ȡ�ҺŻ��۵���Ϣ,�жϹҺŻ��۵��Ƿ���ڣ������ڣ�����������״̬����Ϊ����
    v_����no := Substr(v_ժҪ, Length('����:') + 1);
    Select Count(1) Into n_Count From ������ü�¼ Where NO = v_����no And Mod(��¼����, 10) = 1 And ��¼״̬ = 0;
    If n_Count < 1 Then
      If n_ִ��״̬ = 0 Then
        n_��ʾģʽ := 0;
        v_��ʾ���� := '�ùҺŵ��Ļ��۷��ò����ڣ����˺ź����¹Һ�!';
      End If;
    Else
      If n_ִ��״̬ = -1 Then
        n_��ʾģʽ := 1;
        v_��ʾ���� := '�ò��˴��ڹҺŵ��Ļ��۷��ã�����Ϊ������ʱ��ɾ���ùҺŵ��Ļ��۷��ã�' || Chr(13) || Chr(10) || '���Ҳ����ٻָ�Ϊ����,�Ƿ����?';
      End If;
    End If;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'msg_mode', Nvl(n_��ʾģʽ, 0), 1);
  zlJsonPutValue(v_Output, 'msg_text', v_��ʾ����, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkpativisitstate;
/


Create Or Replace Procedure Zl_Exsesvr_Outpatiforcereceive
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ�ǿ���������
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id            N  1 ����id
  --  reg_no             C  1 �Һ�no
  --  exe_deptid         N  1 ִ�в���id
  --  outp_room_name     C  1 ��������
  --  emg_sign           N  1 �����־
  --  operator_name      C  1 ����Ա����
  --  operator_code      C  1 ����Ա���
  --  operator_id        N  1 ����Աid
  --  rgst_appt_sign     N  1 ԤԼ��־
  --  recv_time          C  1 ����ʱ��
  --  outpno             N  1 �����
  --  reg_id             N  1 �Һ�id

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------

  v_�Һŵ� Varchar2(50);

  v_����Ա���� Varchar2(50);
  v_����Ա��� Varchar2(50);
  n_����Աid   Number;
  n_ִ�в���id Number;

  n_ԤԼ��־ Number;
  n_�����־ Number;

  v_�������� Varchar2(50);
  d_����ʱ�� Date;
  n_����id   Number;
  n_�����   Number;
  n_�Һ�id   Number;

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');
  n_����Աid   := j_Json.Get_Number('operator_id');

  n_ִ�в���id := j_Json.Get_Number('exe_deptid');

  n_ԤԼ��־ := Nvl(j_Json.Get_Number('rgst_appt_sign'), 0);
  n_�����־ := Nvl(j_Json.Get_Number('emg_sign'), 0);
  v_�Һŵ�   := j_Json.Get_String('reg_no');
  v_�������� := j_Json.Get_String('outp_room_name');
  d_����ʱ�� := To_Date(j_Json.Get_String('recv_time'), 'YYYY-MM-DD HH24:MI:SS');
  n_����id   := j_Json.Get_Number('pati_id');

  If d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;
  If n_ԤԼ��־ = 1 Then
    n_����� := j_Json.Get_Number('outpno');
    n_�Һ�id := j_Json.Get_Number('reg_id');
    Zl_����ԤԼ�Һ�_����_s(v_�Һŵ�, v_��������, Null, Null, Null, Null, Null, d_����ʱ��, Null, n_����id, n_�����, n_�Һ�id);
  End If;

  Zl_����䶯��¼_Insert(v_�Һŵ�, 3, 'ǿ������', v_����Ա����, v_����Ա���, Null, n_ִ�в���id, Null, n_����Աid, v_����Ա����);

  Zl_���˽���_s(n_����id, v_�Һŵ�, n_ִ�в���id, v_����Ա����, v_��������, n_�����־, Null, d_����ʱ��);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpatiforcereceive;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatitriagemode
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡ�Һż�¼�ķ��﷽ʽ
  --��Σ�Json_In:��ʽ
  --input 
  --  reg_no  C  1 �Һŵ�

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    triage_mode                     N 1 ���﷽ʽ
  -------------------------------------------

  v_�Һŵ�   Varchar2(50);
  n_�Һ�ģʽ Number(3);
  n_���﷽ʽ Number(3);
  v_Output   Varchar2(1000);
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --�������
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  v_�Һŵ�   := j_Json.Get_String('reg_no');
  n_�Һ�ģʽ := To_Number(Nvl(Substr(zl_GetSysParameter(256), 1, 1), 0));

  Begin
    If Nvl(n_�Һ�ģʽ, 0) = 0 Then
      Select Nvl(Max(a.���﷽ʽ), 0)
      Into n_���﷽ʽ
      From �ҺŰ��� A, ���˹Һż�¼ B
      Where a.���� = b.�ű� And b.No = v_�Һŵ�;
    Else
      Select Nvl(Max(a.���﷽ʽ), 0)
      Into n_���﷽ʽ
      From �ٴ������¼ A, ���˹Һż�¼ B
      Where a.Id = b.�����¼id And b.No = v_�Һŵ�;
    End If;
  Exception
    When Others Then
      Null;
  End;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'triage_mode', n_���﷽ʽ, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatitriagemode;
/


Create Or Replace Procedure Zl_Exsesvr_Getrgsapptpatilist
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --���ܣ�����������ȡԤԼ�����б�
  --��Σ�Json_In:��ʽ
  --input
  --  operator_name          C    1 ����Ա����
  --  outp_recv_dept_id      C    1 ����������ID
  --  outp_recv_Range        N    1 ������ﷶΧ 1-�ұ��˺� 2-�����ң�ԤԼ�Һŵķ�ҩ���������Ԥ�ţ�û�����ң���������
  --  emg_sign               N    0 �����־
  --  err_sign               N    0 �쳣��־

  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_list[]           �����б�֧�ֶ����[����]
  --       reg_id           N   1 �Һ�ID
  --       reg_no           C   1 �Һŵ�
  --       pati_id          N   1 ����id
  --       outpatient_num   C   1 �����
  --       pati_name        C   1 ����
  --       pati_sex         C   1 �Ա�
  --       pati_age         C   1 ����
  --       emg_sign         N   1 ����
  --       happen_time      C   1 ����ʱ��
  --       exe_deptid       N   1 ִ�п���ID
  --       exetr            C   1 ִ����
  --       outp_rfrl_status N   1 ת��״̬
  --       record_sign      N   1 ��¼��־
  --       outptyp_name     C   1 ����
  --       pait_dept        C   1 ���˿���
  --       exe_status       N   1 ִ��״̬
  j_Input Pljson;
  j_Json  Pljson;

  v_����Ա���� Varchar(50);
  n_�������id Number(18);
  n_���ﷶΧ   Number(5);
  n_�����־   Number(5);
  n_�쳣��־   Number(5);

  v_Para     Varchar(50);
  n_�ҺŰ��� Number;
  d_����ʱ�� Date;
  v_Jtmp     Varchar2(32767);
  c_Jtmp     Clob;
Begin
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����Ա���� := j_Json.Get_String('operator_name');
  n_�������id := Nvl(j_Json.Get_Number('outp_recv_dept_id'), 0);
  n_���ﷶΧ   := Nvl(j_Json.Get_Number('outp_recv_Range'), 0);
  n_�����־   := Nvl(j_Json.Get_Number('emg_sign'), 0);
  n_�쳣��־   := Nvl(j_Json.Get_Number('err_sign'), 0);
  v_Para       := zl_GetSysParameter(256);
  If Nvl(Zl_To_Number(Substr(v_Para, 1, 1)), 0) <> 0 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
    If Sysdate >= d_����ʱ�� Then
      n_�ҺŰ��� := 1;
    End If;
  End If;

  If n_�쳣��־ = 1 Then
    For c_�����б� In (Select b.Id, b.No, b.����id, b.�����, b.����, b.�Ա�, b.����, b.����, b.����, b.����, b.����ʱ�� As ʱ��, b.����, e.���� As ���˿���,
                          b.����, b.����, b.����ʱ��, b.����ʱ��, b.ִ�в���id, b.ִ����, b.ת��״̬, f.���� As ת�����, b.ת������, b.ת��ҽ��, b.ִ��״̬,
                          b.��¼��־
                   From ���˹Һż�¼ B, �ٴ������¼ C, ���ű� E, ���ű� F
                   Where b.����id Is Not Null And b.�����¼id = c.Id And b.ִ�в���id = e.Id And b.ת�����id = f.Id(+) And
                         b.��¼���� = 1 And b.��¼״̬ = 1 And ((n_�����־ = 1 And b.���� = 1) Or n_�����־ = 0) And Nvl(b.��¼��־, 0) = -1 And
                         b.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And
                         Sysdate Between c.��ʼʱ�� And c.��ֹʱ�� And
                         (n_���ﷶΧ = 0 Or (n_���ﷶΧ = 1 And b.ִ���� || '' = v_����Ա���� || '') Or
                         ((n_���ﷶΧ = 2 or n_���ﷶΧ = 3) And b.ִ�в���id + 0 = n_�������id And (b.ִ���� || '' = v_����Ա���� Or b.ִ���� Is Null)))
                   Order By ����ʱ��) Loop
    
      v_Jtmp := v_Jtmp || ',{"reg_id":' || c_�����б�.Id;
      v_Jtmp := v_Jtmp || ',"reg_no":"' || c_�����б�.No || '"';
      v_Jtmp := v_Jtmp || ',"pati_id":' || c_�����б�.����id;
      v_Jtmp := v_Jtmp || ',"outpatient_num":"' || c_�����б�.����� || '"';
      v_Jtmp := v_Jtmp || ',"pati_name":"' || Zljsonstr(c_�����б�.����) || '"';
      v_Jtmp := v_Jtmp || ',"pati_sex":"' || c_�����б�.�Ա� || '"';
      v_Jtmp := v_Jtmp || ',"pati_age":"' || c_�����б�.���� || '"';
      v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_�����б�.����, 0);
      v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_�����б�.����ʱ��, 'YYYY-MM-DD HH24:MI') || '"';
      v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_�����б�.ִ�в���id || '', 'null');
      v_Jtmp := v_Jtmp || ',"exetr":"' || Zljsonstr(c_�����б�.ִ����) || '"';
      v_Jtmp := v_Jtmp || ',"outp_rfrl_status":' || Nvl(c_�����б�.ת��״̬ || '', 'null');
      v_Jtmp := v_Jtmp || ',"record_sign":' || Nvl(c_�����б�.��¼��־ || '', 'null');
      v_Jtmp := v_Jtmp || ',"outptyp_name":"' || Zljsonstr(c_�����б�.����) || '"';
      v_Jtmp := v_Jtmp || ',"pait_dept":"' || Zljsonstr(c_�����б�.���˿���) || '"';
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_�����б�.ִ��״̬ || '', 'null');
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
  Else
    If n_�ҺŰ��� = 1 Then
      For c_�����б� In (Select b.Id, b.No, b.����id, b.�����, b.����, b.�Ա�, b.����, b.����, b.����, b.����, b.����ʱ�� As ʱ��, b.����,
                            e.���� As ���˿���, b.����, b.����, b.����ʱ��, b.����ʱ��, b.ִ�в���id, b.ִ����, b.ת��״̬, f.���� As ת�����, b.ת������,
                            b.ת��ҽ��, b.ִ��״̬, b.��¼��־
                     From ���˹Һż�¼ B, �ٴ������¼ C, ���ű� E, ���ű� F
                     Where b.����id Is Not Null And b.�����¼id = c.Id And b.ִ�в���id = e.Id And b.ת�����id = f.Id(+) And
                           b.��¼���� = 2 And b.��¼״̬ = 1 And ((n_�����־ = 1 And b.���� = 1) Or n_�����־ = 0) And
                           Nvl(b.��¼��־, 0) <> -1 And b.����ʱ�� Between Trunc(Sysdate) And
                           Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 And Sysdate Between c.��ʼʱ�� And c.��ֹʱ�� And
                           (n_���ﷶΧ = 0 Or (n_���ﷶΧ = 1 And b.ִ���� || '' = v_����Ա���� || '') Or
                           ((n_���ﷶΧ = 2 or n_���ﷶΧ = 3) And b.ִ�в���id + 0 = n_�������id And (b.ִ���� || '' = v_����Ա���� Or b.ִ���� Is Null)))
                     Order By ����ʱ��) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || c_�����б�.Id;
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_�����б�.No || '"';
        v_Jtmp := v_Jtmp || ',"pati_id":' || c_�����б�.����id;
        v_Jtmp := v_Jtmp || ',"outpatient_num":"' || c_�����б�.����� || '"';
        v_Jtmp := v_Jtmp || ',"pati_name":"' || Zljsonstr(c_�����б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"pati_sex":"' || c_�����б�.�Ա� || '"';
        v_Jtmp := v_Jtmp || ',"pati_age":"' || c_�����б�.���� || '"';
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_�����б�.����, 0);
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_�����б�.����ʱ��, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_�����б�.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exetr":"' || Zljsonstr(c_�����б�.ִ����) || '"';
        v_Jtmp := v_Jtmp || ',"outp_rfrl_status":' || Nvl(c_�����б�.ת��״̬ || '', 'null');
        v_Jtmp := v_Jtmp || ',"record_sign":' || Nvl(c_�����б�.��¼��־ || '', 'null');
        v_Jtmp := v_Jtmp || ',"outptyp_name":"' || Zljsonstr(c_�����б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"pait_dept":"' || Zljsonstr(c_�����б�.���˿���) || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_�����б�.ִ��״̬ || '', 'null');
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
    Else
      For c_�����б� In (Select Null As ID, a.No, a.����id, a.��ʶ�� As �����, a.����, a.�Ա�, a.����, a.�Ƿ��� As ����, a.ִ����, b.����,
                            d.���� As ���˿���, a.����ʱ�� As ʱ��, a.����ʱ��, a.ִ�в���id, 0 As ִ��״̬, 0 As ��¼��־, Null As ת��״̬
                     From ������ü�¼ A, �ҺŰ��� B, ���ű� D
                     Where a.���㵥λ = b.���� And a.ִ�в���id = d.Id And a.��� = 1 And a.��¼���� = 4 And a.��¼״̬ = 0 And
                           ((n_�����־ = 1 And a.�Ƿ��� = 1) Or n_�����־ = 0) And
                           Decode(To_Char(Sysdate, 'D'), '1', b.����, '2', b.��һ, '3', b.�ܶ�, '4', b.����, '5', b.����, '6', b.����,
                                  '7', b.����, Null) In
                           (Select ʱ���
                            From ʱ���
                            Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between
                                  Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                          '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                                  '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')) Or
                                  ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') Between
                                  '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                                  Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                          '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')))) And
                           (n_���ﷶΧ = 0 Or (n_���ﷶΧ = 1 And a.ִ���� || '' = v_����Ա���� || '') Or
                           ((n_���ﷶΧ = 2 or n_���ﷶΧ = 3) And a.ִ�в���id + 0 = n_�������id And (a.ִ���� || '' = v_����Ա���� Or a.ִ���� Is Null))) And
                           a.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60
                     Order By ����ʱ��) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_�����б�.Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_�����б�.No || '"';
        v_Jtmp := v_Jtmp || ',"pati_id":' || c_�����б�.����id;
        v_Jtmp := v_Jtmp || ',"outpatient_num":"' || c_�����б�.����� || '"';
        v_Jtmp := v_Jtmp || ',"pati_name":"' || Zljsonstr(c_�����б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"pati_sex":"' || c_�����б�.�Ա� || '"';
        v_Jtmp := v_Jtmp || ',"pati_age":"' || c_�����б�.���� || '"';
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_�����б�.����, 0);
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_�����б�.����ʱ��, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_�����б�.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exetr":"' || Zljsonstr(c_�����б�.ִ����) || '"';
        v_Jtmp := v_Jtmp || ',"outp_rfrl_status":' || Nvl(c_�����б�.ת��״̬ || '', 'null');
        v_Jtmp := v_Jtmp || ',"record_sign":' || Nvl(c_�����б�.��¼��־ || '', 'null');
        v_Jtmp := v_Jtmp || ',"outptyp_name":"' || Zljsonstr(c_�����б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"pait_dept":"' || Zljsonstr(c_�����б�.���˿���) || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_�����б�.ִ��״̬ || '', 'null');
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
    End If;
  End If;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[' || c_Jtmp || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrgsapptpatilist;
/


Create Or Replace Procedure Zl_Exsesvr_Getvalidreglist
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --���ܣ�����������ȡ���˵���Ч�Һż�¼
  --��Σ�Json_In:��ʽ
  --input
  --  query_type        N    1 �������� ��1-���ݲ���id��ȡ���˵���Ч�Һż�¼��2-���ݹҺŵ���ȡ��Ϣ
  --  pati_id           N    1 ����id
  --  emg_sign          N    1 �����־ ��0-ȫ���Һ�  1-����Һ�
  --  reg_no            C    0 �Һŵ���

  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    reg_list[]           �Һż�¼�б�֧�ֶ����[����]
  --       reg_id           N   1 �Һ�id
  --       reg_no           C   1 �Һŵ�
  --       reg_properties   N   1 ��¼����
  --       exe_deptid       N   1 ִ�в���id
  --       exe_dept         C   1 ִ�в���
  --       fitem_id         N   1 �շ�ϸĿid
  --       fitem_name       C   1 �շ�ϸĿ
  --       exetr            C   1 ִ����
  --       outp_room_name   C   1 ����
  --       happen_time      C   1 ����ʱ��
  --       exe_status       N   1 ִ��״̬
  --       emg_sign         N   1 ����

  j_Input PLJson;
  j_Json  PLJson;

  n_Type             Number(5);
  n_����id           Number(18);
  n_�Ƿ���         Number(18);
  n_�������Һ����� Number(18);
  v_�Һŵ�           Varchar2(60);
  n_�Һ���Ч����     Number;
  n_������Ч����     Number;
  v_Jtmp             Varchar2(32767);
  c_Jtmp             Clob;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Type := Nvl(j_Json.Get_Number('query_type'), 0);
  If n_Type = 1 Then
    n_����id           := Nvl(j_Json.Get_Number('pati_id'), 0);
    n_�������Һ����� := To_Number(Nvl(zl_GetSysParameter(210), '0'));
    n_�Ƿ���         := Nvl(j_Json.Get_Number('emg_sign'), 0);
  
    If n_�������Һ����� = 1 Then
      For c_�Һ��б� In (Select a.Id, a.No, a.��¼����, d.Id As ����id, d.���� As ����, c.Id As ��Ŀid, c.���� As ��Ŀ, a.ִ����, a.����, a.����ʱ��,
                            a.ִ��״̬, a.����
                     From ���˹Һż�¼ A, ������ü�¼ B, �շ���ĿĿ¼ C, ���ű� D
                     Where a.No = b.No And b.��¼���� = 4 And b.��¼״̬ In (1, 0) And b.�շ���� = '1' And a.��¼���� In (1, 2) And
                           a.��¼״̬ = 1 And b.�۸񸸺� Is Null And b.�������� Is Null And b.�շ�ϸĿid = c.Id And a.ִ�в���id = d.Id And
                           a.����ʱ�� <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And a.����id = n_����id And
                           (n_�Ƿ��� = 0 Or (n_�Ƿ��� = 1 And a.���� = 1))
                     Order By ����ʱ�� Desc) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_�Һ��б�.Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_�Һ��б�.No || '"';
        v_Jtmp := v_Jtmp || ',"reg_properties":' || Nvl(c_�Һ��б�.��¼���� || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_�Һ��б�.����id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept":"' || zlJsonStr(c_�Һ��б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"fitem_id":' || Nvl(c_�Һ��б�.��Ŀid || '', 'null');
        v_Jtmp := v_Jtmp || ',"fitem_name":"' || zlJsonStr(c_�Һ��б�.��Ŀ) || '"';
        v_Jtmp := v_Jtmp || ',"exetr":"' || zlJsonStr(c_�Һ��б�.ִ����) || '"';
        v_Jtmp := v_Jtmp || ',"outp_room_name":"' || zlJsonStr(c_�Һ��б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_�Һ��б�.����ʱ��, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_�Һ��б�.ִ��״̬ || '', 'null');
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_�Һ��б�.���� || '', 'null');
        v_Jtmp := v_Jtmp || '}';
      
        If Length(v_Jtmp) > 20000 Then
          If c_Jtmp Is Null Then
            c_Jtmp := Substr(v_Jtmp, 2);
          Else
            c_Jtmp := c_Jtmp || v_Jtmp;
          End If;
          v_Jtmp := Null;
        End If;
      
      End Loop;
    Else
      n_�Һ���Ч���� := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
      n_������Ч���� := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 2, 1));
      If n_�Һ���Ч���� = 0 Then
        n_�Һ���Ч���� := 1;
      End If;
      If n_������Ч���� = 0 Then
        n_������Ч���� := 1;
      End If;
    
      For c_�Һ��б� In (Select a.Id, a.No, a.��¼����, d.Id As ����id, d.���� As ����, c.Id As ��Ŀid, c.���� As ��Ŀ, a.ִ����, a.����, a.����ʱ��,
                            a.ִ��״̬, a.����
                     From ���˹Һż�¼ A, ������ü�¼ B, �շ���ĿĿ¼ C, ���ű� D
                     Where a.No = b.No And b.��¼���� = 4 And b.��¼״̬ In (1, 0) And b.�շ���� = '1' And a.��¼���� In (1, 2) And
                           a.��¼״̬ = 1 And b.�۸񸸺� Is Null And b.�������� Is Null And b.�շ�ϸĿid = c.Id And a.ִ�в���id = d.Id And
                           a.����ʱ�� <= Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And a.����id = n_����id And
                           a.����ʱ�� Between Sysdate - Decode(a.����, 1, n_������Ч����, n_�Һ���Ч����) And
                           Trunc(Sysdate) + 1 - 1 / 24 / 60 / 60 And (n_�Ƿ��� = 0 Or (n_�Ƿ��� = 1 And a.���� = 1))
                     Order By ����ʱ�� Desc) Loop
      
        v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_�Һ��б�.Id || '', 'null');
        v_Jtmp := v_Jtmp || ',"reg_no":"' || c_�Һ��б�.No || '"';
        v_Jtmp := v_Jtmp || ',"reg_properties":' || Nvl(c_�Һ��б�.��¼���� || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_�Һ��б�.����id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept":"' || zlJsonStr(c_�Һ��б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"fitem_id":' || Nvl(c_�Һ��б�.��Ŀid || '', 'null');
        v_Jtmp := v_Jtmp || ',"fitem_name":"' || zlJsonStr(c_�Һ��б�.��Ŀ) || '"';
        v_Jtmp := v_Jtmp || ',"exetr":"' || zlJsonStr(c_�Һ��б�.ִ����) || '"';
        v_Jtmp := v_Jtmp || ',"outp_room_name":"' || zlJsonStr(c_�Һ��б�.����) || '"';
        v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_�Һ��б�.����ʱ��, 'YYYY-MM-DD HH24:MI') || '"';
        v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_�Һ��б�.ִ��״̬ || '', 'null');
        v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_�Һ��б�.���� || '', 'null');
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
    End If;
  Elsif n_Type = 2 Then
    v_�Һŵ� := j_Json.Get_String('reg_no');
    For c_�Һ��б� In (Select a.Id, a.No, a.��¼����, a.ִ�в���id As ����id, Null As ����, Null As ��Ŀid, Null As ��Ŀ, a.ִ����, a.����, a.����ʱ��,
                          a.ִ��״̬, a.����
                   From ���˹Һż�¼ A
                   Where a.No = v_�Һŵ� And a.��¼���� = 1 And a.��¼״̬ = 1) Loop
    
      v_Jtmp := v_Jtmp || ',{"reg_id":' || Nvl(c_�Һ��б�.Id || '', 'null');
      v_Jtmp := v_Jtmp || ',"reg_no":"' || c_�Һ��б�.No || '"';
      v_Jtmp := v_Jtmp || ',"reg_properties":' || Nvl(c_�Һ��б�.��¼���� || '', 'null');
      v_Jtmp := v_Jtmp || ',"exe_deptid":' || Nvl(c_�Һ��б�.����id || '', 'null');
      v_Jtmp := v_Jtmp || ',"exe_dept":"' || zlJsonStr(c_�Һ��б�.����) || '"';
      v_Jtmp := v_Jtmp || ',"fitem_id":' || Nvl(c_�Һ��б�.��Ŀid || '', 'null');
      v_Jtmp := v_Jtmp || ',"fitem_name":"' || zlJsonStr(c_�Һ��б�.��Ŀ) || '"';
      v_Jtmp := v_Jtmp || ',"exetr":"' || zlJsonStr(c_�Һ��б�.ִ����) || '"';
      v_Jtmp := v_Jtmp || ',"outp_room_name":"' || zlJsonStr(c_�Һ��б�.����) || '"';
      v_Jtmp := v_Jtmp || ',"happen_time":"' || To_Char(c_�Һ��б�.����ʱ��, 'YYYY-MM-DD HH24:MI') || '"';
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(c_�Һ��б�.ִ��״̬ || '', 'null');
      v_Jtmp := v_Jtmp || ',"emg_sign":' || Nvl(c_�Һ��б�.���� || '', 'null');
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
  End If;
  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","reg_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","reg_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getvalidreglist;
/

Create Or Replace Procedure Zl_Exsesvr_Outprevisit
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����ﲡ�˻���
  --��Σ�Json_In:��ʽ
  --input
  --  reg_id             N  1 �Һ�id
  --  exe_deptid         N  1 ִ�в���id
  --  outp_room_name     C  1 ��������
  --  outp_dr_name       C  1 ҽ��
  --  revisit_sign       N  1 �����־
  --  appt_mode_name     C  1 ԤԼ��ʽ

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_�Һ�id     Number;
  n_ִ�в���id Number;
  v_��������   Varchar2(50);
  v_ҽ��       Varchar2(50);
  n_�����־   Number;
  v_ԤԼ��ʽ   Varchar2(50);

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�Һ�id     := j_Json.Get_Number('reg_id');
  n_ִ�в���id := j_Json.Get_Number('exe_deptid');
  v_��������   := j_Json.Get_String('outp_room_name');
  v_ҽ��       := j_Json.Get_String('outp_dr_name');
  n_�����־   := j_Json.Get_Number('revisit_sign');
  v_ԤԼ��ʽ   := j_Json.Get_String('appt_mode_name');

  Zl_���˹Һż�¼_����(n_�Һ�id, n_ִ�в���id, v_��������, v_ҽ��, n_�����־, v_ԤԼ��ʽ);
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outprevisit;
/

Create Or Replace Procedure Zl_Exsesvr_Outpfinish
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����ﲡ����ɽ���
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id            N  1 ����ID
  --  reg_no             C  1 �Һ�no
  --  outp_room_name     C  1 ��������
  --  exetr              C  1 ִ����
  --  fnsh_desc          C  1 ���ժҪ
  --  ext_mark           N  1 ���ӱ�־ Ϊ1ʱ,��ʾ��ʿ��ɾ���;2ʱ��ʾ����ϵͳ����ǼǵĹҺ�����(���ַ�ʽ���������ü�¼�͹ҺŻ���,��Ǽ�);3-��������ϵͳͬ����¼

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_����id   Number;
  v_�Һŵ�   Varchar2(50);
  v_������� Varchar2(50);
  v_ִ����   Varchar2(50);
  v_���ժҪ Varchar2(4000);
  n_���ӱ�־ Number;
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  v_�Һŵ�   := j_Json.Get_String('reg_no');
  v_������� := j_Json.Get_String('outp_room_name');
  v_ִ����   := j_Json.Get_String('exetr');
  v_���ժҪ := j_Json.Get_String('fnsh_desc');
  n_���ӱ�־ := j_Json.Get_Number('ext_mark');

  If n_���ӱ�־ = 0 Then
    n_���ӱ�־ := Null;
  End If;

  Zl_���˽������_s(n_����id, v_�Һŵ�, v_�������, v_ִ����, v_���ժҪ, n_���ӱ�־);

  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpfinish;
/


Create Or Replace Procedure Zl_Exsesvr_Outpfinishcancel
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����ﲡ��ȡ����ɽ���
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id            N  1 ����ID
  --  reg_no             C  1 �Һ�no
  --  ext_mark           N  1 ���ӱ�־ Ϊ1ʱ,��ʾ��ʿ��ɾ���;2ʱ��ʾ����ϵͳ����ǼǵĹҺ�����(���ַ�ʽ���������ü�¼�͹ҺŻ���,��Ǽ�);3-��������ϵͳͬ����¼

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_����id   Number;
  v_�Һŵ�   Varchar2(50);
  n_���ӱ�־ Number;
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  v_�Һŵ�   := j_Json.Get_String('reg_no');
  n_���ӱ�־ := j_Json.Get_Number('ext_mark');
  If n_���ӱ�־ = 0 Then
    n_���ӱ�־ := Null;
  End If;
  Zl_���˽������_Cancel_s(n_����id, v_�Һŵ�, n_���ӱ�־);
  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpfinishcancel;
/


Create Or Replace Procedure Zl_Exsesvr_Outpreceive
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����ﲡ�˽���
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id            N  1 ����ID
  --  reg_no             C  1 �Һ�no
  --  exe_deptid         N  1 ִ�в���id
  --  exetr              C  1 ִ����
  --  outp_room_name     C  1 ��������
  --  emg_sign           N  1 �����־
  --  revisit_sign       N  1 �����־
  --  exe_time           C  1 ִ��ʱ��  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_����id     Number;
  v_�Һŵ�     Varchar2(50);
  n_ִ�в���id Number;
  v_ִ����     Varchar2(50);
  v_�������   Varchar2(50);
  n_�����־   Number;
  n_�����־   Number;
  d_ִ��ʱ��   Date;
  j_Input      PLJson;
  j_Json       PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := j_Json.Get_Number('pati_id');
  v_�Һŵ�     := j_Json.Get_String('reg_no');
  n_ִ�в���id := j_Json.Get_Number('exe_deptid');
  v_ִ����     := j_Json.Get_String('exetr');
  v_�������   := j_Json.Get_String('outp_room_name');
  n_�����־   := j_Json.Get_Number('emg_sign');
  n_�����־   := j_Json.Get_Number('revisit_sign');
  d_ִ��ʱ��   := To_Date(j_Json.Get_String('exe_time'), 'YYYY-MM-DD HH24:MI:SS');

  If n_ִ�в���id = 0 Then
    n_ִ�в���id := Null;
  End If;

  Zl_���˽���_s(n_����id, v_�Һŵ�, n_ִ�в���id, v_ִ����, v_�������, n_�����־, n_�����־, d_ִ��ʱ��);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpreceive;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Outpreceivecancel
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����ﲡ��ȡ������
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id            N  1 ����ID
  --  reg_no             C  1 �Һ�no
  --  exe_deptid         N  1 ִ�в���id
  --  exetr              C  1 ִ����
  --  referral_sign      N  1 �Ƿ�ת�� 0-δת��  1-ת��
  --  referral_deptid    N  1 ת�����id
  --  referral_doctor    C  1 ת��ҽ��

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_����id     Number;
  v_�Һŵ�     Varchar2(50);
  n_ִ�в���id Number;
  v_ִ����     Varchar2(50);
  n_ת���־   Number;
  n_ת�����id Number;
  v_ת��ҽ��   Varchar2(50);
  j_Input      PLJson;
  j_Json       PLJson;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := j_Json.Get_Number('pati_id');
  v_�Һŵ�     := j_Json.Get_String('reg_no');
  n_ִ�в���id := j_Json.Get_Number('exe_deptid');
  v_ִ����     := j_Json.Get_String('exetr');
  n_ת���־   := Nvl(j_Json.Get_Number('referral_sign'), 0);
  v_ת��ҽ��   := j_Json.Get_String('referral_doctor');
  n_ת�����id := j_Json.Get_Number('referral_deptid');

  If n_ִ�в���id = 0 Then
    n_ִ�в���id := Null;
  End If;
  If n_ת�����id = 0 Then
    n_ת�����id := Null;
  End If;

  Zl_���˽���_Cancel_s(n_����id, v_�Һŵ�, n_ִ�в���id, v_ִ����, n_ת���־, n_ת�����id, v_ת��ҽ��);
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpreceivecancel;
/

Create Or Replace Procedure Zl_Exsesvr_Outpreferral
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ɲ���ת�ת����գ�ȡ��ת��ܾ�ת�﹦��
  --��Σ�Json_In:��ʽ
  --input
  --  reg_no             C  1 �Һ�no
  --  referral_state     N  1 ת��״̬ 0:ת��(��Ҫ������������),1:����,-1:�ܾ�,Null:ȡ��ת��
  --  referral_deptid    N  1 ת�����id
  --  referral_outproom  C  1 ת������
  --  referral_doctor    C  1 ת��ҽ��

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------

  v_�Һŵ�     Varchar2(50);
  n_ת��״̬   Number;
  n_ת�����id Number;
  v_ת������   Varchar2(50);
  v_ת��ҽ��   Varchar2(50);
  j_Input      PLJson;
  j_Json       PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�     := j_Json.Get_String('reg_no');
  v_ת������   := j_Json.Get_String('referral_outproom');
  v_ת��ҽ��   := j_Json.Get_String('referral_doctor');
  n_ת��״̬   := j_Json.Get_Number('referral_state');
  n_ת�����id := j_Json.Get_Number('referral_deptid');

  Zl_���˹Һż�¼_ת��_s(v_�Һŵ�, n_ת��״̬, n_ת�����id, v_ת������, v_ת��ҽ��);
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outpreferral;
/


Create Or Replace Procedure Zl_Exsesvr_Outprevisitcancel
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����ﲡ��ȡ������
  --��Σ�Json_In:��ʽ
  --input
  --  reg_id             N  1 �Һ�id
  --  revisit_sign       N  1 �����־

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_�Һ�id   Number;
  n_�����־ Number;
  j_Input    PLJson;
  j_Json     PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�Һ�id   := j_Json.Get_Number('reg_id');
  n_�����־ := j_Json.Get_Number('revisit_sign');

  Zl_���˹Һż�¼_ȡ������(n_�Һ�id, n_�����־);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Outprevisitcancel;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Getqueuecallcount
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡ��ǰ�Ŷӽкŵĺ�������
  --��Σ�Json_In:��ʽ
  --input
  --  operator_name  C  1 ����Ա����
  --  emg_sign       N  0 �����־
  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    call_count                      N 1 ��������
  -------------------------------------------

  v_����Ա����   Varchar2(50);
  n_�Һ���Ч���� Number;
  n_������Ч���� Number;
  n_���к�����   Number;
  n_��������     Number;
  n_�����־     Number;

  v_Output Varchar2(1000);
  j_Input  PLJson;
  j_Json   PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����Ա���� := j_Json.Get_String('operator_name');
  n_�����־   := Nvl(j_Json.Get_Number('emg_sign'), 0);

  n_�Һ���Ч���� := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
  n_������Ч���� := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 2, 1));

  If n_�Һ���Ч���� = 0 Then
    n_�Һ���Ч���� := 1;
  End If;
  If n_������Ч���� = 0 Then
    n_������Ч���� := 1;
  End If;
  n_���к����� := To_Number(Nvl(zl_GetSysParameter('��������������', 1260), 0));

  If n_�����־ = 1 Then
    If n_���к����� <> 1 Then
      Select Count(Distinct b.Id) As Count
      Into n_��������
      From ���˹Һż�¼ B, �ŶӽкŶ��� A
      Where a.ҵ��id = b.Id And a.ҵ������ = 0 And Instr(',0,4,', ',' || a.�Ŷ�״̬ || ',') = 0 And b.��¼���� = 1 And b.��¼״̬ = 1 And
            a.ҽ������ || '' = v_����Ա���� And Nvl(a.�������, 0) = 0 And (Nvl(b.����, 0) = 1 And b.����ʱ�� >= Sysdate - n_������Ч����);
    Else
      Select Count(Distinct b.Id) As Count
      Into n_��������
      From ���˹Һż�¼ B, �ŶӽкŶ��� A
      Where a.ҵ��id = b.Id And a.ҵ������ = 0 And Instr(',0,4,6,', ',' || a.�Ŷ�״̬ || ',') = 0 And b.��¼���� = 1 And b.��¼״̬ = 1 And
            a.ҽ������ || '' = v_����Ա���� And Nvl(b.����, 0) = 1 And b.����ʱ�� >= Sysdate - n_������Ч����;
    End If;
  Else
    If n_���к����� <> 1 Then
      Select Count(Distinct b.Id) As Count
      Into n_��������
      From ���˹Һż�¼ B, �ŶӽкŶ��� A
      Where a.ҵ��id = b.Id And a.ҵ������ = 0 And Instr(',0,4,', ',' || a.�Ŷ�״̬ || ',') = 0 And b.��¼���� = 1 And b.��¼״̬ = 1 And
            a.ҽ������ || '' = v_����Ա���� And Nvl(a.�������, 0) = 0 And ((Nvl(b.����, 0) = 1 And b.����ʱ�� >= Sysdate - n_������Ч����) Or
            (Nvl(b.����, 0) <> 1 And b.����ʱ�� >= Sysdate - n_�Һ���Ч����));
    Else
      Select Count(Distinct b.Id) As Count
      Into n_��������
      From ���˹Һż�¼ B, �ŶӽкŶ��� A
      Where a.ҵ��id = b.Id And a.ҵ������ = 0 And Instr(',0,4,6,', ',' || a.�Ŷ�״̬ || ',') = 0 And b.��¼���� = 1 And b.��¼״̬ = 1 And
            a.ҽ������ || '' = v_����Ա���� And ((Nvl(b.����, 0) = 1 And b.����ʱ�� >= Sysdate - n_������Ч����) Or
            (Nvl(b.����, 0) <> 1 And b.����ʱ�� >= Sysdate - n_�Һ���Ч����));
    End If;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'call_count', n_��������, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getqueuecallcount;
/


Create Or Replace Procedure Zl_Exsesvr_Checkqueuedate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ��Դ˹Һ����ڵĶ��н�����Ч���
  --��Σ�Json_In:��ʽ
  --input
  --  reg_id       N  1 �Һ�id
  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    result                          C 1 ����ֵ:��������|��ʾ��Ϣ
  -------------------------------------------
  j_Input  PLJson;
  j_Json   PLJson;
  n_�Һ�id Number;
  v_Out    Varchar2(500);
  v_Output Varchar2(32767);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�Һ�id := j_Json.Get_Number('reg_id');

  v_Out := Zl_Queuedatecheck(n_�Һ�id);

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'result', v_Out, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkqueuedate;
/


Create Or Replace Procedure Zl_Exsesvr_Getqueuereginfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --���ܣ�����������ȡ�ŶӽкŶ�����Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --  query_type        N    1 �������� ��1-ͨ��ҵ��ids��ȡ�ŶӵĹҺ���Ϣ,2-ͨ���Һŵ���ȡ�ŶӵĹҺ���Ϣ��3-ͨ���������Ʋ�ѯ
  --  business_ids      C      ҵ��ids
  --  reg_no            C      �Һŵ�
  --  queue_name        C      ��������
  --  queue_state       C      �Ŷ�״̬
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    queue_list   �Ŷӽк��б�֧�ֶ����[����]
  --       reg_id           N   1 �Һ�id
  --       reg_no           C   1 �Һŵ�
  --       pati_id          N   1 ����id
  --       exec_deptid      N   1 ִ�в���id
  --       exec_state       N   1 ִ��״̬
  --       outp_room        C   1 ����
  --       queue_num        C   1 �ŶӺ���
  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  n_Type     Number(18);
  v_ҵ��ids  Varchar(32767);
  v_״̬     Varchar(4000);
  v_�Һŵ�   Varchar(50);
  v_�������� Varchar(200);
Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Type := j_Json.Get_Number('query_type');
  If n_Type = 1 Then
    v_ҵ��ids := j_Json.Get_String('business_ids');
    v_״̬    := j_Json.Get_String('queue_state');
  
    If v_ҵ��ids Is Null Then
      Json_Out := zlJsonOut('δ����ҵ��id');
      Return;
    End If;
  
    For c_������Ϣ In (Select a.Id, a.No, a.����id, a.ִ�в���id, a.ִ��״̬, b.����, b.�ŶӺ���
                   From ���˹Һż�¼ A, �ŶӽкŶ��� B
                   Where a.Id = b.ҵ��id And b.ҵ������ = 0 And a.Id In (Select Column_Value From Table(f_Str2List(v_ҵ��ids))) And
                         (Instr(',' || v_״̬ || ',', ',' || Nvl(b.�Ŷ�״̬, 0) || ',') > 0 Or v_״̬ Is Null) And
                         a.��¼���� In (1, 2) And a.��¼״̬ = 1) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'reg_id', c_������Ϣ.Id, 1, 1);
      zlJsonPutValue(v_Output, 'reg_no', c_������Ϣ.No);
      zlJsonPutValue(v_Output, 'pati_id', c_������Ϣ.����id, 1);
      zlJsonPutValue(v_Output, 'exec_deptid', c_������Ϣ.ִ�в���id, 1);
      zlJsonPutValue(v_Output, 'exec_state', c_������Ϣ.ִ��״̬, 1);
      zlJsonPutValue(v_Output, 'outp_room', c_������Ϣ.����);
      zlJsonPutValue(v_Output, 'queue_num', c_������Ϣ.�ŶӺ���, 0, 2);
    
    End Loop;
  Elsif n_Type = 2 Then
    v_�Һŵ� := j_Json.Get_String('reg_no');
    v_״̬   := j_Json.Get_String('queue_state');
  
    If v_�Һŵ� Is Null Then
      Json_Out := zlJsonOut('δ����Һŵ�');
      Return;
    End If;
  
    For c_������Ϣ In (Select a.Id, a.No, a.����id, a.ִ�в���id, a.ִ��״̬, b.����, b.�ŶӺ���
                   From ���˹Һż�¼ A, �ŶӽкŶ��� B
                   Where a.Id = b.ҵ��id And b.ҵ������ = 0 And a.No = v_�Һŵ� And
                         (Instr(',' || v_״̬ || ',', ',' || Nvl(b.�Ŷ�״̬, 0) || ',') > 0 Or v_״̬ Is Null) And
                         a.��¼���� In (1, 2) And a.��¼״̬ = 1) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      End If;
      zlJsonPutValue(v_Output, 'reg_id', c_������Ϣ.Id, 1, 1);
      zlJsonPutValue(v_Output, 'reg_no', c_������Ϣ.No);
      zlJsonPutValue(v_Output, 'pati_id', c_������Ϣ.����id, 1);
      zlJsonPutValue(v_Output, 'exec_deptid', c_������Ϣ.ִ�в���id, 1);
      zlJsonPutValue(v_Output, 'exec_state', c_������Ϣ.ִ��״̬, 1);
      zlJsonPutValue(v_Output, 'outp_room', c_������Ϣ.����);
      zlJsonPutValue(v_Output, 'queue_num', c_������Ϣ.�ŶӺ���, 0, 2);
    
    End Loop;
  Elsif n_Type = 3 Then
    v_�������� := j_Json.Get_String('queue_name');
    v_״̬     := j_Json.Get_String('queue_state');
  
    For c_������Ϣ In (Select Distinct /*+ Rule*/ a.Id, a.No, a.����id, a.ִ�в���id, a.ִ��״̬, b.����, b.�ŶӺ���
                   From ���˹Һż�¼ A, �ŶӽкŶ��� B
                   Where a.Id = b.ҵ��id And b.�������� = v_�������� And Nvl(b.ҵ������, 0) = 0 And
                         (Instr(',' || v_״̬ || ',', ',' || Nvl(b.�Ŷ�״̬, 0) || ',') > 0 Or v_״̬ Is Null) And
                         Nvl(a.����id, 0) = 0 And a.��¼���� = 1 And a.��¼״̬ = 1) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      End If;
      zlJsonPutValue(v_Output, 'reg_id', c_������Ϣ.Id, 1, 1);
      zlJsonPutValue(v_Output, 'reg_no', c_������Ϣ.No);
      zlJsonPutValue(v_Output, 'pati_id', c_������Ϣ.����id, 1);
      zlJsonPutValue(v_Output, 'exec_deptid', c_������Ϣ.ִ�в���id, 1);
      zlJsonPutValue(v_Output, 'exec_state', c_������Ϣ.ִ��״̬, 1);
      zlJsonPutValue(v_Output, 'outp_room', c_������Ϣ.����);
      zlJsonPutValue(v_Output, 'queue_num', c_������Ϣ.�ŶӺ���, 0, 2);
    End Loop;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","queue_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","queue_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getqueuereginfo;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Getqueuereglist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------
  --���ܣ�����������ȡ�Ŷӽкŵĺ��ﲡ���б�
  --��Σ�Json_In:��ʽ
  --input
  --  outp_room_names        C    1 ���������ַ���       �Զ��ŷָ�
  --  outp_dr_names          C    1 ����ҽ���ַ���       �Զ��ŷָ�
  --  recipe_exe_status      C    1 ִ��״̬�ַ���       �Զ��ŷָ�
  --  outpque_names          C    1 �ŶӶ��������ַ���   �Զ��ŷָ�
  --  view_type              N    1 �Ŷӷ�����ʾ����     1-�����з��� 2-��ҽ����������  3-�����ҷ���

  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_list           �����б�֧�ֶ����[����]
  --       outpque_id             N   1 �Ŷ�ID
  --       pati_id                N   1 ����id
  --       outpque_name           C   1 �ŶӶ�������
  --       outpque_sno            N   1 �Ŷ����
  --       business_type          C   1 ҵ������
  --       business_id            N   1 ҵ��id
  --       dept_id                N   1 ����id
  --       dept_name              C   1 ��������
  --       outpque_no             C   1 �ŶӺ���
  --       outpque_sign           N   1 �Ŷӱ�־
  --       pati_name              C   1 ��������
  --       pati_age               C   1 ��������
  --       outp_room_name         C   1 ��������
  --       outp_dr_name           C   1 ����ҽ��
  --       call_dr_name           C   1 ����ҽ��
  --       outpat_pri             N   1 ����
  --       revisit_sno            N   1 �������
  --       outpque_time           C   1 �Ŷ�ʱ��
  --       call_time              C   1 ����ʱ��
  --       outpque_state          N   1 �Ŷ�״̬
  --       outpque_revisit_num    N   1 ���������

  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);
  c_Output Clob;

  v_���������ַ��� Varchar2(32767);
  n_���й���       Number;
  v_����ҽ���ַ��� Varchar(4000);
  v_���������ַ��� Varchar(4000);
  v_ִ��״̬�ַ��� Varchar(4000);
  n_��ʾ����       Number;
  n_���ﲡ������   Number;

  l_�������� t_StrList := t_StrList();

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����ҽ���ַ��� := j_Json.Get_String('outp_dr_names');
  v_���������ַ��� := j_Json.Get_String('outp_room_names');
  v_ִ��״̬�ַ��� := j_Json.Get_String('recipe_exe_status');
  n_��ʾ����       := Nvl(j_Json.Get_Number('view_type'), 0);
  n_���ﲡ������   := Nvl(To_Number(zl_GetSysParameter('���ﲡ���Ƿ�����', 1160)), 1);
  v_���������ַ��� := j_Json.Get_String('outpque_names');

  If v_���������ַ��� Is Not Null Then
    n_���й��� := 1;
  End If;

  While v_���������ַ��� Is Not Null Loop
    If Length(v_���������ַ���) <= 4000 Then
      l_��������.Extend;
      l_��������(l_��������.Count) := v_���������ַ���;
      v_���������ַ��� := Null;
    Else
      l_��������.Extend;
      l_��������(l_��������.Count) := Substr(v_���������ַ���, 1, Instr(v_���������ַ���, ',', 3940) - 1);
      v_���������ַ��� := Substr(v_���������ַ���, Instr(v_���������ַ���, ',', 3940) + 1);
    End If;
  End Loop;

  For I In 1 .. l_��������.Count Loop

    If n_��ʾ���� = 1 Then
      For c_�����б� In (Select /*+ Rule*/
                      To_Number(a.Id) As ID, To_Number(a.����id) As ����id, a.��������, a.�Ŷ����, To_Number(a.ҵ������) As ҵ������,
                      To_Number(a.ҵ��id) As ҵ��id, To_Number(a.����id) As ����id, x.���� As ��������, a.�ŶӺ���, a.�Ŷӱ��,
                      a.�������� || Decode(e.ԤԼ, 1, '(Ԥ)', Null) As ��������, e.����, a.����, a.ҽ������,
                      (Select j.���� From ��Ա�� J, �ϻ���Ա�� K Where j.Id = k.��Աid And k.�û��� = a.����ҽ��) As ����ҽ��,
                      To_Number(a.����) As ����, To_Number(a.�������) As �������, To_Char(a.�Ŷ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ŷ�ʱ��,
                      To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, To_Number(a.�Ŷ�״̬) As �Ŷ�״̬,
                      Decode(n_���ﲡ������, 1, To_Number(Nvl(a.�������, 9999999999)), 0) As ���������
                     From �ŶӽкŶ��� A, ���ű� X,
                          (Select /*+cardinality(f,10)*/
                             Column_Value As ��������
                            From Table(f_Str2List(l_��������(I))) F) B,
                          Table(Cast(f_Str2List(v_���������ַ���) As Zltools.t_Strlist)) C,
                          Table(Cast(f_Str2List(v_����ҽ���ַ���) As Zltools.t_Strlist)) D, ���˹Һż�¼ E
                     Where To_Number(a.ҵ��id) = e.Id And
                           (Nvl(a.�Ƿ��ʱ��, 0) = 0 And a.�Ŷ�ʱ�� <= Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 Or
                            Nvl(a.�Ƿ��ʱ��, 0) = 1 And Sysdate > a.�Ŷ�ʱ��) And (a.�������� = b.�������� Or n_���й��� Is Null) And
                           (v_ִ��״̬�ַ��� Is Null Or Instr(v_ִ��״̬�ַ���, a.�Ŷ�״̬) = 0) And x.Id = a.����id And
                           ((a.���� = c.Column_Value And a.ҽ������ Is Null) Or a.ҽ������ = d.Column_Value Or
                            (a.���� Is Null And a.ҽ������ Is Null)) And Nvl(a.�Ŷ�״̬, 0) <> 8
                     Order By �Ŷ�״̬ Desc, �Ŷ����, ���� Desc, ���������, �Ŷ�ʱ��, �ŶӺ���) Loop
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;
        zlJsonPutValue(v_Output, 'outpque_id', c_�����б�.Id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_id', c_�����б�.����id, 1);
        zlJsonPutValue(v_Output, 'outpque_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'outpque_sno', c_�����б�.�Ŷ����, 1);
        zlJsonPutValue(v_Output, 'business_type', c_�����б�.ҵ������);
        zlJsonPutValue(v_Output, 'business_id', c_�����б�.ҵ��id, 1);
        zlJsonPutValue(v_Output, 'dept_id', c_�����б�.����id, 1);
        zlJsonPutValue(v_Output, 'dept_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'outpque_no', c_�����б�.�ŶӺ���);
        zlJsonPutValue(v_Output, 'outpque_sign', c_�����б�.�Ŷӱ��);
        zlJsonPutValue(v_Output, 'pati_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'pati_age', c_�����б�.����);
        zlJsonPutValue(v_Output, 'outp_room_name', c_�����б�.����);
        zlJsonPutValue(v_Output, 'outp_dr_name', c_�����б�.ҽ������);
        zlJsonPutValue(v_Output, 'call_dr_name', c_�����б�.����ҽ��);
        zlJsonPutValue(v_Output, 'outpat_pri', c_�����б�.����, 1);
        zlJsonPutValue(v_Output, 'revisit_sno', c_�����б�.�������, 1);
        zlJsonPutValue(v_Output, 'outpque_time', c_�����б�.�Ŷ�ʱ��);
        zlJsonPutValue(v_Output, 'call_time', c_�����б�.����ʱ��);
        zlJsonPutValue(v_Output, 'outpque_state', c_�����б�.�Ŷ�״̬, 1);
        zlJsonPutValue(v_Output, 'outpque_revisit_num', c_�����б�.���������, 1, 2);

      End Loop;
    Elsif n_��ʾ���� = 2 Then
      For c_�����б� In (Select /*+ Rule*/
                      To_Number(a.Id) As ID, To_Number(a.����id) As ����id, a.��������, a.�Ŷ����, To_Number(a.ҵ������) As ҵ������,
                      To_Number(a.ҵ��id) As ҵ��id, To_Number(a.����id) As ����id, x.���� As ��������, a.�ŶӺ���, a.�Ŷӱ��,
                      a.�������� || Decode(e.ԤԼ, 1, '(Ԥ)', Null) As ��������, e.����, a.����, a.ҽ������,
                      (Select j.���� From ��Ա�� J, �ϻ���Ա�� K Where j.Id = k.��Աid And k.�û��� = a.����ҽ��) As ����ҽ��,
                      To_Number(a.����) As ����, To_Number(a.�������) As �������, To_Char(a.�Ŷ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ŷ�ʱ��,
                      To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, To_Number(a.�Ŷ�״̬) As �Ŷ�״̬,
                      Decode(n_���ﲡ������, 1, To_Number(Nvl(a.�������, 9999999999)), 0) As ���������
                     From �ŶӽкŶ��� A, ���ű� X,
                          (Select /*+cardinality(f,10)*/
                             Column_Value As ��������
                            From Table(f_Str2List(l_��������(I))) F) B,
                          Table(Cast(f_Str2List(v_���������ַ���) As Zltools.t_Strlist)) C,
                          Table(Cast(f_Str2List(v_����ҽ���ַ���) As Zltools.t_Strlist)) D, ���˹Һż�¼ E
                     Where To_Number(a.ҵ��id) = e.Id And
                           (Nvl(a.�Ƿ��ʱ��, 0) = 0 And a.�Ŷ�ʱ�� <= Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 Or
                            Nvl(a.�Ƿ��ʱ��, 0) = 1 And Sysdate > a.�Ŷ�ʱ��) And (a.�������� = b.�������� Or n_���й��� Is Null) And
                           (v_ִ��״̬�ַ��� Is Null Or Instr(v_ִ��״̬�ַ���, a.�Ŷ�״̬) = 0) And x.Id = a.����id And
                           (a.���� = c.Column_Value And (a.ҽ������ Is Null Or a.ҽ������ = d.Column_Value)) And
                           Nvl(a.�Ŷ�״̬, 0) <> 8
                     Order By �Ŷ�״̬ Desc, �Ŷ����, ���� Desc, ���������, �Ŷ�ʱ��, �ŶӺ���) Loop
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;

        zlJsonPutValue(v_Output, 'outpque_id', c_�����б�.Id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_id', c_�����б�.����id, 1);
        zlJsonPutValue(v_Output, 'outpque_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'outpque_sno', c_�����б�.�Ŷ����, 1);
        zlJsonPutValue(v_Output, 'business_type', c_�����б�.ҵ������);
        zlJsonPutValue(v_Output, 'business_id', c_�����б�.ҵ��id, 1);
        zlJsonPutValue(v_Output, 'dept_id', c_�����б�.����id, 1);
        zlJsonPutValue(v_Output, 'dept_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'outpque_no', c_�����б�.�ŶӺ���);
        zlJsonPutValue(v_Output, 'outpque_sign', c_�����б�.�Ŷӱ��);
        zlJsonPutValue(v_Output, 'pati_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'pati_age', c_�����б�.����);
        zlJsonPutValue(v_Output, 'outp_room_name', c_�����б�.����);
        zlJsonPutValue(v_Output, 'outp_dr_name', c_�����б�.ҽ������);
        zlJsonPutValue(v_Output, 'call_dr_name', c_�����б�.����ҽ��);
        zlJsonPutValue(v_Output, 'outpat_pri', c_�����б�.����, 1);
        zlJsonPutValue(v_Output, 'revisit_sno', c_�����б�.�������, 1);
        zlJsonPutValue(v_Output, 'outpque_time', c_�����б�.�Ŷ�ʱ��);
        zlJsonPutValue(v_Output, 'call_time', c_�����б�.����ʱ��);
        zlJsonPutValue(v_Output, 'outpque_state', c_�����б�.�Ŷ�״̬, 1);
        zlJsonPutValue(v_Output, 'outpque_revisit_num', c_�����б�.���������, 1, 2);

      End Loop;
    Elsif n_��ʾ���� = 3 Then
      For c_�����б� In (Select /*+ Rule*/
                      To_Number(a.Id) As ID, To_Number(a.����id) As ����id, a.��������, a.�Ŷ����, To_Number(a.ҵ������) As ҵ������,
                      To_Number(a.ҵ��id) As ҵ��id, To_Number(a.����id) As ����id, x.���� As ��������, a.�ŶӺ���, a.�Ŷӱ��,
                      a.�������� || Decode(e.ԤԼ, 1, '(Ԥ)', Null) As ��������, e.����, a.����, a.ҽ������,
                      (Select j.���� From ��Ա�� J, �ϻ���Ա�� K Where j.Id = k.��Աid And k.�û��� = a.����ҽ��) As ����ҽ��,
                      To_Number(a.����) As ����, To_Number(a.�������) As �������, To_Char(a.�Ŷ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ŷ�ʱ��,
                      To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, To_Number(a.�Ŷ�״̬) As �Ŷ�״̬,
                      Decode(n_���ﲡ������, 1, To_Number(Nvl(a.�������, 9999999999)), 0) As ���������
                     From �ŶӽкŶ��� A, ���ű� X,
                          (Select /*+cardinality(f,10)*/
                             Column_Value As ��������
                            From Table(f_Str2List(l_��������(I))) F) B, Table(f_Str2List(v_����ҽ���ַ���)) D, ���˹Һż�¼ E
                     Where To_Number(a.ҵ��id) = e.Id And
                           (Nvl(a.�Ƿ��ʱ��, 0) = 0 And a.�Ŷ�ʱ�� <= Trunc(Sysdate + 1) - 1 / 24 / 60 / 60 Or
                            Nvl(a.�Ƿ��ʱ��, 0) = 1 And Sysdate > a.�Ŷ�ʱ��) And (a.�������� = b.�������� Or n_���й��� Is Null) And
                           (v_ִ��״̬�ַ��� Is Null Or Instr(v_ִ��״̬�ַ���, a.�Ŷ�״̬) = 0) And x.Id = a.����id And
                           a.ҽ������ = d.Column_Value And Nvl(a.�Ŷ�״̬, 0) <> 8
                     Order By �Ŷ�״̬ Desc, �Ŷ����, ���� Desc, ���������, �Ŷ�ʱ��, �ŶӺ���) Loop

        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;

          v_Output := Null;
        End If;

        zlJsonPutValue(v_Output, 'outpque_id', c_�����б�.Id, 1, 1);
        zlJsonPutValue(v_Output, 'pati_id', c_�����б�.����id, 1);
        zlJsonPutValue(v_Output, 'outpque_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'outpque_sno', c_�����б�.�Ŷ����, 1);
        zlJsonPutValue(v_Output, 'business_type', c_�����б�.ҵ������);
        zlJsonPutValue(v_Output, 'business_id', c_�����б�.ҵ��id, 1);
        zlJsonPutValue(v_Output, 'dept_id', c_�����б�.����id, 1);
        zlJsonPutValue(v_Output, 'dept_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'outpque_no', c_�����б�.�ŶӺ���);
        zlJsonPutValue(v_Output, 'outpque_sign', c_�����б�.�Ŷӱ��);
        zlJsonPutValue(v_Output, 'pati_name', c_�����б�.��������);
        zlJsonPutValue(v_Output, 'pati_age', c_�����б�.����);
        zlJsonPutValue(v_Output, 'outp_room_name', c_�����б�.����);
        zlJsonPutValue(v_Output, 'outp_dr_name', c_�����б�.ҽ������);
        zlJsonPutValue(v_Output, 'call_dr_name', c_�����б�.����ҽ��);
        zlJsonPutValue(v_Output, 'outpat_pri', c_�����б�.����, 1);
        zlJsonPutValue(v_Output, 'revisit_sno', c_�����б�.�������, 1);
        zlJsonPutValue(v_Output, 'outpque_time', c_�����б�.�Ŷ�ʱ��);
        zlJsonPutValue(v_Output, 'call_time', c_�����б�.����ʱ��);
        zlJsonPutValue(v_Output, 'outpque_state', c_�����б�.�Ŷ�״̬, 1);
        zlJsonPutValue(v_Output, 'outpque_revisit_num', c_�����б�.���������, 1, 2);

      End Loop;
    End If;
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","pati_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","pati_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getqueuereglist;
/

 

Create Or Replace Procedure Zl_Exsesvr_Updateregstate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����²��˾���״̬
  --��Σ�Json_In:��ʽ
  --input
  --  reg_no             C  1 �Һŵ�
  --  exe_status         N  1 ִ��״̬ 0:���Ϊ����.-1:���Ϊ������
  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  v_�Һŵ�   Varchar2(50);
  n_ִ��״̬ Number;

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�   := j_Json.Get_String('reg_no');
  n_ִ��״̬ := Nvl(j_Json.Get_Number('exe_status'), 0);

  Zl_���˹Һż�¼_״̬(v_�Һŵ�, n_ִ��״̬);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateregstate;
/

Create Or Replace Procedure Zl_Exsesvr_Updatereginfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����²��˹Һ���Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --  reg_no    C  1 �Һ�no  ͨ���Һ�no��������Ϣ
  --  reg_id    N  1 �Һ�id  ͨ���Һ�id��������Ϣ
  --      update_list            ���¹Һ���Ϣ�б�ֻ��һ��
  --        pay_method     C   ҽ�Ƹ��ʽ
  --        fee_category   C   �ѱ�
  --        community_num  N   �������
  --        pati_name      C   ����
  --        pati_sex       C   �Ա�
  --        pati_age       C   ����

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------

  v_�Һŵ� Varchar2(50);
  n_�Һ�id Number;

  v_ҽ�Ƹ��ʽ Varchar2(50);
  v_�ѱ�         Varchar2(50);

  n_������� Number;
  v_����     ���˹Һż�¼.����%Type;
  v_�Ա�     ���˹Һż�¼.�Ա�%Type;
  v_����     ���˹Һż�¼.����%Type;

  n_ҽ�Ƹ��ʽ_b Number(1);
  n_�ѱ�_b         Number(1);

  n_�������_b Number(1);
  n_����_b     Number(1);
  n_�Ա�_b     Number(1);
  n_����_b     Number(1);

  j_Input PLJson;
  j_Json  PLJson;
  o_Json  PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�Һ�id := Nvl(j_Json.Get_Number('reg_id'), 0);
  v_�Һŵ� := j_Json.Get_String('reg_no');
  o_Json   := j_Json.Get_Pljson('update_list');

  If o_Json.Exist('fee_category') Then
    v_�ѱ�   := o_Json.Get_String('fee_category');
    n_�ѱ�_b := 1;
  End If;

  If o_Json.Exist('pay_method') Then
    v_ҽ�Ƹ��ʽ   := o_Json.Get_String('pay_method');
    n_ҽ�Ƹ��ʽ_b := 1;
  End If;

  If o_Json.Exist('community_num') Then
    n_�������   := o_Json.Get_Number('community_num');
    n_�������_b := 1;
  End If;

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

  If n_�Һ�id <> 0 Then
    Update ���˹Һż�¼
    Set ҽ�Ƹ��ʽ = Decode(n_ҽ�Ƹ��ʽ_b, 1, v_ҽ�Ƹ��ʽ, ҽ�Ƹ��ʽ), �ѱ� = Decode(n_�ѱ�_b, 1, v_�ѱ�, �ѱ�),
        ���� = Decode(n_�������_b, 1, n_�������, ����), ���� = Decode(n_����_b, 1, v_����, ����), �Ա� = Decode(n_�Ա�_b, 1, v_�Ա�, �Ա�),
        ���� = Decode(n_����_b, 1, v_����, ����)
    Where ID = n_�Һ�id;
  Else
    Update ���˹Һż�¼
    Set ҽ�Ƹ��ʽ = Decode(n_ҽ�Ƹ��ʽ_b, 1, v_ҽ�Ƹ��ʽ, ҽ�Ƹ��ʽ), �ѱ� = Decode(n_�ѱ�_b, 1, v_�ѱ�, �ѱ�),
        ���� = Decode(n_�������_b, 1, n_�������, ����), ���� = Decode(n_����_b, 1, v_����, ����), �Ա� = Decode(n_�Ա�_b, 1, v_�Ա�, �Ա�),
        ���� = Decode(n_����_b, 1, v_����, ����)
    Where NO = v_�Һŵ�;
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatereginfo;
/

Create Or Replace Procedure Zl_Exsesvr_Updateregroom
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����˹Һż�¼��������
  --��Σ�Json_In:��ʽ
  --input
  --  reg_no             C  1 �Һ�no
  --  pati_id            N  1 ����id
  --  outp_room          C  1 ����
  --  outpat_dr          C  1 ҽ��
  --  outpat_trg_time    C  1 ����ʱ��
  --  update_room        N  1 ��������
  --  appt_mode          C  1 ԤԼ��ʽ
  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------

  v_�Һŵ�   Varchar2(50);
  n_����id   Number;
  v_����     Varchar2(50);
  v_ҽ��     Varchar2(50);
  d_����ʱ�� Date;
  n_�������� Number;
  v_ԤԼ��ʽ Varchar2(50);

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�   := j_Json.Get_String('reg_no');
  n_����id   := j_Json.Get_Number('pati_id');
  v_����     := j_Json.Get_String('outp_room');
  v_ҽ��     := j_Json.Get_String('outpat_dr');
  d_����ʱ�� := To_Date(j_Json.Get_String('outpat_trg_time'), 'yyyy-mm-dd hh24:mi:ss');
  n_�������� := j_Json.Get_Number('update_room');
  v_ԤԼ��ʽ := j_Json.Get_String('appt_mode');

  Zl_���˹Һż�¼_��������_s(v_�Һŵ�, n_����id, v_����, v_ҽ��, d_����ʱ��, n_��������, v_ԤԼ��ʽ);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateregroom;
/

Create Or Replace Procedure Zl_Exsesvr_Getbillstatubyno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡָ�����ݵ��շѡ��쳣�����ʵ�״̬
  --��Σ�Json_In:��ʽ
  --   input      
  --    fee_no  C 1 ���ݺ�
  --    bill_prop N 1 ��¼����:1-�շѵ�;2-���ʵ�;3-�Զ����ʵ�;4-�Һŵ�;5-���￨;6-Ԥ����
  --����: Json_Out,��ʽ����
  -- output      
  --   code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  statu N 1 �շ�״̬:0-δ�շѻ򻮼�;1-���շѻ��Ѽ���;2-��ȫ�˻�ȫ����;3-�����˷ѻ򲿷�����
  --  err_sign  N 1 �쳣��־:0-��������;1-�տ���쳣;2-�˿���쳣
  --  blnc_sign N 1 ���ʱ�־:��Լ��ʵ���Ч;0-δ����;1-�Ѿ�����
  --  consumeed N 1 Ԥ���Ƿ�������:1-����������;0-δ��������
  ---------------------------------------------------------------------------
  j_Input      PLJson;
  j_Json       PLJson;
  v_���ݺ�     Varchar2(100);
  v_No         ������ü�¼.No%Type;
  n_��¼����   Number(2);
  n_��¼״̬   Number(5);
  n_����ʣ���� Number(2);
  n_�շ��쳣   Number(2);
  n_�˷��쳣   Number(2);
  n_�쳣id     Number(18);
  n_����id     Number(18);
  n_�շ�״̬   Number(2);
  n_Count      Number(2);
  n_���ʱ�־   Number(2);
  n_���ʷ���   Number(2);
  n_У�Ա�־   Number(2);
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  v_Output Varchar2(32767);

  --��װ�ɹ�ʱ���ص�����
  Function Get_Success_Message
  (
    �շ�״̬_In     Number,
    �쳣��־_In     Number,
    ���ʱ�־_In     Number,
    Ԥ�����ѱ�־_In Number
  ) Return Varchar2 Is
  
  Begin
  
    v_Output := '';
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '�ɹ�');
    zlJsonPutValue(v_Output, 'statu', �շ�״̬_In, 1); --�շ�״̬:0-δ�շѻ򻮼�;1-���շѻ��Ѽ���;2-��ȫ�˻�ȫ����;3-�����˷ѻ򲿷�����
  
    zlJsonPutValue(v_Output, 'err_sign', �쳣��־_In, 1); --�쳣��־:0-��������;1-�տ���쳣;2-�˿���쳣
    zlJsonPutValue(v_Output, 'blnc_sign', ���ʱ�־_In, 1); --���ʱ�־:��Լ��ʵ���Ч;0-δ����;1-�Ѿ�����
    zlJsonPutValue(v_Output, 'consumeed_sign', Ԥ�����ѱ�־_In, 1, 2); --Ԥ�����ѱ�־:1-����������;0-δ��������
    v_Output := '{"output":' || v_Output || '}';
  
    Return v_Output;
  End Get_Success_Message;

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ�   := j_Json.Get_String('fee_no');
  n_��¼���� := Nvl(j_Json.Get_Number('bill_prop'), 0);

  If n_��¼���� <= 0 Or v_���ݺ� Is Null Then
    v_Err_Msg := 'δ�����Ҫ�Ĳ�ѯ���������ܻ�ȡ������ص��ݵ�״̬������!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  If n_��¼���� = 1 Then
    --��ȡ�շѵ�;
    Select Max(NO), Nvl(Max(��¼״̬), 0), Max(Decode(Nvl(ʣ������, 0), 0, 0, 1)), Max(�շ��쳣), Max(�˷��쳣), Max(�쳣id)
    Into v_No, n_��¼״̬, n_����ʣ����, n_�շ��쳣, n_�˷��쳣, n_�쳣id
    From (Select NO, Max(��¼״̬) As ��¼״̬, ���, Sum(Nvl(����, 1) * Nvl(����, 0)) As ʣ������,
                  Max(Decode(a.��¼״̬, 2, a.����״̬, 0)) As �˷��쳣, Max(Decode(a.��¼״̬, 2, 0, a.����״̬)) As �շ��쳣,
                  Max(Decode(a.��¼״̬, 2, Decode(Nvl(����״̬, 0), 1, a.����id, 0), 0)) As �쳣id
           From ������ü�¼ A
           Where Mod(��¼����, 10) = 1 And NO = v_���ݺ� And �۸񸸺� Is Null
           Group By NO, ���);
  
    If v_No Is Null Then
      --δ�ҵ�����
      v_Err_Msg := 'δ�ҵ��շѵ���Ϊ' || v_���ݺ� || '���շѵ��ݣ�����������ȡ���ݵ�״̬������!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
  
    If Nvl(n_��¼״̬, 0) = 0 Then
      --���۵�:�շ�״̬_in number,�쳣��־_in number,���ʱ�־_in number,Ԥ�����ѱ�־_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
  
    If n_��¼״̬ = 1 Then
      --�Ѿ��շ�
      Json_Out := Get_Success_Message(1, n_�շ��쳣, 0, 1);
      Return;
    End If;
  
    If Nvl(n_����ʣ����, 0) = 0 Then
    
      n_�շ�״̬ := 2;
    Else
      n_�շ�״̬ := 3;
    End If;
    --�����˷�
    If Nvl(n_�˷��쳣, 0) = 1 Then
      Select Nvl(Max(1), 0) Into n_�˷��쳣 From ����Ԥ����¼ Where ����id = n_�쳣id And Nvl(У�Ա�־, 0) <> 0;
      If Nvl(n_�˷��쳣, 0) = 1 Then
        --�����쳣���϶���û��
        n_�շ�״̬ := 3;
      End If;
    End If;
    Json_Out := Get_Success_Message(n_�շ�״̬, n_�˷��쳣, 1, 0);
    Return;
  End If;

  If n_��¼���� = 2 Or n_��¼���� = 3 Then
    --2-���ʵ�;3-�Զ����ʵ�    
    Select Max(NO), Nvl(Max(��¼״̬), 0), Max(Decode(Nvl(ʣ������, 0), 0, 0, 1)), Max(����id)
    Into v_No, n_��¼״̬, n_����ʣ����, n_�쳣id
    From (Select NO, Max(��¼״̬) As ��¼״̬, ���, Sum(Nvl(����, 1) * Nvl(����, 0)) As ʣ������, Max(����id) As ����id
           From ������ü�¼ A
           Where ��¼���� = n_��¼���� And NO = v_���ݺ� And �۸񸸺� Is Null
           Group By NO, ���
           Union All
           Select NO, Max(��¼״̬) As ��¼״̬, ���, Sum(Nvl(����, 1) * Nvl(����, 0)) As ʣ������, Max(����id) As ����id
           From סԺ���ü�¼ A
           Where ��¼���� = n_��¼���� And NO = v_���ݺ� And �۸񸸺� Is Null
           Group By NO, ���
           
           );
    If v_No Is Null Then
      --δ�ҵ�����
      v_Err_Msg := 'δ�ҵ����ʵ���Ϊ' || v_���ݺ� || '�ļ��ʵ��ݣ�����������ȡ���ݵ�״̬������!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
    n_���ʱ�־ := 0;
    If Nvl(n_��¼״̬, 0) = 0 Then
      --���۵�:�շ�״̬_in number,�쳣��־_in number,���ʱ�־_in number,Ԥ�����ѱ�־_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
  
    If Nvl(n_�쳣id, 0) <> 0 Then
    
      Select Count(1)
      Into n_Count
      From ������ü�¼ A
      Where a.��¼״̬ <> 0 And a.���ʷ��� = 1 Having
       Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)) <> 0 Or
            (Sum(Nvl(a.ʵ�ս��, 0)) = 0 And Sum(Nvl(a.Ӧ�ս��, 0)) <> 0 And Sum(Nvl(a.���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or
            Sum(Nvl(a.���ʽ��, 0)) = 0 And Sum(Nvl(a.Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Group By a.No, a.���, Mod(a.��¼����, 10), Nvl(a.ִ��״̬, 0);
      If Nvl(n_Count, 0) = 0 Then
        --�Ѿ�����
        n_���ʱ�־ := 1;
      Else
        --���ֽ��ʵģ�Ҳ��δ����
        n_���ʱ�־ := 0;
      End If;
    
    End If;
    If n_��¼״̬ = 1 Then
      --�Ѿ�����
      Json_Out := Get_Success_Message(1, 0, n_���ʱ�־, 0);
      Return;
    End If;
  
    If Nvl(n_����ʣ����, 0) = 0 Then
      n_�շ�״̬ := 2;
    Else
      n_�շ�״̬ := 3;
    End If;
  
    Json_Out := Get_Success_Message(n_�շ�״̬, 0, n_���ʱ�־, 0);
    Return;
  End If;
  If n_��¼���� = 4 Then
    --�Һŵ�
    Select Max(NO), Nvl(Max(��¼״̬), 0), Max(Decode(Nvl(ʣ������, 0), 0, 0, 1)), Max(�շ��쳣), Max(�˷��쳣), Max(�쳣id)
    Into v_No, n_��¼״̬, n_����ʣ����, n_�շ��쳣, n_�˷��쳣, n_�쳣id
    From (Select NO, Max(��¼״̬) As ��¼״̬, ���, Sum(Nvl(����, 1) * Nvl(����, 0)) As ʣ������,
                  Max(Decode(a.��¼״̬, 2, a.����״̬, 0)) As �˷��쳣, Max(Decode(a.��¼״̬, 2, 0, a.����״̬)) As �շ��쳣,
                  Max(Decode(a.��¼״̬, 2, Decode(Nvl(a.����״̬, 0), 1, a.����id, 0), 0)) As �쳣id
           From ������ü�¼ A
           Where Mod(��¼����, 10) = 4 And NO = v_���ݺ� And �۸񸸺� Is Null
           Group By NO, ���);
  
    If v_No Is Null Then
      --δ�ҵ�����
      v_Err_Msg := 'δ�ҵ��Һŵ���Ϊ' || v_���ݺ� || '�ĹҺŵ��ݣ�����������ȡ���ݵ�״̬������!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
    If Nvl(n_��¼״̬, 0) = 0 Then
      --���۵�:�շ�״̬_in number,�쳣��־_in number,���ʱ�־_in number,Ԥ�����ѱ�־_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
  
    If n_��¼״̬ = 1 Then
      --�Ѿ��շ�
      Json_Out := Get_Success_Message(1, n_�շ��쳣, 0, 0);
      Return;
    End If;
  
    If Nvl(n_����ʣ����, 0) = 0 Then
    
      n_�շ�״̬ := 2;
    Else
      n_�շ�״̬ := 3;
    End If;
    --�����˷�
    If Nvl(n_�˷��쳣, 0) = 1 Then
      Select Nvl(Max(1), 0) Into n_�˷��쳣 From ����Ԥ����¼ Where ����id = n_�쳣id And Nvl(У�Ա�־, 0) <> 0;
      If Nvl(n_�˷��쳣, 0) = 1 Then
        --�����쳣���϶���û��
        n_�շ�״̬ := 3;
        n_�˷��쳣 := 2;
      End If;
    End If;
    Json_Out := Get_Success_Message(n_�շ�״̬, n_�˷��쳣, 1, 0);
    Return;
    Null;
  End If;

  If n_��¼���� = 5 Then
    --ҽ�ƿ�
    Select Max(NO), Nvl(Max(��¼״̬), 0), Max(Decode(Nvl(ʣ������, 0), 0, 0, 1)), Max(�շ��쳣), Max(�˷��쳣), Max(�쳣id), Max(���ʷ���),
           Max(����id)
    Into v_No, n_��¼״̬, n_����ʣ����, n_�շ��쳣, n_�˷��쳣, n_�쳣id, n_���ʷ���, n_����id
    From (Select NO, Max(��¼״̬) As ��¼״̬, ���, Sum(Nvl(����, 1) * Nvl(����, 0)) As ʣ������,
                  Max(Decode(a.��¼״̬, 2, a.����״̬, 0)) As �˷��쳣, Max(Decode(a.��¼״̬, 2, 0, a.����״̬)) As �շ��쳣,
                  Max(a.���ʷ���) As ���ʷ���, Max(Decode(a.��¼״̬, 2, Decode(Nvl(a.����״̬, 0), 1, a.����id, 0), 0)) As �쳣id,
                  Max(����id) As ����id
           From סԺ���ü�¼ A
           Where Mod(��¼����, 10) = 5 And NO = v_���ݺ� And �۸񸸺� Is Null
           Group By NO, ���);
  
    If v_No Is Null Then
      --δ�ҵ�����
      v_Err_Msg := 'δ�ҵ�����Ϊ' || v_���ݺ� || '�ľ��￨���ݣ�����������ȡ���ݵ�״̬������!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
  
    If Nvl(n_��¼״̬, 0) = 0 Then
      --���۵�:�շ�״̬_in number,�쳣��־_in number,���ʱ�־_in number,Ԥ�����ѱ�־_in number 
      Json_Out := Get_Success_Message(0, 0, 0, 0);
      Return;
    End If;
    n_���ʱ�־ := 0;
    If Nvl(n_���ʷ���, 0) = 1 And Nvl(n_����id, 0) <> 0 Then
      --ֻҪ�Ѿ�����ʵģ��ͷ����ѽ���
      n_���ʱ�־ := 1;
    End If;
    If n_��¼״̬ = 1 Then
      --�Ѿ��շ�
      Json_Out := Get_Success_Message(1, n_�շ��쳣, n_���ʱ�־, 0);
      Return;
    End If;
  
    If Nvl(n_����ʣ����, 0) = 0 Then
    
      n_�շ�״̬ := 2;
    Else
      n_�շ�״̬ := 3;
    End If;
  
    --�����˷�
    If Nvl(n_�˷��쳣, 0) = 1 Then
      Select Nvl(Max(1), 0) Into n_�˷��쳣 From ����Ԥ����¼ Where ����id = n_�쳣id And Nvl(У�Ա�־, 0) <> 0;
      If Nvl(n_�˷��쳣, 0) = 1 Then
        --�����쳣���϶���û��
        n_�շ�״̬ := 3;
        n_�˷��쳣 := 2;
      End If;
    End If;
    --�쳣��־:0-��������;1-�տ���쳣;2-�˿���쳣
    Json_Out := Get_Success_Message(n_�շ�״̬, n_�˷��쳣, 1, 0);
    Return;
  End If;

  If n_��¼���� = 6 Then
    --Ԥ�Լ�¼
    Select Max(NO), Nvl(Max(Decode(��¼����, 11, 0, ��¼״̬)), 0), Decode(Nvl(Sum(��Ԥ��), 0), 0, 0, 1),
           Decode(Nvl(Max(Decode(��¼����, 11, 0, У�Ա�־)), 0), 0, 0, 1)
    Into v_No, n_��¼״̬, n_Count, n_У�Ա�־
    From ����Ԥ����¼
    Where Mod(��¼����, 10) = 1 And NO = v_���ݺ�;
  
    If v_No Is Null Then
      --δ�ҵ�����
      v_Err_Msg := 'δ�ҵ�����Ϊ' || v_���ݺ� || '��Ԥ�����ݣ�����������ȡ���ݵ�״̬������!';
      Json_Out  := zlJsonOut(v_Err_Msg);
      Return;
    End If;
  
    If n_��¼״̬ = 0 Or n_��¼״̬ = 1 And Nvl(n_У�Ա�־, 0) <> 0 Then
      --δ��Ч���쳣����
      Json_Out := Get_Success_Message(0, 1, 0, n_Count);
      Return;
    End If;
  
    If n_��¼״̬ = 1 Then
      --����Ч
      Json_Out := Get_Success_Message(1, 0, 0, n_Count);
      Return;
    End If;
    n_�˷��쳣 := 0;
    If Nvl(n_У�Ա�־, 0) <> 0 Then
      n_�˷��쳣 := 2;
    End If;
    Json_Out := Get_Success_Message(2, n_�˷��쳣, 0, n_Count);
    Return;
  
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillstatubyno;
/


Create Or Replace Procedure Zl_Exsesvr_Updatecardfeeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ��������Ŷ�Ӧ���շѵ�����Ϣ(��Ҫ�Ǹ��¿��ż������ID��Ʊ����Ϣ�ȵ�)
  --���      json
  --input
  --    oper_fun  N  1  ������־:0-ֻ�޸Ŀ��Ѽ�¼;1-�޸ķ��ü�¼��Ʊ��ʹ����ϸ;2-�������޸�
  --    fee_no  C  1  ���õ��ţ�����Ҫ�����ķ��õ���
  --    operator_name  C  1  ����Ա����
  --    operator_code  C  1  ����Ա���
  --    create_time C 1 �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
  --    sendcard_info     ������Ϣ
  --      send_mode N 1 ������ʽ;0-����,1-����,2-����;3-�˿�
  --      cardtype_id C 1 �����id
  --      cardno  C 1 ����:���η��Ż�󶨻򲹿��Ŀ���
  --      recv_id N 1 ����id:Ʊ������ID(����)
  --      cardno_reusing  N 1 ��������:1-���������ظ�ʹ����;0-�������ظ�ʹ��
  --      cardno_old  C 1 ԭ������:����ʱ����Ҫ����ԭ����  --����      json
  --output
  --  code                      C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message                   C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  n_������־   Number(2);
  v_���õ���   סԺ���ü�¼.No%Type;
  n_�����id   ����Ԥ����¼.�����id%Type;
  n_������ʽ   Number(2);
  n_��������   Number(2);
  v_����Ա��� ����Ԥ����¼.����Ա���%Type;
  v_����Ա���� ����Ԥ����¼.����Ա����%Type;
  v_����       סԺ���ü�¼.ʵ��Ʊ��%Type;
  v_ԭ����     סԺ���ü�¼.ʵ��Ʊ��%Type;
  n_����id     Number(18);

  d_�Ǽ�ʱ�� Date;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_������־   := j_Json.Get_Number('oper_fun');
  v_���õ���   := j_Json.Get_String('fee_no');
  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');
  d_�Ǽ�ʱ��   := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');

  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  If Nvl(n_������־, 2) = 2 Then
    --�������޸�
    --ֻ������쳣������
    Update סԺ���ü�¼
    Set ����Ա���� = v_����Ա����, ����Ա��� = v_����Ա���, �Ǽ�ʱ�� = d_�Ǽ�ʱ��
    Where Nvl(����״̬, 0) = 1 And ��¼���� = 5 And ��¼״̬ In (1, 3) And NO = v_���õ���;
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ���Ҫ���µľ��￨�������漰�ĵ�����Ϣ';
      Raise Err_Item;
    End If;
  
    Json_Out := zlJsonOut('�ɹ�', 1);
    Return;
  End If;
  o_Json := j_Json.Get_Pljson('sendcard_info');
  If o_Json Is Null Then
    v_Err_Msg := '����ȷ�������޸ĵĿ�Ƭ��Ϣ�����飡';
    Raise Err_Item;
  End If;

  n_������ʽ := Nvl(o_Json.Get_Number('send_mode'), 0); --������ʽ;;0-����,1-����,2-����;3-�˿�
  n_�����id := o_Json.Get_Number('cardtype_id');
  n_�������� := Nvl(o_Json.Get_Number('cardno_reusing'), 0);
  v_����     := o_Json.Get_String('cardno');
  v_ԭ����   := o_Json.Get_String('cardno_old');
  n_����id   := o_Json.Get_Number('recv_id');

  If Nvl(n_������־, 0) = 1 Then
    --Ʊ��ʹ������
    Update סԺ���ü�¼
    Set ʵ��Ʊ�� = v_����, ���� = Nvl(n_�����id, ����), ���ӱ�־ = Decode(Nvl(���ӱ�־, 0), 8, 8, n_������ʽ)
    Where ��¼���� = 5 And ��¼״̬ In (1, 3) And NO = v_���õ���;
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ���Ҫ���µľ��￨�������漰�ĵ�����Ϣ';
      Raise Err_Item;
    End If;
    --��������=1-���� ;2-����;3-���� ;4-�˿�
    n_������ʽ := Case
                When Nvl(n_������ʽ, 0) = 0 Then
                 1
                When Nvl(n_������ʽ, 0) = 1 Then
                 3
                When Nvl(n_������ʽ, 0) = 2 Then
                 2
                Else
                 4
              End;
  
    Zl_����ҽ�ƿ�Ʊ��_Update_s(n_������ʽ, v_����, v_����Ա����, d_�Ǽ�ʱ��, v_���õ���, n_����id, v_ԭ����, n_��������);
  
  Else
    --ֻ������쳣������
    Update סԺ���ü�¼
    Set ʵ��Ʊ�� = v_����, ���� = Nvl(n_�����id, ����), ����Ա���� = v_����Ա����, ����Ա��� = v_����Ա���, �Ǽ�ʱ�� = d_�Ǽ�ʱ��
    Where Nvl(����״̬, 0) = 1 And ��¼���� = 5 And ��¼״̬ In (1, 3) And NO = v_���õ���;
    If Sql%NotFound Then
      v_Err_Msg := 'δ�ҵ���Ҫ���µľ��￨�������漰�ĵ�����Ϣ';
      Raise Err_Item;
    End If;
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatecardfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Updatepatibaseinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --------------------------------------------------------------------------------------------------
  --����:���²��˷�����صĲ��˻�����Ϣ
  --------------------------------------------------------------------------------------------------
  --��� JSOM��ʽ
  --input
  --  pati_id               N 1 ����id
  --  visit_id              N   ��ҳid
  --  occasion              N   ����
  --  update_info           N   ��Ҫ���µ���Ϣ
  --    pati_name             C   ����
  --    outpatient_num        C   �����
  --    pati_age              C   ����
  --    pati_sex              C 1 �Ա�
  --    explain               C 1 ˵��
  --    regist_no             C 1 �Һŵ�
  --    remark                C 1 ժҪ
  --���� JSON��ʽ
  --output
  --  code                  N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  adjust_explain        C 1 �޸�˵��
  j_Input PLJson;
  j_Json  PLJson;

  o_Json    PLJson;
  n_����id  ������ü�¼.����id%Type;
  v_����    ������ü�¼.����%Type;
  n_�����  ������ü�¼.��ʶ��%Type;
  v_�Ա�    ������ü�¼.�Ա�%Type;
  v_����    ������ü�¼.����%Type; --����ǰ������
  n_����id  Number;
  n_����    Number(1);
  v_˵��    Varchar2(4000);
  v_˵��_In Varchar2(4000);
  ˵��_Out  Clob;
  v_ժҪ    Varchar2(3682);
  v_�Һŵ�  Varchar2(100);

  Procedure p_����
  (
    ����id_In ������ü�¼.����id%Type,
    ����id_In Number,
    ����_In   ������ü�¼.����%Type,
    �Ա�_In   ������ü�¼.�Ա�%Type,
    ����_In   ������ü�¼.����%Type,
    ����_In   Number, --1-����;2-סԺ
    ˵��_In   Varchar2,
    ˵��_Out  Out Varchar2
  ) As
    ------------------------------------------------------------------------------------------
    --����:���·������ҵ�����ݵĲ��˻�����Ϣ
    --���:����id_In:����ID
    --     ����id_In:���ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;Ϊ0˵����������ǹҺž���Ĳ���(����id_InΪ��ʱ,�����ĸò��˵ķ��ò��ֵ�ҵ������)
    --     ����_In:��Ҫ���ĵĲ�������
    --     �Ա�_In:��Ҫ���ĵĲ����Ա�
    --     ����_In:��Ҫ���ĵĲ�������
    --     ����_In:1-����;2-סԺ
    --����:˵��_Out:������Ϣ�������˵����Ϣ��������ʾ����Ա������ز���
    ------------------------------------------------------------------------------------------
    Err_Custom Exception;
    v_Error   Varchar2(2000);
    v_˵��    Varchar2(4000);
    v_No      ������ü�¼.No%Type;
    n_���    Number(2);
    v_Temp    Varchar2(4000);
    n_Split   Number(2);
    d_Maxdate Date;
    v_����    ���ű�.����%Type;
  Begin
    --û��ָ���ľ���ID�������²��˵ķ���ҵ������
    If Nvl(����id_In, 0) = 0 Then
      Return;
    End If;
    v_˵�� := ˵��_In;
    If Nvl(����_In, 0) <= 1 Then
      Begin
        Select NO, ����, �Ǽ�ʱ��
        Into v_No, v_����, d_Maxdate
        From ���˹Һż�¼ A, ���ű� B
        Where a.ִ�в���id = b.Id(+) And a.Id = ����id_In;
      Exception
        When Others Then
          v_No := Null;
      End;
      If v_No Is Null Then
        v_No := '-';
      End If;
    
      n_��� := 0;
      v_Temp := Null;
      For c_���� In (Select Distinct 1 As ����, '�Һ�:' As ���, c.����
                   From ������ü�¼ A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C
                   Where a.No = b.No And b.�������� = 4 And b.Id = c.��ӡid And c.���� = 1 And a.����id = ����id_In And a.��¼���� = 4 And
                         a.No = v_No
                   Union All
                   Select Distinct 2 As ����, '�շ�:' As ���, c.����
                   From ������ü�¼ A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C
                   Where a.No = v_No And b.�������� = 1 And b.Id = c.��ӡid And c.���� = 1 And a.����id = ����id_In And a.��¼���� = 1 And
                         (a.�Һ�id = ����id_In Or ҽ����� Is Null)
                   Union All
                   Select Distinct 3 As ����, 'ҽ�ƿ�:' As ���, c.����
                   From סԺ���ü�¼ A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C
                   Where a.No = b.No And b.�������� = 5 And b.Id = c.��ӡid And c.���� = 1 And Nvl(a.���ʷ���, 0) = 0 And
                         a.����id = ����id_In And a.��¼���� = 5
                   Union All
                   Select Distinct 4 As ����, 'Ԥ��:' As ���, c.����
                   From ����Ԥ����¼ A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C
                   Where a.No = b.No And b.�������� = 2 And b.Id = c.��ӡid And c.���� = 1 And Nvl(a.Ԥ�����, 0) = 1 And
                         a.����id = ����id_In And a.��¼���� = 1
                   Union All
                   Select Distinct 5 As ����, '����:' As ���, c.����
                   From (Select Distinct b.Id, b.No
                          From ������ü�¼ A, ���˽��ʼ�¼ B
                          Where a.����id = b.Id And a.���ʷ��� = 1 And a.����id = ����id_In And a.��¼���� In (2, 12) And
                                (a.�Һ�id = ����id_In Or ҽ����� Is Null)
                          Union All
                          Select Distinct b.Id, b.No
                          From סԺ���ü�¼ A, ���˽��ʼ�¼ B
                          Where a.����id = b.Id And a.���ʷ��� = 1 And a.����id = ����id_In And a.��¼���� = 5) A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C
                   Where a.No = b.No And b.�������� = 3 And b.Id = c.��ӡid And c.���� = 1
                   Order By ����, ����) Loop
      
        If Length(Nvl(v_˵��, '-') || Nvl(v_Temp, '-')) > 3800 Then
          v_˵�� := v_˵�� || '��';
          Exit;
        End If;
      
        If n_��� <> Nvl(c_����.����, 0) Then
          If Not v_Temp Is Null Then
            v_˵�� := Nvl(v_˵��, '') || v_Temp;
          End If;
        
          n_Split := 1;
          If v_Temp Is Null Then
            v_Temp := c_����.���;
          Else
            v_Temp := ';' || c_����.���;
          End If;
          n_��� := Nvl(c_����.����, 0);
        End If;
      
        If n_Split = 1 Then
          v_Temp := Nvl(v_Temp, '') || c_����.����;
        Else
          v_Temp := Nvl(v_Temp, '') || ',' || c_����.����;
        End If;
        n_Split := 0;
      End Loop;
      If Not v_Temp Is Null Then
        If Length(Nvl(v_˵��, '-') || Nvl(v_Temp, '-')) > 4000 Then
          v_˵�� := v_˵�� || '��';
        Else
          v_˵�� := Nvl(v_˵��, '') || v_Temp;
        End If;
      
      End If;
      ˵��_Out := v_˵��;
    
      --������ID��,ֻ������ξ���Ļ�ֱ�ӵǼǵĲ�����Ϣ
      Update ������ü�¼ A
      Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
      Where ����id = ����id_In And a.��¼���� <> 4 And (a.�Һ�id = ����id_In Or ҽ����� Is Null);
    
      Update ������ü�¼ A
      Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
      Where ����id = ����id_In And a.��¼���� = 4 And NO = v_No;
    
      Update סԺ���ü�¼ A
      Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
      Where ����id = ����id_In And ��¼���� = 5;
    
      Update �ŶӽкŶ���
      Set �������� = Nvl(����_In, ��������)
      Where ����id = ����id_In And ҵ������ = 0 And ҵ��id = ����id_In;
      Return;
    End If;
  
    --סԺ:
    --1.�������Ҵ�ӡ�˷�Ʊ��,���������
    n_��� := 0;
    v_Temp := Null;
  
    For c_���� In (Select Distinct 4 As ����, 'Ԥ��:' As ���, c.����
                 From ����Ԥ����¼ A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C
                 Where a.No = b.No And b.�������� = 2 And b.Id = c.��ӡid And c.���� = 1 And Nvl(a.Ԥ�����, 0) = 2 And
                       a.����id = ����id_In And a.��¼���� = 1 And a.��ҳid = ����id_In
                 Union All
                 Select Distinct 5 As ����, '����:' As ���, c.����
                 From (Select Distinct b.Id, b.No
                        From סԺ���ü�¼ A, ���˽��ʼ�¼ B
                        Where a.����id = b.Id And Nvl(a.���ʷ���, 0) = 1 And a.����id = ����id_In And a.��ҳid = ����id_In And
                              a.��¼���� <> 5) A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C
                 Where a.No = b.No And b.�������� = 3 And b.Id = c.��ӡid And c.���� = 1
                 Order By ����, ����) Loop
    
      If Length(Nvl(v_˵��, '-') || Nvl(v_Temp, '-')) > 3800 Then
        v_˵�� := v_˵�� || '��';
        Exit;
      End If;
    
      If n_��� <> Nvl(c_����.����, 0) Then
        If Not v_Temp Is Null Then
          v_˵�� := Nvl(v_˵��, '') || v_Temp;
        End If;
        n_Split := 1;
        If v_Temp Is Null Then
          v_Temp := c_����.���;
        Else
          v_Temp := ';' || c_����.���;
        End If;
        n_��� := Nvl(c_����.����, 0);
      End If;
    
      If n_Split = 1 Then
        v_Temp := Nvl(v_Temp, '') || c_����.����;
      Else
        v_Temp := Nvl(v_Temp, '') || ',' || c_����.����;
      End If;
      n_Split := 0;
    End Loop;
    If Not v_Temp Is Null Then
      If Length(Nvl(v_˵��, '-') || Nvl(v_Temp, '-')) > 4000 Then
        v_˵�� := v_˵�� || '��';
      Else
        v_˵�� := Nvl(v_˵��, '') || v_Temp;
      End If;
    End If;
    ˵��_Out := v_˵��;
  
    Update סԺ���ü�¼
    Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
    Where ����id = ����id_In And ��ҳid = ����id_In And ��¼���� <> 5;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_����;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_����   := j_Json.Get_Number('occasion');
  n_����id := j_Json.Get_Number('visit_id');
  o_Json   := j_Json.Get_Pljson('update_info');
  If o_Json Is Null Then
    Json_Out := zlJsonOut('δ������Ҫ���µ���Ϣ�����飡', 0);
    Return;
  End If;

  v_����    := o_Json.Get_String('pati_name');
  n_�����  := To_Number(o_Json.Get_String('outpatient_num'));
  v_�Ա�    := o_Json.Get_String('pati_sex');
  v_����    := o_Json.Get_String('pati_age');
  v_�Һŵ�  := o_Json.Get_String('regist_no');
  v_˵��_In := o_Json.Get_String('explain');
  v_ժҪ    := o_Json.Get_String('remark');

  If Nvl(v_�Һŵ�, '-') <> '-' Then
    Update ������ü�¼
    Set ��ʶ�� = n_�����, ���� = v_����, �Ա� = v_�Ա�, ���� = v_����, ���� = v_ժҪ
    Where NO = v_�Һŵ� And ��¼���� = 4;
  End If;

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('δ���벡��id�����飡', 0);
    Return;
  End If;

  If Nvl(n_�����, 0) <> 0 Then
    Update ������ü�¼ Set ��ʶ�� = n_����� Where ����id = n_����id;
    Update ���˹Һż�¼ Set ����� = n_����� Where ����id = n_����id;
  End If;
  If Nvl(n_����id, 0) <> 0 Then
    p_����(n_����id, n_����id, v_����, v_�Ա�, v_����, n_����, v_˵��_In, v_˵��);
    If v_˵�� Is Not Null Then
      ˵��_Out := ˵��_Out || Chr(13) || '���ò���:' || Chr(13) || v_˵��;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","adjust_explain":"' || zlJsonStr(˵��_Out, 0) || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatepatibaseinfo;
/
Create Or Replace Procedure Zl_Exsesvr_Getorderfeeexeinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------------------------------------
  --���ܣ������סԺҽ��[ȡ��]ִ����ɼ��ͬʱ��ȡ���õ������Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --    is_finish           N 1 ִ����ɻ�ȡ����ɣ�1-ִ����ɣ�2-ȡ��ִ�����
  --    fee_origin          N 1 ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
  --    fee_nos             C 1 ���õ��ݺ�ƴ������ʽ��NO:��¼����... �磺T000001:1,T000002:2,T000003:3
  --    exe_deptid          N 1 ����ִ�п���ID,0-��ʾ�����ֿ���,����ִ�п���id
  --    order_ids           C 1 ҽ��IDs�����õ����¶�Ӧ��ҽ��id
  --    send_no             N 1 ���ͺ�
  --    wardarea_id         N 1 ���˲���id��סԺ������Ҫ����
  --    order_status        N 1 ҽ��ִ��״̬��ȡ��ִ�����ʱ�д˽��
  --����: Json_Out,��ʽ����
  --  output
  --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fee_ids               C 1 ��Ҫִ����ɵģ�����ids����ͨ��ҩƷ���ķ�����ϸid����ȡ�����ʱ�д˽��
  --    stuffdtl_ids          C 1 ������ϸid,���ŷָ�����Զ���[��]����
  --    rcpdtl_ids            C 1 ҩƷ������ϸid��סԺ���ܻ��У������Զ���[��]ҩƷ
  --    finish_list[]��Ҫִ������Զ���˵ķ�����ϸ�����ʻ��ۣ���ȡ�����ʱ�д˽���б�
  --         pati_id          N 1 ����id
  --         fee_id           N 1 ����id
  --         fee_no           C 1 ���õ��ݺ�
  --         serial_num       N 1 �������
  --         exe_deptid       N 1 ִ�п���id
  --         fee_type         N 1 �շ����ͣ�0-��ͨ���ã�1-ҩƷ�ѣ�2-�����������ķ�
  --    order_list[]����ҽ�����ʹ��ǣ���ȡ�����ʱ�д˽���б�
  --         order_id         N 1 ҽ��ID
  --         send_no          N 1 ���ͺ�
  --         type             N 1 �������ʱ�����ͣ����ڴ�꣬0-ҩƷ�ѣ�1-�����������ķ�
  --    cancel_list[]ȡ��ִ�����ʱ���صķ���״̬�����б���ȡ��ִ�����ʱ�д˽���б�
  --        fee_id            N 1 ����id
  --        exe_status        N 1 ִ��״̬
  --        exe_people        C 1 ִ����
  --        exe_time          C  ִ��ʱ��
  --------------------------------------------------------------------------------------------------------
  n_��˷���      Number(1);
  v_ҽ��ids       Varchar2(32767);
  v_Nos           Varchar2(32767);
  n_ִ�в���id    Number(18);
  v_ִ��ǰ�Ƚ���  Varchar2(4000);
  v_Fee_Item      Varchar2(32767);
  v_Fee_Item_List Varchar2(32767);
  v_Affirm        Varchar2(3000);

  n_Finish    Number; --1-ִ����ɣ�2-ȡ��ִ�����
  n_Origin    Number; --1-������ã�2-סԺ����
  j_Json      Pljson;
  j_Tmp       Pljson;
  j_Output    Pljson;
  v_Jtmp      Varchar2(32767);
  v_Jtmp1     Varchar2(32767);
  v_Jtmp2     Varchar2(32767);
  v_Pati_List Varchar2(32767);
  v_Json_In   Varchar2(32767);
  v_Json_Out  Varchar2(32767);
  v_No        סԺ���ü�¼.No%Type;
  v_���      Varchar2(32767);
  v_Error     Varchar2(32767);
  Err_Custom Exception;

  Procedure p_Checkinfinish As
    ------------------------------------------------------------------------------------------------------
    --���ܣ�סԺҽ��ִ����ɼ��
    --��Σ�Json_In:��ʽ
    --input
    --    fee_nos             C 1 ���õ��ݺ�ƴ������ʽ��NO:��¼����... �磺T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 ִ�п���ID,0-��ʾ�����ֿ���,����ִ�п���id
    --    order_ids           C 1 ҽ��IDs
    --    send_no             N 1 ���ͺ�
    --    wardarea_id         N 1 ����id
    --����: Json_Out,��ʽ����
    --  output
    --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    --    fee_ids               C 1 ��Ҫִ����ɵģ�����ids
    --    stuffdtl_ids          C 1 ������ϸid,���ŷָ�
    --    rcpdtl_ids            C 1 ҩƷ������ϸid
    --    item_list
    --         pati_id          N 1 ����id
    --         pati_pageid      N 1 ��ҳid
    --         fee_id           N 1 ����id
    --         fee_no           C 1 ���õ��ݺ�
    --         serial_num       N 1 �������
    --         exe_deptid       N 1 ִ�п���id
    --         fee_type         N 1 �շ����ͣ�0-��ͨ���ã�1-ҩƷ�ѣ�2-�����������ķ�
    --    order_list
    --         order_id         N 1 ҽ��ID
    --         send_no          N 1 ���ͺ�
    --         type             N 1 �������ʱ�����ͣ����ڴ�꣬0-ҩƷ�ѣ�1-�����������ķ�
    --------------------------------------------------------------------------------------------------------
    v_������ϸids  Varchar2(32767);
    v_ҩƷ��ϸids  Varchar2(32767);
    v_Nos          Varchar2(32767);
    v_Orders       Varchar2(32767);
    n_ִ�в���id   Number;
    v_סԺ�Զ����� Varchar2(300);
    --��ͨ����
    Cursor c_Finish Is
      Select a.Id
      From (Select Distinct a.Id, a.�շ����, a.�շ�ϸĿid
             From סԺ���ü�¼ A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As ��¼����
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And a.��¼���� = n.��¼���� And
                   a.��¼״̬ In (0, 1, 3) And a.ִ��״̬ <> 1 And (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id)) A, �������� B
      Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);
  
    --ִ���а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
    --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
    Cursor c_Stuff Is
      Select a.Id, a.ҽ����� As ҽ��id
      From סԺ���ü�¼ A, �������� D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As ��¼����
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where d.����id = a.�շ�ϸĿid And a.�շ���� = '4' And d.�������� = 1 And a.��¼״̬ = 1 And
            Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And
            (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id Or v_סԺ�Զ����� = '1');
  
    --δ��˵ķ�����(����ҩƷ������)
    Cursor c_Verify Is
      Select a.Id, a.No, a.���, a.ҽ����� As ҽ��id, a.ִ�в���id, a.�շ����, a.�շ�ϸĿid, b.��������, a.����id, a.��ҳid
      From סԺ���ü�¼ A, �������� B,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As ��¼����
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where a.���ʷ��� = 1 And a.��¼״̬ = 0 And a.�۸񸸺� Is Null And Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And
            a.No = n.No And a.��¼���� = n.��¼���� And (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id) And a.�շ�ϸĿid = b.����id(+)
      Order By a.No, a.���;
    v_Feeids   Varchar2(32767);
    n_���ͺ�   Number;
    n_Cnt      Number;
    n_״̬     Number;
    n_��˱�־ Number;
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_���ͺ�     := j_Json.Get_Number('send_no');
    n_ִ�в���id := j_Json.Get_Number('exe_deptid');
    n_ִ�в���id := Nvl(n_ִ�в���id, 0);
    n_��˱�־   := j_Json.Get_Number('fee_audit_status');
    n_״̬       := j_Json.Get_Number('si_inp_status');
  
    Select zl_GetSysParameter(63) Into v_סԺ�Զ����� From Dual;
    For R In c_Finish Loop
      v_Feeids := v_Feeids || ',' || r.Id;
    End Loop;
  
    v_Jtmp  := Null;
    v_Jtmp1 := Null;
    v_Jtmp2 := Null;
    --ִ��ʱ�Զ���˶�Ӧ�ļ��ʻ��۵�����
    --����ҽ����Ӧ��ҩƷ�����ķ��ã���Ϊҽ����ִ�У�����Ӧ����Ч
    For r_Verify In c_Verify Loop
      n_Cnt := 0;
      If r_Verify.�շ���� = '4' And r_Verify.�������� = 1 Then
        n_Cnt := 2;
      Elsif r_Verify.�շ���� In ('5', '6', '7') Then
        n_Cnt := 1;
      End If;
      If n_Cnt <> 0 Then
        v_Jtmp := v_Jtmp || ',{"order_id":' || r_Verify.ҽ��id;
        v_Jtmp := v_Jtmp || ',"send_no":' || n_���ͺ�;
        v_Jtmp := v_Jtmp || ',"type":' || (n_Cnt - 1);
        v_Jtmp := v_Jtmp || '}';
      End If;
    
      v_Jtmp2 := v_Jtmp2 || ',{"pati_id":' || r_Verify.����id;
      v_Jtmp2 := v_Jtmp2 || ',"pati_pageid":' || r_Verify.��ҳid;
      v_Jtmp2 := v_Jtmp2 || ',"fee_id":' || r_Verify.Id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_no":"' || r_Verify.No || '"';
      v_Jtmp2 := v_Jtmp2 || ',"serial_num":' || r_Verify.���;
      v_Jtmp2 := v_Jtmp2 || ',"exe_deptid":' || r_Verify.ִ�в���id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_type":' || n_Cnt; --N  1-��ʾҩƷ��2-��ʾ����
      v_Jtmp2 := v_Jtmp2 || '}';
    
      If r_Verify.�շ���� = '4' And r_Verify.�������� = 1 And (n_ִ�в���id = 0 Or r_Verify.ִ�в���id = n_ִ�в���id Or v_סԺ�Զ����� = '1') Then
        v_������ϸids := v_������ϸids || ',' || r_Verify.Id;
        v_Feeids      := v_Feeids || ',' || r_Verify.Id;
      End If;
      If r_Verify.�շ���� In ('5', '6', '7') And r_Verify.ִ�в���id = n_ִ�в���id Then
        v_ҩƷ��ϸids := v_ҩƷ��ϸids || ',' || r_Verify.Id;
        v_Feeids      := v_Feeids || ',' || r_Verify.Id;
      End If;
    
      If v_Pati_List Is Null Then
        v_Pati_List := '{"pati_id":' || r_Verify.����id;
        v_Pati_List := v_Pati_List || ',"fee_audit_status":' || Nvl(n_��˱�־, 0);
        v_Pati_List := v_Pati_List || ',"si_inp_status":' || Nvl(n_״̬, 0);
        v_Pati_List := v_Pati_List || '}';
        v_Pati_List := ',"pati_list":[' || v_Pati_List || ']';
      End If;
    
      If r_Verify.No <> v_No And v_��� Is Not Null Then
        v_Json_In := '{"fee_nos":"' || v_No || '"';
        v_Json_In := v_Json_In || ',"":"' || v_��� || '"';
        v_Json_In := v_Json_In || v_Pati_List;
        v_Json_In := v_Json_In || '}';
        v_Json_In := '{"input":' || v_Json_In || '}';
        Zl_סԺ���ʼ�¼_Verify_Check(v_Json_In, v_Json_Out);
        v_���   := Null;
        j_Tmp    := Pljson();
        j_Output := Pljson();
        j_Tmp    := Pljson(v_Json_Out);
        j_Output := j_Tmp.Get_Pljson('output');
        If j_Output.Get_Number('code') = 0 Then
          v_Error := j_Output.Get_String('message');
          Raise Err_Custom;
        End If;
      End If;
      v_No   := r_Verify.No;
      v_��� := v_��� || ',' || r_Verify.���;
    End Loop;
  
    If v_��� Is Not Null Then
      v_Json_In := '{"fee_nos":"' || v_No || '"';
      v_Json_In := v_Json_In || ',"":"' || v_��� || '"';
      v_Json_In := v_Json_In || v_Pati_List;
      v_Json_In := v_Json_In || '}';
      v_Json_In := '{"input":' || v_Json_In || '}';
      Zl_סԺ���ʼ�¼_Verify_Check(v_Json_In, v_Json_Out);
      j_Tmp    := Pljson();
      j_Output := Pljson();
      j_Tmp    := Pljson(v_Json_Out);
      j_Output := j_Tmp.Get_Pljson('output');
      If j_Output.Get_Number('code') = 0 Then
        v_Error := j_Output.Get_String('message');
        Raise Err_Custom;
      End If;
    End If;
  
    --����������������Զ�����
    For r_Stuff In c_Stuff Loop
      --��Ҫ�����ĵ���ϸ
      v_������ϸids := v_������ϸids || ',' || r_Stuff.Id;
    End Loop;
  
    --����������������Զ�����
    --���ݴ���Ĳ���id��ȷ���Ƿ���Ҫ�Զ���ҩ
    For r_Drug In (Select a.Id
                   From סԺ���ü�¼ A,
                        (Select /*+cardinality(f,10)*/
                           C1 As NO, To_Number(C2) As ��¼����
                          From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
                   Where a.�շ���� In ('5', '6', '7') And a.��¼״̬ = 1 And
                         Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And a.��¼���� = n.��¼���� And
                         n_ִ�в���id = a.ִ�в���id) Loop
      --��Ҫ��ҩƷ����ϸ
      v_ҩƷ��ϸids := v_ҩƷ��ϸids || ',' || r_Drug.Id;
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"�ɹ�"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_������ϸids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"rcpdtl_ids":"' || Substr(v_ҩƷ��ϸids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"fee_ids":"' || Substr(v_Feeids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"finish_list":[' || Substr(v_Jtmp2, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || ',"order_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  End p_Checkinfinish;

  Procedure p_Checkoutfinish As
    ---------------------------------------------------------------------------
    --���ܣ�����ҽ��ִ����ɼ��
    --��Σ�Json_In:��ʽ
    --input
    --    fee_nos             C 1 ���õ��ݺ�ƴ������ʽ��NO:��¼����... �磺T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 ִ�п���ID,0-��ʾ�����ֿ���,����ִ�п���id
    --    order_ids           C 1 ҽ��IDs
    --    send_no             N 1 ���ͺ�
    --����: Json_Out,��ʽ����
    --  output
    --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    --    fee_ids               C 1 ��Ҫִ����ɵģ�����ids
    --    stuffdtl_ids          C 1 ������ϸid,���ŷָ�
    --    item_list
    --         pati_id          N 1 ����id
    --         fee_id           N 1 ����id
    --         fee_no           C 1 ���õ��ݺ�
    --         serial_num       N 1 �������
    --         exe_deptid       N 1 ִ�п���id
    --         fee_type         N 1 �շ����ͣ�0-��ͨ���ã�1-ҩƷ�ѣ�2-�����������ķ�
    --    order_list
    --         order_id         N 1 ҽ��ID
    --         send_no          N 1 ���ͺ�
    --         type             N 1 �������ʱ�����ͣ����ڴ�꣬0-ҩƷ�ѣ�1-�����������ķ�
    --------------------------------------------------------------------------------
    n_���ͺ�       Number;
    v_Orders       Varchar2(32767);
    n_ִ�в���id   ������ü�¼.ִ�в���id%Type;
    n_Cnt          Number;
    v_Error        Varchar2(2000);
    v_�����Զ����� Varchar2(300);
    Err_Custom Exception;
    v_ִ��ǰ�Ƚ��� Varchar2(500);
    v_������ϸids  Varchar2(32767);
    v_Nos          Varchar2(32767);
    v_Feeids       Varchar2(32767);
    Cursor c_Finishone Is
      Select a.Id, a.ҽ����� As ҽ��id
      From (Select a.Id, a.�շ����, a.�շ�ϸĿid, a.ҽ�����
             From ������ü�¼ A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As ��¼����
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And
                   a.��¼״̬ In (0, 1, 3) And a.ִ��״̬ <> 1 And (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id)) A, �������� B
      Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);
  
    --ִ���а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
    --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
    Cursor c_Stuff Is
      Select a.Id, a.ҽ����� As ҽ��id
      From ������ü�¼ A, �������� D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As ��¼����
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where d.����id = a.�շ�ϸĿid And a.�շ���� = '4' And d.�������� = 1 And a.��¼״̬ = 1 And
            Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And
            (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id Or v_�����Զ����� = '1');
  
    --δ��˵ķ�����(����ҩƷ������)
    Cursor c_Verifyone(P���ʷ��� Number) Is
      Select a.Id, a.No, a.ҽ����� As ҽ��id, a.���, a.ִ�в���id, a.�շ����, a.�շ�ϸĿid, b.��������, a.����id
      From ������ü�¼ A, �������� B,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As ��¼����
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where Nvl(a.���ʷ���, 0) = P���ʷ��� And a.��¼״̬ = 0 And a.�۸񸸺� Is Null And
            Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And Mod(a.��¼����, 10) = n.��¼���� And
            (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id) And a.�շ�ϸĿid = b.����id(+)
      Order By NO, ���;
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_���ͺ�     := j_Json.Get_Number('send_no');
    n_ִ�в���id := j_Json.Get_Number('exe_deptid');
    n_ִ�в���id := Nvl(n_ִ�в���id, 0);
    Select zl_GetSysParameter(92) Into v_�����Զ����� From Dual;
    For R In c_Finishone Loop
      v_Feeids := v_Feeids || ',' || r.Id;
    End Loop;
    v_Feeids := Substr(v_Feeids, 2);
    Select Count(1)
    Into n_Cnt
    From ������ü�¼ A
    Where a.Id In (Select /*+cardinality(x,10)*/
                    x.Column_Value
                   From Table(Cast(f_Num2list(v_Feeids) As Zltools.t_Numlist)) X) And a.����״̬ = 1 And
          Nvl(a.����id, 0) <> 0;
    If n_Cnt > 0 Then
      v_Error := '��ǰִ�е�ҽ����Ӧ�ķ��õ����д����쳣���ݡ�';
      Raise Err_Custom;
    End If;
    Select zl_GetSysParameter(163) Into v_ִ��ǰ�Ƚ��� From Dual;
    --ִ��ʱ�Զ���˶�Ӧ�ļ��ʻ��۵�����
    --����ҽ����Ӧ��ҩƷ�����ķ��ã���Ϊҽ����ִ�У�����Ӧ����Ч
    If Nvl(v_ִ��ǰ�Ƚ���, '0') <> '0' Then
      For r_Verify In c_Verifyone(0) Loop
        v_Error := '��ǰִ�е�ҽ��������δ��ȡ�ķ��á�';
        Raise Err_Custom;
      End Loop;
    End If;
    v_Jtmp  := Null;
    v_Jtmp1 := Null;
    v_Jtmp2 := Null;
    For r_Verify In c_Verifyone(1) Loop
      n_Cnt := 0;
      If r_Verify.�շ���� = '4' And r_Verify.�������� = 1 Then
        n_Cnt := 2;
      Elsif r_Verify.�շ���� In ('5', '6', '7') Then
        n_Cnt := 1;
      End If;
    
      If n_Cnt <> 0 Then
        v_Jtmp := v_Jtmp || ',{"order_id":' || r_Verify.ҽ��id;
        v_Jtmp := v_Jtmp || ',"send_no":' || n_���ͺ�;
        v_Jtmp := v_Jtmp || ',"type":' || (n_Cnt - 1);
        v_Jtmp := v_Jtmp || '}';
      End If;
    
      v_Jtmp2 := v_Jtmp2 || ',{"pati_id":' || r_Verify.����id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_id":' || r_Verify.Id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_no":"' || r_Verify.No || '"';
      v_Jtmp2 := v_Jtmp2 || ',"serial_num":' || r_Verify.���;
      v_Jtmp2 := v_Jtmp2 || ',"exe_deptid":' || r_Verify.ִ�в���id;
      v_Jtmp2 := v_Jtmp2 || ',"fee_type":' || n_Cnt; --N  1-��ʾҩƷ��2-��ʾ����
      v_Jtmp2 := v_Jtmp2 || '}';
    
      If r_Verify.�շ���� = '4' And r_Verify.�������� = 1 And (n_ִ�в���id = 0 Or r_Verify.ִ�в���id = n_ִ�в���id Or v_�����Զ����� = '1') Then
        v_������ϸids := v_������ϸids || ',' || r_Verify.Id;
      End If;
    End Loop;
    --����������������Զ�����
    For r_Stuff In c_Stuff Loop
      --��Ҫ�����ĵ���ϸ
      v_������ϸids := v_������ϸids || ',' || r_Stuff.Id;
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"�ɹ�"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_������ϸids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"fee_ids":"' || v_Feeids || '"';
    v_Jtmp1 := v_Jtmp1 || ',"finish_list":[' || Substr(v_Jtmp2, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || ',"order_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  Exception
    When Err_Custom Then
      Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Checkoutfinish;

  Procedure p_Checkincancel As
    ---------------------------------------------------------------------------
    --���ܣ�סԺҽ��ȡ��ִ����ɼ��
    --��Σ�Json_In:��ʽ
    --input
    --    fee_nos             C 1 ���õ��ݺ�ƴ������ʽ��NO:��¼����... �磺T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 ִ�п���ID,0-��ʾ�����ֿ���,������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
    --    order_ids           C 1 ҽ��IDs
    --    order_status        N 1 ҽ��ִ��״̬
    --    wardarea_id         N 1 ����id
    --����: Json_Out,��ʽ����
    --  output
    --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    --    stuffdtl_ids          C 1 ������ϸid,���ŷָ�
    --    rcpdtl_ids            C 1 ������ϸid,���ŷָ�
    --    item_list       Ҫ���µķ�����ϸ
    --        fee_id            N 1 ����id
    --        exe_status        N 1 ִ��״̬
    --        exe_people        C 1 ִ����
    --        exe_time          C  ִ��ʱ��
    ---------------------------------------------------------------------------
  
    v_Orders      Varchar2(32767);
    n_ִ�в���id  ������ü�¼.ִ�в���id%Type;
    v_������ϸids Varchar2(32767);
    v_Nos         Varchar2(32767);
    v_ҩƷ��ϸids Varchar2(32767);
    n_����id      Number;
    d_ִ��ʱ��    Date;
    v_ִ����      Varchar2(100);
    n_ִ��״̬    Number;
    --Ҫȡ��ִ�еķ�����(������ҩƷ�͸������õ�����)
    Cursor c_Finishone Is
      Select a.Id, a.ִ��ʱ��, a.ִ����
      From (Select a.Id, a.�շ����, a.�շ�ϸĿid, a.ִ��ʱ��, a.ִ����
             From סԺ���ü�¼ A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As ��¼����
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And
                   (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And a.��¼״̬ In (0, 1, 3) And
                   (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id)) A, �������� B
      Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);
  
    --ȡ��ִ���а����������õķ�������ʱ�����ݲ��������Ƿ��Զ�����
    --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
    Cursor c_Stuff Is
      Select a.Id
      From סԺ���ü�¼ A, �������� D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As ��¼����
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where a.�շ���� = '4' And a.��¼״̬ = 1 And a.�շ�ϸĿid = d.����id And d.�������� = 1 And
            Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And
            (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id);
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_ִ�в���id := j_Json.Get_Number('exe_deptid');
    n_ִ�в���id := Nvl(n_ִ�в���id, 0);
    n_ִ��״̬   := j_Json.Get_Number('order_status');
    v_Jtmp       := Null;
    v_Jtmp1      := Null;
    For R In c_Finishone Loop
      Select r.Id As ����id, Decode(n_ִ��״̬, 0, d_ִ��ʱ��, r.ִ��ʱ��) As ִ��ʱ��, Decode(n_ִ��״̬, 0, Null, r.ִ����) As ִ����
      Into n_����id, d_ִ��ʱ��, v_ִ����
      From Dual;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || n_����id;
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(n_ִ��״̬, 0);
      v_Jtmp := v_Jtmp || ',"exe_people":"' || Zljsonstr(v_ִ����) || '"';
      v_Jtmp := v_Jtmp || ',"exe_time":"' || To_Char(d_ִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
      v_Jtmp := v_Jtmp || ',"fee_type":0';
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  
    --����������������Զ�����
    For r_Stuff In c_Stuff Loop
      --��Ҫ�˵�������ϸ
      v_������ϸids := v_������ϸids || ',' || r_Stuff.Id;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || r_Stuff.Id;
      v_Jtmp := v_Jtmp || ',"exe_status":0';
      v_Jtmp := v_Jtmp || ',"exe_people":""';
      v_Jtmp := v_Jtmp || ',"exe_time":""';
      v_Jtmp := v_Jtmp || ',"fee_type":1';
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  
    --���ݴ���Ĳ���id��ȷ���Ƿ���Ҫ����ҩ
    For r_Drug In (Select a.Id
                   From סԺ���ü�¼ A,
                        (Select /*+cardinality(f,10)*/
                           C1 As NO, To_Number(C2) As ��¼����
                          From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
                   Where a.�շ���� In ('5', '6', '7') And a.��¼״̬ = 1 And
                         Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And a.��¼���� = n.��¼���� And
                         n_ִ�в���id = a.ִ�в���id) Loop
      --��Ҫ��ҩƷ����ϸ
      v_ҩƷ��ϸids := v_ҩƷ��ϸids || ',' || r_Drug.Id;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || r_Drug.Id;
      v_Jtmp := v_Jtmp || ',"exe_status":0';
      v_Jtmp := v_Jtmp || ',"exe_people":""';
      v_Jtmp := v_Jtmp || ',"exe_time":""';
      v_Jtmp := v_Jtmp || ',"fee_type":2';
      v_Jtmp := v_Jtmp || '}';
    
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"�ɹ�"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_������ϸids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"rcpdtl_ids":"' || Substr(v_ҩƷ��ϸids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"cancel_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Checkincancel;

  Procedure p_Checkoutcancel As
    ---------------------------------------------------------------------------
    --���ܣ�����ҽ��ȡ��ִ����ɼ��
    --��Σ�Json_In:��ʽ
    --input
    --    fee_nos             C 1 ���õ��ݺ�ƴ������ʽ��NO:��¼����... �磺T000001:1,T000002:2,T000003:3
    --    exe_deptid          N 1 ִ�п���ID,0-��ʾ�����ֿ���,������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в���
    --    order_ids           C 1 ҽ��IDs
    --    order_status        N 1 ҽ��ִ��״̬
    --����: Json_Out,��ʽ����
    --  output
    --    code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
    --    message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
    --    stuffdtl_ids          C 1 ������ϸid,���ŷָ�
    --    item_list       Ҫ���µķ�����ϸ
    --        fee_id            N 1 ����id
    --        exe_status        N 1 ִ��״̬
    --        exe_people        C 1 ִ����
    --        exe_time          C  ִ��ʱ��
    ---------------------------------------------------------------------------
  
    v_Orders      Varchar2(32767);
    n_ִ�в���id  ������ü�¼.ִ�в���id%Type;
    v_������ϸids Varchar2(32767);
    v_Nos         Varchar2(32767);
    n_����id      Number;
    d_ִ��ʱ��    Date;
    v_ִ����      Varchar2(100);
    n_ִ��״̬    Number;
  
    --Ҫȡ��ִ�еķ�����(������ҩƷ�͸������õ�����)
    Cursor c_Finishone Is
      Select a.Id, a.ִ��ʱ��, a.ִ����
      From (Select a.Id, a.�շ����, a.�շ�ϸĿid, a.ִ��ʱ��, a.ִ����
             From ������ü�¼ A,
                  (Select /*+cardinality(f,10)*/
                     C1 As NO, To_Number(C2) As ��¼����
                    From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
             Where Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And
                   (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And a.��¼״̬ In (0, 1, 3) And
                   (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id)) A, �������� B
      Where a.�շ�ϸĿid = b.����id(+) And Not a.�շ���� In ('5', '6', '7') And Not (a.�շ���� = '4' And Nvl(b.��������, 0) = 1);
  
    --ȡ��ִ���а����������õķ�������ʱ�����ݲ��������Ƿ��Զ�����
    --��������ҽ��Ŀǰ�����ڵ��������ִ�е����
    Cursor c_Stuff Is
      Select a.Id
      From ������ü�¼ A, �������� D,
           (Select /*+cardinality(f,10)*/
              C1 As NO, To_Number(C2) As ��¼����
             From Table(Cast(f_Str2list2(v_Nos) As t_Strlist2)) F) N
      Where a.�շ���� = '4' And a.��¼״̬ = 1 And a.�շ�ϸĿid = d.����id And d.�������� = 1 And
            Instr(',' || v_Orders || ',', ',' || a.ҽ����� || ',') > 0 And a.No = n.No And
            (a.��¼���� = n.��¼���� Or a.��¼���� = 11 And n.��¼���� = 1) And (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id);
  Begin
    v_Nos        := j_Json.Get_String('fee_nos');
    v_Orders     := j_Json.Get_String('order_ids');
    n_ִ�в���id := j_Json.Get_Number('exe_deptid');
    n_ִ�в���id := Nvl(n_ִ�в���id, 0);
    n_ִ��״̬   := j_Json.Get_Number('order_status');
    v_Jtmp       := Null;
    v_Jtmp1      := Null;
    For R In c_Finishone Loop
      Select r.Id As ����id, Decode(n_ִ��״̬, 0, d_ִ��ʱ��, r.ִ��ʱ��) As ִ��ʱ��, Decode(n_ִ��״̬, 0, Null, r.ִ����) As ִ����
      Into n_����id, d_ִ��ʱ��, v_ִ����
      From Dual;
      v_Jtmp := v_Jtmp || ',{"fee_id":' || n_����id;
      v_Jtmp := v_Jtmp || ',"exe_status":' || Nvl(n_ִ��״̬, 0);
      v_Jtmp := v_Jtmp || ',"exe_people":"' || Zljsonstr(v_ִ����) || '"';
      v_Jtmp := v_Jtmp || ',"exe_time":"' || To_Char(d_ִ��ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
      v_Jtmp := v_Jtmp || ',"fee_type":0';
      v_Jtmp := v_Jtmp || '}';
    End Loop;
    --����������������Զ�����
    For r_Stuff In c_Stuff Loop
      --��Ҫ�˵�������ϸ
      v_������ϸids := v_������ϸids || ',' || r_Stuff.Id;
    
      v_Jtmp := v_Jtmp || ',{"fee_id":' || r_Stuff.Id;
      v_Jtmp := v_Jtmp || ',"exe_status":' || 0;
      v_Jtmp := v_Jtmp || ',"exe_people":""';
      v_Jtmp := v_Jtmp || ',"exe_time":""';
      v_Jtmp := v_Jtmp || ',"fee_type":1';
      v_Jtmp := v_Jtmp || '}';
    End Loop;
  
    v_Jtmp1 := v_Jtmp1 || '{"code":1';
    v_Jtmp1 := v_Jtmp1 || ',"message":"�ɹ�"';
    v_Jtmp1 := v_Jtmp1 || ',"stuffdtl_ids":"' || Substr(v_������ϸids, 2) || '"';
    v_Jtmp1 := v_Jtmp1 || ',"cancel_list":[' || Substr(v_Jtmp, 2) || ']';
    v_Jtmp1 := v_Jtmp1 || '}';
  
    Json_Out := '{"output":' || v_Jtmp1 || '}';
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Checkoutcancel;

  Procedure p_���ִ��(Json_Out Out Varchar2) As
    n_�Զ����� Number(1);
    Cursor c_Fee Is
      Select a.��Դ, a.Id, a.No, a.���, a.��¼����, a.��¼״̬, a.ִ��״̬, a.�շ�ϸĿid, a.�շ����, a.ִ�в���id, a.����״̬, a.����id, a.���ʷ���, a.�۸񸸺�,
             a.����id, a.ҽ��id, b.��������
      From (Select 2 ��Դ, a.Id, a.No, a.���, a.��¼����, a.��¼״̬, a.ִ��״̬, a.�շ�ϸĿid, a.�շ����, a.ִ�в���id, a.����״̬, a.����id, a.���ʷ���,
                    a.�۸񸸺�, a.����id, a.ҽ����� ҽ��id
             From סԺ���ü�¼ A
             Where a.No In (Select /*+cardinality(X,10)*/
                             x.Column_Value NO
                            From Table(Cast(f_Str2list(v_Nos) As Zltools.t_Strlist)) X) And
                   Instr(',' || v_ҽ��ids || ',', ',' || a.ҽ����� || ',') > 0 And (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id)
             Union All
             Select 1 ��Դ, a.Id, a.No, a.���, a.��¼����, a.��¼״̬, a.ִ��״̬, a.�շ�ϸĿid, a.�շ����, a.ִ�в���id, a.����״̬, a.����id, a.���ʷ���,
                    a.�۸񸸺�, a.����id, a.ҽ����� ҽ��id
             From ������ü�¼ A
             Where a.No In (Select /*+cardinality(X,10)*/
                             x.Column_Value NO
                            From Table(Cast(f_Str2list(v_Nos) As Zltools.t_Strlist)) X) And
                   Instr(',' || v_ҽ��ids || ',', ',' || a.ҽ����� || ',') > 0 And (n_ִ�в���id = 0 Or a.ִ�в���id = n_ִ�в���id)) A,
           �������� B
      Where a.�շ�ϸĿid = b.����id(+)
      Order By a.No, a.�շ�ϸĿid;
  Begin
    Select zl_GetSysParameter(163) Into v_ִ��ǰ�Ƚ��� From Dual;
    v_ִ��ǰ�Ƚ��� := Nvl(v_ִ��ǰ�Ƚ���, '0');
  
    For r_���� In c_Fee Loop
      If r_����.��Դ = 1 Then
        --����������м��
        If r_����.����״̬ = 1 And Nvl(r_����.����id, 0) <> 0 Then
          v_Error := '��ǰִ�е�ҽ����Ӧ�ķ��õ����д����쳣���ݡ�';
          Raise Err_Custom;
        End If;
        If v_ִ��ǰ�Ƚ��� <> '0' Then
          If Nvl(r_����.���ʷ���, 0) = 0 And r_����.��¼״̬ = 0 And r_����.�۸񸸺� Is Null Then
            v_Error := '��ǰִ�е�ҽ��������δ��ȡ�ķ��á�';
            Raise Err_Custom;
          End If;
        End If;
      End If;
      v_Fee_Item := Null;
      n_�Զ����� := 0;
      n_��˷��� := 0;
      --�������ϸ
      If Nvl(r_����.���ʷ���, 0) = 1 And r_����.��¼״̬ = 0 And r_����.�۸񸸺� Is Null Then
        n_��˷��� := 1;
        v_Fee_Item := v_Fee_Item || '{"fee_origin":' || r_����.��Դ;
        v_Fee_Item := v_Fee_Item || ',"fee_id":' || r_����.Id;
        v_Fee_Item := v_Fee_Item || ',"fee_no":"' || r_����.No || '"';
        v_Fee_Item := v_Fee_Item || ',"bill_prop":' || r_����.��¼����;
        v_Fee_Item := v_Fee_Item || ',"rec_state":' || r_����.��¼״̬;
        v_Fee_Item := v_Fee_Item || ',"serial_num":' || r_����.���;
        v_Fee_Item := v_Fee_Item || ',"fee_type":"' || r_����.�շ���� || '"';
        v_Fee_Item := v_Fee_Item || ',"fee_item_id":' || r_����.�շ�ϸĿid;
        v_Fee_Item := v_Fee_Item || ',"order_id":' || r_����.ҽ��id;
        v_Fee_Item := v_Fee_Item || ',"stuff_used":' || Nvl(r_����.��������, 0);
        v_Fee_Item := v_Fee_Item || ',"exe_dept_id":' || Nvl(r_����.ִ�в���id, 0); --ִ�в���id
        v_Fee_Item := v_Fee_Item || ',"is_verify":' || n_��˷���;
      End If;
    
      --ִ�������ϸ
      If r_����.��¼״̬ In (0, 1, 3) And r_����.ִ��״̬ <> 1 Then
        If r_����.��¼״̬ In (0, 1) And r_����.�������� = 1 Then
          If r_����.��Դ = 1 And r_����.��¼���� In (1, 11) Or r_����.��Դ = 2 And r_����.��¼���� = 2 Then
            --�����շѼ�¼�Զ����ϻ�סԺ���ʼ�¼
            n_�Զ����� := 1;
          End If;
        End If;
      
        If v_Fee_Item Is Null Then
          v_Fee_Item := v_Fee_Item || '{"fee_origin":' || r_����.��Դ;
          v_Fee_Item := v_Fee_Item || ',"fee_id":' || r_����.Id;
          v_Fee_Item := v_Fee_Item || ',"fee_no":"' || r_����.No || '"';
          v_Fee_Item := v_Fee_Item || ',"bill_prop":' || r_����.��¼����;
          v_Fee_Item := v_Fee_Item || ',"rec_state":' || r_����.��¼״̬;
          v_Fee_Item := v_Fee_Item || ',"serial_num":' || r_����.���;
          v_Fee_Item := v_Fee_Item || ',"fee_type":"' || r_����.�շ���� || '"';
          v_Fee_Item := v_Fee_Item || ',"fee_item_id":' || r_����.�շ�ϸĿid;
          v_Fee_Item := v_Fee_Item || ',"order_id":' || r_����.ҽ��id;
          v_Fee_Item := v_Fee_Item || ',"stuff_used":' || Nvl(r_����.��������, 0);
          v_Fee_Item := v_Fee_Item || ',"exe_dept_id":' || Nvl(r_����.ִ�в���id, 0); --ִ�в���id
          v_Fee_Item := v_Fee_Item || ',"is_verify":0';
        End If;
        v_Fee_Item := v_Fee_Item || ',"is_finish":1';
      End If;
    
      If v_Fee_Item Is Not Null Then
        v_Fee_Item := v_Fee_Item || '}';
      End If;
    
      If v_Fee_Item Is Not Null Then
        v_Fee_Item_List := v_Fee_Item_List || ',' || v_Fee_Item;
      End If;
    
    End Loop;
  
    If v_Fee_Item_List Is Not Null Then
      v_Fee_Item_List := ',"fee_list":[' || Substr(v_Fee_Item_List, 2) || ']';
    End If;
    v_Affirm := ',"is_affirm":' || Nvl(n_��˷���, 0);
    Json_Out := '{"output":{"code":1,"message":"�ɹ�"' || v_Affirm || v_Fee_Item_List || '}}';
  
  End;
  ---------------------------------------------------------------------------------------------------
Begin
  j_Tmp    := Pljson(Json_In);
  j_Json   := j_Tmp.Get_Pljson('input');
  n_Origin := j_Json.Get_Number('fee_origin');

  n_Finish     := j_Json.Get_Number('is_finish');
  v_ҽ��ids    := j_Json.Get_String('fee_order_ids');
  v_Nos        := j_Json.Get_String('fee_nos');
  n_ִ�в���id := Nvl(j_Json.Get_Number('exe_deptid'), 0);

  If 1 = n_Finish Then
    p_���ִ��(Json_Out);
    Return;
  End If;

  If 1 = n_Origin And 1 = n_Finish Then
    --����ҽ��ִ�����
    p_Checkoutfinish;
  Elsif 2 = n_Origin And 1 = n_Finish Then
    --סԺҽ��ִ�����
    p_Checkinfinish;
  End If;
  If 1 = n_Origin And 2 = n_Finish Then
    --����ҽ��ȡ��ִ�����
    p_Checkoutcancel;
  Elsif 2 = n_Origin And 2 = n_Finish Then
    --סԺҽ��ȡ��ִ�����
    p_Checkincancel;
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(v_Error) || '"}}';
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Exsesvr_Getorderfeeexeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Checkorderrevoke
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�����ҽ��������ؼ�飬��������ҽ������ҽ�����ϣ���һ�η��Ͳ�������һ��ҽ��
  --��Σ�Json_In:��ʽ
  --  input
  --     fee_nos                        C 1 ��ʽ��U0016921,U0016922,,,
  --     order_ids                      C 1 ҽ��id�������δ����һ��ҽ��id���ŷָ�
  --     bill_prop                      N 1 ��¼����,1-�շ�,2-����,������ҽ���˷�����,11��12
  --     after_order_ids                C 1 ������������Ϻ���ҩƷ,�ѷ�ҩ��ҽ���е�ҽ��id��ϸ��,���ŷָ�
  --     exe_fee_ids                    C 1 ��ִ�л�����ִ�еķ���idƴ������ҪӦ�������������Ϻ���ҩ�����������Ϊδ�շѣ�����ҩƷ�ѷ��ϵ����

  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exist_balance                 N 1 �Ƿ�����ѽ��ʵķ���,0-������,1-����
  --    exist_verify                  N 1 �Ƿ��������˵ļ��ʵ�,0-������,1-����
  --    del_list[]����ֱ��ɾ���ļ��ʵ�
  --       fee_source                 N 1 ������Դ:1-������ü�¼��2-סԺ���ü�¼
  --       fee_bill_type              N 1 ��¼����:1-�շѵ���2-���ʵ�
  --       fee_no                     C 1 ���õ��ݺ�
  --       exe_sta_nums               C 1 ��Ŵ�������������ִ��״̬����ʽ�����1,���2,��
  --       serial_num                 C 1 ���ָ�ʽ����¼����=1ʱ��ʽ�����1,���2,�� ��¼����=2ʱ��ʽ�����1:����:ִ��״̬1,���2:����2:ִ��״̬2,��
  ---------------------------------------------------------------------------
  j_Input         Pljson;
  j_Json          Pljson;
  v_Nos           Varchar2(32767);
  v_������ҽ��ids Varchar2(32767);
  v_�޸�ִ��״̬  Varchar2(32767);
  v_Jtmp1         Varchar2(32767);
  v_Serial_Num    Varchar2(32767);
  v_Exe_Fee_Ids   Varchar2(32767);
  v_Order_Ids     Varchar2(32767);
  v_Count         Number;
  n_��¼����      Number(3);
  v_Error         Varchar2(255);
  Err_Custom Exception;
  n_���ѽ��ʷ��� Number;
  n_�м�������� Number;
  v_Tmpno        Varchar2(30);
  v_Del_List     Varchar2(32767);

  --��Ҫ���ʵķ����б�
  Cursor c_Billdel Is
    Select a.��¼����, a.No, f_List2str(Cast(Collect(a.��� || '') As t_Strlist), ',') ��Ŵ�,
           f_List2str(Cast(Collect(a.�����ʵ� || '') As t_Strlist), ',') �����ʵ�
    From (Select Decode(a.��¼����, 11, 1, a.��¼����) As ��¼����, a.��¼״̬, a.No, a.���, a.��� || ':' || (Nvl(����, 1) * ����) || ':0' �����ʵ�
           From ������ü�¼ A
           Where a.ҽ����� Is Not Null And Instr(',' || v_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And
                 Mod(a.��¼����, 10) = n_��¼���� And a.��¼״̬ In (0, 1) And Nvl(Nvl(����, 1) * ����, 0) <> 0 And
                 (v_Exe_Fee_Ids Is Null Or Instr(',' || v_Exe_Fee_Ids || ',', ',' || a.Id || ',') = 0) And
                 a.No In (Select /*+cardinality(X,10)*/
                           x.Column_Value
                          From Table(f_Str2list(v_Nos)) X)) A
    Group By a.��¼����, a.No;

Begin
  --�������
  j_Input         := Pljson(Json_In);
  j_Json          := j_Input.Get_Pljson('input');
  n_��¼����      := j_Json.Get_Number('bill_prop');
  v_Nos           := j_Json.Get_String('fee_nos');
  v_������ҽ��ids := j_Json.Get_String('after_order_ids');
  v_Exe_Fee_Ids   := j_Json.Get_String('exe_fee_ids');
  v_Order_Ids     := j_Json.Get_String('order_ids');

  --����ת���ж�
  Select Count(1)
  Into v_Count
  From H������ü�¼ A
  Where a.ҽ����� Is Not Null And (v_������ҽ��ids Is Null Or Instr(',' || v_������ҽ��ids || ',', ',' || a.ҽ����� || ',') = 0) And
        Instr(',' || v_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And Mod(a.��¼����, 10) = n_��¼���� And
        a.No In (Select /*+cardinality(X,10)*/
                  x.Column_Value
                 From Table(f_Str2list(v_Nos)) X);
  If v_Count > 0 Then
    v_Error := '��ҽ���ķ����Ѿ�ȫ���򲿷�ת���������ݿ⣬�����������' || Chr(13) || '��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�';
    Raise Err_Custom;
  End If;

  --�������ҽ����Ӧ�ķ����Ƿ��������
  Select Count(1)
  Into n_���ѽ��ʷ���
  From (Select Count(1) ��¼����
         From ������ü�¼ A
         Where a.ҽ����� Is Not Null And (v_������ҽ��ids Is Null Or Instr(',' || v_������ҽ��ids || ',', ',' || a.ҽ����� || ',') = 0) And
               Instr(',' || v_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And Mod(a.��¼����, 10) = n_��¼���� And
               a.��¼���� In (2, 12) And a.��¼״̬ = 1 And
               Not (Nvl(a.����״̬, 0) = 1 And Nvl(a.����id, 0) = 0 And Nvl(a.��¼״̬, 0) = 1) And
               a.No In (Select /*+cardinality(X,10)*/
                         x.Column_Value
                        From Table(f_Str2list(v_Nos)) X)
         Group By a.No, Nvl(a.�۸񸸺�, a.���)
         Having Sum(Nvl(a.���ʽ��, 0)) <> 0);

  --����˼��ʷ��ü�����Ƿ�������˵ļ��ʷ���
  Select Count(1)
  Into n_�м��������
  From ������ü�¼ A
  Where a.ҽ����� Is Not Null And (v_������ҽ��ids Is Null Or Instr(',' || v_������ҽ��ids || ',', ',' || a.ҽ����� || ',') = 0) And
        Instr(',' || v_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And Mod(a.��¼����, 10) = n_��¼���� And a.��¼���� In (2, 12) And
        a.��¼״̬ = 1 And Not (Nvl(a.����״̬, 0) = 1 And Nvl(a.����id, 0) = 0 And Nvl(a.��¼״̬, 0) = 1) And a.������ Is Not Null And
        a.������ <> a.����Ա���� And Not (a.����״̬ = 1 And a.����id Is Null And a.��¼״̬ = 1) And
        a.No In (Select /*+cardinality(X,10)*/
                  x.Column_Value
                 From Table(f_Str2list(v_Nos)) X);

  ----�շ��쳣���
  Select Max(a.No)
  Into v_Tmpno
  From ������ü�¼ A
  Where a.ҽ����� Is Not Null And (v_������ҽ��ids Is Null Or Instr(',' || v_������ҽ��ids || ',', ',' || a.ҽ����� || ',') = 0) And
        Instr(',' || v_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And Mod(a.��¼����, 10) = n_��¼���� And a.��¼״̬ In (0, 1) And
        a.ִ��״̬ = 9 And a.No In (Select /*+cardinality(X,10)*/
                                 x.Column_Value
                                From Table(f_Str2list(v_Nos)) X);

  If v_Tmpno Is Not Null Then
    v_Error := 'ҽ�����õ���"' || v_Tmpno || '"�е��շѽ�������쳣���������ϡ�';
    Raise Err_Custom;
  End If;

  If n_��¼���� = 1 Then
    --���շ��շѵ�
    --�����շѵ����ж��Ƿ��Ѿ��շѣ����ų��Զ�ȡ���������Ϻ���ҩ��ҩƷ���ķ���
    Select Max(a.No)
    Into v_Tmpno
    From ������ü�¼ A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And a.ҽ����� Is Not Null And
          (v_������ҽ��ids Is Null Or Instr(',' || v_������ҽ��ids || ',', ',' || a.ҽ����� || ',') = 0) And
          Instr(',' || v_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And Mod(a.��¼����, 10) = n_��¼���� And a.��¼״̬ = 1 And
          a.No In (Select /*+cardinality(X,10)*/
                    x.Column_Value
                   From Table(f_Str2list(v_Nos)) X);
  
    If v_Tmpno Is Not Null Then
      v_Error := 'ҽ�����õ���"' || v_Tmpno || '"�Ѿ��շѣ��������ϡ�';
      Raise Err_Custom;
    End If;
  End If;

  If v_������ҽ��ids Is Not Null Then
    --�����Ϻ���ҩ��ҽ��ҩƷҽ����Ҫ������ҽ�������շѵķ�ҩƷ���ĵķ���ִ��״̬��Ϊδִ�з����˷�
    For R In (Select a.No, f_List2str(Cast(Collect(a.��� || '') As t_Strlist), ',') ���
              From ������ü�¼ A, �������� B
              Where a.�շ�ϸĿid = b.����id(+) And Instr(',' || v_������ҽ��ids || ',', ',' || a.ҽ����� || ',') > 0 And
                    Instr(',' || v_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And Mod(a.��¼����, 10) = n_��¼���� And
                    a.��¼״̬ = 1 And a.ִ��״̬ = 1 And Not (a.�շ���� In ('5', '6', '7') Or a.�շ���� = '4' And b.�������� = 1) And
                    a.No In (Select /*+cardinality(X,10)*/
                              x.Column_Value
                             From Table(f_Str2list(v_Nos)) X)
              Group By a.No) Loop
    
      --�޸�ִ��״̬
      v_�޸�ִ��״̬ := v_�޸�ִ��״̬ || ',{"fee_source":1';
      v_�޸�ִ��״̬ := v_�޸�ִ��״̬ || ',"fee_bill_type":1';
      v_�޸�ִ��״̬ := v_�޸�ִ��״̬ || ',"fee_no":"' || r.No || '"';
      v_�޸�ִ��״̬ := v_�޸�ִ��״̬ || ',"exe_sta_nums":"' || r.��� || '"';
      v_�޸�ִ��״̬ := v_�޸�ִ��״̬ || '}';
    
    End Loop;
  End If;

  v_Del_List := Null;
  For R In c_Billdel Loop
    If r.��¼���� = 1 Then
      v_Serial_Num := r.��Ŵ�;
    Else
      v_Serial_Num := r.�����ʵ�;
    End If;
    --����ɾ���б�
    v_Del_List := v_Del_List || ',{"fee_source":1';
    v_Del_List := v_Del_List || ',"fee_bill_type":' || r.��¼����;
    v_Del_List := v_Del_List || ',"fee_no":"' || r.No || '"';
    v_Del_List := v_Del_List || ',"serial_num":"' || v_Serial_Num || '"';
    v_Del_List := v_Del_List || ',"exe_sta_nums":"' || r.��Ŵ� || '"';
    v_Del_List := v_Del_List || '}';
  End Loop;

  If v_�޸�ִ��״̬ Is Not Null Then
    v_Del_List := v_Del_List || v_�޸�ִ��״̬;
  End If;

  v_Jtmp1 := Null;
  v_Jtmp1 := v_Jtmp1 || ',"exist_balance":' || Nvl(n_���ѽ��ʷ���, 0); --���ѽ��ʵķ���
  v_Jtmp1 := v_Jtmp1 || ',"exist_verify":' || Nvl(n_�м��������, 0); --������˵ļ��ʷ���
  If Not v_Del_List Is Null Then
    v_Jtmp1 := v_Jtmp1 || ',"del_list":[' || Substr(v_Del_List, 2) || ']'; --��ɾ�����ݷ���
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�"' || v_Jtmp1 || '}}';

Exception
  When Err_Custom Then
    Json_Out := Zljsonout(v_Error);
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkorderrevoke;
/


Create Or Replace Procedure Zl_Exsesvr_Checkbabyfee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���Ӥ���Ƿ��Ѿ���������
  --��Σ�Json_In:��ʽ
  --   input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ����id
  --      baby_nums          C 1 Ӥ�����,���������ö��ŷ���;NULL��ʾ��ò��˵�����Ӥ��
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     baby_nums          C 1 Ӥ����� :���������ö��ŷָ�
  ---------------------------------------------------------------------------

  j_Input       PLJson;
  j_Json        PLJson;
  v_Output      Varchar2(4000);
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(5);
  v_Baby_Nums   Varchar2(2000);
  v_Babys       Varchar2(1000);
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  v_Baby_Nums   := j_Json.Get_Number('baby_nums');

  If n_Pati_Id Is Null Then
    Json_Out := zlJsonOut('δ���벡��ID�����飡');
    Return;
  End If;

  If n_Pati_Pageid Is Null Then
    Json_Out := zlJsonOut('δ������ҳID�����飡');
    Return;
  End If;

  If v_Baby_Nums Is Null Then
    Json_Out := zlJsonOut('δ����Ӥ����ţ����飡');
    Return;
  End If;
  v_Babys := Null;
  For R In (Select Distinct Ӥ����
            From סԺ���ü�¼
            Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And
                  ((Instr(',' || Nvl(v_Baby_Nums, '') || ',', ',' || Nvl(Ӥ����, 0) || ',') > 0) Or
                  (v_Baby_Nums Is Null And Nvl(Ӥ����, 0) > 0))) Loop
    v_Babys := Nvl(v_Babys, '') || ',' || r.Ӥ����;
  End Loop;

  If v_Babys Is Not Null Then
    v_Babys := Substr(v_Babys, 2);
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'baby_nums', v_Babys, 0, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkbabyfee;
/


Create Or Replace Procedure Zl_Exsesvr_Updateoutprevstsign
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ����²��˸����־
  --��Σ�Json_In:��ʽ
  --input
  --  reg_id             N  1 �Һ�ID
  --  revst_sign         N  1 �����־ 0:���Ϊ����.1:���Ϊ���� 

  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_�Һ�id   Number;
  n_�����־ Number;

  j_Input PLJson;
  j_Json  PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�Һ�id   := j_Json.Get_Number('reg_id');
  n_�����־ := Nvl(j_Json.Get_Number('revst_sign'), 0);

  Zl_���˹Һż�¼_����(n_�Һ�id, n_�����־);
  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updateoutprevstsign;
/
Create Or Replace Procedure Zl_Exsesvr_Checkorderroll
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�סԺҽ��ҽ��������ط��ü�飬ͬʱ��ȡ���������Ϣ
  --˵�������õ��ݶ�ȫ����������ʱ���ʲ���
  --      �������ִ��״̬������ҩƷȡ��ִ�У�ɾ�����е��ݣ��쳣����
  --��Σ�Json_In:��ʽ
  --  input
  --     outpati_account                N 1 �������
  --     bill_prop                      N 1 ��¼����
  --     fee_nos                        C 1 ��ʽ��T000001,T000002,T000003...
  --     order_ids                      C 1 ҽ��id�������δ����һ��ҽ��id���ŷָ�
  --     check_pacs                     N 0 �����ҽ�����˷���ʱ��Ч��1-�Ƿ����δ��˵���������

  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist                       N 1 �Ƿ���ڣ�0-���棬1-����
  --    fee_nos                       C 1 ����ƴ��������ҽ�������ϴ���no��
  --    del_list[]����ֱ��ɾ���ļ��ʵ�
  --       fee_source                 N 1 ������Դ:1-������ü�¼��2-סԺ���ü�¼
  --       fee_bill_type              N 1 ��¼����:1-�շѵ���2-���ʵ�
  --       fee_no                     C 1 ���õ��ݺ�
  --       exe_sta_nums               C 1 ��Ŵ�������������ִ��״̬����ʽ�����1,���2,��
  --       serial_num                 C 1 ���ָ�ʽ����¼����=1ʱ��ʽ�����1,���2,�� ��¼����=2ʱ��ʽ�����1:����:ִ��״̬1,���2:����2:ִ��״̬2,��
  ---------------------------------------------------------------------------

  j_Input     Pljson;
  j_Json      Pljson;
  v_Nos       Varchar2(32767);
  v_Order_Ids Varchar2(32767);
  c_Order_Ids Clob;
  v_Del_List  Varchar2(32767);
  c_Del_List  Clob;

  c_Out_Tmp Clob;
  I         Number;
  v_Vals    Clob; --���᳤ܻ���Ľ���ô˱��� v_vals , col_vals
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_Vals     Collection_Type;
  n_��¼����   Number;
  n_�������   Number;
  n_Check_Pacs Number;
  v_Jtmp1      Varchar2(32767);
  v_Serial_Num Varchar2(32767);
  v_No_����    Varchar2(32767);
  v_Count      Number(5);
  v_No_���շ�  Varchar2(300);
  v_Error      Varchar2(2000);
  Err_Custom Exception;

  --ҽ���˷����ջ��м�¼����Ϊ��λ��Ŀǰֻ�����ﲡ�˷��͵�ҽ���Ż��У��˴����漰
  Cursor c_Billdel Is
    Select a.��¼����, a.������Դ, a.No, Max(a.ҽ����Ҫ�ĵ���) ҽ����Ҫ�ĵ���, f_List2str(Cast(Collect(a.��� || '') As t_Strlist), ',') ��Ŵ�,
           f_List2str(Cast(Collect(a.�����ʵ� || '') As t_Strlist), ',') �����ʵ���Ŵ�
    From (Select 1 ������Դ, a.��¼״̬, a.No, a.���, a.��¼����, Decode(a.��¼���� || a.��¼״̬, '21', a.No, Null) ҽ����Ҫ�ĵ���,
                  a.��� || ':' || (Nvl(a.����, 1) * a.����) || ':0' �����ʵ�
           From ������ü�¼ A
           Where a.��¼״̬ In (0, 1) And a.��¼���� = n_��¼���� And Instr(',' || c_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And
                 Nvl(Nvl(a.����, 1) * a.����, 0) <> 0 And
                 a.No In (Select /*+cardinality(X,10)*/
                           x.Column_Value
                          From Table(f_Str2list(v_Nos)) X)
           Union All
           Select 2 ������Դ, a.��¼״̬, a.No, a.���, a.��¼����, Decode(a.��¼���� || a.��¼״̬, '21', a.No, Null) ҽ����Ҫ�ĵ���,
                  a.��� || ':' || (Nvl(a.����, 1) * a.����) || ':0' �����ʵ�
           From סԺ���ü�¼ A
           Where a.��¼״̬ In (0, 1) And a.��¼���� = n_��¼���� And Instr(',' || c_Order_Ids || ',', ',' || a.ҽ����� || ',') > 0 And
                 Nvl(Nvl(a.����, 1) * a.����, 0) <> 0 And
                 a.No In (Select /*+cardinality(X,10)*/
                           x.Column_Value
                          From Table(f_Str2list(v_Nos)) X)) A
    Group By a.��¼����, a.������Դ, a.No;

Begin

  --�������
  j_Input      := Pljson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_��¼����   := j_Json.Get_Number('bill_prop');
  n_�������   := j_Json.Get_Number('outpati_account');
  n_Check_Pacs := j_Json.Get_Number('check_pacs');

  If n_Check_Pacs = 1 Then
    --��Ҫ�Ǽ������ҽ�� D �󶨵�ҩƷҪ�Զ�������������
    v_Nos       := j_Json.Get_String('fee_nos');
    v_Order_Ids := j_Json.Get_String('order_ids');
    If n_������� = 1 Then
      Select Count(1)
      Into v_Count
      From ������ü�¼ C, ���˷������� D
      Where c.Id = d.����id And c.��¼״̬ In (0, 1, 3) And d.״̬ = 0 And c.��¼���� = n_��¼���� And
            Instr(',' || v_Order_Ids || ',', ',' || c.ҽ����� || ',') > 0 And
            c.No In (Select /*+cardinality(X,10)*/
                      x.Column_Value
                     From Table(f_Str2list(v_Nos)) X);
    Else
      Select Count(1)
      Into v_Count
      From סԺ���ü�¼ C, ���˷������� D
      Where c.Id = d.����id And c.��¼״̬ In (0, 1, 3) And d.״̬ = 0 And c.��¼���� = n_��¼���� And
            Instr(',' || v_Order_Ids || ',', ',' || c.ҽ����� || ',') > 0 And
            c.No In (Select /*+cardinality(X,10)*/
                      x.Column_Value
                     From Table(f_Str2list(v_Nos)) X);
    End If;
  
    If v_Count > 0 Then
      v_Count := 1;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","isexist":"' || v_Count || '"}}';
  Else
    c_Order_Ids := j_Json.Get_Clob('order_ids');
  
    --���ݺŷֽ�----
    v_Vals := j_Json.Get_Clob('fee_nos');
    I      := 0;
    While v_Vals Is Not Null Loop
      If Length(v_Vals) <= 4000 Then
        Col_Vals(I) := v_Vals;
        v_Vals := Null;
      Else
        Col_Vals(I) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
        v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
      End If;
      I := I + 1;
    End Loop;
    --���ݺŷֽ�----
  
    For Lp In 0 .. Col_Vals.Count - 1 Loop
      v_Nos := Col_Vals(Lp);
      --�ж�����ת��ʱ���Բ���  ҽ�����  �����������Ҳ�ܴﵽ��ͬЧ��
      --�ж������Ƿ��Ѿ�ת��begin
      If n_������� = 1 Then
        Select Count(1)
        Into v_Count
        From H������ü�¼ A
        Where a.��¼���� = n_��¼���� And a.ҽ����� Is Not Null And
              a.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      Else
        Select Count(1)
        Into v_Count
        From HסԺ���ü�¼ A
        Where a.��¼���� = n_��¼���� And a.ҽ����� Is Not Null And
              a.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      End If;
      If v_Count > 0 Then
        v_Error := '��ҽ���ķ����Ѿ�ȫ���򲿷�ת���������ݿ⣬�����������' || Chr(13) || '��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�';
        Raise Err_Custom;
      End If;
      --�ж������Ƿ��Ѿ�ת��end
    
      --δ��˵���������begin
      If n_������� = 1 Then
        Select Count(1)
        Into v_Count
        From ������ü�¼ C, ���˷������� D
        Where c.Id = d.����id And c.��¼״̬ In (0, 1, 3) And d.״̬ = 0 And c.��¼���� = n_��¼���� And
              Instr(',' || c_Order_Ids || ',', ',' || c.ҽ����� || ',') > 0 And
              c.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      Else
        Select Count(1)
        Into v_Count
        From סԺ���ü�¼ C, ���˷������� D
        Where c.Id = d.����id And c.��¼״̬ In (0, 1, 3) And d.״̬ = 0 And c.��¼���� = n_��¼���� And
              Instr(',' || c_Order_Ids || ',', ',' || c.ҽ����� || ',') > 0 And
              c.No In (Select /*+cardinality(X,10)*/
                        x.Column_Value
                       From Table(f_Str2list(v_Nos)) X);
      End If;
      If v_Count > 0 Then
        v_Error := '��ҽ������δ��˵��������룬��ȡ�����������������ٻ��˷��͡�';
        Raise Err_Custom;
      End If;
      --δ��˵���������end
    
      --���շѵ������շѵ����ж�begin
      Select Max(c.No)
      Into v_No_���շ�
      From ������ü�¼ C
      Where c.��¼״̬ = 1 And c.�����־ = 1 And c.��¼���� = 1 And Instr(',' || c_Order_Ids || ',', ',' || c.ҽ����� || ',') > 0 And
            c.No In (Select /*+cardinality(X,10)*/
                      x.Column_Value
                     From Table(f_Str2list(v_Nos)) X);
      If v_No_���շ� Is Not Null Then
        v_Error := '��ҽ�����͵����ﵥ�ݡ�' || v_No_���շ� || '�����շѣ����ܻ��ˡ�';
        Raise Err_Custom;
      End If;
      --���շѵ������շѵ����ж�end
    
      ----�ռ�Ҫɾ���ķ��õ�����Ϣbegin
      For R In c_Billdel Loop
      
        If r.��¼���� = 2 Then
          v_Serial_Num := r.�����ʵ���Ŵ�;
        Else
          v_Serial_Num := r.��Ŵ�;
        End If;
      
        --�ռ�ҽ����Ҫ�ĵ��ݺŴ�
        If r.ҽ����Ҫ�ĵ��� Is Not Null Then
          If Instr(',' || v_No_���� || ',', ',' || r.ҽ����Ҫ�ĵ��� || ',') = 0 Then
            v_No_���� := v_No_���� || ',' || r.ҽ����Ҫ�ĵ���;
          End If;
        End If;
      
        --����ɾ���б�
        v_Del_List := v_Del_List || ',{"fee_source":' || r.������Դ;
        v_Del_List := v_Del_List || ',"fee_bill_type":' || r.��¼����;
        v_Del_List := v_Del_List || ',"fee_no":"' || r.No || '"';
        v_Del_List := v_Del_List || ',"serial_num":"' || v_Serial_Num || '"';
        v_Del_List := v_Del_List || ',"exe_sta_nums":"' || r.��Ŵ� || '"';
        v_Del_List := v_Del_List || '}';
      
        If Length(v_Del_List) > 20000 Then
          If c_Del_List Is Null Then
            c_Del_List := v_Del_List;
          Else
            c_Del_List := c_Del_List || v_Del_List;
          End If;
          v_Del_List := Null;
        End If;
      End Loop;
      ----�ռ�Ҫɾ���ķ��õ�����Ϣend    
    End Loop;
  
    v_Jtmp1 := Null;
    If Not v_No_���� Is Null Then
      v_Jtmp1 := v_Jtmp1 || ',"fee_nos":"' || Substr(v_No_����, 2) || '"'; --����ҽ���ϴ�
    End If;
    c_Out_Tmp := v_Jtmp1;
  
    If Not v_Del_List Is Null Then
      c_Del_List := c_Del_List || v_Del_List;
      c_Del_List := ',"del_list":[' || Substr(c_Del_List, 2) || ']'; --��ɾ�����ݷ���
    End If;
    c_Out_Tmp := c_Out_Tmp || c_Del_List;
    Json_Out  := '{"output":{"code":1,"message":"�ɹ�"' || c_Out_Tmp || '}}';
  End If;
Exception
  When Err_Custom Then
    Json_Out := Zljsonout(v_Error);
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkorderroll;
/


Create Or Replace Procedure Zl_Exsesvr_Billchargeoff
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���������������������
  --��Σ�Json_In:��ʽ
  --  input
  --    request_operator                C 1 ������  
  --    request_code                    C 0 �����˱��
  --    request_time                    C 1 ����ʱ��
  --    request_type                    N 1 �������   
  --    del_tag                         N 1 ɾ����־
  --    reason                          C 1 ����ԭ��
  --    item_list[]������������������б�
  --        fee_id                      N 1 ����ID
  --        request_dept_id             N 1 �����������ID
  --        fee_item_id                 N 1 �շ�ϸĿID
  --        quantity                    N 1 ����
  --        audit_dept_id               N 1 ��˲���id
  --        auto_aduit                  N 0 �Ƿ��Զ����
  --        outpati_account             N 0 �Ƿ��������        
  --        fee_no                      C 0 ���õ��ݺ�
  --        serial_num                  N 0 ��ţ����1:����1:ִ��״̬1,���2:����2:ִ��״̬2,...���n:����n:ִ��״̬n  ��:"1:2:1,2:10:1,3:2:1"      
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  Id_In         ���˷�������.����id%Type;
  �շ�ϸĿid_In ���˷�������.�շ�ϸĿid%Type;
  ���벿��id_In ���˷�������.���벿��id%Type;
  ����_In       ���˷�������.����%Type;
  ������_In     ���˷�������.������%Type;
  �����˱��_In ���˷�������.������%Type;
  ����ʱ��_In   ���˷�������.����ʱ��%Type;
  �������_In   ���˷�������.�������%Type;
  ����ԭ��_In   ���˷�������.����ԭ��%Type;
  ��˲���id_In ���˷�������.��˲���id%Type;
  ɾ����־_In   Integer;
  ���ʱ��_In   ���˷�������.����ʱ��%Type;
  j_Input       PLJson;
  j_Json        PLJson;
  j_Item        PLJson;
  j_List        Pljson_List := Pljson_List();
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  ������_In     := j_Json.Get_String('request_operator');
  �����˱��_In := j_Json.Get_String('request_code');
  ����ʱ��_In   := To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
  �������_In   := j_Json.Get_Number('request_type');
  �������_In   := Nvl(�������_In, 0);
  ɾ����־_In   := j_Json.Get_Number('del_tag');
  ɾ����־_In   := Nvl(ɾ����־_In, 0);
  ����ԭ��_In   := j_Json.Get_String('reason');

  --������Զ����ʱ���ʱ������ã�Ϊ�˽�ʱ��ֿ���һ��
  ���ʱ��_In := ����ʱ��_In + 1 / 24 / 60 / 60;

  j_List := j_Json.Get_Pljson_List('item_list');
  For I In 1 .. j_List.Count Loop
    j_Item        := PLJson(j_List.Get(I));
    Id_In         := j_Item.Get_Number('fee_id');
    �շ�ϸĿid_In := j_Item.Get_Number('fee_item_id');
    ���벿��id_In := j_Item.Get_Number('request_dept_id');
    ����_In       := j_Item.Get_Number('quantity');
    ��˲���id_In := j_Item.Get_Number('audit_dept_id');
    Zl_���˷�������_Insert_s(Id_In, �շ�ϸĿid_In, ���벿��id_In, ����_In, ������_In, ����ʱ��_In, �������_In, ����ԭ��_In, ��˲���id_In, ɾ����־_In);
    If 1 = j_Item.Get_Number('auto_aduit') Then
      Zl_���˷�������_Audit_s(Id_In, ����ʱ��_In, ������_In, ���ʱ��_In, 1, �������_In);
      If 1 = j_Item.Get_Number('outpati_account') Then
        Zl_������ʼ�¼_Delete_s(j_Item.Get_String('fee_no'), j_Item.Get_String('serial_num'), �����˱��_In, ������_In, ���ʱ��_In, 2);
      Else
        Zl_סԺ���ʼ�¼_Delete_s(j_Item.Get_String('fee_no'), j_Item.Get_String('serial_num'), �����˱��_In, ������_In, 2, 0,
                           ���ʱ��_In);
      End If;
    End If;
  End Loop;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Billchargeoff;
/


Create Or Replace Procedure Zl_Exsesvr_Checkbillchargeoff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ָ������ָ�����н������������飬��ȡ��ط�������
  --��Σ�Json_In:��ʽ
  --input
  --     oper_type                       N   1   �������ͣ�0-���������б��鴫��item_list;
  --                                                       1-����������ȡ����fee_ids+fee_source��ȡ���ʷ�����ϸ
  --                                                       2-ȡ�����������ȡ��Ч��������ϸ������fee_ids
  --     fee_ids                         C   1   ����IDs��ϸ
  --     fee_source                      N   1   ������Դ,1-���2-סԺ
  --     item_list[]���������б�
  --         fee_id                      N   1   ����ID
  --         request_dept_id             N   1   �����������ID
  --         item_id                     N   1   �շ�ϸĿID
  --         request_type                N   1   �������:��ҩƷ��������Ч:0-δ��ҩ(��);1-�ѷ�ҩ(��);����Ϊ0
  --         request_num                 N   1   ��������
  --         sended_num                  N   1   �ѷ�����
  --    pati_list[]                 ������Ϣ
  --         pati_id                     N   1   ����ID
  --         pati_name                   C   1   ��������
  --         pati_dept_id                N   1   ��Ժ����id,������ҳ.��Ժ����
  --         fee_audit_status            N   1   ��˱�־:������ҳ.��˱�־
  --         si_inp_status               N   1   סԺ״̬:������ҳ.״̬(0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ)
  --����: Json_Out,��ʽ����
  --output
  --     code                            N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --     message                         C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     fee_ids                         C   1   ������ϸid,oper_type=2����
  --     item_list[]���������б�oper_type=0
  --         fee_id                      N   1   ����ID
  --         request_dept_id             N   1   �����������ID
  --         audit_dept_id               N   1   ������˿���ID
  --     charge_list[]���ʵķ�����ϸ�б�oper_type=1
  --         pati_id                     N   1   ����id
  --         pati_pageid                 N   1   ��ҳid
  --         fee_id                      N   1   ����id
  --         serial_num                  N   1   ���
  --         rec_status                  N   1   ��¼״̬
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_��������   Number;
  v_Feeids     Varchar2(32767);
  v_Tmp        Varchar2(32767);
  n_Fee_Source Number;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�������� := j_Json.Get_Number('oper_type');
  If Nvl(n_��������, 0) = 0 Then
    Zl_���˷�������_Insert_Check(Json_In, Json_Out);
  Elsif n_�������� = 1 Then
    v_Feeids     := j_Json.Get_String('fee_ids');
    n_Fee_Source := j_Json.Get_Number('fee_source');
    v_Tmp        := Null;
    If n_Fee_Source = 2 Then
      For R In (Select e.Id, e.���, e.��¼״̬, e.����id, e.��ҳid
                From סԺ���ü�¼ E
                Where e.Id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_Feeids) As Zltools.t_Numlist)) X)) Loop
      
        v_Tmp := v_Tmp || ',{"pati_id":' || r.����id;
        v_Tmp := v_Tmp || ',"pati_pageid":' || Nvl(r.��ҳid || '', 'null');
        v_Tmp := v_Tmp || ',"fee_id":' || r.Id;
        v_Tmp := v_Tmp || ',"serial_num":' || r.���;
        v_Tmp := v_Tmp || ',"rec_status":' || Nvl(r.��¼״̬, 0);
        v_Tmp := v_Tmp || '}';
      
      End Loop;
    Else
      For R In (Select e.Id, e.���, e.��¼״̬, e.����id, e.��ҳid
                From ������ü�¼ E
                Where e.Id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_Feeids) As Zltools.t_Numlist)) X)) Loop
        v_Tmp := v_Tmp || ',{"pati_id":' || r.����id;
        v_Tmp := v_Tmp || ',"pati_pageid":' || Nvl(r.��ҳid || '', 'null');
        v_Tmp := v_Tmp || ',"fee_id":' || r.Id;
        v_Tmp := v_Tmp || ',"serial_num":' || r.���;
        v_Tmp := v_Tmp || ',"rec_status":' || Nvl(r.��¼״̬, 0);
        v_Tmp := v_Tmp || '}';
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","charge_list":[' || Substr(v_Tmp, 2) || ']}}';
  Elsif n_�������� = 2 Then
    v_Feeids := j_Json.Get_String('fee_ids');
    v_Tmp    := Null;
    For R In (Select Distinct e.����id
              From ���˷������� E
              Where e.����� Is Null And
                    e.����id In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2List(v_Feeids) As Zltools.t_Numlist)) X)) Loop
    
      v_Tmp := v_Tmp || ',' || r.����id;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ids":"' || Substr(v_Tmp, 2) || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkbillchargeoff;
/


Create Or Replace Procedure Zl_Exsesvr_Delbillchargeoff
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ�ɾ����������
  --��Σ�Json_In:��ʽ
  --  input
  --       fee_ids               C 1 ����ids������idƴ��
  --       request_time          C 1 ���������ʱ��                    
  --����: Json_Out,��ʽ����
  --  output
  --       code                  N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --       message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  Zl_���˷�������_Delete_s(j_Json.Get_String('fee_ids'), To_Date(j_Json.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss'));
  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delbillchargeoff;
/


Create Or Replace Procedure Zl_Exsesvr_Checkshareinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��ж��Ƿ����ָ���Ĺ���Ʊ������
  --��Σ�Json_In:��ʽ
  --  input
  --   recv_id          N  1  ����id
  --   invc_type        N  1  Ʊ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist              N   1   �Ƿ���ڣ�0-������ 1-���ڣ�

  j_Input PLJson;
  j_Json  PLJson;
  n_Id    Ʊ�����ü�¼.Id%Type;
  n_Kind  Ʊ�����ü�¼.Ʊ��%Type;
  n_Count Number;

  v_Output  Varchar2(32767);
  n_Isexist Number(1);
Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Kind := j_Json.Get_Number('invc_type');
  n_Id   := j_Json.Get_Number('recv_id');

  Select Count(1) Into n_Count From Ʊ�����ü�¼ Where ID = n_Id And Ʊ�� = n_Kind And ʹ�÷�ʽ = 2;
  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'isexist', n_Isexist, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkshareinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Updpatbaseinfocheck
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ------------------------------------------------------------------------------------------
  --����:���·������ҵ�����ݵĲ��˻�����Ϣ�ļ��
  --���:JSON��ʽ
  --input
  --    vist_id   N 1 ����id��_ ���ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;Ϊ0˵����������ǹҺž���Ĳ���(����id_InΪ��ʱ,�����ĸò��˵ķ��ò��ֵ�ҵ������)
  --    occasion   N 1 ����,1-����;2-סԺ
  --����:JSON��ʽ
  --output
  --  code            N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message        C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  explain   C 1 ˵��
  ------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  d_Maxdate Date;
  v_No      ������ü�¼.No%Type;
  v_����    ���ű�.����%Type;
  v_˵��    Clob;
  n_����id  Number;
  n_����id  Number;
  n_����    Number;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_����id := j_Json.Get_Number('visit_id');
  n_����   := j_Json.Get_Number('occasion');
  If Nvl(n_����id, 0) = 0 Then
    Return;
  End If;
  If Nvl(n_����, 0) <= 1 Then
    --����
    If Nvl(n_����id, 0) <> 0 Then
      Begin
        Select a.No, b.����, a.�Ǽ�ʱ��
        Into v_No, v_����, d_Maxdate
        From ���˹Һż�¼ A, ���ű� B
        Where a.ִ�в���id = b.Id(+) And a.Id = n_����id;
      Exception
        When Others Then
          v_No := Null;
      End;
      If Not v_No Is Null Then
        v_˵�� := '�Һŵ�:' || v_No || LPad(' ', 4) || '�Һſ���:' || v_���� || ' �ջ�Ʊ����Ϣ:';
      End If;
    Else
      v_˵�� := '�ջ�Ʊ����Ϣ:';
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","explain":"' || v_˵�� || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Updpatbaseinfocheck;
/


Create Or Replace Procedure Zl_Exsesvr_Getdynamiccost
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ��̬�ѱ�
  --��Σ�Json_In:��ʽ
  --  input
  --    dept_id             N 1 ����ID
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fee_category        C 1 �ѱ�ƴ�������ŷָ�
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id Number(18);
  v_�ѱ�   Varchar2(32767);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('dept_id');
  For R In (Select ����, ����, ����
            From �ѱ�
            Where Nvl(����, 1) = 2 And Nvl(���ÿ���, 1) = 1 And Nvl(�������, 3) In (1, 3) And
                  Trunc(Sysdate) Between Nvl(��Ч��ʼ, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                  Nvl(��Ч����, To_Date('3000-01-01', 'YYYY-MM-DD'))
            Union All
            Select Distinct a.����, a.����, a.����
            From �ѱ� A, �ѱ����ÿ��� B
            Where a.���� = b.�ѱ� And b.����id = n_����id And Nvl(a.����, 1) = 2 And Nvl(a.���ÿ���, 1) = 2 And
                  Nvl(a.�������, 3) In (1, 3) And Trunc(Sysdate) Between Nvl(a.��Ч��ʼ, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                  Nvl(a.��Ч����, To_Date('3000-01-01', 'YYYY-MM-DD'))
            Order By ����) Loop
    v_�ѱ� := v_�ѱ� || ',' || r.����;
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_category":"' || Substr(v_�ѱ�, 2) || '"}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdynamiccost;
/

Create Or Replace Procedure Zl_Exsesvr_Getneedaudititems
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ�����˵ķ���������Ŀ
  --��Σ�Json_In:��ʽ
  --  input
  --       pati_id          N 1 ����ID
  --       pati_pageid      N 1 ��ҳid
  --       fitem_id         N 1 ��ĿID
  --����: Json_Out,��ʽ����
  --  output
  --       code             N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --       message          C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --       item_list[]
  --          fitem_id         N 1 ��ĿID
  --          limit_quantity  N 1 ʹ������
  --          used_quantity   N 1 ��������
  --          avail_quantity  N 1 ��������

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id Number(18);
  n_��ҳid Number(18);
  n_��Ŀid Number(18);

  v_Output Varchar2(32767);
  c_Output Clob;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');
  n_��Ŀid := j_Json.Get_Number('fitem_id');

  For R In (Select ��Ŀid, ʹ������, ��������, (ʹ������ - ��������) As ��������
            From ����������Ŀ
            Where ����id = n_����id And ��ҳid = n_��ҳid And (Nvl(n_��Ŀid, 0) = 0 Or ��Ŀid = Nvl(n_��Ŀid, 0))) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'item_id', r.��Ŀid, 1, 1); --N
    zlJsonPutValue(v_Output, 'limit_quantity', r.ʹ������, 1); --N
    zlJsonPutValue(v_Output, 'used_quantity', r.��������, 1); --N
    zlJsonPutValue(v_Output, 'avail_quantity', r.��������, 1, 2); --N
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getneedaudititems;
/



Create Or Replace Procedure Zl_Exsesvr_Adviceisexist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  --����ҽ��ID��ѯ�Ƿ��ڷ��ñ���ڼ�¼
  ---------------------------------------------------------------------------
  --input      ����ҽ��ID��ѯҽ��״̬
  --  advice_ids  N  1  ���ҽ��ID���ö��ŷָ�
  --output
  --  code              C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message           C  1  Ӧ����Ϣ��
  --  advice_list       ҽ���б�[����]
  --     advice_id      N    ҽ��ID�����ڷ��õģ�
  --     fee_no         C    ����NO
  --     pati_id        N    ����id
  --     fee_properties N    ��¼����
  --     fee_status     N    ��¼״̬
  --     amount_id      N    ����id
  --     nums           N    ����
  --     packages_num   N    ����
  --     parent_num     N    �۸񸸺�
  --     receipt_type   C    �շ����
  --     receipt_id     N    �շ�ϸĿid
  --     cost_status    N    ����״̬
  --     stdd_price     N    ��׼����
  --     real_amount    N    ʵ�ս��

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_ҽ��id Clob; --��¼ҽ��id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  Col_ҽ��id Collection_Type;
  I          Number;

  v_Output Varchar2(32767);
  c_Output Clob;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_ҽ��id := j_Json.Get_String('advice_ids');

  --�� v_ҽ��id ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Num2list ��������
  I := 0;
  While v_ҽ��id Is Not Null Loop
    If Length(v_ҽ��id) <= 4000 Then
      Col_ҽ��id(I) := v_ҽ��id;
      v_ҽ��id := Null;
    Else
      Col_ҽ��id(I) := Substr(v_ҽ��id, 1, Instr(v_ҽ��id, ',', 3980) - 1);
      v_ҽ��id := Substr(v_ҽ��id, Instr(v_ҽ��id, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;
  I := 0;

  For I In 0 .. Col_ҽ��id.Count - 1 Loop
    For v_ҽ������ In (Select Distinct ҽ�����, ����no, ����id, ��¼����, ��¼״̬, ����id, ����, ����, �۸񸸺�, �շ����, �շ�ϸĿid, ����״̬, ��׼����, ʵ�ս��
                   From (Select /*+cardinality(b,10)*/
                           ҽ�����, NO As ����no, ����id, ��¼����, ��¼״̬, ����id, ����, ����, �۸񸸺�, �շ����, �շ�ϸĿid, ����״̬, ��׼����, ʵ�ս��
                          From ������ü�¼ A, Table(f_Num2List(Col_ҽ��id(I))) B
                          Where a.ҽ����� = b.Column_Value
                          Union All
                          Select /*+cardinality(b,10)*/
                           ҽ�����, NO As ����no, ����id, ��¼����, ��¼״̬, ����id, ����, ����, �۸񸸺�, �շ����, �շ�ϸĿid, ����״̬, ��׼����, ʵ�ս��
                          From סԺ���ü�¼ A, Table(f_Num2List(Col_ҽ��id(I))) B
                          Where a.ҽ����� = b.Column_Value)) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'advice_id', v_ҽ������.ҽ�����, 1, 1);
      zlJsonPutValue(v_Output, 'fee_no', v_ҽ������.����no);
      zlJsonPutValue(v_Output, 'pati_id', v_ҽ������.����id, 1);
      zlJsonPutValue(v_Output, 'fee_properties', v_ҽ������.��¼����, 1);
      zlJsonPutValue(v_Output, 'fee_status', v_ҽ������.��¼״̬, 1);
      zlJsonPutValue(v_Output, 'amount_id', v_ҽ������.����id, 1);
      zlJsonPutValue(v_Output, 'nums', v_ҽ������.����, 1);
      zlJsonPutValue(v_Output, 'packages_num', v_ҽ������.����, 1);
      zlJsonPutValue(v_Output, 'parent_num', v_ҽ������.�۸񸸺�, 1);
      zlJsonPutValue(v_Output, 'receipt_type', v_ҽ������.�շ����);
      zlJsonPutValue(v_Output, 'receipt_id', v_ҽ������.�շ�ϸĿid, 1);
      zlJsonPutValue(v_Output, 'cost_status', v_ҽ������.����״̬, 1);
      zlJsonPutValue(v_Output, 'stdd_price', v_ҽ������.����״̬, 1);
      zlJsonPutValue(v_Output, 'real_amount', v_ҽ������.ʵ�ս��, 1, 2);
    
    End Loop;
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","advice_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","advice_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Adviceisexist;
/

Create Or Replace Procedure Zl_Exsesvr_Getmrbkfeeinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݲ���id��ȡָ���������漰�Ĳ������嵥��Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id  N  1  ����id
  --    fee_no  C    ���ݺ�:���������漰�ĵ��ݺ�
  --    rec_status  N  1  ��¼״̬:1-ԭʼ��¼;2-��������
  --
  --����: Json_Out,��ʽ����
  --output
  --   code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message  C  1  Ӧ����Ϣ��
  --   details_list[]  C    ������ϸ����
  --    fee_no  C  1  ���ݺ�
  --    fee_num  N  1  ���
  --    pati_id  N  1  ����id
  --    pati_name  C  1  ����
  --    pati_sex  C  1  �Ա�
  --    pati_age  C  1  ����
  --    fee_category  C  1  �ѱ�
  --    fee_status  N  1  ����״̬:1-�쳣״̬;0-��������
  --    rec_status  N  1  ��¼״̬:1-������¼;2-���ʼ�¼;3-�����ʵļ�¼
  --    charge_sign  N  1  �շѱ�־:0-����;1-����;2-���۵�
  --    fee_ampaid  N  1  ʵ�ս��
  --    happen_time  C  1  ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --    operator_name  C  1  ����Ա����
  --    memo  C  1  ժҪ
  --    pricebill_no  C  1  ���۵���
  --    price_charged  N  1  �������շ�:1-���۵��Ѿ����շѴ����շ�;0-δ�շ�
  --    balance_info  C    ������Ϣ
  --      blnc_mode  C  1  ���㷽ʽ����
  --      balance_id  N  1  ����ID����ѯ���ϵĵ���ʱΪ����ID
  --      blnc_money  N  1  ���ʽ��
  --      pay_cardno  N  1  ֧������
  --      pay_swapno  C  1  ������ˮ��
  --      pay_swapmemo  C  1  ����˵��
  --      relation_id  N  1  ��������id
  --      cardtype_id  N  1  �����id
  --      consume_card  N  1  �Ƿ����ѿ�:1-��;0-����
  --      blnc_nature  N  1  ��������:1-�ֽ���㷽ʽ,2-������ҽ������ , 8-���㿨����
  --      error_moeny N 1 �����
  --      blnc_statu  N  1  ����״̬:1-δ���ýӿ�;2-�ӿڵ��óɹ�,����δ�շ����,0-��������
  --      consume_card_id  N  1  ���ѿ�id
  --      blnc_no  C  1  �������
  --      blnc_memo  C  1  ժҪ
  --      original_id  N  1  ԭ����ID:����ʱ����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  n_����id   Number;
  v_���ݺ�   ������ü�¼.No%Type;
  n_��¼״̬ ������ü�¼.��¼״̬%Type;
  n_ʵ�ս�� ������ü�¼.ʵ�ս��%Type;
  v_��¼״̬ Varchar2(6);
  v_���۵�   Varchar2(100);
  n_�������� Number(2);
  n_�շѱ�־ Number(2);
  n_����id   Number(18);
  n_ԭ����id Number(18);
  n_Count    Number(18);
  v_Output   Varchar2(32767);
  v_Balance  Varchar2(32767);
  c_Output   Clob;

Begin

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  v_���ݺ�   := j_Json.Get_String('fee_no');
  n_��¼״̬ := j_Json.Get_Number('rec_status');

  v_��¼״̬ := ',1,';
  If Nvl(n_��¼״̬, 0) = 2 Then
    v_��¼״̬ := ',2,';
  End If;

  --�ȶ�ȡ����
  v_���۵� := Null;
  For r_���� In (Select a.No, a.��¼״̬, Nvl(a.�۸񸸺�, a.���) As ���, Max(a.�ѱ�) As �ѱ�, Max(a.����) As ����, Max(a.�Ա�) As �Ա�,
                      Max(a.����) As ����, Max(a.����id) As ����id, a.�շ�ϸĿid, Max(a.ʵ��Ʊ��) As ʵ��Ʊ��, Avg(a.����) As ����,
                      Sum(Decode(n_��¼״̬, 2, -1, 1) * a.Ӧ�ս��) As Ӧ�ս��, Sum(Decode(n_��¼״̬, 2, -1, 1) * a.ʵ�ս��) As ʵ�ս��,
                      Nvl(Max(a.���ʷ���), 0) As ���ʷ���, Nvl(Max(Nvl(a.�Ӱ��־, 0)), 0) As �䶯���, Max(a.����Ա����) As ����Ա����,
                      Max(a.����Ա���) As ����Ա���, To_Char(Max(a.�Ǽ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
                      To_Char(Max(a.����ʱ��), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(Max(a.���ӱ�־), 0), 8, 1, 0) As ������,
                      Max(a.����) As �����id, Max(a.����id) As ����id, Max(a.������) As ������, Max(a.����״̬) As ����״̬, Max(a.ժҪ) As ժҪ,
                      Max(a.ʵ��Ʊ��) As ����
               From סԺ���ü�¼ A
               Where a.��¼���� = 5 And a.���ӱ�־ = 8 And Instr(v_��¼״̬, ',' || a.��¼״̬ || ',') > 0 And ����id = n_����id And
                     ((v_���ݺ� Is Not Null And NO = v_���ݺ�) Or v_���ݺ� Is Null)
               Group By a.No, a.��¼״̬, Nvl(a.�۸񸸺�, a.���), �շ�ϸĿid
               Order By a.No, ���) Loop
    n_ԭ����id := Null;
    v_���۵�   := Null;
    n_����id   := r_����.����id;
    If Nvl(n_��¼״̬, 0) = 2 And Nvl(r_����.���ʷ���, 0) = 0 Then
      Select Max(����id)
      Into n_ԭ����id
      From סԺ���ü�¼
      Where ��¼���� = 5 And ��¼״̬ In (1, 3) And NO = r_����.No;
    End If;
  
    o_Json     := PLJson();
    n_ʵ�ս�� := Nvl(r_����.ʵ�ս��, 0);
    n_�շѱ�־ := Nvl(r_����.���ʷ���, 0);
    n_�������� := 0;
    If r_����.ժҪ Is Not Null And Nvl(n_�շѱ�־, 0) <> 1 Then
      v_���۵� := r_����.ժҪ;
      --һ������Ӧ�ò��࣬���Ե�����ѯ������Ӱ�첻��
      Select Count(1), Sum(ʵ�ս��), Decode(Max(��¼״̬), 0, 0, 1)
      Into n_Count, n_ʵ�ս��, n_��������
      From ������ü�¼
      Where NO = r_����.No And ��¼���� = 1 And Nvl(���ӱ�־, 0) = 8;
    
      If n_Count <> 0 Then
        n_�շѱ�־ := 2; --����
      
      Else
        n_ʵ�ս�� := Nvl(r_����.ʵ�ս��, 0);
        v_���۵�   := Null;
      End If;
    
    End If;
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      c_Output := Nvl(c_Output, '') || To_Clob(v_Output);
      v_Output := '';
    End If;
  
    --1.ȡ������Ϣ
    zlJsonPutValue(v_Output, 'fee_no', r_����.No, 0, 1);
    zlJsonPutValue(v_Output, 'fee_num', r_����.���, 1);
  
    zlJsonPutValue(v_Output, 'pati_id', r_����.����id, 1);
  
    zlJsonPutValue(v_Output, 'pati_name', r_����.����);
    zlJsonPutValue(v_Output, 'pati_sex', Nvl(r_����.�Ա�, ''));
  
    zlJsonPutValue(v_Output, 'pati_age', Nvl(r_����.����, ''));
  
    zlJsonPutValue(v_Output, 'fee_category', Nvl(r_����.�ѱ�, ''));
  
    zlJsonPutValue(v_Output, 'fee_status', Nvl(r_����.����״̬, 0), 1);
    zlJsonPutValue(v_Output, 'rec_status', Nvl(r_����.��¼״̬, 0), 1);
    zlJsonPutValue(v_Output, 'kpbooks_sign', Nvl(n_�շѱ�־, 0), 1);
    zlJsonPutValue(v_Output, 'fee_ampaid', Nvl(n_ʵ�ս��, 0), 1);
    zlJsonPutValue(v_Output, 'happen_time', Nvl(r_����.����ʱ��, ''));
    zlJsonPutValue(v_Output, 'operator_name', Nvl(r_����.����Ա����, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(r_����.ժҪ, ''));
    zlJsonPutValue(v_Output, 'pricebill_no', Nvl(v_���۵�, ''));
  
    zlJsonPutValue(v_Output, 'price_charged', Nvl(n_��������, 0), 1);
  
    --��ȡ������Ϣ
    If Nvl(r_����.���ʷ���, 0) = 0 Then
      v_Balance := '';
      For r_������Ϣ In (
                     
                     Select a.No, Max(Decode(Nvl(b.����, 0), 9, '', a.���㷽ʽ)) As ���㷽ʽ,
                             Sum(Decode(Nvl(b.����, 0), 9, 0, 1) * Decode(n_��¼״̬, 2, -1, 1) * Nvl(a.��Ԥ��, 0)) As ��Ԥ��,
                             Sum(Decode(Nvl(b.����, 0), 9, 1, 0) * Decode(n_��¼״̬, 2, -1, 1) * Nvl(a.��Ԥ��, 0)) As ����,
                             Max(Decode(Nvl(b.����, 0), 9, 0, a.��������id)) As ��������id, Max(a.�����id) As �����id, Max(a.����) As ����,
                             Max(a.���㿨���) As ���㿨���, Max(a.������ˮ��) As ������ˮ��, Max(a.����˵��) As ����˵��,
                             Max(Decode(Nvl(b.����, 0), 9, -1, Nvl(b.����, 0))) As ����, Max(a.У�Ա�־) As У�Ա�־,
                             Max(c.���ѿ�id) As ���ѿ�id, Max(a.�������) As �������, Max(a.ժҪ) As ժҪ
                     From ����Ԥ����¼ A, ���㷽ʽ B, ���˿������¼ C
                     Where a.���㷽ʽ = ����(+) And a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And a.Id = c.����id(+)
                     Group By NO) Loop
      
        zlJsonPutValue(v_Balance, 'blnc_mode', r_������Ϣ.���㷽ʽ, 0, 1);
        zlJsonPutValue(v_Balance, 'balance_id', n_����id, 1);
        zlJsonPutValue(v_Balance, 'blnc_money', Nvl(r_������Ϣ.��Ԥ��, 0), 1);
        zlJsonPutValue(v_Balance, 'pay_cardno', Nvl(r_������Ϣ.����, ''));
        zlJsonPutValue(v_Balance, 'pay_swapno', Nvl(r_������Ϣ.������ˮ��, ''));
        zlJsonPutValue(v_Balance, 'pay_swapmemo', Nvl(r_������Ϣ.����˵��, ''));
      
        zlJsonPutValue(v_Balance, 'relation_id', Nvl(r_������Ϣ.��������id, 0), 1);
      
        If Nvl(r_������Ϣ.���㿨���, 0) <> 0 Then
          zlJsonPutValue(v_Balance, 'cardtype_id', Nvl(r_������Ϣ.���㿨���, 0), 1);
          zlJsonPutValue(v_Balance, 'consume_card', 1, 1);
        Else
          zlJsonPutValue(v_Balance, 'cardtype_id', Nvl(r_������Ϣ.�����id, 0), 1);
          zlJsonPutValue(v_Balance, 'consume_card', 0, 1);
        End If;
      
        zlJsonPutValue(v_Balance, 'consume_card_id', Nvl(r_������Ϣ.���ѿ�id, 0), 1);
        zlJsonPutValue(v_Balance, 'error_moeny', Nvl(r_������Ϣ.����, 0), 1);
        zlJsonPutValue(v_Balance, 'blnc_nature', Nvl(r_������Ϣ.����, 0), 1);
        zlJsonPutValue(v_Balance, 'blnc_statu', Nvl(r_������Ϣ.У�Ա�־, 0), 1);
        zlJsonPutValue(v_Balance, 'blnc_no', Nvl(r_������Ϣ.�������, ''));
        zlJsonPutValue(v_Balance, 'blnc_memo', Nvl(r_������Ϣ.ժҪ, ''));
        zlJsonPutValue(v_Balance, 'original_id', Nvl(n_ԭ����id, 0), 1, 2);
        v_Balance := ',"balance_info":' || v_Balance;
        Exit;
      End Loop;
    Else
      v_Balance := Null;
    End If;
    v_Output := v_Output || Nvl(v_Balance, '') || '}';
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getmrbkfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Checkunauditedfee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------------------------------------
  --���ܣ���鲡���Ƿ������δ��Ч�ļ�����Ŀ
  --��Σ�Json_In,��ʽ����
  --  input
  --    pati_id            N 1 ����ID
  --    pati_pageid        N 1 ��ҳID
  --����: Json_Out,��ʽ����
  --  output
  --    code               N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message            C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    exist              N 1 ִ�б��:0-������;1-����
  --------------------------------------------------------------------------------------------------------------------
  j_Input       PLJson;
  j_Json        PLJson;
  v_Output      Varchar2(400);
  n_Count       Number;
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');

  --���ò��˱���סԺ�Ƿ��Ѿ������
  Select Count(*)
  Into n_Count
  From סԺ���ü�¼
  Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And ���ʷ��� = 1 And ��¼״̬ = 0 And Rownum <= 2;

  If n_Count > 0 Then
    n_Count := 1;
  End If;

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'exist', n_Count, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkunauditedfee;
/


Create Or Replace Procedure Zl_Exsesvr_Getfeebillbycardno
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݿ��Ż�ȡ���õ�����Ϣ
  --��Σ�Json_In:��ʽ
  --   input
  --    pati_id  N  1  ����id
  --    cardtype_id  N  1  �����id
  --    cardno  C  1  ����

  --����: Json_Out,��ʽ����
  --  output
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    feeno  C  1  ���ѵ���
  --    charge_sign  N  1  �շѱ�־:1-�Ѿ��շ���;2-�Ѿ��˷�

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id   Number(18);
  n_�����id Number(18);
  v_����     Varchar2(100);
  v_���ݺ�   סԺ���ü�¼.No%Type;
  n_��¼״̬ סԺ���ü�¼.��¼״̬%Type;
  v_ժҪ     סԺ���ü�¼.ժҪ%Type;
  n_���ʷ��� סԺ���ü�¼.���ʷ���%Type;
  v_���۵�   סԺ���ü�¼.No%Type;
  n_Temp     Number(18);
  n_Count    Number(18);
  n_�շѱ�־ Number(2);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  n_�����id := j_Json.Get_Number('cardtype_id');
  v_����     := j_Json.Get_String('cardno');

  n_�շѱ�־ := 1;
  Select Max(NO), Max(��¼״̬), Max(ժҪ), Max(���ʷ���)
  Into v_���ݺ�, n_��¼״̬, v_ժҪ, n_���ʷ���
  From סԺ���ü�¼
  Where ��¼���� = 5 And ����id = n_����id And ʵ��Ʊ�� = v_���� And To_Number(Nvl(����, '0')) = Nvl(n_�����id, 0) And Nvl(���ӱ�־, 0) <> 8;

  If Nvl(n_��¼״̬, 0) = 0 Then
    n_�շѱ�־ := 0;
  Elsif Nvl(n_��¼״̬, 0) = 1 Then
  
    n_�շѱ�־ := 1;
  Else
    n_�շѱ�־ := 2;
  End If;

  If Nvl(n_���ʷ���, 0) <> 1 And v_ժҪ Is Not Null And Nvl(n_��¼״̬, 0) = 1 Then
  
    Select Max(��¼״̬), Max(NO)
    Into n_Temp, v_���۵�
    From ������ü�¼
    Where ��¼���� = 1 And NO = v_ժҪ And �۸񸸺� Is Null And Nvl(���ӱ�־, 0) <> 8;
    If v_���۵� Is Not Null Then
      If Nvl(n_Temp, 0) = 0 Then
        n_�շѱ�־ := 0; --δ�շ�
      Elsif Nvl(n_Temp, 0) <> 1 Then
        Select Count(1)
        Into n_Count
        From (Select NO, ���, Sum(Nvl(����, 1) * ����) As ʣ����, Max(��¼״̬) As ��¼״̬
               From ������ü�¼
               Where Mod(��¼����, 10) = 1 And NO = v_ժҪ And �۸񸸺� Is Null And Nvl(���ӱ�־, 0) <> 8 Having
                Sum(Nvl(����, 1) * ����) <> 0
               Group By NO, ���);
        If n_Count = 0 Then
          n_�շѱ�־ := 2; --���˷�
        End If;
      Else
        n_�շѱ�־ := 1; --���շ�
      End If;
    End If;
  End If;
  Json_Out := '{"output":{"code":1,"message":"' || '�ɹ�' || '"';
  Json_Out := Json_Out || ',"feeno":"' || Nvl(v_���ݺ�, '') || '"';
  Json_Out := Json_Out || ',"priceno":"' || Nvl(v_���۵�, '') || '"';
  Json_Out := Json_Out || ',"charge_sign":' || Nvl(n_�շѱ�־, 0);
  Json_Out := Json_Out || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeebillbycardno;
/

Create Or Replace Procedure Zl_Exsesvr_Registerinpatient
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ�������Ժ�ǼǷ�����ش���
  --��Σ�Json_In:��ʽ
  --    input
  --      pati_id            N 1  ����id
  --      pati_pageid        N 1  ��ҳID
  --      type               N 1  �Ǽ�ģʽ=0-�����Ǽ�,1-ԤԼ�Ǽ�,2-����ԤԼ
  --      pati_deptid        N 1 ��Ժ����ID
  --����: Json_Out,��ʽ����
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Type        Number(3);
  n_Pati_Deptid Number(18);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  n_Type        := j_Json.Get_Number('type');
  n_Pati_Deptid := j_Json.Get_Number('pati_deptid');

  If n_Type <> 1 Then
    --���˵�����¼
    Update ���˵�����¼
    Set ����ʱ�� = Sysdate
    Where ����id = n_Pati_Id And ����ʱ�� Is Not Null And ����ʱ�� > Sysdate;
    --���˷���������Ŀ
    Delete From ����������Ŀ Where ����id = n_Pati_Id;
  End If;

  If n_Type = 2 Then
    Update ����Ԥ����¼
    Set ��ҳid = n_Pati_Pageid
    Where ����id = n_Pati_Id And ��ҳid Is Null And ����id = n_Pati_Deptid And Ԥ����� = 2 And ��Ԥ�� Is Null And
          Trunc(�տ�ʱ��) = Trunc(Sysdate);
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Registerinpatient;
/

Create Or Replace Procedure Zl_Exsesvr_Unregisterinpatient
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --------------------------------------------------------------------------- 
  --���ܣ�ȡ��������Ժ�ǼǷ�����ش��� 
  --��Σ�Json_In:��ʽ 
  --    input 
  --      pati_id            N 1  ����id 
  --      pati_pageid        N 1  ��ҳID 
  --����: Json_Out,��ʽ���� 
  --    output 
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ� 
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ 
  --------------------------------------------------------------------------- 
  j_Input Pljson;
  j_Json  Pljson;

  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);
  n_Count       Number;
  n_Money       Number(16, 5);
Begin
  --������� 
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');

  Select Sum(���) Into n_Money From ����Ԥ����¼ Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid;
  If n_Money <> 0 Then
    Json_Out := Zljsonout('���˱���סԺ��Ԥ����δ����,�봦�����ִ�д˲�����');
    Return;
  End If;
  Select Sum(���) Into n_Money From ����δ����� Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid;
  If n_Money <> 0 Then
    Json_Out := Zljsonout('���˱���סԺ��δ�����,�봦�����ִ�д˲�����');
    Return;
  End If;

  --����סԺ�������Ԥ����,��Ϊ�������ｻ�� 
  Update ����Ԥ����¼ Set ��ҳid = Null Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid;

  --���η�����,�ı����﷢�� 
  Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And ��¼���� = 5;

  --����סԺ�����з��ü�¼�޽�������ȫ���������򽫶�Ӧ���ü�¼�е�"��ҳID"����� 
  n_Count := 0;
  Select Nvl(Count(*), 0)
  Into n_Count
  From סԺ���ü�¼
  Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And ���ʷ��� = 1 And ����id Is Not Null;

  If n_Count = 0 Then
    Begin
      Select Nvl(Count(*), 0)
      Into n_Count
      From סԺ���ü�¼
      Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And ���ʷ��� = 1
      Group By NO, ��¼����, ���
      Having Nvl(Sum(ʵ�ս��), 0) <> 0;
    Exception
      When Others Then
        n_Count := 0;
    End;
  
    If n_Count = 0 Then
      Delete ����δ����� Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And ��� = 0;
      Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = n_Pati_Id And ��ҳid = n_Pati_Pageid And ���ʷ��� = 1;
    End If;
  End If;
  Json_Out := Zljsonout('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Unregisterinpatient;
/


Create Or Replace Procedure Zl_Exsesvr_Updatepatisurety
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ����˵�����¼���������¡�ɾ������
  --��Σ�Json_In:��ʽ
  --    input
  --      func_id          N 1 ����ID 1-����;2-����;3-ɾ��
  --      pati_id          N 1 ����id
  --      pati_pageid      N 1 ��ҳID
  --      guarantor        c 1 ������
  --      garnt_amount     N 1 ������
  --      garnt_prop       N 1 ��������
  --      garnt_reason     c 1 ����ԭ��
  --      due_time         c 1 ����ʱ��
  --      operator_code    c 1 ����Ա���
  --      operator_name    c 1 ����Ա����
  --      create_time      C 0 �Ǽ�ʱ��   ����ʱ�����ֵ
  --����: Json_Out,��ʽ����
  --    output
  --        code                    N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --        message                 C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input       PLJson;
  j_Json        PLJson;
  n_Func_Id     Number(3);
  n_Pati_Id     Number(18);
  n_Pati_Pageid Number(18);

  v_������   ���˵�����¼.������%Type;
  n_������   ���˵�����¼.������ %Type;
  n_�������� ���˵�����¼.��������%Type;
  v_����ԭ�� ���˵�����¼.����ԭ��%Type;

  d_����ʱ��   ���˵�����¼.����ʱ��%Type;
  v_����Ա��� ���˵�����¼.����Ա���%Type;
  v_����Ա���� ���˵�����¼.����Ա����%Type;
  d_�Ǽ�ʱ��   ���˵�����¼.�Ǽ�ʱ��%Type;

Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_Func_Id     := j_Json.Get_Number('func_id');
  n_Pati_Id     := j_Json.Get_Number('pati_id');
  n_Pati_Pageid := j_Json.Get_Number('pati_pageid');
  If n_Func_Id = 1 Or n_Func_Id = 2 Then
    v_������   := j_Json.Get_String('guarantor');
    n_������   := j_Json.Get_Number('garnt_amount');
    n_�������� := j_Json.Get_Number('garnt_prop');
    v_����ԭ�� := j_Json.Get_String('garnt_reason');
    d_����ʱ�� := To_Date(j_Json.Get_String('due_time'), 'YYYY-MM-DD HH24:MI:SS');
  End If;
  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  If n_Func_Id = 2 Or n_Func_Id = 3 Then
    d_�Ǽ�ʱ�� := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD HH24:MI:SS');
  End If;
  If n_Func_Id = 1 Then
    Zl_���˵�����¼_Insert(n_Pati_Id, n_Pati_Pageid, v_������, n_������, n_��������, v_����ԭ��, Null, d_����ʱ��, v_����Ա���, v_����Ա����);
  Elsif n_Func_Id = 2 Then
    Zl_���˵�����¼_Update(n_Pati_Id, n_Pati_Pageid, v_������, n_������, n_��������, v_����ԭ��, Null, d_����ʱ��, v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��);
  Elsif n_Func_Id = 3 Then
    Zl_���˵�����¼_Delete(n_Pati_Id, n_Pati_Pageid, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����);
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updatepatisurety;
/

Create Or Replace Procedure Zl_Exsesvr_Getorderfeeinfo
(
  Json_In  In Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     query_type         N 1 ��ѯ��ʽ���봫��
  --                               1-��ѯҩռ��,������ν�㣨pati_id+pati_pageid�������� fee_ratio ��ռ�������� 68.5%������һλС�����ַ���
  --                               2-����ҽ��id��ȡδ���ʷ��õ�ҽ��,������ν�㣨pati_id+pati_pageid+baby_num��
  --                               3-����ҽ��id��ȡ������ػ��ܽ��,������ν�㣨order_ids��
  --                               4-����ҽ��id��ȡ���ü�¼������Ϣҽ���嵥�·��ķ����б�,������ν�㣨fee_origin+order_ids��
  --                               5-����ҩƷҽ��ID��ȡ���һ�η��͵ķ��ö�Ӧ���շ���ĿID(ҩƷ���ID)������һ���շ�ϸĿid,������ν�㣨fee_no+bill_prop+order_id+fee_origin��
  --                               6-��ȡҽ����Ӧδ��˵ļ��ʷ��úϼƣ�������ν�㣨fee_origin+fee_no������ʱfee_noΪ������ݺ�ƴ��
  --                               7-���ݷ���id�ж��Ƿ����Ѿ��շѣ�������ν��(fee_origin+fee_ids)�������շѵķ���id����ƴ��
  --                               8-���ݷ�����Դ��ҽ��id��ȡ������ϸ�б�������ν��(fee_origin+fee_no+order_ids)��ҽ��վִ�п��ұ��ʱ�����
  --                               9-���ݷ�����Դ��ҽ��id��ȡ������ϸ�б�������ν��(fee_origin+fee_no+order_ids)����ʿվ���ʺ���ó����쳣�ٴν����쳣�޸�ʱ��ȡ����                             
  --     pati_id            N 0 ����id
  --     pati_pageid        N 0 ��ҳid
  --     baby_num           N 0 Ӥ�����
  --     order_ids          C 0 ҽ��IDƴ��
  --     fee_origin         N 0 ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
  --     fee_no             C 0 ���ݺ�
  --     bill_prop          N 0 ��¼����
  --     order_id           N 0 ҽ��id
  --     fee_ids            C 0 ����id����ƴ��
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fee_ratio               C 1 ҩƷ����ռ�������� 68.5%������һλС�����ַ���[query_type=1ʱ���ش˽��]
  --    order_ids               C 1 ҽ��IDƴ����[query_type=2ʱ���ش˽��]
  --    fee_ampaib              N 1 ʵ�ս�[query_type=6ʱ���ش˽��]
  --    fee_ids                 C 0 ����id����ƴ����[query_type=7ʱ���ش˽��]
  --    fee_am_list[]query_type=3ʱ���ش��б�
  --      fee_amrcvb            N  1 Ӧ�ս��
  --      fee_ampaib            N  1 ʵ�ս��
  --      drug_amrcvb           N  1 ҩƷӦ�ս��
  --      drug_ampaib           N  1 ҩƷʵ�ս��
  --    fee_od_list[]query_type=4ʱ�д��б�
  --         order_id           N 1 ҽ��id
  --         fee_no             C 1 ���õ��ݺ�
  --         bill_prop          N 1 ��¼����
  --         exe_state          N 1 ���ü�¼ִ��״̬
  --         rec_state          N 1 ��¼״̬�����ü�¼�ļ�¼״̬
  --         fee_state          N 1 ����״̬�����ü�¼��
  --         exe_dept_id        N 1 ִ�в���id
  --         exe_dept_name      C 1 ִ�п������ƣ����ü�¼ִ�в���id��Ӧ������
  --         fee_type           C 1 �շ����
  --         nums               N 1 ����
  --         fee_item_id        N 1 �շ�ϸĿid
  --         unit               C 1 ��������ת����ĵ�λ��ҩƷ���� ��Ӧ�İ�װ������ͨ�� �շ���ĿĿ¼ �� ���㵥λ
  --         fee_name           C 1 �շ���Ŀ�����ƣ���ϵͳ�λ�ȡ�˵ġ�����ҩƷ��ʾ��
  --    fee_dept_list[]query_type=8ʱ�д��б�
  --       fee_id               N 1 ����id
  --       exe_dept_id          N 1 ִ�в���id
  --    fee_pivas_list[]�����ҩƷ��ҽ���ķ����б�
  --       fee_id               N 1 ����id
  --       exe_dept_id          N 1 ִ�п���id
  --       drug_id              N 1 �շ�ϸĿid
  --       quantity             N 1 ʣ������
  --       order_id             N 1 ҽ��id,ҽ�����;
  --       fee_no               C 1 ���õ��ݺ�
  --       fee_origin           N 1 ����Դ����1-������ñ�2-סԺ���ñ�
  --       serial_num           N 1 ���
  ---------------------------------------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  n_����id      Number;
  n_��ҳid      Number;
  n_��ѯ��ʽ    Number;
  n_ҩƷ��ʾ    Number;
  v_Vals        Clob;
  l_Vals        t_Strlist;
  n_Baby        Number;
  v_Order_Ids   Varchar2(32767);
  n_Tmp         Number;
  n_Fee_Ampaib  Number;
  n_ҽ��id      Number;
  n_��Դ        Number; --1-���2-����
  v_Fee_No      Varchar2(2000);
  n_��¼����    Number;
  v_Jtmp        Varchar2(32767);
  v_Fee_Ids     Varchar2(32767);
  v_Fee_Ids_Out Varchar2(32767);
  c_Jtmp        Clob;

  Cursor c_Feeoutone(Pҽ��id Number) Is
    Select a.ҽ����� As ҽ��id, a.��¼����, a.No, a.ִ��״̬, a.��¼״̬, a.����״̬, a.����id, a.�������, a.ִ�в���id, c.���� As ִ�п���, a.�շ����,
           (a.���� * a.����) ʣ������, (a.���� * a.���� / Nvl(a.�����װ, 1)) As ��������, a.�շ�ϸĿid,
           Decode(Nvl(Instr('567', a.�շ����), 0), 0, Decode(a.�շ����, '4', b.���㵥λ, b.���㵥λ), a.���ﵥλ) As ��λ,
           Nvl(g.����, b.����) || Decode(b.����, Null, Null, '(' || b.���� || ')') || Decode(b.���, Null, Null, ' ' || b.���) As �շ���Ŀ
    From (Select a.ҽ�����, Min(a.��¼����) As ��¼����, a.No, a.ִ��״̬, Min(a.��¼״̬) As ��¼״̬, Min(a.����״̬) As ����״̬, Min(a.����id) As ����id,
                  a.��� As �������, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.�����װ, b.���ﵥλ, b.����ϵ��
           From ������ü�¼ A, ҩƷ��� B
           Where a.��¼״̬ In (0, 1, 3) And a.�۸񸸺� Is Null And a.�շ�ϸĿid = b.ҩƷid(+) And a.ҽ����� = Pҽ��id
           Group By a.ҽ�����, a.No, a.ִ��״̬, a.���, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.�����װ, b.���ﵥλ, b.����ϵ��) A,
         �շ���ĿĿ¼ B, �շ���Ŀ���� G, ���ű� C
    Where a.ִ�в���id = c.Id(+) And a.�շ�ϸĿid = b.Id And a.�շ�ϸĿid = g.�շ�ϸĿid(+) And g.����(+) = 1 And g.����(+) = n_ҩƷ��ʾ
    Order By a.�������;

  Type t_Fee Is Table Of c_Feeoutone%RowType;
  r_Fee t_Fee;

  Cursor c_Feeout(Pҽ��ids Varchar2) Is
    Select a.ҽ����� As ҽ��id, a.��¼����, a.No, a.ִ��״̬, a.��¼״̬, a.����״̬, a.����id, a.�������, a.ִ�в���id, c.���� As ִ�п���, a.�շ����,
           (a.���� * a.����) ʣ������, (a.���� * a.���� / Nvl(a.�����װ, 1)) As ��������, a.�շ�ϸĿid,
           Decode(Nvl(Instr('567', a.�շ����), 0), 0, Decode(a.�շ����, '4', b.���㵥λ, b.���㵥λ), a.���ﵥλ) As ��λ,
           Nvl(g.����, b.����) || Decode(b.����, Null, Null, '(' || b.���� || ')') || Decode(b.���, Null, Null, ' ' || b.���) As �շ���Ŀ
    From (Select a.ҽ�����, Min(a.��¼����) As ��¼����, a.No, a.ִ��״̬, Min(a.��¼״̬) As ��¼״̬, Min(a.����״̬) As ����״̬, Min(a.����id) As ����id,
                  a.��� As �������, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.�����װ, b.���ﵥλ, b.����ϵ��
           From ������ü�¼ A, ҩƷ��� B
           Where a.��¼״̬ In (0, 1, 3) And a.�۸񸸺� Is Null And a.�շ�ϸĿid = b.ҩƷid(+) And
                 a.ҽ����� In (Select /*+cardinality(x,10)*/
                             x.Column_Value
                            From Table(Cast(f_Num2list(Pҽ��ids) As Zltools.t_Numlist)) X)
           Group By a.ҽ�����, a.No, a.ִ��״̬, a.���, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.�����װ, b.���ﵥλ, b.����ϵ��) A,
         �շ���ĿĿ¼ B, �շ���Ŀ���� G, ���ű� C
    Where a.ִ�в���id = c.Id(+) And a.�շ�ϸĿid = b.Id And a.�շ�ϸĿid = g.�շ�ϸĿid(+) And g.����(+) = 1 And g.����(+) = n_ҩƷ��ʾ
    Order By a.�������;

  Cursor c_Feeinone(Pҽ��id Number) Is
    Select a.ҽ����� As ҽ��id, a.��¼����, a.No, a.ִ��״̬, a.��¼״̬, a.����״̬, a.����id, a.�������, a.ִ�в���id, c.���� As ִ�п���, a.�շ����,
           (a.���� * a.����) ʣ������, (a.���� * a.���� / Nvl(a.סԺ��װ, 1)) As ��������, a.�շ�ϸĿid,
           Decode(Nvl(Instr('567', a.�շ����), 0), 0, Decode(a.�շ����, '4', b.���㵥λ, b.���㵥λ), a.סԺ��λ) As ��λ,
           Nvl(g.����, b.����) || Decode(b.����, Null, Null, '(' || b.���� || ')') || Decode(b.���, Null, Null, ' ' || b.���) As �շ���Ŀ
    From (Select a.ҽ�����, Min(a.��¼����) As ��¼����, a.No, a.ִ��״̬, Min(a.��¼״̬) As ��¼״̬, Min(a.����״̬) As ����״̬, Min(a.����id) As ����id,
                  a.��� As �������, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.סԺ��װ, b.סԺ��λ, b.����ϵ��
           From סԺ���ü�¼ A, ҩƷ��� B
           Where a.��¼״̬ In (0, 1, 3) And a.�۸񸸺� Is Null And a.�շ�ϸĿid = b.ҩƷid(+) And a.ҽ����� = Pҽ��id
           Group By a.ҽ�����, a.No, a.ִ��״̬, a.���, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.סԺ��װ, b.סԺ��λ, b.����ϵ��) A,
         �շ���ĿĿ¼ B, �շ���Ŀ���� G, ���ű� C
    Where a.ִ�в���id = c.Id(+) And a.�շ�ϸĿid = b.Id And a.�շ�ϸĿid = g.�շ�ϸĿid(+) And g.����(+) = 1 And g.����(+) = n_ҩƷ��ʾ
    Order By a.�������;

  Cursor c_Feein(Pҽ��ids Varchar2) Is
    Select a.ҽ����� As ҽ��id, a.��¼����, a.No, a.ִ��״̬, a.��¼״̬, a.����״̬, a.����id, a.�������, a.ִ�в���id, c.���� As ִ�п���, a.�շ����,
           (a.���� * a.����) ʣ������, (a.���� * a.���� / Nvl(a.סԺ��װ, 1)) As ��������, a.�շ�ϸĿid,
           Decode(Nvl(Instr('567', a.�շ����), 0), 0, Decode(a.�շ����, '4', b.���㵥λ, b.���㵥λ), a.סԺ��λ) As ��λ,
           Nvl(g.����, b.����) || Decode(b.����, Null, Null, '(' || b.���� || ')') || Decode(b.���, Null, Null, ' ' || b.���) As �շ���Ŀ
    From (Select a.ҽ�����, Min(a.��¼����) As ��¼����, a.No, a.ִ��״̬, Min(a.��¼״̬) As ��¼״̬, Min(a.����״̬) As ����״̬, Min(a.����id) As ����id,
                  a.��� As �������, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.סԺ��װ, b.סԺ��λ, b.����ϵ��
           From סԺ���ü�¼ A, ҩƷ��� B
           Where a.��¼״̬ In (0, 1, 3) And a.�۸񸸺� Is Null And a.�շ�ϸĿid = b.ҩƷid(+) And
                 a.ҽ����� In (Select /*+cardinality(x,10)*/
                             x.Column_Value
                            From Table(Cast(f_Num2list(Pҽ��ids) As Zltools.t_Numlist)) X)
           Group By a.ҽ�����, a.No, a.ִ��״̬, a.���, a.ִ�в���id, a.�շ����, a.����, a.����, a.�շ�ϸĿid, b.סԺ��װ, b.סԺ��λ, b.����ϵ��) A,
         �շ���ĿĿ¼ B, �շ���Ŀ���� G, ���ű� C
    Where a.ִ�в���id = c.Id(+) And a.�շ�ϸĿid = b.Id And a.�շ�ϸĿid = g.�շ�ϸĿid(+) And g.����(+) = 1 And g.����(+) = n_ҩƷ��ʾ
    Order By a.�������;

Begin
  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ��ʽ  := j_Json.Get_Number('query_type');
  n_����id    := j_Json.Get_Number('pati_id');
  n_��ҳid    := j_Json.Get_Number('pati_pageid');
  n_Baby      := j_Json.Get_Number('baby_num');
  n_ҽ��id    := j_Json.Get_Number('order_id');
  n_��Դ      := j_Json.Get_Number('fee_origin');
  v_Fee_No    := j_Json.Get_String('fee_no');
  n_��¼����  := j_Json.Get_Number('bill_prop');
  v_Order_Ids := j_Json.Get_String('order_ids');

  If n_��ѯ��ʽ = 1 Then
    v_Order_Ids := '0.0%';
    For R In (Select (100 * (a.���з� - a.��ҩ��) / Nvl(a.���з�, 1)) As ����
              From (Select Sum(Decode(a.�շ����, '5', 0, '6', 0, '7', 0, a.ʵ�ս��)) As ��ҩ��, Sum(a.ʵ�ս��) As ���з�
                     From סԺ���ü�¼ A
                     Where a.����id = n_����id And a.��ҳid = n_��ҳid And a.��¼״̬ <> 0 Having Sum(a.ʵ�ս��) > 0) A) Loop
      If r.���� > 0 Then
        v_Order_Ids := Round(r.����, 1) || '%';
      End If;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ratio":"' || v_Order_Ids || '"}}';
  Elsif n_��ѯ��ʽ = 2 Then
    For R In (Select a.ҽ����� As ҽ��id
              From סԺ���ü�¼ A
              Where a.����id = n_����id And a.��ҳid = n_��ҳid And a.��¼״̬ = 0 And (n_Baby Is Null Or Nvl(a.Ӥ����, 0) = n_Baby) And
                    a.ҽ����� Is Not Null
              Group By a.ҽ�����) Loop
      v_Order_Ids := v_Order_Ids || ',' || r.ҽ��id;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","order_ids":"' || Substr(v_Order_Ids, 2) || '"}}';
  Elsif n_��ѯ��ʽ = 3 Then
    For R In (Select Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��, Sum(Decode(Instr('567', a.�շ����), 0, 0, a.Ӧ�ս��)) As ҩƷӦ��,
                     Sum(Decode(Instr('567', a.�շ����), 0, 0, a.ʵ�ս��)) As ҩƷʵ��
              From ������ü�¼ A
              Where a.ҽ����� In (Select /*+cardinality(x,10)*/
                                x.Column_Value
                               From Table(Cast(f_Num2list(v_Order_Ids) As Zltools.t_Numlist)) X)) Loop
    
      n_Tmp  := r.Ӧ�ս��;
      v_Jtmp := v_Jtmp || '"fee_amrcvb":' || Zljsonstr(n_Tmp, 1);
    
      n_Tmp  := r.Ӧ�ս��;
      v_Jtmp := v_Jtmp || ',"fee_ampaib":' || Zljsonstr(n_Tmp, 1);
    
      n_Tmp  := r.ҩƷӦ��;
      v_Jtmp := v_Jtmp || ',"drug_amrcvb":' || Zljsonstr(n_Tmp, 1);
    
      n_Tmp  := r.ҩƷʵ��;
      v_Jtmp := v_Jtmp || ',"drug_ampaib":' || Zljsonstr(n_Tmp, 1);
    
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_am_list":[{' || v_Jtmp || '}]}}';
  Elsif n_��ѯ��ʽ = 4 Then
    n_ҩƷ��ʾ := zl_GetSysParameter('����ҩƷ��ʾ');
    n_ҩƷ��ʾ := Nvl(n_ҩƷ��ʾ, 2);
    If Instr(v_Order_Ids, ',') = 0 Then
      n_ҽ��id := v_Order_Ids;
    End If;
  
    If n_��Դ = 1 Then
      If Nvl(n_ҽ��id, 0) <> 0 Then
        Open c_Feeoutone(n_ҽ��id);
        Fetch c_Feeoutone Bulk Collect
          Into r_Fee;
        Close c_Feeoutone;
      Else
        Open c_Feeout(v_Order_Ids);
        Fetch c_Feeout Bulk Collect
          Into r_Fee;
        Close c_Feeout;
      End If;
    Else
      If Nvl(n_ҽ��id, 0) <> 0 Then
        Open c_Feeinone(n_ҽ��id);
        Fetch c_Feeinone Bulk Collect
          Into r_Fee;
        Close c_Feeinone;
      Else
        Open c_Feein(v_Order_Ids);
        Fetch c_Feein Bulk Collect
          Into r_Fee;
        Close c_Feein;
      End If;
    End If;
    v_Jtmp := Null;
    For I In 1 .. r_Fee.Count Loop
      v_Jtmp := v_Jtmp || ',{"order_id":' || r_Fee(I).ҽ��id;
      v_Jtmp := v_Jtmp || ',"fee_no":"' || r_Fee(I).No || '"';
      v_Jtmp := v_Jtmp || ',"bill_prop":' || Nvl(r_Fee(I).��¼����, 0);
      v_Jtmp := v_Jtmp || ',"exe_state":' || Nvl(r_Fee(I).ִ��״̬, 0);
      v_Jtmp := v_Jtmp || ',"rec_state":' || Nvl(r_Fee(I).��¼״̬, 0);
      v_Jtmp := v_Jtmp || ',"fee_state":' || Nvl(r_Fee(I).����״̬, 0);
      v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r_Fee(I).ִ�в���id || '', 'null'); --ִ�в���id
      v_Jtmp := v_Jtmp || ',"exe_dept_name":"' || Zljsonstr(r_Fee(I).ִ�п���) || '"'; --C 1 ִ�п������ƣ�
      v_Jtmp := v_Jtmp || ',"fee_type":"' || r_Fee(I).�շ���� || '"'; --C 1 �շ����
      v_Jtmp := v_Jtmp || ',"nums":' || Zljsonstr(r_Fee(I).��������, 1); --N 1 ���Σ������˻����
      v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(r_Fee(I).ʣ������, 1); --N 1 ʣ�������������е�ʣ������
      v_Jtmp := v_Jtmp || ',"fee_item_id":' || Nvl(r_Fee(I).�շ�ϸĿid || '', 'null'); --N 1 �շ�ϸĿid
      v_Jtmp := v_Jtmp || ',"unit":"' || Zljsonstr(r_Fee(I).��λ) || '"';
      v_Jtmp := v_Jtmp || ',"fee_name":"' || Zljsonstr(r_Fee(I).�շ���Ŀ) || '"';
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
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_od_list":[' || Substr(v_Jtmp, 2) || ']}}';
    Else
      c_Jtmp   := c_Jtmp || v_Jtmp;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_od_list":[' || c_Jtmp || ']}}';
    End If;
  Elsif n_��ѯ��ʽ = 5 Then
    If n_��Դ = 1 Then
      Select Max(a.�շ�ϸĿid)
      Into n_Tmp
      From ������ü�¼ A
      Where a.No = v_Fee_No And a.��¼���� = n_��¼���� And a.ҽ����� = n_ҽ��id And a.��¼״̬ In (0, 1, 3);
    Else
      Select Max(a.�շ�ϸĿid)
      Into n_Tmp
      From סԺ���ü�¼ A
      Where a.No = v_Fee_No And a.��¼���� = n_��¼���� And a.ҽ����� = n_ҽ��id And a.��¼״̬ In (0, 1, 3);
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_item_id":"' || Nvl(n_Tmp, 0) || '"}}';
  Elsif n_��ѯ��ʽ = 6 Then
  
    v_Vals := v_Fee_No;
  
    l_Vals := t_Strlist();
    While v_Vals Is Not Null Loop
      If Length(v_Vals) <= 4000 Then
        l_Vals.Extend;
        l_Vals(l_Vals.Count) := v_Vals;
        v_Vals := Null;
      Else
        l_Vals.Extend;
        l_Vals(l_Vals.Count) := Substr(v_Vals, 1, Instr(v_Vals, ',', 3980) - 1);
        v_Vals := Substr(v_Vals, Instr(v_Vals, ',', 3980) + 1);
      End If;
    End Loop;
    n_Fee_Ampaib := 0;
    For I In 1 .. l_Vals.Count Loop
      If n_��Դ = 1 Then
        Select Sum(a.ʵ�ս��) As ���
        Into n_Tmp
        From ������ü�¼ A,
             (Select /*+cardinality(f,10)*/
                f.C1 As NO, To_Number(f.C2) As ��¼����
               From Table(f_Str2list2(l_Vals(I), ',', ':')) F) N
        Where a.ҽ����� Is Not Null And a.No = n.No And a.��¼���� = n.��¼���� And a.���ʷ��� = 1 And a.��¼״̬ = 0;
      Else
        Select Sum(a.ʵ�ս��) As ���
        Into n_Tmp
        From סԺ���ü�¼ A,
             (Select /*+cardinality(f,10)*/
                f.C1 As NO, To_Number(f.C2) As ��¼����
               From Table(f_Str2list2(l_Vals(I), ',', ':')) F) N
        Where a.ҽ����� Is Not Null And a.No = n.No And a.��¼���� = n.��¼���� And a.���ʷ��� = 1 And a.��¼״̬ = 0;
      End If;
      n_Fee_Ampaib := n_Fee_Ampaib + Nvl(n_Tmp, 0);
    End Loop;
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ampaib":' || Zljsonstr(n_Fee_Ampaib, 1) || '}}';
  Elsif n_��ѯ��ʽ = 7 Then
    v_Fee_Ids := j_Json.Get_String('fee_ids');
    If n_��Դ = 1 Then
      Select f_List2str(Cast(Collect(a.Id || '') As t_Strlist), ',') ����ids
      Into v_Fee_Ids_Out
      From ������ü�¼ A
      Where a.��¼���� In (2, 1, 11) And a.��¼״̬ = 1 And Nvl(a.����״̬, 0) = 0 And
            a.Id In (Select /*+cardinality(x,10)*/
                      x.Column_Value
                     From Table(Cast(f_Num2list(v_Fee_Ids) As Zltools.t_Numlist)) X);
    Else
      Select f_List2str(Cast(Collect(a.Id || '') As t_Strlist), ',') ����ids
      Into v_Fee_Ids_Out
      From סԺ���ü�¼ A
      Where a.��¼���� = 2 And a.��¼״̬ = 1 And Nvl(a.����״̬, 0) = 0 And
            a.Id In (Select /*+cardinality(x,10)*/
                      x.Column_Value
                     From Table(Cast(f_Num2list(v_Fee_Ids) As Zltools.t_Numlist)) X);
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_ids":"' || v_Fee_Ids_Out || '"}}';
  Elsif n_��ѯ��ʽ = 8 Then
    v_Jtmp := Null;
    --Ҫ���ִ�еķ����У�������ҩƷ������
    If n_��Դ = 1 Then
      For R In (Select ID, ִ�в���id
                From ������ü�¼
                Where �շ���� Not In ('4', '5', '6', '7') And
                      ҽ����� + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As ҽ��id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)) Loop
      
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || '}';
      
      End Loop;
    Else
      For R In (Select ID, ִ�в���id
                From סԺ���ü�¼
                Where �շ���� Not In ('4', '5', '6', '7') And
                      ҽ����� + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As ҽ��id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)) Loop
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || '}';
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_dept_list":[' || Substr(v_Jtmp, 2) || ']}}';
  
  Elsif n_��ѯ��ʽ = 9 Then
    v_Jtmp := Null;
    --�����е�ҩƷҽ��������Ϣ
    If n_��Դ = 1 Then
      For R In (Select ID, ִ�в���id, �շ�ϸĿid, (���� * ����) ʣ������, ҽ�����, NO, ���
                From ������ü�¼
                Where �շ���� In ('5', '6') And ��¼״̬ In (0, 1, 3) And
                      ҽ����� + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As ҽ��id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)
                Order By ��� Desc) Loop
      
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || ',"drug_id":' || Nvl(r.�շ�ϸĿid || '', 'null');
        v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(r.ʣ������, 1);
        v_Jtmp := v_Jtmp || ',"order_id":' || r.ҽ�����;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_origin":1';
        v_Jtmp := v_Jtmp || ',"serial_num":' || r.���;
        v_Jtmp := v_Jtmp || '}';
      
      End Loop;
    Else
      For R In (Select ID, ִ�в���id, �շ�ϸĿid, (���� * ����) ʣ������, ҽ�����, NO, ���
                From סԺ���ü�¼
                Where �շ���� In ('5', '6') And ��¼״̬ In (0, 1, 3) And
                      ҽ����� + 0 In (Select /*+cardinality(b,10) */
                                    Column_Value As ҽ��id
                                   From Table(Cast(f_Str2list(v_Order_Ids, ',') As t_Strlist)) B) And
                      NO In (Select /*+cardinality(b,10) */
                              Column_Value As NO
                             From Table(Cast(f_Str2list(v_Fee_No, ',') As t_Strlist)) B)
                Order By ��� Desc) Loop
        v_Jtmp := v_Jtmp || ',{"fee_id":' || r.Id;
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || ',"drug_id":' || Nvl(r.�շ�ϸĿid || '', 'null');
        v_Jtmp := v_Jtmp || ',"quantity":' || Zljsonstr(r.ʣ������, 1);
        v_Jtmp := v_Jtmp || ',"order_id":' || r.ҽ�����;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_origin":2';
        v_Jtmp := v_Jtmp || ',"serial_num":' || r.���;
        v_Jtmp := v_Jtmp || '}';
      End Loop;
    End If;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_pivas_list":[' || Substr(v_Jtmp, 2) || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getorderfeeinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Getchargeofflist
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡδ������ʵĵ�
  --��Σ�Json_In:��ʽ
  --  input
  --     pati_id                N 1 ����id
  --     pati_pageid            N 1 ��ҳid 
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ 
  --    charge_off_list[]�б�
  --          fee_no            C 1 ���õ��ݺ�
  --          item_name         C 1 �շ���Ŀ����
  --          dept_name         C 1 ��˲������� 
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id Number;
  n_��ҳid Number;
  v_Output Varchar2(32767);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_��ҳid := Nvl(j_Json.Get_Number('pati_pageid'), 0);

  For R In (Select Distinct a.No, d.���� As ��Ŀ, c.���� As ����
            From סԺ���ü�¼ A, ���˷������� B, ���ű� C, �շ���ĿĿ¼ D
            Where a.Id = b.����id And a.�շ�ϸĿid = d.Id And b.��˲���id = c.Id(+) And b.���ʱ�� Is Null And a.����id = n_����id And
                  Nvl(a.��ҳid, 0) = n_��ҳid) Loop
  
    zlJsonPutValue(v_Output, 'fee_no', r.No, 0, 1);
    zlJsonPutValue(v_Output, 'item_name', r.��Ŀ);
    zlJsonPutValue(v_Output, 'dept_name', r.����, 0, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","charge_off_list":[' || v_Output || ']}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getchargeofflist;
/


Create Or Replace Procedure Zl_Exsesvr_Getbillsexestate
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ����ִ�������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     fee_nos               C 1 ���ݺŶ���ƴ����ԭ������˵��ҽ���������ɵķ���no�Ų������ظ�
  --     fee_origin            N 1 ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fee_exe_list[]�б�
  --          order_id          N 1 ҽ��id
  --          fee_no            C 1 ���ݺ�
  --          fee_item_id       N 1 �շ�ϸĿid
  --          exe_dept_id       N 1 ִ�в���id
  --          exe_state         N 1 ִ��״̬
  ---------------------------------------------------------------------------
  j_Input Pljson;
  j_Json  Pljson;

  l_Nos  t_Strlist;
  v_Nos  Varchar2(32767);
  n_��Դ Number; --1-���2-����
  v_Jtmp Varchar2(32767);
  c_Jtmp Clob;

Begin
  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��Դ := j_Json.Get_Number('fee_origin');
  v_Nos  := j_Json.Get_String('fee_nos');
  l_Nos  := t_Strlist();

  While v_Nos Is Not Null Loop
    If Length(v_Nos) <= 4000 Then
      l_Nos.Extend;
      l_Nos(l_Nos.Count) := v_Nos;
      v_Nos := Null;
    Else
      l_Nos.Extend;
      l_Nos(l_Nos.Count) := Substr(v_Nos, 1, Instr(v_Nos, ',', 3980) - 1);
      v_Nos := Substr(v_Nos, Instr(v_Nos, ',', 3980) + 1);
    End If;
  End Loop;

  For I In 1 .. l_Nos.Count Loop
    If 1 = n_��Դ Then
      --����
      For R In (Select a.ҽ����� As ҽ��id, a.No, a.�շ�ϸĿid, a.ִ�в���id, Max(a.ִ��״̬) As ִ��״̬
                From ������ü�¼ A
                Where a.No In (Select /*+cardinality(f,10)*/
                                f.Column_Value As NO
                               From Table(f_Str2list(l_Nos(I))) F) And a.ҽ����� Is Not Null And a.��¼״̬ In (0, 1)
                Group By a.ҽ�����, a.No, a.�շ�ϸĿid, a.ִ�в���id) Loop
      
        v_Jtmp := v_Jtmp || ',{"order_id":' || r.ҽ��id;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_item_id":' || Nvl(r.�շ�ϸĿid || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_state":' || Nvl(r.ִ��״̬ || '', 'null');
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
    Else
      --סԺ
      For R In (Select a.ҽ����� As ҽ��id, a.No, a.�շ�ϸĿid, a.ִ�в���id, Max(a.ִ��״̬) As ִ��״̬
                From סԺ���ü�¼ A
                Where a.No In (Select /*+cardinality(f,10)*/
                                f.Column_Value As NO
                               From Table(f_Str2list(l_Nos(I))) F) And a.ҽ����� Is Not Null And a.��¼״̬ In (0, 1)
                Group By a.ҽ�����, a.No, a.�շ�ϸĿid, a.ִ�в���id
                Order By a.ҽ�����, a.No) Loop
      
        v_Jtmp := v_Jtmp || ',{"order_id":' || r.ҽ��id;
        v_Jtmp := v_Jtmp || ',"fee_no":"' || r.No || '"';
        v_Jtmp := v_Jtmp || ',"fee_item_id":' || Nvl(r.�շ�ϸĿid || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_dept_id":' || Nvl(r.ִ�в���id || '', 'null');
        v_Jtmp := v_Jtmp || ',"exe_state":' || Nvl(r.ִ��״̬ || '', 'null');
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
    End If;
  End Loop;

  If c_Jtmp Is Null Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_exe_list":[' || Substr(v_Jtmp, 2) || ']}}';
  Else
    c_Jtmp   := c_Jtmp || v_Jtmp;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_exe_list":[' || c_Jtmp || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zljsonstr(SQLCode || ':' || SQLErrM) || '"}}';
End Zl_Exsesvr_Getbillsexestate;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Getreturndruginfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ��ҩ�����������Ϣ��ҩ���շ���ѯ��ҩ��Ϣʱ���õ�
  --��Σ�Json_In:��ʽ
  --  input
  --     begin_time            C 1 ��ʼʱ�䣬
  --     end_time              C 1 ����ʱ��
  --     request_dept_id       N 1 ���벿��ID
  --     audit_dept_id         N 1 ��˲���id

  --     type_query            N 0 ��ѯ��ʽ��2-ҩ���շ���ѯʱ����ҩ��ϸ������������Ϣ����ҩ��������ϸ����3-ҩ���շ���ѯʱ����ҩ���ܡ����ػ��ܽ������4-ҩ���շ���ѯ���˻���ʱ�á���ҩ��������ϸ��
  --     effective_time        N 0 ��Ч 0-������1-������2-����������
  --     otherdept_id          N 0 �Է�����id,��ҩ����id
  --     pati_ids              C 0 ����ID����ƴ������������ֲ��ˣ�����
  --     rcp_no                C 0 �����������ݺ�

  --����: Json_Out,��ʽ����
  --  output
  --    code                    N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                 C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    drug_list[]�б�
  --          drug_id           N 1 ҩƷid�����շ�ϸĿid
  --          request_num       N 1 ������
  --          audit_num         N 1 �����

  --    grp_list[]��ҩ���� 
  --          rcp_info                C 1 ҩƷ��Ϣ
  --          in_unit                 C 1 סԺ��λ 
  --          drug_code               C 1 ҩƷ����  
  --          quantity                N 1 Ӧ���� 
  --          back_number             N 1 ��ҩ�� 
  --          reality_number          N 1 ʵ���� 
  --          money                   N 1 ���
  --   detail_list[]��ҩ������ϸ
  --          rcpdtl_id               N 1 ����id��������ϸid
  --          quantity                N 1 ����,��ҩ����
  --          serial_num              N 1 ���
  --          order_id                N 1 ҽ��id 
  --          charge_time             C 1 ����ʱ�� 
  --          rcp_no                  C 1 No ������
  --          charge_people           C 1 ������ 
  --          pati_id                 N 1 ����id
  --          pati_pageid             N 1 ��ҳid

  --   quan_list[]�������
  --         drug_id                  N 1 ҩƷid
  --         quantity                 N 1 ����,��ҩ����
  --         re_money                 N 1 ��� ��ҩ���
  ---------------------------------------------------------------------------
  j_Input      Pljson;
  j_Json       Pljson;
  d_��ʼʱ��   Date;
  d_����ʱ��   Date;
  n_���벿��id ���˷�������.���벿��id%Type;
  n_��˲���id ���˷�������.��˲���id%Type;
  v_Output     Varchar2(32767);
  c_Output     Clob;
  n_Showtype   Number(3);
  n_Type       Number(1);
  v_Jtmp       Varchar2(32767); --��Ҫ���ʹ�ô˱���
  c_Jtmp       Clob; --��Ҫ���ʹ�ô˱���
  n_Ч��       Number(3);
  v_����ids    Varchar2(4000);
  n_�Է�����id Number(18);
  n_����id     Number(18);
  v_��ҩ��Ϣ   Varchar2(32767);
  v_No         Varchar2(30);

  Cursor c_Group_Type Is
    Select a.ժҪ ҩƷ����, a.ժҪ ҩƷ��Ϣ, a.ժҪ סԺ��λ, a.���� ����, a.���� ��ҩ��, a.���� ʵ����, a.���� ���
    From סԺ���ü�¼ A
    Where 0 = 1;
  r_Grp c_Group_Type%RowType;

  Procedure Get����ƴ������ As
  Begin
    v_Jtmp := v_Jtmp || ',';
  
    Zljsonputvalue(v_Jtmp, 'rcp_info', r_Grp.ҩƷ��Ϣ, 0, 1);
    Zljsonputvalue(v_Jtmp, 'drug_code', r_Grp.ҩƷ����, 0);
    Zljsonputvalue(v_Jtmp, 'in_unit', r_Grp.סԺ��λ, 0);
    Zljsonputvalue(v_Jtmp, 'money', r_Grp.���, 1);
    Zljsonputvalue(v_Jtmp, 'quantity', r_Grp.����, 1);
    Zljsonputvalue(v_Jtmp, 'back_number', r_Grp.��ҩ��, 1);
    Zljsonputvalue(v_Jtmp, 'reality_number', r_Grp.ʵ����, 1, 2);
  
    If Length(v_Jtmp) > 30000 Then
      If c_Jtmp Is Null Then
        c_Jtmp := Substr(v_Jtmp, 2);
      Else
        c_Jtmp := c_Jtmp || v_Jtmp;
      End If;
      v_Jtmp := Null;
    End If;
  End;

Begin
  --�������
  j_Input      := Pljson(Json_In);
  j_Json       := j_Input.Get_Pljson('input');
  n_Type       := Nvl(j_Json.Get_Number('type_query'), 0);
  d_��ʼʱ��   := To_Date(j_Json.Get_String('begin_time'), 'YYYY-MM-DD hh24:mi:ss');
  d_����ʱ��   := To_Date(j_Json.Get_String('end_time'), 'YYYY-MM-DD hh24:mi:ss');
  n_���벿��id := Nvl(j_Json.Get_Number('request_dept_id'), 0);
  n_��˲���id := Nvl(j_Json.Get_Number('audit_dept_id'), 0);
  n_Ч��       := j_Json.Get_Number('effective_time');
  v_No         := j_Json.Get_String('rcp_no');

  If n_Type = 0 Then
    For R In (Select a.�շ�ϸĿid As ҩƷid, Sum(a.���� / Nvl(b.סԺ��װ, 1)) As ������,
                     Sum(Decode(a.״̬, 1, a.���� / Nvl(b.סԺ��װ, 1), 0)) As �����
              From ���˷������� A, ҩƷ��� B
              Where a.�շ�ϸĿid = b.ҩƷid And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And a.���벿��id = n_���벿��id And
                    a.��˲���id = n_��˲���id
              Group By a.�շ�ϸĿid) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      Zljsonputvalue(v_Output, 'drug_id', r.ҩƷid, 1, 1);
      Zljsonputvalue(v_Output, 'request_num', r.������, 1);
      Zljsonputvalue(v_Output, 'audit_num', r.�����, 1, 2);
    
    End Loop;
  
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      v_Output := '';
    End If;
    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","drug_list":[') || c_Output || To_Clob(']}}');
    Else
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","drug_list":[' || v_Output || ']}}';
    End If;
  Else
    If d_��ʼʱ�� Is Null Then
      Select Sysdate - 1, Sysdate Into d_��ʼʱ��, d_����ʱ�� From Dual;
    End If;
    If n_Type = 3 Then
      Select zl_GetSysParameter('ҩƷ������ʾ', Null, 100) Into n_Showtype From Dual;
      If Nvl(n_Showtype, 0) = 0 Then
        n_Showtype := 1;
      End If;
      For R In (Select a.ҩƷ����,
                       Nvl(b.����, a.����) || Decode(a.����, Null, Null, '(' || a.���� || ')') ||
                        Decode(a.���, Null, Null, ' ' || a.���) As ҩƷ��Ϣ, a.סԺ��λ, a.����, 0 ��ҩ��, 0 ʵ����, a.���
                From (Select b.ҩƷid, c.���� As ҩƷ����, c.����, c.����, c.���, b.סԺ��λ, Sum(a.���� / Nvl(b.סԺ��װ, 1)) As ����,
                              Sum(a.���) As ���
                       From (Select a.����id, a.����, a.���� * b.��׼���� ���, b.�շ�ϸĿid ҩƷid
                              From ���˷������� A, סԺ���ü�¼ B
                              Where a.����id = b.Id And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And b.ҽ����� Is Not Null And
                                    (b.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And a.��˲���id = n_��˲���id And Nvl(a.״̬, 0) = 0 And
                                    b.�շ���� In ('5', '6', '7') And (b.��ҩ����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And
                                    a.���벿��id = n_���벿��id And
                                    (b.����id + 0 In (Select /*+cardinality(x,10)*/
                                                     x.Column_Value
                                                    From Table(f_Str2list(v_����ids)) X) Or Nvl(v_����ids, 'NONE') = 'NONE') And
                                    (b.ҽ����Ч = n_Ч�� Or Nvl(n_Ч��, 2) = 2)) A, ҩƷ��� B, �շ���ĿĿ¼ C
                       Where a.ҩƷid = b.ҩƷid And a.ҩƷid = c.Id
                       Group By b.ҩƷid, c.����, c.����, c.����, c.���, b.סԺ��λ
                       Having Sum(a.���� / Nvl(b.סԺ��װ, 1)) <> 0 Or Sum(a.���) <> 0) A, �շ���Ŀ���� B
                Where a.ҩƷid = b.�շ�ϸĿid(+) And b.����(+) = 1 And b.����(+) = n_Showtype
                Order By a.ҩƷ����
                -- ���ܷ�ҩ�ţ�ת���ɷ���IDƴ���ٴ��������Ǳ�ȥ��
                -- ��ҩ��ѯ��ʱ���� ��ҩ;����������ʧЧ���ڲ�����������ֻ�����������
                ) Loop
        r_Grp := R;
        Get����ƴ������;
      End Loop;
      If c_Jtmp Is Null Then
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","grp_list":[' || Substr(v_Jtmp, 2) || ']}}';
      Else
        c_Jtmp   := c_Jtmp || v_Jtmp;
        Json_Out := '{"output":{"code":1,"message":"�ɹ�","grp_list":[' || c_Jtmp || ']}}';
      End If;
    Elsif n_Type = 2 Then
      --��һ��Ӧ�÷ŵ�������ķ�����
      --��ҩ���� ����ʱ�� Ϊ������   
      --������ϸ���� ��Ա��ʱ�䣬����ϸ��ʱ���õ� 2 ʱ
      n_����id := 1;
      For R In (Select a.����id, a.����id, a.��ҳid, a.No, a.ҽ����� ҽ��id, Sum(a.����) ����,
                       To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') ����ʱ��, a.������
                From (Select a.����id, a.����, a.����ʱ��, a.������, b.ҽ�����, b.No, b.����id, b.��ҳid
                       From ���˷������� A, סԺ���ü�¼ B
                       Where a.����id = b.Id And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And b.ҽ����� Is Not Null And
                             a.��˲���id = n_��˲���id /* n_�ⷿid*/
                             And Nvl(a.״̬, 0) = 0 And b.�շ���� In ('5', '6', '7') And
                             (b.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And (b.��ҩ����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And
                             a.���벿��id = n_���벿��id /* n_����id*/
                             And (b.����id + 0 In (Select /*+cardinality(x,10)*/
                                                  x.Column_Value
                                                 From Table(f_Str2list(v_����ids)) X) Or Nvl(v_����ids, 'NONE') = 'NONE') And
                             (b.ҽ����Ч = n_Ч�� Or Nvl(n_Ч��, 2) = 2)) A
                Group By a.����id, a.����id, a.��ҳid, a.No, a.ҽ�����, a.����ʱ��, a.������) Loop
      
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',{"rcpdtl_id":' || r.����id;
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"quantity":' || Zljsonstr(r.����, 1);
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"serial_num":' || n_����id;
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"order_id":' || r.ҽ��id;
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"charge_time":"' || r.����ʱ�� || '"';
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"rcp_no":"' || r.No || '"';
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"charge_people":"' || Zljsonstr(r.������) || '"';
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"pati_pageid":' || r.��ҳid;
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"pati_id":' || r.����id;
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || '}';
        n_����id   := n_����id + 1;
      End Loop;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","detail_list":[' || Substr(v_��ҩ��Ϣ, 2) || ']}}';
    Elsif n_Type = 4 Then
      --���λ�����ϸ���� ���˻��ܵ�ʱ���õ� 4 ʱ
      For R In (Select a.ҩƷid, Sum(a.����) ����, Sum(a.���� * a.��׼����) ���
                From (Select a.����, b.��׼����, b.�շ�ϸĿid ҩƷid
                       From ���˷������� A, סԺ���ü�¼ B
                       Where a.����id = b.Id And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And b.ҽ����� Is Not Null And
                             a.��˲���id = n_��˲���id And Nvl(a.״̬, 0) = 0 And b.�շ���� In ('5', '6', '7') And
                             (b.��ҩ����id = n_�Է�����id Or Nvl(n_�Է�����id, 0) = 0) And a.���벿��id = n_���벿��id And
                             (b.No = v_No Or Nvl(v_No, 'NONE') = 'NONE') And
                             (b.����id + 0 In (Select /*+cardinality(x,10)*/
                                              x.Column_Value
                                             From Table(f_Str2list(v_����ids)) X) Or Nvl(v_����ids, 'NONE') = 'NONE') And
                             (b.ҽ����Ч = n_Ч�� Or Nvl(n_Ч��, 2) = 2)
                       -- ���ܷ�ҩ�ţ�ת���ɷ���IDƴ���ٴ��������Ǳ�ȥ��
                       -- ��ҩ��ѯ��ʱ���� ��ҩ;����������ʧЧ���ڲ�����������ֻ�����������
                       ) A
                Group By a.ҩƷid, a.��׼����) Loop
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',{"drug_id":' || r.ҩƷid;
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"quantity":' || Zljsonstr(r.����, 1);
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || ',"re_money":' || Zljsonstr(r.���, 1);
        v_��ҩ��Ϣ := v_��ҩ��Ϣ || '}';
      End Loop;
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","quan_list":[' || Substr(v_��ҩ��Ϣ, 2) || ']}}';
    End If;
  End If;

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getreturndruginfo;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Adddepositinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����Ԥ��������
  --��Σ�Json_In:��ʽ
  --  input
  --    oper_fun            C    ����״̬:0-������Ԥ���� ;1-����Ϊδ��Ч��Ԥ����
  --    deposit_info        C  1  Ԥ�����б�
  --      pati_id           N  1  ����id
  --      pati_pageid       N  1  ��ҳid
  --      pati_name         C  1  ��������
  --      pati_sex          C  1  �Ա�
  --      pati_age          C  1  ����
  --      outpatient_num    C  1  �����
  --      inpatient_num     C  1  סԺ��
  --      mdlpay_name       C  1  ���ʽ����
  --      deposit_id        N  1  Ԥ��ID
  --      deposit_no        C  1  Ԥ�����ݺ�
  --      invc_no           C  1  ��Ʊ��
  --      deposit_type      N     Ԥ�����:1-����;2-סԺ
  --      dept_id           N  1  �ɿ����id
  --      money             N  1  �ɿ���
  --      emp_name          C  1  �ɿλ
  --      emp_bank_name     C  1  ��λ������
  --      emp_bank_actno    C  1  �������˺�
  --      memo              C  1  ժҪ
  --      recv_id           N  1  Ʊ������id
  --    balance_info        C     ������Ϣ:Ŀǰֻ֧��һ�ֽ��㷽ʽ
  --      blnc_mode         C  1  ���㷽ʽ
  --      blnc_no           C  1  �������
  --      cardtype_id       C  1  �����id
  --      consumer_no       C  1  ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --      consume_card_id   N  1  ���ѿ�ID
  --      cardno            C  1  ֧������
  --      swapno            C  1  ������ˮ��
  --      swapmemo          C  1  ����˵��
  --      cprtion_unit      C  1  ������λ
  --      operator_name     C  1  ����Ա����
  --      operator_code     C  1  ����Ա���
  --      create_time       C  1  �Ǽ�ʱ����տ�ʱ��:yyyy-mm-dd hh:mi:ss
  --      insurance_type    N  1  ����
  --      insurance_num     C  1  ҽ����
  --      insurance_pwd     C  1  ҽ������
  --      start_einv        N  1  �Ƿ����õ���Ʊ��:1-����;0-������
  --����: Json_Out,��ʽ����
  --  output
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    deposit_id  N 1 Ԥ��ID
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;
  o_Json  PLJson;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  --���ν�����Ϣ
  n_����״̬ Number(2);

  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;

  d_�Ǽ�ʱ�� ������ü�¼.�Ǽ�ʱ��%Type;
  n_���ѿ�id Number(18);

  --֧����ʽ����
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_�������   ����Ԥ����¼.�������%Type;
  n_�����id   ����Ԥ����¼.�����id%Type;
  n_���㿨��� ����Ԥ����¼.���㿨���%Type;
  v_֧������   ����Ԥ����¼.����%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;
  v_������λ   ����Ԥ����¼.������λ%Type;
  n_��������id ����Ԥ����¼.��������id%Type;
  --Ԥ����ر�������
  n_����id     ����Ԥ����¼.����id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  v_Ԥ������   ����Ԥ����¼.No%Type;
  n_Ԥ�����   ����Ԥ����¼.Ԥ�����%Type;
  n_��ҳid     ����Ԥ����¼.��ҳid%Type;
  n_�ɿ����id ����Ԥ����¼.����id%Type;
  n_�ɿ���   ����Ԥ����¼.���%Type;
  v_�ɿλ   ����Ԥ����¼.�ɿλ%Type;
  v_��λ������ ����Ԥ����¼.��λ������%Type;
  v_�������˺� ����Ԥ����¼.��λ�ʺ�%Type;
  v_ժҪ       ����Ԥ����¼.ժҪ%Type;
  n_����id     Ʊ��ʹ����ϸ.����id%Type;
  v_��Ʊ��     ����Ԥ����¼.ʵ��Ʊ��%Type;

  v_��������     ����Ԥ����¼.����%Type;
  v_�Ա�         ����Ԥ����¼.�Ա�%Type;
  v_����         ����Ԥ����¼.����%Type;
  n_�����       ����Ԥ����¼.�����%Type;
  n_סԺ��       ����Ԥ����¼.סԺ��%Type;
  v_���ʽ���� ����Ԥ����¼.���ʽ����%Type;
  n_Ԥ������Ʊ�� ����Ԥ����¼.Ԥ������Ʊ��%Type;
  n_����         Number(18);
  v_ҽ����       Varchar2(100);
  v_����         Varchar2(100);
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����״̬   := Nvl(j_Json.Get_Number('oper_fun'), 0);
  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');
  d_�Ǽ�ʱ��   := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');

  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  --1���������Ϣ
  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Null Then
    v_Err_Msg := '�����ڽ�����Ϣ������!';
    Raise Err_Item;
  End If;

  v_���㷽ʽ     := o_Json.Get_String('blnc_mode');
  v_�������     := o_Json.Get_String('blnc_no');
  n_�����id     := o_Json.Get_Number('cardtype_id');
  n_���㿨���   := o_Json.Get_Number('consumer_no');
  v_֧������     := o_Json.Get_String('cardno');
  v_������ˮ��   := o_Json.Get_String('swapno');
  v_����˵��     := o_Json.Get_String('swapmemo');
  v_������λ     := o_Json.Get_String('cprtion_unit');
  n_���ѿ�id     := o_Json.Get_Number('consume_card_id');
  n_����         := o_Json.Get_Number('insurance_type');
  v_ҽ����       := o_Json.Get_String('insurance_num');
  v_����         := o_Json.Get_String('insurance_pwd');
  n_Ԥ������Ʊ�� := o_Json.Get_Number('start_einv');

  If Nvl(n_�����id, 0) = 0 Then
    n_�����id := Null;
  End If;
  If Nvl(n_���㿨���, 0) = 0 Then
    n_���㿨��� := Null;
  End If;

  --2.��ȡԤ����Ϣ
  o_Json := PLJson();
  o_Json := j_Json.Get_Pljson('deposit_info');
  If o_Json Is Null Then
    v_Err_Msg := '������Ԥ��������Ϣ,����!';
    Raise Err_Item;
  End If;

  n_����id     := o_Json.Get_Number('pati_id');
  n_��ҳid     := o_Json.Get_Number('pati_pageid');
  n_Ԥ��id     := o_Json.Get_Number('deposit_id');
  v_Ԥ������   := o_Json.Get_String('deposit_no');
  v_��Ʊ��     := o_Json.Get_String('invc_no');
  n_Ԥ�����   := Nvl(o_Json.Get_Number('deposit_type'), 2);
  n_�ɿ����id := o_Json.Get_Number('dept_id');
  n_�ɿ���   := o_Json.Get_Number('money');
  v_�ɿλ   := o_Json.Get_String('emp_name');
  v_��λ������ := o_Json.Get_String('emp_bank_name');
  v_�������˺� := o_Json.Get_String('emp_bank_actno');
  v_ժҪ       := o_Json.Get_String('memo');
  n_����id     := o_Json.Get_Number('recv_id');

  v_��������     := o_Json.Get_String('pati_name');
  v_�Ա�         := o_Json.Get_String('pati_sex');
  v_����         := o_Json.Get_String('pati_age');
  n_�����       := To_Number(o_Json.Get_String('outpatient_num'));
  n_סԺ��       := To_Number(o_Json.Get_String('inpatient_num'));
  v_���ʽ���� := o_Json.Get_String('mdlpay_name');

  If Nvl(n_Ԥ��id, 0) = 0 Then
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
  End If;

  If Nvl(n_��������id, 0) = 0 Then
    n_��������id := n_Ԥ��id;
  End If;

  --����״̬_In:0-�������㣬1-����Ϊ�쳣���ݻ�δ��Ч�ĵ��ݣ�2-����쳣����
  If Nvl(n_����, 0) <> 0 Then
    v_�ɿλ   := n_����;
    v_�������˺� := v_ҽ����;
    v_��λ������ := v_����;
  End If;

  If n_Ԥ������Ʊ�� Is Null Then
    n_Ԥ������Ʊ�� := Zl_Fun_Isstarteinvoice(2, Nvl(n_����, 0), 1, n_Ԥ�����);
  End If;

  Zl_����Ԥ����¼_Insert_s(n_Ԥ��id, v_Ԥ������, v_��Ʊ��, n_����id, n_��ҳid, v_��������, v_�Ա�, v_����, n_�����, n_סԺ��, v_���ʽ����, n_�ɿ����id,
                     n_�ɿ���, v_���㷽ʽ, v_�������, v_�ɿλ, v_��λ������, v_�������˺�, v_ժҪ, v_����Ա���, v_����Ա����, n_����id, n_Ԥ�����, n_�����id,
                     n_���㿨���, v_֧������, v_������ˮ��, v_����˵��, v_������λ, d_�Ǽ�ʱ��, Null, Null, 1, Nvl(n_����״̬, 0), n_��������id, Null,
                     Nvl(n_Ԥ������Ʊ��, 0), n_����);

  If Nvl(n_���㿨���, 0) <> 0 And Nvl(n_���ѿ�id, 0) <> 0 Then
    -- ���ѿ�����
    Zl_���˿������¼_֧��(n_���ѿ�id, v_֧������, 0, n_�ɿ���, n_Ԥ��id, v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��);
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, ' [ ZLSOFT ] ' || v_Err_Msg || ' [ ZLSOFT ] ');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Adddepositinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Getorderchargedinfo
(
  Json_In  In Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��жϵ�ǰ��ִ��ҽ����Ӧ�ķ��õ��Ƿ����շѻ���ʻ��۵��Ƿ�����˺͵��ݵ�״̬
  --��Σ�Json_In:��ʽ
  --  input
  --     fee_origin         N 1 ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
  --     order_ids          C 1 ҽ��IDs�����ŷָ�
  --     fee_nos            C 1 ���õ���ƴ�������ŷָ�
  --     oper_type          N 1 �жϷ�ʽ ��0-����Ƿ����δ�շѼ�¼��1-����Ƿ�������շѼ�¼

  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    isexist             N 1 ����ֵ����٣�0-�٣�1-��
  --    blance_sign         N 1 �Ƿ����쳣���ã�0-������1-�����쳣����
  ---------------------------------------------------------------------------
  j_Input     PLJson;
  j_Json      PLJson;
  n_��Դ      Number;
  v_Order_Ids Varchar2(32767);
  v_Fee_Nos   Varchar2(32767);
  n_���쳣��  Number;
  Int��ʽ     Number;
  n_Cnt       Number;
  n_Blnout    Number;

  v_Output Varchar2(32767);
  Cursor c_Out Is
    Select Nvl(a.��¼״̬, 0) As ��¼״̬, a.ҽ����� As ҽ��id, Nvl(a.ִ��״̬, 0) As ִ��״̬, Nvl(a.����id, 0) As ����id, a.No,
           Nvl(a.����״̬, 0) As ����״̬, Nvl(a.��¼����, 0) As ��¼����
    From ������ü�¼ A
    Where a.��¼״̬ In (0, 1, 3) And
          a.No In (Select /*+cardinality(x,10)*/
                    x.Column_Value
                   From Table(Cast(f_Str2List(v_Fee_Nos) As Zltools.t_Strlist)) X) And
          a.ҽ����� + 0 In (Select /*+cardinality(x,10)*/
                          x.Column_Value
                         From Table(Cast(f_Num2List(v_Order_Ids) As Zltools.t_Numlist)) X);

  Type t_Fee Is Table Of c_Out%RowType;
  r_Fee t_Fee;

  Cursor c_In Is
    Select Nvl(a.��¼״̬, 0) As ��¼״̬, a.ҽ����� As ҽ��id, Nvl(a.ִ��״̬, 0) As ִ��״̬, Nvl(a.����id, 0) As ����id, a.No,
           Nvl(a.����״̬, 0) As ����״̬, Nvl(a.��¼����, 0) As ��¼����
    From סԺ���ü�¼ A
    Where a.��¼״̬ In (0, 1, 3) And
          a.No In (Select /*+cardinality(x,10)*/
                    x.Column_Value
                   From Table(Cast(f_Str2List(v_Fee_Nos) As Zltools.t_Strlist)) X) And
          a.ҽ����� + 0 In (Select /*+cardinality(x,10)*/
                          x.Column_Value
                         From Table(Cast(f_Num2List(v_Order_Ids) As Zltools.t_Numlist)) X);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��Դ      := j_Json.Get_Number('fee_origin');
  v_Order_Ids := j_Json.Get_String('order_ids');
  v_Fee_Nos   := j_Json.Get_String('fee_nos');
  Int��ʽ     := j_Json.Get_Number('oper_type');

  If n_��Դ = 1 Then
    Open c_Out;
    Fetch c_Out Bulk Collect
      Into r_Fee;
    Close c_Out;
  Else
    Open c_In;
    Fetch c_In Bulk Collect
      Into r_Fee;
    Close c_In;
  End If;

  n_Blnout   := 1;
  n_���쳣�� := 0;
  n_Cnt      := r_Fee.Count;

  If Nvl(n_Cnt, 0) = 0 And Int��ʽ = 1 Then
    n_Blnout := 0;
  Else
    For I In 1 .. n_Cnt Loop
    
      --��֧ int��ʽ=0
      If Int��ʽ = 0 Then
        If r_Fee(I).��¼���� = 1 And r_Fee(I).��¼״̬ = 1 And r_Fee(I).����״̬ = 1 And r_Fee(I).����id <> 0 Then
          n_Blnout   := 0;
          n_���쳣�� := 1;
          Exit;
        End If;
        If r_Fee(I).��¼״̬ = 0 Or r_Fee(I).��¼���� = 1 And r_Fee(I).��¼״̬ = 1 And r_Fee(I).����id = 0 Then
          n_Blnout := 0;
          Exit;
        End If;
      Else
        --��֧ int��ʽ=1
        If r_Fee(I).��¼״̬ <> 1 And r_Fee(I).����״̬ <> 1 Then
          n_Blnout := 0;
          Exit;
        End If;
      End If;
    
    End Loop;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'isexist', n_Blnout, 1);
  zlJsonPutValue(v_Output, 'blance_sign', n_���쳣��, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getorderchargedinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Billhavebalance
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ��ж�һ�ż��ʵ�/���Ƿ��Ѿ�����
  --��Σ�Json_In:��ʽ
  --  input
  --       fee_origin         N 1 ������Դ(Ĭ��=2��1-������ã�2-סԺ����)
  --       fee_no             C 1 ���õ��ݺţ�һ��NO    
  --       order_id           N 1 ҽ��id
  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    state         N 1 ���������0-δ���ʣ�1-��ȫ�����ʣ�2-�Ѳ��ֽ���
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Fee_No Varchar2(100);
  n_��Դ   Number;
  n_ҽ��id Number;
  n_������ Number;
  n_������ Number;
  n_����   Number;
Begin
  --������� 
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Fee_No := j_Json.Get_String('fee_no');
  n_��Դ   := j_Json.Get_Number('fee_origin');
  n_ҽ��id := j_Json.Get_Number('order_id');
  n_����   := 0;
  n_������ := 0;
  n_������ := 0;
  If n_��Դ = 1 Then
    For R In (Select Nvl(�۸񸸺�, ���) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
              From ������ü�¼
              Where NO = v_Fee_No And ��¼���� In (2, 12) And (ҽ����� + 0 = n_ҽ��id Or Nvl(n_ҽ��id, 0) = 0)
              Group By Nvl(�۸񸸺�, ���)) Loop
      n_������ := n_������ + 1;
      If Nvl(r.���ʽ��, 0) <> 0 Then
        n_������ := n_������ + 1;
      End If;
    End Loop;
  Else
    For R In (Select Nvl(�۸񸸺�, ���) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
              From סԺ���ü�¼
              Where NO = v_Fee_No And ��¼���� In (2, 12) And (ҽ����� + 0 = n_ҽ��id Or Nvl(n_ҽ��id, 0) = 0)
              Group By Nvl(�۸񸸺�, ���)) Loop
      n_������ := n_������ + 1;
      If Nvl(r.���ʽ��, 0) <> 0 Then
        n_������ := n_������ + 1;
      End If;
    End Loop;
  End If;
  --�޽�����,�൱��δ����
  If n_������ = 0 Then
    n_���� := 0;
  Else
    If n_������ = n_������ Then
      n_���� := 1;
    Else
      n_���� := 2;
    End If;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","state":' || n_���� || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Billhavebalance;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Getdepositinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:��ȡԤ����Ϣ
  --��Σ�Json_In:��ʽ
  --   input
  --    deposit_no  C 1 ���ݺ�:Ԥ�����ݺ�
  --    rec_state N 1 ��¼״̬:1-ԭʼ��ֵ��Ԥ����¼(������������¼);2-�˿��Ԥ����¼,3-ԭʼ��ֵ��¼(�������ʵļ��쳣��)
  --����  json
  --output
  --  code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  deposit_info  C 1 Ԥ������Ϣ
  --    pati_id N 1 ����ID
  --    pati_pageid N 1 ��ҳid
  --    deposit_id  N 1 Ԥ��ID
  --    deposit_no  C 1 Ԥ�����ݺ�
  --    invc_no C 1 ��Ʊ��
  --    deposit_type  N   Ԥ�����:1-����;2-סԺ
  --    dept_id N 1 �ɿ����id
  --    money N 1 �ɿ���
  --    emp_name  C 1 �ɿλ
  --    emp_bank_name C 1 ��λ������
  --    emp_bank_actno  C 1 �������˺�
  --    memo  C 1 ժҪ
  --    operator_name C 1 ����Ա����
  --    operator_code C 1 ����Ա���
  --    create_time C 1 �տ�ʱ��:yyyy-mm-dd hh:mi:ss
  --  balance_info  C   ������Ϣ:Ŀǰֻ֧��һ�ֽ��㷽ʽ
  --    blnc_mode C 1 ���㷽ʽ
  --    blnc_no C 1 �������
  --    cardtype_id N 1 �����id
  --    consumer_no N 1 ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --    consume_card_id N 1 ���ѿ�ID
  --    cardno  C 1 ֧������
  --    swapno  C 1 ������ˮ��
  --    swapmemo  C 1 ����˵��
  --    cprtion_unit  C 1 ������λ
  --    blnc_state  N 1 ����״̬(��У�Ա�־):0��NULL�����ɿ��¼;1-δ���ýӿ�;2-�ӿڵ������
  --    insurance_type  N   ����
  --    insurance_num C   ҽ����
  --    insurance_pwd C   ҽ������
  --    relation_id C 1 ��������ID

  ---------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  j_Input   PLJson;
  j_Json    PLJson;

  v_���ݺ�   ����Ԥ����¼.No%Type;
  n_��¼״̬ Number(4);
  n_Find     Number(2);

  v_Output  Varchar2(32767);
  v_Balance Varchar2(32767);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ�   := j_Json.Get_String('deposit_no');
  n_��¼״̬ := Nvl(j_Json.Get_Number('rec_state'), 1);

  If v_���ݺ� Is Null Then
    v_Err_Msg := 'δ����Ԥ�����ݺţ�����!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  n_Find := 0;
  For c_Ԥ�� In (Select a.Id, a.No, a.ʵ��Ʊ��, a.��¼״̬, a.����id, a.��ҳid, a.����id, a.�ɿλ, a.��λ������, a.��λ�ʺ�, a.ժҪ, a.���, a.���㷽ʽ,
                      a.�������, To_Char(a.�տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �տ�ʱ��, a.����Ա���, a.����Ա����, a.Ԥ�����, a.�����id,
                      a.���㿨���, a.����, a.������ˮ��, a.����˵��, a.������λ, a.�������, a.У�Ա�־, a.��ת��, a.��������, a.�Ự��, a.��������id, a.����ʱ��,
                      a.������Ա, c.���ѿ�id, m.����
               From ����Ԥ����¼ A, ���˿������¼ C, ���㷽ʽ M
               Where a.���㷽ʽ = m.����(+) And a.��¼���� = 1 And a.Id = c.����id(+) And
                     (a.��¼״̬ = n_��¼״̬ Or n_��¼״̬ = 3 And a.��¼״̬ In (0, 1, 3)) And a.no= v_���ݺ�) Loop
    --Ԥ��������Ϣ
    zlJsonPutValue(v_Output, 'pati_id', Nvl(c_Ԥ��.����id, 0), 1, 1);
    zlJsonPutValue(v_Output, 'pati_pageid', Nvl(c_Ԥ��.��ҳid, 0), 1);
    zlJsonPutValue(v_Output, 'deposit_id', Nvl(c_Ԥ��.Id, 0), 1);
    zlJsonPutValue(v_Output, 'deposit_no', Nvl(c_Ԥ��.No, ''));
    zlJsonPutValue(v_Output, 'invc_no', Nvl(c_Ԥ��.ʵ��Ʊ��, ''));
    zlJsonPutValue(v_Output, 'deposit_type', Nvl(c_Ԥ��.Ԥ�����, ''));
    zlJsonPutValue(v_Output, 'dept_id', Nvl(c_Ԥ��.����id, 0), 1);
    zlJsonPutValue(v_Output, 'money', Nvl(c_Ԥ��.���, 0), 1);
    zlJsonPutValue(v_Output, 'emp_name', Nvl(c_Ԥ��.�ɿλ, ''));
    zlJsonPutValue(v_Output, 'emp_bank_name', Nvl(c_Ԥ��.��λ������, ''));
    zlJsonPutValue(v_Output, 'emp_bank_actno', Nvl(c_Ԥ��.��λ�ʺ�, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(c_Ԥ��.ժҪ, ''));
    zlJsonPutValue(v_Output, 'operator_code', Nvl(c_Ԥ��.����Ա���, ''));
    zlJsonPutValue(v_Output, 'operator_name', Nvl(c_Ԥ��.����Ա����, ''));
    zlJsonPutValue(v_Output, 'create_time', Nvl(c_Ԥ��.�տ�ʱ��, ''), 0, 2);
  
    v_Output := '"deposit_info": ' || v_Output;
  
    --������Ϣ
    zlJsonPutValue(v_Balance, 'blnc_mode', Nvl(c_Ԥ��.���㷽ʽ, ''), 0, 1);
    zlJsonPutValue(v_Balance, 'blnc_no', Nvl(c_Ԥ��.�������, ''));
  
    zlJsonPutValue(v_Balance, 'cardtype_id', Nvl(c_Ԥ��.�����id, 0), 1);
    zlJsonPutValue(v_Balance, 'consumer_no', Nvl(c_Ԥ��.���㿨���, 0), 1);
    zlJsonPutValue(v_Balance, 'consume_card_id', Nvl(c_Ԥ��.���ѿ�id, 0), 1);
  
    zlJsonPutValue(v_Balance, 'cardno', Nvl(c_Ԥ��.����, ''));
    zlJsonPutValue(v_Balance, 'swapno', Nvl(c_Ԥ��.������ˮ��, ''));
    zlJsonPutValue(v_Balance, 'swapmemo', Nvl(c_Ԥ��.����˵��, ''));
  
    zlJsonPutValue(v_Balance, 'cprtion_unit', Nvl(c_Ԥ��.������λ, ''));
    zlJsonPutValue(v_Balance, 'relation_id', Nvl(c_Ԥ��.��������id, 0), 1);
  
    If Nvl(c_Ԥ��.����, 0) = 3 Then
      zlJsonPutValue(v_Balance, 'insurance_type', To_Number(Nvl(c_Ԥ��.�ɿλ, '0')));
      zlJsonPutValue(v_Balance, 'insurance_num', Nvl(c_Ԥ��.��λ�ʺ�, ''));
      zlJsonPutValue(v_Balance, 'insurance_pwd', Nvl(c_Ԥ��.��λ������, ''));
	Else
      zlJsonPutValue(v_Balance, 'insurance_type', '0');
      zlJsonPutValue(v_Balance, 'insurance_num', '');
      zlJsonPutValue(v_Balance, 'insurance_pwd',  '');
    End If;
    zlJsonPutValue(v_Balance, 'blnc_state', Nvl(c_Ԥ��.У�Ա�־, 0), 1, 2);
    v_Balance := '"balance_info":' || v_Balance || '';
    v_Output  := v_Output || ',' || v_Balance;
  
    n_Find := 1;
    Exit;
  End Loop;
  If n_Find = 0 Then
    v_Err_Msg := 'δ�ҵ�Ԥ������Ϊ' || v_���ݺ� || '��Ԥ�����ݣ�����!';
    Json_Out  := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�",' || v_Output || '}}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdepositinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Getoutproomlist
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  -------------------------------------------------------------------------------------
  --���ܣ�����������ȡ�йҺŰ��ŵ����������б�
  --��Σ�Json_In:��ʽ
  --input
  --  query_type        N    1  ��ѯ��ʽ 1-����ҽ��վ���������б�,2-����ת�����ת����Ҽ��ؽ�������,3-����ת������ٴ������¼���ؽ�������
  --  site_no           C    0 վ���
  --  outproom_name     C    0 ��������
  --  emg_sign          C    0 �����־
  --  dept_id           N    0 ����id
  --  outp_dr_name      C    0 ����ҽ������
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    outproom_list           ��������б�
  --       outproom_code    C   1 ���ұ���
  --       outproom_name    C   1 ��������
  --       outproom_becode  C   1 �������Ƽ���
  -------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_վ���   Varchar2(50);
  v_�������� Varchar2(50);
  n_�����־ Number(5);

  n_��ѯ��ʽ Number(5);

  n_����id   Number(18);
  v_ҽ������ Varchar2(200);

  v_����ƥ�� Varchar2(50);
  v_Output   Varchar2(32767);
  c_Output   Clob;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ��ʽ := Nvl(j_Json.Get_Number('query_type'), 0);
  v_վ���   := j_Json.Get_String('site_no');
  v_�������� := j_Json.Get_String('outproom_name');
  n_�����־ := Nvl(j_Json.Get_Number('outproom_name'), 0);
  v_ҽ������ := j_Json.Get_String('outp_dr_name');
  n_����id   := Nvl(j_Json.Get_Number('dept_id'), 0);

  If n_��ѯ��ʽ = 1 Then
  
    If Nvl(To_Number(zl_GetSysParameter('����ƥ��')), 0) = 0 Then
      v_����ƥ�� := '%';
    Else
      v_����ƥ�� := '';
    End If;
  
    For c_������� In (Select Distinct e.����, e.����, e.����
                   From �������� E, �ҺŰ������� D, �ҺŰ��� C, ������Ա A, �ϻ���Ա�� B, �ٴ����� F
                   Where a.��Աid = b.��Աid And b.�û��� = User And c.����id = a.����id And c.Id = d.�ű�id And e.���� = d.�������� And
                         a.����id = f.����id And ((n_�����־ = 1 And f.�������� = '20') Or n_�����־ = 0) And
                         ((Upper(e.����) Like v_�������� || '%' Or Upper(e.����) Like v_����ƥ�� || v_�������� || '%' Or
                         Upper(e.����) Like v_����ƥ�� || v_�������� || '%') Or v_�������� Is Null) And (e.վ�� = v_վ��� Or e.վ�� Is Null)) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'outproom_code', c_�������.����, 0, 1);
      zlJsonPutValue(v_Output, 'outproom_name', c_�������.����);
      zlJsonPutValue(v_Output, 'outproom_becode', c_�������.����, 0, 2);
    
    End Loop;
  Elsif n_��ѯ��ʽ = 2 Then
    For c_������� In (Select Distinct e.����, e.����, e.����
                   From �ҺŰ������� A, �������� E
                   Where a.�ű�id In (Select ID
                                    From �ҺŰ���
                                    Where ����id = n_����id And (ҽ������ = v_ҽ������ Or ҽ������ Is Null Or v_ҽ������ Is Null)) And
                         a.�������� = e.����
                   Order By ����) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'outproom_code', c_�������.����, 0, 1);
      zlJsonPutValue(v_Output, 'outproom_name', c_�������.����);
      zlJsonPutValue(v_Output, 'outproom_becode', c_�������.����, 0, 2);
    
    End Loop;
  Elsif n_��ѯ��ʽ = 3 Then
  
    For c_������� In (Select Distinct b.����, b.����, b.����
                   From �ٴ��������Ҽ�¼ A, �������� B
                   Where a.����id = b.Id And
                         a.��¼id In
                         (Select a.Id
                          From �ٴ������¼ A
                          Where a.����id = n_����id And (a.ҽ������ = v_ҽ������ Or a.ҽ������ Is Null Or v_ҽ������ Is Null))
                   Order By b.����) Loop
    
      If Length(Nvl(v_Output, ' ')) > 30000 Then
        If c_Output Is Not Null Then
          c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
        Else
          c_Output := To_Clob(v_Output);
        End If;
      
        v_Output := Null;
      End If;
    
      zlJsonPutValue(v_Output, 'outproom_code', c_�������.����, 0, 1);
      zlJsonPutValue(v_Output, 'outproom_name', c_�������.����);
      zlJsonPutValue(v_Output, 'outproom_becode', c_�������.����, 0, 2);
    
    End Loop;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","outproom_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","outproom_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getoutproomlist;
/

Create Or Replace Procedure Zl_Exsesvr_Getpatiantifee
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡ���˵Ŀ���ҩ����û���
  --��Σ�Json_In:��ʽ
  --input
  --  pati_id        N    1  ����ID
  --  pati_pageid    N    1  ��ҳID
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fee_info            ������Ϣ
  --       anti_fee            N   1 ����ҩ��
  --       drug_fee            N   1 ��ҩ��
  --       inp_fee             N   1 סԺ����

  j_Input  PLJson;
  j_Json   PLJson;
  v_Output Varchar2(32767);

  n_����id Number(18);
  n_��ҳid Number(18);

Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_��ҳid := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  v_Output := Null;
  For c_�����б� In (Select Sum(Decode(Nvl(e.������, 0), 0, 0, a.���ʽ��)) As ����ҩ��,
                        Sum(Decode(a.�շ����, '5', a.���ʽ��, '6', a.���ʽ��, '7', a.���ʽ��, 0)) As ��ҩ��, Sum(a.���ʽ��) As סԺ����
                 From סԺ���ü�¼ A, ҩƷ��� D, ҩƷ���� E
                 Where a.����id = n_����id And a.��ҳid = n_��ҳid And a.��¼״̬ <> 0 And a.�շ�ϸĿid = d.ҩƷid(+) And
                       d.ҩ��id = e.ҩ��id(+)) Loop
  
    zlJsonPutValue(v_Output, 'anti_fee', c_�����б�.����ҩ��, 1, 1);
    zlJsonPutValue(v_Output, 'drug_fee', c_�����б�.��ҩ��, 1);
    zlJsonPutValue(v_Output, 'inp_fee', c_�����б�.סԺ����, 1, 2);
  
  End Loop;

  If v_Output Is Null Then
    zlJsonPutValue(v_Output, 'anti_fee', 0, 1, 1);
    zlJsonPutValue(v_Output, 'drug_fee', 0, 1);
    zlJsonPutValue(v_Output, 'inp_fee', 0, 1, 2);
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_info":' || v_Output || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatiantifee;
/


Create Or Replace Procedure Zl_Exsesvr_Upddepositblncinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�޸�Ԥ��������Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    oper_state  N 1 ����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
  --    pati_id N 1 ����id
  --    deposit_no  C   Ԥ������
  --    deposit_id  N   Ԥ��ID
  --    operator_name C 1 ����Ա����
  --    operator_code C 1 ����Ա���
  --    create_time C 1 ����ʱ��:yyyy-mm-dd hh:mi:ss
  --    invc_no C 1 ��Ʊ��
  --    recv_id N 1 ����id:����Id
  --    balance_info  C   ������Ϣ
  --      blnc_mode C 1 ���㷽ʽ
  --      blnc_no C 1 �������
  --      cardtype_id N 1 �����id
  --      consumer_no N 1 ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --      cardno  C 1 ����
  --      swapno  C 1 ������ˮ��
  --      swapmemo  C 1 ����˵��
  --      memo  C 1 ժҪ
  --      cprtion_unit  C 1 ������λ
  --      other_list[]  C 1 ����������Ϣ
  --        swap_name C 1 ��������
  --        swap_note C 1 ��������
  --����: Json_Out,��ʽ����
  -- output
  --   code                  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  o_Json     PLJson;
  j_Jsonlist Pljson_List := Pljson_List();

  n_����id       Number(18);
  v_Ԥ������     ����Ԥ����¼.No%Type;
  n_Ԥ��id       ����Ԥ����¼.Id%Type;
  n_����id       Ʊ�����ü�¼.Id%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  v_�������     ����Ԥ����¼.�������%Type;
  n_�����id     ����Ԥ����¼.�����id%Type;
  n_���㿨���   ����Ԥ����¼.���㿨���%Type;
  v_֧������     ����Ԥ����¼.����%Type;
  v_������ˮ��   ����Ԥ����¼.������ˮ��%Type;
  v_����˵��     ����Ԥ����¼.����˵��%Type;
  v_ժҪ         ����Ԥ����¼.ժҪ%Type;
  v_������λ     ����Ԥ����¼.������λ%Type;
  v_��������     �������㽻��.������Ŀ%Type;
  v_��������     �������㽻��.��������%Type;
  v_��Ʊ��       ����Ԥ����¼.ʵ��Ʊ��%Type;
  v_����Ա����   ����Ԥ����¼.����Ա����%Type;
  v_����Ա���   ����Ԥ����¼.����Ա���%Type;
  d_�Ǽ�ʱ��     ����Ԥ����¼.�տ�ʱ��%Type;
  n_����״̬     Number(5);
  n_Ԥ������״̬ Number(5);
  n_У�Ա�־     Number(5);

  Err_Item Exception;
  v_Err_Msg Varchar2(500);
  n_Find    Number(2);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����״̬ := Nvl(j_Json.Get_Number('oper_state'), 0);
  n_����id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  v_Ԥ������ := j_Json.Get_String('deposit_no');
  n_Ԥ��id   := j_Json.Get_Number('deposit_id');

  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');
  d_�Ǽ�ʱ��   := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_��Ʊ��     := j_Json.Get_String('invc_no');
  n_����id     := j_Json.Get_Number('recv_id');

  If d_�Ǽ�ʱ�� Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '����ȷ��������Ϣ�����飡';
    Raise Err_Item;
  End If;
  If v_Ԥ������ Is Null Then
    v_Err_Msg := '����ȷ��Ԥ��������Ϣ�����飡';
    Raise Err_Item;
  End If;

  o_Json := j_Json.Get_Pljson('balance_info');
  If o_Json Is Null Then
    v_Err_Msg := '����ȷ��Ԥ������Ϊ' || v_Ԥ������ || '�Ľ�����Ϣ�����飡';
    Raise Err_Item;
  End If;

  v_���㷽ʽ   := o_Json.Get_String('blnc_mode');
  v_�������   := o_Json.Get_String('blnc_no');
  n_�����id   := o_Json.Get_Number('cardtype_id');
  n_���㿨��� := o_Json.Get_Number('consumer_no');
  v_֧������   := o_Json.Get_String('cardno');
  v_������ˮ�� := o_Json.Get_String('swapno');
  v_����˵��   := o_Json.Get_String('swapmemo');
  v_ժҪ       := o_Json.Get_String('memo');
  v_������λ   := o_Json.Get_String('cprtion_unit');
  n_Find       := 0;
  If Nvl(n_���㿨���, 0) = 0 Then
    n_���㿨��� := Null;
  
  End If;
  If Nvl(n_�����id, 0) = 0 Then
    n_�����id := Null;
  
  End If;
  For c_Ԥ�� In (Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����, �Ա�, ����, �����, סԺ��, ���ʽ����, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���,
                      ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
                      �������, У�Ա�־, ��������, �Ự��, ���ӱ�־, ��������id, ����ʱ��, ������Ա, Ԥ������Ʊ��
               From ����Ԥ����¼
               Where ID = n_Ԥ��id And ����id = Nvl(n_����id, 0)) Loop
    --����״̬_In:0-�������㣬1-����Ϊ�쳣���ݻ�δ��Ч�ĵ��ݣ�2-����쳣����;3-��������
    --У�Ա�־
  
    If n_����״̬ = 0 Then
      --����״̬:0-��ɽ���;1-�ӿڵ���ǰ����;2-�ӿڵ��ú�����
      n_Ԥ������״̬ := 2;
      n_У�Ա�־     := Null;
    Elsif n_����״̬ = 1 Or n_����״̬ = 2 Then
      n_Ԥ������״̬ := 3;
      n_У�Ա�־     := n_����״̬;
    Else
      v_Err_Msg := '�����ϱ�Ĳ������ܣ����飡';
      Raise Err_Item;
    End If;
  
    Zl_����Ԥ����¼_Insert_s(n_Ԥ��id, v_Ԥ������, v_��Ʊ��, n_����id, c_Ԥ��.��ҳid, c_Ԥ��.����, c_Ԥ��.�Ա�, c_Ԥ��.����, c_Ԥ��.�����, c_Ԥ��.סԺ��,
                       c_Ԥ��.���ʽ����, c_Ԥ��.����id, c_Ԥ��.���, v_���㷽ʽ, v_�������, c_Ԥ��.�ɿλ, c_Ԥ��.��λ������, c_Ԥ��.��λ�ʺ�, v_ժҪ, v_����Ա���,
                       v_����Ա����, n_����id, c_Ԥ��.Ԥ�����, n_�����id, n_���㿨���, v_֧������, v_������ˮ��, v_����˵��, v_������λ, d_�Ǽ�ʱ��, Null, Null,
                       1, n_Ԥ������״̬, c_Ԥ��.��������id, n_У�Ա�־, Nvl(c_Ԥ��.Ԥ������Ʊ��, 0));
  
    n_Find := 1;
  End Loop;
  If n_Find = 0 Then
    v_Err_Msg := 'δ�ҵ����ݺ�Ϊ' || v_Ԥ������ || '��Ԥ��������Ϣ�����飡';
    Raise Err_Item;
  End If;
  j_Jsonlist := o_Json.Get_Pljson_List('other_list');
  If Not j_Jsonlist Is Null Then
  
    --��ɾ����������
    Delete �������㽻�� Where ����id = n_Ԥ��id;
    Delete �������㽻�� Where ����id = n_Ԥ��id;
  
    For J In 1 .. j_Jsonlist.Count Loop
      o_Json     := PLJson();
      o_Json     := PLJson(j_Jsonlist.Get(J));
      v_�������� := o_Json.Get_String('swap_name');
      v_�������� := o_Json.Get_String('swap_note');
      Insert Into �������㽻�� (����id, ������Ŀ, ��������) Values (n_Ԥ��id, v_��������, v_��������);
    End Loop;
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Upddepositblncinfo;
/
 
Create Or Replace Procedure Zl_Exsesvr_Deldepositerrorrec
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:ɾ���쳣��Ԥ����Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --     deposit_no  C  1  Ԥ������
  --     oper_state  N  1  ����״̬��0-ɾ���쳣��ֵ���ݣ�1-ɾ���쳣�˿�ݣ�2-ɾ���쳣����˿�� 
  --����: Json_Out,��ʽ����
  -- output
  --   code                  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message               C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Ԥ������ ����Ԥ����¼.No%Type;
  n_����״̬ Number(5);
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_Ԥ������ := j_Json.Get_String('deposit_no');
  n_����״̬ := Nvl(j_Json.Get_Number('oper_state'), 0);

  If v_Ԥ������ Is Null Then
    v_Err_Msg := '����ȷ��Ԥ��������Ϣ,���ܽ������ϲ�����';
    Raise Err_Item;
  End If;

  --����_In:0-ɾ���쳣��ֵ���ݣ�1-ɾ���쳣�˿�ݣ�2-ɾ���쳣����˿�� 

  Zl_����Ԥ���쳣��¼_Delete(v_Ԥ������, n_����״̬);

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Deldepositerrorrec;
/

Create Or Replace Procedure Zl_Exsesvr_Checkexeitemvalied
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ������ƺ���㷽ʽ�����ִ����Ŀ�ĺϷ���
  --input   
  --  pati_id      N   1   ����id
  --  register_id   N   1   �Һ�id
  --  receipt_type  C   1   �շ����
  --output
  --  code          C   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message       C   1   Ӧ����Ϣ��
  --  check_flag   N   0   ����־��0-������Ϸ���1-���� ��2-�ܾ�
  --  check_msg    C   0   ���ѻ�ܾ���������ʾ
  -------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id     Number;
  n_�Һ�id     Number;
  v_�շ���� Varchar2(100);
  n_����ģʽ   Number(2);
  v_Check      Varchar2(1000);
  n_Count      Number;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_�Һ�id     := Nvl(j_Json.Get_Number('register_id'), 0);
  v_�շ���� := j_Json.Get_String('receipt_type');

  If n_����id = 0 Then
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","check_flag":0,"check_msg":null}}';
    Return;
  Else
    If n_�Һ�id > 0 Then
      Select Nvl(Max(����ģʽ), 0) As ����ģʽ
      Into n_����ģʽ
      From ���˹Һż�¼
      Where ����id = n_����id And ID = n_�Һ�id;
    Else
      Select Nvl(Max(����ģʽ), 0) As ����ģʽ
      Into n_����ģʽ
      From ���˹Һż�¼
      Where ����id = n_����id And ID In (Select Max(ID) From ���˹Һż�¼ Where ����id = n_����id);
    End If;
  
    --δ���������ƺ����ģʽ
    If n_����ģʽ = 0 Then
      Json_Out := '{"output":{"code":1,"message": "�ɹ�","check_flag":0,"check_msg":null}}';
      Return;
    End If;
  End If;

  --��ҩʱ�������Ƚ���
  If Instr(',' || v_�շ���� || ',', ',5,') <> 0 Or Instr(',' || v_�շ���� || ',', ',6,') <> 0 Or
     Instr(',' || v_�շ���� || ',', ',7,') <> 0 Then
    Select Count(1)
    Into n_Count
    From ����δ�����
    Where ����id = n_����id And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
    If Nvl(n_Count, 0) <> 0 Then
      --����δ�������ݣ������Ƚ���������ִ��
      v_Check := '2|����ҩǰ�������Ƚ���������ҩ';
    End If;
  End If;

  If v_Check Is Null Then
    --���ͨ��ʱ
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","check_flag":0,"check_msg":null}}';
  Else
    --���δͨ������Ҫ���ѻ��ֹʱ
    Json_Out := '{"output":{"code":1,"message": "�ɹ�","check_flag":' || Substr(v_Check, 1, 1) || ',"check_msg":"' ||
                Substr(v_Check, 3) || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkexeitemvalied;
/


Create Or Replace Procedure Zl_Exsesvr_Updbillstartinvoice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�޸ĵ��ݵ���ʼ��Ʊ��:Ԥ�������ʣ��Һŵ�
  --��Σ�Json_In:��ʽ
  --  input     
  --    bill_type N 1 ��������:1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
  --    bill_nos C 1 ���ݺ�:���õ��Ż���㵥�Ż�Ԥ������
  --    inv_no  C   ��Ʊ��:����ʱ����ʾ���
  --����: Json_Out,��ʽ����
  --  output
  --    code                C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1  Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ 
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_�������� Number(5);
  v_���ݺ�   Varchar2(100);
  v_��Ʊ��   Varchar2(100);

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�������� := j_Json.Get_Number('bill_type');
  v_���ݺ�   := j_Json.Get_String('bill_nos');
  v_��Ʊ��   := j_Json.Get_String('inv_no');

  If v_���ݺ� Is Null Then
    v_Err_Msg := 'δ����ָ���ĵ�����Ϣ!';
    Raise Err_Item;
  End If;
  If Instr(v_���ݺ�, ',') > 0 Then
    For c_���� In (Select Column_Value As NO From Table(f_Str2List(v_���ݺ�))) Loop
      Zl_Ʊ����ʼ��_Update(c_����.No, v_��Ʊ��, n_��������);
    End Loop;
  Else
  
    Zl_Ʊ����ʼ��_Update(v_���ݺ�, v_��Ʊ��, n_��������);
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updbillstartinvoice;
/

Create Or Replace Procedure Zl_Exsesvr_Reprintdepositinvc
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���´�ӡԤ����Ʊ
  --��Σ�Json_In:��ʽ
  --input     
  -- deposit_nos C 1 ���ݺ�:Ԥ�����ݺ�
  -- invc_no C 1 ��Ʊ��
  -- invc_id N 1 ����ID
  -- user_name C 1 ʹ��������

  --����: Json_Out,��ʽ����
  --  output
  --    code                C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1  Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ 
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id     Ʊ�����ü�¼.Id%Type;
  v_���ݺ�     Ʊ�ݴ�ӡ����.No%Type;
  v_��Ʊ��     Ʊ��ʹ����ϸ.����%Type;
  v_ʹ�������� Ʊ��ʹ����ϸ.ʹ����%Type;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ� := j_Json.Get_String('deposit_nos');
  v_��Ʊ�� := j_Json.Get_String('invc_no');
  n_����id := j_Json.Get_Number('invc_id');

  v_ʹ�������� := j_Json.Get_String('user_name');

  If v_���ݺ� Is Null Then
    v_Err_Msg := 'δ����ָ����Ԥ��������Ϣ!';
    Raise Err_Item;
  End If;
  Zl_����Ԥ����¼_Reprint(v_���ݺ�, v_��Ʊ��, n_����id, v_ʹ��������);

  Json_Out := zlJsonOut('�ɹ�', 1);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Reprintdepositinvc;
/


Create Or Replace Procedure Zl_Exsesvr_Checkerrordata
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݷ���NO�����ID����շѼ����쳣��Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --      fee_type              C   1 �������'4'-���ģ�'5,6,7'-ҩƷ
  --      rcpdtl_ids            C   1 ������ϸids,����ö��ŷָ�
  --      bill_list[]                  ���飬����NO��Ϣ
  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
  --         rcp_nos                C   1 ����Nos,����ö��ŷָ�
  --����: Json_Out,��ʽ����
  --  output
  --     code                   N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --     message                C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     billid_list[]                   ������ID����ʱ����id�б�
  --        rcpdtl_id           N   1 ������ϸid
  --        fee_status          N   1 ����״̬�� 0-����,1-����
  --        cancel_status       N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
  --        update_status       N   1 �Ƿ�ͬ��״̬:0-����״̬,1-δ����ҩƷ/���ļ���״̬
  --     billno_list[]                 ��NO����ʱ����NO�б�
  --         billtype               N   1 ��������:1-�շѴ���;2-���ʴ���
  --         rcp_no                 C   1 ����no
  --         fee_status             N   1 ����״̬������շ�ʱ,0-δ�շ�,1-���շ�,2-�쳣�շ�;��Լ���ʱ,0-����,1-����
  --         cancel_status          N   1 ����״̬:0-����״̬,1-����ͬ����־�쳣
  --         update_drug_status     N   1 �Ƿ�ͬ��״̬:0-����״̬,2-δ����ҩƷ/�����շ�״̬
  --     expense_list[]               ��ҩƷ����
  --         billtype               N   1 (ԭʼ)��������:1-�շѴ���;2-���ʴ���
  --         rcp_no                 C   1 (ԭʼ)����no
  --         rcpdtl_id              N   1 (ԭʼ)������ϸid
  --         rcp_no_new             C   1 �����ɵĴ���NO
  --         rcpdtl_id_new          N   1 �����ɴ�����ϸid
  --         pati_pageid        N  1  ��ҳID
  ------------------------------------------------------------------------------------------------------------
  j_Input     PLJson;
  j_Json      PLJson;
  v_Output    Varchar2(32767);
  j_Bill_List Pljson_List := Pljson_List();
  v_Nos       Varchar2(4000);
  n_��������  Number(1); -- 1- �շѴ���;2- ���ʵ�����,3 - ���ʱ��� n_Count Number(3);
  v_�շ����  Varchar2(100);

  c_����ids Clob; --����id
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  l_����id Collection_Type;
  c_Output Clob;

  n_����״̬ Number(1);
  Err_Custom Exception;
  v_Err Varchar2(255);

  v_Billno_List  Varchar2(32767);
  v_Expense_List Varchar2(32767);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�շ����  := j_Json.Get_String('fee_type');
  c_����ids   := j_Json.Get_String('rcpdtl_ids');
  j_Bill_List := j_Json.Get_Pljson_List('bill_list');

  If c_����ids Is Null And j_Bill_List Is Null Then
    v_Err := 'δ���벡�˴�����Ϣ�򴦷���ϸID�����飡';
    Raise Err_Custom;
  End If;

  --1.������ID����
  --�� c_����ids ����װ�ɲ�����4000 �ļ��ϴ�����ֹʹ�� f_Num2list ��������
  If c_����ids Is Not Null Then
    While c_����ids Is Not Null Loop
      If Length(c_����ids) <= 4000 Then
        l_����id(l_����id.Count) := c_����ids;
        c_����ids := Null;
      Else
        l_����id(l_����id.Count) := Substr(c_����ids, 1, Instr(c_����ids, ',', 3980) - 1);
        c_����ids := Substr(c_����ids, Instr(c_����ids, ',', 3980) + 1);
      End If;
    End Loop;
  
    --���ݲ���ID�����쳣����
    v_Output := Null;
    For I In 0 .. l_����id.Count - 1 Loop
      For r_���� In (Select /*+cardinality(j,10)*/
                    a.Id, C1.ͬ����־ As ����ͬ����־, c.ͬ����־ As �Ƿ�ͬ����־, Decode(a.��¼״̬, 0, 0, 1) As ����״̬
                   From סԺ���ü�¼ A, Table(f_Num2List(l_����id(I))) J, ���˷����쳣��¼ C, ���˷����쳣��¼ C1
                   Where a.Id = j.Column_Value And Instr(',' || v_�շ���� || ',', ',' || a.�շ���� || ',') > 0 And
                         a.Id = c.����id(+) And c.��������(+) = 0 And a.Id = C1.����id(+) And C1.��������(+) = 1
                   Union All
                   Select /*+cardinality(j,10)*/
                    a.Id, C1.ͬ����־ As ����ͬ����־, c.ͬ����־ As �Ƿ�ͬ����־, Decode(a.��¼״̬, 0, 0, 1) As ����״̬
                   From ������ü�¼ A, Table(f_Num2List(l_����id(I))) J, ���˷����쳣��¼ C, ���˷����쳣��¼ C1
                   Where a.Id = j.Column_Value And Instr(',' || v_�շ���� || ',', ',' || a.�շ���� || ',') > 0 And
                         a.Id = c.����id(+) And c.��������(+) = 0 And a.Id = C1.����id(+) And C1.��������(+) = 1) Loop
      
        If Length(Nvl(v_Output, ' ')) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        zlJsonPutValue(v_Output, 'rcpdtl_id', r_����.Id, 1, 1);
        zlJsonPutValue(v_Output, 'fee_status', r_����.����״̬, 1);
        zlJsonPutValue(v_Output, 'update_status', r_����.�Ƿ�ͬ����־, 1);
        zlJsonPutValue(v_Output, 'cancel_status', r_����.����ͬ����־, 1, 2);
      End Loop;
    End Loop;
  
    If Not c_Output Is Null And Not v_Output Is Null Then
      c_Output := c_Output || ',' || To_Clob(v_Output);
      v_Output := Null;
    End If;
  
    If Not c_Output Is Null Then
      Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","billid_list":[') || c_Output || To_Clob(']}}');
    Else
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","billid_list":[' || v_Output || ']}}';
    End If;
    Return;
  End If;

  --2.������NO����
  For I In 1 .. j_Bill_List.Count Loop
    j_Json := PLJson();
  
    j_Json     := PLJson(j_Bill_List.Get(I));
    n_�������� := j_Json.Get_Number('billtype');
    v_Nos      := j_Json.Get_String('rcp_nos');
    If Nvl(n_��������, 0) = 0 Then
      v_Err := 'δ���뵥�����ͣ����飡';
      Raise Err_Custom;
    End If;
  
    If Nvl(v_Nos, '-') = '-' Then
      v_Err := 'δ���봦��NO�����飡';
      Raise Err_Custom;
    End If;
  
    v_Output := Null;
    For c_������Ϣ In (Select a.No, Max(c.ͬ����־) As �Ƿ�ͬ����־, Max(C1.ͬ����־) As ����ͬ����־, Max(Decode(n_��������, 1, a.����״̬, 0)) As ����״̬,
                          Max(Decode(a.��¼״̬, 0, 0, 1)) As ��¼״̬
                   From ������ü�¼ A, ���˷����쳣��¼ C, ���˷����쳣��¼ C1
                   Where a.��¼���� = n_�������� And a.Id = c.����id(+) And c.��������(+) = 0 And a.Id = C1.����id(+) And C1.��������(+) = 1 And
                         a.No In (Select /*+cardinality(B,10) */
                                   Column_Value
                                  From Table(f_Str2List(v_Nos)) B) And
                         Instr(',' || v_�շ���� || ',', ',' || a.�շ���� || ',') > 0
                   Group By a.No
                   Union All
                   Select a.No, Max(c.ͬ����־) As �Ƿ�ͬ����־, Max(C1.ͬ����־) As ����ͬ����־, Max(Decode(n_��������, 1, a.����״̬, 0)) As ����״̬,
                          Max(Decode(a.��¼״̬, 0, 0, 1)) As ��¼״̬
                   From סԺ���ü�¼ A, ���˷����쳣��¼ C, ���˷����쳣��¼ C1
                   Where a.��¼���� = n_�������� And a.Id = c.����id(+) And c.��������(+) = 0 And a.Id = C1.����id(+) And C1.��������(+) = 1 And
                         a.No In (Select /*+cardinality(B,10) */
                                   Column_Value
                                  From Table(f_Str2List(v_Nos)) B) And
                         Instr(',' || v_�շ���� || ',', ',' || a.�շ���� || ',') > 0
                   Group By a.No) Loop
    
      If Nvl(c_������Ϣ.����״̬, 0) = 1 Then
        n_����״̬ := 2;
      Else
        n_����״̬ := c_������Ϣ.��¼״̬;
      End If;
    
      zlJsonPutValue(v_Output, 'billtype', n_��������, 1, 1);
      zlJsonPutValue(v_Output, 'rcp_no', c_������Ϣ.No);
      zlJsonPutValue(v_Output, 'fee_status', n_����״̬, 1);
      zlJsonPutValue(v_Output, 'cancel_status', c_������Ϣ.����ͬ����־, 1);
      zlJsonPutValue(v_Output, 'update_status', c_������Ϣ.�Ƿ�ͬ����־, 1, 2);
    
    End Loop;
    If v_Output Is Not Null Then
      If v_Billno_List Is Null Then
        v_Billno_List := v_Output;
      Else
        v_Billno_List := v_Billno_List || ',' || v_Output;
      End If;
    End If;
  
    --��ȡ�������תסԺ�ķ�����Ϣ
    v_Output := Null;
    For c_ת������Ϣ In (Select b.Id As ԭʼid, b.No As ԭʼno, a.Id As ת��id, a.No As ת��no, a.��ҳid
                    From סԺ���ü�¼ A, ������ü�¼ B, ������˼�¼ C, ���˷����쳣��¼ D
                    Where a.Id = c.ת��id And b.Id = c.����id And b.��¼���� = n_�������� And a.Id = d.����id And d.�������� = 2 And
                          b.No In (Select /*+cardinality(B,10) */
                                    Column_Value
                                   From Table(f_Str2List(v_Nos)) B) And
                          Instr(',' || v_�շ���� || ',', ',' || a.�շ���� || ',') > 0) Loop
    
      zlJsonPutValue(v_Output, 'billtype', n_��������, 1, 1);
      zlJsonPutValue(v_Output, 'rcpdtl_id', c_ת������Ϣ.ԭʼid, 1);
      zlJsonPutValue(v_Output, 'rcp_no', c_ת������Ϣ.ԭʼno);
      zlJsonPutValue(v_Output, 'rcpdtl_id_new', c_ת������Ϣ.ת��id, 1);
      zlJsonPutValue(v_Output, 'rcp_no_new', c_ת������Ϣ.ת��no, 0);
      zlJsonPutValue(v_Output, 'pati_pageid', c_ת������Ϣ.��ҳid, 1, 2);
    End Loop;
    If v_Output Is Not Null Then
      If v_Expense_List Is Null Then
        v_Expense_List := v_Output;
      Else
        v_Expense_List := v_Expense_List || ',' || v_Output;
      End If;
    End If;
  
  End Loop;
  v_Billno_List  := ',"billno_list":[' || v_Billno_List || ']';
  v_Expense_List := ',"expense_list":[' || v_Expense_List || ']';

  Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�"' || To_Clob(v_Billno_List) || To_Clob(v_Expense_List) || '}}');
Exception
  When Err_Custom Then
    Json_Out := zlJsonOut(v_Err);
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkerrordata;
/


Create Or Replace Procedure Zl_Exsesvr_Checkpatichangeundo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:�������˱䶯��¼ǰ���
  --��Σ�Json_In:��ʽ
  --input
  --    pati_list[]       ���� ��λ�Ի�����ʱ��ͬʱ�����������
  --      pati_id           N 1 ����id 
  --      pati_pageid       N 1 ��ҳID
  --      undo_type         C 1 ��������
  --      create_time       C 1 �Ǽ�ʱ��
  --      fee_item_id       N 1 ������ĿID
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------

  j_Input PLJson;
  j_Json  PLJson;

  j_Temp       PLJson;
  j_Json_List  Pljson_List;
  n_����id     סԺ���ü�¼.����id%Type;
  n_��ҳid     סԺ���ü�¼.��ҳid%Type;
  n_������Ŀid Number(18);
  v_Undo_Type  Varchar2(100);

  d_��ʼʱ�� Date;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  j_Json_List := j_Json.Get_Pljson_List('pati_list');

  If j_Json_List Is Null Then
    Json_Out := zlJsonOut('����ֵ����,���顣');
    Return;
  End If;
  For I In 1 .. j_Json_List.Count Loop
    j_Temp := PLJson(j_Json_List.Get(I));
  
    n_����id     := j_Temp.Get_Number('pati_id');
    n_��ҳid     := j_Temp.Get_Number('pati_pageid');
    n_������Ŀid := j_Temp.Get_Number('fee_item_id');
    v_Undo_Type  := j_Temp.Get_String('undo_type');
    d_��ʼʱ��   := To_Date(j_Temp.Get_String('create_time'), 'YYYY-MM-DD HH24:MI:SS');
    If Nvl(n_����id, 0) = 0 Then
      Json_Out := zlJsonOut('δ���벡��id������!');
      Return;
    End If;
    If Instr(',��Ժ��ס,��ס,ת����ס,����,��λ�Ի�,ת������ס,', ',' || v_Undo_Type || ',') > 0 Then
      For r_Fee In (Select NO, ����
                    From סԺ���ü�¼
                    Where ����id = n_����id And ��ҳid = n_��ҳid And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= d_��ʼʱ��
                    Group By NO, ���, Mod(��¼����, 10), ����
                    Having Sum(���ʽ��) <> 0) Loop
        If v_Undo_Type = '��λ�Ի�' Then
          Json_Out := zlJsonOut('���� ' || r_Fee.���� || ' ���Զ����ʷ����ѽ���,���ܽ��г���������');
        Else
          Json_Out := zlJsonOut('�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������');
        End If;
        Return;
      End Loop;
    Elsif Instr(',��λ�ȼ��䶯,����ȼ��䶯,', ',' || v_Undo_Type || ',') > 0 Then
      -- v_Undo_Type = '��λ�ȼ��䶯' Or v_Undo_Type = '����ȼ��䶯' 
      For r_Fee In (Select NO
                    From סԺ���ü�¼
                    Where ����id = n_����id And ��ҳid = n_��ҳid And Mod(��¼����, 10) = 3 And �շ�ϸĿid = n_������Ŀid And �Ǽ�ʱ�� >= d_��ʼʱ��
                    Group By NO, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        Json_Out := zlJsonOut('�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������');
        Return;
      End Loop;
    End If;
  End Loop;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkpatichangeundo;
/


Create Or Replace Procedure Zl_Exsesvr_Getconsumercardinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���ѿ���Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --   cardno               C 1 ����
  --   cardtype_num         N 1 �ӿڱ��
  --   check_valid          N   ��Ч�Լ�飺1-��飻0-����� 
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    card_id           N 1 ���ѿ�id
  --    card_pwd          C 1 ����
  --    surplus           N 1 ���
  --    limit_type        N 1 �������
  --    occasion          N 1 Ӧ�ó���
  --    pati_id           N 1 ����ID
  --    specpati          N 1 �Ƿ��ض�����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_����     ���ѿ���Ϣ.����%Type;
  n_�ӿڱ�� ���ѿ���Ϣ.�ӿڱ��%Type;

  n_��ֵ       ���ѿ���Ϣ.�ɷ��ֵ%Type;
  n_Id         ���ѿ���Ϣ.Id%Type;
  n_����id     ���ѿ���Ϣ.����id%Type;
  d_��Ч��     ���ѿ���Ϣ.��Ч��%Type;
  v_����       ���ѿ���Ϣ.����%Type;
  d_����ʱ��   ���ѿ���Ϣ.����ʱ��%Type;
  v_��ǰ״̬   Varchar2(20);
  n_���       ���ѿ���Ϣ.���%Type;
  d_ͣ������   ���ѿ���Ϣ.ͣ������%Type;
  v_�������   ���ѿ���Ϣ.�������%Type;
  v_Ӧ�ó���   ���ѿ����Ŀ¼.Ӧ�ó���%Type;
  n_�ض�����   ���ѿ����Ŀ¼.�Ƿ��ض�����%Type;
  n_ʧЧ���   �ʻ��ɿ����.���%Type;
  n_�������   �ʻ��ɿ����.�������%Type;
  n_Count      Number(5);
  v_Message    Varchar2(2000);
  v_Output     Varchar2(32767);
  n_��Ч�Լ�� Number(1);
Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_����       := j_Json.Get_String('cardno');
  n_�ӿڱ��   := j_Json.Get_Number('cardtype_num');
  n_��Ч�Լ�� := Nvl(j_Json.Get_Number('check_valid'), 0);

  If Nvl(v_����, '-') = '-' Or Nvl(n_�ӿڱ��, 0) = 0 Then
    Json_Out := zlJsonOut('δ�����κο��Ż�ӿڱ�ţ�����!');
    Return;
  End If;

  Select Count(1), Max(a.Id), Max(a.�ɷ��ֵ), Max(a.��Ч��), Max(a.����), Max(a.����ʱ��),
         Max(Decode(a.��ǰ״̬, 2, '����', 3, '�˿�', '����')), Max(a.���), Max(a.ͣ������), Max(a.�������), Max(b.Ӧ�ó���), Max(a.����id),
         Max(b.�Ƿ��ض�����)
  Into n_Count, n_Id, n_��ֵ, d_��Ч��, v_����, d_����ʱ��, v_��ǰ״̬, n_���, d_ͣ������, v_�������, v_Ӧ�ó���, n_����id, n_�ض�����
  From ���ѿ���Ϣ A, ���ѿ����Ŀ¼ B
  Where a.�ӿڱ�� = b.��� And a.���� = v_���� And a.�ӿڱ�� = n_�ӿڱ�� And
        ��� = (Select Max(���) From ���ѿ���Ϣ B Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��)
  Order By a.���;

  If n_Count = 0 Then
    Json_Out := zlJsonOut('�ÿ�������Ч��!');
    Return;
  End If;

  --�Ƿ����
  If n_��Ч�Լ�� = 1 And Nvl(d_����ʱ��, To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')) <
     To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss') Then
    v_Message := '�ÿ��Ѿ���' || Nvl(v_��ǰ״̬, '����') || ',����ˢ������!';
    Json_Out  := zlJsonOut(v_Message);
    Return;
  End If;

  --�Ƿ�ͣ��
  If n_��Ч�Լ�� = 1 And Nvl(d_ͣ������, To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')) <
     To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss') Then
    v_Message := '�ÿ��Ѿ���ֹͣʹ��,����ˢ������!';
    Json_Out  := zlJsonOut(v_Message);
    Return;
  End If;

  --�����Ч��
  If Nvl(d_��Ч��, To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')) <
     To_Date('3000-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss') Then
    --�Ƿ������ֵ
    If n_��Ч�Լ�� = 1 And Nvl(n_��ֵ, 0) <> 1 Then
      v_Message := '�ÿ��Ѿ�ʧЧ,����ˢ������!';
      Json_Out  := zlJsonOut(v_Message);
      Return;
    End If;
    --��ȡʵ�ʿ������(���-ʧЧ���)
    --������ķ�����¼(�������>0)��ֱ��ȡʧЧ���
    Select Count(1), Nvl(Max(b.�������), 0), Nvl(Max(b.���), 0)
    Into n_ʧЧ���, n_�������, n_Count
    From ���˿������¼ A, �ʻ��ɿ���� B
    Where a.������� = b.������� And a.���ѿ�id = b.���ѿ�id And a.��¼���� = 1 And a.���ѿ�id = n_Id;
  
    If n_Count > 0 And n_������� = 0 Then
      --����ǰ�ķ�����¼(�������=0)����Ҫͳ��ʧЧ���
      Select Sum(Nvl(ʧЧ���, 0))
      Into n_ʧЧ���
      From (Select ������ As ʧЧ���
             From ���ѿ���Ϣ A
             Where ID = n_Id And ��Ч�� < Sysdate
             Union All
             Select Nvl(Sum(a.Ӧ�ս��), 0) As ʧЧ���
             From ���˿������¼ A, ���ѿ���Ϣ B
             Where a.���ѿ�id = b.Id And a.��¼���� = 4 And a.���ѿ�id = n_Id And
                   a.����ʱ�� <= Nvl(b.��Ч��, To_Date('3000-01-01', 'yyyy-mm-dd')));
    End If;
    n_��� := n_��� - n_ʧЧ���;
  End If;

  --    card_id           N 1 ���ѿ�id
  --    card_pwd          C 1 ����
  --    surplus           N 1 ���
  --    limit_type        N 1 �������
  --    occasion          N 1 Ӧ�ó���
  --    pati_id           N 1 ����ID
  --    specpati          N 1 �Ƿ��ض�����

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'card_id', n_Id, 1);
  zlJsonPutValue(v_Output, 'surplus', Nvl(n_���, 0), 1);
  zlJsonPutValue(v_Output, 'card_pwd', Nvl(v_����, ''));
  zlJsonPutValue(v_Output, 'limit_type', Nvl(v_�������, ''));
  zlJsonPutValue(v_Output, 'occasion', Nvl(v_Ӧ�ó���, '000'));
  zlJsonPutValue(v_Output, 'pati_id', Nvl(n_����id, 0), 1);
  zlJsonPutValue(v_Output, 'specpati', Nvl(n_�ض�����, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getconsumercardinfo;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Sync_Update
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --���ܣ�����ͬ������ռǷ�ͬ����־����NO�򰴷���ID��
  --��Σ�Json_In:��ʽ
  --  input
  --    sign_type           N 1 ��־���ͣ�0-�Ƿ�ͬ����־������ͬ����־,1-ת��ͬ����־
  --    detail_ids          C  1  ������ϸid��(����id��),֧�ֶ��id���á�,���ָ�
  --    bill_list[]
  --      billtype          N   1 ��������:1-�շѴ���;2-���ʴ���
  --      rcp_no            C   1 ����No
  --����: Json_Out,��ʽ����
  --  output
  --    code                 N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message              C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  c_Detailids Clob;
  Type Collection_Type Is Table Of Varchar2(4000) Index By Binary_Integer;
  l_����ids Collection_Type;

  I           Number;
  j_Bill_List Pljson_List;
  o_Json      PLJson;
  n_����      Number(1);
  v_No        Varchar2(20);
  n_��־����  Number(2);

  j_Input PLJson;
  j_Json  PLJson;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��־���� := j_Json.Get_Number('sign_type');

  --1.�����ݸ���
  j_Bill_List := j_Json.Get_Pljson_List('bill_list');
  If j_Bill_List Is Not Null Then
    For I In 1 .. j_Bill_List.Count Loop
      o_Json := PLJson();
      o_Json := PLJson(j_Bill_List.Get(I));
      n_���� := o_Json.Get_Number('billtype');
      v_No   := o_Json.Get_String('rcp_no');

      If Nvl(n_��־����, 0) = 0 Then
        Delete From ���˷����쳣��¼ a
        Where (a.�������� = 0 Or a.�������� = 1) and a.����ID In (Select ID From סԺ���ü�¼ Where ��¼״̬ In (1, 3) And ��¼���� = n_���� And NO = v_No);
        If Sql%NotFound Then
          Delete From ���˷����쳣��¼ a
          Where (a.�������� = 0 Or a.�������� = 1) and a.����ID In (Select ID From ������ü�¼ Where ��¼״̬ In (1, 3) And ��¼���� = n_���� And NO = v_No);
        End If;
      Elsif n_��־���� = 1 Then
        Delete From ���˷����쳣��¼ a
        Where a.�������� = 2 and a.����ID In (Select ID From סԺ���ü�¼ Where ��¼״̬ In (1, 3) And ��¼���� = n_���� And NO = v_No);
      End If;
    End Loop;
  End If;

  --2.������ID����
  c_Detailids := j_Json.Get_Clob('detail_ids');
  I           := 1;
  While c_Detailids Is Not Null Loop
    If Length(c_Detailids) <= 4000 Then
      l_����ids(I) := c_Detailids;
      c_Detailids := Null;
    Else
      l_����ids(I) := Substr(c_Detailids, 1, Instr(c_Detailids, ',', 3980) - 1);
      c_Detailids := Substr(c_Detailids, Instr(c_Detailids, ',', 3980) + 1);
    End If;
    I := I + 1;
  End Loop;

  If Nvl(n_��־����, 0) = 0 Then
    Forall I In 1 .. l_����ids.Count
      Delete ���˷����쳣��¼
      Where (�������� = 0 Or �������� = 1) and ����ID In (Select /*+Cardinality(j,10)*/
                                    j.Column_Value As ID
                                   From Table(f_Num2List(l_����ids(I))) J);
  Elsif n_��־���� = 1 Then
    Forall I In 1 .. l_����ids.Count
      Delete ���˷����쳣��¼
        Where �������� = 2 and ����ID In (Select /*+Cardinality(j,10)*/
                                      j.Column_Value As ID
                                     From Table(f_Num2List(l_����ids(I))) J);
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Sync_Update;
/


Create Or Replace Procedure Zl_Exsesvr_Getrelatedtransinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݹ�������id,��ȡ������Ϣs
  --��Σ�Json_In:��ʽ
  --input
  --  related_ids  C 1 ��������ID:����ö��ŷ���

  --����: Json_Out,��ʽ����
  --  output
  --    code                      N   1   Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message                   C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     swap_list[]  C 1 ������Ϣ�б�
  --      related_id N 1 ��������ID
  --      cardtype_id N 1 �����ID
  --      blnc_Mode C 1 ���㷽ʽ
  --      swapno  C 1 ������ˮ��
  --      swapmemo  C 1 ����˵��
  --      original_money  N 1 ԭʼ���
  --      return_money  N 1 ���˽��

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_��������ids Varchar2(32680);

  v_Output Varchar2(32767);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_��������ids := j_Json.Get_String('related_ids');
  If v_��������ids Is Null Then
    Json_Out := zlJsonOut('δ�������������Ϣ!');
    Return;
  End If;
  For c_������Ϣ In (With �������� As
                    (Select Column_Value As ��������id From Table(f_Num2List(v_��������ids)))
                   Select /*+cardinality(B,10)*/
                    ��������id, �����id, a.���㷽ʽ, a.������ˮ��, a.����˵��, Sum(ԭʼ���) As ԭʼ���, Sum(���˽��) As ���˽��,
                    Sum(ԭʼ���) - Sum(���˽��) As ʣ��δ�˽��
                   From (Select a.��������id, a.�����id, a.���㷽ʽ, a.������ˮ��, a.����˵��,
                                 Decode(a.��¼����, 1, Decode(a.��¼״̬, 2, 0, 1), 1) * Nvl(���, 0) +
                                  Decode(Mod(��¼����, 10), 1, 0, 1) * Decode(Sign(Nvl(��Ԥ��, 0)), 1, 1, 0) * Nvl(��Ԥ��, 0) As ԭʼ���,
                                 (Decode(Sign(Nvl(���, 0)), -1, 1, 0) * Nvl(���, 0) +
                                  Decode(Sign(Nvl(��Ԥ��, 0)), -1, 1, 0) * Nvl(��Ԥ��, 0)) * Decode(Nvl(a.У�Ա�־, 0), 1, 0) As ���˽��
                          From ����Ԥ����¼ A, �������� B
                          Where a.��������id = b.��������id
                          Union All
                          Select a.��������id, a.�����id, a.���㷽ʽ, a.������ˮ��, a.����˵��, 0 As ԭʼ���, -1 * Nvl(b.���, 0) As ���˽��
                          From ����Ԥ����¼ A, �����˿���Ϣ B, �������� C
                          Where a.Id = b.��¼id And a.��������id = c.��������id And b.�Ƿ�ת�� = 1) A
                   Group By a.��������id, a.�����id, a.���㷽ʽ, a.������ˮ��, a.����˵��) Loop
  
    zlJsonPutValue(v_Output, 'related_id', Nvl(c_������Ϣ.��������id, 0), 1, 1);
    zlJsonPutValue(v_Output, 'cardtype_id', Nvl(c_������Ϣ.�����id, 0), 1);
    zlJsonPutValue(v_Output, 'blnc_mode', Nvl(c_������Ϣ.���㷽ʽ, ''));
    zlJsonPutValue(v_Output, 'swapno', Nvl(c_������Ϣ.������ˮ��, ''));
    zlJsonPutValue(v_Output, 'swapmemo', Nvl(c_������Ϣ.����˵��, ''));
    zlJsonPutValue(v_Output, 'original_money', Nvl(c_������Ϣ.ԭʼ���, 0), 1);
    zlJsonPutValue(v_Output, 'return_money', Nvl(c_������Ϣ.���˽��, 0), 1, 2);
  
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","swap_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrelatedtransinfo;
/



Create Or Replace Procedure Zl_Exsesvr_Getspeccalcfeeitem
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --------------------------------------------------------------------------------------
  --���ܣ���ȡ���۷ѱ���ϸ�б�
  --��Σ�Json_In:��ʽ
  --input
  --����      json
  --output
  --    code                N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    feecategory_list           �ѱ���ϸ�б�
  --       fee_category      C   1 �ѱ�����
  --       fee_item_id       N   1 �շ���ĿID
  --       detail_cacfml     N   1 ���㷽ʽ
  --------------------------------------------------------------------------------------
  v_Output Varchar2(32767);
Begin
  For c_�ѱ���ϸ In (Select Distinct �ѱ�, �շ�ϸĿid, ���㷽��
                 From �ѱ���ϸ
                 Where ���㷽�� = 1 And ������Ŀid Is Null And �շ�ϸĿid Is Not Null) Loop
    zlJsonPutValue(v_Output, 'fee_category', c_�ѱ���ϸ.�ѱ�, 0, 1);
    zlJsonPutValue(v_Output, 'fee_item_id', c_�ѱ���ϸ.�շ�ϸĿid, 1);
    zlJsonPutValue(v_Output, 'detail_cacfml', c_�ѱ���ϸ.���㷽��, 1, 2);
  End Loop;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","feecategory_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getspeccalcfeeitem;
/


Create Or Replace Procedure Zl_Exsesvr_Executeturnwardfee
(
  Json_In  Clob,
  Json_Out Out Varchar2
) Is
  ---------------------------------------------------------------------------
  --����:ִ�в���ת�������õ�ת�룬ת������
  --��Σ�Json_In:��ʽ
  --input
  --  oper_type              N   1   �������ͣ�0-�����䶯,1-���������䶯
  --  change_id_old          N   1   ԭ�����ı䶯��¼��ID
  --  change_id_new          N   1   Ŀ�겡���ı䶯��¼��ID
  --  ward_id_old            N   1   ԭ����ID
  --  ward_id_new            N   1   Ŀ�겡��ID
  --  pat_visit_pnurs        C   1   ���λ�ʿ����
  --  operator_code          C   1   ����Ա���
  --  operator_name          C   1   ����Ա����
  --  pati_info              ������Ϣ���������Щ���˵ķ���
  --    pati_id              N   1   ����ID
  --    pati_pageid          N   1   ��ҳID
  --    pati_name            C   1   ��������
  --    fee_audit_status     N   1   ������˱�־:0���-δ���;1-����˻�ʼ���;2-������
  --    si_inp_status        N   1   סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
  --    catalog_date         C   0   ������Ŀ���ڣ�yyyy-mm-dd hh24:mi:ss
  --  bill_list[]            ת���õ�����Ϣ
  --    fee_no               C   1   ���õ��ݺ�
  --    serial_num           N   1   ���
  --    quantity             N   1   ת������
  --  excute_list[]          ������ִ���б�(���ķ���),��ʹ��ִ����Ϊ0ҲҪ����
  --    fee_id               N   1   ����ID
  --    sended_num           N   1   �ѷ�����
  --����: Json_Out,��ʽ����
  --  output
  --    code                C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  --ת�룬ת������:
  --1.����ִ�еķ�ҩƷ���������ϣ��������Ϊ
  --   1)��ԭ��¼�������ʴ���
  --   2)����һ���²����ķ��ã����˿��ң�����ʱ�䲻��
  --2.����ִ�е�ҩƷ����������
  --   ��������˵Ĵ�����ת����ʱ�Ľ����н���ȷ��(���Դ�ӡ�˲��嵥)����ת���������ʱ��ȷ�ϡ�
  --   a)������ԭ����ͨ�����������������²����ֹ������ģ�
  --   b)����ת����ʱ���Զ������������룬����Ѿ���������ˣ���ѯ����ʾ���Ҳ������ķ��ô����ֹ�ȥ����
  j_Input PLJson;
  j_Json  PLJson;

  j_List Pljson_List;
  j_Temp PLJson;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  n_��������   Number(1);
  n_ԭ�䶯id   ���ñ䶯��¼.ԭ�䶯id%Type;
  n_Ŀ��䶯id ���ñ䶯��¼.Ŀ��䶯id%Type;
  n_ԭ����id   ���ű�.Id%Type;
  n_Ŀ�겡��id ���ű�.Id%Type;
  v_���λ�ʿ   ���˷�������.������%Type;
  v_����Ա��� סԺ���ü�¼.����Ա���%Type;
  v_����Ա���� סԺ���ü�¼.����Ա����%Type;

  n_����id   סԺ���ü�¼.����id%Type;
  n_��ҳid   סԺ���ü�¼.��ҳid%Type;
  v_�������� סԺ���ü�¼.����%Type;
  n_��˱�־ Number(2);
  n_סԺ״̬ Number(2);
  v_��Ŀ���� Varchar2(30);

  n_����id   סԺ���ü�¼.Id%Type;
  n_��ִ���� סԺ���ü�¼.����%Type;
  v_No       סԺ���ü�¼.No%Type;
  n_Max���  סԺ���ü�¼.���%Type;
  n_Dec      Number;
  d_�Ǽ�ʱ�� Date;

  n_ת������ סԺ���ü�¼.����%Type;
  n_ִ������ סԺ���ü�¼.����%Type;

  n_δ��������   ���˷�������.����%Type;
  n_����δ������ ���˷�������.����%Type;
  n_�����ѷ����� ���˷�������.����%Type;
  n_������������ ���˷�������.����%Type;
  n_����ȡ������ ���˷�������.����%Type;

  v_ԭ��������   ���ű�.����%Type;
  v_Ŀ�겡������ ���ű�.����%Type;
  n_Ӧ�ս��     סԺ���ü�¼.Ӧ�ս��%Type;
  n_ʵ�ս��     סԺ���ü�¼.ʵ�ս��%Type;

  Type t_Table Is Record(
    NO   ������ü�¼.No%Type,
    ��� ������ü�¼.���%Type,
    ���� ������ü�¼.����%Type);
  Type t_Fee_Table Is Table Of t_Table;

  l_Fee      t_Fee_Table;
  l_Feeno    t_StrList2;
  l_Executed t_NumList2;

  Procedure ��������_Insert
  (
    ����id_In     ���˷�������.����id%Type,
    �������_In   ���˷�������.�������%Type,
    �շ�ϸĿid_In ���˷�������.�շ�ϸĿid%Type,
    ���벿��id_In ���˷�������.���벿��id%Type,
    ��˲���id_In ���˷�������.��˲���id%Type,
    ����_In       ���˷�������.����%Type,
    ������_In     ���˷�������.������%Type,
    ����ʱ��_In   ���˷�������.����ʱ��%Type,
    ״̬_In       ���˷�������.״̬%Type,
    ����ԭ��_In   ���˷�������.����ԭ��%Type
  ) Is
  Begin
    --ȫ����ִ���ˣ��϶���������Ϊ��ִ�е�
    Insert Into ���˷�������
      (����id, �������, �շ�ϸĿid, ��˲���id, ���벿��id, ����, ������, ����ʱ��, ״̬, ����ԭ��)
    Values
      (����id_In, �������_In, �շ�ϸĿid_In, ��˲���id_In, ���벿��id_In, ����_In, ������_In, ����ʱ��_In, ״̬_In, ����ԭ��_In);
  End;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��������   := j_Json.Get_Number('oper_type');
  n_ԭ�䶯id   := j_Json.Get_Number('change_id_old');
  n_Ŀ��䶯id := j_Json.Get_Number('change_id_new');
  n_ԭ����id   := j_Json.Get_Number('ward_id_old');
  n_Ŀ�겡��id := j_Json.Get_Number('ward_id_new');
  v_���λ�ʿ   := j_Json.Get_String('pat_visit_pnurs');
  v_����Ա��� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');

  n_����id   := j_Json.Get_Number('pati_info.pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_info.pati_pageid');
  v_�������� := j_Json.Get_String('pati_info.pati_name');
  n_��˱�־ := j_Json.Get_Number('pati_info.fee_audit_status');
  n_סԺ״̬ := j_Json.Get_Number('pati_info.si_inp_status');
  v_��Ŀ���� := j_Json.Get_String('pati_info.catalog_date');

  v_Err_Msg := Zl_Pati_Charge_Check(v_��������, n_��˱�־, n_סԺ״̬, v_��Ŀ����);
  If v_Err_Msg Is Not Null Then
    Json_Out := zlJsonOut(v_Err_Msg);
    Return;
  End If;

  --����ת����������
  l_Fee   := t_Fee_Table();
  l_Feeno := t_StrList2();
  j_List  := j_Json.Get_Pljson_List('bill_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Temp := PLJson(j_List.Get(I));
      l_Fee.Extend;
      l_Fee(l_Fee.Count).No := j_Temp.Get_String('fee_no');
      l_Fee(l_Fee.Count).��� := j_Temp.Get_Number('serial_num');
      l_Fee(l_Fee.Count).���� := j_Temp.Get_Number('quantity');
    
      l_Feeno.Extend;
      l_Feeno(l_Feeno.Count) := t_StrObj2(l_Fee(l_Fee.Count).No, l_Fee(l_Fee.Count).���);
    End Loop;
  End If;

  --�����������Ϸ��õ���ִ����
  l_Executed := t_NumList2();
  j_List     := Pljson_List();
  j_List     := j_Json.Get_Pljson_List('excute_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Temp     := PLJson();
      j_Temp     := PLJson(j_List.Get(I));
      n_����id   := j_Temp.Get_Number('fee_id');
      n_��ִ���� := j_Temp.Get_Number('sended_num');
    
      l_Executed.Extend;
      l_Executed(l_Executed.Count) := t_NumObj2(n_����id, n_��ִ����);
    End Loop;
  End If;

  Select Max(Decode(ID, n_ԭ����id, ����, Null)), Max(Decode(ID, n_Ŀ�겡��id, ����, Null))
  Into v_ԭ��������, v_Ŀ�겡������
  From ���ű�
  Where ID In (n_ԭ����id, n_Ŀ�겡��id);

  d_�Ǽ�ʱ�� := Sysdate;
  --���С��λ��
  n_Dec := zl_To_Number(Nvl(zl_GetSysParameter(9), '2'));

  n_Max��� := 0;
  v_No      := '-~';

  For r_���� In (Select a.Id As ����id, a.No, Nvl(a.�۸񸸺�, ���) As ���, a.�շ�ϸĿid, a.ҽ����� As ҽ��id, a.��¼״̬,
                      Nvl(a.����, 1) * a.���� As ����, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��, a.�շ���� As �շ����, Nvl(c.��������, 0) As ����, a.ִ��״̬
               From סԺ���ü�¼ A, Table(l_Feeno) B, �������� C
               Where a.��¼���� = 2 And a.No = b.C1 And Nvl(a.�۸񸸺�, a.���) = b.C2 And a.�շ�ϸĿid = c.����id(+) And
                     a.��¼״̬ In (0, 1, 3)
               Order By NO, ���) Loop
  
    If v_No <> r_����.No Then
      v_No := r_����.No;
      Select Nvl(Max(���), 0)
      Into n_Max���
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = 2 And ��¼״̬ In (0, 1, 3);
    End If;
  
    n_ת������ := 0;
    For I In 1 .. l_Fee.Count Loop
      If l_Fee(I).No = r_����.No And l_Fee(I).��� = r_����.��� Then
        n_ת������ := l_Fee(I).����;
        Exit;
      End If;
    End Loop;
  
    --1.���������ڲ���ִ�еģ�ֱ�ӷ�����������
    If Nvl(r_����.����, 0) = 1 Then
      If r_����.��¼״̬ = 0 Then
        v_Err_Msg := '���� ' || r_����.No || ' ��δ������ˣ���ֹת����������';
        Raise Err_Item;
      End If;
      If v_���λ�ʿ Is Null Then
        v_Err_Msg := 'ԭ���������λ�ʿ�����ڣ����ܽ��������������룡';
        Raise Err_Item;
      End If;
    
      Select Max(C2) Into n_ִ������ From Table(l_Executed) Where C1 = r_����.����id;
      n_δ�������� := Nvl(n_ת������, 0) - Nvl(n_ִ������, 0);
    
      Select Sum(Decode(�������, 0, 1, 0) * ����), Sum(Decode(�������, 0, 0, 1) * ����)
      Into n_����δ������, n_�����ѷ�����
      From ���˷�������
      Where ����id = r_����.����id And Nvl(״̬, 0) = 0;
    
      n_������������ := 0;
      If Nvl(n_δ��������, 0) = Nvl(n_ת������, 0) Then
        --��δִ��
        n_����δ������ := Nvl(n_ת������, 0) - Nvl(n_����δ������, 0);
        If n_����δ������ > 0 Then
          n_������������ := n_������������ + n_����δ������;
          ��������_Insert(r_����.����id, 0, r_����.�շ�ϸĿid, n_ԭ����id, n_ԭ����id, n_����δ������, v_���λ�ʿ, d_�Ǽ�ʱ��, 0,
                      '��' || v_ԭ�������� || 'ת��' || v_Ŀ�겡������);
        End If;
      Elsif Nvl(n_δ��������, 0) = 0 Then
        --ȫ����ִ���ˣ��϶���������Ϊ��ִ�е�
        n_�����ѷ����� := Nvl(n_ת������, 0) - Nvl(n_�����ѷ�����, 0);
        If n_�����ѷ����� > 0 Then
          n_������������ := n_������������ + n_�����ѷ�����;
          ��������_Insert(r_����.����id, 1, r_����.�շ�ϸĿid, n_ԭ����id, n_ԭ����id, n_�����ѷ�����, v_���λ�ʿ, d_�Ǽ�ʱ��, 0,
                      '��' || v_ԭ�������� || 'ת��' || v_Ŀ�겡������);
        End If;
      Else
        --�����в��ֶ�ִ�еĽ������ʣ�һ���ֶ�δִ�е�����
        n_����δ������ := Nvl(n_δ��������, 0) - Nvl(n_����δ������, 0);
        If n_����δ������ > 0 Then
          n_������������ := n_������������ + n_����δ������;
          ��������_Insert(r_����.����id, 0, r_����.�շ�ϸĿid, n_ԭ����id, n_ԭ����id, n_����δ������, v_���λ�ʿ, d_�Ǽ�ʱ��, 0,
                      '��' || v_ԭ�������� || 'ת��' || v_Ŀ�겡������);
        End If;
        --��ִ�в���
        n_�����ѷ����� := Nvl(n_ת������, 0) - Nvl(n_δ��������, 0) - Nvl(n_�����ѷ�����, 0);
        If n_�����ѷ����� > 0 Then
          n_������������ := n_������������ + n_�����ѷ�����;
          ��������_Insert(r_����.����id, 1, r_����.�շ�ϸĿid, n_ԭ����id, n_ԭ����id, n_�����ѷ�����, v_���λ�ʿ, d_�Ǽ�ʱ��, 0,
                      '��' || v_ԭ�������� || 'ת��' || v_Ŀ�겡������);
        End If;
      End If;
    
      --���ӱ䶯��¼
      If Nvl(n_������������, 0) > 0 Then
        --���=ʣ����*(׼����/ʣ����)
        n_Ӧ�ս�� := Round(r_����.Ӧ�ս�� * (n_������������ / r_����.����), n_Dec);
        n_ʵ�ս�� := Round(r_����.ʵ�ս�� * (n_������������ / r_����.����), n_Dec);
      
        Insert Into ���ñ䶯��¼
          (ID, ��¼״̬, ����id, ��ҳid, �䶯ʱ��, ԭ�䶯id, Ŀ��䶯id, ԭ����id, Ŀ�겡��id, ����id, NO, �շ����, �շ�ϸĿid, ҽ�����, ����, ����, Ӧ�ս��, ʵ�ս��,
           ״̬, ժҪ, ����Ա���, ����Ա����)
        Values
          (���ñ䶯��¼_Id.Nextval, Decode(Nvl(n_��������, 0), 0, 1, 2), n_����id, n_��ҳid, d_�Ǽ�ʱ��, n_ԭ�䶯id, n_Ŀ��䶯id, n_ԭ����id,
           n_Ŀ�겡��id, r_����.����id, r_����.No, r_����.�շ����, r_����.�շ�ϸĿid, r_����.ҽ��id, n_������������, r_����.��׼����, n_Ӧ�ս��, n_ʵ�ս��, 2,
           Decode(Nvl(n_��������, 0), 0, '�����䶯', '�����䶯����') || '��������������', v_����Ա���, v_����Ա����);
      End If;
    
      --2.�����շ���Ŀ(ҩƷδ����)
    Else
      --�������:
      --1.��ԭʼ��¼��������
      --2.����Ŀ�겡������
      --3.����ǻ��۵���ֱ�Ӹ���ԭ��¼����id��ִ�в���
      If Nvl(r_����.��¼״̬, 0) = 0 Then
        --ֱ���޸�(�������˲�����Ŀ�겡��)
        Update סԺ���ü�¼
        Set ���˲���id = n_Ŀ�겡��id, ִ�в���id = n_Ŀ�겡��id
        Where NO = r_����.No And ��¼���� = 2 And ��¼״̬ = 0 And Nvl(�۸񸸺�, ���) = r_����.���;
      
        Insert Into ���ñ䶯��¼
          (ID, ��¼״̬, ����id, ��ҳid, �䶯ʱ��, ԭ�䶯id, Ŀ��䶯id, ԭ����id, Ŀ�겡��id, ����id, NO, �շ����, �շ�ϸĿid, ҽ�����, ����, ����, Ӧ�ս��, ʵ�ս��,
           ״̬, ժҪ, ����Ա���, ����Ա����)
          Select ���ñ䶯��¼_Id.Nextval, Decode(Nvl(n_��������, 0), 0, 1, 2), n_����id, n_��ҳid, d_�Ǽ�ʱ��, n_ԭ�䶯id, n_Ŀ��䶯id, n_ԭ����id,
                 n_Ŀ�겡��id, ����id, NO, �շ����, �շ�ϸĿid, ҽ�����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, 0,
                 Decode(Nvl(n_��������, 0), 0, '�����䶯', '�����䶯����') || '�޸ļ��ʻ��۵�', v_����Ա���, v_����Ա����
          From (Select Max(Decode(�۸񸸺�, Null, ID, 0)) As ����id, NO, �շ����, �շ�ϸĿid, ҽ�����, Avg(Nvl(����, 1) * ����) As ����,
                        Sum(��׼����) As ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��
                 From סԺ���ü�¼
                 Where NO = r_����.No And ��¼���� = 2 And ��¼״̬ = 0 And Nvl(�۸񸸺�, ���) = r_����.���
                 Group By NO, �շ����, �շ�ϸĿid, ҽ�����, Nvl(�۸񸸺�, ���));
      
      Elsif Nvl(n_ת������, 0) > 0 Then
        --ֱ�����ʴ���
        --��ţ����1:����1:ִ��״̬1,���2:����2:ִ��״̬2,...���n:����n:ִ��״̬n  ��:"1:2:1,2:10:1,3:2:1"
        --1.�Ȳ������ʼ�¼
        Zl_סԺ���ʼ�¼_Delete_s(r_����.No, r_����.��� || ':' || n_ת������ || ':0', v_����Ա���, v_����Ա����, 2, 2, d_�Ǽ�ʱ��);
        --2.Ŀ�겡��ת���¼
        For c_��ϸ In (Select ���˷��ü�¼_Id.Nextval As ����id, NO, ��¼����, 1 As ��¼״̬, n_Max��� + Rownum As ���, ��������,
                            �۸񸸺� + (n_Max��� + Rownum - ���) As �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��,
                            ����, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ���� As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                            ��׼����, -1 * Ӧ�ս�� As Ӧ�ս��, -1 * ʵ�ս�� As ʵ�ս��, ��������id, ������, ������, ִ����, r_����.ִ��״̬ As ִ��״̬, ִ��ʱ��,
                            ����ʱ��, ������Ŀ��, ���մ���id, -1 * ͳ���� As ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���, ����, ҽ��С��id
                     From סԺ���ü�¼
                     Where NO = r_����.No And Nvl(�۸񸸺�, ���) = r_����.��� And ��¼״̬ = 2 And �Ǽ�ʱ�� = d_�Ǽ�ʱ��) Loop
        
          Insert Into סԺ���ü�¼
            (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���˲���id,
             ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������,
             ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���,
             ����, ҽ��С��id)
          Values
            (c_��ϸ.����id, c_��ϸ.No, c_��ϸ.��¼����, c_��ϸ.��¼״̬, c_��ϸ.���, c_��ϸ.��������, c_��ϸ.�۸񸸺�, c_��ϸ.��ҳid, c_��ϸ.����id, c_��ϸ.ҽ�����,
             c_��ϸ.�����־, c_��ϸ.�ಡ�˵�, c_��ϸ.Ӥ����, c_��ϸ.����, c_��ϸ.�Ա�, c_��ϸ.����, c_��ϸ.��ʶ��, c_��ϸ.����, c_��ϸ.�ѱ�, n_Ŀ�겡��id,
             c_��ϸ.���˿���id, c_��ϸ.�շ����, c_��ϸ.�շ�ϸĿid, c_��ϸ.���㵥λ, c_��ϸ.����, c_��ϸ.��ҩ����, c_��ϸ.����, c_��ϸ.�Ӱ��־, c_��ϸ.���ӱ�־,
             c_��ϸ.������Ŀid, c_��ϸ.�վݷ�Ŀ, c_��ϸ.���ʷ���, c_��ϸ.��׼����, c_��ϸ.Ӧ�ս��, c_��ϸ.ʵ�ս��, c_��ϸ.��������id, c_��ϸ.������, n_Ŀ�겡��id,
             c_��ϸ.������, c_��ϸ.ִ����, /*c_��ϸ.ִ��״̬*/ 0, c_��ϸ.ִ��ʱ��, v_����Ա���, v_����Ա����, c_��ϸ.����ʱ��, d_�Ǽ�ʱ��, c_��ϸ.������Ŀ��,
             c_��ϸ.���մ���id, c_��ϸ.ͳ����, c_��ϸ.���ձ���, c_��ϸ.���ʵ�id, c_��ϸ.ժҪ, c_��ϸ.��������, c_��ϸ.�Ƿ���, c_��ϸ.����, c_��ϸ.ҽ��С��id);
        
          Update ���ñ䶯��¼
          Set ���� = Nvl(����, 0) + Nvl(c_��ϸ.��׼����, 0), Ӧ�ս�� = Nvl(Ӧ�ս��, 0) + Nvl(c_��ϸ.Ӧ�ս��, 0),
              ʵ�ս�� = Nvl(ʵ�ս��, 0) + Nvl(c_��ϸ.ʵ�ս��, 0)
          Where ����id = r_����.����id And �䶯ʱ�� = d_�Ǽ�ʱ�� And Ŀ��䶯id = n_Ŀ��䶯id And �շ�ϸĿid = r_����.�շ�ϸĿid And
                ����id + 0 = c_��ϸ.����id;
          If Sql%NotFound Then
            Insert Into ���ñ䶯��¼
              (ID, ��¼״̬, ����id, ��ҳid, �䶯ʱ��, ԭ�䶯id, Ŀ��䶯id, ԭ����id, Ŀ�겡��id, ����id, NO, �շ����, �շ�ϸĿid, ҽ�����, ����, ����, Ӧ�ս��,
               ʵ�ս��, ״̬, ժҪ, ����Ա���, ����Ա����)
            Values
              (���ñ䶯��¼_Id.Nextval, Decode(Nvl(n_��������, 0), 0, 1, 2), n_����id, n_��ҳid, d_�Ǽ�ʱ��, n_ԭ�䶯id, n_Ŀ��䶯id, n_ԭ����id,
               n_Ŀ�겡��id, r_����.����id, r_����.No, r_����.�շ����, r_����.�շ�ϸĿid, r_����.ҽ��id, Round(c_��ϸ.���� * Nvl(c_��ϸ.����, 1), 5),
               c_��ϸ.��׼����, c_��ϸ.Ӧ�ս��, c_��ϸ.ʵ�ս��, 1, Decode(Nvl(n_��������, 0), 0, '�����䶯', '�����䶯����') || '�޸ļ��ʵ�', v_����Ա���,
               v_����Ա����);
          End If;
        
          Update ����������Ŀ
          Set �������� = Nvl(��������, 0) + Round(c_��ϸ.���� * Nvl(c_��ϸ.����, 1), 5)
          Where ����id = n_����id And ��ҳid = n_��ҳid And ��Ŀid = c_��ϸ.�շ�ϸĿid And Nvl(ʹ������, 0) <> 0;
        
          --�������
          Update �������
          Set ������� = Nvl(�������, 0) + c_��ϸ.ʵ�ս��
          Where ����id = c_��ϸ.����id And ���� = 2 And ���� = 1;
          If Sql%RowCount = 0 Then
            Insert Into �������
              (����id, ����, ����, �������, Ԥ�����)
            Values
              (c_��ϸ.����id, 2, 1, c_��ϸ.ʵ�ս��, 0);
          End If;
          --����δ�����
          Update ����δ�����
          Set ��� = Nvl(���, 0) + c_��ϸ.ʵ�ս��
          Where ����id = c_��ϸ.����id And Nvl(��ҳid, 0) = Nvl(c_��ϸ.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(n_Ŀ�겡��id, 0) And
                Nvl(���˿���id, 0) = Nvl(c_��ϸ.���˿���id, 0) And Nvl(��������id, 0) = Nvl(c_��ϸ.��������id, 0) And
                Nvl(ִ�в���id, 0) = Nvl(n_Ŀ�겡��id, 0) And ������Ŀid + 0 = c_��ϸ.������Ŀid And ��Դ;�� + 0 = 2;
          If Sql%RowCount = 0 Then
            Insert Into ����δ�����
              (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
            Values
              (c_��ϸ.����id, c_��ϸ.��ҳid, n_Ŀ�겡��id, c_��ϸ.���˿���id, c_��ϸ.��������id, n_Ŀ�겡��id, c_��ϸ.������Ŀid, 2, c_��ϸ.ʵ�ս��);
          End If;
        
          n_Max��� := c_��ϸ.���;
        End Loop;
      End If;
    End If;
  End Loop;

  If Nvl(n_��������, 0) = 1 Then
    --��������,��Ҫɾ���������ϲ�����δ��˲���
    For c_���� In (Select a.����id, a.����id, a.��ҳid, a.No, a.�շ����, a.�շ�ϸĿid, a.ҽ�����, a.����, a.����, a.Ӧ�ս��, a.ʵ�ս��
                 From ���ñ䶯��¼ A
                 Where a.ԭ�䶯id = n_Ŀ��䶯id And a.Ŀ��䶯id = n_ԭ�䶯id And a.״̬ = 2) Loop
    
      Select Sum(����) Into n_����ȡ������ From ���˷������� Where ����id = c_����.����id And ״̬ In (0, 2);
    
      If Nvl(n_����ȡ������, 0) > 0 Then
        n_Ӧ�ս�� := Round(n_����ȡ������ * Nvl(c_����.����, 0), n_Dec);
        n_ʵ�ս�� := 0;
        If Nvl(c_����.Ӧ�ս��, 0) <> 0 Then
          n_ʵ�ս�� := Round(Nvl(n_Ӧ�ս��, 0) * Nvl(c_����.ʵ�ս��, 0) / c_����.Ӧ�ս��, n_Dec);
        End If;
      
        Insert Into ���ñ䶯��¼
          (ID, ��¼״̬, ����id, ��ҳid, �䶯ʱ��, ԭ�䶯id, Ŀ��䶯id, ԭ����id, Ŀ�겡��id, ����id, NO, �շ����, �շ�ϸĿid, ҽ�����, ����, ����, Ӧ�ս��, ʵ�ս��,
           ״̬, ժҪ, ����Ա���, ����Ա����)
        Values
          (���ñ䶯��¼_Id.Nextval, 2, c_����.����id, c_����.��ҳid, d_�Ǽ�ʱ��, n_ԭ�䶯id, n_Ŀ��䶯id, n_ԭ����id, n_Ŀ�겡��id, c_����.����id, c_����.No,
           c_����.�շ����, c_����.�շ�ϸĿid, c_����.ҽ�����, n_����ȡ������, c_����.����, n_Ӧ�ս��, n_ʵ�ս��, 3, '����������ɾ����������', v_����Ա���, v_����Ա����);
      End If;
    
      Delete ���˷������� Where ����id = c_����.����id And ״̬ In (0, 2);
    End Loop;
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Executeturnwardfee;
/

Create Or Replace Procedure Zl_Exsesvr_Getfeechangerec
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡ���ñ䶯��¼
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id           N   1 ����ID
  --   pati_pageid       N   1 ��ҳID
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]
  --      bill_no          C   ���õ��ݺ�
  --      item_id          N   �շ�ϸĿID
  --      item_name        C   �շ�ϸĿ����
  --      ward_id_old      N   ԭ����id
  --      ward_name_old    N   ԭ��������
  --      ward_id_new      N   Ŀ�겡��id
  --      ward_name_new    N   Ŀ�겡������
  --      quantity         N   ����
  --      price            N   ����
  --      fee_ampaid       N   ʵ�ս��
  --      rec_type         N   ��¼����:0-ֱ�Ӹ���ԭ���ݣ�1-����������ת�����ݣ�2-��������������䶯��3-ȡ����������䶯
  --      change_time      C   �䶯ʱ��:yyyy-mm-dd hh24:mi:ss
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);
  c_Output Clob;

  n_����id ���ñ䶯��¼.����id%Type;
  n_��ҳid ���ñ䶯��¼.��ҳid%Type;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  For r_�䶯 In (Select a.No, a.�շ�ϸĿid, Decode(j.�Ƿ���, 1, '***', b.����) As ��Ŀ����, a.ԭ����id, c.���� As ԭ����, a.Ŀ�겡��id,
                      d.���� As Ŀ�겡��, a.����, a.����, a.ʵ�ս��, a.״̬, To_Char(a.�䶯ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �䶯ʱ��
               From ���ñ䶯��¼ A, �շ���ĿĿ¼ B, ���ű� C, ���ű� D, סԺ���ü�¼ J
               Where a.�շ�ϸĿid = b.Id And a.ԭ����id = c.Id And a.Ŀ�겡��id = d.Id And a.����id = n_����id And a.��ҳid = n_��ҳid And
                     a.����id = j.Id) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'bill_no', r_�䶯.No, 0, 1);
    zlJsonPutValue(v_Output, 'item_id', r_�䶯.�շ�ϸĿid, 1);
    zlJsonPutValue(v_Output, 'item_name', r_�䶯.��Ŀ����);
    zlJsonPutValue(v_Output, 'ward_id_old', r_�䶯.ԭ����id, 1);
    zlJsonPutValue(v_Output, 'ward_name_old', r_�䶯.ԭ����);
    zlJsonPutValue(v_Output, 'ward_id_new', r_�䶯.Ŀ�겡��id, 1);
    zlJsonPutValue(v_Output, 'ward_name_new', r_�䶯.Ŀ�겡��);
    zlJsonPutValue(v_Output, 'quantity', r_�䶯.����, 1);
    zlJsonPutValue(v_Output, 'price', r_�䶯.����, 1);
    zlJsonPutValue(v_Output, 'fee_ampaid', r_�䶯.ʵ�ս��, 1);
    zlJsonPutValue(v_Output, 'rec_type', Nvl(r_�䶯.״̬, 0), 1);
    zlJsonPutValue(v_Output, 'change_time', r_�䶯.�䶯ʱ��, 1, 2);
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeechangerec;
/


Create Or Replace Procedure Zl_Exsesvr_Getturnwardfee
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:��ȡת��������
  --��Σ�Json_In:��ʽ
  -- input
  --   pati_id           N   1 ����ID
  --   pati_pageid       N   1 ��ҳID
  --   exe_deptid        N   1 ִ�в���ID
  --����: Json_Out,��ʽ����
  --  output
  --    code               C  1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message            C  1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]
  --      rec_type         N   ��¼����:1-����ת�Ƽ�¼��2-���������¼
  --      bill_prop        N   ��������:0-���ʻ��۵�,1-���ʵ�
  --      bill_no          C   ���õ��ݺ�
  --      serial_num       N   ���õ������
  --      item_id          N   �շ�ϸĿID
  --      item_name        C   �շ�ϸĿ����
  --      advice_id        N   ҽ�����
  --      quantity         N   ����
  --      price            N   ����
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id     סԺ���ü�¼.����id%Type;
  n_��ҳid     סԺ���ü�¼.��ҳid%Type;
  n_ִ�в���id סԺ���ü�¼.ִ�в���id%Type;

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := j_Json.Get_Number('pati_id');
  n_��ҳid     := j_Json.Get_Number('pati_pageid');
  n_ִ�в���id := j_Json.Get_Number('exe_deptid');
  For r_�䶯 In (Select a.��¼����, Decode(a.��¼״̬, 0, 0, 1) As ��������, a.No, a.���, a.�շ�ϸĿid,
                      Max(Decode(a.�Ƿ���, 1, '***', b.����)) As �շ���Ŀ, a.ҽ�����, Sum(a.ʣ������) As ʣ������, Max(a.��׼����) As ��׼����
               From (
                    
                    With סԺ���� As (Select a.No, a.���, Max(a.�շ�ϸĿid) As �շ�ϸĿid, Sum(����) As ʣ������,
                                         Max(Decode(a.��¼״̬, 2, 0, a.����id)) As ����id, Max(a.��¼״̬) As ��¼״̬,
                                         Max(a.ҽ�����) As ҽ�����, Max(a.�Ƿ���) As �Ƿ���, Max(a.��׼����) As ��׼����
                                  From (Select a.No, ��¼״̬, Nvl(a.�۸񸸺�, ���) As ���, �շ�ϸĿid, Avg(Nvl(a.����, 1) * a.����) As ����,
                                                Max(Decode(a.�۸񸸺�, Null, a.Id, 0)) As ����id, Max(a.ҽ�����) As ҽ�����,
                                                Max(a.�Ƿ���) As �Ƿ���, Sum(a.��׼����) As ��׼����
                                         From סԺ���ü�¼ A, �������� C
                                         Where a.��¼���� = 2 And a.ִ�в���id = n_ִ�в���id And a.ҽ����� Is Not Null And
                                               Nvl(a.�Ƿ񸽷�, 0) = 0 And a.����id = n_����id And a.��ҳid = n_��ҳid And
                                               Instr(',5,6,7,', ',' || a.�շ���� || ',') = 0 And a.�շ�ϸĿid = c.����id And
                                               Nvl(c.��������, 0) = 1
                                         Group By a.No, ��¼״̬, a.ҽ�����, Nvl(a.�۸񸸺�, ���), �շ�ϸĿid, a.ִ��״̬) A
                                  Group By a.No, a.���)
                    --��Ҫ�������������,��ȥ����������
                      Select 2 As ��¼����, a.No, a.���, Max(a.��¼״̬) As ��¼״̬, a.�շ�ϸĿid, Max(a.ҽ�����) As ҽ�����,
                             Nvl(Sum(a.ʣ������), 0) - Nvl(Sum(b.����), 0) As ʣ������, Sum(a.��׼����) As ��׼����, Max(a.�Ƿ���) As �Ƿ���
                      From סԺ���� A,
                           (Select b.����id, Nvl(Sum(b.����), 0) As ����
                             From סԺ���� A, ���˷������� B
                             Where a.����id = b.����id And Nvl(b.״̬, 0) = 0
                             Group By b.����id
                             Having Nvl(Sum(b.����), 0) <> 0) B
                      Where a.����id = b.����id(+)
                      Group By a.No, a.���, a.�շ�ϸĿid
                      Having Nvl(Sum(a.ʣ������), 0) - Nvl(Sum(b.����), 0) <> 0
                      Union All
                      Select 1 As ��¼����, a.No, Nvl(a.�۸񸸺�, a.���) As ���, a.��¼״̬, a.�շ�ϸĿid, a.ҽ�����,
                             Avg(Nvl(a.����, 1) * a.����) As ʣ������, Sum(a.��׼����) As ��׼����, Max(a.�Ƿ���) As �Ƿ���
                      From סԺ���ü�¼ A, �������� C
                      Where a.�շ�ϸĿid = c.����id(+) And a.��¼���� = 2 And a.ִ�в���id = n_ִ�в���id And a.ҽ����� Is Not Null And
                            Nvl(a.�Ƿ񸽷�, 0) = 0 And a.����id = n_����id And a.��ҳid = n_��ҳid And
                            Instr(',5,6,7,', ',' || a.�շ���� || ',') = 0 And Nvl(c.��������, 0) = 0
                      Group By a.No, a.��¼״̬, a.ҽ�����, Nvl(a.�۸񸸺�, a.���), a.�շ�ϸĿid
                      
                      ) A, �շ���ĿĿ¼ B
                      Where a.�շ�ϸĿid = b.Id
                      Group By a.��¼����, Decode(a.��¼״̬, 0, 0, 1), a.No, a.���, a.�շ�ϸĿid, a.ҽ�����
               ) Loop
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'rec_type', r_�䶯.��¼����, 1, 1);
    zlJsonPutValue(v_Output, 'bill_prop', r_�䶯.��������, 1);
    zlJsonPutValue(v_Output, 'bill_no', r_�䶯.No);
    zlJsonPutValue(v_Output, 'serial_num', r_�䶯.���, 1);
    zlJsonPutValue(v_Output, 'item_id', r_�䶯.�շ�ϸĿid, 1);
    zlJsonPutValue(v_Output, 'item_name', r_�䶯.�շ���Ŀ);
    zlJsonPutValue(v_Output, 'advice_id', r_�䶯.ҽ�����, 1);
    zlJsonPutValue(v_Output, 'quantity', r_�䶯.ʣ������, 1);
    zlJsonPutValue(v_Output, 'price', r_�䶯.��׼����, 1, 2);
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","item_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","item_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
     Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getturnwardfee;
/


Create Or Replace Procedure Zl_Exsesvr_Getbillgrpbyfeetype
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ����շ��������ȡ���õ�����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --     query_type   N 1 ��ѯ��ʽ��0-�����ݺ�,1-��ҽ��ID
  --     bill_nos     C 0 ���ݺţ���������,�ö��ŷָ�,��:A00001,A0002,...,A000n
  --     advice_ids   C 0 ҽ��ID����������,�ö��ŷָ�,��:1,2,3,4
  --     pati_id      N 0 ����ID�����ʱ�ʱ�����˻�ȡ
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]
  --      pati_id         N 1 ����ID
  --      pati_pageid     N 1 ��ҳID
  --      pati_name       C 1 ��������
  --      fee_type        C 1 �������
  --      fee_type_name   C 1 �����������
  --      ward_id         N 1 ����ID
  --      fee_ampaid      N 1 ʵ�ս��ϼ�
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_��ѯ��ʽ Number;
  v_Nos      Varchar2(32767);
  v_ҽ��ids  Varchar2(32767);
  n_����id   ������ü�¼.����id%Type;

  n_Firstitem Number(1);
  v_Temp      Varchar2(32767);
  c_Temp      Clob;
Begin
  --�������

  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ��ʽ := j_Json.Get_Number('query_type');

  n_Firstitem := 1;
  v_Temp      := '{"output":{"code":1,"message":"�ɹ�","item_list":[';
  If Nvl(n_��ѯ��ʽ, 0) = 0 Then
    v_Nos    := j_Json.Get_String('bill_nos');
    n_����id := j_Json.Get_Number('pati_id');
  
    For r_���� In (Select m.���, j.��� As ����, m.����id, Sum(m.ʵ�ս��) As ���, m.����id, m.��ҳid, m.����
                 From (Select /*+cardinality(b,10)*/
                         a.����id, 0 As ��ҳid, a.����, a.�շ���� As ���, 0 As ����id, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս��
                        From ������ü�¼ A, Table(f_Str2List(v_Nos)) B
                        Where a.No = b.Column_Value And ���ʷ��� = 1 And ��¼״̬ = 0 And ��¼���� = 2 And
                              (Nvl(n_����id, 0) = 0 Or a.����id = n_����id)
                        Group By a.�շ����, a.����id, a.����
                        Union All
                        Select /*+cardinality(b,10)*/
                         a.����id, a.��ҳid, a.����, a.�շ���� As ���, a.���˲���id As ����id, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս��
                        From סԺ���ü�¼ A, Table(f_Str2List(v_Nos)) B
                        Where a.No = b.Column_Value And ���ʷ��� = 1 And ��¼״̬ = 0 And ��¼���� = 2 And
                              (Nvl(n_����id, 0) = 0 Or a.����id = n_����id)
                        Group By a.�շ����, a.����id, a.��ҳid, a.����, a.���˲���id) M, �շ���� J
                 Where m.��� = j.����
                 Group By m.����id, m.��ҳid, m.����, m.���, j.���, m.����id) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"pati_id":' || r_����.����id;
      v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_����.��ҳid, 0);
      v_Temp := v_Temp || ',"pati_name":"' || zlJsonStr(r_����.����) || '"';
      v_Temp := v_Temp || ',"fee_type":"' || r_����.��� || '"';
      v_Temp := v_Temp || ',"fee_type_name":"' || r_����.���� || '"';
      v_Temp := v_Temp || ',"ward_id":' || Nvl(r_����.����id, 0);
      v_Temp := v_Temp || ',"fee_ampaid":' || zlJsonStr(r_����.���, 1);
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 20000 Then
        c_Temp := c_Temp || To_Clob(v_Temp);
        v_Temp := '';
      End If;
    End Loop;
  
  Else
    v_ҽ��ids := j_Json.Get_String('advice_ids');
  
    For r_���� In (Select m.���, j.��� As ����, m.����id, Sum(m.ʵ�ս��) As ���, m.����id, m.��ҳid, m.����
                 From (Select /*+cardinality(b,10)*/
                         a.����id, 0 As ��ҳid, a.����, a.�շ���� As ���, 0 As ����id, Nvl(Sum(a.ʵ�ս��), 0) As ʵ�ս��
                        From ������ü�¼ A, Table(f_Num2List(v_ҽ��ids)) B
                        Where a.ҽ����� = b.Column_Value And ���ʷ��� = 1 And ��¼״̬ = 0
                        Group By a.�շ����, a.����id, a.����
                        Union All
                        Select /*+cardinality(b,10)*/
                         a.����id, a.��ҳid, a.����, a.�շ���� As ���, a.���˲���id As ����id, Nvl(Sum(ʵ�ս��), 0) As ʵ�ս��
                        From סԺ���ü�¼ A, Table(f_Num2List(v_ҽ��ids)) B
                        Where a.ҽ����� = b.Column_Value And ���ʷ��� = 1 And ��¼״̬ = 0
                        Group By a.�շ����, a.����id, a.��ҳid, a.����, a.���˲���id) M, �շ���� J
                 Where m.��� = j.����
                 Group By m.����id, m.��ҳid, m.����, m.���, j.���, m.����id) Loop
    
      If Nvl(n_Firstitem, 0) = 0 Then
        v_Temp := v_Temp || ',';
      Else
        n_Firstitem := 0;
      End If;
      v_Temp := v_Temp || '{';
      v_Temp := v_Temp || '"pati_id":' || r_����.����id;
      v_Temp := v_Temp || ',"pati_pageid":' || Nvl(r_����.��ҳid, 0);
      v_Temp := v_Temp || ',"pati_name":"' || zlJsonStr(r_����.����) || '"';
      v_Temp := v_Temp || ',"fee_type":"' || r_����.��� || '"';
      v_Temp := v_Temp || ',"fee_type_name":"' || r_����.���� || '"';
      v_Temp := v_Temp || ',"ward_id":' || Nvl(r_����.����id, 0);
      v_Temp := v_Temp || ',"fee_ampaid":' || zlJsonStr(r_����.���, 1);
      v_Temp := v_Temp || '}';
    
      If Length(v_Temp) > 30000 Then
        c_Temp := c_Temp || To_Clob(v_Temp);
        v_Temp := '';
      End If;
    End Loop;
  End If;
  v_Temp := v_Temp || ']}}';

  If c_Temp Is Not Null Then
    Json_Out := c_Temp || To_Clob(v_Temp);
  Else
    Json_Out := v_Temp;
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillgrpbyfeetype;
/

Create Or Replace Procedure Zl_Exsesvr_Getchargeoffapply
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���id����ҳid,��ȡ���˵�����������Ϣ
  --��Σ�Json_In:��ʽ
  --input
  -- pati_id N 1 ����id
  -- pati_pageid N 1 ��ҳid
  --����: Json_Out,��ʽ����
  --output
  --  code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  exists  N 1 �Ƿ����:1-����;0-������
  --    apply_list[]  C   �����б�
  --    fee_no  C 1 ���õ��ݺ�
  --    fitem_name  C 1 �շ���Ŀ����
  --    audit_dept_name C 1 ��˲�������

  ---------------------------------------------------------------------------

  n_����id Number(18);
  n_��ҳid Number(18);

  n_Count Number;
  j_Input PLJson;
  j_Json  PLJson;

  v_Output  Varchar2(32767);
  n_Isexist Number(1);
Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');
  n_��ҳid := j_Json.Get_Number('pati_pageid');

  n_Isexist := 0;
  For c_���� In (Select Distinct a.No, d.���� ��Ŀ����, c.���� ��˿���
               From סԺ���ü�¼ A, ���˷������� B, ���ű� C, �շ���ĿĿ¼ D
               Where a.����id = n_����id And a.��ҳid = n_��ҳid And a.Id = b.����id And b.״̬ = 0 And b.��˲���id = c.Id And
                     b.�շ�ϸĿid = d.Id
               Order By a.No, c.����
               
               ) Loop
  
    zlJsonPutValue(v_Output, 'fee_no', c_����.No, 0, 1);
    zlJsonPutValue(v_Output, 'fitem_name', c_����.��Ŀ����);
    zlJsonPutValue(v_Output, 'audit_dept_name', c_����.��˿���, 0, 2);
  
    n_Isexist := 1;
  End Loop;

  If n_Count > 0 Then
    n_Isexist := 1;
  Else
    n_Isexist := 0;
  End If;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exists":' || n_Isexist || ',"apply_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getchargeoffapply;
/


Create Or Replace Procedure Zl_Exsesvr_Existspricebill
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���id��ҽ��id�ж��Ƿ���ڶ�Ӧ�Ļ��۵�
  --��Σ�Json_In:��ʽ
  --input      
  -- pati_id N 1 ����id
  -- pati_pageid N 1 ��ҳid
  -- advice_ids  C   ҽ��id:����ö���
  -- billtype  N 1 ��������:1-�շѻ��۵�;2-���ʻ��۵�

  --����: Json_Out,��ʽ����
  --  output
  --       code             N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --       message          C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --     exists N 1 �Ƿ����:1-����;0-������
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  n_����id   Number(18);
  n_��ҳid   Number(18);
  v_ҽ��ids  Varchar2(32767);
  n_�������� Number(2);
  n_Count    Number(5);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_��ҳid   := Nvl(j_Json.Get_Number('pati_pageid'), 0);
  v_ҽ��ids  := j_Json.Get_String('advice_ids');
  n_�������� := Nvl(j_Json.Get_Number('billtype'), 1);

  If v_ҽ��ids Is Null Then
    If n_�������� = 1 Then
      Select Max(1)
      Into n_Count
      From ������ü�¼
      Where ��¼���� = 1 And (��¼״̬ = 0 Or ��¼״̬ = 1 And ����id Is Null) And ����id = n_����id And Rownum < 2;
    
    Else
      If Nvl(n_��ҳid, 0) = 0 Then
        Select Max(1)
        Into n_Count
        From (Select 1
               From ������ü�¼
               Where ��¼״̬ = 0 And Nvl(���ʷ���, 0) = 1 And ����id = n_����id And Rownum < 2
               Union All
               Select 1
               From סԺ���ü�¼
               Where ��¼״̬ = 0 And Nvl(���ʷ���, 0) = 1 And �����־ <> 2 And ����id = n_����id And Rownum < 2);
      Else
        Select 1
        Into n_Count
        From סԺ���ü�¼
        Where ��¼״̬ = 0 And Nvl(���ʷ���, 0) = 1 And ����id = n_����id And ��ҳid = n_��ҳid And Rownum < 2;
      End If;
    End If;
  
  Else
  
    Select Max(1)
    Into n_Count
    From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Num2List(v_ҽ��ids)))
           Select /*+cardinality(B,10) */
            1
           From ������ü�¼ A, ҽ������ B
           Where a.ҽ����� = b.ҽ��id And a.��¼״̬ = 0 And Nvl(a.���ʷ���, 0) = 1 And Rownum < 2
           Union All
           Select /*+cardinality(B,10) */
            1
           From סԺ���ü�¼ A, ҽ������ B
           Where a.ҽ����� = b.ҽ��id And a.��¼״̬ = 0 And Nvl(a.���ʷ���, 0) = 1 And Rownum < 2);
  
  
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","exists":' || n_Count || '}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Existspricebill;
/
CREATE OR REPLACE Procedure Zl_Exsesvr_Getdrugerrdata
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID��ҽ����Ϣ���ز��˷�����Ϣ
  --��Σ�Json_In:��ʽ����ֵΪnullʱ��ȡ���в��˵��쳣��Ϣ
  --  input
  --    pati_list[]�����б�
  --       pati_id                    N 1 ����id
  --       bill_list[]                ���õ��ݺ��б����Բ���������ʱ��ʾ��ȡ������ͬ���쳣������
  --         fee_source               N 0 ������Դ��1-���2-סԺ
  --         fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
  --         fee_no                   C 0 ���õ��ݺ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                          N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_bill_list[]
  --       billtype                   N   1 ��������: 1 -�շѴ���  ;2- ���ʵ�����;3- ���ʱ���
  --       pati_source                N   1 ������Դ:1-����;2-סԺ;4-���
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������½ڵ�--------------------------------------
  --       pati_id                    N   1 ����ID
  --       pati_pageid                N   1 ��ҳID
  --       pati_name                  C   1 ��������
  --       pati_sex_code              C   1 �Ա��ţ�������)
  --       pati_sex                   C   1 �Ա�
  --       pati_age                   C   1 ����
  --       pati_deptid                N   1 ���˿���ID
  --       pati_wardarea_id           N     ���˲���ID
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������Ͻڵ�-----------------------------------------
  --       bill_list[]                      ���������б�[����]
  --         fee_source                N  0 ������Դ
  --         rcp_no                    C  1 NO
  --         recipe_type               N  0 ��������:0�Ϳ�-��ͨ,1-����,2-����,3-����,4-��һ,5-����
  --         charge_tag                N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
  --         fee_acnter                C  0 ������
  --         recipe_plcdept_id         C  0 ��������id��������)
  --         recipe_plcdept            C  0 �����������ƣ�������)
  --         recipe_placer_id          C  0 ����ҽʦid��������)
  --         recipe_placer             C  0 ����ҽʦ��������) ����
  --         operator_name             C  1 ����Ա����
  --         operator_code             C  1 ����Ա���
  --         create_time               C  1 �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
  --         item_list[]                    ���������б�[����]

  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������½ڵ�----------------------------------------
  --           pati_id                 N  1 ����ID
  --           pati_pageid             N  0 ��ҳID
  --           pati_name               C  1 ��������
  --           pati_sex_code           C  1 �Ա��ţ�������)
  --           pati_sex                C  1 �Ա�
  --           pati_age                C  1 ����
  --           pati_wardarea_id        N  0 ���˲���ID
  --           pati_deptid             N  1 ���˿���ID
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=10��ʱ�������Ͻڵ�-----------------------------------------

  --           rcpdtl_id               N  1 ������ϸID
  --           serial_num              N  1 ���:(���(�����洢)����ź���ţ�1��2��3��3��3��4��)
  --           pharmacy_id             N  1 ҩ��ID
  --           pharmacy_name           C  1 ҩ������(������)
  --           takedept_id             N  1 ��ҩ����ID:���סԺ�Ŵ���
  --           drug_id                 N  1 ҩƷID
  --           baby_num                N  0  Ӥ�����
  --           advice_id               N  0 ҽ��ID
  --           decoction_method        C  0 �巨
  --           use_mode                N  0 ȡҩ���ԣ�0-������ʽ��1-��Ժ��ҩ��2-��ȡҩ
  --           packages_num            N  1 ��ҩ����
  --           send_num                N  1 ��ҩ����
  --           send_unit               C  1 ��ҩ��λ��zlhis���۵�λ
  --           price                   N  0 �ۼ�
  --           money                   N  0 ���۽��(������)
  --           pharmacy_window         C  0 ��ҩ����
  --           memo                    C  0 ժҪ
  ------------------------------------------------------------------------------------------------------------
  j_Json      PLJson;
  j_Json_In   PLJson;
  j_Pati_List Pljson_List;
  j_Json_Out  PLJson;
  j_Bill_List Pljson_List;

  Json_Temp_Out Clob;
  c_Jtmp        Clob;

  j_Item   PLJson;
  n_����id Number(18);

  v_Json Varchar2(4000);
  n_Code Number;

  n_������Դ Number(1);
  n_��¼���� ������ü�¼.��¼����%Type;
  v_No       ������ü�¼.No%Type;

  l_Outnos t_StrList2 := t_StrList2();
  l_Innos  t_StrList2 := t_StrList2();
Begin
  If Json_In Is Null Then
    --����ϵͳ�мǷ�ͬ���쳣������
    For r_Fee In (Select Distinct 1 As ������Դ, Decode(Mod(a.��¼����, 10), 2, 2, 1) As ��������, a.No
                  From ������ü�¼ A, ���˷����쳣��¼ B
                  Where a.�շ���� In ('5', '6', '7') And a.��¼���� In (1, 2) And
                        a.id = b.����id And (b.�������� = 0 Or b.�������� = 1) And Nvl(b.ͬ����־, 0) = 1 And
                        Exists (Select 1
                         From ������ü�¼
                         Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                         Group By ��¼����, NO, ���
                         Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)
                  Union All
                  Select Distinct 2 As ������Դ, Decode(a.�ಡ�˵�, 1, 3, 2) As ��������, a.No
                  From סԺ���ü�¼ A, ���˷����쳣��¼ B
                  Where a.�շ���� In ('5', '6', '7') And a.��¼���� = 2 And
                        a.id = b.����id And b.�������� = 0 And Nvl(b.ͬ����־, 0) = 1 And Exists
                   (Select 1
                         From סԺ���ü�¼
                         Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                         Group By ��¼����, NO, ���
                         Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)) Loop

      v_Json := Null;
      v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
      v_Json := v_Json || ',"billtype":' || r_Fee.��������;
      v_Json := v_Json || ',"fee_source":' || r_Fee.������Դ;
      v_Json := v_Json || '}';
      v_Json := '{"input":' || v_Json || '}';

      Json_Temp_Out := Null;
      Zl_Drugbill_Build(v_Json, Json_Temp_Out);

      --��������
      j_Json_Out := PLJson();
      j_Json_Out := PLJson(Json_Temp_Out);
      j_Json     := PLJson();
      j_Json     := j_Json_Out.Get_Pljson('output');

      n_Code := Nvl(j_Json.Get_Number('code'), '0');
      If n_Code = 0 Then
        Json_Out := zlJsonOut(j_Json.Get_String('message'));
        Return;
      End If;

      j_Json.Remove('code');
      j_Json.Remove('message');
      Json_Temp_Out := Empty_Clob();
      Dbms_Lob.Createtemporary(Json_Temp_Out, True);
      j_Json.To_Clob(Json_Temp_Out);

      If c_Jtmp Is Null Then
        c_Jtmp := Json_Temp_Out;
      Else
        c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
      End If;
    End Loop;
  Else
    --�������
    j_Json_In   := PLJson(Json_In);
    j_Json      := j_Json_In.Get_Pljson('input');
    j_Pati_List := j_Json.Get_Pljson_List('pati_list');
    
    For I In 1 .. j_Pati_List.Count Loop
    
      j_Item      := PLJson();
      j_Item      := PLJson(j_Pati_List.Get(I));
      n_����id    := j_Item.Get_Number('pati_id');
      j_Bill_List := j_Item.Get_Pljson_List('bill_list');

      If j_Bill_List Is Not Null Then
        For J In 1 .. j_Bill_List.Count Loop
          j_Item     := PLJson();
          j_Item     := PLJson(j_Bill_List.Get(J));
          n_������Դ := j_Item.Get_Number('fee_source');
          n_��¼���� := j_Item.Get_Number('fee_billtype');
          v_No       := j_Item.Get_String('fee_no');

          If n_������Դ = 1 Then
            l_Outnos.Extend;
            l_Outnos(l_Outnos.Count) := t_StrObj2(n_��¼����, v_No);
          Else
            l_Innos.Extend;
            l_Innos(l_Innos.Count) := t_StrObj2(n_��¼����, v_No);
          End If;
        End Loop;
      End If;

      --����ϵͳ�мǷ�ͬ���쳣������
      For r_Fee In (Select Distinct 1 As ������Դ, Decode(Mod(a.��¼����, 10), 2, 2, 1) As ��������, a.No
                    From ������ü�¼ A, ���˷����쳣��¼ B
                    Where a.����id = n_����id And a.�շ���� In ('5', '6', '7') And a.��¼���� In (1, 2) And
                          a.id = b.����id And (b.�������� = 0 Or b.�������� = 1) And Nvl(b.ͬ����־, 0) = 1 And
                          Exists (Select 1
                           From ������ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)
                    Union All
                    Select Distinct 2 As ������Դ, Decode(a.�ಡ�˵�, 1, 3, 2) As ��������, a.No
                    From סԺ���ü�¼ A, ���˷����쳣��¼ B
                    Where a.����id = n_����id And a.�շ���� In ('5', '6', '7') And a.��¼���� = 2 And
                          a.id = b.����id And b.�������� = 0 And Nvl(b.ͬ����־, 0) = 1 And Exists
                     (Select 1
                           From סԺ���ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)) Loop

        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.��������;
        v_Json := v_Json || ',"fee_source":' || r_Fee.������Դ;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';

        Json_Temp_Out := Null;
        Zl_Drugbill_Build(v_Json, Json_Temp_Out);

        --��������
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');

        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;

        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);

        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;

      --�ٴ�ϵͳ�мǷ�ͬ���쳣������
      For r_Fee In (Select /*+Cardinality(j,10)*/
                    Distinct 1 As ������Դ, Decode(Mod(a.��¼����, 10), 2, 2, 1) As ��������, a.No
                    From ������ü�¼ A, Table(l_Outnos) J
                    Where a.����id = n_����id And a.�շ���� In ('5', '6', '7') And a.��¼���� = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From ������ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)
                    Union All
                    Select /*+Cardinality(j,10)*/
                    Distinct 2 As ������Դ, Decode(a.�ಡ�˵�, 1, 3, 2) As ��������, a.No
                    From סԺ���ü�¼ A, Table(l_Innos) J
                    Where a.����id = n_����id And a.�շ���� In ('5', '6', '7') And a.��¼���� = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From סԺ���ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)) Loop

        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.��������;
        v_Json := v_Json || ',"fee_source":' || r_Fee.������Դ;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';

        Json_Temp_Out := Null;
        Zl_Drugbill_Build(v_Json, Json_Temp_Out);

        --��������
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');

        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;

        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);

        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;

    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�","pati_bill_list":[' || c_Jtmp || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getdrugerrdata;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Getstufferrdata
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ����ݲ���ID��ҽ����Ϣ���ز��˷�����Ϣ
  --��Σ�Json_In:��ʽ����ֵΪnullʱ��ȡ���в��˵��쳣��Ϣ
  --  input
  --    pati_list[]�����б�
  --       pati_id                    N 1 ����id
  --       bill_list[]                ���õ��ݺ��б����Բ���������ʱ��ʾ��ȡ������ͬ���쳣������
  --         fee_source               N 0 ������Դ��1-���2-סԺ
  --         fee_billtype             N 0 ���õ������ͣ�1-�շѴ�����2-���ʵ�����
  --         fee_no                   C 0 ���õ��ݺ�
  --����: Json_Out,��ʽ����
  --  output
  --    code                          N   1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C   1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    pati_bill_list[]
  --       billtype                   N   1 ��������: 1 -�շѴ���  ;2- ���ʵ�����;3- ���ʱ���
  --       pati_source                N   1 ������Դ:1-����;2-סԺ;4-���
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�--------------------------------------
  --       pati_id                    N   1 ����ID
  --       pati_pageid                N   1 ��ҳID
  --       pati_name                  C   1 ��������
  --       pati_sex_code              C   1 �Ա��ţ�������)
  --       pati_sex                   C   1 �Ա�
  --       pati_age                   C   1 ����
  --       pati_deptid                N   1 ���˿���ID
  --       pati_wardarea_id           N     ���˲���ID
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------
  --       bill_list[]                      ���������б�[����]
  --         fee_source                N  0 ������Դ
  --         stuff_no                  C  1 NO
  --         charge_tag                N  1 �շѱ�־:0-δ�շѻ���ʻ���;1-���շѻ����
  --         fee_acnter                C  0 ������
  --         plcdept_id                C  0 ��������id��������)
  --         plcdept                   C  0 �����������ƣ�������)
  --         placer_id                 C  0 ����ҽʦid��������)
  --         placer                    C  0 ����ҽʦ��������) ����
  --         operator_name             C  1 ����Ա����
  --         operator_code             C  1 ����Ա���
  --         create_time               C  1 �Ǽ�ʱ��:yyyy-mm-dd hh:mi:ss
  --         item_list[]                    ���������б�[����]
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������½ڵ�----------------------------------------
  --           pati_id                 N  1 ����ID
  --           pati_pageid             N  0 ��ҳID
  --           pati_name               C  1 ��������
  --           pati_sex_code           C  1 �Ա��ţ�������)
  --           pati_sex                C  1 �Ա�
  --           pati_age                C  1 ����
  --           pati_wardarea_id        N    ���˲���ID
  --           pati_deptid             N  1 ���˿���ID
  ---------------------------billtype = 3,���ʱ�����ҩƷ�շ���¼.����=26��ʱ�������Ͻڵ�-----------------------------------------

  --           stuffdtl_id             N  1 ������ϸID(Ŀǰ������Ƿ���id)
  --           serial_num              N  1 ���:(���(�����洢)����ź���ţ�1��2��3��3��3��4��)
  --           warehouse_id            N  1 �ⷿID
  --           is_bakstuff             N  1 �Ƿ񱸻�����:�и�ֵ���Ĳ���Ҫ���룬��0��ʾ�Ǹ�ֵ����ģʽ(��ɨ��ʱʹ��)
  --           bakstuff_batch          N  1 ������������
  --           stuff_id                N  1 ����ID
  --           baby_num                N  0 Ӥ�����
  --           advice_id               N  0 ҽ��ID
  --           packages_num            N  1 ����
  --           outbound_num            N  1 ��������
  --           price                   N  0 �ۼ�
  --           money                   N  0 ���۽��(������)
  --           memo                    C  0 ժҪ
  ------------------------------------------------------------------------------------------------------------
  j_Json      PLJson;
  j_Json_In   PLJson;
  j_Pati_List Pljson_List;
  j_Json_Out  PLJson;
  j_Bill_List Pljson_List;

  Json_Temp_Out Clob;
  c_Jtmp        Clob;

  j_Item   PLJson;
  n_����id Number(18);

  v_Json Varchar2(4000);
  n_Code Number;

  n_������Դ Number(1);
  n_��¼���� ������ü�¼.��¼����%Type;
  v_No       ������ü�¼.No%Type;

  l_Outnos t_StrList2 := t_StrList2();
  l_Innos  t_StrList2 := t_StrList2();
Begin
  If Json_In Is Null Then
    --����ϵͳ�мǷ�ͬ���쳣������
    For r_Fee In (Select Distinct 1 As ������Դ, Decode(Mod(a.��¼����, 10), 2, 2, 1) As ��������, a.No
                  From ������ü�¼ A, ���˷����쳣��¼ B
                  Where a.����id = n_����id And a.�շ���� = '4' And a.��¼���� In (1, 2) And 
                        a.id = b.����id And (b.�������� = 0 Or b.�������� = 1) And Nvl(b.ͬ����־, 0) = 1 And Exists
                   (Select 1
                         From ������ü�¼
                         Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                         Group By ��¼����, NO, ���
                         Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)
                  Union All
                  Select Distinct 2 As ������Դ, Decode(a.�ಡ�˵�, 1, 3, 2) As ��������, a.No
                  From סԺ���ü�¼ A, ���˷����쳣��¼ B
                  Where a.����id = n_����id And a.�շ���� = '4' And a.��¼���� = 2 And 
                        a.id = b.����id And b.�������� = 0 And Nvl(b.ͬ����־, 0) = 1 And Exists
                   (Select 1
                         From סԺ���ü�¼
                         Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                         Group By ��¼����, NO, ���
                         Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)) Loop
    
      v_Json := Null;
      v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
      v_Json := v_Json || ',"billtype":' || r_Fee.��������;
      v_Json := v_Json || ',"fee_source":' || r_Fee.������Դ;
      v_Json := v_Json || '}';
      v_Json := '{"input":' || v_Json || '}';
    
      Json_Temp_Out := Null;
      Zl_Stuffbill_Build(v_Json, Json_Temp_Out);
    
      --��������
      j_Json_Out := PLJson();
      j_Json_Out := PLJson(Json_Temp_Out);
      j_Json     := PLJson();
      j_Json     := j_Json_Out.Get_Pljson('output');
    
      n_Code := Nvl(j_Json.Get_Number('code'), '0');
      If n_Code = 0 Then
        Json_Out := zlJsonOut(j_Json.Get_String('message'));
        Return;
      End If;
    
      j_Json.Remove('code');
      j_Json.Remove('message');
      Json_Temp_Out := Empty_Clob();
      Dbms_Lob.Createtemporary(Json_Temp_Out, True);
      j_Json.To_Clob(Json_Temp_Out);
    
      If c_Jtmp Is Null Then
        c_Jtmp := Json_Temp_Out;
      Else
        c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
      End If;
    End Loop;
  Else
    --�������
    j_Json_In   := PLJson(Json_In);
    j_Json      := j_Json_In.Get_Pljson('input');
    j_Pati_List := j_Json.Get_Pljson_List('pati_list');
    
    For I In 1 .. j_Pati_List.Count Loop
      j_Item      := PLJson();
      j_Item      := PLJson(j_Pati_List.Get(I));
      n_����id    := j_Item.Get_Number('pati_id');
      j_Bill_List := j_Item.Get_Pljson_List('bill_list');
    
      If j_Bill_List Is Not Null Then
        For J In 1 .. j_Bill_List.Count Loop
          j_Item     := PLJson();
          j_Item     := PLJson(j_Bill_List.Get(J));
          n_������Դ := j_Item.Get_Number('fee_source');
          n_��¼���� := j_Item.Get_Number('fee_billtype');
          v_No       := j_Item.Get_String('fee_no');
        
          If n_������Դ = 1 Then
            l_Outnos.Extend;
            l_Outnos(l_Outnos.Count) := t_StrObj2(n_��¼����, v_No);
          Else
            l_Innos.Extend;
            l_Innos(l_Innos.Count) := t_StrObj2(n_��¼����, v_No);
          End If;
        End Loop;
      End If;
    
      --����ϵͳ�мǷ�ͬ���쳣������
      For r_Fee In (Select Distinct 1 As ������Դ, Decode(Mod(a.��¼����, 10), 2, 2, 1) As ��������, a.No
                    From ������ü�¼ A, ���˷����쳣��¼ B
                    Where a.����id = n_����id And a.�շ���� = '4' And a.��¼���� In (1, 2) And 
                          a.id = b.����id And (b.�������� = 0 Or b.�������� = 1) And Nvl(b.ͬ����־, 0) = 1 And Exists
                     (Select 1
                           From ������ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)
                    Union All
                    Select Distinct 2 As ������Դ, Decode(a.�ಡ�˵�, 1, 3, 2) As ��������, a.No
                    From סԺ���ü�¼ A, ���˷����쳣��¼ B
                    Where a.����id = n_����id And a.�շ���� = '4' And a.��¼���� = 2 And 
                          a.id = b.����id And b.�������� = 0 And Nvl(b.ͬ����־, 0) = 1 And Exists
                     (Select 1
                           From סԺ���ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)) Loop
      
        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.��������;
        v_Json := v_Json || ',"fee_source":' || r_Fee.������Դ;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';
      
        Json_Temp_Out := Null;
        Zl_Stuffbill_Build(v_Json, Json_Temp_Out);
      
        --��������
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');
      
        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;
      
        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);
      
        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;
    
      --�ٴ�ϵͳ�мǷ�ͬ���쳣������ 
      For r_Fee In (Select /*+Cardinality(j,10)*/
                    Distinct 1 As ������Դ, Decode(Mod(a.��¼����, 10), 2, 2, 1) As ��������, a.No
                    From ������ü�¼ A, Table(l_Outnos) J
                    Where a.����id = n_����id And a.�շ���� = '4' And a.��¼���� = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From ������ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)
                    Union All
                    Select /*+Cardinality(j,10)*/
                    Distinct 2 As ������Դ, Decode(a.�ಡ�˵�, 1, 3, 2) As ��������, a.No
                    From סԺ���ü�¼ A, Table(l_Innos) J
                    Where a.����id = n_����id And a.�շ���� = '4' And a.��¼���� = j.C1 And a.No = j.C2 And Exists
                     (Select 1
                           From סԺ���ü�¼
                           Where ��¼���� = a.��¼���� And NO = a.No And ��� = a.���
                           Group By ��¼����, NO, ���
                           Having Nvl(Sum(Nvl(����, 1) * ����), 0) <> 0)) Loop
      
        v_Json := Null;
        v_Json := v_Json || '{"fee_no":"' || r_Fee.No || '"';
        v_Json := v_Json || ',"billtype":' || r_Fee.��������;
        v_Json := v_Json || ',"fee_source":' || r_Fee.������Դ;
        v_Json := v_Json || '}';
        v_Json := '{"input":' || v_Json || '}';
      
        Json_Temp_Out := Null;
        Zl_Stuffbill_Build(v_Json, Json_Temp_Out);
      
        --��������
        j_Json_Out := PLJson();
        j_Json_Out := PLJson(Json_Temp_Out);
        j_Json     := PLJson();
        j_Json     := j_Json_Out.Get_Pljson('output');
      
        n_Code := Nvl(j_Json.Get_Number('code'), '0');
        If n_Code = 0 Then
          Json_Out := zlJsonOut(j_Json.Get_String('message'));
          Return;
        End If;
      
        j_Json.Remove('code');
        j_Json.Remove('message');
        Json_Temp_Out := Empty_Clob();
        Dbms_Lob.Createtemporary(Json_Temp_Out, True);
        j_Json.To_Clob(Json_Temp_Out);
      
        If c_Jtmp Is Null Then
          c_Jtmp := Json_Temp_Out;
        Else
          c_Jtmp := c_Jtmp || ',' || Json_Temp_Out;
        End If;
      End Loop;
    End Loop;
  End If;
  Json_Out := '{"output":{"code":1,"message": "�ɹ�","pati_bill_list":[' || c_Jtmp || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getstufferrdata;
/

Create Or Replace Procedure Zl_Exsesvr_Getorderfeestate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ҽ��id,��ȡ��ص��շ�״̬
  --��Σ�Json_In:��ʽ
  --input     
  --  advice_ids  C 1 ҽ��id
  --  bill_nos  C 1 ���ݺ�
  --����: Json_Out,��ʽ����
  -- output      
  --   code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --   state N 1 ״̬:0-δ�շ�,1-��ȫ�շ�;2-�����շ�
  --   billtype  N 1 ��������:0-�������κε���;1-�շѵ�;2-���ʵ�;3-�շѺͼ��ʶ���
  --   advice_ids  C   δ�շѵ�ҽ��ID:����ҽ��idsʱ��Ч

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output Varchar2(32767);

  v_ҽ��ids Varchar2(32767);

  v_���ݺ�      Varchar2(32767);
  n_״̬        Number(2);
  v_δ��ҽ��ids Varchar2(32767);
  n_��������    Number(2);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_ҽ��ids := j_Json.Get_String('advice_ids');
  v_���ݺ�  := j_Json.Get_String('bill_nos');

  n_״̬        := -1;
  v_δ��ҽ��ids := '';
  n_��������    := 0;
  If v_ҽ��ids Is Not Null Then
  
    For c_ҽ�� In (
                 
                 Select /*+ RULE */
                 Distinct ��¼����, ��¼״̬, ҽ�����
                 From (With ҽ������ As (Select Column_Value As ҽ��id From Table(f_Num2List(v_ҽ��ids)))
                         Select Distinct a.��¼����, a.��¼״̬, a.ҽ�����
                         From ������ü�¼ A, ҽ������ B
                         Where a.ҽ����� = b.ҽ��id And a.��¼���� In (1, 2, 3) And a.��¼״̬ In (0, 1, 3)
                         Union All
                         Select Distinct a.��¼����, a.��¼״̬, a.ҽ�����
                         From סԺ���ü�¼ A, ҽ������ B
                         Where a.ҽ����� = b.ҽ��id And a.��¼���� In (1, 2, 3) And a.��¼״̬ In (0, 1, 3))
                 ) Loop
    
      If c_ҽ��.��¼״̬ = 0 Then
        --δ�շ�
        If Nvl(c_ҽ��.ҽ�����, 0) <> 0 Then
          v_δ��ҽ��ids := Nvl(v_δ��ҽ��ids, '') || ',' || Nvl(c_ҽ��.ҽ�����, 0);
        
        End If;
      End If;
    
      If n_״̬ = -1 Then
        If c_ҽ��.��¼״̬ = 0 Then
          n_״̬ := Case
                    When c_ҽ��.��¼״̬ = 0 Then
                     0
                    Else
                     1
                  End;
        End If;
      Elsif n_״̬ = 0 And (c_ҽ��.��¼״̬ = 1 Or c_ҽ��.��¼״̬ = 3) Then
        n_״̬ := 2; --   �����շ�
      Elsif n_״̬ = 1 And c_ҽ��.��¼״̬ = 0 Then
        n_״̬ := 2; --�����շ�
      End If;
    
      If n_�������� = 0 Then
        n_�������� := c_ҽ��.��¼����;
      Elsif n_�������� <> c_ҽ��.��¼���� Then
        --��������
        n_�������� := 3;
      End If;
    
    End Loop;
  
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '�ɹ�');
    zlJsonPutValue(v_Output, 'state', Nvl(n_״̬, 0), 1);
    zlJsonPutValue(v_Output, 'billtype', Nvl(n_��������, 0), 1);
    zlJsonPutValue(v_Output, 'advice_ids', Nvl(v_δ��ҽ��ids, ''), 0, 2);
    Json_Out := '{"output":' || v_Output || '}';
    Return;
  
  End If;

  For c_ҽ�� In (Select /*+ RULE */
               Distinct ��¼����, ��¼״̬, ҽ�����
               From (With ҽ������ As (Select Column_Value As NO From Table(f_Str2List(v_���ݺ�)))
                      Select Distinct a.��¼����, a.��¼״̬, a.ҽ�����
                      From ������ü�¼ A, ҽ������ B
                      Where a.No = b.No And a.��¼���� In (1, 2, 3) And a.��¼״̬ In (0, 1, 3)
                      Union All
                      Select Distinct a.��¼����, a.��¼״̬, a.ҽ�����
                      From סԺ���ü�¼ A, ҽ������ B
                      Where a.No = b.No And a.��¼���� In (1, 2, 3) And a.��¼״̬ In (0, 1, 3))
               ) Loop
  
    If c_ҽ��.��¼״̬ = 0 Then
      --δ�շ�
      If Nvl(c_ҽ��.ҽ�����, 0) <> 0 Then
        v_δ��ҽ��ids := Nvl(v_δ��ҽ��ids, '') || ',' || Nvl(c_ҽ��.ҽ�����, 0);
      End If;
    End If;
  
    If n_״̬ = -1 Then
      If c_ҽ��.��¼״̬ = 0 Then
        n_״̬ := Case
                  When c_ҽ��.��¼״̬ = 0 Then
                   0
                  Else
                   1
                End;
      End If;
    Elsif n_״̬ = 0 And (c_ҽ��.��¼״̬ = 1 Or c_ҽ��.��¼״̬ = 3) Then
      n_״̬ := 2; --   �����շ�
    Elsif n_״̬ = 1 And c_ҽ��.��¼״̬ = 0 Then
      n_״̬ := 2; --�����շ�
    End If;
  
    If n_�������� = 0 Then
      n_�������� := c_ҽ��.��¼����;
    Elsif n_�������� <> c_ҽ��.��¼���� Then
      --��������
      n_�������� := 3;
    End If;
  End Loop;

  --    state  N  1  ״̬:0-δ�շ�,1-��ȫ�շ�;2-�����շ�
  --    billtype  N  1  ��������:0-�������κε���;1-�շѵ�;2-���ʵ�;3-�շѺͼ��ʶ���
  --    advice_ids  C    δ�շѵ�ҽ��ID:����ҽ��idsʱ��Ч

  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'state', Nvl(n_״̬, 0), 1);
  zlJsonPutValue(v_Output, 'billtype', Nvl(n_��������, 0), 1);
  zlJsonPutValue(v_Output, 'advice_ids', Nvl(v_δ��ҽ��ids, ''), 0, 2);
  Json_Out := '{"output":' || v_Output || '}';

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getorderfeestate;
/


Create Or Replace Procedure Zl_Exsesvr_Getfeechargestate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݵ��ݺ���Ϣ����ȡ���ݶ�Ӧ���շ�״̬
  --��Σ�Json_In:��ʽ
  --input     
  --    query_mode  N 1 ��ѯ��ʽ:0-��ѯ�շ�״̬;1-�����Ƿ����δ�շѵ�
  --    bill_nos  C 1 ���ݺ�
  --    bill_type N 1 ��������:1-�շѵ���;���Ժ���չ

  --����: Json_Out,��ʽ����
  -- output      
  --   code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --   message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --   state N 1 ״̬
  --       1.query_mode=0ʱ
  --         ״̬:0-δ�շ�;1-�����շ�/�˷�;2-ȫ���շ�;3-ȫ���˷�
  --       2.query_mode=1ʱ
  --         ״̬:1-����δ�շѵ�;0-������δ�շ�.
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output     Varchar2(32767);
  n_��ѯ��ʽ   Number(2);
  v_���ݺ�     Varchar2(32767);
  n_��������   Number(2);
  n_״̬       Number(2);
  n_�Ƿ�ȫ��   Number(2);
  n_�Ƿ�ȫ��   Number(2);
  n_�Ƿ񲿷��� Number(2);
  n_�Ƿ�δ��   Number(2);

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ��ʽ := Nvl(j_Json.Get_Number('query_mode'), 0);
  v_���ݺ�   := Nvl(j_Json.Get_String('bill_nos'), '');
  n_�������� := Nvl(j_Json.Get_Number('bill_type'), 1);

  If Nvl(n_��������, 0) <> 1 Then
    Json_Out := zlJsonOut('�ݲ�֧�ַ��շѵ��ݡ�');
    Return;
  End If;

  If n_��ѯ��ʽ = 1 Then
    Begin
      --�жϷ���״̬����Ҫ���쳣�ģ�������Ҫ����
      If Instr(v_���ݺ�, ',') > 0 Then
        Select /*+cardinality(b,10)*/
         1
        Into n_״̬
        From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2List(v_���ݺ�))) B
        Where a.��¼���� = 1 And (a.����id Is Null Or Nvl(a.����״̬, 0) = 1 And a.��¼״̬ = 1) And a.No = b.No And Rownum < 2;
      Else
        Select 1
        Into n_״̬
        From ������ü�¼ A
        Where a.��¼���� = 1 And (a.����id Is Null Or Nvl(a.����״̬, 0) = 1 And a.��¼״̬ = 1) And a.No = v_���ݺ� And Rownum < 2;
      End If;
    Exception
      When Others Then
        n_״̬ := 0;
    End;
    zlJsonPutValue(v_Output, 'code', 1, 1, 1);
    zlJsonPutValue(v_Output, 'message', '�ɹ�');
    zlJsonPutValue(v_Output, 'state', Nvl(n_״̬, 0), 1, 2);
    Json_Out := '{"output":' || v_Output || '}';
    Return;
  End If;

  n_״̬       := -1;
  n_�Ƿ�ȫ��   := -1;
  n_�Ƿ񲿷��� := -1;
  n_�Ƿ�ȫ��   := -1;
  n_�Ƿ�δ��   := -1;
  For c_���� In (Select /*+cardinality(b,10)*/
                a.No, a.���, Nvl(Sum(a.���� * Nvl(a.����, 1)), 0) As ʣ������,
                Nvl(Sum(Decode(a.��¼����, 1, 1, 0) * Decode(a.��¼״̬, 2, 0, 1) * a.���� * Nvl(a.����, 1)), 0) As ԭʼ����,
                Nvl(Sum(Decode(a.��¼����, 1, 1, 0) *
                         Decode(a.����id, Null, 1, Decode(a.��¼״̬, 0, 1, 1, Decode(Nvl(a.����״̬, 0), 1, 1, 0), 0)) * a.���� *
                         Nvl(a.����, 1)), 0) As δ������
               From ������ü�¼ A, (Select Column_Value As NO From Table(f_Str2List(v_���ݺ�))) B
               Where Mod(a.��¼����, 10) = 1 And a.�۸񸸺� Is Null And a.No = b.No
               Group By a.No, a.���) Loop
  
    If c_����.ԭʼ���� <> 0 And c_����.ԭʼ���� = c_����.δ������ Then
      --δ�շ�
      n_�Ƿ�δ�� := 1;
    Elsif c_����.ԭʼ���� = c_����.ʣ������ And c_����.δ������ = 0 Then
      --ȫ���շ�
      n_�Ƿ�ȫ�� := 1;
    Elsif c_����.ʣ������ = 0 Then
      --ȫ����
      n_�Ƿ�ȫ�� := 1;
    Else
      --�����շѻ��˷� 
      n_�Ƿ񲿷��� := 1;
      Exit;
    End If;
    If n_�Ƿ�δ�� <> -1 And n_�Ƿ�ȫ�� <> -1 And n_�Ƿ񲿷��� <> -1 Then
      Exit;
    End If;
  End Loop;
  --1-�����ڵ���,0-δ�շ�;1-�����շѻ��˷�;2-ȫ���շ�;3-ȫ���˷�
  If n_�Ƿ񲿷��� = 1 Then
    n_״̬ := 1;
  Elsif n_�Ƿ�ȫ�� = -1 And n_�Ƿ�ȫ�� = 1 And n_�Ƿ�δ�� = -1 Then
    --ȫ��
    n_״̬ := 3;
  Elsif n_�Ƿ�ȫ�� = 1 And n_�Ƿ�ȫ�� = -1 And n_�Ƿ�δ�� = -1 Then
    --ȫ��
    n_״̬ := 2;
  Elsif n_�Ƿ�ȫ�� = -1 And n_�Ƿ�ȫ�� = -1 And n_�Ƿ�δ�� = 1 Then
    n_״̬ := 0;
  Elsif n_�Ƿ�ȫ�� = -1 And n_�Ƿ�ȫ�� = -1 And n_�Ƿ�δ�� = -1 Then
    n_״̬ := -1;
  Else
    n_״̬ := 1; --�����ջ���
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'state', Nvl(n_״̬, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeechargestate;
/

Create Or Replace Procedure Zl_Exsesvr_Getfeebalancestate
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݵ��ݺ���Ϣ����ȡ���ݶ�Ӧ�Ľ���״̬
  --��Σ�Json_In:��ʽ
  --input     
  --    query_mode  N 1 ��ѯ��ʽ:0-�������;1-סԺ����
  --    bill_nos  C 1 ���ݺ�
  --����: Json_Out,��ʽ����
  -- output      
  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    state N 1 ״̬:-1-�����ڼ��ʵ���;0-δ����;1-���ֽ���;2-ȫ������

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output   Varchar2(32767);
  n_��ѯ��ʽ Number(2);
  v_���ݺ�   Varchar2(32767);
  n_���ʱ�־ Number(18);
  n_����     Number(18);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ��ʽ := Nvl(j_Json.Get_Number('query_mode'), 0);
  v_���ݺ�   := Nvl(j_Json.Get_String('bill_nos'), '');

  If Nvl(n_��ѯ��ʽ, 0) = 0 Then
    If Instr(v_���ݺ�, ',') > 0 Then
      Select Decode(Nvl(Sum(Nvl((Case
                                   When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And ͳ���� = 0) Then
                                    0
                                   When δ���� <> 0 And ���ʽ�� <> 0 Then
                                    1
                                   Else
                                    2
                                 End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(����)
      Into n_���ʱ�־, n_����
      From (Select /*+Cardinality(B,10)*/
              a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��,
              Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����,
              Mod(Sum(Decode(Nvl(a.����id, 0), 0, 0, 1)), 2) As ͳ����, Max(1) As ����
             From ������ü�¼ A, Table(f_Str2List(v_���ݺ�)) B
             Where a.No = b.Column_Value And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2
             Group By a.No, Nvl(a.�۸񸸺�, a.���));
    
    Else
    
      Select Decode(Nvl(Sum(Nvl((Case
                                  When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And ͳ���� = 0) Then
                                   0
                                  When δ���� <> 0 And ���ʽ�� <> 0 Then
                                   1
                                  Else
                                   2
                                End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(����)
      Into n_���ʱ�־, n_����
      From (Select a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��,
                    Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��, Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��,
                    Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����,
                    Mod(Sum(Decode(Nvl(a.����id, 0), 0, 0, 1)), 2) As ͳ����, Max(1) As ����
             From ������ü�¼ A
             Where a.No = v_���ݺ� And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2
             Group By a.No, Nvl(a.�۸񸸺�, a.���));
    End If;
  Else
    If Instr(v_���ݺ�, ',') > 0 Then
      Select Decode(Nvl(Sum(Nvl((Case
                                   When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And ͳ���� = 0) Then
                                    0
                                   When δ���� <> 0 And ���ʽ�� <> 0 Then
                                    1
                                   Else
                                    2
                                 End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(����)
      Into n_���ʱ�־, n_����
      From (Select /*+Cardinality(B,10)*/
              a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��,
              Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��, Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����,
              Mod(Sum(Decode(Nvl(a.����id, 0), 0, 0, 1)), 2) As ͳ����, Max(1) As ����
             From סԺ���ü�¼ A, Table(f_Str2List(v_���ݺ�)) B
             Where a.No = b.Column_Value And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2
             Group By a.No, Nvl(a.�۸񸸺�, a.���));
    
    Else
    
      Select Decode(Nvl(Sum(Nvl((Case
                                  When (δ���� <> 0 And ���ʽ�� = 0) Or (δ���� = 0 And (ʵ�ս�� = 0 Or ���ʽ�� = 0) And ͳ���� = 0) Then
                                   0
                                  When δ���� <> 0 And ���ʽ�� <> 0 Then
                                   1
                                  Else
                                   2
                                End), 0)), 0), 0, 0, 2 * Count(1), 2, 1), Max(����)
      Into n_���ʱ�־, n_����
      From (Select a.No, Nvl(a.�۸񸸺�, a.���) As ���, Nvl(Sum(Nvl(a.Ӧ�ս��, 0)), 0) As Ӧ�ս��,
                    Nvl(Sum(Nvl(a.ʵ�ս��, 0)), 0) As ʵ�ս��, Nvl(Sum(Nvl(a.���ʽ��, 0)), 0) As ���ʽ��,
                    Nvl(Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)), 0) As δ����,
                    Mod(Sum(Decode(Nvl(a.����id, 0), 0, 0, 1)), 2) As ͳ����, Max(1) As ����
             From סԺ���ü�¼ A
             Where a.No = v_���ݺ� And a.���ʷ��� = 1 And Mod(a.��¼����, 10) = 2
             Group By a.No, Nvl(a.�۸񸸺�, a.���));
    End If;
  End If;
  If Nvl(n_����, 0) = 0 Then
    n_���ʱ�־ := -1;
  End If;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'state', Nvl(n_���ʱ�־, 0), 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeebalancestate;
/


Create Or Replace Procedure Zl_Exsesvr_Getfeeinfobyblncid
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --����:���ݽ���id��ȡ��Ӧ�ķ�����ϸ����
  --��Σ�Json_In:��ʽ
  --input      
  -- balance_id  N 1 ����ID

  --����: Json_Out,��ʽ����
  --output     
  -- code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  -- message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -- fee_list[]  C 1 ������ϸ�б�
  --   fee_no  C 1 ���õ��ݺ�
  --   serial_num  N 1 ���
  --   receipt_type  C 1 �շ����
  --   fitem_id  N 1 �շ�ϸĿid
  --   fitem_name  C 1 �շ�����
  --   nums  N 1 ����
  --   unit  C 1 ���㵥λ
  --   price N 1 ��׼����
  --   blnc_money  N 1 ���ʽ��
  --   exedept_name  C 1 ִ�п���
  --   happen_time     C 1 ����ʱ��:yyyy-mm-dd HH:MM:SS

  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id Number(18);
  v_Output Varchar2(32767);
  c_Output Clob;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := Nvl(j_Json.Get_Number('balance_id'), 0);

  For c_���� In (Select a.No, a.���, a.�շ����, Nvl(e.����, d.����) As �շ�����, a.���� As �շ�����, a.���ʽ��, a.�շѵ���, a.���㵥λ,
                      Nvl(b.����, 'δ֪') As ִ�п���, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, a.�շ�ϸĿid
               From (
                      
                      Select a.����ʱ��, a.No, Nvl(�۸񸸺�, ���) As ���, a.�շ����, a.�շ�ϸĿid, Avg(Nvl(����, 1)) * Avg(����) As ����, a.���㵥λ,
                              Sum(a.���ʽ��) As ���ʽ��, Sum(a.��׼����) As �շѵ���, a.ִ�в���id
                      From ������ü�¼ A
                      Where a.����id = n_����id
                      Group By a.����ʱ��, a.No, Nvl(�۸񸸺�, ���), a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.ִ�в���id
                      Union All
                      Select a.����ʱ��, a.No, Nvl(�۸񸸺�, ���) As ���, a.�շ����, a.�շ�ϸĿid, Avg(Nvl(����, 1)) * Avg(����) As ����, a.���㵥λ,
                              Sum(a.���ʽ��) As ���ʽ��, Sum(a.��׼����) As �շѵ���, a.ִ�в���id
                      From סԺ���ü�¼ A
                      Where a.����id = n_����id
                      Group By a.����ʱ��, a.No, Nvl(�۸񸸺�, ���), a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.ִ�в���id) A, ���ű� B, �շ���ĿĿ¼ D,
                    �շ���Ŀ���� E
               Where a.ִ�в���id = b.Id(+) And a.�շ�ϸĿid = d.Id And a.�շ�ϸĿid = e.�շ�ϸĿid(+) And e.����(+) = 1 And e.����(+) = 3
               Order By ����ʱ�� Desc, NO Desc, ���) Loop
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
    
      v_Output := Null;
    End If;
  
    zlJsonPutValue(v_Output, 'fee_no', c_����.No, 0, 1);
    zlJsonPutValue(v_Output, 'serial_num', Nvl(c_����.���, 1), 1);
    zlJsonPutValue(v_Output, 'receipt_type', Nvl(c_����.�շ����, ''));
    zlJsonPutValue(v_Output, 'fitem_id', Nvl(c_����.�շ�ϸĿid, 0), 1);
    zlJsonPutValue(v_Output, 'fitem_name', Nvl(c_����.�շ�����, ''));
    zlJsonPutValue(v_Output, 'nums', Nvl(c_����.�շ�����, 1), 1);
    zlJsonPutValue(v_Output, 'unit', Nvl(c_����.���㵥λ, ''));
    zlJsonPutValue(v_Output, 'price', Nvl(c_����.�շѵ���, 0), 1);
    zlJsonPutValue(v_Output, 'blnc_money', Nvl(c_����.���ʽ��, 0), 1);
    zlJsonPutValue(v_Output, 'exedept_name', Nvl(c_����.ִ�п���, ''));
    zlJsonPutValue(v_Output, 'happen_time', Nvl(c_����.����ʱ��, ''), 0, 2);
  
  End Loop;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","fee_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_list":[' || v_Output || ']}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getfeeinfobyblncid;
/


Create Or Replace Procedure Zl_Exsesvr_Getbalanceinfobyid
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:���ݽ���id��ȡ��Ӧ�Ľ�����ϸ����
  --��Σ�Json_In:��ʽ
  --input      
  -- balance_id  N 1 ����ID
  --����: Json_Out,��ʽ����
  --  output      
  --    code  C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    blnc_list[] C 1 ������ϸ�б�
  --      blnc_mode C 1 ���㷽ʽ
  --      blnc_no C 1 �������
  --      blnc_money  N 1 ������
  --      cardtype_id N 1 �����id
  --      consumer_no N 1 ���㿨��ţ��������ѽӿ�Ŀ¼.���
  --      cardno  C 1 ����
  --      swapno  C 1 ������ˮ��
  --      swapmemo  C 1 ����˵��
  --      memo  C 1 ժҪ
  --      cprtion_unit  C 1 ������λ
  --      relation_id N 1 ��������id
  --
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id Number(18);
  v_Output Varchar2(32767);
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := Nvl(j_Json.Get_Number('balance_id'), 0);

  For c_���� In (
               
               Select Decode(Mod(a.��¼����, 10), 1, '[��Ԥ��]', a.���㷽ʽ) As ���㷽ʽ, ��Ԥ�� As ������, a.�������, a.ժҪ, a.�����id, a.���㿨���,
                       a.������ˮ��, a.����˵��, a.����, a.��������id, ������λ
               From ����Ԥ����¼ A
               Where a.����id = n_����id) Loop
  
    zlJsonPutValue(v_Output, 'blnc_mode', c_����.���㷽ʽ, 0, 1);
    zlJsonPutValue(v_Output, 'blnc_no', Nvl(c_����.�������, ''));
    zlJsonPutValue(v_Output, 'blnc_money', Nvl(c_����.������, 0), 1);
    zlJsonPutValue(v_Output, 'cardtype_id', Nvl(c_����.�����id, 0), 1);
    zlJsonPutValue(v_Output, 'consumer_no', Nvl(c_����.���㿨���, 0), 1);
    zlJsonPutValue(v_Output, 'cardno', Nvl(c_����.����, ''));
    zlJsonPutValue(v_Output, 'swapno', Nvl(c_����.������ˮ��, ''));
    zlJsonPutValue(v_Output, 'swapmemo', Nvl(c_����.����˵��, ''));
    zlJsonPutValue(v_Output, 'memo', Nvl(c_����.ժҪ, ''));
    zlJsonPutValue(v_Output, 'cprtion_unit', Nvl(c_����.������λ, ''));
    zlJsonPutValue(v_Output, 'relation_id', Nvl(c_����.��������id, 0), 1, 2);
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","blnc_list":[' || v_Output || ']}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbalanceinfobyid;
/

Create Or Replace Procedure Zl_Exsesvr_Getbalanceinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ------------------------------------------------------------------------------------------------- 
  --���ܣ���ʱ�䷶Χ��ȡ���õ��� 
  --��Σ�json��ʽ 
  --  input  
  --  query_type  N  1  ��ѯ��Χ:0-����ʣ����;1-������ԭʼ������Ϣ
  --    occasion  N  1  ���㳡��:1-�շ�,2-Ԥ��(����Ѻ��),3-����(������),4-�Һ�,5-���￨,6-����ҽ������
  --    fee_nos  C    query_type=2ʱ��Ч:���ݺ�:���㳡��=2ʱ��ΪԤ��NO, ����idδ���룬�ýڵ�ش�
  --���Σ�json��ʽ 
  --  output 
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    data  C    ������Ϣ
  --      pati_info  C    ������Ϣ
  --        pati_id  N  1  ����ID
  --        pati_pageid  N    ��ҳID
  --        pati_name  C  1  ����
  --        pati_sex  C  1  �Ա�
  --        pati_age  C  1  ����
  --        outpatient_num  C  1  �����
  --        inpatient_num  C  1  סԺ��
  --        insurance_type  N  1  ����
  --      balance_info  C    ������Ϣ
  --        invoice_no  C  1  ��Ʊ��
  --        balance_oldid  N  1  ԭ����ID
  --        create_time  C  1  �շ�ʱ��:yyyy-mm-dd hh:mi:ss
  --        total  N  1  �����ܶ�
  --        balance_unit  N  1  �Ƿ��Լ��λ����
  --        balance_type  N  1  Ԥ��ʱ��Ԥ�����:1-����;2-סԺ ;3-�����סԺ;����ʱ����������:1-����;2-סԺ ;3-�����סԺ;
  --        start_einv  N  1  �Ƿ����õ���Ʊ��
  ------------------------------------------------------------------------------------------------- 
  j_Input PLJson;
  j_Json  PLJson;

  n_����     Number(2);
  n_��ѯ���� Number(2);
  v_Nos      Varchar2(32767);
  v_Output   Varchar2(32767);

  n_����id       ������ü�¼.����id%Type;
  n_���ʽ��     ������ü�¼.���ʽ��%Type;
  n_�Ƿ����Ʊ�� ����Ԥ����¼.�Ƿ����Ʊ��%Type;
  v_�տ�ʱ��     Varchar2(30);

  Cursor c_������Ϣ Is(
    Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, a.����id, a.��ҳid, Sum(a.���) As ���ʽ��, Max(a.Ԥ������Ʊ��) As �Ƿ����Ʊ��,
           Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.סԺ��) As סԺ��, Max(a.�����) As �����, Max(m.����) As ����,
           To_Char(Max(Decode(a.��¼״̬, 2, To_Date(Null), a.�տ�ʱ��)), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ��, Max(a.Ԥ�����) As ��������
    From ����Ԥ����¼ A, ���ս����¼ M
    Where a.No = '-' And a.��¼���� = 1 And a.Id = m.��¼id(+) And m.����(+) = 3);
  r_������Ϣ c_������Ϣ%RowType;

  Type Ty_Einvoce Is Ref Cursor;
  c_Balanceinfo Ty_Einvoce; --��̬�α����

Begin
  j_Input    := PLJson(Json_In);
  j_Json     := j_Input.Get_Pljson('input');
  n_����     := Nvl(j_Json.Get_Number('occasion'), 0);
  n_��ѯ���� := Nvl(j_Json.Get_Number('query_type'), 0);
  v_Nos      := j_Json.Get_String('fee_nos');

  If n_��ѯ���� = 1 Then
    --������ԭʼ�Ľ�����Ϣ
    If Nvl(n_����, 0) = 2 Then
      Select a.Id As ����id, Max(a.���) As ���ʽ��, Max(a.Ԥ������Ʊ��) As �Ƿ����Ʊ��,
             To_Char(Max(a.�տ�ʱ��), 'yyyyy-mm-dd hh24:mi:ss') As �շ�ʱ��
      Into n_����id, n_���ʽ��, n_�Ƿ����Ʊ��, v_�տ�ʱ��
      From ����Ԥ����¼ A
      Where a.No = v_Nos And a.��¼���� = 1 And a.��¼״̬ In (1, 3);
    Elsif Nvl(n_����, 0) = 4 Or Nvl(n_����, 0) = 1 Then
    
      Select Max(a.����id) As ����id, Max(a.���ʽ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��,
             To_Char(Max(a.�Ǽ�ʱ��), 'yyyyy-mm-dd hh24:mi:ss') As �շ�ʱ��
      Into n_����id, n_���ʽ��, n_�Ƿ����Ʊ��, v_�տ�ʱ��
      From (Select Max(����id) As ����id, Sum(���ʽ��) As ���ʽ��, Max(�Ǽ�ʱ��) As �Ǽ�ʱ��
             From ������ü�¼
             Where ��¼���� = n_���� And NO = v_Nos And ��¼״̬ In (1, 3)) A, ����Ԥ����¼ B
      Where a.����id = b.����id;
    Elsif Nvl(n_����, 0) = 5 Then
      Select Max(a.����id) As ����id, Max(a.���ʽ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��,
             To_Char(Max(a.�Ǽ�ʱ��), 'yyyyy-mm-dd hh24:mi:ss') As �շ�ʱ��
      Into n_����id, n_���ʽ��, n_�Ƿ����Ʊ��, v_�տ�ʱ��
      From (Select Max(����id) As ����id, Sum(���ʽ��) As ���ʽ��, Max(�Ǽ�ʱ��) As �Ǽ�ʱ��
             From סԺ���ü�¼
             Where ��¼���� = 5 And NO = v_Nos And ��¼״̬ In (1, 3)) A, ����Ԥ����¼ B
      Where a.����id = b.����id;
    Else
      Json_Out := zlJsonOut('���Ͻڵ㴫��ֵ����!');
      Return;
    End If;
    --������Ϣ
    v_Output := v_Output || '"balance_info":';
    v_Output := v_Output || '{"balance_oldid":' || zlJsonStr(n_����id, 1);
    v_Output := v_Output || ',"create_time":"' || zlJsonStr(v_�տ�ʱ��) || '"';
    v_Output := v_Output || ',"start_einv":' || zlJsonStr(n_�Ƿ����Ʊ��, 1);
    v_Output := v_Output || ',"total":' || zlJsonStr(n_���ʽ��, 1);
    --�����ݲ����أ�������,��Ҫʱ�ټ�
    v_Output := v_Output || ',"balance_type":0';
    v_Output := v_Output || ',"balance_unit":0';
    v_Output := v_Output || '}';
  
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":{' || v_Output || '}}}';
  
    Return;
  End If;
  If Nvl(n_����, 0) = 2 Then
    Open c_Balanceinfo For
      Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, a.����id, a.��ҳid, Sum(a.���) As ���ʽ��, Max(a.Ԥ������Ʊ��) As �Ƿ����Ʊ��,
             Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.סԺ��) As סԺ��, Max(a.�����) As �����, Max(m.����) As ����,
             To_Char(Max(Decode(a.��¼״̬, 2, To_Date(Null), a.�տ�ʱ��)), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ��,
             Max(a.Ԥ�����) As ��������
      From ����Ԥ����¼ A, ���ս����¼ M
      Where a.No = v_Nos And a.��¼���� = 1 And a.Id = m.��¼id(+) And m.����(+) = 3
      Group By a.Id, a.No, a.����id, a.��ҳid;
  
  Elsif Nvl(n_����, 0) = 5 Then
    Open c_Balanceinfo For
      Select a.����id As ID, a.����id, a.��ҳid, Max(a.���ʽ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, Max(a.����) As ����,
             Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(b.סԺ��) As סԺ��, Max(Nvl(a.�����, b.�����)) As �����, Max(m.����) As ����,
             Max(a.�շ�ʱ��) As �շ�ʱ��, 1 As ��������
      From (Select Max(Decode(a.��¼״̬, 2, 0, 11, 0, a.����id)) As ����id, Max(a.����id) As ����id, Max(a.��ҳid) As ��ҳid,
                    Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.��ʶ��) As �����, Sum(a.���ʽ��) As ���ʽ��,
                    To_Char(Max(a.�Ǽ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ��
             From סԺ���ü�¼ A
             Where a.No = v_Nos And ��¼���� = 5) A, ����Ԥ����¼ B, ���ս����¼ M
      Where a.����id = b.����id And a.����id = m.��¼id(+) And m.����(+) = 1
      Group By a.����id, a.����id, a.��ҳid;
  Elsif Nvl(n_����, 0) = 4 Then
    --�Һ�
  
    Open c_Balanceinfo For
      Select a.����id As ID, a.����id, a.��ҳid, Max(a.���ʽ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, Max(a.����) As ����,
             Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(b.סԺ��) As סԺ��, Max(Nvl(a.�����, b.�����)) As �����, Max(m.����) As ����,
             Max(a.�շ�ʱ��) As �շ�ʱ��, 1 As ��������
      From (Select Max(Decode(a.��¼״̬, 2, 0, 11, 0, a.����id)) As ����id, Max(a.����id) As ����id, Max(a.��ҳid) As ��ҳid,
                    Max(a.����) As ����, Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.��ʶ��) As �����, Sum(a.���ʽ��) As ���ʽ��,
                    To_Char(Max(a.�Ǽ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ��
             From ������ü�¼ A
             Where a.No = v_Nos And ��¼���� = 4) A, ����Ԥ����¼ B, ���ս����¼ M
      Where a.����id = b.����id And a.����id = m.��¼id(+) And m.����(+) = 1
      Group By a.����id, a.����id, a.��ҳid;
  Elsif Nvl(n_����, 0) = 1 Then
    --�շ�
    --ע�⣺һ�ν���ĵ��ݺű���ȫ����
    Open c_Balanceinfo For
      Select a.����id As ID, a.����id, a.��ҳid, Max(a.���ʽ��) As ���ʽ��, Max(b.�Ƿ����Ʊ��) As �Ƿ����Ʊ��, Max(a.����) As ����,
             Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(b.סԺ��) As סԺ��, Max(Nvl(a.�����, b.�����)) As �����, Max(m.����) As ����,
             Max(a.�շ�ʱ��) As �շ�ʱ��, 1 As ��������
      From (Select /*+ cardinality(b, 10) */
              Max(Decode(a.��¼״̬, 2, 0, 11, 0, a.����id)) As ����id, Max(a.����id) As ����id, Max(a.��ҳid) As ��ҳid, Max(a.����) As ����,
              Max(a.�Ա�) As �Ա�, Max(a.����) As ����, Max(a.��ʶ��) As �����, Sum(a.���ʽ��) As ���ʽ��,
              To_Char(Max(a.�Ǽ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �շ�ʱ��
             From ������ü�¼ A, Table(f_Str2List(v_Nos)) B
             Where a.No = b.Column_Value And Mod(��¼����, 10) = 1) A, ����Ԥ����¼ B, ���ս����¼ M
      Where a.����id = b.����id And a.����id = m.��¼id(+) And m.����(+) = 1
      Group By a.����id, a.����id, a.��ҳid;
    --ELSIF nvl(n_����,0) = 3 THEN 
    --�ݲ�֧��
  Else
    Json_Out := zlJsonOut('���Ͻڵ㴫��ֵ����!');
    Return;
  End If;

  Fetch c_Balanceinfo
    Into r_������Ϣ;

  If c_Balanceinfo %NotFound Then
    Close c_Balanceinfo;
    Json_Out := zlJsonOut('δ�ҵ�ԭʼ����(NO=' || v_Nos || ')�ĵ���Ʊ�ݣ�����!');
    Return;
  End If;

  v_Output := v_Output || '{"pati_id":' || zlJsonStr(r_������Ϣ.����id, 1);
  v_Output := v_Output || ',"pati_pageid":' || zlJsonStr(r_������Ϣ.��ҳid, 1);
  v_Output := v_Output || ',"pati_name":"' || zlJsonStr(r_������Ϣ.����) || '"';
  v_Output := v_Output || ',"pati_sex":"' || zlJsonStr(r_������Ϣ.�Ա�) || '"';
  v_Output := v_Output || ',"pati_age":"' || zlJsonStr(r_������Ϣ.����) || '"';
  v_Output := v_Output || ',"outpatient_num":"' || zlJsonStr(r_������Ϣ.�����) || '"';
  v_Output := v_Output || ',"inpatient_num":"' || zlJsonStr(r_������Ϣ.סԺ��) || '"';
  v_Output := v_Output || ',"insurance_type":' || zlJsonStr(r_������Ϣ.����, 1);

  v_Output := v_Output || '}';

  v_Output := '"pati_info":' || v_Output;
  --������Ϣ
  v_Output := v_Output || ',"balance_info":';
  v_Output := v_Output || '{"balance_oldid":' || zlJsonStr(r_������Ϣ.Id, 1);
  v_Output := v_Output || ',"create_time":"' || zlJsonStr(r_������Ϣ.�շ�ʱ��) || '"';
  v_Output := v_Output || ',"total":' || zlJsonStr(r_������Ϣ.���ʽ��, 1);
  v_Output := v_Output || ',"start_einv":' || zlJsonStr(r_������Ϣ.�Ƿ����Ʊ��, 1);
  v_Output := v_Output || ',"balance_type":' || zlJsonStr(r_������Ϣ.��������, 1);
  v_Output := v_Output || ',"balance_unit":0';
  v_Output := v_Output || '}';

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":{' || v_Output || '}}}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbalanceinfo;
/


Create Or Replace Procedure Zl_Exsesvr_Billverify_Check
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����һ��סԺ���ʻ��۵��ĺϷ��Լ��
  --��Σ�Json_In:��ʽ
  --  input
  --    fee_nos                       C   1   ���ݺţ����������,����ʱ���ö� \�ŷָ�
  --    serials_num                   C   1   ���,����ö��ŷ���,Ϊ��Ϊ����,��fee_nos������ŵ���ʱ��������Ч
  --    pati_list[]������Ϣ���������Щ���˵ķ���
  --      pati_id                     N   1   ����ID
  --      fee_audit_status            N   1   ������˱�־:0���-δ���;1-����˻�ʼ���(��ϲ���:������˷�ʽ������);2-������,��Ͻ���Ȩ��[��ֹδ��˲��˽���]���й������
  --      si_inp_status               N   1   סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
  --����: Json_Out,��ʽ����
  --  output
  --    code                          N   1   Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C   1   Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    item_list[]
  --      rcp_no                      C   1   ��������
  --      stuff_rcpdtl_ids            C   1   ���Ĵ�����ϸIDs:�����������漰�ķ���ids
  --      drug_rcpdtl_ids             C   1   ҩƷ������ϸIDs:ҩƷ���漰�ķ���ids
  --      autosendstuff_rcpdtl_ids    C   1   ���ϴ�����ϸIDs:�Զ����������������漰�ķ���IDs
  ---------------------------------------------------------------------------
Begin
  Zl_סԺ���ʼ�¼_Verify_Check(Json_In, Json_Out);
End Zl_Exsesvr_Billverify_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Updrgstarrangement
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  --���ܣ�������Դ����Ч�İ��š���Ч�ĳ����¼�е�ҽ��������
  --��Σ�n_������ʽ:1-�޸�����,2-ͣ����Ա,3-������Ա
  --      d_����ʱ��:ͣ�ú�����ʱ���룬����ʱ����ԭ����ʱ��
  --˵�����ù��̹���Ա������������Աͣ��/����ʱ���ã�ͬ�������ҺŰ���
  v_Para     Varchar2(200);
  n_�Һ�ģʽ Number(2);
  j_Input    PLJson;
  j_Json     PLJson;

  n_��Աid   �ҺŰ���.ҽ��id%Type;
  n_������ʽ Number(2);
  v_��Ա���� ��Ա��.����%Type;
  d_��ʼʱ�� �ٴ������¼.��ʼʱ��%Type;
  d_����ʱ�� ��Ա��.����ʱ��%Type;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��Աid   := j_Json.Get_Number('rgst_dr_id');
  n_������ʽ := j_Json.Get_Number('oper_type');
  d_����ʱ�� := To_Date(j_Json.Get_String('revoke_time'), 'YYYY-MM-DD hh24:mi:ss');
  If n_������ʽ = 2 And d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;

  Begin
    Select a.����
    Into v_��Ա����
    From ��Ա�� A, ��Ա����˵�� B
    Where a.Id = n_��Աid And a.Id = b.��Աid And b.��Ա���� = 'ҽ��';
  Exception
    When Others Then
      --���账��,�˳�
      Json_Out := zlJsonOut('�ɹ�', 1);
      Return;
  End;

  v_Para     := zl_GetSysParameter(256);
  n_�Һ�ģʽ := Substr(v_Para, 1, 1);

  If n_�Һ�ģʽ = 1 Then
    --�����ģʽ
    If n_������ʽ = 1 Then
      --�޸�
      Update �ٴ������Դ Set ҽ������ = v_��Ա���� Where ҽ��id = n_��Աid;
      Update �ٴ����ﰲ�� Set ҽ������ = v_��Ա���� Where ҽ��id = n_��Աid And ��ֹʱ�� > Sysdate;
      Update �ٴ������¼ Set ҽ������ = v_��Ա���� Where ҽ��id = n_��Աid And �������� >= Trunc(Sysdate);
      Update �ٴ������¼ Set ����ҽ������ = v_��Ա���� Where ����ҽ��id = n_��Աid And �������� >= Trunc(Sysdate);
    Elsif n_������ʽ = 2 Then
      --ͣ��
      Update �ٴ������Դ
      Set ����ʱ�� = d_����ʱ��
      Where ҽ��id = n_��Աid And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'));
    
      --����ǰʱ���Ժ�����г����¼ͣ��
      For c_��¼ In (Select ID, ��ʼʱ��, ��ֹʱ��
                   From �ٴ������¼
                   Where ҽ��id = n_��Աid And �������� >= Trunc(d_����ʱ��) - 1 And ��ֹʱ�� > d_����ʱ�� And �ϰ�ʱ�� Is Not Null) Loop
      
        If c_��¼.��ʼʱ�� < d_����ʱ�� Then
          d_��ʼʱ�� := d_����ʱ��;
        Else
          d_��ʼʱ�� := c_��¼.��ʼʱ��;
        End If;
        Zl_�ٴ������¼_Stopvisit(c_��¼.Id, d_��ʼʱ��, c_��¼.��ֹʱ��, '��Աͣ��', zl_UserName, d_����ʱ��, 0, 1);
      End Loop;
    Elsif n_������ʽ = 3 Then
      --����
      For c_��¼ In (Select a.Id
                   From �ٴ������¼ A, �ٴ������Դ B
                   Where a.��Դid = b.Id And b.����ʱ�� = d_����ʱ�� And b.ҽ��id = n_��Աid And a.�������� >= Trunc(Sysdate) - 1 And
                         a.��ֹʱ�� > Sysdate And a.�ϰ�ʱ�� Is Not Null) Loop
      
        Zl_�ٴ������¼_Stopvisit(c_��¼.Id, Null, Null, Null, zl_UserName, Sysdate, 1, 1);
      End Loop;
    
      Update �ٴ������Դ
      Set ����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')
      Where ҽ��id = n_��Աid And ����ʱ�� = d_����ʱ��;
    
      --���ɳ����¼
      Zl1_Auto_Buildingregisterplan;
    End If;
  Else
    --�Һ��Ű�ģʽ
    If n_������ʽ = 1 Then
      --�޸�
      Update �ҺŰ��� Set ҽ������ = v_��Ա���� Where ҽ��id = n_��Աid And (��ֹʱ�� Is Null Or ��ֹʱ�� > Sysdate);
      Update �ҺŰ��żƻ�
      Set ҽ������ = v_��Ա����
      Where ҽ��id = n_��Աid And (ʧЧʱ�� Is Null Or ʧЧʱ�� > Sysdate);
    Elsif n_������ʽ = 2 Then
      --ͣ��
      Update �ҺŰ��� Set ͣ������ = d_����ʱ�� Where ҽ��id = n_��Աid And (��ֹʱ�� Is Null Or ��ֹʱ�� > Sysdate);
      Update �ҺŰ��żƻ�
      Set ʧЧʱ�� = d_����ʱ��
      Where ҽ��id = n_��Աid And (ʧЧʱ�� Is Null Or ʧЧʱ�� > Sysdate);
    Elsif n_������ʽ = 3 Then
      --���ã�������
      Null;
    End If;
  End If;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updrgstarrangement;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Bulidregistprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --����ҽ��ID��ѯ�Ƿ��ڷ��ñ���ڼ�¼
  ---------------------------------------------------------------------------
  --input      ���ݹҺŵ������շѻ��۵�
  --  rgst_no         C  1  �Һŵ���
  --  pati_id         N     ����ID
  --  pati_name       C     ����
  --  pati_sex        C     �Ա�
  --  pati_age        C     ����
  --  pati_idcard     C     ���֤��
  --  birth_date      C     ��������
  --  rgst_dept_id     N  1  �Һſ���ID
  --  rgst_dr          C  1  ҽ������
  --  operator_name    C  1  ����Ա����
  --  site_no          C    վ��
  --  rgst_visitinfo      ���˾�����Ϣ
  --    outp_room_name  C    �������
  --    emg_sign        N    �����־
  --    revisit_sign    N    �����־
  --    exe_time        C    ִ��ʱ��
  --  ����      json
  --output
  --  code    N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message  C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  fee_no  C  1  ���۵���
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json1 PLJson;

  n_����id     ���˹Һż�¼.����id%Type;
  v_�Һŵ�     ���˹Һż�¼.No%Type;
  n_����id     ���˹Һż�¼.ִ�в���id%Type;
  v_ҽ������   ���˹Һż�¼.ִ����%Type;
  v_����Ա���� ���˹Һż�¼.����Ա����%Type;
  v_����       ���˹Һż�¼.����%Type;
  n_�����־   ���˹Һż�¼.����%Type;
  n_�����־   Integer;
  d_ִ��ʱ��   ���˹Һż�¼.ִ��ʱ��%Type;
  v_վ��       Varchar2(100);
  v_���۵�     ���˹Һż�¼.�շѵ�%Type;
  v_����       ���˹Һż�¼.����%Type;
  v_�Ա�       ���˹Һż�¼.�Ա�%Type;
  v_����       ���˹Һż�¼.����%Type;
  d_��������   Date;
  v_���֤��   Varchar2(18);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�     := j_Json.Get_String('rgst_no');
  n_����id     := j_Json.Get_Number('rgst_dept_id');
  v_ҽ������   := j_Json.Get_String('rgst_dr');
  v_����Ա���� := j_Json.Get_String('operator_name');
  v_վ��       := j_Json.Get_String('site_no');
  n_����id     := j_Json.Get_Number('pait_id');
  v_����       := j_Json.Get_String('pati_name');
  v_�Ա�       := j_Json.Get_String('pati_sex');
  v_����       := j_Json.Get_String('pati_age');
  d_��������   := To_Date(j_Json.Get_String('birth_date'), 'YYYY-MM-DD hh24:mi:ss');
  v_���֤��   := j_Json.Get_String('pati_idcard');

  If Nvl(n_����id, 0) = 0 Then
    n_����id := Null;
  End If;
  Select Zl_Exse_Nextno(12, Null) Into v_���۵� From Dual;

  Zl_���ﻮ�ۼ�¼_Buliding_s(v_�Һŵ�, v_���۵�, n_����id, v_����, v_�Ա�, v_����, d_��������, v_���֤��, n_����id, v_ҽ������, v_����Ա����, v_վ��);

  o_Json1 := j_Json.Get_Pljson('rgst_visitinfo');
  If Not o_Json1 Is Null Then
    n_����id   := j_Json.Get_Number('exe_deptid');
    v_����     := o_Json1.Get_String('outp_room_name');
    n_�����־ := o_Json1.Get_Number('emg_sign');
    n_�����־ := o_Json1.Get_Number('revisit_sign');
    d_ִ��ʱ�� := To_Date(o_Json1.Get_String('exe_time'), 'yyyy-mm-dd hh24:mi:ss');

    Zl_���˽���_s(n_����id, v_�Һŵ�, n_����id, v_ҽ������, v_����, n_�����־, n_�����־, d_ִ��ʱ��);
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_no":"' || v_���۵� || '"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Bulidregistprice;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Checknoischarge
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --����:��鵥���Ƿ����շѣ�Ŀǰ���ڹҺŻ��۵����
  ---------------------------------------------------------------------------
  --input
  --  fee_no          C  1  ���ݺ�
  --  checkCharge      N    ��黮�۵��Ƿ��շ�
  --  rgst_dept_id    N    ִ�в���ID
  --  rgst_dr          C    ִ����
  --output
  --  code        N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message      C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  fee_status  C  1  ����״̬��-1-δ�ҵ���Ӧ�ĹҺŵ���,0-δ�շ�;1-�Һŵ�����;2-��δ�������ۼ�¼; 3-�Һŵ���Ӧ���շѻ��۵���ȫ�շ�(���ڶ��Ż��۵�ʱ������ȫ�յ�);4-�Һŵ���Ӧ�Ļ��۵����ڲ����շ�)
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_���ݺ�       ������ü�¼.No%Type;
  n_��¼����     ���˹Һż�¼.��¼����%Type;
  v_�շѵ�       ���˹Һż�¼.�շѵ�%Type;
  n_ȡ�ű�־     ���˹Һż�¼.ȡ�ű�־%Type;
  n_����id       ������ü�¼.����id%Type;
  n_Min����id    ������ü�¼.����id%Type;
  n_Max����id    ������ü�¼.����id%Type;
  n_ִ�в���id   ������ü�¼.ִ�в���id%Type;
  v_ִ����       ������ü�¼.ִ����%Type;
  n_Checkcharge  Number(2); --��鵥���Ƿ��շ�
  n_Count        Number(2);
  n_ִ�в���id_b Number(2);
  n_ִ����_b     Number(2);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ�      := j_Json.Get_String('fee_no');
  n_Checkcharge := j_Json.Get_Number('checkcharge');

  If j_Json.Exist('input.rgst_dept_id') Then
    n_ִ�в���id   := j_Json.Get_Number('rgst_dept_id ');
    n_ִ�в���id_b := 1;
  End If;
  If j_Json.Exist('input.rgst_dr') Then
    v_ִ����   := j_Json.Get_String('rgst_dr');
    n_ִ����_b := 1;
  End If;

  Select Count(1), Max(a.�շѵ�) As �շѵ�, Max(a.ȡ�ű�־) As ȡ�ű�־, Max(b.����id) As ����id, Max(a.��¼����) As ��¼����
  Into n_Count, v_�շѵ�, n_ȡ�ű�־, n_����id, n_��¼����
  From ���˹Һż�¼ A, ������ü�¼ B
  Where a.No = v_���ݺ� And a.No = b.No And b.��¼���� = 4 And b.��¼״̬ In (0, 1, 3);

  If n_Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ�ҵ���Ӧ�ĹҺŵ���","fee_status":-1}}';
    Return;
  End If;

  If v_�շѵ� Is Null Then
    If Nvl(n_ȡ�ű�־, 0) = 1 Then
      --��Һ�ģʽ��δ���ɻ��۵�
      Json_Out := '{"output":{"code":0,"message":"��δ�������ۼ�¼","fee_status":2}}';
    Elsif Nvl(n_����id, 0) = 0 Or Nvl(n_��¼����, 0) = 2 Then
      --δ���˻�ԤԼδ����
      Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_status":0}}';
    Else
      --���շ�
      Json_Out := '{"output":{"code":0,"message":"�Һŵ����շ�","fee_status":1}}';
    End If;
    Return;
  End If;

  If Nvl(n_Checkcharge, 0) = 0 Then
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_status":0}}';
  End If;

  Select /* +cardinality(b,10) */
   Count(1), Min(Decode(����id, Null, 0, ����id)), Max(����id)
  Into n_Count, n_Min����id, n_Max����id
  From ������ü�¼ A, Table(f_Str2List(v_�շѵ�)) B
  Where a.��¼���� = 1 And a.��¼״̬ In (0, 1, 3) And a.No = b.Column_Value And
        a.ִ�в���id = Decode(n_ִ�в���id_b, 1, n_ִ�в���id, a.ִ�в���id) And a.ִ���� = Decode(n_ִ����_b, 1, v_ִ����, a.ִ����);

  If n_Count = 0 Then
    --û�л��۵�
    Json_Out := '{"output":{"code":0,"message":"��δ�������ۼ�¼","fee_status":2}}';

  Elsif Nvl(n_Max����id, 0) = 0 Then
    --δ�շ�
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_status":0,"fee_status":"' || v_�շѵ� || '"}}';

  Elsif Nvl(n_Min����id, 0) = 0 And Nvl(n_Max����id, 0) > 0 Then
    --δȫ�շ�
    Json_Out := '{"output":{"code":0,"message":"�Һŵ���Ӧ���շѻ��۵���ȫ�շ�","fee_status":3}}';

  Else
    --ȫ�շ�
    Json_Out := '{"output":{"code":0,"message":"�Һŵ���Ӧ�Ļ��۵����ڲ����շ�","fee_status":4}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checknoischarge;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Checkdelregistprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --����:ɾ���ҺŻ��۵�ǰ��鵥���Ƿ�����ɾ��
  ---------------------------------------------------------------------------
  --input
  --  fee_no          C  1  ���ݺ�
  --  checkCharge     N    ��黮�۵��Ƿ��շ�
  --  rgst_dept_id    N    ִ�в���ID
  --  rgst_dr         C    ִ����
  --output
  --  code        N  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message      C  1  Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --  fee_status  C  1  ����״̬��0-��������״̬;1-δ�ҵ��Һŵ�;2-δ���ɻ��۵�;3-δ�ҵ����������Ļ��۵�;4-�����Ѿ��շѵĵ���
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_���ݺ�       ������ü�¼.No%Type;
  v_�շѵ�       ���˹Һż�¼.�շѵ�%Type;
  n_ִ�в���id   ������ü�¼.ִ�в���id%Type;
  v_ִ����       ������ü�¼.ִ����%Type;
  n_Count        Number(2);
  n_Code         Number(2);
  n_ִ�в���id_b Number(2);
  n_ִ����_b     Number(2);
  v_Input        Varchar2(100);
  v_Output       Varchar2(100);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ� := j_Json.Get_String('fee_no');

  If j_Json.Exist('rgst_dept_id') Then
    n_ִ�в���id   := j_Json.Get_Number('rgst_dept_id');
    n_ִ�в���id_b := 1;
  End If;
  If j_Json.Exist('rgst_dr') Then
    v_ִ����   := j_Json.Get_String('rgst_dr');
    n_ִ����_b := 1;
  End If;

  Select Count(1), Max(a.�շѵ�) As �շѵ�
  Into n_Count, v_�շѵ�
  From ���˹Һż�¼ A
  Where a.No = v_���ݺ� And a.��¼״̬ In (0, 1, 3);

  If n_Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ�ҵ��Һŵ�","fee_status":1}}';
    Return;
  End If;

  If v_�շѵ� Is Null Then
    Json_Out := '{"output":{"code":0,"message":"δ���ɻ��۵�","fee_status":2}}';
    Return;
  End If;

  n_Count := 0;
  For c_Price In (Select /* +cardinality(b,10) */
                   a.No, Max(����id) As ����id
                  From ������ü�¼ A, Table(f_Str2List(v_�շѵ�)) B
                  Where a.��¼���� = 1 And a.��¼״̬ In (0, 1, 3) And a.No = b.Column_Value And
                        a.ִ�в���id = Decode(n_ִ�в���id_b, 1, n_ִ�в���id, a.ִ�в���id) And
                        a.ִ���� = Decode(n_ִ����_b, 1, v_ִ����, a.ִ����)
                  Group By a.No) Loop

    If Nvl(c_Price.����id, 0) > 0 Then
      Json_Out := '{"output":{"code":0,"message":"�����Ѿ��շѵĵ���","fee_status":4}}';
      Return;
    End If;

    v_Input := '{"input":{"fee_no":"' || c_Price.No || '"}}';
    Zl_���ﻮ�ۼ�¼_Delete_Check(v_Input, v_Output);
    j_Json  := PLJson();
    j_Json  := PLJson(v_Output);
    j_Json  := j_Input.Get_Pljson('output');
    n_Code  := j_Json.Get_Number('code');
    If Nvl(n_Code, 0) = 0 Then
      Json_Out := '{"output":{"code":0,"message":"δ�ҵ����������Ļ��۵�","fee_status":3}}';
      Return;
    End If;
    n_Count := n_Count + 1;
  End Loop;
  If n_Count = 0 Then
    Json_Out := '{"output":{"code":0,"message":"δ�ҵ����������Ļ��۵�","fee_status":3}}';
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","fee_status":0}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Checkdelregistprice;
/


CREATE OR REPLACE Procedure Zl_Exsesvr_Delregistprice
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) Is
  --����:���ҺŻ��۵��Ƿ�����ɾ��
  ---------------------------------------------------------------------------
  --input
  --  fee_no              C  1  ���ݺ�
  --  rgst_dept_id        N     ����ID
  --  rgst_dr             C     ҽ������
  --  rgst_visitinfo      N
  --     exe_deptid       N   ִ�в���ID
  --     exetr            C   ִ����
  --     referral_sign    N   �Ƿ�ת��: 0-δת��  1-ת��
  --     referral_deptid  N   ת�����ID
  --     referral_doctor  C   ת��ҽ��
  --output
  --  code        N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --  message     C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  o_Json PLJson;

  v_���ݺ�       ������ü�¼.No%Type;
  n_����id       ���˹Һż�¼.����id%Type;
  v_�շѵ�       ���˹Һż�¼.�շѵ�%Type;
  n_ִ�в���id   ���˹Һż�¼.ִ�в���id%Type;
  v_ִ����       ���˹Һż�¼.ִ����%Type;
  n_ת�����id   ���˹Һż�¼.ת�����id%Type;
  v_ת��ҽ��     ���˹Һż�¼.ת������%Type;
  n_ת���־     Number(2);
  n_ִ�в���id_b Number(2);
  n_ִ����_b     Number(2);
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_���ݺ� := j_Json.Get_String('fee_no');

  If j_Json.Exist('rgst_dept_id') Then
    n_ִ�в���id   := j_Json.Get_Number('rgst_dept_id');
    n_ִ�в���id_b := 1;
  End If;
  If j_Json.Exist('rgst_dr') Then
    v_ִ����   := j_Json.Get_String('rgst_dr');
    n_ִ����_b := 1;
  End If;

  Select Max(����id), Max(a.�շѵ�)
  Into n_����id, v_�շѵ�
  From ���˹Һż�¼ A
  Where a.No = v_���ݺ� And a.��¼״̬ In (0, 1, 3);
  If v_�շѵ� Is Null Then
    Json_Out := zlJsonOut('�����ڹҺŻ��۵���');
    Return;
  End If;

  For c_Price In (Select /* +cardinality(b,10) */
                  Distinct a.No
                  From ������ü�¼ A, Table(f_Str2List(v_�շѵ�)) B
                  Where a.��¼���� = 1 And a.��¼״̬ In (0, 1, 3) And a.No = b.Column_Value And
                        a.ִ�в���id = Decode(n_ִ�в���id_b, 1, n_ִ�в���id, a.ִ�в���id) And
                        a.ִ���� = Decode(n_ִ����_b, 1, v_ִ����, a.ִ����)) Loop
    Zl_���ﻮ�ۼ�¼_Delete_s(c_Price.No);
  End Loop;

  o_Json := j_Json.Get_Pljson('rgst_visitinfo');
  If Not o_Json Is Null Then
    n_ִ�в���id := o_Json.Get_Number('exe_deptid');
    v_ִ����     := o_Json.Get_String('exetr');
    n_ת���־   := Nvl(o_Json.Get_Number('referral_sign'), 0);
    n_ת�����id := o_Json.Get_Number('referral_dept_id');
    v_ת��ҽ��   := o_Json.Get_String('referral_dr');

    If n_ִ�в���id = 0 Then
      n_ִ�в���id := Null;
    End If;
    If n_ת�����id = 0 Then
      n_ת�����id := Null;
    End If;

    Zl_���˽���_Cancel_s(n_����id, v_���ݺ�, n_ִ�в���id, v_ִ����, n_ת���־, v_ת��ҽ��, n_ת�����id);
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Delregistprice;
/

Create Or Replace Procedure Zl_Exsesvr_Chkpatichangenurse
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --����:����ȼ��������
  --��Σ�Json_In:��ʽ
  --input
  --      pati_id           N 1 ����id
  --      pati_pageid       N 1 ��ҳID
  --      create_time       C 1 �Ǽ�ʱ��
  --����  json
  --output
  --     code               N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --     message            C 1 Ӧ����Ϣ�� ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_����id   סԺ���ü�¼.����id%Type;
  n_��ҳid   סԺ���ü�¼.��ҳid%Type;
  d_��ʼʱ�� Date;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  d_��ʼʱ�� := To_Date(j_Json.Get_String('create_time'), 'YYYY-MM-DD HH24:MI:SS');
  For r_Fee In (Select NO
                From סԺ���ü�¼
                Where ����id = n_����id And ��ҳid = n_��ҳid And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= d_��ʼʱ�� And �շ���� = 'H'
                Group By NO, ���, Mod(��¼����, 10)
                Having Sum(���ʽ��) <> 0) Loop
    Json_Out := zlJsonOut('�䶯ʱ��֮�������ѽ��ʵ��Զ����ʷ���,���ܸ��Ļ���ȼ���');
    Return;
  End Loop;
  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Chkpatichangenurse;
/



Create Or Replace Procedure Zl_Exsesvr_Getpatireceivables
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡָ�����˵�Ӧ�տ����
  --��Σ�Json_In:��ʽ
  --  input
  --   pati_id            N   ����id
  --����: Json_Out,��ʽ����
  --  output
  --    code              N 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message           C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    fee_amrcvb        N 1 Ӧ�ս��
  ---------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  v_Output   Varchar2(32767);
  n_����id   ����Ԥ����¼.����id%Type;
  n_Ӧ����� ����Ԥ����¼.��Ԥ��%Type;
Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id := j_Json.Get_Number('pati_id');

  If Nvl(n_����id, 0) = 0 Then
    Json_Out := zlJsonOut('����IDδ���롣');
    Return;
  End If;

  Begin
    Select Nvl(a.Ӧ�տ��ܶ�, 0) - Nvl(Sum(���), 0)
    Into n_Ӧ�����
    From (Select a.����id, Sum(a.��Ԥ��) Ӧ�տ��ܶ�
           From ����Ԥ����¼ A, ���㷽ʽ B
           Where a.����id = n_����id And a.���㷽ʽ = b.���� And b.Ӧ�տ� = 1
           Group By ����id) A, ���˽ɿ��¼ B
    Where a.����id = b.����id(+) And b.��¼״̬(+) = 1
    Group By a.����id, Ӧ�տ��ܶ�;
  Exception
    When Others Then
      n_Ӧ����� := 0;
  End;
  zlJsonPutValue(v_Output, 'code', 1, 1, 1);
  zlJsonPutValue(v_Output, 'message', '�ɹ�');
  zlJsonPutValue(v_Output, 'fee_amrcvb', n_Ӧ�����, 1, 2);
  Json_Out := '{"output":' || v_Output || '}';
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getpatireceivables;
/

Create Or Replace Procedure Zl_Exsesvr_Getrgstinfo
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ���ȡԤԼ�Һŵ�����Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --  rgst_no             C  1 �Һŵ�
  --  appt_recv           N    ԤԼ����
  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    close_account_type              N 1 �Һ���Ч�����ڵĽ���ģʽ
  --    fee_list                        ������Ϣ�б�
  --       pati_id          N   1 ����ID
  --       outpatient_num   C   1 �����
  --       rgst_id          N   1 �Һ�id
  --       pati_name        C   1 ����
  --       pati_sex         C   1 �Ա�
  --       pati_age         C   1 ����
  --       fee_category     C   1 �ѱ�
  --       num_category     C   1 �ű�
  --       mdlpay_mode_name C   1 ���ʽ
  --       overtime_sign    N   1 �Ӱ��־
  --       exe_deptid       N   1 ִ�в���id
  --       happen_time      C   1 ����ʱ��
  --       appt_time        C   1 ԤԼʱ��:yyyy-mm-dd hh24:mi:ss
  --       rgst_time        C   1 �Ǽ�ʱ��
  --       operator_code    C   1 ����Ա���
  --       operator_name    C   1 ����Ա����
  --       appt_mode_name   C   1 ԤԼ��ʽ
  --       fee_ampaid       N   1 ʵ�ս��
  --       fee_item_id      N   1 �շ�ϸĿid
  --       outptyp_name     C   1 ����
  -------------------------------------------
  v_Output     Varchar2(32000);
  v_�Һŵ�     ���˹Һż�¼.No%Type;
  n_ԭ����ģʽ ���˹Һż�¼.����ģʽ%Type;
  n_ԤԼ����   Number(2);
  j_Input      Pljson;
  j_Json       Pljson;

Begin
  --�������
  j_Input := Pljson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�   := j_Json.Get_String('rgst_no');
  n_ԤԼ���� := j_Json.Get_Number('appt_recv');

  --��鵱ǰ�Ƿ�Һ���Ϣ�Ƿ����
  For c_������Ϣ In (Select Max(a.����id) As ����id, Max(c.Id) As �Һ�id, Max(a.��ʶ��) As �����, Max(a.����) As ����, Max(a.�Ա�) As �Ա�,
                        Max(a.����) As ����, Max(a.�ѱ�) As �ѱ�, Max(Nvl(c.�ű�, Decode(a.���, 1, a.���㵥λ, ''))) As �ű�,
                        Max(a.�Ӱ��־) As �Ӱ��־, Max(Decode(a.���, 1, a.ִ�в���id, 0)) As ִ�в���id,
                        To_Char(Max(a.����ʱ��), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��,
                        To_Char(Max(a.�Ǽ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, Max(a.����Ա���) As ����Ա���,
                        Max(a.����Ա����) As ����Ա����, Max(c.ԤԼ��ʽ) As ԤԼ��ʽ,
                        To_Char(Max(Nvl(c.ԤԼʱ��, a.����ʱ��)), 'yyyy-mm-dd hh24:mi:ss') As ԤԼʱ��,
                        Max(Nvl(c.�Һ���Ŀid, Decode(a.���, 1, a.�շ�ϸĿid, 0))) As �շ�ϸĿid, Sum(a.ʵ�ս��) As �ҺŽ��,
                        Max(b.����) As ҽ�Ƹ�������, Max(c.����) As ����, Max(c.����) As ����, Max(c.����) As ����
                 From ������ü�¼ A, ҽ�Ƹ��ʽ B, ���˹Һż�¼ C
                 Where a.No = v_�Һŵ� And a.No = c.No And a.��¼���� = 4 And a.��¼״̬ In (0, 1) And a.���ʽ = b.����(+)) Loop
    Zljsonputvalue(v_Output, 'pati_id', c_������Ϣ.����id, 1, 1);
    Zljsonputvalue(v_Output, 'outpatient_num', c_������Ϣ.�����);
    Zljsonputvalue(v_Output, 'rgst_id', c_������Ϣ.�Һ�id, 1);
    Zljsonputvalue(v_Output, 'pati_name', c_������Ϣ.����);
    Zljsonputvalue(v_Output, 'pati_sex', c_������Ϣ.�Ա�);
    Zljsonputvalue(v_Output, 'pati_age', c_������Ϣ.����);
    Zljsonputvalue(v_Output, 'fee_category', c_������Ϣ.�ѱ�);
    Zljsonputvalue(v_Output, 'mdlpay_mode_name', c_������Ϣ.ҽ�Ƹ�������);
    Zljsonputvalue(v_Output, 'num_category', c_������Ϣ.�ű�);
    Zljsonputvalue(v_Output, 'overtime_sign', c_������Ϣ.�Ӱ��־, 1);
    Zljsonputvalue(v_Output, 'exe_deptid', c_������Ϣ.ִ�в���id, 1);
    Zljsonputvalue(v_Output, 'happen_time', c_������Ϣ.����ʱ��);
    Zljsonputvalue(v_Output, 'rgst_time', c_������Ϣ.�Ǽ�ʱ��);
    Zljsonputvalue(v_Output, 'operator_code', c_������Ϣ.����Ա���);
    Zljsonputvalue(v_Output, 'operator_name', c_������Ϣ.����Ա����);
    Zljsonputvalue(v_Output, 'appt_mode_name', c_������Ϣ.ԤԼ��ʽ);
    Zljsonputvalue(v_Output, 'fee_item_id', c_������Ϣ.�շ�ϸĿid, 1);
    Zljsonputvalue(v_Output, 'appt_time', Nvl(c_������Ϣ.ԤԼʱ��, ''));
    Zljsonputvalue(v_Output, 'fee_ampaid', Nvl(c_������Ϣ.�ҺŽ��, 0), 1);
    Zljsonputvalue(v_Output, 'outptyp_name', c_������Ϣ.����);
    Zljsonputvalue(v_Output, 'revst_sign', c_������Ϣ.����, 1);
    Zljsonputvalue(v_Output, 'emg_sign', c_������Ϣ.����, 1, 2);
  
    If Nvl(c_������Ϣ.����id, 0) <> 0 And Nvl(n_ԤԼ����, 0) = 1 Then
      Zl_ԤԼ�ҺŽ���_Check_s(v_�Һŵ�, c_������Ϣ.����id, Sysdate, c_������Ϣ.����Ա����, c_������Ϣ.����, 0, 0, 0, n_ԭ����ģʽ);
    End If;
  End Loop;

  If v_Output Is Null Then
    Json_Out := Zljsonout('�Һ���Ϣ�����ڣ����ܸùҺŵ��ѱ����˴���');
    Return;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","close_account_type":' || Nvl(n_ԭ����ģʽ, 0) || ',"fee_list":[' ||
              v_Output || ']}}';

Exception
  When Others Then
    Json_Out := Zljsonout(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getrgstinfo;
/

CREATE OR REPLACE Procedure Zl_Exsesvr_Rgstapptreceive
(
  Json_In  Varchar2,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ�ԤԼ�ҺŽ���
  --��Σ�Json_In:��ʽ
  --input
  --  rgst_no             C  1 �Һ�no
  --  outp_room_name     C  1 ��������
  --  recv_time          C  1 ����ʱ��
  --  prepay_pati_ids    C  1 ��Ԥ������ids
  --  pati_inhospital    N  1 ��Ԥ������ids

  --  checkout_id        N  1 ����id
  --  cardtype_id        N  1 �����id
  --  pat_card_no        C  1 ����
  --  trans_no           C  1 ������ˮ��
  --  trans_desc         C  1 ����˵��
  --  recv_time          C  1 ����ʱ��
  --  prepay_pati_ids    C  1 ��Ԥ������ids
  --  pati_id            N  1 ����id
  --  outpatient_num     C  1 �����
  --  reg_id             N  1 �Һ�id
  --  blnc_id            N  1 ����ID
  --  relation_id        N  1 ��������ID

  -- ����:
  --  output
  --    code             N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message          C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ

  -------------------------------------------
  n_�Һ�id        ���˹Һż�¼.Id%Type;
  v_�Һŵ�        ���˹Һż�¼.No%Type;
  n_����id        ���˹Һż�¼.����id%Type;
  n_�����        ���˹Һż�¼.�����%Type;
  v_����          ���˹Һż�¼.����%Type;
  v_�Ա�          ���˹Һż�¼.�Ա�%Type;
  v_����          ���˹Һż�¼.����%Type;
  v_���ʽ����  ҽ�Ƹ��ʽ.����%Type;
  v_�ѱ�          ���˹Һż�¼.�ѱ�%Type;
  n_�����¼id    ���˹Һż�¼.�����¼id%Type;
  n_���ʽ��      ������ü�¼.���ʽ��%Type;
  n_����id        ������ü�¼.����id%Type;
  v_���۵�        ������ü�¼.No%Type;
  v_����Ա����    ���˹Һż�¼.����Ա���%Type;
  v_����Ա����    ���˹Һż�¼.����Ա����%Type;
  n_�Һ���Ŀid    ���˹Һż�¼.�Һ���Ŀid%Type;
  n_��������      ���˹Һż�¼.����%Type;
  d_����ʱ��      ���˹Һż�¼.����ʱ��%Type;
  n_ִ�в���id    ���˹Һż�¼.ִ�в���id%Type;
  v_�ű�          ���˹Һż�¼.�ű�%Type;
  v_ҽ������      ���˹Һż�¼.ִ����%Type;
  n_ҽ��id        ��Ա��.Id%Type;
  n_Ԥ��id        ����Ԥ����¼.Id%Type;
  v_��Ԥ������ids Varchar2(1000);
  v_���㷽ʽ      ����Ԥ����¼.���㷽ʽ%Type;
  n_�����id      ����Ԥ����¼.�����id%Type;
  v_֧������      ����Ԥ����¼.����%Type;
  n_�ҺŻ���      Number(2);
  n_��Ժ          Number(2);
  n_�Һ����ɶ���  Number(2);
  o_Json          PLJson;
  j_Input         PLJson;
  j_Json          PLJson;

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�        := j_Json.Get_String('reg_no');
  n_��������      := j_Json.Get_String('outp_room_name');
  n_����id        := j_Json.Get_Number('pati_id');
  n_�����        := j_Json.Get_String('outpatient_num');
  d_����ʱ��      := To_Date(j_Json.Get_String('recv_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_��Ԥ������ids := j_Json.Get_String('prepay_pati_ids');
  n_��Ժ          := j_Json.Get_Number('pati_inhospital');
  n_����id        := j_Json.Get_Number('blnc_id');
  n_Ԥ��id        := j_Json.Get_Number('relation_id');
  v_����Ա����    := j_Json.Get_String('operator_code');
  v_����Ա����    := j_Json.Get_String('operator_name');
  n_�ҺŻ���      := j_Json.Get_Number('pricing_sign');
  If Nvl(n_�ҺŻ���, 0) <> 1 Then
    n_�ҺŻ��� := 0;
  End If;

  If d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;

  Select Max(a.Id), Max(a.No), Max(a.����), Max(a.�Ա�), Max(a.����), Max(c.����), Max(a.�ѱ�), Max(a.�����¼id), Sum(b.ʵ�ս��),
         Max(a.�Һ���Ŀid), Max(a.�ű�), Max(a.ִ�в���id), Max(a.ִ����), Max(d.Id) As ҽ��id
  Into n_�Һ�id, v_�Һŵ�, v_����, v_�Ա�, v_����, v_���ʽ����, v_�ѱ�, n_�����¼id, n_���ʽ��, n_�Һ���Ŀid, v_�ű�, n_ִ�в���id, v_ҽ������, n_ҽ��id
  From ���˹Һż�¼ A, ������ü�¼ B, ҽ�Ƹ��ʽ C, ��Ա�� D
  Where a.No = v_�Һŵ� And a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = b.No And b.��¼���� = 4 And a.ҽ�Ƹ��ʽ = c.����(+) And
        a.ִ���� = d.����(+);

  If v_�Һŵ� Is Null Then
    Json_Out := zlJsonOut('ԤԼ�Һ���Ϣ�����ڣ����ܸ�ԤԼ�Һ��ѱ����ա�');
    Return;
  End If;

  If n_�ҺŻ��� = 1 Then
    n_���ʽ�� := 0;
    Select Nextno(13) Into v_���۵� From Dual;
  
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����, ������Ŀid,
       �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ժҪ, �Ƿ���, �Һ�id, ��ҳid, ���ʽ)
      Select ���˷��ü�¼_Id.Nextval, 1, v_���۵�, 0, a.���, a.��������, a.�۸񸸺�, a.�����־, n_����id, Decode(n_�����, 0, Null, n_�����), a.����,
             a.�Ա�, a.����, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, b.���㵥λ, a.����, a.����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��,
             0, v_����Ա����, n_ִ�в���id, v_����Ա����, a.����ʱ��, d_����ʱ��, a.ִ�в���id, '�Һ�:' || v_�Һŵ�, a.�Ƿ���, a.�Һ�id, a.��ҳid, a.���ʽ
      From ������ü�¼ A, �շ���ĿĿ¼ B
      Where a.No = v_�Һŵ� And a.��¼���� = 4 And a.��¼״̬ = 0 And a.�շ�ϸĿid = b.Id;
  
    Update ������ü�¼
    Set Ӧ�ս�� = 0, ʵ�ս�� = 0, ժҪ = '����' || v_���۵�
    Where NO = v_�Һŵ� And ��¼���� = 4 And ��¼״̬ = 0;
  End If;

  If Nvl(n_�����¼id, 0) <> 0 Then
    Zl_ԤԼ�ҺŽ���_����_Insert_s(v_�Һŵ�, Null, n_����id, n_�����, v_����, v_�Ա�, v_����, v_���ʽ����, v_�ѱ�, n_��������, n_����id, n_���ʽ��, d_����ʱ��,
                          d_����ʱ��, v_����Ա����, v_����Ա����, n_�ҺŻ���, Null, 0, Null, v_���۵�, n_�Һ���Ŀid);
  Else
    Zl_ԤԼ�ҺŽ���_Insert_s(v_�Һŵ�, Null, n_����id, n_�����, v_����, v_�Ա�, v_����, v_���ʽ����, v_�ѱ�, n_��������, n_����id, n_���ʽ��, d_����ʱ��,
                       d_����ʱ��, v_����Ա����, v_����Ա����, n_�ҺŻ���, Null, 0, Null, v_���۵�, n_�Һ���Ŀid);
  End If;

  n_�Һ����ɶ��� := zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
  If Nvl(n_�Һ����ɶ���, 0) <> 0 Then
    n_�Һ����ɶ��� := 1;
  End If;
  Update ������ü�¼ Set ����ʱ�� = d_����ʱ�� Where ��¼���� = 4 And ��¼״̬ = 1 And NO = v_�Һŵ�;
  Update ���˹Һż�¼ Set ����ʱ�� = d_����ʱ�� Where ��¼״̬ = 1 And NO = v_�Һŵ�;

  If Nvl(n_�ҺŻ���, 0) = 1 Or v_��Ԥ������ids Is Not Null Then
    If v_��Ԥ������ids Is Not Null Then
      Zl_���˹Һ��շ�_Modify_s(v_�Һŵ�, n_����id, n_���ʽ�� || '|' || v_��Ԥ������ids, 3, 1);
    End If;
    Zl_���˹Һż�¼_��ɹҺ�_s(v_�Һŵ�, n_��Ժ, 2, n_�Һ����ɶ���);
  
    --ҽ��վ�Զ�����
    Update ������ü�¼ Set ִ���� = v_����Ա����, ִ��ʱ�� = d_����ʱ��, ִ��״̬ = 2 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = v_�Һŵ�;
    Update ���˹Һż�¼ Set ִ���� = v_����Ա����, ִ��ʱ�� = d_����ʱ��, ִ��״̬ = 2 Where NO = v_�Һŵ�;
    If Nvl(n_�Һ����ɶ���, 0) = 1 Then
      --���պ�,�������
      Update �ŶӽкŶ��� Set �Ŷ�״̬ = 2 Where ҵ������ = 0 And ҵ��id = n_�Һ�id;
    End If;
  
  Else
    o_Json := j_Json.Get_Pljson('balance_info');
    If o_Json Is Not Null Then
      n_���ʽ�� := o_Json.Get_Number('blnc_money');
      v_���㷽ʽ := o_Json.Get_String('blnc_mode');
      n_�����id := o_Json.Get_Number('cardtype_id');
      v_֧������ := o_Json.Get_String('pay_cardno');
    
      Zl_���˹Һ��շ�_Modify_s(v_�Һŵ�, n_����id, v_���㷽ʽ || ',' || n_���ʽ�� || ', , ', 1, 0, 0, n_Ԥ��id, n_�����id, v_֧������, Null, Null,
                         0, 1);
    End If;
  End If;

  Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_�Һ���Ŀid, n_ִ�в���id, d_����ʱ��, 2, v_�ű�, 0, n_�����¼id);

  Json_Out := '{"output":{"code":1,"message":"�ɹ�","blnc_id":' || Nvl(n_����id || '', 'null') || ',"relation_id":' ||
              Nvl(n_Ԥ��id || '', 'null') || '}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Rgstapptreceive;
/

Create Or Replace Procedure Zl_Exsesvr_Updrgstbalanceinfo
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  -------------------------------------------
  --���ܣ�ҽ��վԤԼ���ո��¹Һ�֧����Ϣ
  --��Σ�Json_In:��ʽ
  --input
  --  rgst_no            C  1 �Һ�no
  --  blnc_id            N  1 ����ID
  --  relation_id        N  1 ��������ID
  --  pati_inhospital    N  1 ��Ժ����
  --  totalmoney         N  1 ֧���ܽ��
  --  cardtype_id        N  1 ֧�������ID
  --  rgst_recv_time     C  1 ����ʱ��
  --  recharge           C    �쳣���½���
  --  operator_code      C    ����Ա����
  --  operator_name      C    ����Ա����
  --  balance_list[]
  --     blnc_mode       N  1 ���㷽ʽ
  --     swapmoney       C  1 ������
  --     swapno          C  1 ������ˮ��
  --     swapmemo        C  1 ����˵��
  --     blnc_no         C  1 �������
  --     blnc_memo       C  1 ����ժҪ
  --     card_no         C  1 ֧������
  --     cardtype_id     N    �����ID
  --  otherswap_list[]   C    ����������Ϣ
  --     swap_name       C  1  ��������
  --     swap_note       C  1  ��������
  -- ����:
  --  output
  --    code                            N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                         C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  -------------------------------------------
  n_�Һ�id     ���˹Һż�¼.Id%Type;
  v_�Һŵ�     ���˹Һż�¼.No%Type;
  d_����ʱ��   ���˹Һż�¼.����ʱ��%Type;
  v_����Ա���� ���˹Һż�¼.����Ա���%Type;
  v_����Ա���� ���˹Һż�¼.����Ա����%Type;
  n_������   ������ü�¼.���ʽ��%Type;
  n_����id     ������ü�¼.����id%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;
  v_�������   ����Ԥ����¼.�������%Type;
  v_����ժҪ   ����Ԥ����¼.ժҪ%Type;
  n_��������id ����Ԥ����¼.��������id%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  n_�����id   ����Ԥ����¼.�����id%Type;
  v_֧������   ����Ԥ����¼.����%Type;

  n_֧���ܽ��   ������ü�¼.���ʽ��%Type;
  n_�ϼƽ��     ������ü�¼.���ʽ��%Type;
  v_��������     �������㽻��.������Ŀ%Type;
  v_��������     �������㽻��.��������%Type;
  v_��չ��Ϣ     Varchar2(4000);
  n_����         Number(2);
  n_��ͨ����     Number(2);
  n_��Ժ         Number(2);
  n_�Һ����ɶ��� Number(2);
  n_��������     Number(2);
  n_��ɹҺ�     Number(2);
  o_Json         PLJson;
  j_Input        PLJson;
  j_Json         PLJson;

  j_Jsonlist Pljson_List := Pljson_List();

Begin
  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  v_�Һŵ�     := j_Json.Get_String('rgst_no');
  n_����id     := j_Json.Get_Number('blnc_id');
  n_��������id := j_Json.Get_Number('relation_id');
  n_��Ժ       := j_Json.Get_Number('pati_inhospital');
  n_֧���ܽ�� := j_Json.Get_Number('totalmoney');
  n_����       := j_Json.Get_Number('recharge');
  d_����ʱ��   := To_date(j_Json.Get_String('rgst_recv_time'), 'YYYY-MM-DD hh24:mi:ss');
  v_����Ա���� := j_Json.Get_String('operator_code');
  v_����Ա���� := j_Json.Get_String('operator_name');
  
  If Nvl(n_����, 0) = 1 Then
    zl_���˹Һż�¼_�쳣����_s(v_�Һŵ�, Null, d_����ʱ��, v_����Ա����, v_����Ա����);
  End If;
  
  n_�������� := 0;
  n_��ɹҺ� := 0;
  j_Jsonlist := j_Json.Get_Pljson_List('balance_list');
  If j_Jsonlist Is Not Null Then
    For I In 1 .. j_Jsonlist.Count Loop
      o_Json       := PLJson();
      o_Json       := PLJson(j_Jsonlist.Get(I));
      v_���㷽ʽ   := o_Json.Get_String('blnc_mode');
      n_������   := o_Json.Get_Number('swapmoney');
      v_������ˮ�� := o_Json.Get_String('swapno');
      v_����˵��   := o_Json.Get_String('swapmemo');
      v_�������   := o_Json.Get_String('blnc_no');
      v_����ժҪ   := o_Json.Get_String('blnc_memo');
      v_֧������   := o_Json.Get_String('card_no');
      n_�����id   := o_Json.Get_Number('cardtype_id');
    
      n_�ϼƽ�� := Nvl(n_�ϼƽ��, 0) + n_������;
      If I > 1 Then
        n_�������� := 1;
      End If;
      If I = j_Jsonlist.Count And n_֧���ܽ�� = n_�ϼƽ�� Then
        n_��ɹҺ� := 1;
      End If;
      If Nvl(n_�����id, 0) = 0 Then
        n_��ͨ���� := 1;
      End If;
      Zl_���˹Һ��շ�_Modify_s(v_�Һŵ�, n_����id, v_���㷽ʽ || ',' || n_������ || ',' || v_������� || ',' || v_����ժҪ, 1, n_��ɹҺ�, n_��������,
                         n_��������id, n_�����id, v_֧������, v_������ˮ��, v_����˵��, Nvl(n_��ͨ����, 0), 2);
    End Loop;
  End If;

  Begin
    n_�����id := j_Json.Get_Number('cardtype_id');
    j_Jsonlist := Pljson_List();
    j_Jsonlist := j_Json.Get_Pljson_List('otherswap_list');
    If j_Jsonlist Is Not Null Then
      For I In 1 .. j_Jsonlist.Count Loop
        o_Json     := PLJson();
        o_Json     := PLJson(j_Jsonlist.Get(I));
        v_�������� := o_Json.Get_String('swap_name');
        v_�������� := o_Json.Get_String('swap_note');
        v_��չ��Ϣ := v_��չ��Ϣ || '||' || v_�������� || '|' || v_�������� || '|' || v_��������;
      
        If Lengthb(v_��չ��Ϣ) > 2000 Then
          Zl_�������㽻��_Insert(n_�����id, 0, v_֧������, n_����id, v_��չ��Ϣ);
          v_��չ��Ϣ := Null;
        End If;
      End Loop;
    End If;
  
    If v_��չ��Ϣ Is Not Null Then
      Zl_�������㽻��_Insert(n_�����id, 0, v_֧������, n_����id, v_��չ��Ϣ);
    End If;
  Exception
    When Others Then
      Null;
  End;

  If n_��ɹҺ� = 1 Then
    n_�Һ����ɶ��� := zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If Nvl(n_�Һ����ɶ���, 0) <> 0 Then
      n_�Һ����ɶ��� := 1;
    End If;
  
    Zl_���˹Һż�¼_��ɹҺ�_s(v_�Һŵ�, n_��Ժ, 2, n_�Һ����ɶ���);
  
    Update ������ü�¼ Set ִ���� = v_����Ա����, ִ��ʱ�� = d_����ʱ��, ִ��״̬ = 2 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = v_�Һŵ�;
    Update ���˹Һż�¼ Set ִ���� = v_����Ա����, ִ��ʱ�� = d_����ʱ��, ִ��״̬ = 2 Where NO = v_�Һŵ�;
    If Nvl(n_�Һ����ɶ���, 0) = 1 Then
      --���պ�,�������
      Select Max(ID) Into n_�Һ�id From ���˹Һż�¼ Where NO = v_�Һŵ�;
      Update �ŶӽкŶ��� Set �Ŷ�״̬ = 2 Where ҵ������ = 0 And ҵ��id = n_�Һ�id;
    End If;
  End If;

  Json_Out := zlJsonOut('�ɹ�', 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Updrgstbalanceinfo;
/

Create Or Replace Procedure Zl_Exsesvr_Odr_Check
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҽ�����ڷ����շ��ûؼ������ݻ�ȡ
  --��Σ�Json_In:��ʽ
  --  input
  --    check_type                  N 1 ������ͣ�
  --                                ˵����1-���Ҫ�ջص�ҽ����Ӧ�ķ��ý������������ order_ids��ѯ
  --                                      10-��ʽ�����ջش�ʱorder_ids����һ��ҽ��id,���[�ջص���������]
  --    order_ids                   C 1 ҽ��IDƴ��
  -----------[�ջص���������]----check_type=10ʱ����------------------------------------------------------------------------------------------
  --    roll_num                    N 1 ����Ҫ�ջص�����check_type=10ʱ����
  --    fee_no                      N 1 ���õ�����Ϣ,��Ҫ�������ֵ�ǰ�ջ�ģ:null-��������; �������۵�-��ʾ�ǵ������۵�,���嵥�ݺ�-���Ǹ�������
  --    fee_nos                     C 1 �����ҽ�� order_ids ��Ӧ�ĵ��ݺ�,����ƴ��,��ʱ�� order_ids ֻ��һ��ֵ
  --    advice_dosage               N 1 ��������
  --    advice_note                 N 1 ҽ������
  --    clinic_type                 C 1 �������
  --    is_stuff_order                    N 1 ������������ҽ��
  --    price_list[]ҽ���Ƽ��б�
  --          order_id               N 1 ҽ��id
  --          fee_item_id            N 1 �շ�ϸĿid
  --          refer_num              N 1 ��������
  --          fee_way                N 1 �շѷ�ʽ����ͨ�շѷ�ʽΪ0������ȡ
  --    price_exe_list[]ҽ��ִ�мƼ������б�
  --          fee_item_id            N 1 �շ�ϸĿid
  --          roll_num               N 1 �ջ�����
  --    excute_list[]           ������ִ���б�(ҩƷ�����ķ���),��ʹ��ִ����Ϊ0ҲҪ����
  --          fee_id              N   1   ����ID
  --          sended_num          N   1   �ѷ�����
  --    advice_excute_list[]    ������ִ���б�(ҽ������),��ʹ��ִ����Ϊ0ҲҪ����
  --          advice_id           N   1   ҽ��ID
  --          fee_item_id         N   1   �շ�ϸĿID
  --          execute_num         N   1   ��ִ����
  --    pati_list[]             ������Ϣ���������Щ���˵ķ���
  --          pati_id             N   1   ����ID
  --          pati_name           C   1   ��������
  --          fee_audit_status    N   1   ������˱�־:0���-δ���;1-����˻�ʼ���(��ϲ���:������˷�ʽ������);2-������,��Ͻ���Ȩ��[��ֹδ��˲��˽���]���й������
  --          si_inp_status       N   1   סԺ״̬:0-����סԺ;1-��δ���;2-����ת�ƻ�����ת����;3-��Ԥ��Ժ
  --          catalog_date        C   0   ������Ŀ���ڣ�yyyy-mm-dd hh24:mi:ss

  --����: Json_Out,��ʽ����
  --  output
  --    code          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ

  --    charge_list[]���������б�
  --                  outpati_account        N 1 �������0-סԺ����,1-�������
  --                  fee_id                 N 1 ����id
  --                  fee_item_id            N 1 �շ�ϸĿid
  --                  request_dept_id        N 1 �������id
  --                  audit_dept_id          N 1 ��˿���id
  --                  request_num            N 1 ��������

  --    del_list[]����ɾ���б�
  --                  outpati_account        N 1 �������0-סԺ����,1-�������
  --                  fee_no                 C 1 ���õ��ݺ�
  --                  serial_num             C 1 ɾ�����,��Ÿ�ʽ:����:ִ��״̬

  --   del_drug[]ҩƷɾ���б�
  --                  rcpdtl_id              N 1 ������ϸid,Ŀǰ����ķ���ID
  --                  chargeoffs_num         N 1 ��������

  --   del_stuff[]����ɾ���б�
  --                  stuffdtl_id            N 1 ������ϸid,Ŀǰ����ķ���ID
  --                  return_num             N 1 ��������

  --   pivas_list[]���������б�
  --                  pivas_id               N 1 ��Һid
  --                  auto_aduit             N 1 �Ƿ��Զ���� 0-�����,1-Ҫ�Զ����
  --                  request_time           C 1 ����ʱ��
  --                  reason                 C 1 ����ԭ��

  --    roll_list[]���������б����
  --                  outpati_account        N 1 �������0-סԺ����,1-�������
  --                  clinic_type            C 1 ҽ���������
  --                  fee_no                 C 1 ���ݺ�
  --                  item_type              C 1 �շ�ϸĿ���
  --                  fee_id                 N 1 ����id
  --                  fee_id_old             N 1 ����id,ԭʼ����id
  --                  packages_num           N 1 ����
  --                  send_num               N 1 ����
  --                  is_stuff_order         N 1 �����Ƿ��ǰ󶨵����ķ���0-������ҽ��,1-����ҽ��
  --                  stuff_used             N 1 �Ƿ��Ǹ����������ŷ���
  --                  exe_status             N 1 ִ��״̬

  --     roll_drug_list[]�����ջ��б�ҩƷ
  --                  clinic_type            C 1 ҽ���������
  --                  rcp_no                 C 1 ������,���õ���
  --                  rcpdtl_id              N 1 ������ϸID
  --                  rcpdtl_id_old          N 1 ������ϸID,ԭʼ��ϸid
  --                  packages_num           N 1 ��ҩ����
  --                  send_num               N 1 ��ҩ����
  --                  item_type              C 1 �շ���Ŀ���

  --     roll_stuff_list[]�����ջ��б�����
  --                  clinic_type            C 1 ҽ���������
  --                  stuff_no               C 1 �������ʵĵ��ݺ�
  --                  stuffdtl_id            N 1 ������ϸid,���ü�¼id
  --                  stuffdtl_id_old        N 1 ԭʼ������ϸid,���ü�¼id
  --                  packages_num           N 1 ����
  --                  outbound_num           N 1 ����
  --                  is_stuff_order         N 1 �����Ƿ��ǰ󶨵����ķ���0-������ҽ��,1-����ҽ��

  ---------------------------------------------------------------------------

  --    price_list[]ҽ���Ƽ��б�
  Type t_Rs_�Ƽ� Is Record(
    ҽ��id     Number,
    �շ�ϸĿid Number,
    ��������   Number,
    �ջ�����   Number,
    �շѷ�ʽ   Number);
  Type t_�Ƽ� Is Table Of t_Rs_�Ƽ�;
  Rs_�Ƽ� t_�Ƽ�;
  Rs_�ջ� t_�Ƽ�;

  Type t_Rs_ִ�� Is Record(
    ҽ��id   Number,
    ����id   Number,
    �շ�id   Number,
    δִ���� Number,
    ��ִ���� Number);
  Type t_ִ�� Is Table Of t_Rs_ִ��;
  Rs_ִ�� t_ִ��; --ҩƷ�����Ѿ�ִ������

  Type t_Rs_���� Is Record(
    ҽ��id     Number,
    ����id     Number,
    �շ�ϸĿid Number,
    ����       Number,
    ��Һid     Number,
    ����״̬   Number);
  Type t_���� Is Table Of t_Rs_����;
  Rs_���� t_����;

  Type t_Rs_���� Is Record(
    ����id       Number(18),
    ���         Number,
    ����         Number,
    �������     Number,
    ����         Number,
    �շ�ϸĿid   Number,
    �������id   Number,
    ��˿���id   Number,
    ��ִ����     Number,
    ����ʱ��     Date,
    �������     Number,
    �����Զ���� Number,
    ��Һid       Number,
    NO           Varchar2(60),
    ���         Number, --  0-��ͨ,1-ҩƷ,2-����
    ҽ��id       Number);
  Type t_���� Is Table Of t_Rs_����;
  Rs_����    t_����;
  Rs_����    t_����;
  Rs_Dellist t_����;

  j_Json       PLJson;
  j_Tmp        PLJson;
  j_Output     PLJson;
  j_Item       PLJson;
  j_List       Pljson_List := Pljson_List();
  j_List_Order Pljson_List := Pljson_List();

  Lngtmp Number;
  --v_Dec                Number;
  v_Json_In   Varchar2(32767);
  v_Json_Out  Varchar2(32767);
  v_Item_List Varchar2(32767); --���+����
  --v_�������           Varchar2(322);
  v_Orderfeenos        Varchar2(32767); --ҽ����Ӧ�ķ��õ��ݺ�,����ƴ��,ҽ��ID+���ݺ� ����Ψһȷ��������
  v_Pati_List          Varchar2(32767);
  v_Chk��������        Varchar2(32767);
  v_Excute_List        Varchar2(32767);
  v_Advice_Excute_List Varchar2(32767);
  v_Charge_List        Varchar2(32767);
  v_Del_List           Varchar2(32767);
  v_ҽ������           Varchar2(4000);
  v_Del_Drug           Varchar2(32767);
  v_Del_Stuff          Varchar2(32767);
  v_�Զ�����           Varchar2(4000);

  �ջ���_In     Number;
  v_�շ�ϸĿid  Number;
  Nt_�շ�ϸĿid Number;
  Nt_��������   Number;
  Nt_����ԭ��   Varchar2(4000);
  Nt_���͵���   Varchar2(200);
  Nt_��������   Number;
  Nt_�Զ����   Number;
  Nt_�շѷ�ʽ   Number;
  Nt_��������   Number;
  Nt_�������   Varchar2(30);
  Nt_�շѱ�־   Number;
  Nt_ҽ������   Varchar2(4000);
  v_�շ�����    Varchar2(4000);
  v_����ids     Varchar2(32767);
  v_�������    Varchar2(32767);
  v_����id      Number(18);
  v_��ǰ����    Number;
  v_ʣ������    Number;
  v_���ʽ��    Number;
  v_�ջ�ʣ��    Number;
  v_�ջ�����    Number;
  v_�ջ�����tmp Number;
  v_�ջ���      Number;
  v_��������    Number;
  v_����ϵ��    ҩƷ���.����ϵ��%Type;
  v_סԺ��װ    ҩƷ���.סԺ��װ%Type;
  v_���ʲ���    Varchar2(4000);
  No_In         Varchar2(2000);
  v_����        Number;
  v_Tmp         Varchar2(32767);
  v_Pivas_Ids   Varchar2(32767);
  v_Pivas_Out   Varchar2(32767);
  v_Roll_List   Varchar2(32767);
  v_Roll_List_d Varchar2(32767); --�����ջص�ҩƷ
  v_Roll_List_s Varchar2(32767); --�����ջص�����
  n_��鷽ʽ    Number;
  l_ҽ��ids     t_StrList;
  c_ҽ��ids     Clob;
  Vo_Vals       Clob;
  v_Error       Varchar2(255);
  b_In          Clob;
  b_Out         Clob;
  �ջ�ʱ��_In   Date;
  Err_Custom Exception;
  ҽ��id_In Number;

  n_�������ʲ��� Number;

  --��������ʱ��,�������,��������,�����Զ����,����ģʽ

  --����ṹ
  Cursor c_Fee_List_Type Is
    Select a.Id ����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.ִ��״̬ ��������, a.�շ���� �������, a.ժҪ ҽ������, a.���� ��������, a.���� ʣ������,
           a.���� ��ִ����, a.���� δִ����, a.ִ��״̬ ִ�б�־, a.��¼״̬, a.����ʱ�� �Ǽ�ʱ��, a.ִ��״̬ �շѷ�ʽ, a.ִ��״̬ �������, a.����ʱ�� ��������ʱ��, a.ִ��״̬ �������,
           a.���� ��������, a.ִ��״̬ �����Զ����, a.ִ��״̬ ����ģʽ, a.ִ�в���id, a.Id ��Һid
    From סԺ���ü�¼ A
    Where 0 = 1;
  r_Detail c_Fee_List_Type%RowType;

  --ֱ�����ʵ����
  Cursor c_Detail Is
    Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, Null �������, Null ҽ������, -null ��������, a.ʣ������, -null ��ִ����,
           -null δִ����, a.ִ�б�־, a.��¼״̬, a.�Ǽ�ʱ��, -null �շѷ�ʽ, a.�������
    From (Select 0 As �������, Max(Decode(b.��¼״̬, 2, 0, b.Id)) As ����id, b.No, Nvl(b.�۸񸸺�, b.���) As ���, b.�շ�ϸĿid, b.���˲���id,
                  Sum(Nvl(b.����, 1) * b.����) As ʣ������, b.�շ����, Max(Nvl(b.ִ��״̬, 0)) As ִ�б�־, d.��������, Max(b.��¼״̬) As ��¼״̬,
                  Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��
           From סԺ���ü�¼ B, �������� D
           Where b.ҽ����� = ҽ��id_In And b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.����id(+) And
                 Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0
           Group By b.No, b.��¼����, Nvl(b.�۸񸸺�, b.���), b.�շ�ϸĿid, b.���˲���id, b.�շ����, d.��������
           Having Sum(Nvl(b.����, 1) * b.����) > 0
           Union All
           Select 1 As �������, Max(Decode(b.��¼״̬, 2, 0, b.Id)) As ����id, b.No, Nvl(b.�۸񸸺�, b.���) As ���, b.�շ�ϸĿid, b.���˲���id,
                  Sum(Nvl(b.����, 1) * b.����) As ʣ������, b.�շ����, Max(Nvl(b.ִ��״̬, 0)) As ִ�б�־, d.��������, Max(b.��¼״̬) As ��¼״̬,
                  Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��
           From ������ü�¼ B, �������� D
           Where b.ҽ����� = ҽ��id_In And b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.����id(+) And
                 Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0
           Group By b.No, b.��¼����, Nvl(b.�۸񸸺�, b.���), b.�շ�ϸĿid, b.���˲���id, b.�շ����, d.��������
           Having Sum(Nvl(b.����, 1) * b.����) > 0) A
    Order By a.�շ�ϸĿid, a.ִ�б�־, a.�Ǽ�ʱ�� Desc;

  --δ��Ч�����ʻ��۵�
  Cursor c_Del Is
    Select b.Id As ����id, b.No, b.���, b.�շ����, b.�շ�ϸĿid, Nvl(b.����, 1) * b.���� As ʣ������, d.��������, b.�������
    From (Select 0 As �������, a.Id, a.No, a.���, a.�շ�ϸĿid, a.����, a.����, a.�۸񸸺�, a.ҽ�����, a.�շ����
           From סԺ���ü�¼ A
           Where a.ҽ����� = ҽ��id_In And a.��¼״̬ = 0
           Union All
           Select 1 As �������, a.Id, a.No, a.���, a.�շ�ϸĿid, a.����, a.����, a.�۸񸸺�, a.ҽ�����, a.�շ����
           From ������ü�¼ A
           Where a.ҽ����� = ҽ��id_In And a.��¼״̬ = 0) B, �������� D
    Where b.ҽ����� = ҽ��id_In And b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.����id(+) And
          Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0
    Order By b.�շ�ϸĿid, b.No Desc;

  --���������������Ǿ���
  Cursor c_Negdrug Is
    Select b.Id As ����id, b.No, b.���, b.�շ����, b.�շ�ϸĿid, b.���˲���id, Nvl(b.����, 1) * b.���� As ʣ������, b.�������, b.��¼״̬,
           b.ִ��״̬ As ִ�б�־, d.��������, b.ִ�в���id
    From (Select 0 As �������, a.Id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.����, a.����, a.�۸񸸺�, a.ҽ�����, a.ִ��״̬, a.��¼״̬, a.�շ����,
                  a.ִ�в���id
           From סԺ���ü�¼ A
           Where a.ҽ����� = ҽ��id_In And a.��¼״̬ In (0, 1, 3)
           Union All
           Select 1 As �������, a.Id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.����, a.����, a.�۸񸸺�, a.ҽ�����, a.ִ��״̬, a.��¼״̬, a.�շ����,
                  a.ִ�в���id
           From ������ü�¼ A
           Where a.ҽ����� = ҽ��id_In And a.��¼״̬ In (0, 1, 3)) B, �������� D
    Where Instr(',' || v_Orderfeenos || ',', ',' || b.No || ',') > 0 And b.�շ�ϸĿid = d.����id(+)
    Order By b.�շ�ϸĿid, b.No Desc;

  --������ҩ����(����ҩ;��)����ʱ�������ķ���(����������ж�����¼)
  --�Է�ҩҽ��,ֱ���ջ�ָ����,���ܶ�η���(�����η��ͼ۸�ͬ,���ջصļ۸��������εģ���Ȼ��Ҫ���ݶ���������μ��ջ���)��
  --���ı������ۼ۵�λ������סԺ��λת��
  --��ҩ��������д�˷��ͼ�¼(�����˶���������ȼ�)
  --һ��ֻ��һ�λ�һ�η���ֻ��һ�ε���Ŀ��ʱ��֧�ָ�������
  Cursor c_Other Is
    Select a.�������, a.No, a.���, a.����id, a.ʣ������, a.�շ�ϸĿid, a.���˲���id, a.��¼״̬, a.ִ�б�־, a.��������, a.�շѷ�ʽ, a.�շ����, d.��������,
           a.ִ�в���id
    From (Select 0 As �������, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.Id As ����id, a.���� As ʣ������, Nvl(a.ִ��״̬, 0) As ִ�б�־,
                  a.ҽ�����, Null ���ͺ�, Null ��������, Null �շѷ�ʽ, a.�շ����, a.ִ�в���id
           From סԺ���ü�¼ A
           Where a.No = Nt_���͵��� And a.��¼״̬ In (0, 1, 3) And a.ҽ����� + 0 = ҽ��id_In
           Union All
           Select 1 As �������, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.Id As ����id, a.���� As ʣ������, Nvl(a.ִ��״̬, 0) As ִ��״̬,
                  a.ҽ�����, Null ���ͺ�, Null ��������, Null �շѷ�ʽ, a.�շ����, a.ִ�в���id
           From ������ü�¼ A
           Where a.No = Nt_���͵��� And a.��¼״̬ In (0, 1, 3) And a.ҽ����� + 0 = ҽ��id_In) A, �������� D
    Where a.�շ�ϸĿid = d.����id(+)
    Order By a.�շ�ϸĿid, a.���, a.��¼״̬;

  Procedure p_Add_Negbill As
    P����     Number;
    P�Զ����� Number;
  Begin
  
    If Nt_������� = '7' Then
      -- ��ҩҽ�����䷽ʽ������ԭ;
      P����             := �ջ���_In;
      r_Detail.�������� := r_Detail.�������� / P����;
    Else
      P���� := 1;
    End If;
  
    --���ڷ�ҩƷ����ҽ����,Ҫ�Է��ý����Զ����,ҩƷ������Ҫ�Զ�����
    --������������
    Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
  
    v_Roll_List := v_Roll_List || ',{"outpati_account":' || r_Detail.�������; --0-סԺ����,1-�������
    v_Roll_List := v_Roll_List || ',"clinic_type":"' || Nt_������� || '"'; --ҽ�����������
    v_Roll_List := v_Roll_List || ',"fee_no":"' || No_In || '"';
    v_Roll_List := v_Roll_List || ',"item_type":"' || r_Detail.�շ���� || '"';
    v_Roll_List := v_Roll_List || ',"fee_id":' || v_����id;
    v_Roll_List := v_Roll_List || ',"fee_id_old":' || r_Detail.����id;
    v_Roll_List := v_Roll_List || ',"packages_num":' || zlJsonStr(P����, 1);
    v_Roll_List := v_Roll_List || ',"send_num":' || zlJsonStr(r_Detail.��������, 1);
    v_Roll_List := v_Roll_List || ',"is_stuff_order":' || Nvl(Nt_��������, 0); --�Ƿ��Ǹ������õ�ҽ��
    v_Roll_List := v_Roll_List || ',"stuff_used":' || Nvl(r_Detail.��������, 0); --�Ƿ��Ǹ����������ķ���
    v_Roll_List := v_Roll_List || ',"exe_status":' || Nvl(r_Detail.ִ�б�־, 0);
    v_Roll_List := v_Roll_List || ',"exe_deptid":' || r_Detail.ִ�в���id; --�⼸����㽫�������ڼ���ɱ���,��ʱδ��
    v_Roll_List := v_Roll_List || ',"fee_item_id":' || r_Detail.�շ�ϸĿid;
    v_Roll_List := v_Roll_List || ',"charge_tag":' || Nvl(Nt_�շѱ�־, 0);
    v_Roll_List := v_Roll_List || '}';
  
    If r_Detail.�շ���� In ('5', '6', '7') Then
      --�Ƿ��շ�
      v_Roll_List_d := v_Roll_List_d || ',{"clinic_type":"' || Nt_������� || '"'; --ҽ�����������
      v_Roll_List_d := v_Roll_List_d || ',"rcp_no":"' || No_In || '"';
      v_Roll_List_d := v_Roll_List_d || ',"rcpdtl_id":' || v_����id;
      v_Roll_List_d := v_Roll_List_d || ',"rcpdtl_id_old":' || r_Detail.����id;
      v_Roll_List_d := v_Roll_List_d || ',"packages_num":' || zlJsonStr(P����, 1);
      v_Roll_List_d := v_Roll_List_d || ',"send_num":' || zlJsonStr(r_Detail.��������, 1);
      v_Roll_List_d := v_Roll_List_d || ',"charge_tag":' || Nvl(Nt_�շѱ�־, 0);
      v_Roll_List_d := v_Roll_List_d || '}';
    End If;
  
    If r_Detail.�������� = 1 Then
      If v_�Զ����� = '1' And r_Detail.ִ�б�־ = 1 Then
        P�Զ����� := 1;
      End If;
      --�Ƿ��շ�
      v_Roll_List_s := v_Roll_List_s || ',{"clinic_type":"' || Nt_������� || '"'; --ҽ�����������
      v_Roll_List_s := v_Roll_List_s || ',"stuff_no":"' || No_In || '"';
      v_Roll_List_s := v_Roll_List_s || ',"stuffdtl_id":' || v_����id;
      v_Roll_List_s := v_Roll_List_s || ',"stuffdtl_id_old":' || r_Detail.����id;
      v_Roll_List_s := v_Roll_List_s || ',"packages_num":' || zlJsonStr(P����, 1);
      v_Roll_List_s := v_Roll_List_s || ',"outbound_num":' || zlJsonStr(r_Detail.��������, 1);
      v_Roll_List_s := v_Roll_List_s || ',"is_stuff_order":' || Nvl(Nt_��������, 0);
      v_Roll_List_s := v_Roll_List_s || ',"stuff_auto_send":' || Nvl(P�Զ�����, 0);
      v_Roll_List_s := v_Roll_List_s || ',"charge_tag":' || Nvl(Nt_�շѱ�־, 0);
      v_Roll_List_s := v_Roll_List_s || '}';
    End If;
  End;

  Procedure p_Getoutlist As
    --�����б���װ
    P���    Varchar2(32767);
    Pno      Varchar2(30);
    p_Deltmp Varchar2(32767);
  Begin
    If v_Charge_List Is Not Null Then
      v_Json_Out := v_Json_Out || ',"charge_list":[' || Substr(v_Charge_List, 2) || ']';
    End If;
  
    If v_Del_List Is Not Null Then
      v_Del_List := Null;
      --��Ҫ������
      For I In 1 .. Rs_Dellist.Count Loop
        If Pno Is Not Null And Pno <> Rs_Dellist(I).No Then
          p_Deltmp := '{"outpati_account":' || Rs_Dellist(I - 1).�������;
          p_Deltmp := p_Deltmp || ',"fee_no":"' || Rs_Dellist(I - 1).No || '"';
          p_Deltmp := p_Deltmp || ',"serial_num":"' || Substr(P���, 2) || '"';
          p_Deltmp := p_Deltmp || '}';
        
          If v_Del_List Is Null Then
            v_Del_List := p_Deltmp;
          Else
            v_Del_List := p_Deltmp || ',' || v_Del_List;
          End If;
          P��� := Null;
        End If;
        Pno   := Rs_Dellist(I).No;
        P��� := P��� || ',' || Rs_Dellist(I).��� || ':' || Rs_Dellist(I).���� || ':0';
      End Loop;
    
      p_Deltmp := '{"outpati_account":' || Rs_Dellist(Rs_Dellist.Count).�������;
      p_Deltmp := p_Deltmp || ',"fee_no":"' || Rs_Dellist(Rs_Dellist.Count).No || '"';
      p_Deltmp := p_Deltmp || ',"serial_num":"' || Substr(P���, 2) || '"';
      p_Deltmp := p_Deltmp || '}';
    
      If v_Del_List Is Null Then
        v_Del_List := ',' || p_Deltmp;
      Else
        v_Del_List := ',' || p_Deltmp || ',' || v_Del_List;
      End If;
    
      v_Json_Out := v_Json_Out || ',"del_list":[' || Substr(v_Del_List, 2) || ']';
    
    End If;
    If v_Del_Drug Is Not Null Then
      v_Json_Out := v_Json_Out || ',"del_drug":[' || Substr(v_Del_Drug, 2) || ']';
    End If;
    If v_Del_Stuff Is Not Null Then
      v_Json_Out := v_Json_Out || ',"del_stuff":[' || Substr(v_Del_Stuff, 2) || ']';
    End If;
    If v_Roll_List Is Not Null Then
      v_Json_Out := v_Json_Out || ',"roll_list":[' || Substr(v_Roll_List, 2) || ']';
    End If;
  
    If v_Roll_List_d Is Not Null Then
      v_Json_Out := v_Json_Out || ',"roll_drug_list":[' || Substr(v_Roll_List_d, 2) || ']';
    End If;
  
    If v_Roll_List_s Is Not Null Then
      v_Json_Out := v_Json_Out || ',"roll_stuff_list":[' || Substr(v_Roll_List_s, 2) || ']';
    End If;
  
    If v_Pivas_Out Is Not Null Then
      v_Json_Out := v_Json_Out || ',"pivas_list":[' || Substr(v_Pivas_Out, 2) || ']';
    End If;
  End;

  Procedure p_Add_Delitem As
    --���Ҫɾ���ķ����к����������б�Ԫ��
  Begin
    Rs_����.Extend;
    Lngtmp := Rs_����.Count;
    Rs_����(Lngtmp).����id := r_Detail.����id;
    Rs_����(Lngtmp).No := r_Detail.No;
    Rs_����(Lngtmp).��� := r_Detail.���;
    Rs_����(Lngtmp).���� := r_Detail.��������;
    Rs_����(Lngtmp).������� := r_Detail.�������;
    Rs_����(Lngtmp).��Һid := r_Detail.��Һid;
    If r_Detail.�շ���� In ('5', '6', '7') Then
      Rs_����(Lngtmp).��� := 1;
    Elsif r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
      Rs_����(Lngtmp).��� := 2;
    Else
      Rs_����(Lngtmp).��� := 0;
    End If;
    Rs_����(Lngtmp).���� := r_Detail.����ģʽ;
    If Rs_����(Lngtmp).���� = 1 Then
      Rs_����(Lngtmp).�շ�ϸĿid := r_Detail.�շ�ϸĿid;
      Rs_����(Lngtmp).�������id := r_Detail.���˲���id;
      Rs_����(Lngtmp).�����Զ���� := r_Detail.�����Զ����;
      Rs_����(Lngtmp).������� := r_Detail.�������;
      Rs_����(Lngtmp).��ִ���� := r_Detail.��ִ����;
      Rs_����(Lngtmp).����ʱ�� := r_Detail.��������ʱ��;
    End If;
  End;

  Procedure p_Delbill_Check(Prownum Number) As
    --ֱ�ӿ���ɾ���ĵ��ݼ��
    Rp    Number;
    Phave Number := 0;
  Begin
    Rp          := Prownum;
    v_Item_List := '{"serial_num":' || Rs_����(Rp).���;
    v_Item_List := v_Item_List || ',"quantity":' || zlJsonStr(Rs_����(Rp).����, 1);
    v_Item_List := v_Item_List || '}';
    v_Json_Out  := Null;
    If Rs_����(Rp).������� = 1 Then
      v_Json_In := '{"fee_no":"' || Rs_����(Rp).No || '"';
      v_Json_In := v_Json_In || ',"fee_bill_type":2';
      v_Json_In := v_Json_In || ',"item_list":[' || v_Item_List || ']';
      v_Json_In := v_Json_In || v_Excute_List;
      v_Json_In := v_Json_In || v_Advice_Excute_List;
      v_Json_In := v_Json_In || '}';
      v_Json_In := '{"input":' || v_Json_In || '}';
      Zl_������ʼ�¼_Delete_Check(v_Json_In, v_Json_Out);
    Else
      v_Json_In := '{"fee_no":"' || Rs_����(Rp).No || '"';
      v_Json_In := v_Json_In || ',"fee_bill_type":2';
      v_Json_In := v_Json_In || ',"balance_ban_writeoffs":0';
      v_Json_In := v_Json_In || ',"part_ban_writeoffs":0';
      v_Json_In := v_Json_In || ',"item_list":[' || v_Item_List || ']';
      v_Json_In := v_Json_In || v_Excute_List;
      v_Json_In := v_Json_In || v_Advice_Excute_List;
      v_Json_In := v_Json_In || v_Pati_List;
      v_Json_In := v_Json_In || '}';
      v_Json_In := '{"input":' || v_Json_In || '}';
      Zl_סԺ���ʼ�¼_Delete_Check(v_Json_In, v_Json_Out);
    End If;
  
    j_Tmp    := PLJson();
    j_Output := PLJson();
    j_Tmp    := PLJson(v_Json_Out);
    j_Output := j_Tmp.Get_Pljson('output');
    If j_Output.Get_Number('code') = 0 Then
      v_Error := j_Output.Get_String('message');
      Raise Err_Custom;
    End If;
    j_List := Pljson_List();
    j_List := j_Output.Get_Pljson_List('item_list');
    j_Tmp  := PLJson();
    j_Tmp  := PLJson(j_List.Get(1));
  
    --����
    Lngtmp := j_Tmp.Get_Number('quantity');
  
    For I In 1 .. Rs_Dellist.Count Loop
      If Rs_Dellist(I).No = Rs_����(Rp).No And Rs_Dellist(I).������� = Nvl(Rs_����(Rp).�������, 0) And Rs_Dellist(I)
         .��� = j_Tmp.Get_Number('serial_num') Then
        Rs_Dellist(I).���� := Rs_Dellist(I).���� + j_Tmp.Get_Number('quantity');
        Phave := 1;
      End If;
    End Loop;
  
    If Phave <> 1 Then
      Rs_Dellist.Extend;
      Rs_Dellist(Rs_Dellist.Count).No := Rs_����(Rp).No;
      Rs_Dellist(Rs_Dellist.Count).������� := Nvl(Rs_����(Rp).�������, 0);
      Rs_Dellist(Rs_Dellist.Count).��� := j_Tmp.Get_Number('serial_num');
      Rs_Dellist(Rs_Dellist.Count).���� := j_Tmp.Get_Number('quantity');
    End If;
  
    v_Tmp := j_Tmp.Get_Number('serial_num');
    v_Tmp := v_Tmp || ':' || j_Tmp.Get_Number('quantity');
    v_Tmp := v_Tmp || ':' || j_Tmp.Get_Number('execute_tag');
  
    --����ɾ���б�Ҫ����
    v_Del_List := v_Del_List || ',{"outpati_account":' || Nvl(Rs_����(Rp).�������, 0);
    v_Del_List := v_Del_List || ',"fee_no":"' || Rs_����(Rp).No || '"';
    v_Del_List := v_Del_List || ',"serial_num":"' || v_Tmp || '"';
    v_Del_List := v_Del_List || '}';
    v_Del_List := '��ɾ��';
    --ҩƷ�б�
    If Rs_����(Rp).��� = 1 Then
      v_Del_Drug := v_Del_Drug || ',{"rcpdtl_id":' || Rs_����(Rp).����id;
      v_Del_Drug := v_Del_Drug || ',"chargeoffs_num":' || zlJsonStr(Lngtmp, 1);
      If Rs_����(Rp).��Һid Is Not Null Then
        v_Del_Drug := v_Del_Drug || ',"pivas_id":' || Rs_����(Rp).��Һid;
      End If;
      v_Del_Drug := v_Del_Drug || '}';
    End If;
  
    --�����б�
    If Rs_����(Rp).��� = 2 Then
      v_Del_Stuff := v_Del_Stuff || ',{"stuffdtl_id":' || Rs_����(Rp).����id;
      v_Del_Stuff := v_Del_Stuff || ',"return_num":' || zlJsonStr(Lngtmp, 1);
      v_Del_Stuff := v_Del_Stuff || '}';
    End If;
  End;

  Procedure p_Charge_Check(Prownum Number) As
    --����������
  Begin
    Lngtmp        := Prownum;
    v_Chk�������� := ',{"fee_id":' || Rs_����(Lngtmp).����id;
    v_Chk�������� := v_Chk�������� || ',"fee_item_id":' || Rs_����(Lngtmp).�շ�ϸĿid;
    v_Chk�������� := v_Chk�������� || ',"request_dept_id":' || Rs_����(Lngtmp).�������id;
    v_Chk�������� := v_Chk�������� || ',"audit_dept_id":0'; --��˲��Ų�ȷ����0ͨ����鷽����ȷ��
    v_Chk�������� := v_Chk�������� || ',"request_type":' || Nvl(Rs_����(Lngtmp).�������, 0);
    v_Chk�������� := v_Chk�������� || ',"request_num":' || zlJsonStr(Rs_����(Lngtmp).����, 1);
    v_Chk�������� := v_Chk�������� || ',"sended_num":' || zlJsonStr(Rs_����(Lngtmp).��ִ����, 1);
    v_Chk�������� := v_Chk�������� || '}';
  
    v_Tmp := '{"input":{';
    v_Tmp := v_Tmp || '"item_list":[' || Substr(v_Chk��������, 2) || ']';
    v_Tmp := v_Tmp || v_Pati_List;
    v_Tmp := v_Tmp || '}}';
    b_In  := v_Tmp;
    Zl_���˷�������_Insert_Check(b_In, b_Out);
    j_Tmp    := PLJson();
    j_Output := PLJson();
    j_Tmp    := PLJson(b_Out);
    j_Output := j_Tmp.Get_Pljson('output');
    If j_Output.Get_Number('code') = 0 Then
      v_Error := j_Output.Get_String('message');
      Raise Err_Custom;
    End If;
    j_List := Pljson_List();
    j_List := j_Output.Get_Pljson_List('item_list');
    j_Tmp  := PLJson();
    j_Tmp  := PLJson(j_List.Get(1));
  
    Rs_����(Lngtmp).��˿���id := j_Tmp.Get_Number('audit_dept_id');
    Nt_�Զ���� := 0;
    --v_����:�����ж�
    If Rs_����(Lngtmp).��˿���id = Rs_����(Lngtmp).�������id And (v_���� = 1 Or Rs_����(Lngtmp).�����Զ���� = 1) Then
    
      v_Tmp := '{"fee_id":' || Rs_����(Lngtmp).����id;
      v_Tmp := v_Tmp || ',"stuff_auto_return":' || 0;
      v_Tmp := v_Tmp || ',"request_time":""';
      v_Tmp := v_Tmp || ',"request_type":' || Nvl(Rs_����(Lngtmp).�������, 0);
      v_Tmp := v_Tmp || ',"sended_num":' || zlJsonStr(Rs_����(Lngtmp).��ִ����, 1);
      v_Tmp := v_Tmp || '}';
    
      v_Json_In := '{"input":{"no_consistence":1';
      v_Json_In := v_Json_In || ',"item_list":[' || v_Tmp || ']';
      v_Json_In := v_Json_In || v_Pati_List;
      v_Json_In := v_Json_In || '}}';
    
      Zl_���˷�������_Audit_Check(v_Json_In, v_Json_Out);
    
      j_Tmp    := PLJson();
      j_Output := PLJson();
      j_Tmp    := PLJson(v_Json_Out);
      j_Output := j_Tmp.Get_Pljson('output');
      If j_Output.Get_Number('code') = 0 Then
        v_Error := j_Output.Get_String('message');
        Raise Err_Custom;
      End If;
      Nt_�Զ���� := 1;
      Rs_����(Lngtmp).���� := 0;
    End If;
  
    --���������б�
    v_Charge_List := v_Charge_List || ',{"outpati_account":' || Rs_����(Lngtmp).�������;
    v_Charge_List := v_Charge_List || ',"fee_id":' || Rs_����(Lngtmp).����id;
    v_Charge_List := v_Charge_List || ',"fee_item_id":' || Rs_����(Lngtmp).�շ�ϸĿid;
    v_Charge_List := v_Charge_List || ',"request_dept_id":' || Rs_����(Lngtmp).�������id;
    v_Charge_List := v_Charge_List || ',"audit_dept_id":' || Nvl(Rs_����(Lngtmp).��˿���id || '', 'null');
    v_Charge_List := v_Charge_List || ',"request_type":' || Nvl(Rs_����(Lngtmp).�������, 0);
    v_Charge_List := v_Charge_List || ',"request_num":' || zlJsonStr(Rs_����(Lngtmp).����, 1);
    v_Charge_List := v_Charge_List || ',"auto_aduit":' || Nt_�Զ����;
    v_Charge_List := v_Charge_List || ',"request_time":"' || To_Char(Rs_����(Lngtmp).����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
    v_Charge_List := v_Charge_List || ',"reason":"' || zlJsonStr(Nt_����ԭ��) || '"';
    v_Charge_List := v_Charge_List || '}';
  End;

  Procedure p_Get������Ϣ(P�ջ����� Out Number) As
    --��Ҫ��һ��ȷ����ֵ:v_����ϵ��,v_סԺ��װ,r_Detail.�շѷ�ʽ,r_Detail.�������,r_Detail.��������,p�ջ�����,v_��������
    --�÷����ڲ������Щ��������һ�θ�ֵ v_����ϵ��,v_סԺ��װ,�շѷ�ʽ,�������,p�ջ�����,v_��������
  Begin
    r_Detail.������� := Nt_�������;
    r_Detail.�������� := Nt_��������;
    r_Detail.ҽ������ := Nt_ҽ������;
    --ҩƷ�ջ�������������͹��Ϊ׼����ģ��Դ˼�����ջ��ۼ�����
    Begin
      Select ����ϵ��, סԺ��װ Into v_����ϵ��, v_סԺ��װ From ҩƷ��� Where ҩƷid = r_Detail.�շ�ϸĿid;
    Exception
      When Others Then
        v_����ϵ�� := 1;
        v_סԺ��װ := 1;
    End;
  
    --��ҽ���Ƽ��л�ȡ�շѷ�ʽ�Ͷ�������
    v_��������        := Null;
    r_Detail.�շѷ�ʽ := Null;
    For Lngtmp In 1 .. Rs_�Ƽ�.Count Loop
      If Rs_�Ƽ�(Lngtmp).�շ�ϸĿid = r_Detail.�շ�ϸĿid And Rs_�Ƽ�(Lngtmp).ҽ��id = ҽ��id_In Then
        v_��������        := Rs_�Ƽ�(Lngtmp).��������;
        r_Detail.�շѷ�ʽ := Rs_�Ƽ�(Lngtmp).�շѷ�ʽ;
      End If;
    End Loop;
    v_��������        := Nvl(v_��������, 1);
    r_Detail.�շѷ�ʽ := Nvl(r_Detail.�շѷ�ʽ, 0);
    r_Detail.��ִ���� := 0;
    If r_Detail.�շ���� In ('5', '6', '7') Or r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
      For Rl In 1 .. Rs_ִ��.Count Loop
        If Rs_ִ��(Rl).����id = r_Detail.����id And Rs_ִ��(Rl).ҽ��id = ҽ��id_In Then
          r_Detail.��ִ���� := Rs_ִ��(Rl).��ִ����;
        End If;
      End Loop;
    End If;
  
    --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
    If r_Detail.�շѷ�ʽ = 0 Then
      If r_Detail.������� = '7' Then
        --��ҩ�䷽ҩƷ������*����
        P�ջ����� := Round(�ջ���_In * r_Detail.�������� / Nvl(v_����ϵ��, 1), 5);
      Else
        P�ջ����� := Round(�ջ���_In * Nvl(v_סԺ��װ, 1), 5) * v_��������;
      End If;
    Else
      P�ջ����� := 0;
      For Lngtmp In 1 .. Rs_�ջ�.Count Loop
        If Rs_�ջ�(Lngtmp).�շ�ϸĿid = r_Detail.�շ�ϸĿid And Rs_�ջ�(Lngtmp).ҽ��id = ҽ��id_In Then
          P�ջ����� := Rs_�ջ�(Lngtmp).�ջ�����;
        End If;
      End Loop;
      P�ջ����� := Round(P�ջ�����, 5);
    End If;
  End;

  Procedure p_���������м�����
  (
    Pִ�б�־ Number,
    P����     Number
  ) As
  Begin
    --�����м����ݿ�ͨ�� r_Detail ���ͻ�ȡ
    --ϵͳ��������ִ�к��Ƿ���˻��۵������ԣ���ִ�е���Ȼ�����ǻ��۵�
    If Pִ�б�־ = 0 And r_Detail.��¼״̬ = 0 Then
      r_Detail.�������� := P����;
      r_Detail.����ģʽ := 0;
      --��������ʱ��,�������,��������,�����Զ����,����ģʽ
      p_Add_Delitem;
    Else
      If Not (r_Detail.�շ���� = '7' And Pִ�б�־ <> 0) Then
        r_Detail.�������� := P����;
        r_Detail.����ģʽ := 1;
        r_Detail.ִ�б�־ := Pִ�б�־;
      
        If Nvl(No_In, '�������۵�') <> '�������۵�' And Nvl(n_�������ʲ���, 0) = 0 Then
          p_Add_Negbill;
        Else
          p_Add_Delitem;
        End If;
      End If;
    End If;
  End;

  Procedure p_����ִ�в������(P�ջ��� Number) As
    Lngִ������ Number;
    Lngδִ���� Number;
  Begin
  
    Lngִ������ := Nvl(r_Detail.��ִ����, 0);
    --��������ʱ��,�������,��������,�����Զ����,����ģʽ
    If Nvl(Lngִ������, 0) <= 0 Then
      --����δִ��,�������
      r_Detail.������� := 0;
      p_���������м�����(0, P�ջ���);
    Else
      Lngδִ���� := Nvl(r_Detail.ʣ������, 0) - Nvl(Lngִ������, 0);
      If Lngδִ���� <= 0 Then
        r_Detail.������� := 1;
        p_���������м�����(1, P�ջ���);
      Else
        --��ִ��������ȷ��õ�ʣ��������,�����������ȫ���Ѿ�ִ��,�Ѿ�ִ����Ҳ����Ϊ����Ϊ����Ҳ������δִ��
        If Lngδִ���� >= P�ջ��� Then
          --����δִ��,�������
          r_Detail.������� := 0;
          p_���������м�����(0, P�ջ���);
        Else
          p_���������м�����(0, Lngδִ����);
          r_Detail.������� := 1;
          p_���������м�����(1, P�ջ��� - Lngδִ����);
        End If;
      End If;
    End If;
  End;

  Procedure p_��������
  (
    P�ջ��� Number,
    P��     Out Number
  ) As
    p_�ջ�ʱ�� Date;
    P��������  Number;
    P��������  Number;
  Begin
    P��������  := 0;
    P��        := 0;
    p_�ջ�ʱ�� := �ջ�ʱ��_In;
    For R In 1 .. Rs_����.Count Loop
      If Rs_����(R).����id = r_Detail.����id And Rs_����(R).�շ�ϸĿid = r_Detail.�շ�ϸĿid And Rs_����(R).ҽ��id = ҽ��id_In Then
        n_�������ʲ���        := 1;
        P��������             := P�������� + Rs_����(R).����;
        r_Detail.��Һid       := Rs_����(R).��Һid;
        r_Detail.��������ʱ�� := p_�ջ�ʱ��;
        If Rs_����(R).����״̬ = 1 Then
          r_Detail.�������     := 0;
          r_Detail.�����Զ���� := 1;
          p_���������м�����(0, Rs_����(R).����);
        Else
          r_Detail.�������     := 1;
          r_Detail.�����Զ���� := 0;
          p_���������м�����(1, Rs_����(R).����);
        End If;
        r_Detail.��Һid := Null;
        If Instr(',' || v_Pivas_Ids || ',', ',' || Rs_����(R).��Һid || ',') = 0 Then
          v_Pivas_Out := v_Pivas_Out || ',{"pivas_id":' || Rs_����(R).��Һid;
          v_Pivas_Out := v_Pivas_Out || ',"auto_aduit":' || r_Detail.�����Զ����;
          v_Pivas_Out := v_Pivas_Out || ',"request_time":"' || To_Char(p_�ջ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '"';
          v_Pivas_Out := v_Pivas_Out || ',"reason":"' || zlJsonStr(Nt_����ԭ��) || '"';
          v_Pivas_Out := v_Pivas_Out || '}';
          v_Pivas_Ids := v_Pivas_Ids || ',' || Rs_����(R).��Һid;
        End If;
        p_�ջ�ʱ��     := p_�ջ�ʱ�� + 1 / 24 / 60 / 60;
        n_�������ʲ��� := 0;
      End If;
    End Loop;
  
    If P�������� <> 0 Then
      --ʣ�²��ֶ����������뷽ʽ
      P�������� := P�ջ��� - P��������;
      If P�������� > 0 Then
        r_Detail.�����Զ���� := 0;
        r_Detail.�������     := 1;
        r_Detail.��������ʱ�� := �ջ�ʱ��_In;
        p_���������м�����(1, P��������);
      End If;
      P�� := 1;
    End If;
    --������ԭ
    r_Detail.�����Զ���� := Null;
    r_Detail.�������     := Null;
    r_Detail.��������ʱ�� := �ջ�ʱ��_In;
  End;

  Procedure p_Get_Json_Out As
  Begin
    If Rs_����.Count > 0 Then
      For Rp In 1 .. Rs_����.Count Loop
        If Nvl(Rs_����(Rp).����, 0) = 1 Then
          p_Charge_Check(Rp);
        
        End If;
        If Nvl(Rs_����(Rp).����, 0) = 0 Then
          p_Delbill_Check(Rp);
        End If;
      End Loop;
    End If;
    v_Json_Out := '{"code":1,"message":"�ɹ�"';
    p_Getoutlist;
    v_Json_Out := v_Json_Out || '}';
    Json_Out   := '{"output":' || v_Json_Out || '}';
  End;

Begin

  --�������
  j_Tmp      := PLJson(Json_In);
  j_Json     := j_Tmp.Get_Pljson('input');
  n_��鷽ʽ := j_Json.Get_Number('check_type');

  If 1 = n_��鷽ʽ Then
    c_ҽ��ids := j_Json.Get_Clob('order_ids');
    l_ҽ��ids := t_StrList();
    While c_ҽ��ids Is Not Null Loop
      If Length(c_ҽ��ids) <= 4000 Then
        l_ҽ��ids.Extend;
        l_ҽ��ids(l_ҽ��ids.Count) := c_ҽ��ids;
        c_ҽ��ids := Null;
      Else
        l_ҽ��ids.Extend;
        l_ҽ��ids(l_ҽ��ids.Count) := Substr(c_ҽ��ids, 1, Instr(c_ҽ��ids, ',', 3980) - 1);
        c_ҽ��ids := Substr(c_ҽ��ids, Instr(c_ҽ��ids, ',', 3980) + 1);
      End If;
    End Loop;
    For I In 1 .. l_ҽ��ids.Count Loop
      For R In (Select a.ҽ����� As ҽ��id
                From סԺ���ü�¼ A
                Where a.��¼���� In (2, 12) And a.��¼״̬ = 1 And
                      a.ҽ����� In (Select /*+cardinality(b,10) */
                                  Column_Value
                                 From Table(f_Num2List(l_ҽ��ids(I))) B)
                
                Group By a.ҽ�����
                Having Sum(Nvl(a.���ʽ��, 0)) <> 0) Loop
        Vo_Vals := Vo_Vals || ',' || r.ҽ��id;
      End Loop;
      For R In (Select a.ҽ����� As ҽ��id
                From ������ü�¼ A
                Where a.��¼���� In (2, 12) And a.��¼״̬ = 1 And
                      a.ҽ����� In (Select /*+cardinality(b,10) */
                                  Column_Value
                                 From Table(f_Num2List(l_ҽ��ids(I))) B)
                
                Group By a.ҽ�����
                Having Sum(Nvl(a.���ʽ��, 0)) <> 0) Loop
        Vo_Vals := Vo_Vals || ',' || r.ҽ��id;
      End Loop;
    End Loop;
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","order_ids":"' || Substr(Vo_Vals, 2) || '"}}';
  Elsif 2 = n_��鷽ʽ Then
    --����ջش�����ҽ����Ӧ�����Ƿ�ȫ��δ��˵Ļ��۵����Ա�ȷ��ֱ���޸Ļ��۵�������ȡ�µĵ��ݺ�
    �ջ���_In := j_Json.Get_Number('roll_num');
    ҽ��id_In := j_Json.Get_String('order_ids'); --Ŀǰ��һ��ҽ����һ�Σ������Ż�
    Select Sum(a.ʣ������) ʣ������
    Into v_��ǰ����
    From (Select Nvl(a.����, 1) * a.���� As ʣ������
           From סԺ���ü�¼ A
           Where a.ҽ����� = ҽ��id_In And a.��¼״̬ = 0
           Union All
           Select Nvl(a.����, 1) * a.���� As ʣ������
           From ������ü�¼ A
           Where a.ҽ����� = ҽ��id_In And a.��¼״̬ = 0) A;
    If Nvl(�ջ���_In, 0) > Nvl(v_��ǰ����, 0) Then
      ҽ��id_In := Null;
    End If;
    --˵����order_ids ���صĽ��Ϊ��ֵ˵�����ܸĻ��۵���Ҫ��������
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","order_ids":"' || ҽ��id_In || '"}}';
  Else
  
    Rs_�Ƽ�    := t_�Ƽ�();
    Rs_�ջ�    := t_�Ƽ�();
    Rs_ִ��    := t_ִ��();
    Rs_����    := t_����();
    Rs_����    := t_����();
    Rs_����    := t_����();
    Rs_Dellist := t_����();
    j_List     := j_Json.Get_Pljson_List('price_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_�Ƽ�.Extend;
        Lngtmp := Rs_�Ƽ�.Count;
        Rs_�Ƽ�(Lngtmp).�շ�ϸĿid := j_Item.Get_Number('fee_item_id');
        Rs_�Ƽ�(Lngtmp).�������� := j_Item.Get_Number('refer_num');
        Rs_�Ƽ�(Lngtmp).�շѷ�ʽ := j_Item.Get_Number('fee_way');
        Rs_�Ƽ�(Lngtmp).ҽ��id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
  
    --�󶨶��շ��õ��ջ�����
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('price_exe_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_�ջ�.Extend;
        Lngtmp := Rs_�ջ�.Count;
        Rs_�ջ�(Lngtmp).�շ�ϸĿid := j_Item.Get_Number('fee_item_id');
        Rs_�ջ�(Lngtmp).�ջ����� := j_Item.Get_Number('roll_num');
        Rs_�ջ�(Lngtmp).ҽ��id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
  
    --�����б�
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('pivas_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_����.Extend;
        Lngtmp := Rs_����.Count;
        Rs_����(Lngtmp).����id := j_Item.Get_Number('fee_id');
        Rs_����(Lngtmp).���� := j_Item.Get_Number('quantity');
        Rs_����(Lngtmp).�շ�ϸĿid := j_Item.Get_Number('fee_item_id');
        Rs_����(Lngtmp).��Һid := j_Item.Get_Number('pivas_id');
        Rs_����(Lngtmp).����״̬ := j_Item.Get_Number('operator_status');
        Rs_����(Lngtmp).ҽ��id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
  
    --�����б�
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('pati_list');
    If j_List Is Not Null Then
      v_Pati_List := j_List.To_Char(False);
      If v_Pati_List Is Not Null Then
        v_Pati_List := ',"pati_list":' || v_Pati_List;
      End If;
    End If;
  
    --ҩƷ����ִ���б�
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('excute_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_ִ��.Extend;
        Lngtmp := Rs_ִ��.Count;
        Rs_ִ��(Lngtmp).����id := j_Item.Get_Number('fee_id');
        Rs_ִ��(Lngtmp).��ִ���� := j_Item.Get_Number('sended_num');
        Rs_ִ��(Lngtmp).ҽ��id := j_Item.Get_Number('order_id');
      End Loop;
      v_Excute_List := j_List.To_Char(False);
      If v_Excute_List Is Not Null Then
        v_Excute_List := ',"excute_list":' || v_Excute_List;
      End If;
    End If;
  
    --ҽ��ִ���б�
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('advice_excute_list');
    If j_List Is Not Null Then
      v_Advice_Excute_List := j_List.To_Char(False);
      If v_Advice_Excute_List Is Not Null Then
        v_Advice_Excute_List := ',"advice_excute_list":' || v_Advice_Excute_List;
      End If;
    End If;
  
    --���������б�
    j_List := Pljson_List();
    j_List := j_Json.Get_Pljson_List('other_send_list');
    If j_List Is Not Null Then
      For I In 1 .. j_List.Count Loop
        j_Item := PLJson();
        j_Item := PLJson(j_List.Get(I));
        Rs_����.Extend;
        Lngtmp := Rs_����.Count;
        Rs_����(Lngtmp).No := j_Item.Get_String('fee_no');
        Rs_����(Lngtmp).���� := j_Item.Get_Number('send_num');
        Rs_����(Lngtmp).ҽ��id := j_Item.Get_Number('order_id');
      End Loop;
    End If;
    v_���� := zl_GetSysParameter('�����ջط��ñ����Զ����', 1254);
    --��ʼ��ҽ��ѭ���ջ�
    j_List_Order := j_Json.Get_Pljson_List('order_list');
    For Lp_Order In 1 .. j_List_Order.Count Loop
      j_Item := PLJson();
      j_Item := PLJson(j_List_Order.Get(Lp_Order));
    
      �ջ���_In := j_Item.Get_Number('roll_num');
      If �ջ���_In <= 0 Then
        v_Error := 'Ҫ�ջص�����Ϊ������';
        Raise Err_Custom;
      End If;
    
      No_In         := j_Item.Get_String('fee_no'); --������ʽ�����ĵ��ݺ�
      v_ҽ������    := j_Item.Get_String('advice_note');
      ҽ��id_In     := j_Item.Get_Number('order_id'); --ֻ��һ��ҽ��id
      �ջ�ʱ��_In   := To_Date(j_Item.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_Orderfeenos := j_Item.Get_String('fee_nos');
      Nt_����ԭ��   := j_Item.Get_String('reason');
    
      Nt_��������  := j_Item.Get_Number('advice_dosage');
      Nt_ҽ������  := j_Item.Get_String('advice_note');
      Nt_�������  := j_Item.Get_String('clinic_type');
      Nt_��������  := j_Item.Get_Number('is_stuff_order');
      v_�շ�ϸĿid := Null;
    
      If No_In Is Null Then
        --a.���������ջ�ģʽ
        --��Һ��ҩ��¼������������
        v_���ʲ��� := zl_GetSysParameter(23);
        --�����ջ���������ԭʼ���ý��з�̯����
      
        For r_Fee In c_Detail Loop
          --��ֵ
          Select r_Fee.����id, r_Fee.No, r_Fee.���, r_Fee.�շ�ϸĿid, r_Fee.���˲���id, r_Fee.�շ����, r_Fee.��������, r_Fee.�������,
                 r_Fee.ҽ������, r_Fee.��������, r_Fee.ʣ������, r_Fee.��ִ����, r_Fee.δִ����, r_Fee.ִ�б�־, r_Fee.��¼״̬, r_Fee.�Ǽ�ʱ��,
                 r_Fee.�շѷ�ʽ, r_Fee.�������, �ջ�ʱ��_In, Null, Null, Null, Null
          Into r_Detail.����id, r_Detail.No, r_Detail.���, r_Detail.�շ�ϸĿid, r_Detail.���˲���id, r_Detail.�շ����, r_Detail.��������,
               r_Detail.�������, r_Detail.ҽ������, r_Detail.��������, r_Detail.ʣ������, r_Detail.��ִ����, r_Detail.δִ����, r_Detail.ִ�б�־,
               r_Detail.��¼״̬, r_Detail.�Ǽ�ʱ��, r_Detail.�շѷ�ʽ, r_Detail.�������, r_Detail.��������ʱ��, r_Detail.�������, r_Detail.��������,
               r_Detail.�����Զ����, r_Detail.����ģʽ
          From Dual;
        
          --��Ҫ��һ��ȷ����ֵ:v_����ϵ��,v_סԺ��װ,�շѷ�ʽ,�������,v_�ջ�����,v_��������
          v_�ջ�����tmp := 0;
          p_Get������Ϣ(v_�ջ�����tmp);
        
          --ȷ�����շ�ϸĿid���ջ�������
          If Nvl(v_�շ�ϸĿid, 0) <> r_Detail.�շ�ϸĿid And (r_Detail.������� Not In ('5', '6', '7') Or Nvl(v_�շ�ϸĿid, 0) = 0) Then
            --����δ��̯���
            If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
              v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
              Raise Err_Custom;
            End If;
            v_�ջ����� := v_�ջ�����tmp;
            v_ҽ������ := r_Detail.ҽ������;
          End If;
          --���շ�ϸĿ��ÿ��������ϸ��̯�ջ�
          If v_�ջ����� > 0 Then
            --����Ӧ�����Ƿ��ѽ��ʣ�����ֹʱ
            v_���ʽ�� := 0;
            If v_���ʲ��� = '2' And r_Detail.��¼״̬ <> 0 Then
              Select Sum(���ʽ��)
              Into v_���ʽ��
              From סԺ���ü�¼
              Where NO = r_Detail.No And ��¼���� In (2, 12) And Nvl(�۸񸸺�, ���) = r_Detail.���;
            End If;
          
            If Nvl(v_���ʽ��, 0) = 0 Then
              v_ʣ������ := r_Detail.ʣ������;
              If v_�ջ����� > v_ʣ������ Then
                v_��ǰ���� := v_ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
            
              If r_Detail.�շ���� In ('5', '6', '7') Or r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
                p_��������(v_��ǰ����, Lngtmp);
                If Lngtmp = 0 Then
                  --��ΪҩƷ���Ĵ��ڲ���ִ�е����,������ܾ�Ҫ����Ϊ����������,������ʱ������ͳһ����
                  p_����ִ�в������(v_��ǰ����);
                End If;
              Else
                p_���������м�����(r_Detail.ִ�б�־, v_��ǰ����);
              End If;
              v_����ids := v_����ids || ',' || r_Detail.����id;
            End If;
          End If;
          v_�շ�ϸĿid := r_Detail.�շ�ϸĿid;
        End Loop;
      
        --����δ��̯���
        If v_�ջ����� > 0 Then
          v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
          Raise Err_Custom;
        End If;
      
      Elsif No_In = '�������۵�' Then
        --ֱ�ӵ�������
        For r_Fee In c_Del Loop
          --��ֵ
          Select r_Fee.����id, r_Fee.No, r_Fee.���, r_Fee.�շ�ϸĿid, Null, r_Fee.�շ����, r_Fee.��������, Null, Null, 0, r_Fee.ʣ������,
                 0, 0, 0, 0, Null, Null, r_Fee.�������, �ջ�ʱ��_In, Null, Null, Null, Null
          Into r_Detail.����id, r_Detail.No, r_Detail.���, r_Detail.�շ�ϸĿid, r_Detail.���˲���id, r_Detail.�շ����, r_Detail.��������,
               r_Detail.�������, r_Detail.ҽ������, r_Detail.��������, r_Detail.ʣ������, r_Detail.��ִ����, r_Detail.δִ����, r_Detail.ִ�б�־,
               r_Detail.��¼״̬, r_Detail.�Ǽ�ʱ��, r_Detail.�շѷ�ʽ, r_Detail.�������, r_Detail.��������ʱ��, r_Detail.�������, r_Detail.��������,
               r_Detail.�����Զ����, r_Detail.����ģʽ
          From Dual;
          v_�ջ�����tmp := 0;
          p_Get������Ϣ(v_�ջ�����tmp);
          r_Detail.��ִ���� := 0; --�ǻ��۵�����δִ��
          If Nvl(v_�շ�ϸĿid, 0) <> r_Fee.�շ�ϸĿid Then
            --����δ��̯���
            If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
              v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
              Raise Err_Custom;
            End If;
            v_�ջ����� := v_�ջ�����tmp;
            v_ҽ������ := r_Detail.ҽ������;
          End If;
          If v_�ջ����� > 0 Then
            v_ʣ������ := r_Detail.ʣ������;
            If v_�ջ����� > v_ʣ������ Then
              v_��ǰ���� := v_ʣ������;
            Else
              v_��ǰ���� := v_�ջ�����;
            End If;
            v_�ջ����� := v_�ջ����� - v_��ǰ����;
            If r_Detail.�շ���� In ('5', '6', '7') Or r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
              p_��������(v_��ǰ����, Lngtmp);
              If Lngtmp = 0 Then
                p_����ִ�в������(v_��ǰ����);
              End If;
            Else
              p_���������м�����(r_Detail.ִ�б�־, v_��ǰ����);
            End If;
          End If;
          v_�շ�ϸĿid := r_Fee.�շ�ϸĿid;
        End Loop;
        --����δ��̯���
        If v_�ջ����� > 0 Then
          v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
          Raise Err_Custom;
        End If;
      
      Elsif Nvl(No_In, '�������۵�') <> '�������۵�' Then
        Select zl_GetSysParameter(63) Into v_�Զ����� From Dual;
        Select zl_GetSysParameter(80) Into v_������� From Dual;
        --������������--�������������ܴ��ڻ��۵�����ʵ���ϵ����
        --����������ʱ�����Ҫ���߷ֿ���,����������ʱ��û���������ƣ��ж��پͳ���ټ�ʹ��ʣ��Ҳ���ж�
        --���ü�¼���շ���¼��һ��һ�Ĺ�ϵ,����������Ѿ����Ʋ��ܸ���������.
        --�����������ǰ����������ǲ��ֲ��ø����ݷ�ʽ�����,����ģû��,�������벿�ָ�Ϊ���ɵ���,����һ������ǲ���̯����
        --������ţ��շ���ţ�˳�������
        --һ��ҽ����ҩƷֻ��һ�У������ѭ����Ϊ�˴����η��͵����������ҩƷ�ڽ����ѽ��ø����ջ�
        If Nt_������� In ('5', '6', '7') Or (Nt_������� = '4' And Nvl(Nt_��������, 0) = 1) Then
        
          Select Decode(Nvl(Instr(v_�������, Decode(Nt_�������, '4', '4', '5')), 0), 0, 1, 0)
          Into Nt_�շѱ�־
          From Dual;
        
          For r_Drug In c_Negdrug Loop
            --��ֵ
            Select r_Drug.����id, r_Drug.No, r_Drug.���, r_Drug.�շ�ϸĿid, r_Drug.���˲���id, r_Drug.�շ����, r_Drug.��������, Null, Null,
                   0, r_Drug.ʣ������, 0, 0, 0, r_Drug.��¼״̬, Null, Null, r_Drug.�������, �ջ�ʱ��_In, Null, Null, Null, Null,
                   r_Drug.ִ�в���id
            Into r_Detail.����id, r_Detail.No, r_Detail.���, r_Detail.�շ�ϸĿid, r_Detail.���˲���id, r_Detail.�շ����, r_Detail.��������,
                 r_Detail.�������, r_Detail.ҽ������, r_Detail.��������, r_Detail.ʣ������, r_Detail.��ִ����, r_Detail.δִ����, r_Detail.ִ�б�־,
                 r_Detail.��¼״̬, r_Detail.�Ǽ�ʱ��, r_Detail.�շѷ�ʽ, r_Detail.�������, r_Detail.��������ʱ��, r_Detail.�������,
                 r_Detail.��������, r_Detail.�����Զ����, r_Detail.����ģʽ, r_Detail.ִ�в���id
            From Dual;
          
            v_�ջ�����tmp := 0;
            p_Get������Ϣ(v_�ջ�����tmp);
            v_�ջ����� := v_�ջ�����tmp;
          
            If v_�ջ����� > 0 Then
              v_ʣ������ := r_Detail.ʣ������;
              If v_�ջ����� > v_ʣ������ Then
                v_��ǰ���� := v_ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
              If r_Detail.�շ���� In ('5', '6', '7') Or r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
                p_��������(v_��ǰ����, Lngtmp);
                If Lngtmp = 0 Then
                  p_����ִ�в������(v_��ǰ����);
                End If;
              Else
                p_���������м�����(r_Detail.ִ�б�־, v_��ǰ����);
              End If;
            End If;
            If v_�ջ����� <= 0 Then
              Exit;
            End If;
          End Loop;
        
          If v_�ջ����� <> 0 Then
            --û���ջ���������,�շ���¼����������(���¼��ȫ������Ϊ��)
            Null;
          End If;
        Else
          --��ҩƷ����
          --ҩƷ����ִ���б�
          v_�ջ�ʣ�� := �ջ���_In;
          For I���� In 1 .. Rs_����.Count Loop
            If Rs_����(I����).ҽ��id = ҽ��id_In Then
              Nt_���͵��� := Rs_����(I����).No;
              Nt_�������� := Rs_����(I����).����;
              If Nt_�������� < v_�ջ�ʣ�� Then
                --һ���ջض�η��ͣ�����ÿ�η��ͷ��������䶯���Ƽۣ�
                v_�ջ�ʣ�� := v_�ջ�ʣ�� - Nt_��������;
                v_�ջ���   := Nt_��������;
              Else
                --һ�η������ջ�ʣ�ࣻ
                v_�ջ���   := v_�ջ�ʣ��;
                v_�ջ�ʣ�� := 0;
              End If;
            
              v_�շ����� := '';
              For r_Other In c_Other Loop
                Nt_�շ�ϸĿid := r_Other.�շ�ϸĿid;
                If Nvl(v_�շ�����, '0') <> r_Other.�շ�ϸĿid || ',' || r_Other.��� Then
                  --����ͨ������,����ҩƷ����ϵ����ϵ,���Ѿ�ִ������δִ��������
                  --��ҽ���Ƽ��л�ȡ�շѷ�ʽ�Ͷ�������
                  v_��������  := Null;
                  Nt_�շѷ�ʽ := Null;
                  For Lngtmp In 1 .. Rs_�Ƽ�.Count Loop
                    If Rs_�Ƽ�(Lngtmp).�շ�ϸĿid = Nt_�շ�ϸĿid And Rs_�Ƽ�(Lngtmp).ҽ��id = ҽ��id_In Then
                      v_��������  := Rs_�Ƽ�(Lngtmp).��������;
                      Nt_�շѷ�ʽ := Rs_�Ƽ�(Lngtmp).�շѷ�ʽ;
                    End If;
                  End Loop;
                  v_��������  := Nvl(v_��������, 1);
                  Nt_�շѷ�ʽ := Nvl(Nt_�շѷ�ʽ, 0);
                  --�������һ�η��͵ķ��ü�¼������Ҫ�ջص�����ȫ���ջ�
                  --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
                  If Nt_�շѷ�ʽ = 0 Then
                    --���»�ȡ�շѷ�ʽ�Ͷ�������
                    v_�ջ����� := v_�ջ��� * Nvl(v_��������, 1);
                  Else
                    v_�ջ����� := 0;
                    For Lngtmp In 1 .. Rs_�ջ�.Count Loop
                      If Rs_�ջ�(Lngtmp).�շ�ϸĿid = Nt_�շ�ϸĿid And Rs_�ջ�(Lngtmp).ҽ��id = ҽ��id_In Then
                        v_�ջ����� := Rs_�ջ�(Lngtmp).�ջ�����;
                      End If;
                    End Loop;
                    v_�ջ����� := Round(v_�ջ�����, 5);
                  End If;
                End If;
              
                If v_�ջ����� > 0 Then
                  If r_Other.��¼״̬ = 0 Then
                    If v_�ջ����� > r_Other.ʣ������ Then
                      v_��ǰ���� := r_Other.ʣ������;
                    Else
                      v_��ǰ���� := v_�ջ�����;
                    End If;
                  Else
                    v_��ǰ���� := v_�ջ�����;
                  End If;
                  v_�ջ����� := v_�ջ����� - v_��ǰ����;
                
                  --��ֵ
                  Select r_Other.����id, r_Other.No, r_Other.���, r_Other.�շ�ϸĿid, r_Other.���˲���id, r_Other.�շ����,
                         r_Other.��������, Null, Null, 0, 0, 0, 0, r_Other.ִ�б�־, 0, Null, Null, r_Other.�������, �ջ�ʱ��_In, Null,
                         Null, Null, Null, r_Other.ִ�в���id
                  Into r_Detail.����id, r_Detail.No, r_Detail.���, r_Detail.�շ�ϸĿid, r_Detail.���˲���id, r_Detail.�շ����,
                       r_Detail.��������, r_Detail.�������, r_Detail.ҽ������, r_Detail.��������, r_Detail.ʣ������, r_Detail.��ִ����,
                       r_Detail.δִ����, r_Detail.ִ�б�־, r_Detail.��¼״̬, r_Detail.�Ǽ�ʱ��, r_Detail.�շѷ�ʽ, r_Detail.�������,
                       r_Detail.��������ʱ��, r_Detail.�������, r_Detail.��������, r_Detail.�����Զ����, r_Detail.����ģʽ, r_Detail.ִ�в���id
                  From Dual;
                
                  Select Decode(Nvl(Instr(v_�������, r_Detail.�շ����), 0), 0, 1, 0) Into Nt_�շѱ�־ From Dual;
                
                  If r_Detail.ִ�б�־ = 1 Then
                    Nt_�շѱ�־ := 1;
                  End If;
                  If r_Other.��¼״̬ = 0 Then
                    p_���������м�����(r_Other.��¼״̬, v_��ǰ����);
                  Else
                    p_���������м�����(1, v_��ǰ����);
                  End If;
                  v_�շ����� := r_Other.�շ�ϸĿid || ',' || r_Other.���;
                End If;
              End Loop;
              If v_�ջ�ʣ�� <= 0 Then
                Exit;
              End If;
            End If;
          End Loop;
        
        End If;
      End If;
    End Loop;
    p_Get_Json_Out;
  End If;
Exception
  When Err_Custom Then
    Json_Out := '{"output":{"code":0,"message":"' || zlJsonStr(v_Error) || '"}}';
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Odr_Check;
/

Create Or Replace Procedure Zl_Exsesvr_Overdue_Recovery
(
  Json_In  Clob,
  Json_Out Out Varchar2
) As
  ---------------------------------------------------------------------------
  --���ܣ����ڷ����ջط�����ش���
  --��Σ�Json_In:��ʽ
  --  input
  --     operator_name                      C 1 ����Ա����
  --     operator_code                      C 1 ����Ա���
  --     operator_time                      C 1 ����ʱ��
  --     charge_list[]���������б�
  --                  outpati_account        N 1 �������0-סԺ����,1-�������
  --                  fee_id                 N 1 ����id
  --                  fee_item_id            N 1 �շ�ϸĿid
  --                  request_dept_id        N 1 �������id
  --                  audit_dept_id          N 1 ��˿���id
  --                  request_num            N 1 ��������
  --    del_list[]����ɾ���б�
  --                  outpati_account        N 1 �������0-סԺ����,1-�������
  --                  fee_no                 C 1 ���õ��ݺ�
  --                  serial_num             C 1 ɾ�����,��Ÿ�ʽ:����:ִ��״̬
  --    roll_list[]���������б�
  --                  outpati_account        N 1 �������0-סԺ����,1-�������
  --                  clinic_type            C 1 ҽ���������
  --                  fee_no                 C 1 ���ݺ�
  --                  item_type              C 1 �շ�ϸĿ���
  --                  fee_id                 N 1 ����id
  --                  fee_id_old             N 1 ����id,ԭʼ����id
  --                  packages_num           N 1 ����
  --                  send_num               N 1 ����
  --                  is_stuff_order         N 1 �����Ƿ��ǰ󶨵����ķ���0-������ҽ��,1-����ҽ��
  --                  stuff_used             N 1 �Ƿ��Ǹ����������ŷ���
  --                  exe_status             N 1 ִ��״̬

  --����: Json_Out,��ʽ����
  --  output
  --    code                          N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message                       C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  ---------------------------------------------------------------------------
  j_Input        Pljson;
  j_Item         Pljson;
  j_List         Pljson_List := Pljson_List();
  No_In          סԺ���ü�¼.No%Type;
  v_No           סԺ���ü�¼.No%Type;
  n_����id       סԺ���ü�¼.Id%Type;
  n_�շ�ϸĿid   סԺ���ü�¼.Id%Type;
  n_���벿��id   סԺ���ü�¼.Id%Type;
  n_��˲���id   סԺ���ü�¼.Id%Type;
  n_����         סԺ���ü�¼.����%Type;
  v_������       Varchar2(300);
  d_����ʱ��     ���˷�������.����ʱ��%Type;
  n_�������     ���˷�������.�������%Type;
  v_����ԭ��     ���˷�������.����ԭ��%Type;
  n_�Զ����     Number;
  n_�������     Number;
  v_���         Varchar2(30000);
  v_����Ա���   Varchar2(300);
  v_����Ա����   Varchar2(300);
  v_��Ա���     Varchar2(300);
  v_��Ա����     Varchar2(300);
  d_�Ǽ�ʱ��     Date;
  d_����ʱ��     Date;
  v_�������     Number;
  v_����id       סԺ���ü�¼.Id%Type;
  Old_����id     סԺ���ü�¼.Id%Type;
  v_Dec          Number;
  v_�������     Varchar2(3000);
  v_�Զ�����     Varchar2(4000);
  v_��ʼ���     Number;
  n_����         Number;
  v_�������     Number;
  �ջ�ʱ��_In    Date;
  v_Temp         Varchar2(4000);
  v_�������     Varchar2(4000);
  v_�շ����     Varchar2(300);
  v_��ǰ����     Number;
  v_��ǰ����     Number;
  v_ʵ�ս��     Number;
  n_��������ҽ�� Number;
  n_�������÷��� Number;
  n_ִ��״̬     Number;
  v_ҽ��ִ��     Number;
  n_��¼״̬     Number;
  n_����         Number;

  n_ִ��״̬���� Number(2);
  d_ִ��ʱ��     Date;
  v_ִ����       Varchar2(300);

  --���α����ڴ��������ػ��ܱ�
  Cursor c_Money
  (
    v_Start סԺ���ü�¼.���%Type,
    v_End   סԺ���ü�¼.���%Type
  ) Is
    Select ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Nvl(Ӧ�ս��, 0)) As Ӧ�ս��, Sum(Nvl(ʵ�ս��, 0)) As ʵ�ս��
    From סԺ���ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And NO = No_In And ��� Between v_Start And v_End
    Group By ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid
    Union All
    Select ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Nvl(Ӧ�ս��, 0)) As Ӧ�ս��, Sum(Nvl(ʵ�ս��, 0)) As ʵ�ս��
    From ������ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And NO = No_In And ��� Between v_Start And v_End
    Group By ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid;

Begin
  j_Item  := Pljson(Json_In);
  j_Input := j_Item.Get_Pljson('input');

  v_����Ա��� := j_Input.Get_String('operator_code');
  v_����Ա���� := j_Input.Get_String('operator_name');
  d_����ʱ��   := To_Date(j_Input.Get_String('operator_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_��Ա��� := v_����Ա���;
  v_��Ա���� := v_����Ա����;

  v_������    := v_����Ա����;
  d_�Ǽ�ʱ��  := d_����ʱ��;
  �ջ�ʱ��_In := d_�Ǽ�ʱ��;

  --�����б�
  j_List := j_Input.Get_Pljson_List('charge_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item       := Pljson();
      j_Item       := Pljson(j_List.Get(I));
      n_����id     := j_Item.Get_Number('fee_id');
      n_�շ�ϸĿid := j_Item.Get_Number('fee_item_id');
      n_���벿��id := j_Item.Get_Number('request_dept_id');
      n_��˲���id := j_Item.Get_Number('audit_dept_id');
      n_�������   := j_Item.Get_Number('request_type');
      n_����       := j_Item.Get_Number('request_num');
      n_�Զ����   := j_Item.Get_Number('auto_aduit');
      d_����ʱ��   := To_Date(j_Item.Get_String('request_time'), 'yyyy-mm-dd hh24:mi:ss');
      v_����ԭ��   := j_Item.Get_String('reason'); --����ԭ��
      Zl_���˷�������_Insert_s(n_����id, n_�շ�ϸĿid, n_���벿��id, n_����, v_������, d_����ʱ��, n_�������, v_����ԭ��, n_��˲���id, 2);
      If n_�Զ���� = 1 Then
        Zl_���˷�������_Audit_s(n_����id, d_����ʱ��, v_������, d_����ʱ��, 1, n_�������);
      End If;
    End Loop;
  End If;

  --����ɾ���б�
  j_List := Pljson_List();
  j_List := j_Input.Get_Pljson_List('del_list');
  If j_List Is Not Null Then
    For I In 1 .. j_List.Count Loop
      j_Item     := Pljson();
      j_Item     := Pljson(j_List.Get(I));
      n_������� := j_Item.Get_Number('outpati_account');
      v_No       := j_Item.Get_String('fee_no');
      v_���     := j_Item.Get_String('serial_num');
      If n_������� = 1 Then
        Zl_������ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, d_�Ǽ�ʱ��, 2);
      Else
        Zl_סԺ���ʼ�¼_Delete_s(v_No, v_���, v_����Ա���, v_����Ա����, 2, 0, d_�Ǽ�ʱ��);
      End If;
    End Loop;
  End If;

  --���ø��������б�
  j_List := Pljson_List();
  j_List := j_Input.Get_Pljson_List('roll_list');
  If j_List Is Not Null Then
    --�������������ܴ��ڻ��۵�����ʵ���ϵ����
    --���С��λ��
    Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
    --���ɻ��۵�ϵͳ����
    Select zl_GetSysParameter(80) Into v_������� From Dual;
    Select zl_GetSysParameter(63) Into v_�Զ����� From Dual;
  
    For I In 1 .. j_List.Count Loop
      j_Item := Pljson();
      j_Item := Pljson(j_List.Get(I));
    
      --���ﻹ��Ҫ��NO_In��������,��NO���ύ����
      v_�������     := j_Item.Get_String('clinic_type');
      No_In          := j_Item.Get_String('fee_no');
      v_�շ����     := j_Item.Get_String('item_type');
      v_����id       := j_Item.Get_Number('fee_id');
      Old_����id     := j_Item.Get_Number('fee_id_old');
      v_��ǰ����     := j_Item.Get_Number('packages_num');
      v_��ǰ����     := j_Item.Get_Number('send_num');
      n_�������     := j_Item.Get_Number('outpati_account');
      n_��������ҽ�� := j_Item.Get_Number('is_stuff_order');
      n_�������÷��� := j_Item.Get_Number('stuff_used');
      v_ҽ��ִ��     := j_Item.Get_Number('exe_status');
    
      Select Decode(n_ִ��״̬, 1, Decode(v_�շ����, '4', Decode(n_�������÷���, 1, 0, 1), Decode(Instr(',5,6,7,', v_�շ����), 0, 1, 0)),
                     0)
      Into n_ִ��״̬����
      From Dual;
    
      If v_�շ���� = '4' And n_�������÷��� = 1 Then
        If v_�Զ����� = '1' Then
          n_ִ��״̬���� := 1;
        End If;
      End If;
    
      If n_ִ��״̬���� = 1 Then
        d_ִ��ʱ�� := d_����ʱ��;
        v_ִ����   := v_����Ա����;
      Else
        d_ִ��ʱ�� := Null;
        v_ִ����   := Null;
      End If;
    
      If n_������� = 1 Then
        --������ü�¼
        -------------------------------------------------------------------------------------
      
        Select Nvl(Max(���), 0) + 1
        Into v_�������
        From ������ü�¼
        Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
      
        --��¼��ŷ�Χ�Դ�����ܱ�
        If v_��ʼ��� Is Null Then
          v_��ʼ��� := v_�������;
        End If;
        v_������� := v_�������;
      
        If v_������� In ('5', '6', '7') Or n_��������ҽ�� = 1 Then
          Select Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0) Into n_���� From Dual;
          Insert Into ������ü�¼
            (�Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����,
             ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��,
             ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ҽ�����, ������, ����Ա���, ����Ա����)
            Select �Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, v_����id, 2, No_In, Decode(n_����, 1, 0, 1), v_�������, Null, Null, 1, ����id,
                   ��ҳid, ��ʶ��, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, v_��ǰ����, -1 * v_��ǰ����,
                   �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec),
                   Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Null, 1, ��������id, ������, �ջ�ʱ��_In, �ջ�ʱ��_In, ִ�в���id, v_ִ����,
                   n_ִ��״̬����, d_ִ��ʱ��, ҽ�����, Decode(n_����, 1, v_��Ա����, Null), Decode(n_����, 1, Null, v_��Ա���),
                   Decode(n_����, 1, Null, v_��Ա����)
            From ������ü�¼
            Where ID = Old_����id;
        Else
          --��ҩƷ����ҽ��
          --ҽ����ִ�У��ջصķ���Ҳ��Ϊ��ִ�У�������ҩƷ�͸������õ����ģ���Ϊʵ�ʷ��ű�ʾִ��
          --����ִ��״ֱ̬�Ӹ���Ϊ���ʵ��������ٵ�����˼��ʻ��۵�
          n_ִ��״̬ := v_ҽ��ִ��;
          Select Decode(n_ִ��״̬, 1, 1, Decode(Nvl(Instr(v_�������, v_�շ����), 0), 0, 1, 0))
          Into n_��¼״̬
          From Dual;
        
          Insert Into ������ü�¼
            (�Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����,
             ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��,
             ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ��ʱ��, ִ����, ҽ�����, ������, ����Ա���, ����Ա����)
            Select �Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, v_����id, 2, No_In, n_��¼״̬, v_�������, Null,
                   Decode(a.�۸񸸺�, Null, Null, v_������� + a.�۸񸸺� - a.���), 1, a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�, a.����,
                   a.���˲���id, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, 1, -1 * v_��ǰ����, a.�Ӱ��־, a.���ӱ�־,
                   a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, Round(-1 * v_��ǰ���� * a.��׼����, v_Dec),
                   Round(-1 * v_��ǰ���� * a.��׼����, v_Dec), Null, 1, a.��������id, a.������, �ջ�ʱ��_In, �ջ�ʱ��_In, a.ִ�в���id, n_ִ��״̬����,
                   d_ִ��ʱ��, v_ִ����, a.ҽ�����, Decode(n_����, 1, v_��Ա����, Null), Decode(n_����, 1, Null, v_��Ա���),
                   Decode(n_����, 1, Null, v_��Ա����)
            From ������ü�¼ A
            Where a.Id = Old_����id;
        End If;
      
        --����˵����Ӧ�õ�ҩƷ���ķ������ �ɱ���,�˴��� ��׼���� ��Ϊ�ɱ���,Ӱ�첻��
        Select Zl_Actualmoney_s(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ��׼����, ҽ�����)
        Into v_Temp
        From ������ü�¼
        Where ID = v_����id;
        v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
        Update ������ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
        v_������� := v_�������;
        v_������� := v_������� + 1;
        n_����     := 1;
      Else
      
        --סԺ���ü�¼
        -------------------------------------------------------------------------------------
        Select Nvl(Max(���), 0) + 1
        Into v_�������
        From סԺ���ü�¼
        Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
      
        --��¼��ŷ�Χ�Դ�����ܱ�
        If v_��ʼ��� Is Null Then
          v_��ʼ��� := v_�������;
        End If;
        v_������� := v_�������;
      
        If v_������� In ('5', '6', '7') Or n_��������ҽ�� = 1 Then
          Select Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0) Into n_���� From Dual;
          Insert Into סԺ���ü�¼
            (�Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, ��ҩ����id, ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��,
             ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
             �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ҽ�����, ������, ����Ա���,
             ����Ա����)
            Select �Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, ��ҩ����id, v_����id, 2, No_In, Decode(n_����, 1, 0, 1), v_�������, Null, Null,
                   �ಡ�˵�, 2, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id,
                   v_��ǰ����, -1 * v_��ǰ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec),
                   Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Null, 1, ��������id, ������, �ջ�ʱ��_In, �ջ�ʱ��_In, ִ�в���id, v_ִ����,
                   n_ִ��״̬����, d_ִ��ʱ��, ҽ�����, Decode(n_����, 1, v_��Ա����, Null), Decode(n_����, 1, Null, v_��Ա���),
                   Decode(n_����, 1, Null, v_��Ա����)
            From סԺ���ü�¼
            Where ID = Old_����id;
        Else
          --��ҩƷ����ҽ��
          --ҽ����ִ�У��ջصķ���Ҳ��Ϊ��ִ�У�������ҩƷ�͸������õ����ģ���Ϊʵ�ʷ��ű�ʾִ��
          --����ִ��״ֱ̬�Ӹ���Ϊ���ʵ��������ٵ�����˼��ʻ��۵�
          n_ִ��״̬ := v_ҽ��ִ��;
          Select Decode(n_ִ��״̬, 1, 1, Decode(Nvl(Instr(v_�������, v_�շ����), 0), 0, 1, 0))
          Into n_��¼״̬
          From Dual;
          Insert Into סԺ���ü�¼
            (�Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, ��ҩ����id, ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��,
             ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
             �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ��ʱ��, ִ����, ҽ�����, ������, ����Ա���,
             ����Ա����)
            Select �Ƿ���, ����, ���ʵ�id, ҽ����Ч, �Ƿ���, ����, ��ҩ����id, v_����id, 2, No_In, n_��¼״̬, v_�������, Null,
                   Decode(a.�۸񸸺�, Null, Null, v_������� + a.�۸񸸺� - a.���), a.�ಡ�˵�, 2, a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�,
                   a.����, a.����, a.���˲���id, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, 1, -1 * v_��ǰ����,
                   a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, Round(-1 * v_��ǰ���� * a.��׼����, v_Dec),
                   Round(-1 * v_��ǰ���� * a.��׼����, v_Dec), Null, 1, a.��������id, a.������, �ջ�ʱ��_In, �ջ�ʱ��_In, a.ִ�в���id, n_ִ��״̬����,
                   d_ִ��ʱ��, v_ִ����, a.ҽ�����, Decode(n_����, 1, v_��Ա����, Null), Decode(n_����, 1, Null, v_��Ա���),
                   Decode(n_����, 1, Null, v_��Ա����)
            From סԺ���ü�¼ A
            Where a.Id = Old_����id;
        End If;
      
        --����˵����Ӧ�õ�ҩƷ���ķ������ �ɱ���,�˴��� ��׼���� ��Ϊ�ɱ���,Ӱ�첻��
        Select Zl_Actualmoney_s(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ��׼����, ҽ�����)
        Into v_Temp
        From סԺ���ü�¼
        Where ID = v_����id;
        v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
        Update סԺ���ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
        v_������� := v_�������;
        v_������� := v_������� + 1;
        n_����     := 2;
      End If;
    
      --���������ػ��ܱ�
      For r_Money In c_Money(v_��ʼ���, v_�������) Loop
        --�������
        Update �������
        Set ������� = Nvl(�������, 0) + r_Money.ʵ�ս��
        Where ����id = r_Money.����id And ���� = 1 And ���� = n_����;
      
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, ����, �������, Ԥ�����)
          Values
            (r_Money.����id, 1, n_����, r_Money.ʵ�ս��, 0);
        End If;
      
        --����δ�����
        Update ����δ�����
        Set ��� = Nvl(���, 0) + r_Money.ʵ�ս��
        Where ����id = r_Money.����id And ��ҳid = r_Money.��ҳid And Nvl(���˲���id, 0) = Nvl(r_Money.���˲���id, 0) And
              Nvl(���˿���id, 0) = Nvl(r_Money.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Money.��������id, 0) And
              Nvl(ִ�в���id, 0) = Nvl(r_Money.ִ�в���id, 0) And ������Ŀid + 0 = r_Money.������Ŀid And ��Դ;�� + 0 = n_����;
      
        If Sql%RowCount = 0 Then
          Insert Into ����δ�����
            (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
          Values
            (r_Money.����id, r_Money.��ҳid, r_Money.���˲���id, r_Money.���˿���id, r_Money.��������id, r_Money.ִ�в���id, r_Money.������Ŀid,
             n_����, r_Money.ʵ�ս��);
        End If;
      End Loop;
    End Loop;
  End If;

  Json_Out := '{"output":{"code":1,"message":"�ɹ�"}}';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Exsesvr_Overdue_Recovery;
/
Create Or Replace Procedure Zl_Exsesvr_Getbillbytime
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  -------------------------------------------------------------------------------------------------
  --���ܣ���ʱ�䷶Χ��ȡ���õ���
  --��Σ�json��ʽ
  --  input
  --    query_type          N 0 ��ѯ��ʽ:0-��ȡҩƷҽ�����õ��ݣ�1-��ȡ����ҽ�����õ���
  --    fee_source          N 1 ������Դ:0-������;1-����;2-סԺ
  --    start_time          C 1 ��ʼʱ�䣬��ʽ��yyyy-mm-dd hh24:mi:ss
  --    end_time            C 1 ����ʱ�䣬��ʽ��yyyy-mm-dd hh24:mi:ss
  --    exe_deptids         C 0 ִ�в���ID������ö�Ӣ�ĺŷָ�
  --    excp_exe_deptids    C 0 ��������ִ�в���ID������ö�Ӣ�ĺŷָ�
  --���Σ�json��ʽ
  --  output
  --    code                C 1 Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ���ɹ�ʱ���سɹ���Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    bill_nos            C 1 ������Ϣ:��ʽ����������1:NO1,��������2:NO2,...
  --                            ���У���������: 1-�շѴ���;2-���ʵ�����;3-���ʱ���
  -------------------------------------------------------------------------------------------------
  j_Input PLJson;
  j_Json  PLJson;

  n_��ѯ���� Number(2);
  n_������Դ Number(2);
  d_��ʼʱ�� ������ü�¼.����ʱ��%Type;
  d_����ʱ�� ������ü�¼.����ʱ��%Type;

  v_ִ�в���id     Varchar2(32767);
  v_����ִ�в���id Varchar2(32767);

  v_Output Varchar2(32767);
  c_Output Clob;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_��ѯ���� := j_Json.Get_Number('query_type');
  n_������Դ := j_Json.Get_Number('fee_source');
  d_��ʼʱ�� := To_Date(j_Json.Get_String('start_time'), 'yyyy-mm-dd hh24:mi:ss');
  d_����ʱ�� := To_Date(j_Json.Get_String('end_time'), 'yyyy-mm-dd hh24:mi:ss');

  v_ִ�в���id     := j_Json.Get_String('exe_deptids');
  v_����ִ�в���id := j_Json.Get_String('excp_exe_deptids');

  If d_��ʼʱ�� Is Null Or d_����ʱ�� Is Null Then
    Json_Out := zlJsonOut('��ѯʱ�䷶Χ��Ч��');
    Return;
  End If;

  --0-��ȡҩƷҽ�����õ���
  If Nvl(n_��ѯ����, 0) = 0 Then
    --����
    If Nvl(n_������Դ, 0) = 0 Or n_������Դ = 1 Then
      For r_���� In (Select Distinct a.No, a.��¼���� As ��������
                   From ������ü�¼ A
                   Where a.��¼���� In (1, 2) And a.ҽ����� Is Not Null And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And
                         a.�շ���� In ('5', '6', '7') And
                         (v_ִ�в���id Is Null Or Instr(',' || v_ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') > 0) And
                         (v_����ִ�в���id Is Null Or Instr(',' || v_����ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_����.�������� || ':' || r_����.No;
        Else
          v_Output := v_Output || ',' || r_����.�������� || ':' || r_����.No;
        End If;
      End Loop;
    End If;
  
    --סԺ
    If Nvl(n_������Դ, 0) = 0 Or n_������Դ = 2 Then
      For r_���� In (Select Distinct a.No, Decode(Nvl(a.�ಡ�˵�, 0), 1, 3, 2) As ��������
                   From סԺ���ü�¼ A
                   Where a.��¼���� In (1, 2) And a.ҽ����� Is Not Null And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And
                         a.�շ���� In ('5', '6', '7') And
                         (v_ִ�в���id Is Null Or Instr(',' || v_ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') > 0) And
                         (v_����ִ�в���id Is Null Or Instr(',' || v_����ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_����.�������� || ':' || r_����.No;
        Else
          v_Output := v_Output || ',' || r_����.�������� || ':' || r_����.No;
        End If;
      End Loop;
    End If;
  End If;

  --1-��ȡ����ҽ�����õ���
  If Nvl(n_��ѯ����, 0) = 1 Then
    --����
    If Nvl(n_������Դ, 0) = 0 Or n_������Դ = 1 Then
      For r_���� In (Select Distinct a.No, a.��¼���� As ��������
                   From ������ü�¼ A
                   Where a.��¼���� In (1, 2) And a.ҽ����� Is Not Null And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And a.�շ���� = '4' And
                         (v_ִ�в���id Is Null Or Instr(',' || v_ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') > 0) And
                         (v_����ִ�в���id Is Null Or Instr(',' || v_����ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_����.�������� || ':' || r_����.No;
        Else
          v_Output := v_Output || ',' || r_����.�������� || ':' || r_����.No;
        End If;
      End Loop;
    End If;
  
    --סԺ
    If Nvl(n_������Դ, 0) = 0 Or n_������Դ = 2 Then
      For r_���� In (Select Distinct a.No, Decode(Nvl(a.�ಡ�˵�, 0), 1, 3, 2) As ��������
                   From סԺ���ü�¼ A
                   Where a.��¼���� In (1, 2) And a.ҽ����� Is Not Null And a.����ʱ�� Between d_��ʼʱ�� And d_����ʱ�� And a.�շ���� = '4' And
                         (v_ִ�в���id Is Null Or Instr(',' || v_ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') > 0) And
                         (v_����ִ�в���id Is Null Or Instr(',' || v_����ִ�в���id || ',', ',' || Nvl(a.ִ�в���id, 0) || ',') = 0)) Loop
      
        If Length(v_Output) > 30000 Then
          If c_Output Is Not Null Then
            c_Output := c_Output || ',' || To_Clob(v_Output);
          Else
            c_Output := To_Clob(v_Output);
          End If;
          v_Output := Null;
        End If;
      
        If v_Output Is Null Then
          v_Output := r_����.�������� || ':' || r_����.No;
        Else
          v_Output := v_Output || ',' || r_����.�������� || ':' || r_����.No;
        End If;
      End Loop;
    End If;
  End If;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := c_Output || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","bill_nos":"') || c_Output || To_Clob('"}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","bill_nos":"' || v_Output || '"}}';
  End If;
Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getbillbytime;
/


Create Or Replace Procedure Zl_Exsesvr_Outnewbill
(
  Json_In  Clob,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ�ҽ���������ɷ��õ��ݣ������շѵ��������ʵ���סԺ���ʵ�
  --��Σ�Json_In:��ʽ
  --  input
  --    billtype              N  1  1-�շѵ���2-���ʵ�
  --    pati_id               N  1  ����id
  --    pati_pageid           N  0  ��ҳid:��Ҫ���������۲��˼��ʻ�סԺ�����������ʱ����,�������Բ�����ýӵ�
  --    sgin_no               C  1  �����
  --    pati_name             C  1  ��������
  --    pati_sex              C  1  �Ա�
  --    pati_age              C  1  ����
  --    fee_category          C  1  �ѱ�
  --    pati_deptid           N  1  ���˿���id
  --    operator_name         C  1  ����Ա���� 
  --    operator_code         C  1  ����Ա����
  --    outpati_tag           N  0  �����ʶ:1-����;3-���￨;4-��첻��ʱ��ȱʡΪ1
  --    rgst_id               N  1  �Һ�id
  --    emg_sign              N  0  �Ƿ���
  --    charge_tag            N  1  �Ƿ񻮼�:�������ʱ���룬1-��ʾ������ʻ��۵�;0-��ʾ������ʵ�
  --    placer                C  1  ������
  --    plcdept_id            N  1  ��������id
  --    happen_time           C  0  ����ʱ��:����ʱ���Ե�ǰʱ��Ϊ׼,��ʽΪyyyy-mm-dd hh24:mi:ss
  --    create_time           C  0  �Ǽ�ʱ��:����ʱ���Ե�ǰʱ��Ϊ׼,��ʽΪyyyy-mm-dd hh24:mi:ss
  --    site_no               C  0  վ���:Ժ��
  --    mdlpay_mode_name      C  0  ҽ�Ƹ��ʽ����
  --    bill_list[]           C     �����б�
  --      fee_no              C     δ����ýӵ�ʱ����ϵͳ�Զ����ɡ�
  --      apply_id            C  1  ����ID:�ⲿ�ٴ�ϵͳ������ID,Ŀǰδ�洢�����ڷ�����Ϣ
  --      item_list[]      
  --        fitem_id          N  1  �շ�ϸĿid
  --        packages_num      N  1  ����
  --        send_num          N  1  ����
  --        drug_price        N  1  ʵ��ҩƷ�����ļ۸�:ʵ�۱ش�;����������ʱ���շѼ�Ŀ.�ּ�Ϊ׼.
  --        exe_deptid        N  1  ִ�в���id
  --        memo              C  1  ժҪ
  --        order_id          N  1  ҽ��ID:ZLHIS�ڲ���ҽ��ID
  --        decoction_method  C  1  �巨
  --        morphology        C  1  ��ҩ��̬
  --        bakstuff_batch    N     ����:��Ը�������ʱ��Ч��
  --        receipt_issecret  N  1  �Ƿ��ܣ�0-�����ܣ�1-����
  --        excute_tag        N  1  ִ��״̬:0-δִ��;1-��ִ��    �ڷ���������Զ����ϵ�����£�����ִ��״̬=1�����
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    bill_list[]             �����б�
  --      fee_no            C 1 NO
  --      apply_id          N 1 �ⲿ�ٴ�ϵͳ������ID��HISϵͳĿǰδ����
  --      item_list[]           ��Ŀ�б�
  --      fee_id            N 1 ����ID
  --      order_id          N 1 ҽ��ID
  --      fitem_id          N 1 �շ�ϸĿid
  --      fee_amrcvb        N 1 Ӧ�ս��
  --      fee_ampaib        N 1 ʵ�ս��
  ---------------------------------------------------------------------------
  j_Input    PLJson;
  j_Json     PLJson;
  j_Billlist Pljson_List;
  j_Jsonbill PLJson;
  j_Itemlist Pljson_List;
  j_Jsonitem PLJson;
  v_Output   Varchar2(32767);
  c_Output   Clob;
  v_Bill     Varchar2(32767);
  c_Bill     Clob;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;
  --   input
  n_��������     Number(2); --1-�շѵ���2-���ʵ�
  n_����id       ������ü�¼.����id%Type;
  n_��ҳid       ������ü�¼.��ҳid%Type;
  n_�����       ������ü�¼.��ʶ��%Type;
  v_����         ������ü�¼.����%Type;
  v_�Ա�         ������ü�¼.�Ա�%Type;
  v_����         ������ü�¼.����%Type;
  v_�ѱ�         ������ü�¼.�ѱ�%Type;
  n_���˿���id   ������ü�¼.���˿���id%Type;
  v_����Ա����   ������ü�¼.����Ա����%Type;
  v_����Ա���   ������ü�¼.����Ա���%Type;
  n_�����ʶ     Number(2); --1-����;3-���￨;4-��첻��ʱ��ȱʡΪ1
  n_�Һ�id       Number(18);
  n_����         Number(2);
  n_����         Number(2);
  v_������       ������ü�¼.������%Type;
  n_��������id   ������ü�¼.��������id%Type;
  v_����ʱ��     Varchar2(100);
  d_����ʱ��     ������ü�¼.����ʱ��%Type;
  v_�Ǽ�ʱ��     Varchar2(100);
  d_�Ǽ�ʱ��     ������ü�¼.�Ǽ�ʱ��%Type;
  v_Ժ��         Varchar2(50);
  v_���ʽ���� ������ü�¼.���ʽ%Type;
  --    bill_list[] 
  v_No     ������ü�¼.No%Type;
  v_����id varchar2(100);
  --    item_list[]      
  n_�շ�ϸĿid ������ü�¼.�շ�ϸĿid%Type;
  n_����       ������ü�¼.����%Type;
  n_����       ������ü�¼.����%Type;
  n_��׼����   ������ü�¼.��׼����%Type;
  n_����       ������ü�¼.��׼����%Type;
  n_ִ�п���id ������ü�¼.ִ�в���id%Type;
  v_ժҪ       ������ü�¼.ժҪ%Type;
  n_ҽ�����   ������ü�¼.ҽ�����%Type;
  v_�巨       ������ü�¼.����%Type;
  v_��ҩ��̬   ������ü�¼.����%Type;
  n_����       ������ü�¼.����%Type;
  n_����       Number(2);
  n_ִ��״̬   Number(2); --0-δִ��;1-��ִ��    �ڷ���������Զ����ϵ�����£�����ִ��״̬=1�����
  n_���       ������ü�¼.���%Type;
  n_�۸񸸺�   ������ü�¼.�۸񸸺�%Type;
  v_�շ����   ������ü�¼.�շ����%Type;
  v_�۸�ȼ�   Varchar2(100);
  v_��ͨ�ȼ�   Varchar2(100);
  v_ҩƷ�ȼ�   Varchar2(100);
  v_���ĵȼ�   Varchar2(100);
  v_Pricegrade Varchar2(500);
  n_Ӧ�ս��   ������ü�¼.Ӧ�ս��%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  v_Tmp        Varchar2(500);
  n_����id     ������ü�¼.Id%Type;
  n_Count      Number(5);
  n_Money_Dec  Number(2); --���С��
  n_Price_Dec  Number(2); --����С��
Begin

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_�������� := j_Json.Get_Number('billtype');
  n_����id   := j_Json.Get_Number('pati_id');
  n_��ҳid   := j_Json.Get_Number('pati_pageid');
  n_�����   := j_Json.Get_String('sgin_no');

  v_���� := j_Json.Get_String('pati_name');
  v_�Ա� := j_Json.Get_String('pati_sex');
  v_���� := j_Json.Get_String('pati_age');
  v_�ѱ� := j_Json.Get_String('fee_category');

  n_���˿���id := j_Json.Get_Number('pati_deptid');
  v_����Ա���� := j_Json.Get_String('operator_name');
  v_����Ա��� := j_Json.Get_String('operator_code');

  n_�����ʶ := j_Json.Get_Number('outpati_tag');
  n_�Һ�id   := j_Json.Get_Number('rgst_id');
  n_����     := j_Json.Get_Number('emg_sign');
  n_����     := j_Json.Get_Number('charge_tag');

  v_������ := j_Json.Get_String('placer');

  If v_������ Is Null Then
    v_Err_Msg := 'û�д��뿪���ˣ����飡';
    Raise Err_Item;
  End If;

  n_��������id := j_Json.Get_Number('plcdept_id');

  If Nvl(n_��������id, 0) = 0 Then
    v_Err_Msg := 'û�д��뿪������id�����飡';
    Raise Err_Item;
  End If;

  v_����ʱ�� := j_Json.Get_String('happen_time');
  If v_����ʱ�� Is Not Null Then
    d_����ʱ�� := To_Date(v_����ʱ��, 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_����ʱ�� := Sysdate;
  End If;

  v_�Ǽ�ʱ�� := j_Json.Get_String('create_time');
  If v_�Ǽ�ʱ�� Is Not Null Then
    d_�Ǽ�ʱ�� := To_Date(v_�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_�Ǽ�ʱ�� := Sysdate;
  End If;

  v_Ժ��         := j_Json.Get_String('site_no');
  v_���ʽ���� := j_Json.Get_String('mdlpay_mode_name');
  If Nvl(v_Ժ��, '-') = '-' And Nvl(v_���ʽ����, '-') = '-' Then
    v_��ͨ�ȼ� := Null;
    v_ҩƷ�ȼ� := Null;
    v_���ĵȼ� := Null;
  Else
    v_Pricegrade := Zl_Get_Pricegrade_s(v_Ժ��, v_���ʽ����);
    For c_�۸�ȼ� In (Select Rownum As ���, Column_Value As �۸�ȼ� From Table(f_Str2List(v_Pricegrade, '|'))) Loop
      If c_�۸�ȼ�.��� = 1 Then
        v_��ͨ�ȼ� := c_�۸�ȼ�.�۸�ȼ�;
      End If;
      If c_�۸�ȼ�.��� = 2 Then
        v_ҩƷ�ȼ� := c_�۸�ȼ�.�۸�ȼ�;
      End If;
      If c_�۸�ȼ�.��� = 3 Then
        v_���ĵȼ� := c_�۸�ȼ�.�۸�ȼ�;
      End If;
    End Loop;
  End If;

  n_Count := 0;

  n_Money_Dec := zl_To_Number(Nvl(zl_GetSysParameter(9), '2'));
  n_Price_Dec := zl_To_Number(Nvl(zl_GetSysParameter(157), '5'));

  If Not j_Json.Exist('bill_list') Then
    v_Err_Msg := 'δ����ġ�bill_list���ڵ㣬���飡';
    Raise Err_Item;
  End If;

  j_Billlist := j_Json.Get_Pljson_List('bill_list');
  If j_Billlist.Count = 0 Then
    v_Err_Msg := 'δ����ġ�bill_list���ڵ㣬���飡';
    Raise Err_Item;
  End If;
  For I In 1 .. j_Billlist.Count Loop
  
    j_Jsonbill := PLJson();
    j_Jsonbill := PLJson(j_Billlist.Get(I));
  
    v_No := j_Jsonbill.Get_String('fee_no');
    If v_No Is Null Then
      If n_�������� = 1 Then
        v_No := Zl_Exse_Nextno(13, 0);
      Else
        v_No := Zl_Exse_Nextno(14, 0);
      End If;
    End If;
  
    v_����id := j_Jsonbill.Get_String('apply_id');
	 
    n_���   := 1;
    n_����id := Null;
    n_Count  := n_Count + 1;
    v_Bill   := Null;
  
    If Not j_Jsonbill.Exist('item_list') Then
      v_Err_Msg := 'δ����ġ�item_list���ڵ㣬���飡';
      Raise Err_Item;
    End If;
  
    j_Itemlist := j_Jsonbill.Get_Pljson_List('item_list');
    If j_Itemlist.Count = 0 Then
      v_Err_Msg := 'δ����ġ�item_list���ڵ㣬���飡';
      Raise Err_Item;
    End If;
    For J In 1 .. j_Itemlist.Count Loop
      j_Jsonitem   := PLJson();
      j_Jsonitem   := PLJson(j_Itemlist.Get(J));
      n_�շ�ϸĿid := j_Jsonitem.Get_Number('fitem_id');
      n_�۸񸸺�   := Null;
      Select Max(���) Into v_�շ���� From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;
      If v_�շ���� Is Null Then
        v_Err_Msg := '��ǰ������շ�ϸĿID�޶�Ӧ���շ���ĿĿ¼��¼�����飡';
        Raise Err_Item;
      End If;
    
      If Instr(',5,6,7,', ',' || v_�շ���� || ',') > 0 Then
        v_�۸�ȼ� := v_ҩƷ�ȼ�;
      Elsif v_�շ���� = '4' Then
        v_�۸�ȼ� := v_���ĵȼ�;
      Else
        v_�۸�ȼ� := v_��ͨ�ȼ�;
      End If;
    
      n_���� := j_Jsonitem.Get_Number('packages_num');
      If Nvl(n_����, 0) = 0 Then
        n_���� := 1;
      End If;
    
      n_���� := j_Jsonitem.Get_Number('send_num');
    
      If Nvl(n_����, 0) = 0 Then
        v_Err_Msg := '��ǰ���������Ϊ0�����飡';
        Raise Err_Item;
      End If;
    
      n_��׼����   := j_Jsonitem.Get_Number('drug_price');
      n_ִ�п���id := j_Jsonitem.Get_Number('exe_deptid');
    
      If Nvl(n_ִ�п���id, 0) = 0 Then
        v_Err_Msg := 'û�д���ִ�п���id�����飡';
        Raise Err_Item;
      End If;
    
      v_ժҪ     := j_Jsonitem.Get_String('memo');
      n_ҽ����� := j_Jsonitem.Get_Number('order_id');
      v_�巨     := j_Jsonitem.Get_String('decoction_method');
      v_��ҩ��̬ := j_Jsonitem.Get_String('morphology');
      n_����     := j_Jsonitem.Get_Number('bakstuff_batch');
      n_����     := j_Jsonitem.Get_Number('receipt_issecret');
      n_ִ��״̬ := j_Jsonitem.Get_Number('excute_tag');
    
      If Nvl(n_ҽ�����, 0) = 0 Then
        n_ҽ����� := Null;
      End If;
    
      For r_������Ŀ In (Select a.Id As �շ�ϸĿid, b.������Ŀid, c.����, c.�վݷ�Ŀ, b.�ּ�, b.ԭ��, b.�Ӱ�Ӽ���, b.�����շ���, b.ȱʡ�۸�, a.���㵥λ, a.��������,
                            a.���ηѱ�, a.��� As �շ����
                     From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                     Where b.�շ�ϸĿid = a.Id And c.Id = b.������Ŀid And Sysdate Between b.ִ������ And
                           Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And a.Id = n_�շ�ϸĿid And
                           ((b.�۸�ȼ� Is Null And Nvl(v_�۸�ȼ�, '-') = '-') Or
                           (b.�۸�ȼ� = v_�۸�ȼ� Or
                           (b.�۸�ȼ� Is Null And Not Exists
                            (Select 1
                               From �շѼ�Ŀ
                               Where b.�շ�ϸĿid = �շ�ϸĿid And �۸�ȼ� = v_�۸�ȼ� And Sysdate Between ִ������ And
                                     Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))))))
                     Order By �շ�ϸĿid, ������Ŀid) Loop
        If Nvl(n_��׼����, 0) = 0 Then
          n_���� := r_������Ŀ.�ּ�;
        Else
          n_���� := n_��׼����;
        End If;
        n_����     := Round(n_����, n_Price_Dec);
        n_Ӧ�ս�� := Round(Nvl(n_����, 0) * Nvl(n_����, 1) * n_����, n_Money_Dec);
      
        If Nvl(r_������Ŀ.���ηѱ�, 0) = 1 Then
          n_ʵ�ս�� := n_Ӧ�ս��;
        Else
          --��ȡʵ�ս��
          v_Tmp      := Zl_Actualmoney_s(v_�ѱ�, n_�շ�ϸĿid, r_������Ŀ.������Ŀid, n_Ӧ�ս��, n_����, n_����, n_ҽ�����);
          n_ʵ�ս�� := Round(zl_To_Number(Substr(v_Tmp, Instr(v_Tmp, ':') + 1)), n_Money_Dec);
        End If;
      
        Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
      
        --����������ʵ��ݻ��շѵ���
        If n_�������� = 2 Then
          Zl_������ʼ�¼_Insert_s(v_No, n_���, n_����id, n_�����, v_����, v_�Ա�, v_����, v_�ѱ�, 0, 0, n_���˿���id, n_��������id, v_������, Null,
                             r_������Ŀ.�շ�ϸĿid, r_������Ŀ.�շ����, r_������Ŀ.���㵥λ, n_����, n_����, 0, n_ִ�п���id, n_�۸񸸺�, r_������Ŀ.������Ŀid,
                             r_������Ŀ.�վݷ�Ŀ, n_����, n_Ӧ�ս��, n_ʵ�ս��, d_����ʱ��, d_�Ǽ�ʱ��, n_����, Null, v_����Ա���, v_����Ա����, n_����id,
                             Null, v_ժҪ, n_ҽ�����, n_�����ʶ, v_��ҩ��̬, v_�巨, n_��ҳid, Null, n_����, Null, n_�Һ�id, n_����, 1, n_����);
        Else
          Zl_���ﻮ�ۼ�¼_Insert_s(v_No, n_���, n_����id, n_��ҳid, n_�����, Null, v_����, v_�Ա�, v_����, v_�ѱ�, 0, n_���˿���id, n_��������id,
                             v_������, Null, r_������Ŀ.�շ�ϸĿid, r_������Ŀ.�շ����, r_������Ŀ.���㵥λ, Null, n_����, n_����, 0, n_ִ�п���id, n_�۸񸸺�,
                             r_������Ŀ.������Ŀid, r_������Ŀ.�վݷ�Ŀ, n_����, n_Ӧ�ս��, n_ʵ�ս��, d_����ʱ��, d_�Ǽ�ʱ��, v_����Ա����, n_����id, v_ժҪ,
                             n_ҽ�����, v_�巨, 1, Null, r_������Ŀ.��������, Null, Null, v_��ҩ��̬, Null, Null, n_����, Null, n_�Һ�id,
                             n_����, 1, n_����);
        
        End If;
      
        If n_�۸񸸺� Is Null Then
          n_�۸񸸺� := n_���;
        End If;
        n_��� := n_��� + 1;
      
        v_Bill := v_Bill || ',{"fee_id":' || n_����id;
        v_Bill := v_Bill || ',"order_id":' || zlJsonStr(n_ҽ�����, 1);
        v_Bill := v_Bill || ',"fitem_id":' || zlJsonStr(r_������Ŀ.�շ�ϸĿid, 1);
        v_Bill := v_Bill || ',"fee_amrcvb":' || zlJsonStr(n_Ӧ�ս��, 1);
        v_Bill := v_Bill || ',"fee_ampaib":' || zlJsonStr(n_ʵ�ս��, 1);
        v_Bill := v_Bill || '}';
      
        If Length(v_Bill) > 30000 Then
          If c_Bill Is Null Then
            c_Bill := Substr(v_Bill, 2);
          Else
            c_Bill := c_Bill || v_Bill;
          End If;
          v_Bill := Null;
        End If;
      End Loop;
    
      If n_����id Is Null Then
        v_Err_Msg := '���ݴ���ġ�item_list���ڵ�δ�ҵ���Ч�շ���Ŀ�����飡';
        Raise Err_Item;
      End If;
    
    End Loop;
  
    v_Output := Null;
    v_Output := v_Output || ',{"fee_no":"' || v_No || '"';
    v_Output := v_Output || ',"apply_id":' || zlJsonStr(v_����id);
  
    If n_Count = 1 Then
      c_Output := Substr(v_Output, 2);
    Else
      c_Output := c_Output || v_Output;
    End If;
  
    If c_Bill Is Null Then
      c_Output := c_Output || ',"item_list":[' || Substr(v_Bill, 2) || ']';
    Else
      c_Output := c_Output || ',"item_list":[' || c_Bill || v_Bill || ']';
    End If;
  
    c_Output := c_Output || '}';
  
  End Loop;
  Json_Out := '{"output":{"code":1,"message":"�ɹ�","bill_list":[' || c_Output || ']}}';

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(SQLCode || SQLErrM) || '"}}';
End Zl_Exsesvr_Outnewbill;
/


Create Or Replace Procedure Zl_Exsesvr_Getwaitingpati
(
  Json_In  Varchar2,
  Json_Out Out Clob
) As
  ---------------------------------------------------------------------------
  --���ܣ���ȡ���ﲡ����Ϣ
  --��Σ�Json_In:��ʽ
  --  input
  --    pati_id             N    ����ID:����ʱ����ʾ���ò���ID��ȡ������Ϣ
  --    exe_deptid          N    ִ�п���id:����ʱ����ʾ��ִ�в���Id��ȡ������Ϣ
  --����: Json_Out,��ʽ����
  --  output
  --    code                N 1 Ӧ����0-ʧ�ܣ�1-�ɹ�
  --    message             C 1 Ӧ����Ϣ��ʧ��ʱ���ؾ���Ĵ�����Ϣ
  --    reg_list[]          C   �Һ���Ϣ�б�
  --      pati_id           N 1 ����ID
  --      pati_name         C 1 ����
  --      pati_sex          C 1 �Ա�
  --      pati_age          C 1 ����
  --      insurance_type    C 1 ����
  --      insurance_name    C 1 ��������
  --      reg_no            C 1 �Һ�no
  --      reg_id            C 1 �Һ�ID
  --      exe_deptid        N 1 ִ�в���id
  --      exer_id           N 1 ҽ��ID
  --      exetr             C 1 ҽ��
  --      outp_room_name    C   ��������
  --      emg_sign          N   �����־
  ---------------------------------------------------------------------------
  j_Input        PLJson;
  j_Json         PLJson;
  v_Output       Varchar2(32767);
  c_Output       Clob;
  v_Para         Varchar2(100);
  n_����id       ������ü�¼.����id%Type;
  n_ִ�п���id   ������ü�¼.ִ�в���id%Type;
  n_��ͨ��Ч���� Number(2);
  n_������Ч���� Number(2);
Begin

  v_Para         := Nvl(zl_GetSysParameter(21), '11');
  n_��ͨ��Ч���� := To_Number(Substr(v_Para, 1, 1));
  n_������Ч���� := To_Number(Substr(v_Para, 2));

  --�������
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_����id     := Nvl(j_Json.Get_Number('pati_id'), 0);
  n_ִ�п���id := Nvl(j_Json.Get_Number('exe_deptid'), 0);
  For r_���� In (Select a.����id, a.����, a.�Ա�, a.����, a.����, b.���� As ��������, a.No As �Һ�no, a.Id As �Һ�id, a.ִ�в���id, a.ִ���� As ҽ��,
                      c.Id As ҽ��id, a.����, Nvl(a.����, 0) As ����
               From ���˹Һż�¼ A, ������� B, ��Ա�� C
               Where Nvl(a.ִ��״̬, 0) = 0 And a.���� = b.���(+) And (a.����id = n_����id Or n_����id = 0) And
                     (a.ִ�в���id = n_ִ�п���id Or n_ִ�п���id = 0) And
                     ((a.�Ǽ�ʱ�� >= Trunc(Sysdate) - n_��ͨ��Ч���� And Nvl(a.����, 0) = 0) Or
                      (a.�Ǽ�ʱ�� >= Trunc(Sysdate) - n_������Ч���� And Nvl(a.����, 0) = 1)) And a.ִ���� = c.����(+)) Loop
  
    zlJsonPutValue(v_Output, 'pati_id', r_����.����id, 1, 1);
    zlJsonPutValue(v_Output, 'pati_name', r_����.����);
    zlJsonPutValue(v_Output, 'pati_sex', r_����.�Ա�);
    zlJsonPutValue(v_Output, 'pati_age', r_����.����);
    zlJsonPutValue(v_Output, 'insurance_type', r_����.����, 1);
    zlJsonPutValue(v_Output, 'insurance_name', r_����.��������);
    zlJsonPutValue(v_Output, 'reg_no', r_����.�Һ�no);
    zlJsonPutValue(v_Output, 'reg_id', r_����.�Һ�id, 1);
    zlJsonPutValue(v_Output, 'exe_deptid', r_����.ִ�в���id, 1);
    zlJsonPutValue(v_Output, 'exer_id', r_����.ҽ��id, 1);
    zlJsonPutValue(v_Output, 'exetr', r_����.ҽ��);
    zlJsonPutValue(v_Output, 'outp_room_name', r_����.����);
    zlJsonPutValue(v_Output, 'emg_sign', r_����.����, 1, 2);
  
    If Length(Nvl(v_Output, ' ')) > 30000 Then
      If c_Output Is Not Null Then
        c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
      Else
        c_Output := To_Clob(v_Output);
      End If;
      v_Output := Null;
    End If;
  
  End Loop;

  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := Nvl(c_Output, '') || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;

  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","reg_list":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","reg_list":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := '{"output":{"code":0,"message":"' || Zltools.Zljsonstr(SQLCode || SQLErrM) || '"}}';
End Zl_Exsesvr_Getwaitingpati;
/

Create Or Replace Procedure Zl_Exsesvr_Getusebillinfo
(
  Json_In  Varchar2,
  Json_Out Out Clob
) Is
  ------------------------------------------------------------------------------------------------- 
  --���ܣ���ȡƱ��ʹ����ϸ���� 
  --��Σ�json��ʽ 
  --  input      
  --    occasion  N  1  ҵ�񳡺�:1-�շ�,2-Ԥ��(����Ѻ��),3-����,4-�Һ�,5-���￨ 
  --    inv_type  N  1  Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --    fee_nos  C  1  ���õ��ݺ�,����ö��ŷ���
  --      exits_history C
  --���Σ�json��ʽ 
  --   output      
  --    code  C  1  Ӧ���룺0-ʧ�ܣ�1-�ɹ�
  --    message  C  1  "Ӧ����Ϣ��  �ɹ�ʱ���سɹ���Ϣ  ʧ��ʱ���ؾ���Ĵ�����Ϣ"
  --    data[]  C  1  ʹ����ϸ����
  --      use_id  N  1  ʹ��id
  --      invoice_no  C  1  ��Ʊ��
  --      use_note  C  1  ʹ��ԭ��
  --      use_time  C  1  ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
  --      inv_user  C  1  ��Ʊʹ����
  --      recv_id  C  1  ����ID
  ------------------------------------------------------------------------------------------------- 
  j_Input PLJson;
  j_Json  PLJson;

  n_ҵ�񳡺� Number(2);
  v_Nos      Varchar2(32767);
  n_Ʊ��     Number(2);
  v_Output   Varchar2(32767);
  c_Output   Clob;
  n_Nomoved  Number(2);

  Cursor c_Ʊ����Ϣ Is(
    Select b.Id, b.���� As Ʊ�ݺ�, Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��,
           To_Char(b.ʹ��ʱ��, 'YYYY-MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����, b.����id
    From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B
    Where a.�������� = 5 And a.Id = b.��ӡid And a.No = '-' And b.Ʊ�� = 1);

  r_Ʊ����Ϣ c_Ʊ����Ϣ%RowType;

  Type Ty_Invoce Is Ref Cursor;
  c_Invoice Ty_Invoce; --��̬�α����

  v_No ������ü�¼.No%Type;
Begin
  j_Input := PLJson(Json_In);
  j_Json  := j_Input.Get_Pljson('input');

  n_ҵ�񳡺� := Nvl(j_Json.Get_Number('occasion'), 0);
  n_Ʊ��     := j_Json.Get_Number('inv_type');
  v_Nos      := j_Json.Get_String('fee_nos');

  If v_Nos Is Null Then
    Json_Out := zlJsonOut('δ������Ҫ��ѯ�ķ��õ���!');
    Return;
  End If;

  --�����־��ֻ�м��ʲ��У�����ȱʡΪNULL
  If Instr(v_Nos, ',') > 0 Then
    v_Nos := Substr(v_Nos, 1, Instr(v_Nos, ',') - 1);
  Else
    v_No := v_Nos;
  End If;
  n_Nomoved := Zl_Fun_Checkinhistory(n_ҵ�񳡺�, v_No, Null);

  If Instr(v_Nos, ',') > 0 Then
    If Nvl(n_Nomoved, 0) = 1 Then
      Open c_Invoice For
        Select b.Id, b.���� As Ʊ�ݺ�, Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��,
               To_Char(b.ʹ��ʱ��, 'YYYY-MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����, b.����id
        From HƱ�ݴ�ӡ���� A, HƱ��ʹ����ϸ B
        Where a.�������� = n_ҵ�񳡺� And a.Id = b.��ӡid And a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And
              b.Ʊ�� = n_Ʊ��;
    
    Else
      Open c_Invoice For
        Select b.Id, b.���� As Ʊ�ݺ�, Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��,
               To_Char(b.ʹ��ʱ��, 'YYYY-MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����, b.����id
        From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B
        Where a.�������� = n_ҵ�񳡺� And a.Id = b.��ӡid And a.No In (Select Column_Value From Table(f_Str2List(v_Nos))) And
              b.Ʊ�� = n_Ʊ��;
    End If;
  Else
    If Nvl(n_Nomoved, 0) = 1 Then
      Open c_Invoice For
        Select b.Id, b.���� As Ʊ�ݺ�, Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��,
               To_Char(b.ʹ��ʱ��, 'YYYY-MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����, b.����id
        From HƱ�ݴ�ӡ���� A, HƱ��ʹ����ϸ B
        Where a.�������� = n_ҵ�񳡺� And a.Id = b.��ӡid And a.No = v_Nos And b.Ʊ�� = n_Ʊ��;
    Else
      Open c_Invoice For
        Select b.Id, b.���� As Ʊ�ݺ�, Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��,
               To_Char(b.ʹ��ʱ��, 'YYYY-MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����, b.����id
        From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B
        Where a.�������� = n_ҵ�񳡺� And a.Id = b.��ӡid And a.No = v_Nos And b.Ʊ�� = n_Ʊ��;
    End If;
  End If;

  --����Ʊ����Ϣ
  v_Output := Null;

  Loop
    Fetch c_Invoice
      Into r_Ʊ����Ϣ;
    Exit When c_Invoice %NotFound;
  
    If v_Output Is Not Null Then
      v_Output := v_Output || ',';
    End If;
  
    --      use_id  N  1  ʹ��id
    --      invoice_no  C  1  ��Ʊ��
    --      use_note  C  1  ʹ��ԭ��
    --      use_time  C  1  ʹ��ʱ��:yyyy-mm-dd hh24:mi:ss
    --      inv_user  C  1  ��Ʊʹ����
    --      recv_id  C  1  ����ID
    v_Output := v_Output || '{"use_id":' || zlJsonStr(r_Ʊ����Ϣ.Id, 1);
    v_Output := v_Output || ',"invoice_no":"' || zlJsonStr(r_Ʊ����Ϣ.Ʊ�ݺ�) || '"';
    v_Output := v_Output || ',"use_note":"' || zlJsonStr(r_Ʊ����Ϣ.ʹ��ԭ��) || '"';
    v_Output := v_Output || ',"use_time":"' || zlJsonStr(r_Ʊ����Ϣ.ʹ��ʱ��) || '"';
    v_Output := v_Output || ',"inv_user":"' || zlJsonStr(r_Ʊ����Ϣ.ʹ����) || '"';
    v_Output := v_Output || ',"recv_id":' || zlJsonStr(Nvl(r_Ʊ����Ϣ.����id, 0), 1);
    v_Output := v_Output || '}';
    If Length(v_Output) > 30000 Then
      If c_Output Is Null Then
        c_Output := Substr(v_Output, 2);
      Else
        c_Output := c_Output || v_Output;
      End If;
      v_Output := Null;
    End If;
  End Loop;
  Close c_Invoice;
  If Not c_Output Is Null And Not v_Output Is Null Then
    c_Output := c_Output || ',' || To_Clob(v_Output);
    v_Output := '';
  End If;
  If Not c_Output Is Null Then
    Json_Out := To_Clob('{"output":{"code":1,"message":"�ɹ�","data":[') || c_Output || To_Clob(']}}');
  Else
    Json_Out := '{"output":{"code":1,"message":"�ɹ�","data":[' || v_Output || ']}}';
  End If;

Exception
  When Others Then
    Json_Out := zlJsonOut(SQLCode || ':' || SQLErrM);
End Zl_Exsesvr_Getusebillinfo;
/