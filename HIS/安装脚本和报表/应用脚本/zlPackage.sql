--����Ŀ¼
--1.��������,2.ҽ������,3.���˲�������,4.���û���,5.ҩƷ���Ļ���
--6.�ٴ�����,7.�ٴ�·������,8.��������,9.�������,10.�������
--11.������,12.ҽ��ҵ��,13.���˲���ҵ��,14.����ҵ��,15.ҩƷ����ҵ��
--16.�ٴ�ҽ��,17.�ٴ�·��,18.����ҵ��,19.����ҵ��,20.����ҵ��,21.���ҵ��
----------------------------------------------------------------------------
--[[1.��������]]
----------------------------------------------------------------------------
Create Or Replace Package b_Einvoice_Request Is
  ------------------------------------------------------------------
  --����Ʊ��ҵ����
  --  1.Einvoice_Start-�жϵ���Ʊ���Ƿ�����(����:1-����;0-δ����)
  --  2.EInvoice_Create-����Ʊ�ݿ���(����1-�ɹ�;0-ʧ��)
  --  3.Einvoice_Cancel_Check-����Ʊ������ǰ���(����:1-�Ϸ�;0-���Ϸ�)
  --  4.Einvoice_Cancel-����Ʊ������(����1-�ɹ�;0-ʧ��)
  ------------------------------------------------------------------
  --1.�жϵ���Ʊ���Ƿ�����
  Function Einvoice_Start
  (
    ҵ�񳡾�_In Integer,
    ����_In     ���ս����¼.����%Type,
    ����_In     Integer := Null
  ) Return Number;

  --2.����Ʊ�ݿ���
  Function Einvoice_Create
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;

  --3.����Ʊ�����ϼ��
  Function Einvoice_Cancel_Check
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;

  --4.����Ʊ������
  Function Einvoice_Cancel
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number;
End b_Einvoice_Request;
/
Create Or Replace Package b_Common_Context As
  --�ٶȣ�������>ȫ�������ģ�3-6����>���ѯ��3-6����
  --��������ȫ�������ġ�
  Procedure Set_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2,
    Value_In     In Varchar2
  );
  --���������ġ�
  Procedure Clear_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2 := Null
  );
End b_Common_Context;
/
--144329:��˶,2019-09-04,ȫ�������Ĵ��滺��
CREATE OR REPLACE Package Body b_Einvoice_Request Is
  ------------------------------------------------------------------
  --����Ʊ��ҵ����
  --  1.Einvoice_Start-�жϵ���Ʊ���Ƿ�����(����:1-����;0-δ����)
  --  2.EInvoice_Create-����Ʊ�ݿ���(����1-�ɹ�;0-ʧ��)
  --  3.Einvoice_Cancel_Check-����Ʊ������ǰ���(����:1-�Ϸ�;0-���Ϸ�)
  --  4.Einvoice_Cancel-����Ʊ������(����1-�ɹ�;0-ʧ��)
  ------------------------------------------------------------------

  Function Einvoice_Start
  (
    ҵ�񳡾�_In Integer,
    ����_In     ���ս����¼.����%Type,
    ����_In     Integer := Null
  ) Return Number Is
    ------------------------------------------------------------------
    --����:�жϵ���Ʊ���Ƿ�����
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����_in:NULL-����������;��Գ���Ϊ���˼�Ԥ��:1-����;2-סԺ;
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-����;0-δ����
    -------------------------------------------------------------------
    v_������   ����Ʊ�����.������%Type;
    v_Sql      Varchar2(1000);
    n_Return   Number(2);
    n_����     Number(2);
    n_Err_Code Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    If Nvl(ҵ�񳡾�_In, 0) = 2 And Nvl(����_In, 0) = 1 Then
      --����Ԥ�����ݲ�֧��
      Return 0;
    End If;

    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 0 Or v_������ Is Null Then
      --δ���û��ް����ƣ�ֱ�ӷ���0����ʾ�ɹ�;
      Return 0;
    End If;

    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Start(:1,:2,:3); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In ҵ�񳡾�_In, ����_In, ����_In;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code := SQLCode;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      Return 0;
    End If;
    Return 0;

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Start;

  Function Einvoice_Create
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --����:����Ʊ�ݿ���
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ��ԭԤ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ��ԭ����ID
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ���˿��Ԥ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ����ǰ�˷ѵĽ���ID,�����˷�ʱ��Ч;
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-�ɹ�;0-ʧ��
    -------------------------------------------------------------------
    v_������      ����Ʊ�����.������%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_����        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 0 Or v_������ Is Null Then
      --δ���û��ް����ƣ�ֱ�ӷ���1����ʾ�ɹ�;
      Return 1;
    End If;

    n_Err_Code := Null;
    Begin
      v_Sql := 'begin :n_return:=' || v_������ || '.EInvoice_Create(:1,:2,:3,:v_Err_Msg_out); end;';
      Execute Immediate v_Sql
        Using Out n_Return, In ҵ�񳡾�_In, ����id_In, ����id_In, Out v_Err_Msg_Out;
      ������Ϣ_Out := v_Err_Msg_Out;
      Return n_Return;
    Exception
      When Others Then
        n_Err_Code    := SQLCode;
        v_Err_Msg_Out := SQLErrM;
    End;
    If n_Err_Code = -6550 Or n_Err_Code Is Null Then
      --û�д˹��̣�����true
      Return 1;
    End If;
    Raise Err_Item;

  Exception
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Create;

  Function Einvoice_Cancel_Check
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --����:����Ʊ������
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ��ԭԤ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ��ԭ����ID
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-�ɹ�;0-ʧ��
    -------------------------------------------------------------------
    v_������      ����Ʊ�����.������%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_����        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    If ҵ�񳡾�_In = 2 Then
      --Ԥ����
      Select Max(Nvl(Ԥ������Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    Else
      --��Ԥ�����շѡ����ʡ��Һż����￨
      Select Max(Nvl(�Ƿ����Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --�ü�¼��δ���õ���Ʊ�ݵģ�ֱ�ӷ���1;
      Return 1;
    End If;

    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;

    If n_���� = 0 Or v_������ Is Null Then
      ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
      Return 0;
    End If;

    --����Ƿ񻻿�������Ʊ��
    For c_����Ʊ�� In (Select ID, �Ƿ񻻿�, ֽ�ʷ�Ʊ��
                   From ����Ʊ��ʹ�ü�¼
                   Where Ʊ�� = ҵ�񳡾�_In And ��¼״̬ = 1 And ����id = ����id_In) Loop
      --��Ե���Ʊ�ݽ��д���
      If Nvl(c_����Ʊ��.�Ƿ񻻿�, 0) = 1 Then
        --����ֽ�ʷ�Ʊ�ţ���ֹ���ϲ���
        ������Ϣ_Out := '�Ѿ�����ֽ�ʷ�Ʊ(' || c_����Ʊ��.ֽ�ʷ�Ʊ�� || ')���ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      n_Err_Code := Null;
      Begin
        v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In ҵ�񳡾�_In, c_����Ʊ��.Id, Out v_Err_Msg_Out;
        ������Ϣ_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;

      If n_Err_Code = -6550 Then
        ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;

    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel_Check;

  Function Einvoice_Cancel
  (
    ҵ�񳡾�_In  Integer,
    ����id_In    ����Ԥ����¼.����id%Type,
    ������Ϣ_Out Out Varchar2
  ) Return Number Is
    ------------------------------------------------------------------
    --����:����Ʊ������
    --���:ҵ�񳡾�_In-1-�շ�,2-Ԥ��,3-����,4-�Һ�;5-���￨
    --     ����ID_In-ҵ�񳡾�_In=2(Ԥ��)ʱ��ԭԤ��ID,ҵ�񳡾�_In<>2(Ԥ��)ʱ��ԭ����ID
    --����:������Ϣ_Out-���صĴ�����Ϣ
    --����:1-�ɹ�;0-ʧ��
    -------------------------------------------------------------------
    v_������      ����Ʊ�����.������%Type;
    v_Sql         Varchar2(1000);
    n_Return      Number(2);
    v_Err_Msg_Out Varchar2(4000);
    n_����        Number(2);
    n_Err_Code    Number(18);
    Err_Item   Exception;
    Err_Custom Exception;
  Begin

    If ҵ�񳡾�_In = 2 Then
      --Ԥ����
      Select Max(Nvl(Ԥ������Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    Else
      --��Ԥ�����շѡ����ʡ��Һż����￨
      Select Max(Nvl(�Ƿ����Ʊ��, 0)) Into n_Return From ����Ԥ����¼ Where ����id = ����id_In;
    End If;
    If Nvl(n_Return, 0) = 0 Then
      --�ü�¼��δ���õ���Ʊ�ݵģ�ֱ�ӷ���1;
      Return 1;
    End If;

    Begin
      n_���� := 1;
      Select ������ Into v_������ From ����Ʊ����� Where �Ƿ����� = 1;
    Exception
      When Others Then
        n_���� := 0;
    End;

    If n_���� = 0 Or v_������ Is Null Then
      ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
      Return 0;
    End If;

    --����Ƿ񻻿�������Ʊ��
    For c_����Ʊ�� In (Select ID, �Ƿ񻻿�, ֽ�ʷ�Ʊ��
                   From ����Ʊ��ʹ�ü�¼
                   Where Ʊ�� = ҵ�񳡾�_In And ��¼״̬ = 1 And ����id = ����id_In) Loop
      --��Ե���Ʊ�ݽ��д���
      If Nvl(c_����Ʊ��.�Ƿ񻻿�, 0) = 1 Then
        --����ֽ�ʷ�Ʊ�ţ���ֹ���ϲ���
        ������Ϣ_Out := '�Ѿ�����ֽ�ʷ�Ʊ(' || c_����Ʊ��.ֽ�ʷ�Ʊ�� || ')���ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      n_Err_Code := Null;

      --���Ⲣ��ԭ�򣬻�����Ҫ�Ƚ��м�����Ʊ���Ƿ������졣
      Begin
        v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Cancel_Check(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In ҵ�񳡾�_In, ����id_In, Out v_Err_Msg_Out;
        ������Ϣ_Out := v_Err_Msg_Out;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;

      If n_Err_Code = -6550 Then
        ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Or n_Return = 0 Then
        Raise Err_Item;
      End If;

      --���е���Ʊ�ݳ�촦��
      n_Return := 0;
      Begin
        v_Sql := 'begin :n_return:=' || v_������ || '.Einvoice_Cancel(:1,:2,:v_Err_Msg_out); end;';
        Execute Immediate v_Sql
          Using Out n_Return, In ҵ�񳡾�_In, ����id_In, Out v_Err_Msg_Out;
        ������Ϣ_Out := v_Err_Msg_Out;
        Return n_Return;
      Exception
        When Others Then
          n_Err_Code    := SQLCode;
          v_Err_Msg_Out := SQLErrM;
      End;

      If n_Err_Code = -6550 Then
        ������Ϣ_Out := '����Ʊ��δ���ã����ڴ����н����˷ѻ��˿�ڴ˴���ֹ�������Ʊ�ݳ�졣';
        Return 0;
      End If;
      If Not n_Err_Code Is Null Then
        Raise Err_Item;
      End If;
    End Loop;
    Return 1;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg_Out || '[ZLSOFT]');
    When Err_Item Then
      zl_ErrorCenter(n_Err_Code, v_Err_Msg_Out);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Einvoice_Cancel;
End b_Einvoice_Request;
/
Create Or Replace Package Body b_Common_Context As
  Procedure Set_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2,
    Value_In     In Varchar2
  ) As
  Begin
    Dbms_Session.Set_Context(Namespace_In, Name_In, Value_In, Null, Null);
  End Set_Context;
  Procedure Clear_Context
  (
    Namespace_In In Varchar2,
    Name_In      In Varchar2 := Null
  ) As
  Begin
    Dbms_Session.Clear_Context(Namespace_In, Null, Name_In);
  End Clear_Context;
End b_Common_Context;
/
create or replace context zlMessageCtx using b_Common_Context ACCESSED GLOBALLY;
Create Or Replace Package b_Zlmsg_Cache Is
  --1�����ƺ�b_Message_Cache�Լ�b_Message���𣬷�����ܻᵼ����ͣʱ����õ�
  --2���ð����b_Message_Cache(10.35.130)���Լ����b_Message(<10.35.130)�Ļ��沿��

  --������Ϣ�����棬��ֹ�ֹ����뵼������ 
  --1���ð����޸��Լ������뾡������ҵ��͹�ʱ���� 
  --2�������ڸð�������ȫ�ֱ������磺 
  --Message_Creator Zlmsg_Todo.Creator%Type := Null; 
  --��ȷʵ��Ҫ����ȫ�ֱ��������Ѿ�����ȫ�ֱ�����������PLSQL�б����ִ��������䣺 
  --ALTER PACKAGE b_zlMessage_Cache COMPILE SPECIFICATION 

  --�ж���Ϣ�Ƿ����� 
  Function Is_Message_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number;
  --���õ�ǰ�Ự��ƽ̨���� 
  Procedure Set_Platform_Call(Platform_Call Number);
  --��ȡ��Ϣ������ 
  Function Get_Message_Creator Return Varchar2;
End b_Zlmsg_Cache;
/
Create Or Replace Package Body b_Zlmsg_Cache Is
  --1�����ƺ�b_Message_Cache�Լ�b_Message���𣬷�����ܻᵼ����ͣʱ����õ�
  --2���ð����b_Message_Cache(10.35.130)���Լ����b_Message(<10.35.130)�Ļ��沿��

  --������Ϣ�����棬��ֹ�ֹ����뵼������ 
  --1���ð����޸��Լ������뾡������ҵ��͹�ʱ���� 
  --2�������ڸð�������ȫ�ֱ������磺 
  --Message_Creator Zlmsg_Todo.Creator%Type := Null; 
  --��ȷʵ��Ҫ����ȫ�ֱ��������Ѿ�����ȫ�ֱ�����������PLSQL�б����ִ��������䣺 
  --ALTER PACKAGE b_zlMessage_Cache COMPILE SPECIFICATION 

  --�Ƿ���ƽ̨���� 
  Is_Platform_Call Number(1) := 0;
  --��Ϣ�������� 
  Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --�Ƿ�������Ϣ
  Function Is_Message_Using(Msg_Code_In Zlmsg_Lists.Code%Type) Return Number As
    n_Using Zlmsg_Lists.Using%Type;
    v_Code  Zlmsg_Lists.Code%Type;
  Begin
    If Is_Platform_Call = 1 Then
      Return 0;
    End If;
    v_Code := Upper(Msg_Code_In);
    --����ת��������ת��
    n_Using := Sys_Context('zlMessageCtx', v_Code);
    If n_Using Is Null Then
      --����ȡMax�ݴ��������൱�����,�û�����û�в�ȡͬ���޸Ļ��Լ���������Ϣ���͵���δע�ᵽZlmsg_Lists���������������ִ���  
      Select Nvl(Using, 0) Into n_Using From Zlmsg_Lists A Where Code = v_Code;
      --��ʱ������ͬһ��ʵ���´�����ͬ��Zlmsg_Lists�����⣬��Ϊ����ʵ��ֻ��һ��Zlmsg_Lists
      b_Common_Context.Set_Context('zlMessageCtx', v_Code, n_Using);
    End If;
    Return n_Using;
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || 'δ��Zlmsg_Lists���ҵ���Ϣ"' || v_Code || '"������ϵ����Ա���д���' || '[ZLSOFT]');
      Return 0;
  End;
  --���õ�ǰ�ỰΪƽ̨���� 
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    Is_Platform_Call := Platform_Call;
  End Set_Platform_Call;
  --��ѯ������Ϣ����Ա
  Function Get_Message_Creator Return Varchar2 As
  Begin
    Return Message_Creator;
  End;
Begin
  --������ʵ����ִ��һ��
  Message_Creator := zl_UserName;
End b_Zlmsg_Cache;
/

Create Or Replace Package b_Message Is
  --������Ϣ�����棬��ֹ�ֹ����뵼������
  --1���ð����޸��Լ������뾡������ҵ��͹�ʱ����
  --2�������ڸð�������ȫ�ֱ������磺
  --Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --��ȷʵ��Ҫ����ȫ�ֱ��������Ѿ�����ȫ�ֱ�����������PLSQL�б����ִ��������䣺
  --ALTER PACKAGE b_Message COMPILE SPECIFICATION

  Type c_Dynamic Is Ref Cursor;

  --����ƽ̨��������
  Procedure Set_Platform_Call(Platform_Call Number);
  --��������
  Procedure Zlhis_Dict_001(Id_In Number);
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In Number);
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In Number);
  --���ò���
  Procedure Zlhis_Dict_004(����id_In Number);
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In Number);
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In Number);
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In Number);
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In Number);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In Number);
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In Number);
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In Number);
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In Number);
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In Number);
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In Number);
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In Number);
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In Number);
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In Number);
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In Number);
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In Number,
    ����_In   Varchar2,
    ������_In Varchar2,
    Ӣ����_In Varchar2
  );

  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In Number);
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In Number);
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In Number);
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In Number);
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In Number,
    Id_In   Number
  );
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In Number,
    Id_In   Number
  );
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In Number,
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2
  );
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In Number,
    Id_In   Number
  );
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In Number,
    Id_In   Number
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In Varchar2,
    Id_In   Number
  );
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In Varchar2,
    Id_In   Number
  );
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In Varchar2,
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2
  );
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In Varchar2,
    Id_In   Number
  );
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In Varchar2,
    Id_In   Number
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  );
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  );
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   Varchar2,
    ҩƷid_In Number,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ���_In   Varchar2,
    ����_In   Varchar2
  );
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  );
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  );
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  );
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  );
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In Number);
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In Number);
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In Number);
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2,
    ���_In Varchar2,
    ����_In Varchar2
  );
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In Number);
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In Number);
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In Number,
    �շ�ϸĿid_In In Number
  );
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In Number,
    �շ�ϸĿid_In In Number,
    ��Ŀ����_In   In Varchar2,
    ��Ŀ����_In   In Varchar2,
    ҽ������_In   In Varchar2,
    ҽ������_In   In Varchar2
  );
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In Number,
    Id_In   Number
  );
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In Number,
    Id_In   Number
  );
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In Number,
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2
  );
  --�շѼ�Ŀ�䶯
  Procedure Zlhis_Dict_053(�շ���Ŀid_In Number);
  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054(������Ŀid_In Number);
  --���Ĵ洢�ⷿ�䶯
  Procedure Zlhis_Dict_055(ϸĿid_In Varchar2);
  --�������Ƽ������
  Procedure Zlhis_Dictpacs_001
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ������_In Number
  );
  --�޸����Ƽ������
  Procedure Zlhis_Dictpacs_002
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ������_In Number
  );
  --ɾ�����Ƽ������
  Procedure Zlhis_Dictpacs_003
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ������_In Number
  );
  --�������Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_004
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ��ע_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Number
  );
  --�޸����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_005
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ��ע_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Number
  );
  --ɾ�����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_006
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ��ע_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Number
  );
  --����������Ŀ��λ
  Procedure Zlhis_Dictpacs_007
  (
    Id_In       Number,
    ��Ŀid_In   Number,
    ����_In     Varchar2,
    ��λ_In     Varchar2,
    ����_In     Varchar2,
    Ĭ��_In     Number,
    �ϼ�����_In Varchar2
  );
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictpacs_008
  (
    Id_In       Number,
    ��Ŀid_In   Number,
    ����_In     Varchar2,
    ��λ_In     Varchar2,
    ����_In     Varchar2,
    Ĭ��_In     Number,
    �ϼ�����_In Varchar2
  );
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictpacs_009
  (
    Id_In       Number,
    ��Ŀid_In   Number,
    ����_In     Varchar2,
    ��λ_In     Varchar2,
    ����_In     Varchar2,
    Ĭ��_In     Number,
    �ϼ�����_In Varchar2
  );
  --�������Ƽ���걾
  Procedure Zlhis_Dictlis_004
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Varchar2
  );
  --�޸����Ƽ���걾
  Procedure Zlhis_Dictlis_005
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Varchar2
  );
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictlis_006
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Varchar2
  );
  --������Ѫ������
  Procedure Zlhis_Dictlis_007
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ��Ӽ�_In Varchar2,
    ��Ѫ��_In Varchar2,
    ���_In   Varchar2,
    ��ɫ_In   Number,
    ����id_In Number
  );
  --�޸Ĳ�Ѫ������
  Procedure Zlhis_Dictlis_008
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ��Ӽ�_In Varchar2,
    ��Ѫ��_In Varchar2,
    ���_In   Varchar2,
    ��ɫ_In   Number,
    ����id_In Number
  );
  --ɾ����Ѫ������
  Procedure Zlhis_Dictlis_009
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ��Ӽ�_In Varchar2,
    ��Ѫ��_In Varchar2,
    ���_In   Varchar2,
    ��ɫ_In   Number,
    ����id_In Number
  );

  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In Varchar2);
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In Varchar2);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In Varchar2);
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004
  (
    No_In       Varchar2,
    ���_In     Number,
    ��¼״̬_In Number
  );
  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In Number,
    �շ�id_In Number
  );
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In Number,
    �����շ�id_In Number,
    ����_In       Number,
    ����id_In     Number
  );
  --ҩƷ����
  Procedure Zlhis_Drug_007(�۸�id_In Number);
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2);
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In Number,
    ʱ��_In   Number
  );
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In Number);
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In Number,
    ʱ��_In   Number
  );
  --���ķ���
  Procedure Zlhis_Drug_012
  (
    �ⷿid_In Number,
    �շ�id_In Number
  );
  --��������
  Procedure Zlhis_Drug_013
  (
    �����շ�id_In Number,
    �����շ�id_In Number,
    ����_In       Number,
    ����id_In     Number
  );
  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In Number,
    ��ҳid_In  In Number,
    ҽ��id_In  In Number,
    ҽ��ids_In In Varchar2
  );
  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ҽ��id_In In Number
  );

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  );

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  );

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );

  --7.�������߻�����ҽ����סԺ
  Procedure Zlhis_Cis_007
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  );

  --���ﻼ�߽���
  Procedure Zlhis_Cis_008
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  );

  --���ﻼ��ȡ������
  Procedure Zlhis_Cis_009
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  );

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In Number,
    ����id_In In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In Number
  );
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In Number,
    ����id_In   In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In Number,
    ����id_In   In Number,
    ���id_In   In Number,
    �������_In In Varchar2
  );

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  );

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In Number,
    ����id_In In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In Number,
    ��Ϣid_In In Number
  );

  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number
  );
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number
  );
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In Number,
    ����id_In In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In Number,
    ��Ϣid_In In Number
  );

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    No_In       In Varchar2,
    ������Դ_In In Number
  );

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    No_In       In Varchar2,
    ������Դ_In In Number
  );

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  );

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  );

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  );

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  );

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  );

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  );

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    No_In       In Varchar2,
    ��������_In In Number,
    �״�ʱ��_In In Date,
    ĩ��ʱ��_In In Date,
    ��������_In In Varchar2
  );
  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    Ҫ��ʱ��_In In Date,
    ִ��ʱ��_In In Date
  );

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    Ҫ��ʱ��_In In Date,
    ִ��ʱ��_In In Date,
    ��������_In In Number,
    ִ�н��_In In Number,
    ִ��ժҪ_In In Varchar2,
    ִ�п���_In In Number,
    ִ����_In   In Varchar2,
    �˶���_In   In Varchar2,
    ��¼��Դ_In In Number
  );
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  );
  --�������뷢�ͺ��޸�
  Procedure Zlhis_Cis_056
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number
  );

  --���ﻼ����ɾ���
  Procedure Zlhis_Cis_057
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  );

  --���ﻼ��ȡ����ɾ���
  Procedure Zlhis_Cis_058
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  );

  --ȷ��ֹͣ����ҽ��
  Procedure Zlhis_Cis_059
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  );

  --����Σ��ֵ����
  Procedure Zlhis_Cis_060
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    Σ��ֵid_In In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number
  );

  --����Ƥ�Խ����д
  Procedure Zlhis_Cis_061
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ҽ��id_In In Number
  );

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Number,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Number,
    ԭ״̬_In In Number,
    ��״̬_In In Number
  );
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Number,
    ԭ״̬_In In Number,
    ��״̬_In In Number
  );
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Number,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  );
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Number);
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Number,
    ԤԼid_In In Number
  );
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Number,
    ԤԼid_In       In Number,
    ԤԼ����_In     In Date,
    ԤԼ���_In     In Number,
    ����豸����_In In Varchar2
  );

  --36.���߷���
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In Number,
    ����id_In   In Number,
    �����id_In In Number,
    ����_In     In Varchar2
  );

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In Number,
    ����id_In   In Number,
    �����id_In In Number,
    ����_In     In Varchar2
  );

  --38.�����˿�
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In Number,
    ����id_In   In Number,
    �����id_In In Number,
    ԭ����_In   In Varchar2,
    �¿���_In   In Varchar2
  );

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In Number,
    No_In     In Varchar2
  );

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In Number,
    No_In     In Varchar2,
    ����_In   In Varchar2
  );

  --41.�����˺�
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In Number,
    No_In     In Varchar2
  );

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In Number,
    �䶯id_In   In Number
  );

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In Varchar2,
    �䶯ԭ��_In   In Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In In Number
  );

  --���������շѼ��������
  --��������_In:1-�շѽ��㣬2-�������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In Number
  );

  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In Number
  );

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In Number,
    ���ݺ�_In In Varchar2
  );

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In Number,
    ���ݺ�_In   In Varchar2
  );

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In Varchar2,
    ����id_In   In Number
  );

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In Varchar2,
    ����id_In   In Number,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  );

  --�����������룬
  Procedure Zlhis_Charge_009
  (
    ����id_In   Number,
    �������_In Number,
    ����ʱ��_In Date
  );

  --ȡ����������
  Procedure Zlhis_Charge_010
  (
    ����id_In     Number,
    �������_In   Number,
    ����ʱ��_In   Date,
    ����_In       Number,
    ���벿��id_In Number,
    ������_In     Varchar2
  );

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    ������ʽ_In In Varchar2
  );
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    Ӥ�����_In Number
  );
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    Ӥ�����_In Number
  );
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In Number);
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In Number);

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017
  (
    ����id_In   In Number,
    ԭ����id_In In Number,
    �仯ids_In  In Varchar2
  );

  --69.����ת����ת��
  Procedure Zlhis_Patient_026
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );

  Procedure Zlhis_Patient_028(����id_In In Number);

  --79.���۲���תסԺ����
  Procedure Zlhis_Patient_029
  (
    ����id_In In Number,
    ��ҳid_In In Number
  );

  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In Number);
  --Ѫ��:������Ѫ�ܾ�
  Procedure Zlhis_Blood_002(ҽ��id_In In Number);

  --70.����걾���
  Procedure Zlhis_Lis_001(�걾id_In In Number);
  --71.����걾��˳���
  Procedure Zlhis_Lis_002(�걾id_In In Number);
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In Varchar2,
    ҽ��id_In   In Number,
    ҽ��ids_In  In Varchar2
  );
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In Varchar2,
    ҽ��id_In   In Number,
    ҽ��ids_In  In Varchar2
  );
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In Number);
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In Number);
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In Number);
  --��������
  Procedure Zlhis_Emr_018
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �ļ�id_In In Number
  );
  --�������ϻ���Ա�䶯��Ϣ
  Procedure Zltools_Users_001
  (
    �û���_In In Varchar2,
    ��Աid_In In Number
  );
  Procedure Zltools_Users_002
  (
    �û���_In In Varchar2,
    ��Աid_In In Number
  );
End b_Message;
/

Create Or Replace Package Body b_Message Is
  --������Ϣ�����棬��ֹ�ֹ����뵼������
  --1���ð����޸��Լ������뾡������ҵ��͹�ʱ����
  --2�������ڸð�������ȫ�ֱ������磺
  --Message_Creator Zlmsg_Todo.Creator%Type := Null;
  --��ȷʵ��Ҫ����ȫ�ֱ��������Ѿ�����ȫ�ֱ�����������PLSQL�б����ִ��������䣺
  --ALTER PACKAGE b_Message COMPILE SPECIFICATION

  Procedure p_Msg_Todo_Insert
  (
    Msg_Code_In  Zlmsg_Todo.Msg_Code%Type,
    Key_Value_In Zlmsg_Todo.Key_Value%Type
  ) Is
  Begin
    If b_Zlmsg_Cache.Is_Message_Using(Msg_Code_In) = 0 Then
      Return;
    End If;
    Insert Into Zlmsg_Todo
      (ID, Msg_Code, Key_Value, State, Create_Time, Creator)
    Values
      (Zlmsg_Todo_Id.Nextval, Upper(Msg_Code_In), Key_Value_In, 0, Sysdate, b_Zlmsg_Cache.Get_Message_Creator);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Msg_Todo_Insert;
  --���õ�ǰ�ỰΪƽ̨����
  Procedure Set_Platform_Call(Platform_Call Number) Is
  Begin
    b_Zlmsg_Cache.Set_Platform_Call(Platform_Call);
  End Set_Platform_Call;
  --��ϢZlhis_Dict_001
  Procedure Zlhis_Dict_001(Id_In Number) Is
    v_Define Xmltype;
    v_Value  Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_001', v_Value);
  End Zlhis_Dict_001;
  --�޸Ĳ���
  Procedure Zlhis_Dict_002(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_002', v_Value);
  End Zlhis_Dict_002;
  --ͣ�ò���
  Procedure Zlhis_Dict_003(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_003', v_Value);
  End Zlhis_Dict_003;
  --���ò���
  Procedure Zlhis_Dict_004(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_004', v_Value);
  End Zlhis_Dict_004;
  --������Ա
  Procedure Zlhis_Dict_005(��Աid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_005', v_Value);
  End Zlhis_Dict_005;
  --�޸���Ա
  Procedure Zlhis_Dict_006(��Աid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_006', v_Value);
  End Zlhis_Dict_006;
  --ͣ����Ա
  Procedure Zlhis_Dict_007(��Աid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_007', v_Value);
  End Zlhis_Dict_007;
  --������Ա
  Procedure Zlhis_Dict_008(��Աid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_008', v_Value);
  End Zlhis_Dict_008;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_009(ϸĿid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_009', v_Value);
  End Zlhis_Dict_009;
  --�޸��շ���Ŀ
  Procedure Zlhis_Dict_010(ϸĿid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_010', v_Value);
  End Zlhis_Dict_010;
  --ͣ���շ���Ŀ
  Procedure Zlhis_Dict_011(ϸĿid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_011', v_Value);
  End Zlhis_Dict_011;
  --�����շ���Ŀ
  Procedure Zlhis_Dict_012(ϸĿid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ϸĿID>' || ϸĿid_In || '</ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_012', v_Value);
  End Zlhis_Dict_012;
  --����������Ŀ
  Procedure Zlhis_Dict_013(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_013', v_Value);
  End Zlhis_Dict_013;
  --�޸�������Ŀ
  Procedure Zlhis_Dict_014(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_014', v_Value);
  End Zlhis_Dict_014;
  --ͣ��������Ŀ
  Procedure Zlhis_Dict_015(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_015', v_Value);
  End Zlhis_Dict_015;
  --����������Ŀ
  Procedure Zlhis_Dict_016(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_016', v_Value);
  End Zlhis_Dict_016;
  --����������Ŀ
  Procedure Zlhis_Dict_017(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_017', v_Value);
  End Zlhis_Dict_017;
  --�޸ļ�����Ŀ
  Procedure Zlhis_Dict_018(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_018', v_Value);
  End Zlhis_Dict_018;
  --ɾ��������Ŀ
  Procedure Zlhis_Dict_019
  (
    ����id_In Number,
    ����_In   Varchar2,
    ������_In Varchar2,
    Ӣ����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID>' || '<����>' || ����_In || '</����>' || '<������>' || ������_In || '</������>' ||
               '<Ӣ����>' || Ӣ����_In || '</Ӣ����>' || '<ϵͳ>1</ϵͳ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_019', v_Value);
  End Zlhis_Dict_019;
  --������������Ŀ¼
  Procedure Zlhis_Dict_021(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_021', v_Value);
  End Zlhis_Dict_021;
  --�޸ļ�������Ŀ¼
  Procedure Zlhis_Dict_022(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_022', v_Value);
  End Zlhis_Dict_022;
  --ͣ�ü�������Ŀ¼
  Procedure Zlhis_Dict_023(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_023', v_Value);
  End Zlhis_Dict_023;
  --���ü�������Ŀ¼
  Procedure Zlhis_Dict_024(����id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_024', v_Value);
  End Zlhis_Dict_024;
  --����ҩƷ����
  Procedure Zlhis_Dict_025
  (
    ����_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_025', v_Value);
  End Zlhis_Dict_025;
  --�޸�ҩƷ����
  Procedure Zlhis_Dict_026
  (
    ����_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_026', v_Value);
  End Zlhis_Dict_026;
  --ɾ��ҩƷ����
  Procedure Zlhis_Dict_027
  (
    ����_In Number,
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_027', v_Value);
  End Zlhis_Dict_027;
  --ͣ��ҩƷ����
  Procedure Zlhis_Dict_028
  (
    ����_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_028', v_Value);
  End Zlhis_Dict_028;
  --����ҩƷ����
  Procedure Zlhis_Dict_029
  (
    ����_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_029', v_Value);
  End Zlhis_Dict_029;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_030
  (
    ���_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_030', v_Value);
  End Zlhis_Dict_030;
  --�޸�ҩƷƷ��
  Procedure Zlhis_Dict_031
  (
    ���_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_031', v_Value);
  End Zlhis_Dict_031;
  --ɾ��ҩƷƷ��
  Procedure Zlhis_Dict_032
  (
    ���_In Varchar2,
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_032', v_Value);
  End Zlhis_Dict_032;
  --ͣ��ҩƷƷ��
  Procedure Zlhis_Dict_033
  (
    ���_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_033', v_Value);
  End Zlhis_Dict_033;
  --����ҩƷƷ��
  Procedure Zlhis_Dict_034
  (
    ���_In Varchar2,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_034', v_Value);
  End Zlhis_Dict_034;
  --����ҩƷ���
  Procedure Zlhis_Dict_035
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_035', v_Value);
  End Zlhis_Dict_035;
  --�޸�ҩƷ���
  Procedure Zlhis_Dict_036
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_036', v_Value);
  End Zlhis_Dict_036;
  --ɾ��ҩƷ���
  Procedure Zlhis_Dict_037
  (
    ���_In   Varchar2,
    ҩƷid_In Number,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ���_In   Varchar2,
    ����_In   Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><���>' || ���_In || '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_037', v_Value);
  End Zlhis_Dict_037;
  --ͣ��ҩƷ���
  Procedure Zlhis_Dict_038
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_038', v_Value);
  End Zlhis_Dict_038;
  --����ҩƷ���
  Procedure Zlhis_Dict_039
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_039', v_Value);
  End Zlhis_Dict_039;
  --����ҩƷ�洢�ⷿ
  Procedure Zlhis_Dict_040
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_040', v_Value);
  End Zlhis_Dict_040;
  --����ҩƷ�����޶�
  Procedure Zlhis_Dict_041
  (
    ���_In   Varchar2,
    ҩƷid_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���>' || ���_In || '</���><ҩƷID>' || ҩƷid_In || '</ҩƷID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_041', v_Value);
  End Zlhis_Dict_041;
  --��������Ʒ��
  Procedure Zlhis_Dict_042(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_042', v_Value);
  End Zlhis_Dict_042;
  --�������Ĺ��
  Procedure Zlhis_Dict_043(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_043', v_Value);
  End Zlhis_Dict_043;
  --�޸����Ĺ��
  Procedure Zlhis_Dict_044(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_044', v_Value);
  End Zlhis_Dict_044;
  --ɾ�����Ĺ��
  Procedure Zlhis_Dict_045
  (
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2,
    ���_In Varchar2,
    ����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID><����>' || ����_In || '</����><����>' || ����_In || '</����><���>' || ���_In ||
               '</���><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_045', v_Value);
  End Zlhis_Dict_045;
  --ͣ�����Ĺ��
  Procedure Zlhis_Dict_046(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_046', v_Value);
  End Zlhis_Dict_046;
  --�������Ĺ��
  Procedure Zlhis_Dict_047(Id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || Id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_047', v_Value);
  End Zlhis_Dict_047;
  --ҽ������
  Procedure Zlhis_Dict_048
  (
    ����_In       In Number,
    �շ�ϸĿid_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_048', v_Value);
  End Zlhis_Dict_048;
  --ɾ��ҽ������
  Procedure Zlhis_Dict_049
  (
    ����_In       In Number,
    �շ�ϸĿid_In In Number,
    ��Ŀ����_In   In Varchar2,
    ��Ŀ����_In   In Varchar2,
    ҽ������_In   In Varchar2,
    ҽ������_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><�շ�ϸĿID>' || �շ�ϸĿid_In || '</�շ�ϸĿID><��Ŀ����>' || ��Ŀ����_In || '</��Ŀ����><��Ŀ����>' ||
               ��Ŀ����_In || '</��Ŀ����><ҽ������>' || ҽ������_In || '</ҽ������><ҽ������>' || ҽ������_In || '</ҽ������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_049', v_Value);
  End Zlhis_Dict_049;
  --�������ķ���
  Procedure Zlhis_Dict_050
  (
    ����_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_050', v_Value);
  End Zlhis_Dict_050;
  --�޸����ķ���
  Procedure Zlhis_Dict_051
  (
    ����_In Number,
    Id_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_051', v_Value);
  End Zlhis_Dict_051;
  --ɾ�����ķ���
  Procedure Zlhis_Dict_052
  (
    ����_In Number,
    Id_In   Number,
    ����_In Varchar2,
    ����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><ID>' || Id_In || '</ID><����>' || ����_In || '</����><����>' || ����_In ||
               '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_052', v_Value);
  End Zlhis_Dict_052;
  --�շѼ�Ŀ�䶯
  Procedure Zlhis_Dict_053(�շ���Ŀid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ���ĿID>' || �շ���Ŀid_In || '</�շ���ĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_053', v_Value);
  End Zlhis_Dict_053;

  --�����շѶ��ձ䶯
  Procedure Zlhis_Dict_054(������Ŀid_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������ĿID>' || ������Ŀid_In || '</������ĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_054', v_Value);
  End Zlhis_Dict_054;

  --���Ĵ洢�ⷿ�䶯
  Procedure Zlhis_Dict_055(ϸĿid_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ�ϸĿID>' || ϸĿid_In || '</�շ�ϸĿID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICT_055', v_Value);
  End Zlhis_Dict_055;

  --�������Ƽ������
  Procedure Zlhis_Dictpacs_001
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ������_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_001', v_Value);
  End Zlhis_Dictpacs_001;

  --�޸����Ƽ������
  Procedure Zlhis_Dictpacs_002
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ������_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_002', v_Value);
  End Zlhis_Dictpacs_002;
  --ɾ�����Ƽ������
  Procedure Zlhis_Dictpacs_003
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ������_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><������>' || ������_In ||
               '</������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_003', v_Value);
  End Zlhis_Dictpacs_003;
  --�������Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_004
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ��ע_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_004', v_Value);
  End Zlhis_Dictpacs_004;
  --�޸����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_005
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ��ע_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_005', v_Value);
  End Zlhis_Dictpacs_005;
  --ɾ�����Ƽ�鲿λ
  Procedure Zlhis_Dictpacs_006
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    ��ע_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In ||
               '</����><��ע>' || ��ע_In || '</��ע><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In || '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_006', v_Value);
  End Zlhis_Dictpacs_006;
  --����������Ŀ��λ
  Procedure Zlhis_Dictpacs_007
  (
    Id_In       Number,
    ��Ŀid_In   Number,
    ����_In     Varchar2,
    ��λ_In     Varchar2,
    ����_In     Varchar2,
    Ĭ��_In     Number,
    �ϼ�����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��><�ϼ�����>' || �ϼ�����_In || '</�ϼ�����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_007', v_Value);
  End Zlhis_Dictpacs_007;
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictpacs_008
  (
    Id_In       Number,
    ��Ŀid_In   Number,
    ����_In     Varchar2,
    ��λ_In     Varchar2,
    ����_In     Varchar2,
    Ĭ��_In     Number,
    �ϼ�����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��><�ϼ�����>' || �ϼ�����_In || '</�ϼ�����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_008', v_Value);
  End Zlhis_Dictpacs_008;
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictpacs_009
  (
    Id_In       Number,
    ��Ŀid_In   Number,
    ����_In     Varchar2,
    ��λ_In     Varchar2,
    ����_In     Varchar2,
    Ĭ��_In     Number,
    �ϼ�����_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ID>' || Id_In || '</ID><��ĿID>' || ��Ŀid_In || '</��ĿID><����>' || ����_In || '</����><��λ>' || ��λ_In ||
               '</��λ><����>' || ����_In || '</����><Ĭ��>' || Ĭ��_In || '</Ĭ��><�ϼ�����>' || �ϼ�����_In || '</�ϼ�����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DICTPACS_009', v_Value);
  End Zlhis_Dictpacs_009;
  --����������Ŀ��λ
  Procedure Zlhis_Dictlis_004
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_004', v_Value);
  End Zlhis_Dictlis_004;
  --�޸�������Ŀ��λ
  Procedure Zlhis_Dictlis_005
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_005', v_Value);
  End Zlhis_Dictlis_005;
  --ɾ��������Ŀ��λ
  Procedure Zlhis_Dictlis_006
  (
    ����_In     Varchar2,
    ����_In     Varchar2,
    ����_In     Varchar2,
    �����Ա�_In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><�����Ա�>' || �����Ա�_In ||
               '</�����Ա�></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_006', v_Value);
  End Zlhis_Dictlis_006;
  --������Ѫ������
  Procedure Zlhis_Dictlis_007
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ��Ӽ�_In Varchar2,
    ��Ѫ��_In Varchar2,
    ���_In   Varchar2,
    ��ɫ_In   Number,
    ����id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_007', v_Value);
  End Zlhis_Dictlis_007;
  --������Ѫ������
  Procedure Zlhis_Dictlis_008
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ��Ӽ�_In Varchar2,
    ��Ѫ��_In Varchar2,
    ���_In   Varchar2,
    ��ɫ_In   Number,
    ����id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_008', v_Value);
  End Zlhis_Dictlis_008;
  --������Ѫ������
  Procedure Zlhis_Dictlis_009
  (
    ����_In   Varchar2,
    ����_In   Varchar2,
    ����_In   Varchar2,
    ��Ӽ�_In Varchar2,
    ��Ѫ��_In Varchar2,
    ���_In   Varchar2,
    ��ɫ_In   Number,
    ����id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����>' || ����_In || '</����><����>' || ����_In || '</����><����>' || ����_In || '</����><��Ӽ�>' || ��Ӽ�_In ||
               '</��Ӽ�><��Ѫ��>' || ��Ѫ��_In || '</��Ѫ��><��ɫ>' || ��ɫ_In || '</��ɫ><���>' || ���_In || '</���><����ID_In>' || ����id_In ||
               '</����ID_In></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_DictLis_009', v_Value);
  End Zlhis_Dictlis_009;
  --ҩƷ��ҩ����
  Procedure Zlhis_Drug_001(No_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_001', v_Value);
  End Zlhis_Drug_001;
  --ȡ��ҩƷ��ҩ����
  Procedure Zlhis_Drug_002(No_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_002', v_Value);
  End Zlhis_Drug_002;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_003(No_In Varchar2) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_003', v_Value);
  End Zlhis_Drug_003;
  --ҩƷ�ƿⵥ����
  Procedure Zlhis_Drug_004
  (
    No_In       Varchar2,
    ���_In     Number,
    ��¼״̬_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><���ݺ�>' || No_In || '</���ݺ�><���>' || ���_In || '</���><��¼״̬>' || ��¼״̬_In || '</��¼״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_004', v_Value);
  End Zlhis_Drug_004;
  --���ŷ�ҩ
  Procedure Zlhis_Drug_005
  (
    �ⷿid_In Number,
    �շ�id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�ⷿID>' || �ⷿid_In || '</�ⷿID><�շ�ID>' || �շ�id_In || '</�շ�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_005', v_Value);
  End Zlhis_Drug_005;
  --������ҩ
  Procedure Zlhis_Drug_006
  (
    �����շ�id_In Number,
    �����շ�id_In Number,
    ����_In       Number,
    ����id_In     Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������¼ID>' || �����շ�id_In || '</������¼ID><������¼ID>' || �����շ�id_In || '</������¼ID><����>' || ����_In ||
               '</����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_006', v_Value);
  End Zlhis_Drug_006;
  --ҩƷ����
  Procedure Zlhis_Drug_007(�۸�id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_007', v_Value);
  End Zlhis_Drug_007;
  --���䷢��
  Procedure Zlhis_Drug_008(��¼ids_In Varchar2) Is
    v_Value  Zlmsg_Todo.Key_Value%Type;
    n_��¼id Number(18);
    v_Tmp    Varchar2(4000);
    n_Length Number(18);
  Begin
    If ��¼ids_In Is Null Then
      v_Tmp := Null;
    Else
      v_Tmp := ��¼ids_In || ',';
    End If;
  
    v_Value := '<root><��¼IDS>';
  
    While v_Tmp Is Not Null Loop
      --�ֽⵥ��ID��
      n_��¼id := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1));
      v_Tmp    := Replace(',' || v_Tmp, ',' || n_��¼id || ',');
    
      --�жϵ�ǰ�����Ƿ񼴽���������
      Select Lengthb(v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>') Into n_Length From Dual;
      If n_Length > 950 Then
        v_Value := v_Value || '</��¼IDs></root>';
        b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
        v_Value := '<root><��¼IDs>';
      End If;
    
      v_Value := v_Value || '<��¼ID>' || n_��¼id || '</��¼ID>';
    End Loop;
  
    v_Value := v_Value || '</��¼IDS></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_008', v_Value);
  End Zlhis_Drug_008;
  --ҩƷ���ۼ�
  Procedure Zlhis_Drug_009
  (
    �۸�id_In Number,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_009', v_Value);
  End Zlhis_Drug_009;
  --���ĵ��ɱ���
  Procedure Zlhis_Drug_010(�۸�id_In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_010', v_Value);
  End Zlhis_Drug_010;
  --���ĵ��ۼ�
  Procedure Zlhis_Drug_011
  (
    �۸�id_In Number,
    ʱ��_In   Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�۸�ID>' || �۸�id_In || '</�۸�ID><ʱ��>' || ʱ��_In || '</ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_011', v_Value);
  End Zlhis_Drug_011;
  --���ķ���
  Procedure Zlhis_Drug_012
  (
    �ⷿid_In Number,
    �շ�id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�ⷿID>' || �ⷿid_In || '</�ⷿID><�շ�ID>' || �շ�id_In || '</�շ�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_012', v_Value);
  End Zlhis_Drug_012;
  --��������
  Procedure Zlhis_Drug_013
  (
    �����շ�id_In Number,
    �����շ�id_In Number,
    ����_In       Number,
    ����id_In     Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><������¼ID>' || �����շ�id_In || '</������¼ID><������¼ID>' || �����շ�id_In || '</������¼ID><����>' || ����_In ||
               '</����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_DRUG_013', v_Value);
  End Zlhis_Drug_013;
  --2.ֹͣ����ҽ����סԺ
  Procedure Zlhis_Cis_002
  (
    ����id_In  In Number,
    ��ҳid_In  In Number,
    ҽ��id_In  In Number,
    ҽ��ids_In In Varchar2
  ) Is
    r_Data c_Dynamic;
    v_Id   Number(18);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CIS_002') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                  '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                   '</ID></root>');
    Else
      Open r_Data For 'Select ID From ����ҽ����¼ Where ID In (Select Column_Value From Table(f_Num2list(:1))) And ���id Is Null'
        Using ҽ��ids_In;
      Loop
        Fetch r_Data
          Into v_Id;
        Exit When r_Data%NotFound;
      
        b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_002',
                                    '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || v_Id ||
                                     '</ID></root>');
      End Loop;
    End If;
  End Zlhis_Cis_002;

  --3.���ϻ���ҽ��������/סԺ
  Procedure Zlhis_Cis_003
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ҽ��id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_003', v_Value);
  End Zlhis_Cis_003;

  --4.��������ҽ����סԺ
  Procedure Zlhis_Cis_004
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_004;

  --5.������������ҽ����סԺ
  Procedure Zlhis_Cis_005
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_005;

  --6.���߻�����ҽ����סԺ
  Procedure Zlhis_Cis_006
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_006;

  --7.�������߻�����ҽ����סԺ
  Procedure Zlhis_Cis_007
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('Zlhis_Cis_007', v_Value);
  End Zlhis_Cis_007;

  --���ﻼ�߽���
  Procedure Zlhis_Cis_008
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_008', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_008;

  --���ﻼ��ȡ������
  Procedure Zlhis_Cis_009
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_009', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_009;

  --10.�´ﻼ����ϣ�����/סԺ
  Procedure Zlhis_Cis_010
  (
    ����id_In In Number,
    ����id_In In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ���id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_010',
                                '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || ���id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_010;
  --11.�����������
  Procedure Zlhis_Cis_011
  (
    ����id_In   In Number,
    ����id_In   In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    Id_In       In Number,
    ����id_In   In Number,
    ���id_In   In Number,
    �������_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ID>' || Id_In || '</ID><����ID>' ||
               ����id_In || '</����ID><���ID>' || ���id_In || '</���ID><�������>' || �������_In || '</�������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_011', v_Value);
  End Zlhis_Cis_011;

  --����ִ��ҽ��У��
  Procedure Zlhis_Cis_012
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_012;

  --13.����Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_014
  (
    ����id_In In Number,
    ����id_In In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In Number,
    ��Ϣid_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_014', v_Value);
  End Zlhis_Cis_014;
  --15.���߼������룬����/סԺ
  Procedure Zlhis_Cis_016
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_016', v_Value);
  End Zlhis_Cis_016;
  --16.���߼�����룬����/סԺ
  Procedure Zlhis_Cis_017
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� Varchar2(20);
  Begin
    Execute Immediate 'Select Max(a.��������) From ������ĿĿ¼ A, ����ҽ����¼ B Where b.������Ŀid = a.Id And b.Id = :1'
      Into v_��������
      Using ҽ��id_In;
  
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    If v_�������� = '����' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_054', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_017', v_Value);
    End If;
  End Zlhis_Cis_017;
  --17.�����������룬����/סԺ
  Procedure Zlhis_Cis_018
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_018', v_Value);
  End Zlhis_Cis_018;
  --18.������Ѫ���룬סԺ
  Procedure Zlhis_Cis_019
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_019', v_Value);
  End Zlhis_Cis_019;
  --19.���߻������룬סԺ
  Procedure Zlhis_Cis_020
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_020', v_Value);
  End Zlhis_Cis_020;
  --20.��������ҽ����סԺ
  Procedure Zlhis_Cis_021
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_021', v_Value);
  End Zlhis_Cis_021;
  --21.��������ҽ����סԺ
  Procedure Zlhis_Cis_022
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_022', v_Value);
  End Zlhis_Cis_022;
  --22.������������ҽ����סԺ
  Procedure Zlhis_Cis_023
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_023', v_Value);
  End Zlhis_Cis_023;

  --24.���Σ��ֵ�Ķ�֪ͨ
  Procedure Zlhis_Cis_025
  (
    ����id_In In Number,
    ����id_In In Number, --���ﲡ�� �Һ�ID��סԺ���� ��ҳID
    ҽ��id_In In Number,
    ��Ϣid_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><����ID>' || ����id_In || '</����ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ID>' ||
               ��Ϣid_In || '</ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_025', v_Value);
  End Zlhis_Cis_025;

  --����ִ��ҽ������
  Procedure Zlhis_Cis_026
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In ||
                                 '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_026;

  --�������߼�������
  Procedure Zlhis_Cis_036
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    No_In       In Varchar2,
    ������Դ_In In Number --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_036', v_Value);
  End Zlhis_Cis_036;

  --�������߼������
  Procedure Zlhis_Cis_037
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    No_In       In Varchar2,
    ������Դ_In In Number --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value    Zlmsg_Todo.Key_Value%Type;
    v_�������� Varchar2(20);
  Begin
    Execute Immediate 'Select Max(a.��������) From ������ĿĿ¼ A, ����ҽ����¼ B Where b.������Ŀid = a.Id And b.Id = :1'
      Into v_��������
      Using ҽ��id_In;
  
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO><������Դ>' || ������Դ_In ||
               '</������Դ></root>';
    If v_�������� = '����' Then
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_055', v_Value);
    Else
      b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_037', v_Value);
    End If;
  End Zlhis_Cis_037;

  --����������������
  Procedure Zlhis_Cis_038
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_038', v_Value);
  End Zlhis_Cis_038;

  --����������Ѫ����
  Procedure Zlhis_Cis_039
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_039', v_Value);
  End Zlhis_Cis_039;

  --�������߻�������
  Procedure Zlhis_Cis_040
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_040', v_Value);
  End Zlhis_Cis_040;

  --������������ҽ��
  Procedure Zlhis_Cis_041
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_041', v_Value);
  End Zlhis_Cis_041;

  --������������ҽ��
  Procedure Zlhis_Cis_042
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_042', v_Value);
  End Zlhis_Cis_042;

  --������������ҽ��
  Procedure Zlhis_Cis_043
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ���ͺ�_In In Number,
    ҽ��id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_043', v_Value);
  End Zlhis_Cis_043;

  --��������ִ��ҽ��
  Procedure Zlhis_Cis_044
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    No_In       In Varchar2,
    ��������_In In Number,
    �״�ʱ��_In In Date,
    ĩ��ʱ��_In In Date,
    ��������_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' ||
               ҽ��id_In || '</ID><NO>' || No_In || '</NO><��������>' || ��������_In || '</��������><�״�ʱ��>' ||
               To_Char(�״�ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</�״�ʱ��><ĩ��ʱ��>' ||
               To_Char(ĩ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ĩ��ʱ��><��������>' || ��������_In || '</��������></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_044', v_Value);
  End Zlhis_Cis_044;

  --����ҽ��ִ�еǼ�
  Procedure Zlhis_Cis_050
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    Ҫ��ʱ��_In In Date,
    ִ��ʱ��_In In Date
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_050', v_Value);
  End Zlhis_Cis_050;

  --����ҽ��ȡ��ִ�еǼ�
  Procedure Zlhis_Cis_051
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    Ҫ��ʱ��_In In Date,
    ִ��ʱ��_In In Date,
    ��������_In In Number,
    ִ�н��_In In Number,
    ִ��ժҪ_In In Varchar2,
    ִ�п���_In In Number,
    ִ����_In   In Varchar2,
    �˶���_In   In Varchar2,
    ��¼��Դ_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><Ҫ��ʱ��>' || To_Char(Ҫ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</Ҫ��ʱ��><ִ��ʱ��>' || To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</ִ��ʱ��><��������>' || ��������_In ||
               '</��������><ִ�н��>' || ִ�н��_In || '</ִ�н��><ִ��ժҪ>' || ִ��ժҪ_In || '</ִ��ժҪ><ִ�п���ID>' || ִ�п���_In ||
               '</ִ�п���ID><ִ����>' || ִ����_In || '</ִ����><�˶���>' || �˶���_In || '</�˶���><��¼��Դ>' || ��¼��Դ_In || '</��¼��Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_051', v_Value);
  End Zlhis_Cis_051;
  --����ҽ��ִ�����
  Procedure Zlhis_Cis_052
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_052',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_052;
  --����ҽ������ִ�����
  Procedure Zlhis_Cis_053
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ���ͺ�_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_053',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><���ͺ�>' || ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_053;

  --�������뷢�ͺ��޸�
  Procedure Zlhis_Cis_056
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    ���ͺ�_In   In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number --1-����;2-סԺ;3-����(������ڸ��ﲿ�Ž�����������);4-��첡��
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In || '</�Һŵ�><���ͺ�>' ||
               ���ͺ�_In || '</���ͺ�><ID>' || ҽ��id_In || '</ID><������Դ>' || ������Դ_In || '</������Դ></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_056', v_Value);
  End Zlhis_Cis_056;

  --���ﻼ����ɾ���
  Procedure Zlhis_Cis_057
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_057', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_057;

  --���ﻼ��ȡ����ɾ���
  Procedure Zlhis_Cis_058
  (
    ����id_In In Number,
    �Һŵ�_In In Varchar2
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_058', '<root><����ID>' || ����id_In || '</����ID><NO>' || �Һŵ�_In || '</NO></root>');
  End Zlhis_Cis_058;

  --ȷ��ֹͣ����ҽ��
  Procedure Zlhis_Cis_059
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_059',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><ID>' || ҽ��id_In ||
                                 '</ID></root>');
  End Zlhis_Cis_059;

  --����Σ��ֵ����
  Procedure Zlhis_Cis_060
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    �Һŵ�_In   In Varchar2,
    Σ��ֵid_In In Number,
    ҽ��id_In   In Number,
    ������Դ_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_060',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><ID>' || Σ��ֵid_In || '</ID><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><������Դ>' || ������Դ_In ||
                                 '</������Դ></root>');
  End Zlhis_Cis_060;

  --����Ƥ�Խ����д
  Procedure Zlhis_Cis_061
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �Һŵ�_In In Varchar2,
    ҽ��id_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_CIS_061',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�Һŵ�>' || �Һŵ�_In ||
                                 '</�Һŵ�><ID>' || ҽ��id_In || '</ID></root>');
  End Zlhis_Cis_061;

  --26.��鱨����ɣ�������ʱ
  Procedure Zlhis_Pacs_001
  (
    ҽ��id_In   In Number,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PACS_001') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '</��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_001', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_001;
  --27.���״̬ͬ�������״̬�ı��
  Procedure Zlhis_Pacs_002
  (
    ҽ��id_In In Number,
    ԭ״̬_In In Number,
    ��״̬_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_002', v_Value);
  End Zlhis_Pacs_002;
  --28.���״̬���ˣ����״̬���˺�
  Procedure Zlhis_Pacs_003
  (
    ҽ��id_In In Number,
    ԭ״̬_In In Number,
    ��״̬_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԭ״̬>' || ԭ״̬_In || '</ԭ״̬><��״̬>' || ��״̬_In || '</��״̬></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_003', v_Value);
  End Zlhis_Pacs_003;
  --29.��鱨�泷��������������ʱ
  Procedure Zlhis_Pacs_004
  (
    ҽ��id_In   In Number,
    ����id_Ins  In Varchar2,
    ��������_In In Number --1-�ϰ�PACS���棬2-�ϰ没���༭�����棬3-�°�༭������
  ) Is
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PACS_004') = 0 Then
      Return;
    End If;
    For R In (Select '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><����ID>' || Column_Value || '</����ID><��������>' || ��������_In ||
                      '</��������></root>' As Xml_Value
              From Table(f_Str2list(����id_Ins))) Loop
      b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_004', r.Xml_Value);
    End Loop;
  End Zlhis_Pacs_004;
  --30.���Σ��ֵ֪ͨ����鷢��Σ��ֵʱ
  Procedure Zlhis_Pacs_005(ҽ��id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_005', v_Value);
  End Zlhis_Pacs_005;
  -- ���ԤԼ֪ͨ�����ԤԼʱ
  Procedure Zlhis_Pacs_006
  (
    ҽ��id_In In Number,
    ԤԼid_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_006', v_Value);
  End Zlhis_Pacs_006;
  -- ȡ�����ԤԼ��ȡ��ԤԼʱ
  Procedure Zlhis_Pacs_007
  (
    ҽ��id_In       In Number,
    ԤԼid_In       In Number,
    ԤԼ����_In     In Date,
    ԤԼ���_In     In Number,
    ����豸����_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ԤԼID>' || ԤԼid_In || '</ԤԼID><ԤԼ����>' || ԤԼ����_In || '</ԤԼ����><ԤԼ���>' ||
               ԤԼ���_In || '</ԤԼ���><����豸����>' || ����豸����_In || '</����豸����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PACS_007', v_Value);
  End Zlhis_Pacs_007;

  --36.���߷�����󶨿�
  Procedure Zlhis_Patient_018
  (
    �䶯id_In   In Number,
    ����id_In   In Number,
    �����id_In In Number,
    ����_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_018', v_Value);
  End;

  --37.�����˿�
  Procedure Zlhis_Patient_019
  (
    �䶯id_In   In Number,
    ����id_In   In Number,
    �����id_In In Number,
    ����_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><����>' || ����_In || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_019', v_Value);
  End;

  --38.���߲���/����
  Procedure Zlhis_Patient_020
  (
    �䶯id_In   In Number,
    ����id_In   In Number,
    �����id_In In Number,
    ԭ����_In   In Varchar2,
    �¿���_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ID>' || �䶯id_In || '</�䶯ID><����ID>' || ����id_In || '</����ID><�����ID>' || �����id_In ||
               '</�����ID><ԭ����>' || ԭ����_In || '</ԭ����><�¿���>' || �¿���_In || '</�¿���></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_020', v_Value);
  End;

  --39.���˹ҺŵǼǣ�����ԤԼ�Ǽ�)
  Procedure Zlhis_Regist_001
  (
    �Һ�id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_001', v_Value);
  End;

  --40.���˷���
  Procedure Zlhis_Regist_002
  (
    �Һ�id_In In Number,
    No_In     In Varchar2,
    ����_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO><����>' || Nvl(����_In, '') || '</����></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_002', v_Value);
  End;

  --41.�����˺ţ���ȡ��ԤԼ)
  Procedure Zlhis_Regist_003
  (
    �Һ�id_In In Number,
    No_In     In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�Һ�ID>' || �Һ�id_In || '</�Һ�ID><NO>' || No_In || '</NO></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_003', v_Value);
  End;

  --42.�ٴ����ﰲ�ŵ���
  Procedure Zlhis_Regist_004
  (
    �䶯ԭ��_In In Integer, --1-ͣ��;2-����;3-���ұ䶯
    ��¼id_In   In Number,
    �䶯id_In   In Number
    
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><��¼ID>' || ��¼id_In || '</��¼ID><�䶯ID>' || �䶯id_In ||
               '</�䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_004', v_Value);
  End;

  --43.���ﻼ�߹ҺŻ��Ų���
  Procedure Zlhis_Regist_005
  (
    No_In         In Varchar2,
    �䶯ԭ��_In   Integer, --1-����;2-����;3-ԤԼ���ڱ䶯,
    ����䶯id_In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><NO>' || No_In || '</NO><�䶯ԭ��>' || �䶯ԭ��_In || '</�䶯ԭ��><����䶯ID>' || ����䶯id_In ||
               '</����䶯ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_REGIST_005', v_Value);
  End;

  --���������շѼ��������
  Procedure Zlhis_Charge_002
  (
    ��������_In In Number,
    ����id_In   In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --��������_In:1-�շѽ��㣬2-�������
    v_Value := '<root><��������>' || ��������_In || '</��������><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_002', v_Value);
  End;

  --46.�����˷ѵ���
  Procedure Zlhis_Charge_004
  (
    �˷�����_In In Number,
    ����id_In   In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    --�˷�����_In:1-�շѽ��㣬2-�������
    v_Value := '<root><�˷�����>' || �˷�����_In || '</�˷�����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_004', v_Value);
  End;

  --47.��Ԥ����
  Procedure Zlhis_Charge_005
  (
    Ԥ��id_In In Number,
    ���ݺ�_In In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><Ԥ��ID>' || Ԥ��id_In || '</Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_005', v_Value);
  End;

  --48.��Ԥ����(����������Ԥ�����)
  Procedure Zlhis_Charge_006
  (
    ��Ԥ��id_In In Number,
    ���ݺ�_In   In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><��Ԥ��ID>' || ��Ԥ��id_In || '</��Ԥ��ID><���ݺ�>' || ���ݺ�_In || '</���ݺ�></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_006', v_Value);
  End;

  --סԺ���ʵ���
  Procedure Zlhis_Charge_007
  (
    �շ����_In In Varchar2,
    ����id_In   In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_007', v_Value);
  End;

  --סԺ���ʵ�������
  Procedure Zlhis_Charge_008
  (
    �շ����_In In Varchar2,
    ����id_In   In Number,
    �շ�ids_In  In Varchar2 := Null --���ܷ���ID��Ӧ����շ�id����Ӧ��ʽ���շ�id,����|�շ�id,��������ҩƷ����
  ) Is
    v_Value   Zlmsg_Todo.Key_Value%Type;
    v_Tmp     Varchar2(4000);
    v_Infotmp Varchar2(4000);
    v_Fields  Varchar2(4000);
    v_�շ�id  Varchar2(50);
    v_����    Varchar2(20);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CHARGE_008') = 0 Then
      Return;
    End If;
    v_Value := '<root><�շ����>' || �շ����_In || '</�շ����><����ID>' || ����id_In || '</����ID>';
  
    If �շ�ids_In Is Null Then
      v_Infotmp := Null;
      v_Tmp     := '<�շ�IDS>' || '<�շ�ID>' || '</�շ�ID>' || '<����>' || '</����>' || '</�շ�IDS>';
    Else
      v_Infotmp := �շ�ids_In || '|';
      While v_Infotmp Is Not Null Loop
        --�ֽ��շ�ID��
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        v_�շ�id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        v_����    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        v_Tmp := v_Tmp || '<�շ�IDS>' || '<�շ�ID>' || v_�շ�id || '</�շ�ID>' || '<����>' || v_���� || '</����>' || '</�շ�IDS>';
      End Loop;
    End If;
  
    v_Value := v_Value || v_Tmp || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_008', v_Value);
  End;

  --������������
  Procedure Zlhis_Charge_009
  (
    ����id_In   Number,
    �������_In Number,
    ����ʱ��_In Date
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CHARGE_009') = 0 Then
      Return;
    End If;
    v_Value := '<root><�������>' || �������_In || '</�������><����ID>' || ����id_In || '</����ID><����ʱ��>' ||
               To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '</����ʱ��>' || '</root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_009', v_Value);
  End;

  --ȡ����������
  Procedure Zlhis_Charge_010
  (
    ����id_In     Number,
    �������_In   Number,
    ����ʱ��_In   Date,
    ����_In       Number,
    ���벿��id_In Number,
    ������_In     Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_CHARGE_010') = 0 Then
      Return;
    End If;
    v_Value := '<root><�������>' || �������_In || '</�������><����ID>' || ����id_In || '</����ID><����>' || ����_In || '</����><���벿��ID>' ||
               ���벿��id_In || '</���벿��ID><������>' || ������_In || '</������><����ʱ��>' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') ||
               '</����ʱ��></root>';
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_CHARGE_010', v_Value);
  End;

  --53.סԺ������Ժ�Ǽ�
  Procedure Zlhis_Patient_001
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    n_�䶯id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_001') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id
      Using ����id_In, ��ҳid_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_001',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_001;
  --54.סԺ������Ժ���
  Procedure Zlhis_Patient_002
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    n_�䶯id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_002') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id
      Using ����id_In, ��ҳid_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_002',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_002;
  --56.סԺ���ߴ�λ���
  Procedure Zlhis_Patient_004
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    v_ԭ����   Varchar2(255);
    v_�´���   Varchar2(255);
    n_�䶯id   Number(18);
    n_��ʼԭ�� Number(3);
    d_��ʼʱ�� Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_004') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, ����, ��ʼʱ��, ��ʼԭ�� From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id, v_�´���, d_��ʼʱ��, n_��ʼԭ��
      Using ����id_In, ��ҳid_In;
  
    Execute Immediate 'Select Max(����) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� = :3 And ��ֹԭ�� = :4 And Nvl(���Ӵ�λ, 0) = 0'
      Into v_ԭ����
      Using ����id_In, ��ҳid_In, d_��ʼʱ��, n_��ʼԭ��;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_004',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����>' ||
                                 v_ԭ���� || '</ԭ����>' || '<�´���>' || v_�´��� || '</�´���>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' ||
                                 '</root>');
  End Zlhis_Patient_004;
  --57.סԺ���߲�����
  Procedure Zlhis_Patient_005
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    n_�䶯id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_005') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id
      Using ����id_In, ��ҳid_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_005',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_005;
  --58.סԺ���߱������
  Procedure Zlhis_Patient_006
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    ������ʽ_In In Varchar2
  ) Is
    n_Id         Number(18);
    n_����id     Number(18);
    n_����id     Number(18);
    n_����ȼ�id Number(18);
    n_ҽ��С��id Number(18);
    v_����       Varchar2(20);
    v_���λ�ʿ   Varchar2(50);
    v_����ҽʦ   Varchar2(50);
    v_����ҽʦ   Varchar2(50);
    v_����ҽʦ   Varchar2(50);
    v_����       Varchar2(50);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_006') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select  Max(id), Max(����id), Max(����id), Max(����ȼ�id), Max(ҽ��С��id), Max(����), Max(���λ�ʿ), Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(����) ' ||
                      'From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) And Nvl(���Ӵ�λ, 0) = 0'
      Into n_Id, n_����id, n_����id, n_����ȼ�id, n_ҽ��С��id, v_����, v_���λ�ʿ, v_����ҽʦ, v_����ҽʦ, v_����ҽʦ, v_����
      Using ����id_In, ��ҳid_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_006',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><������ʽ>' || ������ʽ_In ||
                                 '</������ʽ><����ID>' || n_����id || '</����ID>' || '<����ID>' || n_����id || '</����ID>' || '<����ȼ�ID>' ||
                                 n_����ȼ�id || '</����ȼ�ID>' || '<ҽ��С��ID>' || n_ҽ��С��id || '</ҽ��С��ID>' || '<����>' || v_���� ||
                                 '</����>' || '<���λ�ʿ>' || v_���λ�ʿ || '</���λ�ʿ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' ||
                                 '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����ҽʦ>' || v_����ҽʦ || '</����ҽʦ>' || '<����>' || v_���� ||
                                 '</����>' || '<ID>' || n_Id || '</ID>' || '</root>');
  End Zlhis_Patient_006;
  --59.סԺ����ҽ�����
  Procedure Zlhis_Patient_007
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    v_ԭסԺҽ�� Varchar2(100);
    v_��סԺҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ����ҽ�� Varchar2(100);
    v_������ҽ�� Varchar2(100);
    v_ԭ���λ�ʿ Varchar2(100);
    v_�����λ�ʿ Varchar2(100);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_007') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, ����ҽʦ, ����ҽʦ, ����ҽʦ, ���λ�ʿ, ��ʼʱ��, ��ʼԭ�� ' ||
                      'From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id, v_��סԺҽ��, v_������ҽ��, v_������ҽ��, v_�����λ�ʿ, d_��ʼʱ��, n_��ʼԭ��
      Using ����id_In, ��ҳid_In;
  
    Execute Immediate 'Select Max(����ҽʦ), Max(����ҽʦ), Max(����ҽʦ), Max(���λ�ʿ) ' ||
                      'From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� = :3 And ��ֹԭ�� = :4 And Nvl(���Ӵ�λ, 0) = 0'
      Into v_ԭסԺҽ��, v_ԭ����ҽ��, v_ԭ����ҽ��, v_ԭ���λ�ʿ
      Using ����id_In, ��ҳid_In, d_��ʼʱ��, n_��ʼԭ��;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_007',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭסԺҽ��>' ||
                                 v_ԭסԺҽ�� || '</ԭסԺҽ��>' || '<��סԺҽ��>' || v_��סԺҽ�� || '</��סԺҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� ||
                                 '</ԭ����ҽ��>' || '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ����ҽ��>' || v_ԭ����ҽ�� || '</ԭ����ҽ��>' ||
                                 '<������ҽ��>' || v_������ҽ�� || '</������ҽ��>' || '<ԭ���λ�ʿ>' || v_ԭ���λ�ʿ || '</ԭ���λ�ʿ>' || '<�����λ�ʿ>' ||
                                 v_�����λ�ʿ || '</�����λ�ʿ>' || '<�䶯ID>' || n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_007;
  --סԺ���߻���ȼ����
  Procedure Zlhis_Patient_008
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    v_ԭ����ȼ�id Number(18);
    v_�»���ȼ�id Number(18);
    n_�䶯id       Number(18);
    n_��ʼԭ��     Number(3);
    d_��ʼʱ��     Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_008') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, ����ȼ�id, ��ʼʱ��, ��ʼԭ�� From ���˱䶯��¼ ' ||
                      'Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id, v_�»���ȼ�id, d_��ʼʱ��, n_��ʼԭ��
      Using ����id_In, ��ҳid_In;
  
    Execute Immediate 'Select Max(����ȼ�id) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� = :3 And ��ֹԭ�� = :4 And Nvl(���Ӵ�λ, 0) = 0'
      Into v_ԭ����ȼ�id
      Using ����id_In, ��ҳid_In, d_��ʼʱ��, n_��ʼԭ��;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_008',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ԭ����ȼ�ID>' ||
                                 v_ԭ����ȼ�id || '</ԭ����ȼ�ID>' || '<�»���ȼ�ID>' || v_�»���ȼ�id || '</�»���ȼ�ID>' || '<�䶯ID>' ||
                                 n_�䶯id || '</�䶯ID>' || '</root>');
  End Zlhis_Patient_008;
  --60.סԺ����Ԥ��Ժ
  Procedure Zlhis_Patient_009
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    n_�䶯id Number;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_009') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select Max(ID) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id
      Using ����id_In, ��ҳid_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_009',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_009;
  --61.סԺ���߳�Ժ
  Procedure Zlhis_Patient_010
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_010',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID></root>');
  End Zlhis_Patient_010;
  --62.סԺ�����������Ǽ�
  Procedure Zlhis_Patient_011
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    Ӥ�����_In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_011',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_011;
  --63.סԺ����ת�����
  Procedure Zlhis_Patient_012
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    v_ת������id Number(18);
    v_ת�����id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_012') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, ����id, ��ʼʱ��, ��ʼԭ�� From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id, v_ת�����id, d_��ʼʱ��, n_��ʼԭ��
      Using ����id_In, ��ҳid_In;
  
    Execute Immediate 'Select Max(����id) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� = :3 And ��ֹԭ�� = :4 And Nvl(���Ӵ�λ, 0) = 0'
      Into v_ת������id
      Using ����id_In, ��ҳid_In, d_��ʼʱ��, n_��ʼԭ��;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_012',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�����ID>' || v_ת�����id || '</ת�����ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_012;
  --64.�������Ǽ�����
  Procedure Zlhis_Patient_013
  (
    ����id_In   In Number,
    ��ҳid_In   In Number,
    Ӥ�����_In Number
  ) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_013',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><Ӥ�����>' || Ӥ�����_In ||
                                 '</Ӥ�����></root>');
  End Zlhis_Patient_013;
  --65.���ﻼ�ߵǼ�
  Procedure Zlhis_Patient_015(����id_In In Number) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_015', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_015;
  --66.������Ϣ�޸�
  Procedure Zlhis_Patient_016(����id_In In Number) Is
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_016', '<root><����ID>' || ����id_In || '</����ID></root>');
  End Zlhis_Patient_016;

  --67.���ߺϲ�
  Procedure Zlhis_Patient_017
  (
    ����id_In   In Number,
    ԭ����id_In In Number,
    �仯ids_In  In Varchar2
  ) Is
    --������ 1����id,1��ҳid:1ԭ����id,1ԭ��ҳid; 2����id,2��ҳid:2ԭ����id,2ԭ��ҳid;��.
  Begin
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_017',
                                '<root><����ID>' || ����id_In || '</����ID><ԭ����ID>' || ԭ����id_In || '</ԭ����ID><CINFO>' ||
                                 �仯ids_In || '</CINFO></root>');
  End Zlhis_Patient_017;

  --69.סԺ����ת�벡��
  Procedure Zlhis_Patient_026
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    v_ת������id Number(18);
    v_ת�벡��id Number(18);
    n_�䶯id     Number(18);
    n_��ʼԭ��   Number(3);
    d_��ʼʱ��   Date;
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_026') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ID, ����id, ��ʼʱ��, ��ʼԭ�� From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0'
      Into n_�䶯id, v_ת�벡��id, d_��ʼʱ��, n_��ʼԭ��
      Using ����id_In, ��ҳid_In;
  
    Execute Immediate 'Select Max(����id) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� = :3 And ��ֹԭ�� = :4 And Nvl(���Ӵ�λ, 0) = 0'
      Into v_ת������id
      Using ����id_In, ��ҳid_In, d_��ʼʱ��, n_��ʼԭ��;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_026',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID>' || '<ת������ID>' ||
                                 v_ת������id || '</ת������ID>' || '<ת�벡��ID>' || v_ת�벡��id || '</ת�벡��ID>' || '<�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID>' || '</root>');
  End Zlhis_Patient_026;

  Procedure Zlhis_Patient_028(����id_In In Number) Is
    v_����     Varchar2(100);
    v_�Ա�     Varchar2(10);
    v_����     Varchar2(20);
    v_�����   Number(18);
    v_���֤�� Varchar2(20);
    v_�������� Varchar2(50);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_PATIENT_028') = 0 Then
      Return;
    End If;
  
    Execute Immediate 'Select ����, �Ա�, ����, To_Char(��������, ''yyyymmdd''), �����, ���֤�� From ������Ϣ Where ����id = :1'
      Into v_����, v_�Ա�, v_����, v_��������, v_�����, v_���֤��
      Using ����id_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_028',
                                '<root><����ID>' || ����id_In || '</����ID><����>' || v_���� || '</����>' || '<�Ա�>' || v_�Ա� ||
                                 '</�Ա�>' || '<����>' || v_���� || '</����>' || '<��������>' || v_�������� || '</��������>' || '<�����>' ||
                                 v_����� || '</�����>' || '<���֤��>' || v_���֤�� || '</���֤��>' || '</root>');
  End Zlhis_Patient_028;

  --79.���۲���תסԺ����
  Procedure Zlhis_Patient_029
  (
    ����id_In In Number,
    ��ҳid_In In Number
  ) Is
    n_�䶯id Number(18);
  Begin
    Execute Immediate 'Select max(ID) From ���˱䶯��¼ Where ����id = :1 And ��ҳid = :2 And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0 And ��ʼԭ�� = 9'
      Into n_�䶯id
      Using ����id_In, ��ҳid_In;
  
    b_Message.p_Msg_Todo_Insert('ZLHIS_PATIENT_029',
                                '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�䶯ID>' || n_�䶯id ||
                                 '</�䶯ID></root>');
  End Zlhis_Patient_029;

  --Ѫ��:������Ѫ���
  Procedure Zlhis_Blood_001(ҽ��id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_001', v_Value);
    End If;
  End Zlhis_Blood_001;

  --Ѫ��:���Ҿܾ���Ѫ
  Procedure Zlhis_Blood_002(ҽ��id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><ҽ��ID>' || ҽ��id_In || '</ҽ��ID></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_BLOOD_002', v_Value);
    End If;
  End Zlhis_Blood_002;

  --70.���鱨�����
  Procedure Zlhis_Lis_001(�걾id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_001', v_Value);
    End If;
  End Zlhis_Lis_001;
  --71.���鱨����˳���
  Procedure Zlhis_Lis_002(�걾id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_002', v_Value);
    End If;
  End Zlhis_Lis_002;
  --73.����걾�����ӡ
  Procedure Zlhis_Lis_004
  (
    ��������_In In Varchar2,
    ҽ��id_In   In Number,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    r_Data  c_Dynamic;
    v_Id    Number(18);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_LIS_004') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004', v_Value);
    Else
      Open r_Data For 'Select ҽ��ID From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(f_Num2list(:1)))'
        Using ҽ��ids_In;
      Loop
        Fetch r_Data
          Into v_Id;
        Exit When r_Data%NotFound;
      
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_004',
                                    '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || v_Id || '</ҽ��ID><ϵͳ>1</ϵͳ></root>');
      End Loop;
    End If;
  End Zlhis_Lis_004;
  --74.����걾�����ӡ����
  Procedure Zlhis_Lis_005
  (
    ��������_In In Varchar2,
    ҽ��id_In   In Number,
    ҽ��ids_In  In Varchar2
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
    r_Data  c_Dynamic;
    v_Id    Number(18);
  Begin
    If b_Zlmsg_Cache.Is_Message_Using('ZLHIS_LIS_005') = 0 Then
      Return;
    End If;
    If ҽ��id_In Is Not Null Then
      v_Value := '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || ҽ��id_In || '</ҽ��ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005', v_Value);
    Else
      Open r_Data For 'Select ҽ��id From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(f_Num2list(:1)))'
        Using ҽ��ids_In;
      Loop
        Fetch r_Data
          Into v_Id;
        Exit When r_Data%NotFound;
      
        b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_005',
                                    '<root><��������>' || ��������_In || '</��������><ҽ��ID>' || v_Id || '</ҽ��ID><ϵͳ>1</ϵͳ></root>');
      End Loop;
    End If;
  End Zlhis_Lis_005;
  --75.����걾����
  Procedure Zlhis_Lis_006(�걾id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_006', v_Value);
    End If;
  End Zlhis_Lis_006;
  --76.����걾���ճ���
  Procedure Zlhis_Lis_007(�걾id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_007', v_Value);
    End If;
  End Zlhis_Lis_007;
  --77.����걾����
  Procedure Zlhis_Lis_008(�걾id_In In Number) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    If �걾id_In Is Not Null Then
      v_Value := '<root><�걾ID>' || �걾id_In || '</�걾ID><ϵͳ>1</ϵͳ></root>';
      b_Message.p_Msg_Todo_Insert('ZLHIS_LIS_008', v_Value);
    End If;
  End Zlhis_Lis_008;

  --��������
  Procedure Zlhis_Emr_018
  (
    ����id_In In Number,
    ��ҳid_In In Number,
    �ļ�id_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><����ID>' || ����id_In || '</����ID><��ҳID>' || ��ҳid_In || '</��ҳID><�ļ�ID>' || �ļ�id_In ||
               '</�ļ�ID></root>';
    b_Message.p_Msg_Todo_Insert('ZLHIS_EMR_018', v_Value);
  End Zlhis_Emr_018;
  --�������ϻ���Ա�䶯��Ϣ
  Procedure Zltools_Users_001
  (
    �û���_In In Varchar2,
    ��Աid_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�û���>' || �û���_In || '</�û���><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_001', v_Value);
  End Zltools_Users_001;
  Procedure Zltools_Users_002
  (
    �û���_In In Varchar2,
    ��Աid_In In Number
  ) Is
    v_Value Zlmsg_Todo.Key_Value%Type;
  Begin
    v_Value := '<root><�û���>' || �û���_In || '</�û���><��ԱID>' || ��Աid_In || '</��ԱID></root>';
    b_Message.p_Msg_Todo_Insert('ZLTOOLS_USERS_002', v_Value);
  End Zltools_Users_002;
End b_Message;
/

----------------------------------------------------------------------------
--[[11.������]]
----------------------------------------------------------------------------



----------------------------------------------------------------------------
--[[16.�ٴ�ҽ��]]
----------------------------------------------------------------------------

Create Or Replace Package Pkg_Zyedit As
  -----------------------------------------------------
  --��ȡ��ҩ����
  -----------------------------------------------------
  Procedure Get_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --��ȡ��ҩ֤��
  -----------------------------------------------------
  Procedure Get_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��ҩ����
  -----------------------------------------------------
  Procedure Get_Fjlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ����б�
  -----------------------------------------------------
  Procedure Get_Diaglist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --��ȡ��ҩ�������
  -----------------------------------------------------
  Procedure Get_Fjitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��֤��֢
  -----------------------------------------------------
  Procedure Get_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��֢�η�
  -----------------------------------------------------
  Procedure Get_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��֢�η���Ӧ��ҩ
  -----------------------------------------------------
  Procedure Get_Additems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��ҩ��ע(ֱ��ȡ�Խ�ϵͳ�Ľ�ע)
  -----------------------------------------------------
  Procedure Get_Jzitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��ҩ�巨(ֱ��ȡ�Խ�ϵͳ�ļ巨)
  -----------------------------------------------------
  Procedure Get_Jftype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��ҩ�÷�(ֱ��ȡ�Խ�ϵͳ���÷�)
  -----------------------------------------------------
  Procedure Get_Usetype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��ҩƵ��(ֱ��ȡ�Խ�ϵͳ��Ƶ��)
  -----------------------------------------------------
  Procedure Get_Usetime
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���õķ�ҩҩ��(ֱ��ȡ�Խ�ϵͳ�ķ�ҩҩ��)
  -----------------------------------------------------
  Procedure Get_Drugdept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡҩƷ���(ֱ��ȡ�Խ�ϵͳ��ҩƷ���)
  -----------------------------------------------------
  Procedure Get_Drugstock
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --������ҽ�����Ϣ
  -----------------------------------------------------
  Procedure Load_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --���ز�ҩ��ϸ
  -----------------------------------------------------
  Procedure Load_Zyinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --������ҽ�����Ϣ
  -----------------------------------------------------
  Procedure Save_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --ɾ����ҽ���(����ɾ��)
  -----------------------------------------------------
  Procedure Del_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --ͨ��HISҽ��ID��ȡ����ID�����ID
  -----------------------------------------------------
  Procedure Get_Diagid
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���з���
  -----------------------------------------------------
  Procedure Get_Fjall
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��ҩĿ¼
  -----------------------------------------------------
  Procedure Get_Drugitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡHISƷ���б�
  -----------------------------------------------------
  Procedure Get_Hisdrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�޸Ĳ�ҩ��Ϣ
  -----------------------------------------------------
  Procedure Save_Drugitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --HISƷ����������
  -----------------------------------------------------
  Procedure Set_Autodrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�޸ķ�����Ϣ
  -----------------------------------------------------
  Procedure Save_Fjitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�޸���ҽ����
  -----------------------------------------------------
  Procedure Set_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�޸���ҽ֤��
  -----------------------------------------------------
  Procedure Set_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --����֤�ͷ�����Ӧ
  -----------------------------------------------------
  Procedure Set_Zxtofj
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�޸���֤��֢
  -----------------------------------------------------
  Procedure Set_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�޸ļ�֢�η�
  -----------------------------------------------------
  Procedure Set_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�����η���ҩ��Ӧ
  -----------------------------------------------------
  Procedure Set_Zftozy
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��Ŀɾ������
  -----------------------------------------------------
  Procedure Del_Zydata
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --��ȡ���ݿ�ϵͳʱ��
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�쳣����
  -----------------------------------------------------
  Procedure Errorcenter
  (
    Err_Num In Number,
    Err_Msg In Varchar2
  );

End Pkg_Zyedit;
/
Create Or Replace Package Body Pkg_Zyedit As
  -----------------------------------------------------
  --��ȡ��ҩ����
  -----------------------------------------------------
  Procedure Get_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.����id As ID, a.�Ʊ�, a.��������, a.���� From ��ҽ���� A Order By a.�Ʊ�, a.��������;
  End Get_Distype;

  -----------------------------------------------------
  --��ȡ��ҩ֤��
  -----------------------------------------------------
  Procedure Get_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_����id ��ҽ֤��.����id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_����id := To_Number(Jsonobj.Get_String('����ID'));
  
    Open Output_Out For
      Select a.֤��id As ID, a.֤������, a.����, a.֤���η�, a.֤������, a.֢״����
      From ��ҽ֤�� A
      Where a.����id = n_����id
      Order By a.֤������;
  End Get_Zxtype;

  -----------------------------------------------------
  --��ȡ��ҩ����
  -----------------------------------------------------
  Procedure Get_Fjlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_֤��id  ֤�ͷ�������.֤��id%Type;
    v_ƥ����  Varchar2(100);
    n_Usetype Number;
  Begin
    Jsonobj   := Pljson(Input_In);
    n_֤��id  := To_Number(Nvl(Jsonobj.Get_String('֤��ID'), 0));
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_֤��id = 0 Then
      v_ƥ���� := '%' || Upper(Jsonobj.Get_String('ƥ����')) || '%';
    
      If n_Usetype = 0 Then
        Open Output_Out For
          Select b.����id As ID, b.��������, b.���� As ����, b.��Դ, Decode(Nvl(b.�Ƿ���, 0), 0, b.���ժҪ, '���ܷ���') As ���ժҪ, b.��������,
                 b.��Ӧ֤����, b.����, b.��������, b.�Ƿ���
          From �η����� B
          Where b.���� Like v_ƥ���� Or b.�������� Like v_ƥ���� Or b.���� Like v_ƥ���� Or b.�������� Like v_ƥ����
          Order By b.��������;
      Else
        Open Output_Out For
          Select b.����id As ID, b.��������, b.���� As ����, b.��Դ, b. ���ժҪ, b.��������, b.��Ӧ֤����, b.����, b.��������
          From �η����� B
          Where b.���� Like v_ƥ���� Or b.�������� Like v_ƥ���� Or b.���� Like v_ƥ���� Or b.�������� Like v_ƥ����
          Order By b.��������;
      End If;
    Else
      If n_Usetype = 0 Then
        Open Output_Out For
          Select a.����id As ID, b.��������, b.���� As ����, b.����, b.��������, b.��Դ, b.��������,
                 Decode(Nvl(b.�Ƿ���, 0), 0, b.���ժҪ, '���ܷ���') As ���ժҪ, b.��Ӧ֤����, b.�Ƿ���
          From ֤�ͷ������� A, �η����� B
          Where a.����id = b.����id And a.֤��id = n_֤��id And a.״̬ = 1
          Order By a.����id, b.��������;
      Else
        Open Output_Out For
          Select a.����id As ID, b.��������, b.���� As ����, b.����, b.��������, b.��Դ, b.��������, b.���ժҪ, b.��Ӧ֤����, a.״̬, a.����id
          From ֤�ͷ������� A, �η����� B
          Where a.����id = b.����id And a.֤��id = n_֤��id
          Order By -a.״̬, a.����id;
      End If;
    End If;
  End Get_Fjlist;

  -----------------------------------------------------
  --��ȡ����б�
  -----------------------------------------------------
  Procedure Get_Diaglist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    v_ƥ���� Varchar2(100);
  Begin
    Jsonobj := Pljson(Input_In);
  
    v_ƥ���� := '%' || Upper(Jsonobj.Get_String('ƥ����')) || '%';
    Open Output_Out For
      Select a.֤��id As ID, b.�������� || '-' || a.֤������ As �������, b.���� || a.���� As ����, a.֤���η� As �η�, a.֤������ As ����, a.֢״����
      From ��ҽ֤�� A, ��ҽ���� B
      Where a.����id = b.����id And (b.�������� || '-' || a.֤������ Like v_ƥ���� Or b.���� || a.���� Like v_ƥ����)
      Order By b.����id;
  End Get_Diaglist;

  -----------------------------------------------------
  --��ȡ��ҩ�������
  -----------------------------------------------------
  Procedure Get_Fjitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_����id ��������.����id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_����id := To_Number(Nvl(Jsonobj.Get_String('����ID'), 0));
  
    Open Output_Out For
      Select b.����id, b.��ҩid, a.��ҩ����, b.�÷���ע, b.�ŷ�����, b.����, a.��λ, Nvl(a.HisƷ��id, 0) As HisƷ��id
      From �������� B, ��ҩĿ¼ A
      Where b.��ҩid = a.��ҩid And b.����id = n_����id
      Order By b.����id;
  End Get_Fjitems;

  -----------------------------------------------------
  --��ȡ��֤��֢
  -----------------------------------------------------
  Procedure Get_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_Usetype Number; --0 ��Ч����/-1 ȫ������
  Begin
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_Usetype = 0 Then
      Open Output_Out For
        Select b.��֢id As ID, b.��֢����, b.���� From ��֤��֢ B Where b.״̬ = 1 Order By b.��֢id;
    Else
      Open Output_Out For
        Select b.״̬, b.��֢id As ID, b.��֢����, b.���� From ��֤��֢ B Order By -b.״̬, b.��֢id;
    End If;
  End Get_Adddis;

  -----------------------------------------------------
  --��ȡ��֢�η�
  -----------------------------------------------------
  Procedure Get_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_��֢id  ��֢�η�.��֢id%Type;
    n_Usetype Number; --0 ��Ч����/-1 ȫ������
  Begin
    Jsonobj   := Pljson(Input_In);
    n_��֢id  := To_Number(Nvl(Jsonobj.Get_String('��֢ID'), 0));
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_Usetype = 0 Then
      Open Output_Out For
        Select b.�η�id As ID, b.�η�����, b.����
        From ��֢�η� B
        Where b.��֢id = n_��֢id And b.״̬ = 1
        Order By b.�η�id;
    Else
      Open Output_Out For
        Select b.�η�id As ID, b.�η�����, b.����, b.״̬, b.��֢id
        From ��֢�η� B
        Where b.��֢id = n_��֢id
        Order By -b.״̬, b.�η�id;
    End If;
  End Get_Addzf;

  -----------------------------------------------------
  --��ȡ��֢�η���Ӧ��ҩ
  -----------------------------------------------------
  Procedure Get_Additems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_�η�id  ��֢��ҩ.�η�id%Type;
    n_Usetype Number; --0 ��Ч����/-1 ȫ������
  Begin
    Jsonobj   := Pljson(Input_In);
    n_�η�id  := To_Number(Nvl(Jsonobj.Get_String('�η�ID'), 0));
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    If n_Usetype = 0 Then
      Open Output_Out For
        Select b.��ҩid As ID, a.��ҩ����, a.����, b.����, a.��λ, Nvl(a.HisƷ��id, 0) As HisƷ��id
        From ��ҩĿ¼ A, ��֢��ҩ B
        Where a.��ҩid = b.��ҩid And b.�η�id = n_�η�id And b.״̬ = 1
        Order By b.��ҩid;
    Else
      Open Output_Out For
        Select b.��ҩid, a.��ҩ����, a.����, b.����, a.��λ, b.״̬, b.��ҩid As ID, b.�η�id
        From ��ҩĿ¼ A, ��֢��ҩ B
        Where a.��ҩid = b.��ҩid And b.�η�id = n_�η�id
        Order By -b.״̬, b.��ҩid;
    End If;
  
  End Get_Additems;

  -----------------------------------------------------
  --��ȡ��ҩ��ע(ֱ��ȡ�Խ�ϵͳ�Ľ�ע)
  -----------------------------------------------------
  Procedure Get_Jzitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(2, Null, Output_Out);
  End Get_Jzitems;

  -----------------------------------------------------
  --��ȡ��ҩ�巨(ֱ��ȡ�Խ�ϵͳ�ļ巨)
  -----------------------------------------------------
  Procedure Get_Jftype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(1, Input_In, Output_Out);
  End Get_Jftype;

  -----------------------------------------------------
  --��ȡ��ҩ�÷�(ֱ��ȡ�Խ�ϵͳ���÷�)
  -----------------------------------------------------
  Procedure Get_Usetype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(3, Input_In, Output_Out);
  End Get_Usetype;

  -----------------------------------------------------
  --��ȡ��ҩƵ��(ֱ��ȡ�Խ�ϵͳ���÷�)
  -----------------------------------------------------
  Procedure Get_Usetime
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(4, Input_In, Output_Out);
  End Get_Usetime;

  -----------------------------------------------------
  --��ȡ���õķ�ҩҩ��(ֱ��ȡ�Խ�ϵͳ�ķ�ҩҩ��)
  -----------------------------------------------------
  Procedure Get_Drugdept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(5, Input_In, Output_Out);
  End Get_Drugdept;

  -----------------------------------------------------
  --��ȡҩƷ���(ֱ��ȡ�Խ�ϵͳ��ҩƷ���)
  -----------------------------------------------------
  Procedure Get_Drugstock
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(6, Input_In, Output_Out);
  End Get_Drugstock;

  -----------------------------------------------------
  --������ҽ�����Ϣ
  -----------------------------------------------------
  Procedure Load_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_���id ������ҽ��ϼ�¼.���id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_���id := To_Number(Jsonobj.Get_String('���ID'));
  
    Open Output_Out For
      Select a.����id, a.����id, a.��������, a.����, a.��ҩ�÷�, a.��ҩ�巨, a.����, a.��ҩƵ��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.ҽ������, a.His�巨id,
             a.His�÷�id, a.Hisҩ��id, b.���id, b.���﷽ʽ, b.�Ʊ�, b.����id, b.��������, b.֤��id, b.֤������, b.��ҽ���, b.��ҽ�η�, b.����ʱ��, b.������,
             b.His���id, b.Hisҽ��id, Nvl(c.�Ƿ���, 0) As �Ƿ���
      From ������ҽ������¼ A, ������ҽ��ϼ�¼ B, �η����� C
      Where b.����id = a.����id And a.����id = c.����id And b.���id = n_���id;
  End Load_Zyedit;

  -----------------------------------------------------
  --���ز�ҩ��ϸ
  -----------------------------------------------------
  Procedure Load_Zyinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_����id ������ҽ��ϼ�¼.����id%Type;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_����id := To_Number(Jsonobj.Get_String('����ID'));
  
    Open Output_Out For
      Select a.������ϸid, a.����id, a.���, a.��ҩid, a.�Ƿ��ҩ, a.��Դ, a.��ҩ����, a.����, a.��λ, a.��ע, a.HisƷ��id, a.His���id
      From ������ҽ������ϸ A
      Where a.����id = n_����id
      Order By a.���;
  End Load_Zyinfo;

  -----------------------------------------------------
  --������ҽ�����Ϣ
  -----------------------------------------------------
  Procedure Save_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/1-�޸�
    n_����id     ������ҽ��ϼ�¼.����id%Type;
    v_�Һŵ�     ������ҽ��ϼ�¼.�Һŵ�%Type;
    n_�����     ������ҽ��ϼ�¼.�����%Type;
    n_���id     ������ҽ��ϼ�¼.���id%Type;
    n_����id     ������ҽ������¼.Hisҩ��id%Type;
    v_��������   Varchar2(100);
    n_����Աid   Number;
    v_����Ա���� ������ҽ��ϼ�¼.������%Type;
    v_����       ������ҽ��ϼ�¼.����%Type;
    v_�Ա�       ������ҽ��ϼ�¼.�Ա�%Type;
    v_����       ������ҽ��ϼ�¼.����%Type;
    v_����       ������ҽ��ϼ�¼.����%Type;
    v_��������   Varchar2(100);
    n_���﷽ʽ   ������ҽ��ϼ�¼.���﷽ʽ%Type;
    v_�Ʊ�       ������ҽ��ϼ�¼.�Ʊ�%Type;
    n_����id     ������ҽ��ϼ�¼.����id%Type;
    v_��������   ������ҽ��ϼ�¼.��������%Type;
    n_֤��id     ������ҽ��ϼ�¼.֤��id%Type;
    v_֤������   ������ҽ��ϼ�¼.֤������%Type;
    v_��ҽ���   ������ҽ��ϼ�¼.��ҽ���%Type;
    v_��ҽ�η�   ������ҽ��ϼ�¼.��ҽ�η�%Type;
    n_����id     ������ҽ������¼.����id%Type;
    v_��������   ������ҽ������¼.��������%Type;
    n_����       ������ҽ������¼.����%Type;
    v_��ҩ�÷�   ������ҽ������¼.��ҩ�÷�%Type;
    n_His�÷�id  ������ҽ������¼.His�÷�id%Type;
    v_��ҩ�巨   ������ҽ������¼.��ҩ�巨%Type;
    n_His�巨id  ������ҽ������¼.His�巨id%Type;
    v_����       ������ҽ������¼.����%Type;
    v_��ҩƵ��   ������ҽ������¼.��ҩƵ��%Type;
    n_Ƶ�ʴ���   ������ҽ������¼.Ƶ�ʴ���%Type;
    n_Ƶ�ʼ��   ������ҽ������¼.Ƶ�ʼ��%Type;
    v_�����λ   ������ҽ������¼.�����λ%Type;
    v_ҽ������   ������ҽ������¼.ҽ������%Type;
    n_Hisҩ��id  ������ҽ������¼.Hisҩ��id%Type;
  
    n_��ҩid    ������ҽ������ϸ.��ҩid%Type;
    n_�Ƿ��ҩ  ������ҽ������ϸ.�Ƿ��ҩ%Type;
    v_��Դ      ������ҽ������ϸ.��Դ%Type;
    v_��ҩ����  ������ҽ������ϸ.��ҩ����%Type;
    n_����      ������ҽ������ϸ.����%Type;
    v_��λ      ������ҽ������ϸ.��λ%Type;
    v_��ע      ������ҽ������ϸ.��ע%Type;
    n_HisƷ��id ������ҽ������ϸ.HisƷ��id%Type;
    n_His���id ������ҽ������ϸ.His���id%Type;
  
    n_Hisҽ��id   Number;
    n_His���id   Number;
    n_�Ƿ���    Number;
    v_��ҽ���old ������ҽ��ϼ�¼.��ҽ���%Type;
  
    n_����id    ������ҽ������¼.����id%Type;
    d_Now       Date;
    Jsonlistobj Pljson_List;
  
    v_Out Clob;
  
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
    --�������
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_����id     := To_Number(Jsonobj.Get_String('����ID'));
    v_�Һŵ�     := Jsonobj.Get_String('�Һŵ�');
    n_�����     := To_Number(Jsonobj.Get_String('�����'));
    n_���id     := To_Number(Jsonobj.Get_String('���ID'));
    n_����id     := To_Number(Jsonobj.Get_String('����ID'));
    v_��������   := Jsonobj.Get_String('��������');
    n_����Աid   := To_Number(Jsonobj.Get_String('����ԱID'));
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    v_����       := Jsonobj.Get_String('����');
    v_�Ա�       := Jsonobj.Get_String('�Ա�');
    v_����       := Jsonobj.Get_String('����');
    v_����       := Jsonobj.Get_String('����');
    v_��������   := Jsonobj.Get_String('��������');
    n_���﷽ʽ   := To_Number(Jsonobj.Get_String('���﷽ʽ'));
    v_�Ʊ�       := Jsonobj.Get_String('�Ʊ�');
    n_����id     := To_Number(Jsonobj.Get_String('����ID'));
    v_��������   := Jsonobj.Get_String('��������');
    n_֤��id     := To_Number(Jsonobj.Get_String('֤��ID'));
    v_֤������   := Jsonobj.Get_String('֤������');
    v_��ҽ���   := Jsonobj.Get_String('��ҽ���');
    v_��ҽ�η�   := Jsonobj.Get_String('��ҽ�η�');
    n_����id     := To_Number(Jsonobj.Get_String('����ID'));
    v_��������   := Jsonobj.Get_String('��������');
    n_����       := To_Number(Jsonobj.Get_String('����'));
    v_��ҩ�÷�   := Jsonobj.Get_String('��ҩ�÷�');
    n_His�÷�id  := To_Number(Jsonobj.Get_String('HIS�÷�ID'));
    v_��ҩ�巨   := Jsonobj.Get_String('��ҩ�巨');
    n_His�巨id  := To_Number(Jsonobj.Get_String('HIS�巨ID'));
    v_����       := Jsonobj.Get_String('����');
    v_��ҩƵ��   := Jsonobj.Get_String('��ҩƵ��');
    n_Ƶ�ʴ���   := To_Number(Jsonobj.Get_String('Ƶ�ʴ���'));
    n_Ƶ�ʼ��   := To_Number(Jsonobj.Get_String('Ƶ�ʼ��'));
    v_�����λ   := Jsonobj.Get_String('�����λ');
    v_ҽ������   := Jsonobj.Get_String('ҽ������');
    n_Hisҩ��id  := To_Number(Jsonobj.Get_String('HISҩ��ID'));
    Jsonlistobj  := Jsonobj.Get_Pljson_List('������ϸ');
  
    Select Sysdate Into d_Now From Dual;
  
    --����
    If n_Usetype = 0 Then
      --�༭����
      Select ������ҽ��ϼ�¼_���id.Nextval Into n_���id From Dual;
      Select ������ҽ������¼_����id.Nextval Into n_����id From Dual;
    
      Insert Into ������ҽ������¼
        (����id, ����id, ��������, ����, ��ҩ�÷�, ��ҩ�巨, ����, ��ҩƵ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ҽ������, His�巨id, His�÷�id, Hisҩ��id)
      Values
        (n_����id, n_����id, v_��������, n_����, v_��ҩ�÷�, v_��ҩ�巨, v_����, v_��ҩƵ��, n_Ƶ�ʴ���, n_Ƶ�ʼ��, v_�����λ, v_ҽ������, n_His�巨id,
         n_His�÷�id, n_Hisҩ��id);
    
      Insert Into ������ҽ��ϼ�¼
        (���id, ����id, �Һŵ�, ����, �����, �Ա�, ����, ����, ��������, ���﷽ʽ, �Ʊ�, ����id, ��������, ֤��id, ֤������, ��ҽ���, ��ҽ�η�, ����id, ����ʱ��, ������)
      Values
        (n_���id, n_����id, v_�Һŵ�, v_����, n_�����, v_�Ա�, v_����, v_����, To_Date(v_��������, 'yyyy-mm-dd'), n_���﷽ʽ, v_�Ʊ�, n_����id,
         v_��������, n_֤��id, v_֤������, v_��ҽ���, v_��ҽ�η�, n_����id, d_Now, v_����Ա����);
    
      For I In 1 .. Jsonlistobj.Count Loop
        Jsonobj     := Pljson();
        Jsonobj     := Pljson(Jsonlistobj.Get(I));
        n_��ҩid    := To_Number(Jsonobj.Get_String('��ҩID'));
        n_�Ƿ��ҩ  := To_Number(Jsonobj.Get_String('�Ƿ��ҩ'));
        v_��Դ      := Jsonobj.Get_String('��Դ');
        v_��ҩ����  := Jsonobj.Get_String('��ҩ����');
        n_����      := To_Number(Jsonobj.Get_String('����'));
        v_��λ      := Jsonobj.Get_String('��λ');
        v_��ע      := Jsonobj.Get_String('��ע');
        n_HisƷ��id := To_Number(Jsonobj.Get_String('HISƷ��ID'));
        n_His���id := To_Number(Jsonobj.Get_String('HIS���ID'));
      
        Insert Into ������ҽ������ϸ
          (������ϸid, ����id, ���, ��ҩid, �Ƿ��ҩ, ��Դ, ��ҩ����, ����, ��λ, ��ע, HisƷ��id, His���id)
        Values
          (������ҽ������ϸ_������ϸid.Nextval, n_����id, I, n_��ҩid, n_�Ƿ��ҩ, v_��Դ, v_��ҩ����, n_����, v_��λ, v_��ע, n_HisƷ��id, n_His���id);
      End Loop;
    Else
      Select Max(����id), Max(Hisҽ��id), Max(His���id), Max(��ҽ���)
      Into n_����id, n_Hisҽ��id, n_His���id, v_��ҽ���old
      From ������ҽ��ϼ�¼
      Where ���id = n_���id;
    
      If Nvl(n_���id, 0) = 0 Or Nvl(n_����id, 0) = 0 Then
        v_Err_Msg := 'δ�ҵ�������϶�Ӧ�Ĵ������ݡ�';
        Raise Err_Item;
      End If;
    
      Update ������ҽ������¼
      Set ����id = n_����id, �������� = v_��������, ���� = n_����, ��ҩ�÷� = v_��ҩ�÷�, ��ҩ�巨 = v_��ҩ�巨, ���� = v_����, ��ҩƵ�� = v_��ҩƵ��, Ƶ�ʴ��� = n_Ƶ�ʴ���,
          Ƶ�ʼ�� = n_Ƶ�ʼ��, �����λ = v_�����λ, ҽ������ = v_ҽ������, His�巨id = n_His�巨id, His�÷�id = n_His�÷�id, Hisҩ��id = n_Hisҩ��id
      Where ����id = n_����id;
    
      Update ������ҽ��ϼ�¼
      Set ����id = n_����id, �Һŵ� = v_�Һŵ�, ���� = v_����, ����� = n_�����, �Ա� = v_�Ա�, ���� = v_����, ���� = v_����,
          �������� = To_Date(v_��������, 'yyyy-mm-dd'), ���﷽ʽ = n_���﷽ʽ, �Ʊ� = v_�Ʊ�, ����id = n_����id, �������� = v_��������, ֤��id = n_֤��id,
          ֤������ = v_֤������, ��ҽ��� = v_��ҽ���, ��ҽ�η� = v_��ҽ�η�, ����id = n_����id, ����ʱ�� = d_Now, ������ = v_����Ա����
      
      Where ���id = n_���id;
    
      Delete From ������ҽ������ϸ Where ����id = n_����id;
    
      For I In 1 .. Jsonlistobj.Count Loop
        Jsonobj     := Pljson();
        Jsonobj     := Pljson(Jsonlistobj.Get(I));
        n_��ҩid    := To_Number(Jsonobj.Get_String('��ҩID'));
        n_�Ƿ��ҩ  := To_Number(Jsonobj.Get_String('�Ƿ��ҩ'));
        v_��Դ      := Jsonobj.Get_String('��Դ');
        v_��ҩ����  := Jsonobj.Get_String('��ҩ����');
        n_����      := To_Number(Jsonobj.Get_String('����'));
        v_��λ      := Jsonobj.Get_String('��λ');
        v_��ע      := Jsonobj.Get_String('��ע');
        n_HisƷ��id := To_Number(Jsonobj.Get_String('HISƷ��ID'));
        n_His���id := To_Number(Jsonobj.Get_String('HIS���ID'));
      
        Insert Into ������ҽ������ϸ
          (������ϸid, ����id, ���, ��ҩid, �Ƿ��ҩ, ��Դ, ��ҩ����, ����, ��λ, ��ע, HisƷ��id, His���id)
        Values
          (������ҽ������ϸ_������ϸid.Nextval, n_����id, I, n_��ҩid, n_�Ƿ��ҩ, v_��Դ, v_��ҩ����, n_����, v_��λ, v_��ע, n_HisƷ��id, n_His���id);
      End Loop;
    End If;
  
    Select Nvl(Max(�Ƿ���), 0) Into n_�Ƿ��� From �η����� Where ����id = n_����id;
  
    --ͬ��HISҽ�����
    Zl_��ҽ���_Save(Input_In, n_Hisҽ��id, n_His���id, v_��ҽ���old, n_�Ƿ���, v_Out);
    If Nvl(v_Out, '��') != '��' Then
      Jsonobj := Pljson();
      Jsonobj := Pljson(v_Out);
      If To_Number(Jsonobj.Get_String('His���id')) != 0 And To_Number(Nvl(Jsonobj.Get_String('His���id'), 0)) != 0 Then
        Update ������ҽ��ϼ�¼
        Set Hisҽ��id = To_Number(Nvl(Jsonobj.Get_String('Hisҽ��id'), 0)),
            His���id = To_Number(Nvl(Jsonobj.Get_String('His���id'), 0))
        Where ���id = n_���id;
      Else
        v_Err_Msg := '��ҽ��ϱ���ʧ��,����HISͬ���ӿڡ�';
        Raise Err_Item;
      End If;
    Else
      v_Err_Msg := '��ҽ��ϱ���ʧ��,����HISͬ���ӿڡ�';
      Raise Err_Item;
    End If;
  
    Open Output_Out For
      Select To_Number(Nvl(Jsonobj.Get_String('Hisҽ��id'), 0)) As Hisҽ��id,
             To_Number(Nvl(Jsonobj.Get_String('His���id'), 0)) As His���id, n_���id As ���id, n_����id As ����id
      From Dual;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End Save_Zyedit;

  -----------------------------------------------------
  --ɾ����ҽ���(����ɾ��)
  -----------------------------------------------------
  Procedure Del_Zyedit
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj  Pljson;
    n_���id ������ҽ��ϼ�¼.���id%Type;
    n_����id ������ҽ��ϼ�¼.����id%Type;
  
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
    Jsonobj  := Pljson(Input_In);
    n_���id := To_Number(Jsonobj.Get_String('���ID'));
  
    Select Max(����id) Into n_����id From ������ҽ��ϼ�¼ Where ���id = n_���id;
  
    If Nvl(n_���id, 0) = 0 Or Nvl(n_����id, 0) = 0 Then
      v_Err_Msg := 'δ�ҵ�������϶�Ӧ�Ĵ������ݡ�';
      Raise Err_Item;
    End If;
  
    Delete From ������ҽ������¼ Where ����id = n_����id;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  End Del_Zyedit;

  -----------------------------------------------------
  --ͨ��HISҽ��ID����HIS���ID��ȡ����ID�����ID
  -----------------------------------------------------
  Procedure Get_Diagid
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj     Pljson;
    n_Hisҽ��id ������ҽ��ϼ�¼.Hisҽ��id%Type;
    n_His���id ������ҽ��ϼ�¼.His���id%Type;
  Begin
    Jsonobj     := Pljson(Input_In);
    n_Hisҽ��id := To_Number(Jsonobj.Get_String('HISҽ��ID'));
    n_His���id := To_Number(Jsonobj.Get_String('HIS���ID'));
    Open Output_Out For
      Select a.����id, a.���id From ������ҽ��ϼ�¼ A Where a.Hisҽ��id = n_Hisҽ��id Or a.His���id = n_His���id;
  End Get_Diagid;

  -----------------------------------------------------
  --��ȡ��ҩĿ¼
  -----------------------------------------------------
  Procedure Get_Drugitems
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.��ҩid As ID, a.��ҩ����, a.����, a.����, a.��������, a.��λ, Null As His����, a.��Դ, a.HisƷ��id, a.��ҩ����, a.��״, a.ҩ��, a.��Ӧ֤,
             a.�÷�, a.����, a.����, a.�ɷ�, a.ҩ������, a.������, To_Char(a.����ʱ��, 'yyyy-MM-dd hh24:mi') As ����ʱ��, a.����޸���,
             To_Char(a.����޸�ʱ��, 'yyyy-MM-dd hh24:mi') As ����޸�ʱ��
      From ��ҩĿ¼ A
      Order By a.��ҩid, a.��ҩ����;
  End Get_Drugitems;

  -----------------------------------------------------
  --��ȡ���з���
  -----------------------------------------------------
  Procedure Get_Fjall
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����id As ID, ��������, ����, ����, ��������, ��Դ, ���ժҪ, ��������, ��������, �Ʒ�����, ��Ӧ֤����, ���������������, ������,
             To_Char(����ʱ��, 'yyyy-MM-dd hh24:mi') As ����ʱ��, ����޸���, To_Char(����޸�ʱ��, 'yyyy-MM-dd hh24:mi') As ����޸�ʱ��,
             Nvl(�Ƿ���, 0) As �Ƿ���, Decode(Nvl(�Ƿ���, 0), 1, '��', '') As ��
      From �η�����
      Order By ����id, ��������;
  End Get_Fjall;

  -----------------------------------------------------
  --��ȡHISƷ���б�
  -----------------------------------------------------
  Procedure Get_Hisdrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(7, Input_In, Output_Out);
  End Get_Hisdrug;

  -----------------------------------------------------
  --�޸Ĳ�ҩ��Ϣ
  -----------------------------------------------------
  Procedure Save_Drugitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/1-�޸�/2-�޸�HIS����
    n_��ҩid     ��ҩĿ¼.��ҩid%Type;
    v_��ҩ����   ��ҩĿ¼.��ҩ����%Type;
    v_����       ��ҩĿ¼.����%Type;
    v_����       ��ҩĿ¼.����%Type;
    v_��������   ��ҩĿ¼.��������%Type;
    v_��Դ       ��ҩĿ¼.��Դ%Type;
    v_��λ       ��ҩĿ¼.��λ%Type;
    v_��ҩ����   ��ҩĿ¼.��ҩ����%Type;
    v_��״       ��ҩĿ¼.��״%Type;
    v_ҩ��       ��ҩĿ¼.ҩ��%Type;
    v_��Ӧ֤     ��ҩĿ¼.��Ӧ֤%Type;
    v_�÷�       ��ҩĿ¼.�÷�%Type;
    v_����       ��ҩĿ¼.����%Type;
    v_����       ��ҩĿ¼.����%Type;
    v_�ɷ�       ��ҩĿ¼.�ɷ�%Type;
    v_ҩ������   ��ҩĿ¼.ҩ������%Type;
    n_HisƷ��id  ��ҩĿ¼.HisƷ��id%Type;
    v_����Ա���� ��ҩĿ¼.����޸���%Type;
    n_����Աid   Number;
  Begin
  
    --�������
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_��ҩid     := To_Number(Nvl(Jsonobj.Get_String('��ҩID'), 0));
    v_��ҩ����   := Jsonobj.Get_String('��ҩ����');
    v_����       := Jsonobj.Get_String('����');
    v_����       := Jsonobj.Get_String('����');
    v_��������   := Jsonobj.Get_String('��������');
    v_��λ       := Jsonobj.Get_String('��λ');
    v_��Դ       := Jsonobj.Get_String('��Դ');
    v_��ҩ����   := Jsonobj.Get_String('��ҩ����');
    v_��״       := Jsonobj.Get_String('��״');
    v_ҩ��       := Jsonobj.Get_String('ҩ��');
    v_��Ӧ֤     := Jsonobj.Get_String('��Ӧ֤');
    v_�÷�       := Jsonobj.Get_String('�÷�');
    v_����       := Jsonobj.Get_String('����');
    v_����       := Jsonobj.Get_String('����');
    v_�ɷ�       := Jsonobj.Get_String('�ɷ�');
    v_ҩ������   := Jsonobj.Get_String('ҩ������');
    n_HisƷ��id  := To_Number(Nvl(Jsonobj.Get_String('HISƷ��ID'), 0));
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    n_����Աid   := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    If n_HisƷ��id = 0 Then
      n_HisƷ��id := Null;
    End If;
    If n_Usetype = 0 Then
      Select ��ҩĿ¼_��ҩid.Nextval Into n_��ҩid From Dual;
      Insert Into ��ҩĿ¼
        (��ҩid, ��ҩ����, ����, ����, ��������, ��λ, ��Դ, ��ҩ����, ��״, ҩ��, ��Ӧ֤, �÷�, ����, ����, �ɷ�, ҩ������, HisƷ��id, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_��ҩid, v_��ҩ����, v_����, v_����, v_��������, v_��λ, v_��Դ, v_��ҩ����, v_��״, v_ҩ��, v_��Ӧ֤, v_�÷�, v_����, v_����, v_�ɷ�, v_ҩ������,
         n_HisƷ��id, v_����Ա����, Sysdate, v_����Ա����, Sysdate);
    Elsif n_Usetype = 1 Then
      Update ��ҩĿ¼
      Set ��ҩ���� = v_��ҩ����, ���� = v_����, ���� = v_����, �������� = v_��������, ��λ = v_��λ, ��Դ = v_��Դ, ��ҩ���� = v_��ҩ����, ��״ = v_��״, ҩ�� = v_ҩ��,
          ��Ӧ֤ = v_��Ӧ֤, �÷� = v_�÷�, ���� = v_����, ���� = v_����, �ɷ� = v_�ɷ�, ҩ������ = v_ҩ������, HisƷ��id = n_HisƷ��id, ����޸��� = v_����Ա����,
          ����޸�ʱ�� = Sysdate
      Where ��ҩid = n_��ҩid;
    Elsif n_Usetype = 2 Then
      Update ��ҩĿ¼
      Set HisƷ��id = n_HisƷ��id, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate
      Where ��ҩid = n_��ҩid;
    End If;
    Open Output_Out For
      Select n_��ҩid As ��ҩid From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Save_Drugitem;

  -----------------------------------------------------
  --HISƷ����������
  -----------------------------------------------------
  Procedure Set_Autodrug
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Zl_��ҩ����_Edit(8, Input_In, Output_Out);
  End Set_Autodrug;

  -----------------------------------------------------
  --�޸ķ�����Ϣ
  -----------------------------------------------------
  Procedure Save_Fjitem
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj            Pljson;
    n_Usetype          Number; --0-����/1-�޸�
    n_����id           �η�����.����id%Type;
    v_��������         �η�����.��������%Type;
    v_����             �η�����.����%Type;
    v_����             �η�����.����%Type;
    v_��������         �η�����.��������%Type;
    v_��Դ             �η�����.��Դ%Type;
    v_���ժҪ         �η�����.���ժҪ%Type;
    v_��������         �η�����.��������%Type;
    v_��������         �η�����.��������%Type;
    v_�Ʒ�����         �η�����.�Ʒ�����%Type;
    v_��Ӧ֤����       �η�����.��Ӧ֤����%Type;
    v_��������������� �η�����.���������������%Type;
    n_�Ƿ���         �η�����.�Ƿ���%Type;
    v_����Ա����       �η�����.����޸���%Type;
    n_����Աid         Number;
    Jsonlistobj        Pljson_List;
  
    n_��ҩid   ��������.��ҩid%Type;
    n_����     ��������.����%Type;
    v_�÷���ע ��������.�÷���ע%Type;
    v_�ŷ����� ��������.�ŷ�����%Type;
  Begin
  
    --�������
    Jsonobj            := Pljson(Input_In);
    n_Usetype          := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_����id           := To_Number(Nvl(Jsonobj.Get_String('����ID'), 0));
    v_��������         := Jsonobj.Get_String('��������');
    v_����             := Jsonobj.Get_String('����');
    v_����             := Jsonobj.Get_String('����');
    v_��������         := Jsonobj.Get_String('��������');
    v_��Դ             := Jsonobj.Get_String('��Դ');
    v_���ժҪ         := Jsonobj.Get_String('���ժҪ');
    v_��������         := Jsonobj.Get_String('��������');
    v_��������         := Jsonobj.Get_String('��������');
    v_�Ʒ�����         := Jsonobj.Get_String('�Ʒ�����');
    v_��Ӧ֤����       := Jsonobj.Get_String('��Ӧ֤����');
    v_��������������� := Jsonobj.Get_String('���������������');
    n_�Ƿ���         := To_Number(Nvl(Jsonobj.Get_String('�Ƿ���'), 0));
    v_����Ա����       := Jsonobj.Get_String('����Ա����');
    n_����Աid         := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    Jsonlistobj        := Jsonobj.Get_Pljson_List('��������');
  
    If n_Usetype = 0 Then
      Select �η�����_����id.Nextval Into n_����id From Dual;
      Insert Into �η�����
        (����id, ��������, ����, ����, ��������, ��Դ, ���ժҪ, ��������, ��������, �Ʒ�����, ��Ӧ֤����, ���������������, �Ƿ���, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_����id, v_��������, v_����, v_����, v_��������, v_��Դ, v_���ժҪ, v_��������, v_��������, v_�Ʒ�����, v_��Ӧ֤����, v_���������������, n_�Ƿ���, v_����Ա����,
         Sysdate, v_����Ա����, Sysdate);
    Else
      Update �η�����
      Set �������� = v_��������, ���� = v_����, ���� = v_����, �������� = v_��������, ��Դ = v_��Դ, ���ժҪ = v_���ժҪ, �������� = v_��������, �������� = v_��������,
          �Ʒ����� = v_�Ʒ�����, ��Ӧ֤���� = v_��Ӧ֤����, ��������������� = v_���������������, �Ƿ��� = n_�Ƿ���, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate
      Where ����id = n_����id;
    End If;
  
    --���·�������
    If n_Usetype = 1 Then
      Delete From �������� Where ����id = n_����id;
    End If;
    For I In 1 .. Jsonlistobj.Count Loop
      Jsonobj    := Pljson(Jsonlistobj.Get(I));
      n_��ҩid   := To_Number(Jsonobj.Get_String('��ҩID'));
      n_����     := To_Number(Jsonobj.Get_String('����'));
      v_�÷���ע := Jsonobj.Get_String('�÷���ע');
      v_�ŷ����� := Jsonobj.Get_String('�ŷ�����');
    
      Insert Into ��������
        (����id, ����id, ��ҩid, �÷���ע, �ŷ�����, ����, ������, ����ʱ��)
      Values
        (��������_����id.Nextval, n_����id, n_��ҩid, v_�÷���ע, v_�ŷ�����, n_����, v_����Ա����, Sysdate);
    End Loop;
  
    Open Output_Out For
      Select n_����id As ����id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Save_Fjitem;

  -----------------------------------------------------
  --�޸���ҽ����
  -----------------------------------------------------
  Procedure Set_Distype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/1-�޸�
    n_����id     ��ҽ����.����id%Type;
    v_��������   ��ҽ����.��������%Type;
    v_����       ��ҽ����.����%Type;
    v_�Ʊ�       ��ҽ����.�Ʊ�%Type;
    v_����Ա���� ��ҽ����.����޸���%Type;
    n_����Աid   Number;
  Begin
  
    --�������
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_����id     := To_Number(Nvl(Jsonobj.Get_String('����ID'), 0));
    v_��������   := Jsonobj.Get_String('��������');
    v_����       := Jsonobj.Get_String('����');
    v_�Ʊ�       := Jsonobj.Get_String('�Ʊ�');
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    n_����Աid   := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    If n_Usetype = 0 Then
      Select ��ҽ����_����id.Nextval Into n_����id From Dual;
      Insert Into ��ҽ����
        (����id, ��������, ����, �Ʊ�, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_����id, v_��������, v_����, v_�Ʊ�, v_����Ա����, Sysdate, v_����Ա����, Sysdate);
    Else
      Update ��ҽ����
      Set �������� = v_��������, ���� = v_����, �Ʊ� = v_�Ʊ�, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate
      Where ����id = n_����id;
    End If;
    Open Output_Out For
      Select n_����id As ����id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Distype;

  -----------------------------------------------------
  --�޸���ҽ֤��
  -----------------------------------------------------
  Procedure Set_Zxtype
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/1-�޸�
    n_֤��id     ��ҽ֤��.֤��id%Type;
    v_֤������   ��ҽ֤��.֤������%Type;
    v_����       ��ҽ֤��.����%Type;
    n_����id     ��ҽ֤��.����id%Type;
    v_֤������   ��ҽ֤��.֤������%Type;
    v_֤���η�   ��ҽ֤��.֤���η�%Type;
    v_֢״����   ��ҽ֤��.֢״����%Type;
    v_����Ա���� ��ҽ֤��.����޸���%Type;
    n_����Աid   Number;
  Begin
  
    --�������
    Jsonobj    := Pljson(Input_In);
    n_Usetype  := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_֤��id   := To_Number(Nvl(Jsonobj.Get_String('֤��ID'), 0));
    v_֤������ := Jsonobj.Get_String('֤������');
    v_����     := Jsonobj.Get_String('����');
    n_����id   := To_Number(Nvl(Jsonobj.Get_String('����ID'), 0));
    v_֤������ := Jsonobj.Get_String('֤������');
    v_֤���η� := Jsonobj.Get_String('֤���η�');
    v_֢״���� := Jsonobj.Get_String('֢״����');
  
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    n_����Աid   := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    If n_Usetype = 0 Then
      Select ��ҽ֤��_֤��id.Nextval Into n_֤��id From Dual;
      Insert Into ��ҽ֤��
        (֤��id, ֤������, ����, ����id, ֤������, ֤���η�, ֢״����, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_֤��id, v_֤������, v_����, n_����id, v_֤������, v_֤���η�, v_֢״����, v_����Ա����, Sysdate, v_����Ա����, Sysdate);
    Else
      Update ��ҽ֤��
      Set ֤������ = v_֤������, ���� = v_����, ֤������ = v_֤������, ֤���η� = v_֤���η�, ֢״���� = v_֢״����, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate
      Where ֤��id = n_֤��id;
    End If;
    Open Output_Out For
      Select n_֤��id As ֤��id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Zxtype;

  -----------------------------------------------------
  --����֤�ͷ�����Ӧ
  -----------------------------------------------------
  Procedure Set_Zxtofj
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/1-����״̬
    n_����id     ֤�ͷ�������.����id%Type;
    n_֤��id     ֤�ͷ�������.֤��id%Type;
    n_����id     ֤�ͷ�������.����id%Type;
    n_״̬       ֤�ͷ�������.״̬%Type;
    v_����Ա���� ��ҽ֤��.����޸���%Type;
    n_����Աid   Number;
  Begin
    --�������
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    n_����Աid   := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    If n_Usetype = 0 Then
      n_֤��id := To_Number(Nvl(Jsonobj.Get_String('֤��ID'), 0));
      n_����id := To_Number(Nvl(Jsonobj.Get_String('����ID'), 0));
      Select ֤�ͷ�������_����id.Nextval Into n_����id From Dual;
      Insert Into ֤�ͷ�������
        (����id, ֤��id, ����id, ״̬, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_����id, n_֤��id, n_����id, 1, v_����Ա����, Sysdate, v_����Ա����, Sysdate);
    Elsif n_Usetype = 1 Then
      n_״̬   := To_Number(Nvl(Jsonobj.Get_String('״̬'), 0));
      n_����id := To_Number(Nvl(Jsonobj.Get_String('����ID'), 0));
      Update ֤�ͷ������� Set ״̬ = n_״̬, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate Where ����id = n_����id;
    End If;
    Open Output_Out For
      Select n_����id As ����id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Zxtofj;

  -----------------------------------------------------
  --�޸���֤��֢
  -----------------------------------------------------
  Procedure Set_Adddis
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/1-�޸�/2-�ı�״̬
    n_��֢id     ��֤��֢.��֢id%Type;
    v_��֢����   ��֤��֢.��֢����%Type;
    v_����       ��֤��֢.����%Type;
    n_״̬       ��֤��֢.״̬%Type;
    v_����Ա���� ��֤��֢.����޸���%Type;
    n_����Աid   Number;
  Begin
  
    --�������
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_��֢id     := To_Number(Nvl(Jsonobj.Get_String('��֢ID'), 0));
    v_��֢����   := Jsonobj.Get_String('��֢����');
    v_����       := Jsonobj.Get_String('����');
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    n_����Աid   := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    If n_Usetype = 0 Then
      Select ��֤��֢_��֢id.Nextval Into n_��֢id From Dual;
      Insert Into ��֤��֢
        (��֢id, ��֢����, ����, ״̬, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_��֢id, v_��֢����, v_����, 1, v_����Ա����, Sysdate, v_����Ա����, Sysdate);
    Elsif n_Usetype = 1 Then
      Update ��֤��֢
      Set ��֢���� = v_��֢����, ���� = v_����, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate
      Where ��֢id = n_��֢id;
    Elsif n_Usetype = 2 Then
      n_״̬ := To_Number(Nvl(Jsonobj.Get_String('״̬'), 0));
      Update ��֤��֢ Set ״̬ = n_״̬, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate Where ��֢id = n_��֢id;
    End If;
    Open Output_Out For
      Select n_��֢id As ��֢id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Adddis;

  -----------------------------------------------------
  --�޸ļ�֢�η�
  -----------------------------------------------------
  Procedure Set_Addzf
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/1-�޸�/2-�ı�״̬
    n_�η�id     ��֢�η�.�η�id%Type;
    v_�η�����   ��֢�η�.�η�����%Type;
    n_��֢id     ��֢�η�.��֢id%Type;
    v_����       ��֢�η�.����%Type;
    n_״̬       ��֢�η�.״̬%Type;
    v_����Ա���� ��֢�η�.����޸���%Type;
    n_����Աid   Number;
  Begin
    --�������
    Jsonobj      := Pljson(Input_In);
    n_Usetype    := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_�η�id     := To_Number(Nvl(Jsonobj.Get_String('�η�ID'), 0));
    n_��֢id     := To_Number(Nvl(Jsonobj.Get_String('��֢ID'), 0));
    v_�η�����   := Jsonobj.Get_String('�η�����');
    v_����       := Jsonobj.Get_String('����');
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    n_����Աid   := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    If n_Usetype = 0 Then
      Select ��֢�η�_�η�id.Nextval Into n_�η�id From Dual;
      Insert Into ��֢�η�
        (�η�id, �η�����, ����, ��֢id, ״̬, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_�η�id, v_�η�����, v_����, n_��֢id, 1, v_����Ա����, Sysdate, v_����Ա����, Sysdate);
    Elsif n_Usetype = 1 Then
      Update ��֢�η�
      Set �η����� = v_�η�����, ���� = v_����, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate
      Where �η�id = n_�η�id;
    Elsif n_Usetype = 2 Then
      n_״̬ := To_Number(Nvl(Jsonobj.Get_String('״̬'), 0));
      Update ��֢�η� Set ״̬ = n_״̬, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate Where �η�id = n_�η�id;
    End If;
    Open Output_Out For
      Select n_�η�id As �η�id From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Addzf;

  -----------------------------------------------------
  --�����η���ҩ��Ӧ
  -----------------------------------------------------
  Procedure Set_Zftozy
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj      Pljson;
    n_Usetype    Number; --0-����/2-�޸�/1-����״̬
    n_��ҩid     ��֢��ҩ.��ҩid%Type;
    n_�η�id     ��֢��ҩ.�η�id%Type;
    n_��ҩid     ��֢��ҩ.��ҩid%Type;
    n_����       ��֢��ҩ.����%Type;
    n_״̬       ��֢��ҩ.״̬%Type;
    v_����Ա���� ��֢��ҩ.����޸���%Type;
    n_����Աid   Number;
  Begin
    --�������
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
  
    v_����Ա���� := Jsonobj.Get_String('����Ա����');
    n_����Աid   := To_Number(Nvl(Jsonobj.Get_String('����ԱID'), 0));
    If n_Usetype = 0 Then
      n_�η�id := To_Number(Nvl(Jsonobj.Get_String('�η�ID'), 0));
      n_��ҩid := To_Number(Nvl(Jsonobj.Get_String('��ҩID'), 0));
      n_����   := To_Number(Nvl(Jsonobj.Get_String('����'), 0));
      Select ��֢��ҩ_��ҩid.Nextval Into n_��ҩid From Dual;
    
      Insert Into ��֢��ҩ
        (��ҩid, �η�id, ��ҩid, ����, ״̬, ������, ����ʱ��, ����޸���, ����޸�ʱ��)
      Values
        (n_��ҩid, n_�η�id, n_��ҩid, n_����, 1, v_����Ա����, Sysdate, v_����Ա����, Sysdate);
    Elsif n_Usetype = 1 Then
      n_��ҩid := To_Number(Nvl(Jsonobj.Get_String('��ҩID'), 0));
      n_�η�id := To_Number(Nvl(Jsonobj.Get_String('�η�ID'), 0));
      n_��ҩid := To_Number(Nvl(Jsonobj.Get_String('��ҩID'), 0));
      n_����   := To_Number(Nvl(Jsonobj.Get_String('����'), 0));
      n_��ҩid := To_Number(Nvl(Jsonobj.Get_String('��ҩID'), 0));
      Update ��֢��ҩ
      Set �η�id = n_�η�id, ��ҩid = n_��ҩid, ���� = n_����, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate
      Where ��ҩid = n_��ҩid;
    Elsif n_Usetype = 2 Then
      n_״̬   := To_Number(Nvl(Jsonobj.Get_String('״̬'), 0));
      n_��ҩid := To_Number(Nvl(Jsonobj.Get_String('��ҩID'), 0));
      Update ��֢��ҩ Set ״̬ = n_״̬, ����޸��� = v_����Ա����, ����޸�ʱ�� = Sysdate Where ��ҩid = n_��ҩid;
    End If;
    Open Output_Out For
      Select n_��ҩid As ��ҩid From Dual;
  Exception
    When Others Then
      Errorcenter(SQLCode, SQLErrM);
  End Set_Zftozy;

  -----------------------------------------------------
  --��Ŀɾ������
  -----------------------------------------------------
  Procedure Del_Zydata
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj   Pljson;
    n_Usetype Number;
    n_Id      Number;
    v_Name    Varchar2(50);
    v_Table   Varchar2(4000);
    v_Err_Msg Varchar2(255);
    Err_Item Exception;
  Begin
    --�������
    Jsonobj   := Pljson(Input_In);
    n_Usetype := To_Number(Nvl(Jsonobj.Get_String('USETYPE'), 0));
    n_Id      := To_Number(Nvl(Jsonobj.Get_String('ID'), 0));
  
    If n_Usetype = 0 Then
      --��ҩĿ¼
      Select Nvl(Max(������), '��') Into v_Name From ��ҩĿ¼ Where ��ҩid = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From ��ҩĿ¼ Where ��ҩid = n_Id;
      End If;
    Elsif n_Usetype = 1 Then
      --��ҽ����
      Select Nvl(Max(������), '��') Into v_Name From ��ҽ���� Where ����id = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From ��ҽ���� Where ����id = n_Id;
      End If;
    Elsif n_Usetype = 2 Then
      --��ҽ֤��
      Select Nvl(Max(������), '��') Into v_Name From ��ҽ֤�� Where ֤��id = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From ��ҽ֤�� Where ֤��id = n_Id;
      End If;
    Elsif n_Usetype = 3 Then
      --֤�ͷ�������
      Select Nvl(Max(������), '��') Into v_Name From ֤�ͷ������� Where ����id = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From ֤�ͷ������� Where ����id = n_Id;
      End If;
    Elsif n_Usetype = 4 Then
      --�η�����
      Select Nvl(Max(������), '��') Into v_Name From �η����� Where ����id = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From �η����� Where ����id = n_Id;
      End If;
    Elsif n_Usetype = 5 Then
      --��֤��֢
      Select Nvl(Max(������), '��') Into v_Name From ��֤��֢ Where ��֢id = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From ��֤��֢ Where ��֢id = n_Id;
      End If;
    Elsif n_Usetype = 6 Then
      --��֢�η�
      Select Nvl(Max(������), '��') Into v_Name From ��֢�η� Where �η�id = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From ��֢�η� Where �η�id = n_Id;
      End If;
    Elsif n_Usetype = 7 Then
      --��֢��ҩ
      Select Nvl(Max(������), '��') Into v_Name From ��֢��ҩ Where ��ҩid = n_Id;
      If v_Name = 'ϵͳ����' Then
        v_Err_Msg := '��ǰɾ����ĿΪϵͳ������Ŀ,����ɾ����';
        Raise Err_Item;
      Else
        Delete From ��֢��ҩ Where ��ҩid = n_Id;
      End If;
    End If;
    Open Output_Out For
      Select '1' As ��� From Dual;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
    When Others Then
      v_Table := SQLErrM;
      If SQLCode = -2292 Then
        Select Table_Name
        Into v_Table
        From All_Constraints
        Where Instr(v_Table, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
        v_Err_Msg := '[ZLSOFT]�ü�¼�� ' || v_Table || ' ���Ѿ�ʹ��,' || Chr(13) || '����ɾ�����޸�[ZLSOFT]';
        Raise_Application_Error(-20005, v_Err_Msg);
      End If;
  End Del_Zydata;

  -----------------------------------------------------
  --��ȡ���ݿ�ϵͳʱ��
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select Sysdate As ��ǰʱ�� From Dual;
  End Get_Now_Time;

  -----------------------------------------------------
  --�쳣����
  -----------------------------------------------------
  Procedure Errorcenter
  (
    Err_Num In Number,
    Err_Msg In Varchar2
  ) Is
    v_Outnum Number := 0;
    v_Outmsg Varchar2(1000) := '';
    v_Count  Number;
    v_Temp   Varchar2(1000) := '';
  
    Cursor Cur_Ind_Cols Is
      Select Table_Name, Column_Name From All_Ind_Columns Where Instr(Err_Msg, Index_Owner || '.' || Index_Name) > 0;
  
    Cursor Cur_Con_Cols Is
      Select Table_Name, Column_Name
      From All_Cons_Columns
      Where (Owner, Constraint_Name) =
            (Select r_Owner, r_Constraint_Name
             From All_Constraints
             Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 0 And Rownum < 2);
  Begin
    If Err_Num = -1 Then
      For Row_Cols In Cur_Ind_Cols Loop
        v_Temp   := Row_Cols.Table_Name;
        v_Outmsg := v_Outmsg || '��' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]' || v_Temp || '��(' || Substr(v_Outmsg, 2) || ')�����ظ���[ZLSOFT]';
      v_Outnum := -20000;
    Elsif Err_Num = -1000 Then
      v_Outmsg := '[ZLSOFT]�򿪵����ݱ�̫�࣬��Ҫʱ��ϵͳ����Ա�޸����ݿ��Open_Cursors���á�';
      v_Outnum := -20001;
    Elsif Err_Num = -1400 Or Err_Num = -1407 Then
      Select Table_Name, Column_Name
      Into v_Temp, v_Outmsg
      From All_Tab_Columns
      Where Instr(Err_Msg, '"' || Owner || '"."' || Table_Name || '"."' || Column_Name || '"') > 0 And Rownum < 2;
      v_Outmsg := '[ZLSOFT]' || v_Temp || '��(' || v_Outmsg || ')�������룡[ZLSOFT]';
      v_Outnum := -20002;
    Elsif Err_Num = -1401 Then
      v_Outmsg := '[ZLSOFT]���ڸ����ֵ�������п����ƣ��������ӻ����ʧ�ܡ�[ZLSOFT]';
      v_Outnum := -20003;
    Elsif Err_Num = -2290 Then
      Select Table_Name, Search_Condition
      Into v_Temp, v_Outmsg
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      If Instr(v_Outmsg, 'IS NOT NULL') > 0 Then
        v_Outmsg := '[ZLSOFT]' || v_Temp || ' �� ' || Replace(v_Outmsg, 'IS NOT NULL', '�������룡') || '[ZLSOFT]';
        v_Outnum := -20004;
      Else
        v_Outmsg := Err_Msg;
        v_Outnum := -20999;
      End If;
    Elsif Err_Num = -2292 Then
      Select Table_Name
      Into v_Temp
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      For Row_Cols In Cur_Con_Cols Loop
        v_Outmsg := v_Outmsg || '��' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]�ü�¼�� ' || v_Temp || ' ���Ѿ�ʹ��,' || Chr(13) || '����ɾ�����޸�(' || Substr(v_Outmsg, 2) || ')[ZLSOFT]';
      v_Outnum := -20005;
    Else
      v_Outmsg := Err_Msg;
      v_Outnum := -20999;
    End If;
  
    ------------------------
    --��󲹳���д�����¼�Ĵ���
    ------------------------
    Raise_Application_Error(v_Outnum, Substr(v_Outmsg, 1, 100));
  End Errorcenter;

End Pkg_Zyedit;
/




----------------------------------------------------------------------------
--[[21.���ҵ��]]
----------------------------------------------------------------------------
--רҵ��RIS�ӿ�
CREATE OR REPLACE Package b_Zlxwinterface Is
  Type t_Refcur Is Ref Cursor;

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  );

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  );

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  );

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  );

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ҽ��ID_In   In Risҽ��ʧ�ܼ�¼.ҽ��id%Type,
    ��������_In In Risҽ��ʧ�ܼ�¼.��������%Type
  );

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  );

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  );

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type);

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  );

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  );

  --16���Ǽ�Σ��ֵ
  Procedure ����Σ��ֵ��¼_Insert
  (
    Id_In         In ����Σ��ֵ��¼.Id%Type,
    ������Դ_In   In ����Σ��ֵ��¼.������Դ%Type,
    ����id_In     In ����Σ��ֵ��¼.����id%Type,
    ��ҳid_In     In ����Σ��ֵ��¼.��ҳid%Type,
    �Һŵ�_In     In ����Σ��ֵ��¼.�Һŵ�%Type,
    Ӥ��_In       In ����Σ��ֵ��¼.Ӥ��%Type,
    ����_In       In ����Σ��ֵ��¼.����%Type,
    �Ա�_In       In ����Σ��ֵ��¼.�Ա�%Type,
    ����_In       In ����Σ��ֵ��¼.����%Type,
    ҽ��id_In     In ����Σ��ֵ��¼.ҽ��id%Type,
    �걾id_In     In ����Σ��ֵ��¼.�걾id%Type,
    Σ��ֵ����_In In ����Σ��ֵ��¼.Σ��ֵ����%Type,
    ����ʱ��_In   In ����Σ��ֵ��¼.����ʱ��%Type,
    �������id_In In ����Σ��ֵ��¼.�������id%Type,
    ������_In     In ����Σ��ֵ��¼.������%Type
  );

  --17��ȡ��Σ��ֵ
  Procedure ����Σ��ֵ��¼_Delete(ҽ��id_In In ����Σ��ֵ��¼.ҽ��id%Type);

  --18�������ٴ�ҽ��
  Function ����ҽ����¼_Send(ҽ��id_In In ����ҽ������.ҽ��id%Type) Return Varchar2;

End b_Zlxwinterface;
/

CREATE OR REPLACE Package Body b_Zlxwinterface Is

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  ) Is
  
    --������ҽ��ID_IN - ����ִ�е�ҽ��ID��
    --      ״̬_IN - -1-ɾ����0-ԤԼ��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�15-����
    --     ����ִ��_In -0-ȫ��ִ�У�1-����ִ�У����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  
    Cursor c_Adviceinfo Is
      Select a.Id, a.���id, Nvl(a.���id, a.Id) As ��id, a.�������, a.������Դ, a.ִ�п���id, b.ִ�й���
      From ����ҽ����¼ A, ����ҽ������ B
      Where a.Id = b.ҽ��id And ID = ҽ��id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_ִ��״̬ ����ҽ������.ִ��״̬%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    n_ִ��     Number; --����Ƿ���Ҫ����״̬��1����Ҫ���£���������Ҫ����
    v_Count    Number;
    v_�����   ����ҽ������.�����%Type;
    v_���ʱ�� ����ҽ������.���ʱ��%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    v_ִ��״̬ := 0;
    v_ִ�й��� := 0;
  
    --��ȡҽ������ҽ��ID������ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --����״̬_INִ��ҽ��
    ---1-ɾ����0-ԤԼ(��RIS��ʵ���Ͼ���ɾ��)��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�13-ȡ����ˣ�14-����ɾ����15-����
  
    If ״̬_In = -1 Or ״̬_In = 0 Then
      v_ִ��״̬ := 0; --δִ��
      v_ִ�й��� := 0;
    Elsif ״̬_In = 1 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 2; --�ѱ���
    Elsif ״̬_In = 3 Or ״̬_In = 14 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 3; --�Ѽ��
    Elsif ״̬_In = 4 Then
      --���ı�
      v_ִ��״̬ := v_ִ��״̬;
    Elsif ״̬_In = 9 Or ״̬_In = 13 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 4; --�ѱ���
    Elsif ״̬_In = 12 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 5; --�����
    Elsif ״̬_In = 15 Then
      v_ִ��״̬ := 1; --��ȫִ��
      v_ִ�й��� := 6; --�����
      v_�����   := ������Ա_In;
      v_���ʱ�� := ִ��ʱ��_In;
    End If;
  
    n_ִ�� := 1; --Ĭ�϶�Ҫ����״̬
  
    If ״̬_In = 13 Or ״̬_In = 14 Then
      --ɾ����Ӧ��������
      Delete From ���Ӳ�����¼
      Where ID = (Select ����id From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In);
      Delete From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In;
    
      --ɾ�����ж��Ƿ񻹴��ڱ��棬��������ҽ��״̬���ֲ��䣬������ȫ��ɾ�������ҽ��״̬
      Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In;
    
      If v_Count > 0 Then
        n_ִ�� := 0; --��������ҽ��״̬���ֲ���
      End If;
    End If;
  
    --�����ɾ������ɾ�����е�ԤԼ��Ϣ
    If ״̬_In = -1 Or ״̬_In = 0 Then
      Zl_Ris���ԤԼ_Delete(ҽ��id_In);
    End If;
  
    --����ǵǼǣ����жϴ˼���Ƿ�δִ��
    If ״̬_In = 1 Then
      If r_Adviceinfo.ִ�й��� >= 3 Then
        v_Error := '�����Ѿ���������ˣ������ظ��Ǽǡ�';
        Raise Err_Custom;
      End If;
    End If;
  
    --��ʼִ��ҽ��
    If n_ִ�� = 1 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        -- ������λҽ������ִ��
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id = ҽ��id_In;
      Else
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = r_Adviceinfo.��id Or ���id = r_Adviceinfo.��id));
      End If;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receiverisstate;

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id, �������, ������Դ From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_Temp     Varchar2(255);
    v_��Ա��� ��Ա��.���%Type;
    v_��Ա���� ��Ա��.����%Type;
    v_����id   ���ű�.Id%Type;
    v_�������� ����ҽ������.��¼����%Type;
    v_���ͺ�   ����ҽ������.���ͺ�%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    Select ���ͺ�, ִ�й��� Into v_���ͺ�, v_ִ�й��� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --�ǼǺ���ɲ�ִ�з���  2-�Ǽǣ�3-��飬4-���棬5-��ˣ�6-���
    If v_ִ�й��� >= 2 Or v_ִ�й��� <= 6 Then
      --ȡ��ǰ������Ա
      If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null And ִ�в���id_In Is Not Null Then
        v_��Ա��� := ����Ա���_In;
        v_��Ա���� := ����Ա����_In;
        v_����id   := ִ�в���id_In;
      Else
        v_Temp     := Zl_Identity;
        v_����id   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      End If;
    
      If r_Advice.������Դ = 2 Then
        Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
        Into v_��������
        From ����ҽ������
        Where ���ͺ� = v_���ͺ� And ҽ��id = ҽ��id_In;
      Else
        v_�������� := 1;
      End If;
    
      --ִ�з��ú��Զ�����
      If v_�������� = 1 Then
        Zl_����ҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      Else
        Zl_סԺҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      End If;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��;

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������
    --      ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
  
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_���ͺ� ����ҽ������.���ͺ�%Type;
    v_Count  Number;
    v_Error  Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ���߳�Ժ�ļ�����룬����ִ�з���
    Select Count(*)
    Into v_Count
    From ����ҽ����¼ A, ������ҳ B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And (b.��Ժ���� Is Not Null Or b.״̬ = 3) And a.Id = r_Advice.��id;
  
    If v_Count > 0 Then
      v_Error := 'סԺ�����Ѿ���Ժ����Ԥ��Ժ������ȡ�����á�';
      Raise Err_Custom;
    End If;
  
    Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --����ͳһ��ҽ��ִ��Cancel����
    Zl_����ҽ��ִ��_Cancel(ҽ��id_In, v_���ͺ�, Null, ����ִ��_In, ִ�в���id_In, ����Ա���_In, ����Ա����_In);
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��_Cancel;

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  ) Is
    --��ȡ����ҽ��������������Ϣ
    Cursor c_Advice
    (
      v_��id  Number,
      v_Risid Number
    ) Is
      Select e.Id, e.������Դ, e.����id, e.��ҳid, e.Ӥ��, e.���˿���id, e.�ļ�id, e.��������, e.��������, f.����id, e.ִ�п���id
      From (Select c.Id, c.������Դ, c.����id, c.��ҳid, c.Ӥ��, c.���˿���id, c.�ļ�id, d.���� ��������, d.���� ��������, c.ִ�п���id
             From (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.Ӥ��, a.���˿���id, b.�����ļ�id �ļ�id, a.ִ�п���id
                    From ����ҽ����¼ A, ��������Ӧ�� B
                    Where a.Id = v_��id And a.������Ŀid = b.������Ŀid(+) And b.Ӧ�ó���(+) = Decode(a.������Դ, 2, 2, 4, 4, 1)) C,
                  �����ļ��б� D
             Where c.�ļ�id = d.Id(+)) E, ����ҽ������ F
      Where e.Id = f.ҽ��id(+) And f.Risid(+) = v_Risid;
  
    --�����ļ������Ԫ��
    Cursor c_File(v_File Number) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where a.�ļ�id = v_File
      Order By a.�������;
  
    Cursor c_Report(v_���Ӳ�����¼id Number) Is
      Select b.Id, a.�����ı�
      From ���Ӳ������� A, ���Ӳ������� B
      Where a.�������� = 3 And a.Id = b.��id And b.�������� = 2 And b.��ֹ�� = 0 And a.�ļ�id = v_���Ӳ�����¼id;
  
    Cursor c_Content
    (
      v_�ļ�id Number,
      v_���id Number
    ) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where �ļ�id = v_�ļ�id And ��id = v_���id;
  
    r_Advice        c_Advice%RowType;
    v_����id        ���Ӳ�������.�ļ�id%Type;
    v_��������id    ���Ӳ�������.Id%Type;
    v_��������idnew ���Ӳ�������.Id%Type;
    v_�������      ���Ӳ�������.�������%Type;
    v_��id          ���Ӳ�������.��id%Type;
    v_�����ı�      ���Ӳ�������.�����ı�%Type;
    v_�������id    ���Ӳ�������.�������id%Type;
    --v_��ʽ����    ���Ӳ�����ʽ.����%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_��ҽ��id ����ҽ������.ҽ��id%Type;
    v_���     Varchar2(300);
    n_����     Number;
    n_Rptcount Number;
    v_�������� ���Ӳ�����¼.��������%Type;
    v_�Һŵ�id ���˹Һż�¼.Id%Type;
  
    Function Getrptno
    (
      v_ҽ��idin   ����ҽ������.ҽ��id%Type,
      v_��������in ���Ӳ�����¼.��������%Type
    ) Return Varchar As
      v_Return Number;
      v_No     Number;
      v_Count  Number;
    Begin
      Select Count(ҽ��id) + 1 Into v_No From ����ҽ������ Where ҽ��id = v_ҽ��idin;
      v_Count := 1;
      While v_Count = 1 Loop
        Select Count(ID)
        Into v_Count
        From ����ҽ������ A, ���Ӳ�����¼ B
        Where a.ҽ��id = v_ҽ��idin And a.����id = b.Id And b.�������� = v_��������in || v_No;
        If v_Count = 1 Then
          v_No := v_No + 1;
        End If;
      End Loop;
      v_Return := v_No;
      Return v_Return;
    End Getrptno;
  
  Begin
  
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��id From ����ҽ����¼ Where ID = ҽ��id_In;
  
    Open c_Advice(v_��ҽ��id, Nvl(Risid_In, 0));
    Fetch c_Advice
      Into r_Advice;
  
    If Nvl(r_Advice.�ļ�id, 0) = 0 Then
      v_Error := '���μ����Ŀû�ж�Ӧ��صļ�鱨�棬�������Ա��ϵ��';
      Raise Err_Custom;
    Else
      If Nvl(r_Advice.����id, 0) > 0 Then
        ----����������
        --�ҳ��������д�ı�������к���"%����%","%����%","%����%","%���%",���ô���Ĳ�������
        For r_Report In c_Report(r_Advice.����id) Loop
          If r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ��������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%���%' Then
            Update ���Ӳ������� Set �����ı� = �������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ���潨��_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          End If;
        End Loop;
        --���±���ʱ��
        Update ���Ӳ�����¼
        Set ���ʱ�� = Sysdate, ������ = ����ҽ��_In, ����ʱ�� = Sysdate
        Where ID = r_Advice.����id;
      Else
        --���жϵ������Ƿ��ж�Ӧ����ٺͱ��
        If Nvl(��������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ�����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(�������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%���%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(���潨��_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ������顿��Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
      
        If r_Advice.������Դ = 1 Then
          --�����ȡ�Һŵ�ID
          Select nvl(Max(c.Id), 0)
          Into v_�Һŵ�id
          From ����ҽ����¼ B, ���˹Һż�¼ C
          Where b.�Һŵ� = c.No(+) And c.��¼״̬ In (1, 3) And b.Id = v_��ҽ��id;
        Else
          --����������޹Һŵ�ID��ֱ������Ϊ0
          v_�Һŵ�id := 0;
        End If;
      
        --�������Ӳ�����¼
        Select ���Ӳ�����¼_Id.Nextval Into v_����id From Dual;
        n_Rptcount := Getrptno(ҽ��id_In, r_Advice.��������);
        If n_Rptcount > 1 Then
          v_�������� := r_Advice.�������� || n_Rptcount;
        Else
          v_�������� := r_Advice.��������;
        End If;
        Insert Into ���Ӳ�����¼
          (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ���ʱ��, ������, ����ʱ��, ���汾, ǩ������)
        Values
          (v_����id, r_Advice.������Դ, r_Advice.����id, Decode(r_Advice.������Դ, 2, r_Advice.��ҳid, v_�Һŵ�id), r_Advice.Ӥ��,
           r_Advice.���˿���id, r_Advice.��������, r_Advice.�ļ�id, v_��������, ����ҽ��_In, Sysdate, Sysdate, ����ҽ��_In, Sysdate, 1, 2);
      
        --����ҽ�������¼
        Insert Into ����ҽ������ (ҽ��id, ����id, Risid) Values (v_��ҽ��id, v_����id, Risid_In);
      
        v_������� := 0;
      
        --�²�����������
        For r_File In c_File(r_Advice.�ļ�id) Loop
          Select ���Ӳ�������_Id.Nextval Into v_��������id From Dual;
          v_�����ı�   := r_File.�����ı�;
          v_�������id := 0;
        
          If Nvl(r_File.��������, 0) = 1 And Nvl(r_File.��id, 0) = 0 Then
            --���
            v_�������id := r_File.Id;
            v_��id       := v_��������id;
          End If;
        
          If Nvl(r_File.��������, 0) = 4 And r_File.Ҫ������ Is Not Null Then
            --Ԫ��
            v_�����ı� := Zl_Replace_Element_Value(r_File.Ҫ������, r_Advice.����id, r_Advice.��ҳid, r_Advice.������Դ, r_Advice.Id);
          End If;
        
          If Nvl(r_File.��id, 0) <> 0 Then
            v_�������id := 0;
          End If;
        
          v_������� := v_������� + 1;
        
          If Instr(v_���, '|' || r_File.��id || '|') > 0 Then
            Null;
          Else
            Insert Into ���Ӳ�������
              (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��,
               Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
            Values
              (v_��������id, v_����id, 1, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������, r_File.������, r_File.��������,
               r_File.��������, Null, v_�����ı�, r_File.�Ƿ���, r_File.Ԥ�����id, r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id,
               r_File.�滻��, r_File.Ҫ������, r_File.Ҫ������, r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ, r_File.Ҫ�ر�ʾ, r_File.������̬,
               r_File.Ҫ��ֵ��, Decode(v_�������id, 0, Null, v_�������id));
          End If;
        
          --Ϊ���ʱ�������ı�����
          If Nvl(r_File.��������, 0) = 3 And Nvl(r_File.��id, 0) <> 0 Then
            v_��� := v_��� || ',|' || r_File.Id || '|';
          
            If r_File.�����ı� Like '%����%' Then
              v_�����ı� := ��������_In || Chr(13) || Chr(13);
            Elsif r_File.�����ı� Like '%���%' Then
              v_�����ı� := �������_In || Chr(13) || Chr(13);
            Else
              v_�����ı� := ���潨��_In || Chr(13) || Chr(13);
            End If;
          
            For r_Con In c_Content(r_Advice.�ļ�id, r_File.Id) Loop
              Select ���Ӳ�������_Id.Nextval Into v_��������idnew From Dual;
              v_������� := v_������� + 1;
            
              Insert Into ���Ӳ�������
                (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id,
                 �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
              Values
                (v_��������idnew, v_����id, 1, 0, v_��������id, v_�������, 2, r_Con.������, r_Con.��������, r_Con.��������, Null, v_�����ı�,
                 r_Con.�Ƿ���, r_Con.Ԥ�����id, r_Con.�������, r_Con.ʹ��ʱ��, r_Con.����Ҫ��id, r_Con.�滻��, r_Con.Ҫ������, r_Con.Ҫ������,
                 r_Con.Ҫ�س���, r_Con.Ҫ��С��, r_Con.Ҫ�ص�λ, r_Con.Ҫ�ر�ʾ, r_Con.������̬, r_Con.Ҫ��ֵ��,
                 Decode(v_�������id, 0, Null, v_�������id));
            End Loop;
          End If;
        End Loop;
      
        --����Ӳ�����ʽ�к����������ָ�ʽ�����ַ�������֮���������ֽ����ɼ�
        --Select ���� Into v_��ʽ���� From �����ļ���ʽ Where �ļ�ID=r_Advice.�ļ�ID;
        --Insert Into ���Ӳ�����ʽ (�ļ�ID,����) Values (v_����id,v_��ʽ����);
      
      End If;
    End If;
    Close c_Advice;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receivereport;

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  ) As
  
    v_����        Varchar2(20);
    v_���䵥λ    Varchar2(20);
    v_��������    Date;
    v_������Դ    ����ҽ����¼.������Դ%Type;
    v_����id      ����ҽ����¼.����id%Type;
    v_Strtmpbefor Varchar2(4000);
    v_Msg         Varchar2(4000);
  Begin
    Begin
      Select ������Դ, ����id Into v_������Դ, v_����id From ����ҽ����¼ Where ID = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    If ��������_In Is Null And ����_In Is Not Null Then
      --�����������������
      v_���䵥λ := Substr(����_In, Length(����_In), 1);
      If Instr('��,��,��', v_���䵥λ) <= 0 Then
        v_���䵥λ := Null;
      Else
        v_���� := Replace(����_In, v_���䵥λ, '');
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
        Into v_��������
        From Dual;
      End If;
    Else
      v_�������� := ��������_In;
    End If;
    Select Zl_Fun_Checkidentify(0, v_����id, v_Strtmpbefor) Into v_Strtmpbefor From Dual;
    If v_������Դ = 3 Then
      Update ������Ϣ
      Set ���� = ����_In, �Ա� = Nvl(�Ա�_In, �Ա�), ���� = ����_In, �������� = v_��������, �ѱ� = Nvl(�ѱ�_In, �ѱ�),
          ҽ�Ƹ��ʽ = Nvl(ҽ�Ƹ��ʽ_In, ҽ�Ƹ��ʽ), ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ),
          ���֤�� = ���֤��_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
      --�޸Ķ�Ӧ��ҽ����¼
      Update ����ҽ����¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    Else
      Update ������Ϣ
      Set ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ), ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In,
          ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
    End If;
    Select Zl_Fun_Checkidentify(1, v_����id, v_Strtmpbefor) Into v_Msg From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ������Ϣ_�޸�;

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  ) As
    --������ҽ��ID_IN=����ִ�е�ҽ��ID
  
    v_���ͺ� ����ҽ��ִ��.���ͺ�%Type;
  
  Begin
  
    Begin
      Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    Zl_����ҽ��ִ��_�ܾ�ִ��(ҽ��id_In, v_���ͺ�, ����Ա���_In, ����Ա����_In, ִ�в���id_In, �ܾ�ԭ��_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ȡ��������뵥;

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ҽ��ID_In   In Risҽ��ʧ�ܼ�¼.ҽ��id%Type,
    ��������_In In Risҽ��ʧ�ܼ�¼.��������%Type
  ) Is
  Begin
    Insert Into Risҽ��ʧ�ܼ�¼
      (ID, ҽ��ID, ��������, ����ʱ��, �ط�����)
    Values
      (Risҽ��ʧ�ܼ�¼_Id.Nextval, ҽ��ID_In, ��������_In, Sysdate, 0);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_Insert;

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  ) Is
    v_�ط����� Risҽ��ʧ�ܼ�¼.�ط�����%Type;
  Begin
    --��������_In -- 1 �ط��ɹ���ɾ����¼��2--�ط�ʧ��
  
    If ��������_In = 1 Then
      Delete From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
    Else
      Select �ط����� Into v_�ط����� From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
      If v_�ط����� >= 99 Then
        v_�ط����� := 99;
      Else
        v_�ط����� := v_�ط����� + 1;
      End If;
      Update Risҽ��ʧ�ܼ�¼ Set ����ʱ�� = Sysdate, �ط����� = v_�ط����� Where ID = Id_In;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_�ط�;

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  ) Is
    -- Action_In: 1 �ؽ����ݣ�2 ȡ���ؽ�����
    v_No ����ҽ������.No%Type;
  Begin
    If Action_In = 1 Then
      Select Nextno(14) Into v_No From Dual;
    
      Update ����ҽ������
      Set NO = v_No, �Ʒ�״̬ = 0
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
      Update סԺ���ü�¼ Set ҽ����� = Null Where NO = No_In;
    Elsif Action_In = 2 Then
      Update סԺ���ü�¼ Set ҽ����� = ҽ��id_In Where NO = No_In;
      Update ����ҽ������
      Set NO = No_In, �Ʒ�״̬ = 4
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ����ҽ��_�ؽ�����;

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type) Is
    v_Temp     Varchar2(255);
    v_��Ա���� ��Ա��.����%Type;
  Begin
    --ȡ��ǰ������Ա
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  
    Update Ris���ԤԼ Set �Ƿ��ӡ = 1, ��ӡ�� = v_��Ա����, ��ӡʱ�� = Sysdate Where ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ԤԼ_��ӡ;

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  ) Is
  
    l_����id   t_Numlist := t_Numlist();
    v_����ris  Ris���ÿ���.�Ƿ�����ris%Type;
    v_����ԤԼ Ris���ÿ���.�Ƿ�����ԤԼ%Type;
  
    Cursor c_Dept(Dept_In Varchar2) Is
      Select Column_Value From Table(f_Num2list(Dept_In));
  Begin
  
    If ��������_In = 1 Then
      v_����ris  := 1;
      v_����ԤԼ := Null;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ris = 1;
    Else
      v_����ris  := Null;
      v_����ԤԼ := 1;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ԤԼ = 1;
    End If;
  
    If ����ids_In Is Null Then
      Insert Into Ris���ÿ���
        (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
      Values
        (Ris���ÿ���_Id.Nextval, �������_In, ����_In, Null, v_����ris, v_����ԤԼ);
    Else
      Open c_Dept(����ids_In);
      Fetch c_Dept Bulk Collect
        Into l_����id;
      Close c_Dept;
    
      Forall I In 1 .. l_����id.Count
        Insert Into Ris���ÿ���
          (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
        Values
          (Ris���ÿ���_Id.Nextval, �������_In, ����_In, l_����id(I), v_����ris, v_����ԤԼ);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Update;

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete Is
  
  Begin
    Delete From Ris���ÿ���;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2 Is
    v_Return Varchar2(4000) := Null;
    Cursor c_Patient Is
      Select ����, �Ա�, Decode(�Ա�, '��', 'M', 'Ů', 'F', 'O') As �Ա����, ��������, ����id, ��ϵ�˵�ַ, ��ͥ�绰, ��ϵ�˵绰, ����״��, ���֤��, ��ǰ����id,
             ��ǰ����id, ��ǰ���� As ����, ���￨��, ��Ժʱ��, ��Ժʱ��
      From ������Ϣ
      Where ����id = ����id_In;
    r_Patient c_Patient%RowType;
  
    Cursor c_Order Is
      Select ��ҳid, Ӥ��, Decode(������Դ, 1, 'OUTPAT', 2, 'INPAT', 'UNK') As ������Դ, ����ҽ��, ����ʱ��, У�Ի�ʿ, ҽ������, ������־, ִ�п���id
      From ����ҽ����¼
      Where ID = ҽ��id_In;
    r_Order c_Order%RowType;
  
    Cursor c_Diagnose Is
      Select ������� || Decode(Nvl(�Ƿ�����, 0), 0, '', ' (��)') As �ٴ����
      From �������ҽ�� A, ������ϼ�¼ B
      Where a.ҽ��id = ҽ��id_In And a.���id = b.Id;
    r_Diagnose c_Diagnose%RowType;
  
    --��ȡָ�����������
    Procedure p_Get_Rowtype(Table_In In Varchar2) Is
    Begin
      If Table_In = '������Ϣ' Then
        Open c_Patient;
        Fetch c_Patient
          Into r_Patient;
      Elsif Table_In = '����ҽ����¼' Then
        Open c_Order;
        Fetch c_Order
          Into r_Order;
      Elsif Table_In = '������ϼ�¼' Then
        Open c_Diagnose;
        Fetch c_Diagnose
          Into r_Diagnose;
      End If;
    Exception
      When Others Then
        Null;
    End p_Get_Rowtype;
  
  Begin
    Case
    --ֱ�ӷ��ص�����Ԫ��
      When Ԫ����_In = 'ҽ��ID' Then
        v_Return := ҽ��id_In;
      When Ԫ����_In = '����ID' Then
        v_Return := ����id_In;
      
    --�������Ա𵥶�����������Ӥ��
      When Instr(',����,�Ա�,�Ա����,��������,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        p_Get_Rowtype('������Ϣ');
        If Nvl(r_Order.Ӥ��, 0) = 0 Then
          If Ԫ����_In = '����' Then
            v_Return := r_Patient.����;
          Elsif Ԫ����_In = '�Ա�' Then
            v_Return := r_Patient.�Ա�;
          Elsif Ԫ����_In = '�Ա����' Then
            v_Return := r_Patient.�Ա����;
          Elsif Ԫ����_In = '��������' Then
            v_Return := To_Char(r_Patient.��������, 'YYYYMMDDMISS');
          End If;
        Else
          If Ԫ����_In = '����' Then
            Select Decode(Ӥ������, Null, r_Patient.���� || '֮Ӥ' || Trim(To_Char(���, '9')), Ӥ������) As Ӥ������
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
          Elsif Instr('�Ա�', Ԫ����_In) > 0 Then
            Select Ӥ���Ա�
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            If Ԫ����_In = '�Ա����' Then
              Select Decode(v_Return, '��', 'M', 'Ů', 'F', 'O') Into v_Return From Dual;
            End If;
          Elsif Ԫ����_In = '��������' Then
            Select ����ʱ��
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            v_Return := To_Char(v_Return, 'YYYYMMDDMISS');
          End If;
        End If;
      
    --��ѯ������Ϣ���ص�Ԫ��
      When Instr(',��ϵ�˵�ַ,��ͥ�绰,��ϵ�˵绰,����״��,���֤��,����,���￨��,��Ժʱ��,��Ժʱ��,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('������Ϣ');
        Case Ԫ����_In
          When '��ϵ�˵�ַ' Then
            v_Return := r_Patient.��ϵ�˵�ַ;
          When '��ͥ�绰' Then
            v_Return := r_Patient.��ͥ�绰;
          When '��ϵ�˵绰' Then
            v_Return := r_Patient.��ϵ�˵绰;
          When '����״��' Then
            v_Return := r_Patient.����״��;
          When '���֤��' Then
            v_Return := r_Patient.���֤��;
          When '����' Then
            v_Return := r_Patient.����;
          When '���￨��' Then
            v_Return := r_Patient.���￨��;
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          Else
            v_Return := '';
        End Case;
        --��ѯҽ�����ص�Ԫ��
      When Instr(',������Դ,����ҽ��,����ʱ��,У�Ի�ʿ,ҽ������,������־,������־����,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        Case Ԫ����_In
          When '������Դ' Then
            v_Return := r_Order.������Դ;
          When '����ҽ��' Then
            v_Return := r_Order.����ҽ��;
          When '����ʱ��' Then
            v_Return := To_Char(r_Order.����ʱ��, 'YYYYMMDDMISS');
          When 'У�Ի�ʿ' Then
            v_Return := r_Order.У�Ի�ʿ;
          When 'ҽ������' Then
            v_Return := r_Order.ҽ������;
          When '������־' Then
            v_Return := r_Order.������־;
        End Case;
        --��ѯ��ϼ�¼���ص�Ԫ��
      When Ԫ����_In = '�ٴ����' Then
        p_Get_Rowtype('������ϼ�¼');
        v_Return := r_Diagnose.�ٴ����;
      
      Else
        --���в�ѯSQL����ֵ��Ԫ��
        If Ԫ����_In = 'ִ��վ��' Then
          p_Get_Rowtype('����ҽ����¼');
          Select Decode(վ��, 1, 'SITE0002', 2, 'SITE0001', 3, 'SITE0003', 'SITE0001')
          Into v_Return
          From ���ű�
          Where ID = r_Order.ִ�п���id;
        End If;
        If Ԫ����_In = '��ǰ��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��ʶ��' Then
          Select Decode(a.������Դ, 1, c.�����, 2, Decode(c.סԺ��, Null, c.�����, c.סԺ��), 4, c.������, c.�����)
          Into v_Return
          From ����ҽ����¼ A, ������Ϣ C
          Where a.����id = c.����id And a.Id = ҽ��id_In;
        End If;
    End Case;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      Return Null;
  End Ris_Replace_Element_Value;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete Is
  Begin
    Delete From Ris��Ժ����;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  ) Is
  
  Begin
  
    Insert Into Ris��Ժ����
      (ID, ҽԺ����, ҽԺ����, �û���, ����, ���ݿ������)
    Values
      (Id_In, ҽԺ����_In, ҽԺ����_In, �û���_In, ����_In, ���ݿ������_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Update;

  --16���Ǽ�Σ��ֵ
  Procedure ����Σ��ֵ��¼_Insert
  (
    Id_In         In ����Σ��ֵ��¼.Id%Type,
    ������Դ_In   In ����Σ��ֵ��¼.������Դ%Type,
    ����id_In     In ����Σ��ֵ��¼.����id%Type,
    ��ҳid_In     In ����Σ��ֵ��¼.��ҳid%Type,
    �Һŵ�_In     In ����Σ��ֵ��¼.�Һŵ�%Type,
    Ӥ��_In       In ����Σ��ֵ��¼.Ӥ��%Type,
    ����_In       In ����Σ��ֵ��¼.����%Type,
    �Ա�_In       In ����Σ��ֵ��¼.�Ա�%Type,
    ����_In       In ����Σ��ֵ��¼.����%Type,
    ҽ��id_In     In ����Σ��ֵ��¼.ҽ��id%Type,
    �걾id_In     In ����Σ��ֵ��¼.�걾id%Type,
    Σ��ֵ����_In In ����Σ��ֵ��¼.Σ��ֵ����%Type,
    ����ʱ��_In   In ����Σ��ֵ��¼.����ʱ��%Type,
    �������id_In In ����Σ��ֵ��¼.�������id%Type,
    ������_In     In ����Σ��ֵ��¼.������%Type
  ) Is
  Begin
  
    Zl_����Σ��ֵ��¼_Insert(Id_In, ������Դ_In, ����id_In, ��ҳid_In, �Һŵ�_In, Ӥ��_In, ����_In, �Ա�_In, ����_In, ҽ��id_In, �걾id_In, Σ��ֵ����_In,
                      ����ʱ��_In, �������id_In, ������_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ����Σ��ֵ��¼_Insert;

  --17��ȡ��Σ��ֵ
  Procedure ����Σ��ֵ��¼_Delete(ҽ��id_In In ����Σ��ֵ��¼.ҽ��id%Type) Is
    Cursor c_Critical Is
      Select a.id From ����Σ��ֵ��¼ A Where a.ҽ��id = ҽ��id_In;
  Begin
    For r_Critical In c_Critical Loop
      zl_����Σ��ֵ��¼_delete(r_Critical.id);
    End Loop;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ����Σ��ֵ��¼_Delete;

  --18�������ٴ�ҽ��
  Function ����ҽ����¼_Send(ҽ��id_In In ����ҽ������.ҽ��id%Type) Return Varchar2 Is
    --����ֵ�ּ��������
    --1����ѯ����ҽ�������ؿգ�
    --2����ѯ��ҽ������֯ҽ����Ϣ��
    --3����ѯ��ҽ������֯ҽ����Ϣʧ�ܣ����ؿգ�
    Cursor c_Order Is
      Select a.id, a.���id, a.����id, a.��ҳID, a.������Դ, a.�Һŵ� As NO, a.����, a.�Ա�, a.����, e.������, a.���˿���id,
             To_Char(e.��������, 'YYYY-MM-DD') As ��������, e.��ͥ��ַ As ��ַ, e.��ͥ�绰 As ��ϵ�绰, a.����ҽ�� As ����ҽ��, e.�����, e.סԺ��,
             To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ��������, c.Ӱ����� As �������, a.������Ŀid As ��Ŀ����, a.ҽ������ As ��Ŀ����, a.������־,
             a.������Ŀid || ';' || a.�걾��λ || ';' || a.��鷽�� As �����Ŀ����, a.��������ID As �������ID, a.ִ�п���ID, b.No As ���ݺ�, e.�ѱ�,
             e.ҽ�Ƹ��ʽ, nvl(a.Ӥ��, 0) As Ӥ��, e.���￨��, a.ҽ������, e.����, e.���֤��, b.���ͺ�, Null As ���Ŀ��, 0 As ״̬
      From ����ҽ����¼ A, ����ҽ������ B, Ӱ������Ŀ C, ������Ϣ E
      Where a.id = b.ҽ��id And a.������Ŀid = c.������Ŀid And a.����id = e.����id And (a.ID = ҽ��id_In Or a.���ID = ҽ��id_In) And
            a.ҽ��״̬ = 8;
    r_Order c_Order%RowType;
  
    v_Return       Varchar2(4000) := Null;
    v_�ٴ����     Varchar2(2000) := Null;
    v_ҽ������     Varchar2(800) := Null;
    v_Val          Varchar2(4000) := Null;
    v_Col          Varchar2(1000) := Null;
    v_���˿���     ���ű�.����%Type := Null;
    v_ִ�п���     ���ű�.����%Type := Null;
    v_�������     ���ű�.����%Type := Null;
    v_����         ���ű�.����%Type;
    v_����ID       ������ҳ.��ǰ����id%Type;
    v_��������     ������ҳ.��������%Type;
    v_����         ������ҳ.��Ժ����%Type;
    v_����         ���˹Һż�¼.����%Type;
    v_�Һ�ID       ���˹Һż�¼.id%Type;
    n_Baby         ����ҽ����¼.Ӥ��%Type;
    v_Ӥ������     ������Ϣ.����%Type;
    v_Ӥ���Ա�     ������Ϣ.�Ա�%Type;
    v_Ӥ������     ������Ϣ.����%Type;
    v_Ӥ���������� ������Ϣ.��������%Type;
    v_�Ƿ���     ���˹Һż�¼.����%Type;
    v_����         ������Ϣ.����%Type;
    v_ҽ������     Number; --1 ��ҽ����2 ��λҽ����
  
    --��ȡ�ٴ����
    Function f_GetDiagnose
    (
      v_������Դ In ����ҽ����¼.������Դ%Type,
      v_ҽ��id   In ����ҽ����¼.id%Type
    ) Return Varchar2 Is
    
      --סԺ�ٴ���ϣ�ֻ��ȡ��Ҫ���
      Cursor c_DiagnoseIn Is
        Select a.id, e.�������
        From ����ҽ����¼ A, ������ϼ�¼ E
        Where a.����ID = e.����id And a.��ҳid = e.��ҳid And e.��¼��Դ = 3 And e.������� In (2, 12) And e.��ϴ��� = 1 And e.������� = 1 And
              a.id = v_ҽ��id;
      r_DiagnoseIn c_DiagnoseIn%RowType;
    
      --����������ٴ���ϣ���ȡҽ����Ӧ�����
      Cursor c_DiagnoseOut Is
        Select a.id, e.�������
        From ����ҽ����¼ A, �������ҽ�� D, ������ϼ�¼ E
        Where d.ҽ��id = a.id And d.���id = e.id And a.id = v_ҽ��id;
      r_DiagnoseOut c_DiagnoseOut%RowType;
    
      v_Return Varchar2(2000);
      iCount   Number;
    Begin
      iCount := 0;
      If v_������Դ = 2 Then
        Open c_DiagnoseIn;
        Fetch c_DiagnoseIn
          Into r_DiagnoseIn;
        While c_DiagnoseIn%Found Loop
          iCount := iCount + 1;
          If iCount = 1 Then
            If lengthb(iCount || '��' || r_DiagnoseIn.������� || '��') < 2000 Then
              v_Return := iCount || '��' || r_DiagnoseIn.������� || '��';
            End If;
          Else
            If lengthb(v_Return || Chr(10) || iCount || '��' || r_DiagnoseIn.������� || '��') < 2000 Then
              v_Return := v_Return || Chr(10) || iCount || '��' || r_DiagnoseIn.������� || '��';
            End If;
          End If;
        
          Fetch c_DiagnoseIn
            Into r_DiagnoseIn;
        End Loop;
      
      Else
        Open c_DiagnoseOut;
        Fetch c_DiagnoseOut
          Into r_DiagnoseOut;
        While c_DiagnoseOut%Found Loop
          iCount := iCount + 1;
          If iCount = 1 Then
            If lengthb(iCount || '��' || r_DiagnoseOut.������� || '��') < 2000 Then
              v_Return := iCount || '��' || r_DiagnoseOut.������� || '��';
            End If;
          Else
            If lengthb(v_Return || Chr(10) || iCount || '��' || r_DiagnoseOut.������� || '��') < 2000 Then
              v_Return := v_Return || Chr(10) || iCount || '��' || r_DiagnoseOut.������� || '��';
            End If;
          End If;
        
          Fetch c_DiagnoseOut
            Into r_DiagnoseOut;
        End Loop;
      End If;
    
      If iCount = 1 Then
        v_Return := substr(v_return, 3);
      End If;
      Return v_Return;
    
    End f_GetDiagnose;
  
    --��ȡҽ������
    Function f_GetAttachment(v_ҽ��id In ����ҽ����¼.id%Type) Return Varchar2 Is
      Cursor c_Attachment Is
        Select a.��Ŀ, a.���� From ����ҽ������ A Where a.ҽ��ID = v_ҽ��id Order By ����;
      r_Attachment c_Attachment%RowType;
    
      v_Return Varchar2(800);
    Begin
      Open c_Attachment;
      Fetch c_Attachment
        Into r_Attachment;
      While c_Attachment%Found Loop
        If r_Attachment.���� Is Not Null Then
          If v_Return Is Null Then
            If lengthb('��' || nvl(r_Attachment.��Ŀ, '') || '��' || chr(10) || nvl(r_Attachment.����, '')) < 800 Then
              v_Return := '��' || nvl(r_Attachment.��Ŀ, '') || '��' || chr(10) || nvl(r_Attachment.����, '');
            End If;
          Else
            If lengthb(v_Return || Chr(10) || '��' || nvl(r_Attachment.��Ŀ, '') || '��' || chr(10) ||
                       nvl(r_Attachment.����, '')) < 800 Then
              v_Return := v_Return || Chr(10) || '��' || nvl(r_Attachment.��Ŀ, '') || '��' || chr(10) ||
                          nvl(r_Attachment.����, '');
            End If;
          End If;
        End If;
        Fetch c_Attachment
          Into r_Attachment;
      End Loop;
      Return v_Return;
    End f_GetAttachment;
  
    --��ȡ��������
    Function f_GetDeptName(v_����id In ���ű�.id%Type) Return Varchar2 Is
      v_Return ���ű�.����%Type;
    Begin
      Select Max(����) Into v_Return From ���ű� Where ID = v_����id;
      Return v_Return;
    End f_GetDeptName;
  
  Begin
  
    Open c_Order;
    Fetch c_Order
      Into r_Order;
    While c_Order%Found Loop
      --���ݲ�����Դ����ѯ���ߵ� ������ҳ�����˹Һż�¼,�ٴ���ϵ���Ϣ��ֻ��һ��
      If v_���˿��� Is Null Then
        v_����     := 0;
        v_�Һ�ID   := '';
        v_�������� := '';
        v_����ID   := '';
        v_����     := '';
      
        --ֻ��סԺ���������ȡ ������ҳ���Һż�¼���ٴ����
        If r_Order.������Դ = 2 Then
          Select b.��������, b.��ǰ����id As ����id, b.��Ժ���� As ����
          Into v_��������, v_����ID, v_����
          From ����ҽ����¼ A, ������ҳ B
          Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.id = ҽ��id_In;
        
          v_�ٴ���� := f_GetDiagnose(r_Order.������Դ, ҽ��id_In);
        Elsif r_Order.������Դ = 1 Then
          Select b.����, b.id As �Һ�id
          Into v_����, v_�Һ�ID
          From ����ҽ����¼ A, ���˹Һż�¼ B
          Where a.�Һŵ� = b.no And a.id = ҽ��id_In;
        
          v_�ٴ���� := f_GetDiagnose(r_Order.������Դ, ҽ��id_In);
        End If;
      
        Select decode(nvl(r_Order.������־, 0), 1, 1, nvl(v_����, 0)) Into v_�Ƿ��� From dual;
        v_ҽ������ := f_GetAttachment(ҽ��id_In);
      
        v_����     := f_GetDeptName(v_����ID);
        v_ִ�п��� := f_GetDeptName(r_Order.ִ�п���id);
        v_������� := f_GetDeptName(r_Order.�������ID);
        v_���˿��� := f_GetDeptName(r_Order.���˿���id);
      End If;
    
      --������� ���޸ĳ���ȡID��Ȼ��ͨ��f_GetDeptName��ȡ����
    
      --ѭ��ҽ����¼���������
      If r_Order.���id Is Null Then
        --������ҽ������Ҫ����Ӥ��ҽ��
        v_ҽ������ := 1;
        If r_Order.Ӥ�� = 0 Then
          v_Val  := r_Order.ID || '[;]' || r_Order.����ID || '[;]' || r_Order.������Դ || '[;]' || r_Order.��ҳID || '[;]' ||
                    r_Order.NO || '[;]' || r_Order.���� || '[;]' || r_Order.�Ա� || '[;]' || r_Order.�������� || '[;]' ||
                    r_Order.��ַ || '[;]' || r_Order.��ϵ�绰 || '[;]' || v_������� || '[;]' || r_Order.����ҽ�� || '[;]' ||
                    r_Order.����� || '[;]' || r_Order.סԺ�� || '[;]' || v_���� || '[;]' || v_����;
          v_���� := r_Order.����;
        Else
          n_Baby := r_Order.Ӥ��;
          Select Decode(a.Ӥ������, Null, b.���� || '֮��' || Trim(To_Char(a.���, '9')), a.Ӥ������) As Ӥ������, Ӥ���Ա�,
                 round(Sysdate - a.����ʱ��) || '��' As Ӥ������, To_Char(a.����ʱ��, 'YYYY-MM-DD') As ����ʱ��
          Into v_Ӥ������, v_Ӥ���Ա�, v_Ӥ������, v_Ӥ����������
          From ������������¼ A, ������Ϣ B
          Where a.����id = r_Order.����ID And a.��ҳid = r_Order.��ҳID And a.����id = b.����id And a.��� = n_Baby;
        
          v_���� := v_Ӥ������;
          v_Val  := r_Order.ID || '[;]' || r_Order.����ID || '[;]' || r_Order.������Դ || '[;]' || r_Order.��ҳID || '[;]' ||
                    r_Order.NO || '[;]' || v_Ӥ������ || '[;]' || v_Ӥ���Ա� || '[;]' || v_Ӥ���������� || '[;]' || r_Order.��ַ ||
                    '[;]' || r_Order.��ϵ�绰 || '[;]' || v_������� || '[;]' || r_Order.����ҽ�� || '[;]' || r_Order.����� || '[;]' ||
                    r_Order.סԺ�� || '[;]' || v_���� || '[;]' || v_����;
        End If;
        --v_����ID Ϊ�գ��Ƿ�����                
        v_Val := v_Val || '[;]' || Trim(v_�ٴ����) || '[;]' || r_Order.�������� || '[;]' || r_Order.���Ŀ�� || '[;]' ||
                 r_Order.������� || '[;]' || r_Order.��Ŀ���� || '[;]' || r_Order.��Ŀ���� || '[;]' || r_Order.״̬ || '[;]' ||
                 v_�Ƿ��� || '[;]' || v_���˿��� || '[;]' || v_���� || '[;]' || r_Order.������ || '[;]' || r_Order.���ͺ� || '[;]' ||
                 nvl(r_Order.�������ID, '') || '[;]' || r_Order.���ݺ� || '[;]' || v_�Һ�ID || '[;]' || Trim(v_ҽ������) || '[;]' ||
                 v_ִ�п��� || '[;]' || r_Order.�ѱ� || '[;]' || r_Order.ҽ�Ƹ��ʽ || '[;]' || r_Order.ִ�п���id || '[;]' ||
                 nvl(r_Order.���￨��, 0) || '[;]' || r_Order.ҽ������ || '[;]' || r_Order.���� || '[;]' || v_�������� || '[;]' ||
                 r_Order.���֤��;
        v_col := 'appno[;]patid[;]patsource[;]pageid[;]regno[;]name[;]sex[;]birthdate[;]address[;]phoneno[;]dept[;]doctor[;]outpatno[;]inpatno[;]ward[;]bedno[;]clinicdiag[;]appdate[;]clinicdesc[;]modality[;]patno[;]partname[;]status[;]emergency[;]patdept[;]age[;]physicalexamid[;]sendno[;]deptno[;]billno[;]regid[;]clinicdiagex[;]executdept[;]feekind[;]paykind[;]executdeptID[;]medicalCardID[;]DoctorEntrust[;]Nation[;]PatientType[;]IDCard';
      
      Else
        --���Ͳ�λҽ��
        v_ҽ������ := 2;
        v_Val      := r_Order.���id || '[;]' || r_Order.ID || '[;]' || r_Order.�����Ŀ����;
        v_col      := 'AppNO[;]AppPartNo[;]ExamPlace';
      End If;
    
      If v_Return Is Null Then
        v_Return := v_ҽ������ || '[:]' || v_Col || '[:]' || v_Val;
      Else
        v_Return := v_Return || '{;}' || v_ҽ������ || '[:]' || v_Col || '[:]' || v_Val;
      End If;
    
      Fetch c_Order
        Into r_Order;
    End Loop;
  
    If c_Order%RowCount = 0 Then
      v_Return := '';
    End If;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
      Return Null;
  End ����ҽ����¼_Send;

End b_Zlxwinterface;
/

--Pacs�ĵ��༭��

--Ӱ�񱨸�ԭ�͹���(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_Common Is
  Type t_Refcur Is Ref Cursor;

--1 ��ȡ�����Ļ�������
Procedure p_GetParInfBuf(
  Val           Out t_Refcur,
  ϵͳ_In In Ӱ�����˵��.ϵͳ%Type,
  ģ��_In  In Ӱ�����˵��.ģ��%Type
  );
  
--2 ���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�
  Function f_Str2list
  (
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
  ) Return t_Strlist
    Pipelined;
  
--3  ��ȡ����ֵ�Ļ�������
--��ǰ�û����ڼ�����Ĳ���ֵ
Procedure p_GetParValueBuf(
  Val           Out t_Refcur,
  ϵͳ_In In Ӱ�����˵��.ϵͳ%Type,
  ģ��_In In Ӱ�����˵��.ģ��%Type,
  ����ID_In In Varchar2,
  ������_In In Varchar2,
  �û�ID_In In Number);

--4  ��ȡȨ�޵Ļ�������
Procedure p_GetPopedomBuf(
  Val           Out t_Refcur,
  ϵͳ_In In Ӱ�����˵��.ϵͳ%Type,
  ģ��_IN In Ӱ�����˵��.ģ��%Type,
  �û���_In In Varchar2);

--5  ���ò���ֵ
Procedure p_SetParameterValue(
  ����ID_In    In Ӱ�����ȡֵ.����ID%Type,
  ������ʶ_In In Ӱ�����ȡֵ.������ʶ%Type,
  ����ֵ_In    In Ӱ�����ȡֵ.����ֵ%Type);

--6  ��ȡ�û��˺���Ϣ
Function f_Get_Personal_Info_By_Account(
	Account_In In Varchar2
) Return Xmltype;

end b_PACS_Common ;

/



--*************************************************************************************
--*                  Ӱ�񱨸�ԭ�͹���(---ʵ�ֲ���---)                                                        *
--*************************************************************************************
CREATE OR REPLACE Package Body b_PACS_Common  Is

--1 ��ȡ�����Ļ�������
Procedure p_GetParInfBuf(
  Val           Out t_Refcur,
  ϵͳ_In In Ӱ�����˵��.ϵͳ%Type,
  ģ��_In  In Ӱ�����˵��.ģ��%Type
) Is
Begin
Open  Val For
   Select RawToHex(ID) As ID,RawToHex(PID) As PID,ϵͳ,ģ��,������,Ĭ��ֵ,��������,��������
   From Ӱ�����˵��
   Where ϵͳ=ϵͳ_In And ģ��=ģ��_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End p_GetParInfBuf;

--2 ���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�
Function f_Str2list(
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
  ) Return t_Strlist
    Pipelined As
    v_Str Long;
    P     Number;
    --���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�
    --������Str_In,��:�׿�,θ����,θ��Ѫ...,Split_In,�ָ���,ȱʡΪ,��
    --˵����
    --1����SQL������漰��IN(����1, ����2,��) ���Ӿ�ʱʹ�����ַ�ʽ�Ա����ð󶨱�����
    --2��ʹ������������ʱ����Ҫ��SQL����м��롰/*+ Rule*/����ʾ����ΪCbo����ʱ�ڴ��û��ͳ������,��
    --3�����ֵ���ʾ��
    --Select /*+ Rule*/ * From Sample_List Where Title In (Select * From Table(f_Str2list('�׿�,θ����,θ��Ѫ'));
    --Select /*+ Rule*/ A.* From Sample_List A, Table(f_Str2list('�׿�,θ����,θ��Ѫ')) B Where A.Title = B.Column_Value;
  Begin
    If Str_In Is Null Then
      Return;
    End If;
    v_Str := Str_In || Split_In;
    Loop
      P := Instr(v_Str, Split_In);
      Exit When(Nvl(P, 0) = 0);
      Pipe Row(Trim(Substr(v_Str, 1, P - 1)));
      v_Str := Substr(v_Str, P + 1);
    End Loop;
    Return;
  End;

--3  ��ȡ����ֵ�Ļ�������
--��ǰ�û����ڼ�����Ĳ���ֵ
Procedure p_GetParValueBuf(
  Val           Out t_Refcur,
  ϵͳ_In In Ӱ�����˵��.ϵͳ%Type,
  ģ��_In In Ӱ�����˵��.ģ��%Type,
  ����ID_In In Varchar2,
  ������_In In Varchar2,
  �û�ID_In In Number
) Is
Begin
	Open  Val For
	    Select RawToHex(b.ID) As ID, RawToHex(b.����ID) As ����ID,b.������ʶ,b.����ֵ
		From Ӱ�����˵�� a, Ӱ�����ȡֵ b
		Where  a.id=b.����id And a.ϵͳ=ϵͳ_In And a.ģ��=ģ��_In and (a.��������=0 or a.��������=1 or (a.��������=2 and b.������ʶ=����ID_In) or (a.��������=3 and b.������ʶ=�û�ID_In) or (a.��������=4 and b.������ʶ=������_In));
Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
End 	p_GetParValueBuf;

--4  ��ȡȨ�޵Ļ�������
Procedure p_GetPopedomBuf(
  Val           Out t_Refcur,
  ϵͳ_In In Ӱ�����˵��.ϵͳ%Type,
  ģ��_In In Ӱ�����˵��.ģ��%Type,
  �û���_In In Varchar2
)Is
Begin
    --�����û�, ģ���,����
	Open  Val For
	    Select a.�û�,b.ϵͳ, b.��� as ģ��, b.����
		From zluserroles a, zlrolegrant b
		Where a.��ɫ=b.��ɫ And a.�û�=�û���_In And b.ϵͳ=ϵͳ_In And b.���=ģ��_In;
Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm);
End p_GetPopedomBuf;

--5  ���²���ֵ
Procedure p_SetParameterValue(
  ����ID_In    In Ӱ�����ȡֵ.����ID%Type,
  ������ʶ_In In Ӱ�����ȡֵ.������ʶ%Type,
  ����ֵ_In    In Ӱ�����ȡֵ.����ֵ%Type
)Is
Begin
	Update Ӱ�����ȡֵ Set ����ֵ=����ֵ_In Where ����ID=����ID_In And ������ʶ=������ʶ_In;
	If Sql%RowCount = 0 Then
	  Insert Into Ӱ�����ȡֵ(ID, ����ID,������ʶ,����ֵ)
	  Values(sys_guid(), ����ID_In,������ʶ_In,����ֵ_In);
	End If;
Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm);
End p_SetParameterValue;

--6  ��ȡ�û��˺���Ϣ
Function f_Get_Personal_Info_By_Account(
	Account_In In Varchar2
) Return Xmltype Is
  Docxml   Xmltype;
Begin 
  Select Xmltype('<root></root>') Into Docxml From Dual;  
  Select Appendchildxml(Docxml, '/root',
                         Xmlconcat(Xmlelement("code", a.Id), Xmlelement("full_name", a.����),
                                    Xmlelement("sex", Xmlattributes(a.�Ա� As "display"),
                                                Decode(a.�Ա�, '��', '1', 'Ů', '2', 'δ֪', '0', '9')),
                                    Xmlelement("birthday", To_Char(a.��������, 'yyyy-mm-dd')),
                                    Xmlelement("idcard_num", a.���֤��)))
  Into Docxml
  From ��Ա�� A, �ϻ���Ա�� B
  Where b.�û��� = Account_In And b.��Աid = a.Id And Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate;  

  Select Appendchildxml(Docxml, '/root',
                         Xmlelement("departments",
                                     Xmlagg(Xmlelement("department", Xmlattributes(c.���� As "display", d.ȱʡ As "current"),
                                                        Xmlelement("dept_value", c.Id)))))
  Into Docxml
  From �ϻ���Ա�� B, ���ű� C, ������Ա D
  Where b.�û��� = Account_In And b.��Աid = d.��Աid And d.����id = c.Id;

  For r_Record In (Select d.����id, Xmlelement("subjects", Xmlagg(Xmlelement("subject", c.����))) As ����ѧ��
                   From �ٴ����� A, �ϻ���Ա�� B, ������Ա D, �ٴ����� C
                   Where b.�û��� = Account_In And b.��Աid = d.��Աid And a.����id = d.����id And a.�������� = c.����
                   Group By d.����id
                   Order By d.����id) Loop
    Select Appendchildxml(Docxml, '/root/departments/department[dept_value=' || r_Record.����id || ']', r_Record.����ѧ��)
    Into Docxml
    From Dual;
  End Loop;

  Return Docxml;
Exception
  When Others Then
    Return Null;
End f_Get_Personal_Info_By_Account;

End b_PACS_Common;

/



--*************************************************************************************
--*								   (---��������---)                                    *
--*************************************************************************************
Create Or Replace Package b_PACS_Config Is
  Type t_Refcur Is Ref Cursor;

  -- ��    �ܣ���ȡӰ���ֵ��嵥
  Procedure p_GetAllDictList(
	Val			Out t_Refcur
  );

  -- ��    �ܣ���ȡӰ���ֵ�����
  Procedure p_GetAllDictItems(
    Val           Out t_Refcur,
	�ֵ�ID_In	In Ӱ���ֵ�����.�ֵ�ID%Type
  );

  -- ��    �ܣ��������޸�Ӱ���ֵ�����
  Procedure p_EditDictItem(
	�ֵ�ID_In		In Ӱ���ֵ�����.�ֵ�ID%Type,
	�ɱ��_In		In Ӱ���ֵ�����.���%Type,
	���_In			In Ӱ���ֵ�����.���%Type,
	����_In			In Ӱ���ֵ�����.����%Type,
	����_In			In Ӱ���ֵ�����.����%Type,
	˵��_In			In Ӱ���ֵ�����.˵��%Type
  );

  -- ��    �ܣ�ɾ��Ӱ���ֵ�����
  Procedure p_DelDictItem(
	�ֵ�ID_In		In Ӱ���ֵ�����.�ֵ�ID%Type,
	���_In			In Ӱ���ֵ�����.���%Type
  );
End b_PACS_Config;
/

--*************************************************************************************
--*								   (---ʵ�ֲ���---)                                    *
--*************************************************************************************
Create Or Replace Package Body b_PACS_Config  Is
  -- ��    �ܣ���ȡӰ���ֵ��嵥
  Procedure p_GetAllDictList(
	Val			Out t_Refcur
  )
  Is
	strSql varchar2(100);
  Begin
	strSql := 'select Rawtohex(ID) ID,���,����,˵��,�Ƿ�ϵͳ,���� From Ӱ���ֵ��嵥';
	Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAllDictList;

  -- ��    �ܣ���ȡӰ���ֵ�����
  Procedure p_GetAllDictItems(
    Val           Out t_Refcur,
	�ֵ�ID_In		In Ӱ���ֵ�����.�ֵ�ID%Type
  )
  Is
	strSql varchar2(200);
  Begin
	strSql := 'Select Rawtohex(A.�ֵ�id) Rid, A.���, A.����, A.����, A.˵�� '||
			  'From Ӱ���ֵ����� A Where A.�ֵ�id = '''|| �ֵ�ID_In ||'''';
	Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAllDictItems;

  -- ��    �ܣ��������޸�Ӱ���ֵ�����
  Procedure p_EditDictItem(
	�ֵ�ID_In		In Ӱ���ֵ�����.�ֵ�ID%Type,
	�ɱ��_In		In Ӱ���ֵ�����.���%Type,
	���_In			In Ӱ���ֵ�����.���%Type,
	����_In			In Ӱ���ֵ�����.����%Type,
	����_In			In Ӱ���ֵ�����.����%Type,
	˵��_In			In Ӱ���ֵ�����.˵��%Type
  )
  Is
	n_Num Number;
	v_Msg Varchar2(50);
	Err	  Exception;
  Begin
	If �ɱ��_In<>'-1' Then
	  Select Count(�ֵ�ID) Into n_Num From Ӱ���ֵ����� Where �ֵ�ID = �ֵ�ID_In And ��� = ���_In And ���<>�ɱ��_In;
	  If n_Num > 0 Then
		v_Msg:='�����ֵ�ID�ͱ���ظ�!';
		Raise Err;
	  End IF;

	  Update Ӱ���ֵ����� A
	  Set A.��� = ���_In,A.���� = ����_In,A.���� = ����_In,A.˵�� = ˵��_In
	  Where A.�ֵ�ID = �ֵ�ID_In And A.��� = �ɱ��_In;
	Else
	  Select Count(�ֵ�ID) Into n_Num From Ӱ���ֵ����� Where �ֵ�ID = �ֵ�ID_In And ��� = ���_In;
	  If n_Num > 0 Then
		v_Msg:='�����ֵ�ID�ͱ���ظ�!';
		Raise Err;
	  End IF;
	
	  Insert Into Ӱ���ֵ�����(�ֵ�ID,���,����,����,˵��)
	  Values(�ֵ�ID_In,���_In,����_In,����_In,˵��_In);
	End If;
  Exception
	When Err Then
	  Raise_Application_Error(-20101,v_Msg);
	When Others Then
	  Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditDictItem;

  -- ��    �ܣ�ɾ��Ӱ���ֵ�����
  Procedure p_DelDictItem(
	�ֵ�ID_In		In Ӱ���ֵ�����.�ֵ�ID%Type,
	���_In			In Ӱ���ֵ�����.���%Type
  )
  Is
  Begin
	Delete From Ӱ���ֵ����� Where �ֵ�ID = �ֵ�ID_In And ��� = ���_In;
  Exception
	When Others Then
	  Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelDictItem;
End b_PACS_Config;
/





Create Or Replace Package b_PACS_RptPublic Is
  Type t_Strlist Is Table Of Varchar2(4000);
  --���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�
  Function f_Str2list(
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
    ) Return t_Strlist
    Pipelined;

  --����:���ݴ���ı�������ȡ��������Code������
  Function f_Get_Nextcode(
    Tablename_In Varchar2,
    Len_In       Number := 0,
    Mount_In     Number := 0,
    Pre_In       Varchar2 := Null
    ) Return Varchar2;

  --���ܣ������ַ���ƴ������
  Function f_Spellcode(
    v_Instr  In Varchar2,
    v_Outnum In Integer := 10
    ) Return Varchar2;

  --����������
  Procedure zl_ErrorCenter(
    Err_Num In Number,
    Err_Msg In Varchar2
    );

 --�Ӵ����XML����ȡ�༭��¼
  Function f_Geteditlist(
    Content_In In Xmltype
	) Return t_Editlist;

  Function Xml2clob(
    Xml_In Xmltype
	) Return Clob;

  Function f_Getlastedit(
    Content_In In Xmltype
	) Return t_Editlist;
  
  ----�����������
  --Function f_Disjoin_Anonym
  --(
    --Content_In    In Xmltype,
    --x_Anonym_Data Out Xmltype
  --) Return Xmltype;

  ----�ϲ���������
  --Function f_Incorporate_Anonym
  --(
    --Content_In    In Xmltype,
    --x_Anonym_Data In Xmltype
  --) Return Clob;

  --����XML�Ľڵ�ֵ,���ڵ�Ϊ<ele></ele>����պϽڵ�ʱ��Updatexml������Ч
  Procedure p_Set_Elementtext(
    Texture In Out Xmltype,
    Ename   In Varchar2,
    Eaname  In Varchar2,
    Eatext  In Varchar2,
    Etext   In Varchar2
    );
  --�������ĵ�ID��ȡ���ĵ���ǰ״̬
 Function f_Get_Docstatus(
    Content_In In Xmltype
    ) Return Varchar2;

  Function f_If_Intersect(
    Str1 Varchar2,
    Str2 Varchar2
    ) Return Number;

End b_PACS_RptPublic;
/

Create Or Replace Package Body b_PACS_RptPublic Is

  Function Xml2clob(
    Xml_In Xmltype
	) Return Clob As
  Begin
    Return Xml_In.Getclobval();
  End Xml2clob;

  Function f_Str2list(
    Str_In   In Varchar2,
    Split_In In Varchar2 := ','
  ) Return t_Strlist
    Pipelined As
    v_Str Long;
    P     Number;
    --���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�
    --������Str_In,��:�׿�,θ����,θ��Ѫ...,Split_In,�ָ���,ȱʡΪ,��
    --˵����
    --1����SQL������漰��IN(����1, ����2,��) ���Ӿ�ʱʹ�����ַ�ʽ�Ա����ð󶨱�����
    --2��ʹ������������ʱ����Ҫ��SQL����м��롰/*+ Rule*/����ʾ����ΪCbo����ʱ�ڴ��û��ͳ������,��
    --3�����ֵ���ʾ��
    --Select /*+ Rule*/ * From Sample_List Where Title In (Select * From Table(f_Str2list('�׿�,θ����,θ��Ѫ'));
    --Select /*+ Rule*/ A.* From Sample_List A, Table(f_Str2list('�׿�,θ����,θ��Ѫ')) B Where A.Title = B.Column_Value;
  Begin
    If Str_In Is Null Then
      Return;
    End If;
    v_Str := Str_In || Split_In;
    Loop
      P := Instr(v_Str, Split_In);
      Exit When(Nvl(P, 0) = 0);
      Pipe Row(Trim(Substr(v_Str, 1, P - 1)));
      v_Str := Substr(v_Str, P + 1);
    End Loop;
    Return;
  End;

  Function f_Get_Nextcode(
    Tablename_In Varchar2,
    Len_In       Number := 0,
    Mount_In     Number := 0,
    Pre_In       Varchar2 := Null
  ) Return Varchar2 Is
    --���ݴ���ı�������ȡ��������Code������
    --Len_In        ��ָ������ʱ����ָ���������Code���������ڱ��ֶγ��Ȼ򲻴�ʱΪ���ֶγ���
    --Mount_In      ��ָ������ʱ�����Codeÿλ�����ٽ�λʱ�Ƿ����ݳ��ȣ�=0�����ݣ���1���� ���統ǰ���CodeΪ Z99,��������򷵻�1A00���򷵻�Z99
    --������������0123456789 ��ĸA...Z,��������9��Zʱǰһλ�������ҵ�ǰλתΪ0��A�����ǰһ��Ϊ�����ֻ���ĸ����ǰһλ����
    --��ĸȫ��Ϊ��д����Сд,����ָ������Codeʱ����ָ������01������ָ������Ϊ5���򷵻� 00001
    v_Sql      Varchar2(100);
    v_Maxcode  Varchar2(150);
    v_Origcode Varchar2(150);
    v_Old      Varchar2(6);
    v_New      Varchar2(6);
    v_Return   Varchar2(150);
    n_Collen   Number;
    n_Length   Number;
    Err_Custom Exception;
    v_Msg Varchar2(200);
    Function f_Char_Carry(Word_In Varchar2) Return Varchar2 As
      v_Temp Varchar2(6);
      n_Asc  Number;
    Begin
      Select Ascii(Upper(Word_In)) Into n_Asc From Dual;
      If n_Asc = 57 Then
        v_Temp := '0';
      Elsif n_Asc = 90 Then
        v_Temp := 'A';
      Elsif n_Asc >= 48 And n_Asc <= 56 Or n_Asc >= 65 And n_Asc <= 89 Then
        v_Temp := Chr(Ascii(Word_In) + 1);
      Else
        v_Temp := Word_In;
      End If;
      Return v_Temp;
    End;
  Begin
    Begin
      Select Data_Length
      Into n_Collen
      From User_Tab_Cols
      Where Table_Name = Upper(Tablename_In) And Upper(Column_Name) = '����';
    Exception
      When Others Then
        Null;
        v_Msg := 'û�е�ǰҪ���ҵı������û�С����롿�ֶΣ�';
        Raise Err_Custom;
    End;
  
    --�����볤��Ϊ0������ֶγ���ʱȡ��ǰ��󳤶ȣ�����ȡ���볤���൱��������
    If Len_In = 0 Or Len_In > n_Collen Then
      v_Sql := 'Select Max(Length(����)) From ' || Tablename_In;
      Execute Immediate v_Sql
        Into n_Length;
    Else
      n_Length := Len_In;
    End If;
  
    If Nvl(n_Length, 0) = 0 Then
      Return '1';
    End If;
    
    If (Pre_In Is Not Null) And Length(Pre_In) >= n_Length Then
      v_Msg := 'ָ������ĳ���Ӧ�ô���ǰ׺����';
      Raise Err_Custom;
    End If;
  
    --����ָ��ǰ׺��������ֵ
    If Pre_In Is Not Null Then
      v_Sql := 'Select Max(����) From ' || Tablename_In || ' Where upper(substr(code,1,length(''' || Pre_In ||
               '''))) =' || '' || 'upper(''' || Pre_In || ''')';
      Execute Immediate v_Sql
        Into v_Maxcode;
    Else
      v_Sql := 'Select Max(����) From ' || Tablename_In || ' Where Length(����)=' || n_Length;
      Execute Immediate v_Sql
        Into v_Maxcode;
    End If;
  
    --�������codeΪ�գ���ô��ֵΪ1,ǰ������0������������Ⱦ���
    If v_Maxcode Is Null Then
      If Pre_In Is Null Then
        Select LPad('1', n_Length, '0') Into v_Maxcode From Dual;
        Return v_Maxcode;
      Else
        Select Pre_In || LPad('1', n_Length - Length(Pre_In), '0') Into v_Maxcode From Dual;
        Return v_Maxcode;
      End If;
    Else
      If Pre_In Is Null Then
        v_Maxcode  := Upper(v_Maxcode);
        v_Origcode := v_Maxcode;
      Else
        --����Ϊָ������
        v_Maxcode := Upper(Pre_In || LPad(Substr(v_Maxcode, Length(Pre_In) + 1), n_Length - Length(Pre_In), '0'));
        --ָ��ǰ׺ʱ������ǰ׺����ַ�����Ϊ�����ֵ
        v_Origcode := Substr(v_Maxcode, Length(Pre_In) + 1);
        v_Maxcode  := v_Origcode;
      End If;
      
      For I In 0 .. Length(v_Maxcode) Loop
        If I = Length(v_Maxcode) Then
          If Len_In <> 0 And Mount_In = 0 Then
            --ָ�����Ȳ��Ҳ����ݳ��ȣ����ڵ����ʱ�����ٽ�λ
            Return v_Origcode;
          Else
            v_Old := 'Z';
            v_New := '1';
          End If;
        Else
          v_Old := Substr(v_Maxcode, Length(v_Maxcode) - I, 1);
          v_New := f_Char_Carry(v_Old);
        End If;
      
        --�¾�ֵ��ȱ���Ϊ�����ֻ���ĸ,��Ҫ��ǰ����
        If v_Old != v_New Then
          v_Return := Substr(v_Maxcode, 0, Length(v_Maxcode) - I - 1) || v_New ||
                      Substr(v_Maxcode, Length(v_Maxcode) - I + 1);
          If v_New != '0' And v_New != 'A' Then
            If Pre_In Is Null Then
              Return v_Return;
            Else
              Return Pre_In || v_Return;
            End If;
          Else
            --���ڡ�0��������ǰ��λ,ǰһλ����
            v_Maxcode := v_Return;
          End If;
        End If;
      End Loop;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

  --���ܣ������ַ���ƴ������
  --������v_Instr��Ҫ����ƴ�����ַ�����v_Outnum �������볤�ȣ�Ĭ��10������40���ַ����10
  Function f_Spellcode(
    v_Instr  In Varchar2,
    v_Outnum In Integer := 10
  ) Return Varchar2 Is
    v_Spell     Varchar2(40);
    v_Input     Varchar2(1000);
    v_Bitchar   Varchar2(100);
    r_Bitchar   Varchar2(100);
    v_Bitnum    Integer;
    v_Outmaxnum Integer;
    Function f_Nlssort(p_Word In Varchar2) Return Varchar2 As
    Begin
      Return Nlssort(p_Word, 'NLS_SORT=SCHINESE_PINYIN_M');
    End;
  Begin
    If v_Outnum < 1 Or v_Outnum > 40 Then
      v_Outmaxnum := 10;
    Else
      v_Outmaxnum := v_Outnum;
    End If;
  
    If v_Instr Is Null Or Length(LTrim(v_Instr)) = 0 Then
      v_Spell := '';
    Else
      v_Input := Upper(v_Instr);
      v_Spell := '';
      For v_Bitnum In 1 .. Length(v_Input) Loop
        v_Bitchar := Substr(v_Input, v_Bitnum, 1);
        r_Bitchar := f_Nlssort(v_Bitchar);
        If r_Bitchar >= f_Nlssort('߹') And r_Bitchar <= f_Nlssort('�') Then
          v_Spell := v_Spell || 'A';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('��') Then
          v_Spell := v_Spell || 'B';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�e') Then
          v_Spell := v_Spell || 'C';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�z') Then
          v_Spell := v_Spell || 'D';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('��') Then
          v_Spell := v_Spell || 'E';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�g') Then
          v_Spell := v_Spell || 'F';
        Elsif r_Bitchar >= f_Nlssort('�') And r_Bitchar <= f_Nlssort('�B') Then
          v_Spell := v_Spell || 'G';
        Elsif r_Bitchar >= f_Nlssort('�o') And r_Bitchar <= f_Nlssort('��') Then
          v_Spell := v_Spell || 'H';
        Elsif r_Bitchar >= f_Nlssort('آ') And r_Bitchar <= f_Nlssort('�h') Then
          v_Spell := v_Spell || 'J';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�i') Then
          v_Spell := v_Spell || 'K';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�^') Then
          v_Spell := v_Spell || 'L';
        Elsif r_Bitchar >= f_Nlssort('�`') And r_Bitchar <= f_Nlssort('��') Then
          v_Spell := v_Spell || 'M';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('��') Then
          v_Spell := v_Spell || 'N';
        Elsif r_Bitchar >= f_Nlssort('�p') And r_Bitchar <= f_Nlssort('�a') Then
          v_Spell := v_Spell || 'O';
        Elsif r_Bitchar >= f_Nlssort('�r') And r_Bitchar <= f_Nlssort('��') Then
          v_Spell := v_Spell || 'P';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�d') Then
          v_Spell := v_Spell || 'Q';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�U') Then
          v_Spell := v_Spell || 'R';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�R') Then
          v_Spell := v_Spell || 'S';
        Elsif r_Bitchar >= f_Nlssort('�@') And r_Bitchar <= f_Nlssort('�X') Then
          v_Spell := v_Spell || 'T';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('�F') Then
          v_Spell := v_Spell || 'W';
        Elsif r_Bitchar >= f_Nlssort('Ϧ') And r_Bitchar <= f_Nlssort('�R') Then
          v_Spell := v_Spell || 'X';
        Elsif r_Bitchar >= f_Nlssort('Ѿ') And r_Bitchar <= f_Nlssort('�') Then
          v_Spell := v_Spell || 'Y';
        Elsif r_Bitchar >= f_Nlssort('��') And r_Bitchar <= f_Nlssort('��') Then
          v_Spell := v_Spell || 'Z';
        Elsif Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', v_Bitchar) > 0 Then
          v_Spell := v_Spell || v_Bitchar;
        Elsif Instr('���������������', v_Bitchar) > 0 Then
          v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41664);
        Elsif Instr('���£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣�', v_Bitchar) > 0 Then
          v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41856);
        Elsif Instr('����', v_Bitchar) > 0 Then
          v_Spell := v_Spell || 'A';
        Elsif Instr('����', v_Bitchar) > 0 Then
          v_Spell := v_Spell || 'B';
        Elsif Instr('����', v_Bitchar) > 0 Then
          v_Spell := v_Spell || 'G';
        End If;
        Exit When Length(v_Spell) > Nvl(v_Outmaxnum, 40) - 1;
      End Loop;
    End If;
    Return(v_Spell);
  End;

  Procedure zl_ErrorCenter(
    Err_Num In Number,
    Err_Msg In Varchar2
  ) Is
    v_Outnum Number := 0;
    v_Outmsg Varchar2(1000) := '';
    v_Temp   Varchar2(1000) := '';
  
    Cursor Cur_Ind_Cols Is
      Select Table_Name, Column_Name From All_Ind_Columns Where Instr(Err_Msg, Index_Owner || '.' || Index_Name) > 0;
  
    Cursor Cur_Con_Cols Is
      Select Table_Name, Column_Name
      From All_Cons_Columns
      Where (Owner, Constraint_Name) =
            (Select r_Owner, r_Constraint_Name
             From All_Constraints
             Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 0 And Rownum < 2);
  Begin
    If Err_Num = -1 Then
      For Row_Cols In Cur_Ind_Cols Loop
        v_Temp   := Row_Cols.Table_Name;
        v_Outmsg := v_Outmsg || '��' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]' || v_Temp || '��(' || Substr(v_Outmsg, 2) || ')�����ظ���[ZLSOFT]';
      v_Outnum := -20000;
    Elsif Err_Num = -1000 Then
      v_Outmsg := '[ZLSOFT]�򿪵����ݱ�̫�࣬��Ҫʱ��ϵͳ����Ա�޸����ݿ��Open_Cursors���á�';
      v_Outnum := -20001;
    Elsif Err_Num = -1400 Or Err_Num = -1407 Then
      Select Table_Name, Column_Name
      Into v_Temp, v_Outmsg
      From All_Tab_Columns
      Where Instr(Err_Msg, '"' || Owner || '"."' || Table_Name || '"."' || Column_Name || '"') > 0 And Rownum < 2;
      v_Outmsg := '[ZLSOFT]' || v_Temp || '��(' || v_Outmsg || ')�������룡[ZLSOFT]';
      v_Outnum := -20002;
    Elsif Err_Num = -1401 Then
      v_Outmsg := '[ZLSOFT]���ڸ����ֵ�������п����ƣ��������ӻ����ʧ�ܡ�[ZLSOFT]';
      v_Outnum := -20003;
    Elsif Err_Num = -2290 Then
      Select Table_Name, Search_Condition
      Into v_Temp, v_Outmsg
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      If Instr(v_Outmsg, 'IS NOT NULL') > 0 Then
        v_Outmsg := '[ZLSOFT]' || v_Temp || ' �� ' || Replace(v_Outmsg, 'IS NOT NULL', '�������룡') || '[ZLSOFT]';
        v_Outnum := -20004;
      Else
        v_Outmsg := Err_Msg;
        v_Outnum := -20999;
      End If;
    Elsif Err_Num = -2292 Then
      Select Table_Name
      Into v_Temp
      From All_Constraints
      Where Instr(Err_Msg, Owner || '.' || Constraint_Name) > 1 And Rownum < 2;
    
      For Row_Cols In Cur_Con_Cols Loop
        v_Outmsg := v_Outmsg || '��' || Row_Cols.Column_Name;
      End Loop;
    
      v_Outmsg := '[ZLSOFT]�ü�¼�� ' || v_Temp || ' ���Ѿ�ʹ��,' || Chr(13) || '����ɾ�����޸�(' || Substr(v_Outmsg, 2) || ')[ZLSOFT]';
      v_Outnum := -20005;
    Else
      v_Outmsg := Err_Msg;
      v_Outnum := -20999;
    End If;
    Raise_Application_Error(v_Outnum, Substr(v_Outmsg, 1, 100));
  End zl_ErrorCenter;

  --���ĵ�����ȡ�ı༭��¼
  Function f_Geteditlist(
    Content_In In Xmltype
	) Return t_Editlist As
    --��ȡ�ĵ��༭��ǩ�����޶���¼,���ظ�ʽ�������£������ĵ�SUBIIDΪ�գ���
    --  Subiid  Subaid  �༭��  �༭ʱ��             ǩ�� ��ǩ��
    --  AAAAAA  AID     Null    Null                 0    0     ��һ�����ڱ�ʾ������¼
    --  AAAAAA  AID     ���ջ�  2012-05-31 12:01:02  1    0
    --  AAAAAA  AID     ����    2012-05-31 12:05:02  0    0
    --  AAAAAA  AID     ����    2012-05-31 12:06:02  1    1
    --  AAAAAA  AID     ����    2012-05-31 12:07:02  0    0
    --  AAAAAA  AID     ���ջ�  2012-05-31 12:08:02  0    0
    --  AAAAAA  AID     ���ջ�  2012-05-31 12:09:02  1    1
    --  BBBBBB............
    Content_c Clob;
    Xcdoc     Xmldom.Domdocument;
  
    Targetdoc Dbms_Xmldom.Domdocument;
  
    Signlist     Xmldom.Domnodelist;
    l_s          Number;
    n_Isnull     Number;
    Signname     Varchar2(64);
    Signtime     Date;
    Isaduit      Number(1);
    Xxdoc        Xmltype;
    Xaudit       Xmltype;
    Xa_Text      Xmltype;
    Textlist     Xmldom.Domnodelist;
    l_t          Number;
    Starttime    Date;
    Aftertime    Date;
    Aduitname    Varchar2(64);
    Aduittime    Date;
    Revisiontime Varchar2(20);
    Ts_Editlist  t_Editlist;
    Ta_Editlist  t_Editlist;
    r_Editlist   t_Editlist := t_Editlist();
    
    Function Sortbytime(t_e t_Editlist) Return t_Editlist Is
      Tm_List t_Editlist := t_Editlist();
    Begin
      For Rs In (Select * From Table(Cast(t_e As t_Editlist)) A Order By a.�༭ʱ��) Loop
        Tm_List.Extend;
        Tm_List(Tm_List.Count) := t_Edits( Rs.�༭��, Rs.�༭ʱ��, Rs.ǩ��, Rs.��ǩ��);
      End Loop;
      Return Tm_List;
    End Sortbytime;
    
  Begin
    If Content_In Is Null Then
      r_Editlist := t_Editlist();
      Return r_Editlist;
    End If;
    --ͼƬ���ܻᳬ��64K,���ڵ㳬��64KʱNewdomdocument��崻�,Newdomdocument(clob)��ʽ��ż������"ͨ��ͨ���ļ�����"
    Content_c := Xml2clob(Content_In);
    
    --�����ĵ�,ֱ�Ӹ��ĵ���ֵ
    Xcdoc   := Xmldom.Newdomdocument(Content_c);
    
    Signlist    := Xmldom.Getelementsbytagname(Xcdoc, 'signature');
    l_s         := Xmldom.Getlength(Signlist);
    l_s         := Nvl(l_s, 0);
    Ts_Editlist := t_Editlist();
    
    --��������ǩ����¼
    For L In 0 .. l_s - 1 Loop
      --��ȡǩ��λ��1-ǩ��λ��0-��ʵǩ��
      n_Isnull := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'isnull');
      
      If Nvl(n_Isnull, 0) = 0 Then
        --�������ǩ��λ���Ǿ���һ����ʵǩ��
        --displayinfo ǩ����ʾ��Ϣ
        Signname := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'displayinfo');
        If Signname Is Not Null Then
          --��ǩ����ʾ��Ϣ����Ϊ��
          Select Substr(Signname, 1, Decode(Instr(Signname, ','), 0, Length(Signname) + 1, Instr(Signname, ',')) - 1)
          Into Signname
          From Dual;
          --signtime ǩ��ʱ��
          Revisiontime := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'signtime');
          Signtime     := To_Date(Revisiontime, 'yyyy-mm-dd hh24:mi:ss');
          --isaudit ��ǩ���
          Isaduit      := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Signlist, L)), 'isaudit');
          Isaduit      := Nvl(Isaduit, 0);
          Ts_Editlist.Extend;
          Ts_Editlist(Ts_Editlist.Count) := t_Edits( Signname, Signtime, 1, Isaduit);
        End If;
      End If;
    End Loop;
    
    Xxdoc := Xmldom.Getxmltype(Xcdoc);
    --����������������ɾ��Ϊ��ǵ��޶���¼
    --ratag �޶��������,ȡֵΪϵͳ��¼�˺�����;  rdtag �޶�ɾ�����,ȡֵΪϵͳ��¼�˺�����
    Xa_Text     := Xxdoc.Extract('//*[@ratag!=""]|//*[@rdtag!=""]');
    Xaudit      := Xmltype('<root></root>');
    Xaudit      := Xaudit.Appendchildxml('/root', Xa_Text);
    Textlist    := Xmldom.Getchildnodes(Xmldom.Getfirstchild(Xmldom.Makenode(Xmldom.Newdomdocument(Xaudit))));
    l_t         := Xmldom.Getlength(Textlist);
    Ta_Editlist := t_Editlist();
    For L In 0 .. l_t - 1 Loop
      --ratag �޶��������,ȡֵΪϵͳ��¼�˺�����
      Aduitname := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'ratag');
      If Nvl(Aduitname, 'a') = 'a' Then
        --��ratagȡ����ֵ��˵����ɾ����¼
        Revisiontime := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'rdtime');
        Aduittime    := To_Date(Revisiontime, 'yyyy-mm-dd hh24:mi:ss');
        Aduitname    := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'rdtag');
      Else
        Revisiontime := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Textlist, L)), 'ratime');
        Aduittime    := To_Date(Revisiontime, 'yyyy-mm-dd hh24:mi:ss');
      End If;
      Ta_Editlist.Extend;
      Ta_Editlist(Ta_Editlist.Count) := t_Edits( Aduitname, Aduittime, 0, 0);
    End Loop;
    --�Ȱ�ʱ������
    Ts_Editlist := Sortbytime(Ts_Editlist);
    Ta_Editlist := Sortbytime(Ta_Editlist);
    --�����һ�����ڱ�ʾ������¼
    r_Editlist.Extend;
    r_Editlist(r_Editlist.Count) := t_Edits( Null,
                                            To_Date('1945-08-06 09:16:02', 'yyyy-mm-dd hh24:mi:ss'), 0, 0);
    
    --����ǩ�����´�ǩ��֮�䱻�϶�Ϊ�޶���¼,�������ɱ༭�б�
    For L In 1 .. Ts_Editlist.Count Loop
      Starttime := Ts_Editlist(L).�༭ʱ��;
      If L = Ts_Editlist.Count Then
        --ֻ����Ϊ��ʱ��������ǩ��ʱ��֮����жϲ��գ���ѭ�������һ��ǩ��ʱ���˲���ʧȥ����,���Ը�ֵ���Դ��ڵ�ǰϵͳʱ��
        Aftertime := Sysdate + 1;
      Else
        Aftertime := Ts_Editlist(L + 1).�༭ʱ��;
      End If;
      r_Editlist.Extend;
      r_Editlist(r_Editlist.Count) := Ts_Editlist(L);
      
      Aduitname := 'A';
      For N In 1 .. Ta_Editlist.Count Loop
        Aduittime := Ta_Editlist(N).�༭ʱ��;
        If Aduittime Between Starttime And Aftertime Then
          Starttime := Aduittime;
          If Aduitname <> Ta_Editlist(N).�༭�� Then
            --��ͬ�˵��޶���¼���һ���޶���¼
            Aduitname := Ta_Editlist(N).�༭��;
            r_Editlist.Extend;
            r_Editlist(r_Editlist.Count) := t_Edits( Aduitname, Aduittime, 0, 0);
          Else
            --ͬһ�˲�ͬʱ��ദ�޶�,ֻȡ���һ��ʱ��
            r_Editlist(r_Editlist.Count).�༭ʱ�� := Aduittime;
          End If;
        Elsif Aduittime>Aftertime then
          --��Ϊ�޶���¼����ʱ�������������ǩ����¼֮�������´�ǩ������޶�
          Exit;
        End If;
      End Loop;
    End Loop;
  
    If Not Xmldom.Isnull(Xcdoc) Then
      Xmldom.Freedocument(Xcdoc);
    End If;
  
    If Not Xmldom.Isnull(Targetdoc) Then
      Xmldom.Freedocument(Targetdoc);
    End If;
  
    Return r_Editlist;
  End f_Geteditlist;

  Function f_Getlastedit(
    Content_In In Xmltype
	) Return t_Editlist As
    t_List   t_Editlist := t_Editlist();
    r_List   t_Editlist := t_Editlist();
    
    Function Lastlist(
      t_e       t_Editlist
    ) Return t_Editlist Is
      Tm_List t_Editlist := t_Editlist();
    Begin
      For Rs In (Select * From Table(Cast(t_e As t_Editlist)) A  Order By a.�༭ʱ�� Desc) Loop
        Tm_List.Extend;
        Tm_List(Tm_List.Count) := t_Edits( Rs.�༭��, Rs.�༭ʱ��, Rs.ǩ��, Rs.��ǩ��);
        Return Tm_List;
      End Loop;
    End Lastlist;
    
  Begin
    Select f_Geteditlist(Content_In) Into t_List From Dual;
    r_List.Extend;
    r_List(r_List.Count) := Lastlist(t_List) (1);
    Return r_List;
  End f_Getlastedit;
  
  
  --����XML�Ľڵ�ֵ,���ڵ�Ϊ<ele></ele>����պϽڵ�ʱ��Updatexml������Ч
  Procedure p_Set_Elementtext(
    Texture In Out Xmltype,
    Ename   In Varchar2,
    Eaname  In Varchar2,
    Eatext  In Varchar2,
    Etext   In Varchar2
  ) Is
    --������     Texture ������XML
    --           Ename ��Ҫ���õĽڵ�����
    --           Eaname ��Ҫ���õĽڵ����������ƣ����ھ�ȷ��λ
    --           Eatext ��Ҫ���õĽڵ�������ֵ�����ھ�ȷ��λ
    --           Ttext ��Ҫ���õ�Ԫ��ֵ
    x_Dom   Xmldom.Domdocument;
    x_Nlist Xmldom.Domnodelist;
    x_Text  Xmldom.Domnode;
    x_Node  Xmldom.Domnode;
    n_Len   Number;
    v_Val   Varchar2(2000);
    Procedure Freeall Is
    Begin
      If Not Xmldom.Isnull(x_Text) Then
        Xmldom.Freenode(x_Text);
      End If;
    
      If Not Xmldom.Isnull(x_Node) Then
        Xmldom.Freenode(x_Node);
      End If;
    
      If Not Xmldom.Isnull(x_Dom) Then
        Xmldom.Freedocument(x_Dom);
      End If;
    End Freeall;
  Begin
    If Texture Is Null Then
      Return;
    End If;
  
    x_Dom   := Xmldom.Newdomdocument(Texture);
    x_Nlist := Xmldom.Getelementsbytagname(x_Dom, Ename);
    n_Len   := Xmldom.Getlength(x_Nlist);
    For I In 0 .. n_Len - 1 Loop
      If Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(x_Nlist, I)), Eaname) = Eatext Then
        For J In 0 .. Xmldom.Getlength(Xmldom.Getchildnodes(Xmldom.Item(x_Nlist, I))) - 1 Loop
          x_Node := Xmldom.Item(Xmldom.Getchildnodes(Xmldom.Item(x_Nlist, I)), J);
          If Xmldom.Getnodetype(x_Node) = Xmldom.Text_Node Then
            --�ҵ��ı��ڵ�
            v_Val := Xmldom.Getnodevalue(x_Node);
            Exit;
          Else
            v_Val := Null;
          End If;
        End Loop;
      
        If v_Val Is Null Then
          x_Text := Xmldom.Makenode(Xmldom.Createtextnode(x_Dom, Etext));
          x_Text := Xmldom.Importnode(x_Dom, x_Text, True);
          x_Node := Xmldom.Appendchild(Xmldom.Item(x_Nlist, I), x_Text);
        Else
          Xmldom.Setnodevalue(x_Node, Etext);
        End If;
        Texture := Xmldom.Getxmltype(x_Dom);
        Freeall;
        Return;
      End If;
    End Loop;
  
    Freeall;
  End p_Set_Elementtext;

  --�������ĵ�ID��ȡ���ĵ���ǰ״̬
  Function f_Get_Docstatus(
    Content_In In Xmltype
  ) Return Varchar2 Is
    n_Sign   Number;
    n_Audit  Number;
    v_Editor Varchar2(200);
    v_n      Varchar2(20);
  Begin
    For Rs In (Select *
               From Table(Cast((Select b_PACS_RptPublic.f_Geteditlist(Content_In) From Dual) As t_Editlist))
               Order By �༭ʱ�� Asc) Loop
      n_Sign   := Rs.ǩ��;
      n_Audit  := Rs.��ǩ��;
      v_Editor := Rs.�༭��;
    End Loop;
    If n_Sign = 0 And n_Audit = 0 And v_Editor Is Null Then
      v_n := '�༭��';
    Elsif n_Sign = 1 And n_Audit = 0 Then
      v_n := '��ǩ��';
    Elsif n_Sign = 0 And n_Audit = 0 And v_Editor Is Not Null Then
      v_n := '����';
    Elsif n_Sign = 1 And n_Audit = 1 Then
      v_n := '����ǩ';
    End If;
    Return v_n;
  End f_Get_Docstatus;

  Function f_If_Intersect
  (
    Str1 Varchar2,
    Str2 Varchar2
  ) Return Number As
    n_Num Number;
  Begin
  
    Select Count(*)
    Into n_Num
    From (Select a.Column_Value Value
           From Table(b_PACS_RptPublic.f_Str2list(Str1, ',')) A
           Intersect
           Select b.Column_Value Value From Table(b_PACS_RptPublic.f_Str2list(Str2, ',')) B);
  
    Return n_Num;
  End;

Begin
  -- Initialization
  Null;

End b_PACS_RptPublic;
/


--Ӱ�񱨸�ԭ�͹���(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptCommon Is
  Type t_Refcur Is Ref Cursor;

  --��ȡԤ�����>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	);

  --Ԫ�ط���>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Eleclass(
    Val Out t_Refcur
	);

  --ԭ��Ƭ��>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	);

  --ԭ���б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_By_Id(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	);

  --ԭ������>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_Content(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	);

  --�����嵥>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Samplelist_By_Aid(
    Val           Out t_Refcur,
	Antetypelist_Id_In Varchar2,
	Condition_In       Ӱ�񱨸淶���嵥.����%Type,
	Author_In          Ӱ�񱨸淶���嵥.����%Type,
	Subjects_In        Ӱ�񱨸淶���嵥.ѧ��%Type
	);

  --��ȡ������ø��ݲ��ID��ȡ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigById(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸���.ID%Type
	);

  --��ȡ������ø���ԭ���嵥��ȡ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigByAId(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ���嵥.ID%Type
	);

  --��ȡԪ��>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Element(
    Val Out t_Refcur
	);

  --��ȡƬ���б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	);

  --��ȡֵ���б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Range_List(
    Val Out t_Refcur
	);

  --��ȡֵ���б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Combo_List(
    Val Out t_Refcur
	);

  --��ȡԭ��Ƭ�θ���ԭ��ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentDirectory_ByAid(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	);

  --����ԭ��id��ȡƬ������
  Procedure p_Get_FragmentData_ByAid(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	);

  --��ȡ���ݱ��������ʱ��>>>>>>>>>>>>>>>>>>>>>>>>>>>
  procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
	Table_Name_In Varchar2
	);

  --��ȡƬ���б�����ϼ�ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Bypid(
    Val           Out t_Refcur,
	Pid_In Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type
	);

  --��ȡƬ���б���ݽڵ�����>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Byleaf(
    Val           Out t_Refcur,
	Leaf_In Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type
	);

  --��ȡֵ����Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Range_List_Byid(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ֵ���嵥.Id%Type
	);

  --����Ԫ��ID��ȡֵ��ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Getelementrid_By_Eid(
    Val           Out t_Refcur,
	Eid_In Ӱ�񱨸�Ԫ���嵥.Id%Type
	);

  --��ȡ������λ�б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_GetMasure_UnitList(
    Val Out t_Refcur
	);

  --��ȡ�ĵ�������Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Doc_Kind(
    Val Out t_Refcur
	);

  --���ܣ���ȡ����ѧ����Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Subjects(
    Val Out t_Refcur
	);

  --�鿴�Ƿ������Ӧ�ı����������(���ڵ��뵼��)>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exits_Doc_Kinds(
    Val           Out t_Refcur,
	����_In      Varchar2,
	����_In      Varchar2,
	Tablename_In Varchar2
	);

  --�Ƿ������ͬ��ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exist_Id(
    Val           Out t_Refcur,
	Id_In        Number,
	Tablename_In Varchar2
	);

  --ͨ�����ƻ�ȡID��Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Id_By_Title(
    Val           Out t_Refcur,
	����_In      Varchar2,
	Tablename_In Varchar2,
	Type_In      Varchar2
	);

  --ͨ�����Ƭ���嵥
  procedure p_Get_FragmentSampleName(
    Val           Out t_Refcur,
	���_In Varchar2
	);

  --����ID��Ӧ��Ƭ������
  procedure p_Update_PhraseContent(
    Id_In      Ӱ�񱨸�Ƭ���嵥.ID%type,
	Name_In		Ӱ�񱨸�Ƭ���嵥.����%Type,
	Content_In Varchar2
	);
  --��ȡԭ��ID��Ӧ�ĵ�һ��Ƭ�νڵ�
  procedure p_Get_FragmentData_LevelOne(
    Val           Out t_Refcur,
	AId_In Ӱ�񱨸�ԭ���嵥.ID%type
	);

  -- ��ȡƬ�ε��²�ڵ�
  procedure p_GetFragmentDataListByFID(
    Val           Out t_Refcur,
	FId_In Ӱ�񱨸�Ƭ���嵥.ID%type
	);
end b_PACS_RptCommon;
/



--*************************************************************************************
--*                  Ӱ�񱨸�ԭ�͹���(---ʵ�ֲ���---)                                                        *
--*************************************************************************************
CREATE OR REPLACE Package Body b_PACS_RptCommon Is
  -- ��    �ܣ��÷���ֻ������ʾ...

  --��ȡԤ�����>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) As ID, a.����, a.����, a.˵��
        From Ӱ�񱨸�Ԥ����� A
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --Ԫ�ط���>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Eleclass(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(A.ID) As ID,
             A.����,
             A.����,
             A.˵��,
             RawToHex(A.�ϼ�ID) �ϼ�ID
        From Ӱ�񱨸�Ԫ�ط��� A
       Order By ����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --ԭ��Ƭ��>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(A.Ƭ��ID) As Ƭ��ID
        From Ӱ�񱨸�ԭ��Ƭ�� A
       Where a.ԭ��ID = Aid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --ԭ���嵥>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_By_Id(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select /*+rule*/
       RawToHex(A.ID) As ID,
       a.����,
       a.����,
       a.����,
       a.˵��,
       a.�ɷ�����ҳ�� As ҳ������,
       a.�ɷ����ø�ʽ As ��ʽ����,
       Extractvalue(b.Column_Value, '/root/print_hf_mode') Printhfmode,
       Extractvalue(b.Column_Value, '/root/print_follow_pages') Printfollowpages,
       Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit,
       (Nvl(a.����ѡ��, XmlType('<NULL/>'))).GetClobVal() As ����ѡ��,
       a.������,
       a.����ʱ��,
       a.�޸���,
       a.�޸�ʱ��,
       a.�Ƿ����,
       A.����
        From Ӱ�񱨸�ԭ���嵥 A,
             Table(Xmlsequence(Extract(a.����ѡ��, '/root'))) B
       Where a.Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --ԭ������>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Antetypelist_Content(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select (Nvl(a.����, XmlType('<ZLXML/>'))).GetClobVal() As ����
        From Ӱ�񱨸�ԭ���嵥 A
       Where a.Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --ͨ��ԭ��ID�����Ӧ�ķ�����Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Samplelist_By_Aid(
    Val           Out t_Refcur,
	Antetypelist_Id_In Varchar2,
	Condition_In       Ӱ�񱨸淶���嵥.����%Type,
	Author_In          Ӱ�񱨸淶���嵥.����%Type,
	Subjects_In        Ӱ�񱨸淶���嵥.ѧ��%Type
	) As
  Begin
    --ֱ�ӻ�ȡ��ԭ���µķ����б�
    If Length(Antetypelist_Id_In) > 30 Then
      Open Val For
        Select /*+rule*/
         RawToHex(A.ID) as ID,
         a.����,
         a.����,
         a.˵��,
         a.ѧ��,
         a.���,
         a.��ǩ,
         a.�Ƿ�˽��
          From Ӱ�񱨸淶���嵥 A
         Where a.ԭ��ID = Hextoraw(Antetypelist_Id_In)
           And (a.���� = Author_In Or (a.ѧ�� Is Null And a.�Ƿ�˽�� = 0) Or
               Subjects_In Is Null Or
               (a.ѧ�� Is Not Null And
               b_PACS_RptPublic.f_If_Intersect(a.ѧ��, Subjects_In) > 0 And
               a.�Ƿ�˽�� = 0));
    Else
      --���һ������ԭ����Ϣ�ķ������νṹ
      Open Val For
        Select Distinct a.���� As ID,
                        a.���� As ����,
                        '' as ˵��,
                        '' As ԭ��ID,
                        'category' As ����,
                        '' As ����,
                        '' As ѧ��,
                        Null as ���༭ʱ��,
                        '' As ��ǩ,
                        0 As �Ƿ�˽��,
                        0 As Imgindex
          From Ӱ�񱨸�ԭ���嵥 A
         Where a.���� = Antetypelist_Id_In
           And Exists
         (Select ID From Ӱ�񱨸淶���嵥 C Where c.ԭ��ID = a.Id)
           And a.���� Is Not Null
        Union
        Select m.*
          From (Select RawToHex(B.ID) As ID,
                       b.����,
                       b.˵��,
                       b.���� As ԭ��ID,
                       'antetype' As ����,
                       '' As ����,
                       '' As ѧ��,
                       Null as ���༭ʱ��,
                       '' As ��ǩ,
                       0 As �Ƿ�˽��,
                       0 As Imgindex
                  From Ӱ�񱨸�ԭ���嵥 B
                 Where b.���� = Antetypelist_Id_In
                   And Exists (Select ID
                          From Ӱ�񱨸淶���嵥 C
                         Where c.ԭ��ID = b.Id)
                 Order By b.����) M
        
        Union All
        Select n.*
          From (Select /*+rule*/
                 RawToHex(A.ID) As ID,
                 a.����,
                 a.˵��,
                 RawToHex(A.ԭ��ID) As ԭ��ID,
                 'sample' As ����,
                 a.����,
                 a.ѧ��,
                 a.���༭ʱ��,
                 a.��ǩ,
                 a.�Ƿ�˽��,
                 Decode(a.�Ƿ�˽��, 1, 2, 1) As Imgindex
                  From Ӱ�񱨸淶���嵥 A, Ӱ�񱨸�ԭ���嵥 C
                 Where a.ԭ��ID = c.Id
                   And c.���� = Antetypelist_Id_In
                   And ((a.���� Like '%' || Condition_In || '%' And
                       Condition_In Is Not Null) Or Condition_In Is Null)
                   And (a.���� = Author_In Or (a.ѧ�� Is Null And a.�Ƿ�˽�� = 0) Or
                       Subjects_In Is Null Or
                       (a.ѧ�� Is Not Null And
                       b_PACS_RptPublic.f_If_Intersect(a.ѧ��, Subjects_In) > 0 And
                       a.�Ƿ�˽�� = 0))
                 Order By a.���, a.����) N;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡ������ø��ݲ��ID��ȡ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigById(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸���.ID%Type
	) As
    v_Sql Varchar2(1000);
  Begin
    If (Id_In Is Not Null) Then
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.����, t.˵��, t.����, t.����, t.����, t.�Ƿ����,t.����,T.��ʾ��ʽ  From Ӱ�񱨸��� T Where t.Id =:Id_In And Rownum = 1';
      
        Open Val For v_Sql
          Using Id_In;
      End;
    Else
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.����, t.˵��, t.����, t.����, t.����, t.�Ƿ����,t.����,T.��ʾ��ʽ From Ӱ�񱨸��� T where �Ƿ���� = 0 order by t.����';
      
        Open Val For v_Sql;
      End;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡ������ø���ԭ���嵥��ȡ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Plugin_ConfigByAId(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ���嵥.ID%Type
	) As
    v_Sql Varchar2(1000);
  Begin
    If (Aid_In Is Not Null) Then
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.����, t.˵��, t.����, t.����, t.����, t.�Ƿ����,t.����,T.��ʾ��ʽ ' ||
                 ' From Ӱ�񱨸��� T ' || ' Where t.Id in ( ' ||
                 ' Select X.pluginid from Ӱ�񱨸�ԭ���嵥 K, ' ||
                 '  (XMLTable(''*//pluginid''  Passing K.ר�ò�� Columns pluginid varchar2(32) Path ''/pluginid''))  X ' ||
                 ' Where K.id=:Aid_In) And �Ƿ���� = 0' || ' Union All ' ||
                 'Select RawToHex(T.ID) as ID, t.����, t.˵��, t.����, t.����, t.����, t.�Ƿ����,t.����,T.��ʾ��ʽ ' ||
                 ' From Ӱ�񱨸��� T ' || ' Where �Ƿ���� = 0 And t.����=0 ';
      
        Open Val For v_Sql
          Using Aid_In;
      End;
    Else
      Begin
        v_Sql := 'Select RawToHex(T.ID) as ID, t.����, t.˵��, t.����, t.����, t.����, t.�Ƿ����,t.����,T.��ʾ��ʽ From Ӱ�񱨸��� T where �Ƿ���� = 0 order by t.����';
      
        Open Val For v_Sql;
      End;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡ����Ԫ��>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Element(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) as ID,
             RawToHex(T.����ID) as ����ID,
             T.����,
             T.����,
             T.ǰ׺,
             T.��׺,
             T.˵��,
             T.��������,
             T.��ֵ��̬,
             T.��С����,
             T.��󳤶�,
             T.��СС��λ,
             T.���С��λ,
             T.������λ,
             (Nvl(T.��չ����, XmlType('<NULL/>'))).GetClobVal() As ��չ����,
             T.ֵ��ID,
             T.ֵ������
        From Ӱ�񱨸�Ԫ���嵥 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡƬ���б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.�ϼ�ID) As �ϼ�ID,
             t.����,
             t.����,
             t.˵��,
             t.�ڵ�����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             t.ѧ��,
             t.��ǩ,
             t.�Ƿ�˽��,
             t.����,
             t.���༭ʱ��,
			 (Nvl(t.��Ӧ����, XmlType('<NULL/>'))).GetClobVal() As ��Ӧ����
        From Ӱ�񱨸�Ƭ���嵥 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡֵ���б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Range_List(
    Val Out t_Refcur
	) as
  begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.����ID) As ����ID,
             T.����,
             T.����,
             T.˵��,
             T.��������,
             T.ֵ������,
             (Nvl(t.ֵ������, XmlType('<NULL/>'))).GetClobVal() As ֵ������,
             t.���༭ʱ��
        From Ӱ�񱨸�ֵ���嵥 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end;

  --��ȡ����б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Combo_List(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             T.����,
             T.����,
             T.˵��,
             T.����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             T.�༭��,
             T.���༭ʱ��,
             T.����
        From Ӱ�񱨸�����嵥 T;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡԭ��Ƭ��Ŀ¼����ԭ��ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentDirectory_ByAid(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	) As
  Begin
    Open Val For
      Select ����,
             RawToHex(ID) As ID,
             4 as ImageIndex,
             RawToHex(�ϼ�ID) As �ϼ�ID,
             '<NULL/>' As ���,
             ����,
             �ڵ�����,
             �Ƿ�˽��,
             ����,
             ��ǩ,
			 ˵��,
             ѧ��
        From Ӱ�񱨸�Ƭ���嵥
       Where ID In (Select ID
                      From Ӱ�񱨸�Ƭ���嵥
                     Start With ID In (Select Ƭ��ID
                                         From Ӱ�񱨸�ԭ��Ƭ��
                                        Where ԭ��ID = Aid_In)
                    Connect By Prior �ϼ�ID = ID)
       order by ����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡԭ��Ƭ�����ݸ���ԭ��ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_FragmentData_ByAid(
    Val           Out t_Refcur,
	Aid_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	) As
  Begin
    Open Val For
      With TabFragmentId As
       (Select Ƭ��ID From Ӱ�񱨸�ԭ��Ƭ�� Where ԭ��ID = Aid_In)
      Select RawToHex(ID) As ID,
             RawToHex(�ϼ�ID) As �ϼ�ID,
             ����,
             ����,
             ˵��,
             �ڵ�����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             ѧ��,
             ��ǩ,
             �Ƿ�˽��,
             ����,
             ���༭ʱ��,
             EXTRACTValue( ��Ӧ����, '/Root/Rad') ��Ŀ,
             EXTRACTValue( ��Ӧ����, '/Root/Part') ��λ, 
             EXTRACTValue( ��Ӧ����, '/Root/Kind') ���,
             EXTRACTValue( ��Ӧ����, '/Root/Sex') �Ա�,
             0 as ���״̬,
             0 as ��Ӧ״̬
        From Ӱ�񱨸�Ƭ���嵥 t
       Where Id Not In (Select Ƭ��ID From TabFragmentId)
       Start With ID In (Select Ƭ��ID From TabFragmentId)
      Connect By Prior �ϼ�ID = ID
      Union All
      Select RawToHex(ID) As ID,
             RawToHex(�ϼ�ID) As �ϼ�ID,
             ����,
             ����,
             ˵��,
             �ڵ�����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             ѧ��,
             ��ǩ,
             �Ƿ�˽��,
             ����,
             ���༭ʱ��,
             EXTRACTValue( ��Ӧ����, '/Root/Rad') ��Ŀ,
             EXTRACTValue( ��Ӧ����, '/Root/Part') ��λ, 
             EXTRACTValue( ��Ӧ����, '/Root/Kind') ���,
             EXTRACTValue( ��Ӧ����, '/Root/Sex') �Ա�,
             0 as ���״̬,
             0 as ��Ӧ״̬
        From Ӱ�񱨸�Ƭ���嵥 t
       Start With ID In (Select Ƭ��ID From TabFragmentId)
      Connect By Prior ID = �ϼ�ID
       order by ����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡ���ݱ��������ʱ��>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
	Table_Name_In Varchar2
	) As
    v_sql Varchar2(4000);
  Begin
    v_sql := 'select max(���༭ʱ��) maxvalue from ' || Table_Name_In;
    Open val For v_sql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡƬ���б�����ϼ�ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Bypid(
    Val           Out t_Refcur,
	Pid_In Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.�ϼ�ID) As �ϼ�ID,
             T.����,
             T.����,
             T.˵��,
             T.�ڵ�����,
             (Nvl(T.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             T.ѧ��,
             T.��ǩ,
             T.�Ƿ�˽��,
             T.����,
			 (Nvl(T.��Ӧ����, XmlType('<NULL/>'))).GetClobVal() As ��Ӧ����,
             T.���༭ʱ��
        From Ӱ�񱨸�Ƭ���嵥 T
       Where T.�ϼ�ID = Pid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --ͨ���ڵ����ͻ�ȡ�ʾ��б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Fragment_List_Byleaf(
    Val           Out t_Refcur,
	Leaf_In Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.�ϼ�ID) As �ϼ�ID,
             T.����,
             T.����,
             T.˵��,
             T.�ڵ�����,
             (Nvl(T.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             T.ѧ��,
             T.��ǩ,
             T.�Ƿ�˽��,
             T.����,
             T.���༭ʱ��
        From Ӱ�񱨸�Ƭ���嵥 T
       Where t.�ڵ����� = Leaf_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡֵ����Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Range_List_Byid(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ֵ���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.����ID) As ����ID,
             T.����,
             T.����,
             T.˵��,
             T.��������,
             T.ֵ������,
             (Nvl(T.ֵ������, XmlType('<NULL/>'))).GetClobVal() As ֵ������
        From Ӱ�񱨸�ֵ���嵥 T
       Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --����Ԫ��ID��ȡֵ��ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Getelementrid_By_Eid(
    Val           Out t_Refcur,
	Eid_In Ӱ�񱨸�Ԫ���嵥.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(A.ֵ��ID) As ֵ��ID
        From Ӱ�񱨸�Ԫ���嵥 A
       Where a.Id = Eid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --��ȡ������λ�б�>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_GetMasure_UnitList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ����, ����, ˵��, ǰ׺ From Ӱ�񱨸������λ;
  End p_GetMasure_UnitList;

  --��ȡ�ĵ�������Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Doc_Kind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select a.����, a.����, a.˵�� From Ӱ�񱨸����� A Order By a.����;
  End;

  --���ܣ���ȡ����ѧ����Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_All_Subjects(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select rawtohex(b.�ֵ�id) ID, b.��� As ����, b.����, b.����, b.˵��
        From Ӱ���ֵ��嵥 A, Ӱ���ֵ����� B
       Where a.���� = 'רҵѧ��'
         And a.Id = b.�ֵ�id
       Order By ����;
  End;

  --�鿴�Ƿ������Ӧ�ı����������(���ڵ��뵼��)>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exits_Doc_Kinds(
    Val           Out t_Refcur,
	����_In      Varchar2,
	����_In      Varchar2,
	Tablename_In Varchar2
	) As
    v_Type Varchar2(50);
    n_Num  Number;
    v_Sql  Varchar2(100);
  Begin
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE ���� =''' ||
             ����_In || ''' AND ���� =''' || ����_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '1';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE ���� <>''' ||
             ����_In || ''' AND ���� = ''' || ����_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '2';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE ���� =''' ||
             ����_In || ''' or ���� =''' || ����_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num = 0 Then
      v_Type := '3';
    End If;
    v_Sql := 'SELECT COUNT(*) FROM ' || Tablename_In || ' WHERE ���� =''' ||
             ����_In || ''' AND ���� <>''' || ����_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    If n_Num > 0 Then
      v_Type := '4';
    End If;
    Open Val For
      Select v_Type As Type From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --�Ƿ������ͬ��ID>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_If_Exist_Id(
    Val           Out t_Refcur,
	Id_In        Number,
	Tablename_In Varchar2
	) As
    v_Sql Varchar2(100);
    n_Num Number;
  Begin
    v_Sql := 'select count(id) from ' || Tablename_In || ' where id=''' ||
             Id_In || '''';
    Execute Immediate v_Sql
      Into n_Num;
    Open Val For
      Select Decode(n_Num, 0, 0, 1) Num From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --ͨ�����ƻ�ȡID��Ϣ>>>>>>>>>>>>>>>>>>>>>>>>>>>
  Procedure p_Get_Id_By_Title(
    Val           Out t_Refcur,
	����_In      Varchar2,
	Tablename_In Varchar2,
	Type_In      Varchar2
	) As
    v_Id  Varchar2(50);
    v_Sql Varchar2(100);
  Begin
    If Type_In = '1' Then
      v_Sql := 'select id from ' || Tablename_In || ' where ����=''' || ����_In || '''';
    Else
      v_Sql := 'select ���� from ' || Tablename_In || ' where ����=''' || ����_In || '''';
    End If;
    v_Id := '';
    Execute Immediate v_Sql
      Into v_Id;
    Open Val For
      Select v_Id ID From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --ͨ�����Ƭ���嵥
  Procedure p_Get_FragmentSampleName(
	Val           Out t_Refcur,
	���_In Varchar2
	) as
  Begin
    Open Val For
      Select RawToHex(T.ID) As ID,
             RawToHex(T.�ϼ�ID) As �ϼ�ID,
             t.����,
             t.����,
             t.˵��,
             t.�ڵ�����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             t.ѧ��,
             t.��ǩ,
             t.�Ƿ�˽��,
             t.����,
             t.���༭ʱ��
        From Ӱ�񱨸�Ƭ���嵥 T
      Where t.���� LIKE '%' || ���_In || '%';
      --Where  F_TRANS_PINYIN_CAPITAL(t.����) LIKE '%' || ���_In || '%';
  End p_Get_FragmentSampleName;

  --����ID��Ӧ��Ƭ������
  Procedure p_Update_PhraseContent(
	Id_In      Ӱ�񱨸�Ƭ���嵥.ID%Type,
	Name_In		Ӱ�񱨸�Ƭ���嵥.����%Type,
	Content_In Varchar2
	) as
  Begin
    Update Ӱ�񱨸�Ƭ���嵥 t 
	Set t.��� = Content_In, t.����=Name_In 
	Where t.id = Id_In;
  End p_Update_PhraseContent;

  --��ȡԭ��ID��Ӧ�ĵ�һ��Ƭ�νڵ�
  Procedure p_Get_FragmentData_LevelOne(
	Val           Out t_Refcur,
	AId_In Ӱ�񱨸�ԭ���嵥.ID%Type
	) as
  Begin
    Open Val For
      With TabFragmentId As
       (Select Ƭ��ID From Ӱ�񱨸�ԭ��Ƭ�� Where ԭ��ID = Aid_In)
      Select RawToHex(ID) As ID,
             RawToHex(�ϼ�ID) As �ϼ�ID,
             ����,
             ����,
             ˵��,
             �ڵ�����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             ѧ��,
             ��ǩ,
             �Ƿ�˽��,
             ����,
             ���༭ʱ��
        From Ӱ�񱨸�Ƭ���嵥 t
       Where Id In (Select Ƭ��ID From TabFragmentId);
  
  End p_Get_FragmentData_LevelOne;

  -- ��ȡƬ�ε��²�ڵ�
  Procedure p_GetFragmentDataListByFID(
	Val           Out t_Refcur,
	FId_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	) as
  Begin
    Open Val For
      Select RawToHex(ID) As ID,
             RawToHex(�ϼ�ID) As �ϼ�ID,
             ����,
             ����,
             ˵��,
             �ڵ�����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             ѧ��,
             ��ǩ,
             �Ƿ�˽��,
             ����,
             ���༭ʱ��
        From Ӱ�񱨸�Ƭ���嵥 t
       Where �ϼ�ID = FId_In;
  End p_GetFragmentDataListByFID;

End b_PACS_RptCommon;

/




   --Ӱ�񱨸����---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptParam Is
  Type t_Refcur Is Ref Cursor;

  --����1����û�Ĳ����б�
  Procedure p_GetPrograms(
    Val Out t_Refcur
	);
  --����2��ͨ��ģ��Ż�ȡӰ�������Ϣ
  Procedure p_GetParamByQum(
    Val           Out t_Refcur,
	ģ��_In Ӱ�����˵��.ģ��%Type
	);
  --����3��ͨ������ID�Ż�ȡӰ�����ȡֵ��Ϣ
  Procedure p_GetParamValue(
    Val           Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	);
  --����4����ò�����Ϣ
  Procedure p_GetDepart(
    Val Out t_Refcur
	);
  --����5�������Ա��Ϣ
  Procedure p_GetUsersInfo(
    Val Out t_Refcur
	);
  --����6����û�������Ϣ
  Procedure p_GetMachinesInfo(
    Val Out t_Refcur
	);
  --����7����ȡ���е�Ӱ�������Ϣ
  Procedure p_GetAllParam(
    Val Out t_Refcur
	);
  --����8��������в��ŵ����еĲ���ȡֵ��Ϣ
  Procedure p_GetValueAllDepart(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	);
  --����9����ID��Ӧ���ŵ����еĲ���ȡֵ��Ϣ
  Procedure p_GetValueByDepart(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type,
	����ID_In ���ű�.ID%Type
	);
  --����9����ȡ���е��û���Ӧ�Ĳ���ֵ
  Procedure p_GetValueAllUser(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	);
  --����10����ȡ�û�ID��Ӧ�Ĳ�����Ϣ
  Procedure p_GetValueByUser(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type,
	�û�ID_In ��Ա��.ID%Type
	);
  --����11����ȡ���еĹ���վ��Ӧ�Ĳ���ֵ
  Procedure p_GetValueAllMachine(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	);
  --����12����ȡ����վ���ƶ�Ӧ�Ĳ�����Ϣ
  Procedure p_GetValueByMachine(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type,
	������_In zlclients.����վ%Type
	);
  --����13����Ӳ�����Ϣ
  Procedure p_AddParamValue(
    ID_In       Ӱ�����ȡֵ.ID%Type,
    ����ID_In   Ӱ�����ȡֵ.����ID%Type,
    ������ʶ_In Ӱ�����ȡֵ.������ʶ%Type,
    ����ֵ_In   Ӱ�����ȡֵ.����ֵ%Type
	);

  --����14���޸Ĳ�����Ϣ
  Procedure p_EditParamValue(
    ID_In       Ӱ�����ȡֵ.ID%Type,
    ������ʶ_In Ӱ�����ȡֵ.������ʶ%Type,
    ����ֵ_In   Ӱ�����ȡֵ.����ֵ%Type
	);

  --����15:ͨ��ID��ò�����Ϣ
  Procedure p_GetParamByID(
    Val Out t_Refcur,
	ID_In Ӱ�����˵��.ID%Type
	);
  --����16���޸�ID��Ӧ�Ĳ�������
  Procedure p_ChangeAdjustByID(
    ID_In     Ӱ�����˵��.ID%Type,
	Adjust_In Ӱ�����˵��.��������%Type
	);
  --����17����ö�Ӧ������ʶ�Ĳ���ȡֵ��Ϣ
  Procedure p_GetValueBySign(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	);
  --����18���޸�ID��Ӧ�Ĳ�����Ϣ��Ĭ��ֵ
  Procedure p_EditDaultValue(
    ID_In     Ӱ�����˵��.ID%Type,
	Ĭ��ֵ_In Ӱ�����˵��.Ĭ��ֵ%Type);

  --����19��ͨ�����Ż����Ա��Ϣ
  Procedure p_GetUserByDID(
    Val Out t_Refcur,
	DID_In ������Ա.����ID%Type
	);
  --����21:ͨ��ID��ò���ȡֵ
  Procedure p_GetParamValueByCID(
    Val Out t_Refcur,
	CID_In Ӱ�����ȡֵ.����ID%Type
	);
  --����22:ͨ��ID���ģ��ŵĲ���ȡֵ
  Procedure p_GetValueLevel0(
    Val Out t_Refcur,
	����ID_In   Ӱ�����ȡֵ.����ID%Type,
	������ʶ_In Ӱ�����ȡֵ.������ʶ%Type
	);
end b_PACS_RptParam;
/

--Ӱ�񱨸����---ʵ�ֲ���---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptParam Is
  --create by hwei;

  --����1����û�Ĳ����б�
  Procedure p_GetPrograms(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ���, ����, Decode(A.����, '', ���, ��� || '-' || ����) ����
        From (Select Distinct (t.ģ��) ���,
                              (Select y.����
                                 From zlprograms y
                                Where to_char(y.���) = t.ģ��) ����
                From Ӱ�����˵�� t) A;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetPrograms;
  --����2��ͨ��ģ��Ż�ȡӰ�������Ϣ
  Procedure p_GetParamByQum(
    Val Out t_Refcur,
	ģ��_In Ӱ�����˵��.ģ��%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(PID) PID,
             t.����,
             t.�������,
             t.������,
             t.Ĭ��ֵ,
             t.��������,
             t.ȡֵ��Χ,
             t.��������,
             t.˵��,
             '�� ��' ����ֵ
        From Ӱ�����˵�� t
       Where t.ģ�� = ģ��_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamByQum;

  --����3��ͨ������ID�Ż�ȡӰ�����ȡֵ��Ϣ
  Procedure p_GetParamValue(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	) As
    paramPiont nvarchar2(50);
    paramLevel Number := -1;
    paramCount Number := -1;
  Begin
    Select a.��������
      Into paramLevel
      From Ӱ�����˵�� a
     Where a.id = ����ID_In
       And rownum <= 1;
    If paramLevel = 1 Then
      Select count(t.id)
        Into paramCount
        From Ӱ�����ȡֵ t
       Where t.����id = ����ID_In;
      IF paramCount <> 0 THEN
        Select t.������ʶ
          Into paramPiont
          From Ӱ�����ȡֵ t
         Where t.����id = ����ID_In;
      End If;
      IF paramCount <> 0 And
         Replace(translate(paramPiont, '0123456789', '0'), '0', '') IS NULL THEN
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(����ID) ����ID,
                 t.������ʶ,
                 (Select a.����
                    From zlprograms a
                   Where a.��� = t.������ʶ
                     And rownum <= 1) As ��ʶ����,
                 t.����ֵ
            From Ӱ�����ȡֵ t
           Where t.����id = ����ID_In;
      Else
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(����ID) ����ID,
                 t.������ʶ,
                 t.������ʶ As ��ʶ����,
                 t.����ֵ
            From Ӱ�����ȡֵ t
           Where t.����id = ����ID_In;
      End If;
    Elsif paramLevel = 2 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(����ID) ����ID,
               t.������ʶ,
               (Select a.���� From ���ű� a Where a.id = t.������ʶ) as ��ʶ����,
               t.����ֵ
          From Ӱ�����ȡֵ t
         Where t.����id = ����ID_In;
    Elsif paramLevel = 3 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(����ID) ����ID,
               t.������ʶ,
               (Select a.����
                  From ��Ա�� a
                 Where a.id = t.������ʶ
                   And rownum <= 1) As ��ʶ����,
               t.����ֵ
          From Ӱ�����ȡֵ t
         Where t.����id = ����ID_In;
    Else
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(����ID) ����ID,
               t.������ʶ,
               t.������ʶ As ��ʶ����,
               t.����ֵ
          From Ӱ�����ȡֵ t
         Where t.����id = ����ID_In;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamValue;

  --����4����ò�����Ϣ
  Procedure p_GetDepart(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ID, �ϼ�ID, t.����, t.���� from ���ű� t Order by t.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDepart;
  --����5�������Ա��Ϣ
  Procedure p_GetUsersInfo(Val Out t_Refcur) As
  Begin
    Open Val For
      Select ID, t.���, t.����, t.����, t.���֤��
        From ��Ա�� t
       Order by t.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetUsersInfo;

  --����6����û�������Ϣ
  Procedure p_GetMachinesInfo(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select t.����վ, t.ip, t.���� From zlclients t Order by t.����վ;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetMachinesInfo;

  --����7����ȡ���е�Ӱ�������Ϣ
  Procedure p_GetAllParam(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(PID) PID,
             t.����,
             t.�������,
             t.������,
             t.Ĭ��ֵ,
             t.��������,
             t.ȡֵ��Χ,
             t.��������,
             t.˵��
        From Ӱ�����˵�� t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAllParam;
  --����8��������в��ŵ����еĲ���ȡֵ��Ϣ
  Procedure p_GetValueAllDepart(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	) As
  Begin
    Open Val For
      Select t.id, t.����id, t.������ʶ, t.����ֵ
        From Ӱ�����ȡֵ t
       Where t.����id = ����ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueAllDepart;

  --����9����ID��Ӧ���ŵ����еĲ���ȡֵ��Ϣ
  Procedure p_GetValueByDepart(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type,
	����ID_In ���ű�.ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.id) As ID,
             RawToHex(����ID_In) As ����ID,
             s.id As ������ʶ,
             (Select t.����ֵ
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.id) As ����ֵ
        From ���ű� s
       Where s.id = ����ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueByDepart;

  --����9����ȡ���е��û���Ӧ�Ĳ���ֵ
  Procedure p_GetValueAllUser(
    Val Out t_Refcur,
    ����ID_In Ӱ�����ȡֵ.����ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.id) As ID,
             RawToHex(����ID_In) As ����ID,
             s.id As ������ʶ,
             (Select t.����ֵ
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.id) As ����ֵ
        From ��Ա�� s;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueAllUser;

  --����10����ȡ�û�ID��Ӧ�Ĳ�����Ϣ
  Procedure p_GetValueByUser(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type,
	�û�ID_In ��Ա��.ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.id) As ID,
             RawToHex(����ID_In) As ����ID,
             s.id As ������ʶ,
             (Select t.����ֵ
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.id) As ����ֵ
        From ��Ա�� s
       Where s.id = �û�ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueByUser;

  --����11����ȡ���еĹ���վ��Ӧ�Ĳ���ֵ
  Procedure p_GetValueAllMachine(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.����վ) As ID,
             RawToHex(����ID_In) As ����ID,
             s.����վ As ������ʶ,
             (Select t.����ֵ
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.����վ) As ����ֵ
        From zlclients s;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueAllMachine;

  --����12����ȡ����վ���ƶ�Ӧ�Ĳ�����Ϣ
  Procedure p_GetValueByMachine(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type,
	������_In zlclients.����վ%Type
	) As
  Begin
    Open Val For
      Select (Select RawToHex(ID) ID
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.����վ) as ID,
             RawToHex(����ID_In) as ����ID,
             s.����վ as ������ʶ,
             (Select t.����ֵ
                From Ӱ�����ȡֵ t
               Where t.����id = ����ID_In
                 And t.������ʶ = s.����վ) as ����ֵ
        From zlclients s
       Where s.����վ = ������_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueByMachine;
  --����13����Ӳ�����Ϣ
  Procedure p_AddParamValue(
    ID_In       Ӱ�����ȡֵ.ID%Type,
    ����ID_In   Ӱ�����ȡֵ.����ID%Type,
    ������ʶ_In Ӱ�����ȡֵ.������ʶ%Type,
    ����ֵ_In   Ӱ�����ȡֵ.����ֵ%Type
	) As
  Begin
    Insert Into Ӱ�����ȡֵ t
      (ID, ����ID, ������ʶ, ����ֵ)
    ValueS
      (ID_In, ����ID_In, ������ʶ_In, ����ֵ_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddParamValue;

  --����14���޸Ĳ�����Ϣ
  Procedure p_EditParamValue(
    ID_In       Ӱ�����ȡֵ.ID%Type,
    ������ʶ_In Ӱ�����ȡֵ.������ʶ%Type,
    ����ֵ_In   Ӱ�����ȡֵ.����ֵ%Type
	) As
  Begin
    Update Ӱ�����ȡֵ t
       Set ������ʶ = ������ʶ_In, ����ֵ = ����ֵ_In
     Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditParamValue;
  --����15:ͨ��ID��ò�����Ϣ
  Procedure p_GetParamByID(
    Val Out t_Refcur,
	ID_In Ӱ�����˵��.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(PID) PID,
             t.����,
             t.�������,
             t.������,
             t.Ĭ��ֵ,
             t.��������,
             t.ȡֵ��Χ,
             t.��������,
             t.˵��
        From Ӱ�����˵�� t
       Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamByID;
  --����16���޸�ID��Ӧ�Ĳ�������
  Procedure p_ChangeAdjustByID(
    ID_In     Ӱ�����˵��.ID%Type,
	Adjust_In Ӱ�����˵��.��������%Type) As
  Begin
    Delete From Ӱ�����ȡֵ a Where a.����id = ID_In;
    Update Ӱ�����˵�� t Set t.�������� = Adjust_In Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_ChangeAdjustByID;

  --����17����ö�Ӧ������ʶ�Ĳ���ȡֵ��Ϣ
  Procedure p_GetValueBySign(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	) As
  Begin
    Open Val For
      Select t.id, t.����id, t.������ʶ, t.����ֵ
        From Ӱ�����ȡֵ t
       Where t.����id = ����ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueBySign;

  --����18���޸�ID��Ӧ�Ĳ�����Ϣ��Ĭ��ֵ
  Procedure p_EditDaultValue(
    ID_In     Ӱ�����˵��.ID%Type,
	Ĭ��ֵ_In Ӱ�����˵��.Ĭ��ֵ%Type) As
  Begin
    Update Ӱ�����˵�� t Set t.Ĭ��ֵ = Ĭ��ֵ_In Where t.id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditDaultValue;

  --����19��ͨ�����Ż����Ա��Ϣ
  Procedure p_GetUserByDID(
    Val Out t_Refcur,
	DID_In ������Ա.����ID%Type
	) As
  Begin
    Open Val For
      Select ID, t.���, t.����, t.����, t.���֤��
        From ��Ա�� t
       Where t.id In
             (Select a.��Աid From ������Ա a Where a.����id = DID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetUserByDID;
  --����20��ͨ������ID�Ż�ȡӰ�����ȡֵ��Ϣ
  Procedure p_GetParamValue1(
    Val Out t_Refcur,
	����ID_In Ӱ�����ȡֵ.����ID%Type
	) As
    paramLevel number := -1;
    paramPiont nvarchar2(50);
    paramCount number := -1;
  Begin
    Select a.��������
      Into paramLevel
      From Ӱ�����˵�� a
     Where a.id = ����ID_In
       And rownum <= 1;
    If paramLevel = 1 then
      Select Count(t.id)
        into paramCount
        From Ӱ�����ȡֵ t
       Where t.����id = ����ID_In;
      Select t.������ʶ
        into paramPiont
        From Ӱ�����ȡֵ t
       Where t.����id = ����ID_In;
      IF paramCount <> 0 and
         Replace(translate(paramPiont, '0123456789', '0'), '0', '') IS NULL THEN
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(����ID) ����ID,
                 (Select a.����
                    From zlprograms a
                   Where a.��� = t.������ʶ
                     And rownum <= 1) As ������ʶ,
                 t.����ֵ
            From Ӱ�����ȡֵ t
           Where t.����id = ����ID_In;
      Else
        Open Val For
          Select RawToHex(ID) ID,
                 RawToHex(����ID) ����ID,
                 t.������ʶ,
                 t.����ֵ
            From Ӱ�����ȡֵ t
           Where t.����id = ����ID_In;
      End if;
    Elsif paramLevel = 2 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(����ID) ����ID,
               (Select a.����
                  From ���ű� a
                 Where a.id = t.������ʶ
                   And rownum <= 1) As ������ʶ,
               t.����ֵ
          From Ӱ�����ȡֵ t
         Where t.����id = ����ID_In;
    Elsif paramLevel = 3 Then
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(����ID) ����ID,
               (Select a.����
                  From ��Ա�� a
                 Where a.id = t.������ʶ
                   And rownum <= 1) As ������ʶ,
               t.����ֵ
          From Ӱ�����ȡֵ t
         Where t.����id = ����ID_In;
    Else
      Open Val For
        Select RawToHex(ID) ID,
               RawToHex(����ID) ����ID,
               t.������ʶ,
               t.����ֵ
          From Ӱ�����ȡֵ t
         Where t.����id = ����ID_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamValue1;
  --����21:ͨ��ID��ò���ȡֵ
  Procedure p_GetParamValueByCID(
    Val Out t_Refcur,
	CID_In Ӱ�����ȡֵ.����ID%Type
	) As
  Begin
    Open Val For
      Select t.������ʶ, t.����ֵ
        From Ӱ�����ȡֵ t
       Where t.����id = CID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetParamValueByCID;

  --����22:ͨ��ID���ģ��ŵĲ���ȡֵ
  Procedure p_GetValueLevel0(
    Val Out t_Refcur,
	����ID_In   Ӱ�����ȡֵ.����ID%Type,
	������ʶ_In Ӱ�����ȡֵ.������ʶ%Type
	) As
  Begin
    Open Val For
      Select t.������ʶ, t.����ֵ
        From Ӱ�����ȡֵ t
       Where t.����id = ����ID_In
         And t.������ʶ = ������ʶ_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetValueLevel0;
End b_PACS_RptParam;
/


--Ӱ�񱨸�Ԫ��ֵ��(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptElement Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  Procedure p_GetElementClassList(
    Val Out t_Refcur
	);

  --2.��  �ܣ�����Ӱ�񱨸�Ԫ�ط�����Ϣ
  Procedure p_AddElementClass(
    ID_In   In Ӱ�񱨸�Ԫ�ط���.ID%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	˵��_In In Ӱ�񱨸�Ԫ�ط���.˵��%Type,
	�ϼ�ID_In In Ӱ�񱨸�Ԫ�ط���.�ϼ�ID%Type
	);

  --3.��  �ܣ��޸�Ӱ�񱨸�Ԫ�ط�����Ϣ
  Procedure p_EditElementClass(
    ID_In   In Ӱ�񱨸�Ԫ�ط���.ID%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	˵��_In In Ӱ�񱨸�Ԫ�ط���.˵��%Type,
	�ϼ�ID_In In Ӱ�񱨸�Ԫ�ط���.�ϼ�ID%Type
	);

  --4.��  �ܣ�ɾ��Ӱ�񱨸�Ԫ�ط�����Ϣ
  Procedure p_DelelEmentClass(
    ID_In In Ӱ�񱨸�Ԫ�ط���.ID%Type
	);

  --5.��  �ܣ���÷����Ӧ��Ӱ�񱨸�ֵ����Ϣ�б�
  Procedure p_GetRangeByClass(
    Val           Out t_Refcur,
	����ID_In In Ӱ�񱨸�ֵ���嵥.����ID%Type
	);

  --6.��  �ܣ����ID��Ӧ��Ӱ�񱨸�ֵ����Ϣ
  Procedure p_GetRangeByID(
    Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�ֵ���嵥.ID%Type
	);

  --7.��  �ܣ�����Ӱ�񱨸�ֵ����Ϣ
  Procedure p_AddRange(
    ID_In          In Ӱ�񱨸�ֵ���嵥.ID%Type,
	����ID_In    In Ӱ�񱨸�ֵ���嵥.����ID%Type,
	����_In       In Ӱ�񱨸�ֵ���嵥.����%Type,
	����_In       In Ӱ�񱨸�ֵ���嵥.����%Type,
	˵��_In       In Ӱ�񱨸�ֵ���嵥.˵��%Type,
	��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
	ֵ������_In In Varchar2);

  --8.��  �ܣ��޸�Ӱ�񱨸�ֵ����Ϣ
  Procedure p_EditRange(
    ID_In         In Ӱ�񱨸�ֵ���嵥.ID%Type,
	����ID_In   In Ӱ�񱨸�ֵ���嵥.����ID%Type,
	����_In       In Ӱ�񱨸�ֵ���嵥.����%Type,
	����_In       In Ӱ�񱨸�ֵ���嵥.����%Type,
	˵��_In       In Ӱ�񱨸�ֵ���嵥.˵��%Type,
	��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
	ֵ������_In In Varchar2
	);

  --9.��  �ܣ�ɾ��Ӱ�񱨸�ֵ����Ϣ
  Procedure p_DelRange(
    ID_In In Ӱ�񱨸�ֵ���嵥.ID%Type
	);

  --10.��  �ܣ���÷����Ӧ��Ӱ�񱨸�Ԫ���б�
  Procedure p_GetElementByClass(
    Val           Out t_Refcur,
    ����ID_In In Ӱ�񱨸�Ԫ���嵥.����ID%Type
	);

  --11.��  �ܣ����ID��Ӧ��Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_GetElementByID(
    Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�Ԫ���嵥.ID%Type
	);

  --12.�� �ܣ�����Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_AddElement(
    ID_In           In Ӱ�񱨸�Ԫ���嵥.ID%Type,
	����ID_In     In Ӱ�񱨸�Ԫ���嵥.����ID%Type,
	����_In         In Ӱ�񱨸�Ԫ���嵥.����%Type,
	����_In         In Ӱ�񱨸�Ԫ���嵥.����%Type,
	˵��_In         In Ӱ�񱨸�Ԫ���嵥.˵��%Type,
	ǰ׺_In         In Ӱ�񱨸�Ԫ���嵥.ǰ׺%Type,
	��׺_In         In Ӱ�񱨸�Ԫ���嵥.��׺%Type,
	��������_In   In Ӱ�񱨸�Ԫ���嵥.��������%Type,
	��ֵ��̬_In   In Ӱ�񱨸�Ԫ���嵥.��ֵ��̬%Type,
	��С����_In   In Ӱ�񱨸�Ԫ���嵥.��С����%Type,
	��󳤶�_In   In Ӱ�񱨸�Ԫ���嵥.��󳤶�%Type,
	��СС��λ_In In Ӱ�񱨸�Ԫ���嵥.��СС��λ%Type,
	���С��λ_In In Ӱ�񱨸�Ԫ���嵥.���С��λ%Type,
	������λ_In   In Ӱ�񱨸�Ԫ���嵥.������λ%Type,
	��չ����_In   In Varchar2,
	ֵ��ID_In      In Ӱ�񱨸�Ԫ���嵥.ֵ��ID%Type,
	ֵ������_In   In Ӱ�񱨸�Ԫ���嵥.ֵ������%Type
	);

  --13.�� �ܣ��޸�Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_EditElement(
    ID_In         In Ӱ�񱨸�Ԫ���嵥.ID%Type,
	����ID_In     In Ӱ�񱨸�Ԫ���嵥.����ID%Type,
	����_In       In Ӱ�񱨸�Ԫ���嵥.����%Type,
	����_In       In Ӱ�񱨸�Ԫ���嵥.����%Type,
	ǰ׺_In       In Ӱ�񱨸�Ԫ���嵥.ǰ׺%Type,
    ��׺_In       In Ӱ�񱨸�Ԫ���嵥.��׺%Type,
    ˵��_In       In Ӱ�񱨸�Ԫ���嵥.˵��%Type,
    ��������_In   In Ӱ�񱨸�Ԫ���嵥.��������%Type,
    ��ֵ��̬_In   In Ӱ�񱨸�Ԫ���嵥.��ֵ��̬%Type,
    ��С����_In   In Ӱ�񱨸�Ԫ���嵥.��С����%Type,
    ��󳤶�_In   In Ӱ�񱨸�Ԫ���嵥.��󳤶�%Type,
    ��СС��λ_In In Ӱ�񱨸�Ԫ���嵥.��СС��λ%Type,
    ���С��λ_In In Ӱ�񱨸�Ԫ���嵥.���С��λ%Type,
    ������λ_In   In Ӱ�񱨸�Ԫ���嵥.������λ%Type,
    ��չ����_In   In Varchar2,
    ֵ��ID_In     In Ӱ�񱨸�Ԫ���嵥.ֵ��ID%Type,
    ֵ������_In   In Ӱ�񱨸�Ԫ���嵥.ֵ������%Type
	);

  --14.�� �ܣ�ɾ��Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_DelElement(
    ID_In Ӱ�񱨸�Ԫ���嵥.ID%Type
	);

  --15.�� �ܣ�ͨ��ID��ȡӰ�񱨸������Ϣ
  Procedure p_GetElementClassByID(
    Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�Ԫ�ط���.ID%Type
	);
  --16.��  �ܣ���ȡԪ�ص���һ������
  Procedure p_Get_ElementNextCode(
    Val Out t_Refcur
	);
  --17.��  �ܣ���ȡԪ�ط������һ������
  Procedure p_Get_ElementClassNextCode(
    Val Out t_Refcur
	);
  --18.��  �ܣ���ȡ��Ӧ��ֵ���������ڵ�Ԫ�����
  Procedure p_Get_ElementClassByKind(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type
	);
  --19.��  �ܣ���ȡֵ�����Ͷ�Ӧ��ֵ����Ϣ
  Procedure p_Get_RangeByKind(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type
	);
  --20.��  �ܣ���ȡ��Ӧ��ֵ�����ͺ������������ڵ�Ԫ�����
  Procedure p_Get_ElementClassByKindType(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
	��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type
	);
  --21.��  �ܣ���ȡֵ�����ͺ��������Ͷ�Ӧ��ֵ����Ϣ
  Procedure p_Get_RangeByKindAndType(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
	��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type
	);
  --22.�� �ܣ���ȡ����޸�Ӱ�񱨸�Ԫ�ط�����Ϣ
  Procedure p_GetElementClassLastID(
    Val Out t_Refcur
	);
  --23.�� �ܣ���ȡ�༭�˶�Ӧ������޸�Ӱ�񱨸�Ԫ����ϢID
  Procedure p_GetElementLastID(
    Val Out t_Refcur
	);
  --24.�� �ܣ���ȡ����޸�Ӱ�񱨸�ֵ����ϢID
  Procedure p_GetRangeLastID(
    Val Out t_Refcur
	);
  --25.�� �ܣ���Ӽ�����λ��Ϣ
  Procedure p_AddMasure_Unit(
    ����_In Ӱ�񱨸������λ.����%Type,
    ����_In Ӱ�񱨸������λ.����%Type,
    ˵��_In Ӱ�񱨸������λ.˵��%Type,
    ǰ׺_In Ӱ�񱨸������λ.ǰ׺%Type
	);
  --26.��  �ܣ��޸ļ�����λ
  Procedure p_EditMasure_Unit(
    ԭ����_In Ӱ�񱨸������λ.����%Type,
    ����_In   Ӱ�񱨸������λ.����%Type,
    ����_In   Ӱ�񱨸������λ.����%Type,
    ˵��_In   Ӱ�񱨸������λ.˵��%Type,
    ǰ׺_In   Ӱ�񱨸������λ.ǰ׺%Type
	);
  --27.��  �ܣ�ɾ��������λ
  Procedure p_DelMasure_Unit(
    ����_In Ӱ�񱨸������λ.����%Type
	);
  --28.��  �ܣ��жϼ�����λ�ı����Ƿ��Ѵ���
  Procedure p_If_Exist_Masure_Unit(
    Val           Out t_Refcur,
	����_In Ӱ�񱨸������λ.����%Type
	);
  --29.��  �ܣ� �ж�Ԫ�ر����Ƿ��Ѵ���
  Procedure p_If_Exist_ElementCode(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�Ԫ���嵥.ID%Type,
	����_In Ӱ�񱨸�Ԫ���嵥.����%Type
	);
  --30.��  �ܣ� �ж�Ԫ�������Ƿ��Ѵ���
  Procedure p_If_Exist_ElementName(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�Ԫ���嵥.ID%Type,
	����_In Ӱ�񱨸�Ԫ���嵥.����%Type
	);
  --31.��  �ܣ� �ж�ֵ������Ƿ��Ѵ���
  Procedure p_If_Exist_RangeCode(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�ֵ���嵥.ID%Type,
	����_In Ӱ�񱨸�ֵ���嵥.����%Type
	);
  --32.��  �ܣ� �ж�ֵ������Ƿ��Ѵ���
  Procedure p_If_Exist_RangeName(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�ֵ���嵥.ID%Type,
	����_In Ӱ�񱨸�ֵ���嵥.����%Type
	);
  --33.���Ԫ���б�
  Procedure p_GetElementList(
    Val Out t_Refcur
	);
  --34.���ֵ���б�
  Procedure p_GetRangeList(
    Val Out t_Refcur
	);
  --35.�ж�Ԫ�ط���ı���ͱ����Ƿ����
  Procedure p_If_Exist_ElementClass(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�Ԫ�ط���.ID%Type,
	����_In Ӱ�񱨸�Ԫ�ط���.����%Type,
	����_In Ӱ�񱨸�Ԫ�ط���.����%Type
	) ;
    --36.�жϸ�Ԫ�ط��������Ƿ���ֵ�����Ԫ��
  Procedure p_Is_CanDel_ElementClass(
    Val           Out t_Refcur,
	ID_In Ӱ�񱨸�Ԫ�ط���.ID%Type
	);
End b_PACS_RptElement;
/

	--Ӱ�񱨸�Ԫ��ֵ��(---ʵ�ֲ���---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptElement Is

  --1.��   �ܣ���ȡȫ����Ӱ�񱨸�Ԫ�ط���
  Procedure p_GetElementClassList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(�ϼ�ID) �ϼ�ID,
             ����,
             ����,
             '[' || ���� || ']' || ���� As ����,
             ˵��,
             ���༭ʱ��
        From Ӱ�񱨸�Ԫ�ط���;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementClassList;

  --2.��  �ܣ�����Ӱ�񱨸�Ԫ�ط�����Ϣ
  Procedure p_AddElementClass(
    ID_In   In Ӱ�񱨸�Ԫ�ط���.ID%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	˵��_In In Ӱ�񱨸�Ԫ�ط���.˵��%Type,
	�ϼ�ID_In In Ӱ�񱨸�Ԫ�ط���.�ϼ�ID%Type
	) As
  Begin
    Insert Into Ӱ�񱨸�Ԫ�ط���
      (ID, ����, ����, ˵��, ���༭ʱ��,�ϼ�ID)
    Values
      (ID_In, ����_In, ����_In, ˵��_In, Sysdate,�ϼ�ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddElementClass;

  --3.��  �ܣ��޸�Ӱ�񱨸�Ԫ�ط�����Ϣ
  Procedure p_EditElementClass(
    ID_In   In Ӱ�񱨸�Ԫ�ط���.ID%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	����_In In Ӱ�񱨸�Ԫ�ط���.����%Type,
	˵��_In In Ӱ�񱨸�Ԫ�ط���.˵��%Type,
	�ϼ�ID_In In Ӱ�񱨸�Ԫ�ط���.�ϼ�ID%Type
	) As
  Begin
    Update Ӱ�񱨸�Ԫ�ط���
       Set ����         = ����_In,
           ����         = ����_In,
           ˵��         = ˵��_In,
           ���༭ʱ�� = Sysdate,
           �ϼ�ID=�ϼ�ID_In
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditElementClass;

  --4.��  �ܣ�ɾ��Ӱ�񱨸�Ԫ�ط�����Ϣ
  Procedure p_DelelEmentClass(
    ID_In In Ӱ�񱨸�Ԫ�ط���.ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸�Ԫ�ط��� Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelelEmentClass;

  --5.��  �ܣ���÷����Ӧ��Ӱ�񱨸�ֵ����Ϣ�б�
  Procedure p_GetRangeByClass(
    Val           Out t_Refcur,
	����ID_In In Ӱ�񱨸�ֵ���嵥.����ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(����ID) ����ID,
             ����,
             ����,
             ˵��,
             ��������,
             ֵ������,
             (Nvl(t.ֵ������, XmlType('<NULL/>'))).GetClobVal() As ֵ������,
             ���༭ʱ��
        From Ӱ�񱨸�ֵ���嵥 t
       Where ����ID = ����ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeByClass;

  --6.��  �ܣ����ID��Ӧ��Ӱ�񱨸�ֵ����Ϣ
  Procedure p_GetRangeByID(
    Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�ֵ���嵥.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(����ID) ����ID,
             ����,
             ����,
             '[' || ���� || ']' || ���� ����,
             ˵��,
             ��������,
             ֵ������,
             (Nvl(t.ֵ������, XmlType('<NULL/>'))).GetClobVal() As ֵ������,
             ���༭ʱ��
        From Ӱ�񱨸�ֵ���嵥 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeByID;

  --7.��  �ܣ�����Ӱ�񱨸�ֵ����Ϣ
  Procedure p_AddRange(
    ID_In       In Ӱ�񱨸�ֵ���嵥.ID%Type,
	����ID_In   In Ӱ�񱨸�ֵ���嵥.����ID%Type,
    ����_In     In Ӱ�񱨸�ֵ���嵥.����%Type,
    ����_In     In Ӱ�񱨸�ֵ���嵥.����%Type,
    ˵��_In     In Ӱ�񱨸�ֵ���嵥.˵��%Type,
    ��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type,
    ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
    ֵ������_In In Varchar2
	) As
  Begin
    Insert Into Ӱ�񱨸�ֵ���嵥
      (ID, ����ID, ����, ����, ˵��, ��������, ֵ������, ֵ������, ���༭ʱ��)
    Values
      (ID_In, ����ID_In, ����_In, ����_In, ˵��_In, ��������_In, ֵ������_In, ֵ������_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddRange;

  --8.��  �ܣ��޸�Ӱ�񱨸�ֵ����Ϣ
  Procedure p_EditRange(
    ID_In       In Ӱ�񱨸�ֵ���嵥.ID%Type,
    ����ID_In   In Ӱ�񱨸�ֵ���嵥.����ID%Type,
    ����_In     In Ӱ�񱨸�ֵ���嵥.����%Type,
    ����_In     In Ӱ�񱨸�ֵ���嵥.����%Type,
    ˵��_In     In Ӱ�񱨸�ֵ���嵥.˵��%Type,
    ��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type,
    ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
    ֵ������_In In Varchar2
	) As
  Begin
    Update Ӱ�񱨸�ֵ���嵥
       Set ����ID       = ����ID_In,
           ����         = ����_In,
           ����         = ����_In,
           ˵��         = ˵��_In,
           ��������     = ��������_In,
           ֵ������     = ֵ������_In,
           ֵ������     = ֵ������_In,
           ���༭ʱ�� = Sysdate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditRange;

  --9.��  �ܣ�ɾ��Ӱ�񱨸�ֵ����Ϣ
  Procedure p_DelRange(
    ID_In In Ӱ�񱨸�ֵ���嵥.ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸�ֵ���嵥 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelRange;

  --10.��  �ܣ���÷����Ӧ��Ӱ�񱨸�Ԫ���б�
  Procedure p_GetElementByClass(
    Val           Out t_Refcur,
	����ID_In In Ӱ�񱨸�Ԫ���嵥.����ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(����ID) ����ID,
             ����,
             ����,
             ǰ׺,
             ��׺,
             ˵��,
             ��������,
             ��ֵ��̬,
             ��С����,
             ��󳤶�,
             ��СС��λ,
             ���С��λ,
             ������λ,
             (Nvl(t.��չ����, XmlType('<NULL/>'))).GetClobVal() As ��չ����,
             RawToHex(ֵ��ID) ֵ��ID,
             ֵ������,
             ���༭ʱ��
        From Ӱ�񱨸�Ԫ���嵥 t
       Where ����ID = ����ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementByClass;

  --11.��  �ܣ����ID��Ӧ��Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_GetElementByID(
    Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�Ԫ���嵥.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(����ID) ����ID,
             ����,
             ����,
             ǰ׺,
             ��׺,
             ˵��,
             ��������,
             ��ֵ��̬,
             ��С����,
             ��󳤶�,
             ��СС��λ,
             ���С��λ,
             ������λ,
             (Nvl(t.��չ����, XmlType('<NULL/>'))).GetClobVal() As ��չ����,
             RawToHex(ֵ��ID) ֵ��ID,
             ֵ������,
             ���༭ʱ��
        From Ӱ�񱨸�Ԫ���嵥 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementByID;

  --12.�� �ܣ�����Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_AddElement(
    ID_In            In Ӱ�񱨸�Ԫ���嵥.ID%Type,
    ����ID_In     In Ӱ�񱨸�Ԫ���嵥.����ID%Type,
    ����_In         In Ӱ�񱨸�Ԫ���嵥.����%Type,
    ����_In         In Ӱ�񱨸�Ԫ���嵥.����%Type,
    ˵��_In         In Ӱ�񱨸�Ԫ���嵥.˵��%Type,
    ǰ׺_In         In Ӱ�񱨸�Ԫ���嵥.ǰ׺%Type,
    ��׺_In         In Ӱ�񱨸�Ԫ���嵥.��׺%Type,
    ��������_In   In Ӱ�񱨸�Ԫ���嵥.��������%Type,
    ��ֵ��̬_In   In Ӱ�񱨸�Ԫ���嵥.��ֵ��̬%Type,
    ��С����_In   In Ӱ�񱨸�Ԫ���嵥.��С����%Type,
    ��󳤶�_In   In Ӱ�񱨸�Ԫ���嵥.��󳤶�%Type,
    ��СС��λ_In In Ӱ�񱨸�Ԫ���嵥.��СС��λ%Type,
    ���С��λ_In In Ӱ�񱨸�Ԫ���嵥.���С��λ%Type,
    ������λ_In   In Ӱ�񱨸�Ԫ���嵥.������λ%Type,
    ��չ����_In   In Varchar2,
    ֵ��ID_In      In Ӱ�񱨸�Ԫ���嵥.ֵ��ID%Type,
    ֵ������_In   In Ӱ�񱨸�Ԫ���嵥.ֵ������%Type
	) As
  Begin
    Insert Into Ӱ�񱨸�Ԫ���嵥
      (ID, ����ID, ����, ����,  ǰ׺, ��׺, ˵��,  ��������,  ��ֵ��̬,  ��С����,  ��󳤶�,
       ��СС��λ,  ���С��λ,  ������λ,   ��չ����, ֵ��ID,  ֵ������, ���༭ʱ��)
    Values
      (ID_In, ����ID_In, ����_In, ����_In, ǰ׺_In, ��׺_In, ˵��_In, ��������_In, ��ֵ��̬_In, ��С����_In, ��󳤶�_In,
       ��СС��λ_In, ���С��λ_In, ������λ_In, ��չ����_In, ֵ��ID_In, ֵ������_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddElement;

  --13.�� �ܣ��޸�Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_EditElement(
    ID_In           In Ӱ�񱨸�Ԫ���嵥.ID%Type,
    ����ID_In     In Ӱ�񱨸�Ԫ���嵥.����ID%Type,
    ����_In        In Ӱ�񱨸�Ԫ���嵥.����%Type,
    ����_In        In Ӱ�񱨸�Ԫ���嵥.����%Type,
    ǰ׺_In        In Ӱ�񱨸�Ԫ���嵥.ǰ׺%Type,
    ��׺_In        In Ӱ�񱨸�Ԫ���嵥.��׺%Type,
    ˵��_In        In Ӱ�񱨸�Ԫ���嵥.˵��%Type,
    ��������_In   In Ӱ�񱨸�Ԫ���嵥.��������%Type,
    ��ֵ��̬_In   In Ӱ�񱨸�Ԫ���嵥.��ֵ��̬%Type,
    ��С����_In   In Ӱ�񱨸�Ԫ���嵥.��С����%Type,
    ��󳤶�_In   In Ӱ�񱨸�Ԫ���嵥.��󳤶�%Type,
    ��СС��λ_In In Ӱ�񱨸�Ԫ���嵥.��СС��λ%Type,
    ���С��λ_In In Ӱ�񱨸�Ԫ���嵥.���С��λ%Type,
    ������λ_In    In Ӱ�񱨸�Ԫ���嵥.������λ%Type,
    ��չ����_In    In Varchar2,
    ֵ��ID_In      In Ӱ�񱨸�Ԫ���嵥.ֵ��ID%Type,
    ֵ������_In   In Ӱ�񱨸�Ԫ���嵥.ֵ������%Type
	) As
  Begin

    Update Ӱ�񱨸�Ԫ���嵥
       Set ����ID       = ����ID_In,
           ����         = ����_In,
           ����         = ����_In,
           ǰ׺         = ǰ׺_In,
           ��׺         = ��׺_In,
           ˵��         = ˵��_In,
           ��������     = ��������_In,
           ��ֵ��̬     = ��ֵ��̬_In,
           ��С����     = ��С����_In,
           ��󳤶�     = ��󳤶�_In,
           ��СС��λ   = ��СС��λ_In,
           ���С��λ   = ���С��λ_In,
           ������λ     = ������λ_In,
           ��չ����     = ��չ����_In,
           ֵ��ID       = ֵ��ID_In,
           ֵ������     = ֵ������_In,
           ���༭ʱ�� = Sysdate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditElement;

  --14.�� �ܣ�ɾ��Ӱ�񱨸�Ԫ����Ϣ
  Procedure p_DelElement(
    ID_In Ӱ�񱨸�Ԫ���嵥.ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸�Ԫ���嵥 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelElement;

  --15.�� �ܣ�ͨ��ID��ȡӰ�񱨸������Ϣ
  Procedure p_GetElementClassByID(
    Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�Ԫ�ط���.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID, ����, ����, ˵��, ���༭ʱ��,RawToHex(�ϼ�id) �ϼ�ID
        From Ӱ�񱨸�Ԫ�ط��� t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementClassByID;

  --16.��  �ܣ���ȡԪ�ص���һ������
  Procedure p_Get_ElementNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('Ӱ�񱨸�Ԫ���嵥') As ����
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementNextCode;

  --17.��  �ܣ���ȡԪ�ط������һ������
  Procedure p_Get_ElementClassNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('Ӱ�񱨸�Ԫ�ط���') As ����
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementClassNextCode;

  --18.��  �ܣ���ȡ��Ӧ��ֵ���������ڵ�Ԫ�����
  Procedure p_Get_ElementClassByKind(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type
	) As
  Begin
    Open Val For
      Select Distinct ��������, ����ID
        From (Select RawToHex(����ID) ����ID,
                     (Select a.����
                        From Ӱ�񱨸�Ԫ�ط��� A
                       Where a.Id = t.����id) As ��������
                From Ӱ�񱨸�ֵ���嵥 T
               Where t.ֵ������ = ֵ������_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementClassByKind;

  --19.��  �ܣ���ȡֵ�����Ͷ�Ӧ��ֵ����Ϣ
  Procedure p_Get_RangeByKind(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             ����,
             RawToHex(����ID) ����ID,
             ����,
             ��������,
             (Select a.���� From Ӱ�񱨸�Ԫ�ط��� A Where a.Id = t.����id) As ��������,
             ֵ������
        From Ӱ�񱨸�ֵ���嵥 T
       Where t.ֵ������ = ֵ������_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_RangeByKind;

  --20.��  �ܣ���ȡ��Ӧ��ֵ�����ͺ������������ڵ�Ԫ�����
  Procedure p_Get_ElementClassByKindType(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
	��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type
	) As
  Begin
    Open Val For
      Select Distinct ��������, ����ID
        From (Select RawToHex(����id) ����ID,
                     (Select a.����
                        From Ӱ�񱨸�Ԫ�ط��� A
                       Where a.Id = t.����id) As ��������
                From Ӱ�񱨸�ֵ���嵥 T
               Where t.ֵ������ = ֵ������_In
                 and t.�������� = ��������_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ElementClassByKindType;

  --21.��  �ܣ���ȡֵ�����ͺ��������Ͷ�Ӧ��ֵ����Ϣ
  Procedure p_Get_RangeByKindAndType(
    Val           Out t_Refcur,
	ֵ������_In In Ӱ�񱨸�ֵ���嵥.ֵ������%Type,
	��������_In In Ӱ�񱨸�ֵ���嵥.��������%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             ����,
             RawToHex(����ID) ����ID,
             ����,
             ��������,
             (Select a.���� From Ӱ�񱨸�Ԫ�ط��� A Where a.Id = t.����id) As ��������,
             ֵ������
        From Ӱ�񱨸�ֵ���嵥 T
       Where t.ֵ������ = ֵ������_In
         and t.�������� = ��������_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_RangeByKindAndType;

  --22.�� �ܣ���ȡ�༭�˶�Ӧ������޸�Ӱ�񱨸�ԭ�ͷ�����ϢID
  Procedure p_GetElementClassLastID(
    Val Out t_Refcur
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, ���༭ʱ��
        From Ӱ�񱨸�Ԫ�ط��� t1
       Where Not Exists (Select 1
                From Ӱ�񱨸�Ԫ�ط���
               Where ���༭ʱ�� > t1.���༭ʱ��);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementClassLastID;

  --23.�� �ܣ���ȡ�༭�˶�Ӧ������޸�Ӱ�񱨸�Ԫ����ϢID
  Procedure p_GetElementLastID(
    Val Out t_Refcur
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, ���༭ʱ��
        From Ӱ�񱨸�Ԫ���嵥 t1
       Where Not Exists (Select 1
                From Ӱ�񱨸�Ԫ���嵥
               Where ���༭ʱ�� > t1.���༭ʱ��);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementLastID;

  --24.�� �ܣ���ȡ����޸�Ӱ�񱨸�ֵ����ϢID
  Procedure p_GetRangeLastID(
    Val Out t_Refcur
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, ���༭ʱ��
        From Ӱ�񱨸�ֵ���嵥 t1
       Where Not Exists (Select 1
                From Ӱ�񱨸�ֵ���嵥
               Where ���༭ʱ�� > t1.���༭ʱ��);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeLastID;

  --25.�� �ܣ���Ӽ�����λ��Ϣ
  Procedure p_AddMasure_Unit(
    ����_In Ӱ�񱨸������λ.����%Type,
    ����_In Ӱ�񱨸������λ.����%Type,
    ˵��_In Ӱ�񱨸������λ.˵��%Type,
    ǰ׺_In Ӱ�񱨸������λ.ǰ׺%Type
	) As
  Begin
    Insert Into Ӱ�񱨸������λ
      (����, ����, ˵��, ǰ׺)
    Values
      (����_In, ����_In, ˵��_In, ǰ׺_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddMasure_Unit;

  --26.��  �ܣ��޸ļ�����λ
  Procedure p_EditMasure_Unit(
    ԭ����_In Ӱ�񱨸������λ.����%Type,
    ����_In   Ӱ�񱨸������λ.����%Type,
    ����_In   Ӱ�񱨸������λ.����%Type,
    ˵��_In   Ӱ�񱨸������λ.˵��%Type,
    ǰ׺_In   Ӱ�񱨸������λ.ǰ׺%Type
	) As
  Begin
    Update Ӱ�񱨸������λ
       Set ���� = ����_In, ���� = ����_In, ˵�� = ˵��_In, ǰ׺ = ǰ׺_In
     Where ���� = ԭ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditMasure_Unit;

  --27.��  �ܣ�ɾ��������λ
  Procedure p_DelMasure_Unit(
    ����_In Ӱ�񱨸������λ.����%Type
	) As
  Begin
    Delete from Ӱ�񱨸������λ Where ���� = ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelMasure_Unit;

  --28.��  �ܣ��жϼ�����λ�ı����Ƿ��Ѵ���
  Procedure p_If_Exist_Masure_Unit(
    Val           Out t_Refcur,
	����_In Ӱ�񱨸������λ.����%Type
	) As
  Begin
    Open Val For
      Select Count(t.����) ����
        From Ӱ�񱨸������λ t
       Where t.���� = ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_Masure_Unit;

  --29.��  �ܣ� �ж�Ԫ�ر����Ƿ��Ѵ���
  Procedure p_If_Exist_ElementCode(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�Ԫ���嵥.ID%Type,
	����_In Ӱ�񱨸�Ԫ���嵥.����%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) ����
        From Ӱ�񱨸�Ԫ���嵥 t
       Where t.���� = ����_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_ElementCode;

  --30.��  �ܣ� �ж�Ԫ�������Ƿ��Ѵ���
  Procedure p_If_Exist_ElementName(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�Ԫ���嵥.ID%Type,
	����_In Ӱ�񱨸�Ԫ���嵥.����%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) ����
        From Ӱ�񱨸�Ԫ���嵥 t
       Where t.���� = ����_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_ElementName;

  --31.��  �ܣ� �ж�ֵ������Ƿ��Ѵ���
  Procedure p_If_Exist_RangeCode(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�ֵ���嵥.ID%Type,
	����_In Ӱ�񱨸�ֵ���嵥.����%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) ����
        From Ӱ�񱨸�ֵ���嵥 t
       Where t.���� = ����_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_RangeCode;

  --32.��  �ܣ� �ж�ֵ�������Ƿ��Ѵ���
  Procedure p_If_Exist_RangeName(
    Val           Out t_Refcur,
	ID_In   Ӱ�񱨸�ֵ���嵥.ID%Type,
	����_In Ӱ�񱨸�ֵ���嵥.����%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) ����
        From Ӱ�񱨸�ֵ���嵥 t
       Where t.���� = ����_In
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_RangeName;

  --33.���Ԫ���б�
  Procedure p_GetElementList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(t.ID) ID,
             RawToHex(t.����ID) ����ID,
             t.����,
             t.����,
             t.ǰ׺,
             t.��׺,
             t.˵��,
             t.��������,
             t.��ֵ��̬,
             t.��С����,
             t.��󳤶�,
             t.��СС��λ,
             t.���С��λ,
             t.������λ,
             (Nvl(t.��չ����, XmlType('<NULL/>'))).GetClobVal() As ��չ����,
             RawToHex(t.ֵ��ID) ֵ��ID,
             (select a.���� from Ӱ�񱨸�ֵ���嵥 a Where a.id=t.ֵ��ID)as ֵ������,
             t.ֵ������,
             t.���༭ʱ��
        From Ӱ�񱨸�Ԫ���嵥 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetElementList;

  --34.���ֵ���б�
  Procedure p_GetRangeList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             RawToHex(����ID) ����ID,
             ����,
             ����,
             ˵��,
             ��������,
             ֵ������,
             (Nvl(t.ֵ������, XmlType('<NULL/>'))).GetClobVal() As ֵ������,
             ���༭ʱ��
        From Ӱ�񱨸�ֵ���嵥 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetRangeList;

  --35.�ж�Ԫ�ط���ı���ͱ����Ƿ����
  Procedure p_If_Exist_ElementClass(
    Val           Out t_Refcur,
    ID_In   Ӱ�񱨸�Ԫ�ط���.ID%Type,
	����_In Ӱ�񱨸�Ԫ�ط���.����%Type,
	����_In Ӱ�񱨸�Ԫ�ط���.����%Type
	) As
  Begin
    Open Val For
      Select Count(t.ID) ����
        From Ӱ�񱨸�Ԫ�ط��� t
       Where (t.���� = ����_In or t.���� = ����_In)
         and t.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_If_Exist_ElementClass;

  --36.�жϸ�Ԫ�ط��������Ƿ���ֵ�����Ԫ��
  Procedure p_Is_CanDel_ElementClass(
    Val           Out t_Refcur,
	ID_In Ӱ�񱨸�Ԫ�ط���.ID%Type
	) As
    ElementCount int;
    RangeCount   int;
    ElementClassCout int;
  Begin
    Select Count(*)
      into ElementCount
      From Ӱ�񱨸�Ԫ���嵥 a
     Where a.����id = ID_In;
    Select Count(*)
      into RangeCount
      From Ӱ�񱨸�ֵ���嵥 b
     Where b.����id = ID_In;
     Select Count(*)
      into ElementClassCout
      From Ӱ�񱨸�Ԫ�ط��� b
     Where b.�ϼ�id = ID_In;
    ElementCount := ElementCount + RangeCount+ElementClassCout;
    Open Val For
      Select ElementCount Count From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Is_CanDel_ElementClass;

End b_PACS_RptElement;
/




--Ӱ�񱨸�������(---���岿��---)***************************************************
--Ӱ�񱨸�������(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  --1.��  �ܣ����Ӱ�񱨸�����б�
  Procedure p_GetComboList(
    Val Out t_Refcur
	);
  --2.��  �ܣ����Ӱ�񱨸������Ϣ
  Procedure p_AddComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	);
  --3.��  ��;�޸�Ӱ�񱨸������Ϣ
  Procedure p_EditComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	);
  --4.��  �ܣ�ͨ��IDɾ��Ӱ�񱨸������Ϣ
  Procedure p_DelComboInfo(
    ID_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --5.��  �ܣ�����ID���Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --6.��  �ܣ����Ӱ�񱨸��������з�����Ϣ
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	);
  --7.��  �ܣ����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --8.��  �ܣ�����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_EditComboContent(
	ID_In   In Ӱ�񱨸�����嵥.ID%Type,
	���_In  In Ӱ�񱨸�����嵥.���%Type
	);
  --9.�� �ܣ���ȡ�༭�˶�Ӧ������޸�Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	�༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	);
  --10.��  �ܣ�����Ƭ�ε���Ͼ�
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
    Id_In   In Ӱ�񱨸�����嵥.ID%Type
	);

  --11.��  �ܣ��޸�Ƭ�ε���Ͼ�
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In Ӱ�񱨸�����嵥.ID%Type,
    Pid_In  In Varchar2
	);
  --12.��  �ܣ����ݷ���ID��ѯ�ʾ�
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --13.��  �ܣ���ȡ��һ������
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	);
end b_PACS_RptCombo;
/

--Ӱ�񱨸�������(---ʵ�ֲ���---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25

  --1.��  �ܣ����Ӱ�񱨸�����б�
  Procedure p_GetComboList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             ����,
             ����,
             ˵��,
             ����,
             ����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             �༭��,
             ���༭ʱ��
        From Ӱ�񱨸�����嵥 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboList;

  --2.��  �ܣ����Ӱ�񱨸������Ϣ
  Procedure p_AddComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	) As
  Begin
    Insert Into Ӱ�񱨸�����嵥
      (ID, ����, ����, ˵��, ����, ����, ���, �༭��, ���༭ʱ��)
    Values
      (ID_In, ����_In, ����_In, ˵��_In, ����_In, ����_In, ���_In, �༭��_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddComboInfo;

  --3.��  ��;�޸�Ӱ�񱨸������Ϣ
  Procedure p_EditComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	) As
  Begin
    Update Ӱ�񱨸�����嵥
       set ����         = ����_In,
           ����         = ����_In,
           ˵��         = ˵��_In,
           ����         = ����_In,
           ����         = ����_In,
           ���         = ���_In,
           �༭��       = �༭��_In,
           ���༭ʱ�� = SysDate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboInfo;

  --4.��  �ܣ�ͨ��IDɾ��Ӱ�񱨸������Ϣ
  Procedure p_DelComboInfo(
    ID_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸�����嵥 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelComboInfo;

  --5.��  �ܣ�����ID���Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             ����,
             ����,
             ˵��,
             ����,
             ����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             �༭��,
             ���༭ʱ��
        From Ӱ�񱨸�����嵥 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByID;

  --6.��  �ܣ����Ӱ�񱨸��������з�����Ϣ
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Distinct ���� From Ӱ�񱨸�����嵥;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboAllGroup;

  --7.��  �ܣ����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Open Val For
      Select (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���
        From Ӱ�񱨸�����嵥 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboContent;

  --8.��  �ܣ�����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_EditComboContent(
    ID_In   In Ӱ�񱨸�����嵥.ID%Type,
    ���_In In Ӱ�񱨸�����嵥.���%Type
	) As
  Begin
    Update Ӱ�񱨸�����嵥 Set ��� = ���_In Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboContent;

  --9.�� �ܣ���ȡ�༭�˶�Ӧ������޸�Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	�༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, �༭��, ���༭ʱ��
        From Ӱ�񱨸�����嵥 t1
       Where Not Exists (Select 1
                From Ӱ�񱨸�����嵥
               Where ���༭ʱ�� > t1.���༭ʱ��)
         And �༭�� = �༭��_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByEditor;

  --10.��  �ܣ�����Ƭ�ε���Ͼ�
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
	Id_In   In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Update Ӱ�񱨸�����嵥 A
       Set a.��� = Appendchildxml(a.���, '/root', Text_In)
     Where a.ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Append_Fragment_Tocombo;

  --11.��  �ܣ��޸�Ƭ�ε���Ͼ�
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In Ӱ�񱨸�����嵥.ID%Type,
    Pid_In  In Varchar2
	) As
  Begin
    Update Ӱ�񱨸�����嵥 A
       Set a.��� = Updatexml(a.���,
                            '/root/sentence[@sid="' || Pid_In || '"]',
                            Text_In)
     Where a.ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Combo_Fragment;

  --12.��  �ܣ����ݷ���ID��ѯ�ʾ�
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             �ϼ�id,
             ����,
             ����,
             ˵��,
             �ڵ�����,
             (Nvl(a.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             ѧ��,
             ��ǩ,
             �Ƿ�˽��,
             ����,
			 (Nvl(a.��Ӧ����, XmlType('<NULL/>'))).GetClobVal() As ��Ӧ����, 
             ���༭ʱ��
        From Ӱ�񱨸�Ƭ���嵥 A
       Where a.�ϼ�id = Id_In
         And a.�ڵ����� <> 0;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Fragment_By_Typeid;

  --13.��  �ܣ���ȡ��һ������
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('Ӱ�񱨸�����嵥') As ����
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ComboNextCode;
End b_PACS_RptCombo;
/



CREATE OR REPLACE Package b_PACS_RptFragments Is
  Type t_Refcur Is Ref Cursor;


  --���ܣ���ȡ����Ԥ�����
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	);
  --���ܣ���ȡ���ж������
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	);
  --���ܣ���ȡ��ǰ�û�ѧ�����ж���������ڵ�
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type
	) ;


  --���ܣ����ݷ���ID���Ҷ���
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	) ;


   Procedure p_Get_Label_By_Typeid(
     Val           Out t_Refcur,
	 Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	 ) ;

  --���ܣ������������
  Procedure p_Add_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

  --���ܣ��޸Ķ������
  Procedure p_Edit_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

  --���ܣ�ɾ���������
   Procedure p_Del_Fragmenttype(
     Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	 );

    --���ܣ���Ӷ���
  Procedure p_Add_Fragment(
     Id_In      Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

   --���ܣ��޸Ķ���
  Procedure p_Edit_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    );
   --���ܣ�ɾ������
  Procedure p_Del_Fragment(
    Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	);

  procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	);

  --���ܣ��������
  Procedure p_Import_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

procedure p_Get_Data_Last_Edit_Time(
  Val           Out t_Refcur,
  Table_Name_In varchar2
  );

   --���ܣ��ж�Ƭ�η����ܷ�ɾ��
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	);

  --���ܣ�����Ƭ��ID�����õ�ǰƬ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  );
  
  --���ܣ�����Ƭ�εĸ�ID����������Ŀ¼����Ŀ¼Ƭ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionByPid
  (
    �ϼ�ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In    In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  );

  --���ܣ���ȡ��ǰ����Ƭ����Ӧ����
  Procedure p_Get_FraConditionByOrderId
  (
    Val           Out t_Refcur,
	ҽ��ID_In    Ӱ�����¼.ҽ��ID%Type
  );

  --���ܣ���ȡӰ�������
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  );
  
  --���ܣ���������ȡ���Ƽ�鲿λ
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --���ܣ���������ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --���ܣ��������Ʊ����ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  );

  --�ж��Ƿ�����ͬ�Ĵ���
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  Code_In  Ӱ�񱨸�Ƭ���嵥.����%Type
  );

  --�ж��Ƿ�����ͬ������
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  PID_In    In Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
  Name_In  In Ӱ�񱨸�Ƭ���嵥.����%Type,
  Author_In In  Ӱ�񱨸�Ƭ���嵥.����%Type
  );

  End  b_PACS_RptFragments;
/
CREATE OR REPLACE Package Body b_PACS_RptFragments Is

  ------------------------------------------------------------------------
  --Ƭ��ģ��
  ------------------------------------------------------------------------

  --���ܣ���ȡ����Ԥ�����
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(ID) ID, ����, ���� From Ӱ�񱨸�Ԥ����� Order By ����;
  End p_Get_All_Phr_Onlines;

  --���ܣ���ȡ���ж������
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, Rawtohex(a.�ϼ�id) As �ϼ�id, a.����, a.����, a.˵��, a.�ڵ�����
      From Ӱ�񱨸�Ƭ���嵥 A
      Where a.�ڵ����� = 0
      Start With �ϼ�id Is Null
      Connect By Prior ID = �ϼ�id
      Order By �ϼ�ID Desc, ����;
  End p_Get_All_Fragment_Class;

  --���ܣ���ȡ��ǰ�û�ѧ�����ж���������ڵ�
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type
	) As
  Begin
    If Subjects_In <> '' Then
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.�ϼ�id) As �ϼ�id, a.����, a.����, a.˵��, a.�ڵ�����, (Nvl(a.���, XmlType('<NULL/>'))).GetClobVal() As ���, 
			a.ѧ��, a.��ǩ, a.�Ƿ�˽��, a.����, (Nvl(a.��Ӧ����, XmlType('<NULL/>'))).GetClobVal() As ��Ӧ����,a.���༭ʱ��, a.�ڵ����� As Image
        From Ӱ�񱨸�Ƭ���嵥 A
        Where (a.ѧ�� In (Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(Subjects_In, ','))
                        Intersect
                        Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(a.ѧ��, ','))) And a.�ڵ����� <> 0) Or a.�ڵ����� = 0 Or a.ѧ�� Is Null
        Order By ����, �ϼ�id;
    Else
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.�ϼ�id) As �ϼ�id, a.����, a.����, a.˵��, a.�ڵ�����, (Nvl(a.���, XmlType('<NULL/>'))).GetClobVal() As ���, 
			a.ѧ��, a.��ǩ, a.�Ƿ�˽��, a.����, (Nvl(a.��Ӧ����, XmlType('<NULL/>'))).GetClobVal() As ��Ӧ����,a.���༭ʱ��, a.�ڵ����� As Image
        From Ӱ�񱨸�Ƭ���嵥 A
        Order By �ϼ�id, �ڵ�����, ����, ����;
    End If;
  End p_Get_All_Fragment;

  --���ܣ����ݷ���ID���Ҷ���
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
  ) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, a.�ϼ�ID,a.����, a.����, a.˵��, a.�ڵ�����, (Nvl(a.���, XmlType('<NULL/>'))).GetClobVal() As ���, 
				a.ѧ��, a.��ǩ, a.�Ƿ�˽��, a.����, (Nvl(a.��Ӧ����, XmlType('<NULL/>'))).GetClobVal() As ��Ӧ����, a.���༭ʱ��,a.�ڵ����� As Image
      From Ӱ�񱨸�Ƭ���嵥 A
      Where a.�ϼ�id = Hextoraw(Id_In) And a.�ڵ����� <> 0;
  End p_Get_Fragment_By_Typeid;

  --���ܣ�����ĳ���������ж����ǩ
  Procedure p_Get_Label_By_Typeid(
    Val           Out t_Refcur,
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
    ) As
  Begin
    Open Val For
      Select Distinct ��ǩ From Ӱ�񱨸�Ƭ���嵥 Where �ϼ�id = Hextoraw(Id_In) And ��ǩ Is Not Null;
  End p_Get_Label_By_Typeid;

  --���ܣ������������
  Procedure p_Add_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where ���� = Code_In Or ���� = Title_In And �ڵ����� = 0 And �ϼ�id = Hextoraw(Pid_In);

    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]�������ƻ�����Ѿ����ڣ�[ZLSOFT]';
      Raise Err_Item;
    Else
      Insert Into Ӱ�񱨸�Ƭ���嵥
        (ID, �ϼ�id, ����, ����, ˵��, �ڵ�����, ����, ���༭ʱ��)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Author_In, Sysdate);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragmenttype;

  --���ܣ��޸Ķ������
  Procedure p_Edit_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where (���� = Code_In Or ���� = Title_In) And �ڵ����� = 0 And �ϼ�id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]�������ƻ�����Ѿ����ڣ�[ZLSOFT]';
      Raise Err_Item;
    Else
      Update Ӱ�񱨸�Ƭ���嵥
      Set �ϼ�id = Hextoraw(Pid_In), ���� = Code_In, ���� = Title_In, ˵�� = Note_In, �ڵ����� = Leaf_In, ���� = Author_In,
          ���༭ʱ�� = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragmenttype;

  --���ܣ�ɾ���������
  Procedure p_Del_Fragmenttype(
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where �ڵ����� <> 0 And
          ID In (Select ID From Ӱ�񱨸�Ƭ���嵥 Connect By Prior ID = �ϼ�id Start With ID = Hextoraw(Id_In));

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]�÷����´��ڶ���ݲ���ɾ����[ZLSOFT]';
      Raise Err_Item;
    Else
      Delete Ӱ�񱨸�Ƭ���嵥 Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragmenttype;

  --���ܣ���Ӷ���
  Procedure p_Add_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
  Begin

      Insert Into Ӱ�񱨸�Ƭ���嵥
        (ID, �ϼ�id, ����, ����, ˵��, �ڵ�����, ���, ѧ��, ��ǩ, �Ƿ�˽��, ����, ���༭ʱ��)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragment;

  --���ܣ��޸Ķ���
  Procedure p_Edit_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where (���� = Code_In Or ���� = Title_In) And �ڵ����� <> 0 And �ϼ�id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]��������ƻ�����Ѿ����ڣ�[ZLSOFT]';
      Raise Err_Item;
    Else
      Update Ӱ�񱨸�Ƭ���嵥
      Set �ϼ�id = Hextoraw(Pid_In), ���� = Code_In, ���� = Title_In, ˵�� = Note_In, �ڵ����� = Leaf_In, ��� = Content_In,
          ѧ�� = Subjects_In, ��ǩ = Label_In, �Ƿ�˽�� = Private_In, ���� = Author_In, ���༭ʱ�� = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragment;

  --
  Procedure p_Get_All_Fragment_List(Val Out t_Refcur) As
  Begin
    Open Val For
      Select Rawtohex(t.Id) As ID, Rawtohex(t.�ϼ�id) As �ϼ�id, t.����, t.����, t.˵��, t.�ڵ�����, (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���, t.ѧ��, t.��ǩ, t.�Ƿ�˽��, t.����,
             t.���༭ʱ��
      From Ӱ�񱨸�Ƭ���嵥 T;
  End p_Get_All_Fragment_List;

  --���ܣ�ɾ������
  Procedure p_Del_Fragment(
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	) As
  Begin
    Delete Ӱ�񱨸�Ƭ���嵥 Where ID = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragment;

  --���ܣ��������
  Procedure p_Import_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    v_Num Number(2);
  Begin
    Select Count(ID)
    Into v_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where ((���� = Code_In Or ���� = Title_In) And �ϼ�id = Hextoraw(Pid_In)) Or
          (�ϼ�id Is Null And (���� = Code_In Or ���� = Title_In));

    If v_Num > 0 Then
      Update Ӱ�񱨸�Ƭ���嵥
      Set ��� = Content_In, ���༭ʱ�� = Sysdate, �Ƿ�˽�� = 0
      Where ((���� = Code_In Or ���� = Title_In) And �ϼ�id = Hextoraw(Pid_In)) Or
            (�ϼ�id Is Null And (���� = Code_In Or ���� = Title_In));
    Else
      Insert Into Ӱ�񱨸�Ƭ���嵥
        (ID, �ϼ�id, ����, ����, ˵��, �ڵ�����, ���, ѧ��, ��ǩ, �Ƿ�˽��, ����, ���༭ʱ��)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);
    End If;

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Import_Fragment;

  --
  Procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
    Table_Name_In Varchar2
    ) As
    v_Sql Varchar2(4000);
  Begin
    v_Sql := 'select max(���༭ʱ��) maxvalue from ' || Table_Name_In;
    Open Val For v_Sql;
  End p_Get_Data_Last_Edit_Time;
  
   --���ܣ��ж�Ƭ�η����ܷ�ɾ��
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select Count(t.id) Count
        From Ӱ�񱨸�Ƭ���嵥 t
       Where �ϼ�id = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_IsCanDel_FragmentType;
  
  --���ܣ�����Ƭ��ID�����õ�ǰƬ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  )As
  Begin
    Update Ӱ�񱨸�Ƭ���嵥 Set ��Ӧ���� = ��Ӧ����_In Where ID = Hextoraw(ID_In) And �ڵ����� != 0;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionById;
  
  --���ܣ�����Ƭ�εĸ�ID����������Ŀ¼����Ŀ¼Ƭ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionByPid
  (
    �ϼ�ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  )As
  Begin
    Update Ӱ�񱨸�Ƭ���嵥 Set ��Ӧ���� = ��Ӧ����_In Where �ϼ�ID = Hextoraw(�ϼ�ID_In) And �ڵ����� != 0;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionByPid;

  --���ܣ���ȡ��ǰ����Ƭ����Ӧ����
  Procedure p_Get_FraConditionByOrderId(
    Val           Out t_Refcur,
	  ҽ��ID_In    Ӱ�����¼.ҽ��ID%Type
	) As
  Begin
    Open Val For
	  Select a.id, a.�Ա�,c.Ӱ�����, d.����||' - '||d.���� ������, c.Ӱ�����||' - '||e.����||' - '||e.���� �����Ŀ, A.ҽ������
      From ����ҽ����¼ a, ����ҽ������ b, Ӱ�����¼ c, Ӱ������� d, ������ĿĿ¼ e
      Where a.id = b.ҽ��id and b.ҽ��id=c.ҽ��id and c.Ӱ����� = d.���� and a.������Ŀid = e.id and a.id = ҽ��ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_FraConditionByOrderId;

  --���ܣ���ȡӰ�������
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  ) As
  Begin
    Open Val For
      Select ����||' - '||���� ������ From Ӱ�������;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckLueKind;
  
  --���ܣ���������ȡ���Ƽ�鲿λ
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select Distinct ����||���� IID, '' �ϼ�ID, ����||' - '||���� ���Ʋ�λ From ���Ƽ�鲿λ a,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) b Where a.���� = b.Column_Value
      Union Select ����||����||���� IID, ����||���� �ϼ�ID, ����||' - '||���� ���Ʋ�λ From ���Ƽ�鲿λ c,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) d Where c.���� = d.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckPartList;
  
  --���ܣ���������ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.����, r.Ӱ�����||' - '||I.����||' - '||I.���� �����Ŀ
      From ������ĿĿ¼ I, Ӱ������Ŀ R, Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.������Ŀid And R.Ӱ�����=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByKind;
  
  --���ܣ��������Ʊ����ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.����, r.Ӱ�����||' - '||I.����||' - '||I.���� �����Ŀ
      From ������ĿĿ¼ I, Ӱ������Ŀ R, Table(Cast(f_Str2list(''||Code_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.������Ŀid And I.����=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByCode;

  --�ж��Ƿ�����ͬ�Ĵ���
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  Code_In  Ӱ�񱨸�Ƭ���嵥.����%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From Ӱ�񱨸�Ƭ���嵥 Where ID<>ID_In And ����=Code_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameCode;

  --�ж��Ƿ�����ͬ������
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  PID_In    In Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
  Name_In  In Ӱ�񱨸�Ƭ���嵥.����%Type,
  Author_In In  Ӱ�񱨸�Ƭ���嵥.����%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From Ӱ�񱨸�Ƭ���嵥 Where �ϼ�ID=PID_In And ����=Author_In And ID<>ID_In And ����=Name_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameName;

End  b_PACS_RptFragments;
/





--Ӱ�񱨸�ԭ�͹���(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptAntetype Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  --1.��ȡ�ļ�ԭ�����
  Procedure p_Get_Antetypelistkind(
    Val Out t_Refcur
	);
  --2.�����ĵ����ͻ�ȡ�ĵ���Ϣ
  Procedure p_Get_Antetypelis_By_Kind(
	Val           Out t_Refcur,
	����_In      Ӱ�񱨸�ԭ���嵥.����%Type,
	Stop_Flag    Number,
	Condition_In Varchar2
	);
  --3.���һ���ĵ�ԭ��
  Procedure p_Add_Antetypelist(
    Id_In           Ӱ�񱨸�ԭ���嵥.ID%Type,
	����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
	����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
	����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    �豸��_In		Ӱ���豸Ŀ¼.�豸��%Type,
	˵��_In         Ӱ�񱨸�ԭ���嵥.˵��%Type,
	�ɷ�����ҳ��_In Ӱ�񱨸�ԭ���嵥.�ɷ�����ҳ��%Type,
	�ɷ����ø�ʽ_In Ӱ�񱨸�ԭ���嵥.�ɷ����ø�ʽ%Type,
    �ɷ���д���_In Ӱ�񱨸�ԭ���嵥.�ɷ���д���%Type,
	�Ƿ����_In     Ӱ�񱨸�ԭ���嵥.�Ƿ����%Type,
	������_In       Ӱ�񱨸�ԭ���嵥.������%Type,
	����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
	����ѡ��_In     Ӱ�񱨸�ԭ���嵥.����ѡ��%Type,
	�ʾ����ʱ��_In Ӱ�񱨸�ԭ���嵥.�ʾ����ʱ��%Type,
	�������ʱ��_In Ӱ�񱨸�ԭ���嵥.�������ʱ��%Type,
	ר�ò��_In     Ӱ�񱨸�ԭ���嵥.ר�ò��%Type,
	Copy_Id_In      Ӱ�񱨸�ԭ���嵥.ID%Type,
	Only_Head_In    Varchar2,
	����_In         Ӱ�񱨸�ԭ���嵥.����%Type
	);
  --4.�޸�һ���ĵ�ԭ��
  Procedure p_Edit_Antetypelist(
    Id_In           Ӱ�񱨸�ԭ���嵥.ID%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    �豸��_In		Ӱ���豸Ŀ¼.�豸��%Type,
    ˵��_In         Ӱ�񱨸�ԭ���嵥.˵��%Type,
    �ɷ�����ҳ��_In Ӱ�񱨸�ԭ���嵥.�ɷ�����ҳ��%Type,
    �ɷ����ø�ʽ_In Ӱ�񱨸�ԭ���嵥.�ɷ����ø�ʽ%Type,
    �ɷ���д���_In Ӱ�񱨸�ԭ���嵥.�ɷ���д���%Type,
    �Ƿ����_In     Ӱ�񱨸�ԭ���嵥.�Ƿ����%Type,
    �޸���_In       Ӱ�񱨸�ԭ���嵥.�޸���%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����ѡ��_In     Ӱ�񱨸�ԭ���嵥.����ѡ��%Type,
	�ʾ����ʱ��_In Ӱ�񱨸�ԭ���嵥.�ʾ����ʱ��%Type,
	�������ʱ��_In Ӱ�񱨸�ԭ���嵥.�������ʱ��%Type,
    ר�ò��_In     Ӱ�񱨸�ԭ���嵥.ר�ò��%Type,
    Copy_Id_In      Ӱ�񱨸�ԭ���嵥.ID%Type,
    Only_Head_In    Varchar2,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type
	);
  --5.ɾ��һ���ļ�ԭ��
  Procedure p_Del_Antetypelist(
    Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	);
  --6.����ID��ȡ�ļ�ԭ��
  Procedure p_Get_Antetypelist_By_Id(
	Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	);
  --7.��ȡԭ��XML����
  Procedure p_Get_Antetypelist_Content(
	Val           Out t_Refcur,
	Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	);
  --8.ͣ�û������ļ�ԭ��
  Procedure p_Stop_Antetypelist(
    Id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	);

  --9.�����ĵ�������Ϣ
  Procedure p_Add_Doc_Kind(
    ����_In Ӱ�񱨸�����.����%Type,
    ����_In Ӱ�񱨸�����.����%Type,
    ˵��_In Ӱ�񱨸�����.˵��%Type
	);
  --10.ɾ���ĵ�������Ϣ
  Procedure p_Del_Doc_Kind;
  --11.��ȡԤ�������Ϣ
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	);
  --12.���Ԥ�������Ϣ
  Procedure p_Add_Pre_Outline(
    ID_In   Ӱ�񱨸�Ԥ�����.ID%Type,
	����_In Ӱ�񱨸�Ԥ�����.����%Type,
	����_In Ӱ�񱨸�Ԥ�����.����%Type,
	˵��_In Ӱ�񱨸�Ԥ�����.˵��%Type
	);
  --13.ɾ��Ԥ�������Ϣ
  Procedure p_Del_Pre_Outline;
  --14.��ȡ�������ĵ�ԭ����Ϣ
  Procedure p_Output_Antetypelist(
    Val Out t_Refcur
	);
  --15.���ԭ��Ƭ��
  Procedure p_Add_Antetype_Fragments(
    ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type,
	Ƭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.Ƭ��ID%Type
	);
  --16.ɾ��ԭ��Ƭ��
  Procedure p_Del_Antetype_Fragments(
    ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	);
  --17.��ȡԭ��Ƭ��
  Procedure p_Get_Antetype_Fragments(
    Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	);
  --18.��ȡĳ��ԭ�͹�����ĳ��Ƭ�η���
  Procedure p_Get_Antetype_f_Byaidfid(
    Val           Out t_Refcur,		
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type,
    Ƭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.Ƭ��ID%Type
	);
  --19.�����ĵ�ԭ��XML����
  Procedure p_Edit_Antetypelist_Content(
    Id_In     Ӱ�񱨸�ԭ���嵥.Id%Type,
	����_In   Ӱ�񱨸�ԭ���嵥.����%Type,
	�޸���_In Ӱ�񱨸�ԭ���嵥.�޸���%Type
	);
  --20.��ȡ����ԭ��
  Procedure p_Get_All_Antetype_Lists(
    Val Out t_Refcur
	);
  --21.��ȡ�Ѿ������˹�����ԭ��Ƭ��������Ϣ

  Procedure p_Get_Antetype_Fragments_Info(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	);
  --22.��ȡѡ����������Ķ�������
  Procedure p_Get_Selected_Fragments(
	Val           Out t_Refcur,
	ԭ��id_In Varchar2
	);
  --23.��ȡ�ܸ��Ƶ�ԭ������

  Procedure p_Get_Copy_Antetype(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�ԭ���嵥.����%Type
	);
  --24.��ȡԭ�͵ķ�����Ϣ
  Procedure p_Get_Antetype_Category(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�ԭ���嵥.����%Type
	);
  --25.����ԭ��ͬ���������
  Procedure p_Synchronous_Sample(
    ԭ��id_In Ӱ�񱨸�ԭ���嵥.Id%Type
	);
  --26.��ȡ������ԭ���б�
  Procedure p_Get_Out_Antetypelist(
    Val Out t_Refcur
	);
  --27.ͨ�������ȡԭ��������Ϣ
  Procedure p_Get_Antetype_Kind_By_Code(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�����.����%Type
	);
  --28.��ȡ�¼���Ϣ���������̶��¼�
  Procedure p_Get_Doc_Event(
    Val Out t_Refcur
	);
  --29.��ȡ����ԭ�͵������ظ���Ϣ

  Procedure p_Get_Antetypelist_Same_Info(
	Val           Out t_Refcur,
	Tablename_In Varchar2,
    Id_In        Ӱ�񱨸�ԭ���嵥.Id%Type,
    ����_In      Varchar2,
    ����_In      Varchar2
	);
  --30.��ȡ�¼��ظ�����Ϣ
  Procedure p_Event_Same_Info(
	Val           Out t_Refcur,	
	Id_In      Ӱ�񱨸��¼�.Id%Type,
    ԭ��ID_In  Ӱ�񱨸��¼�.ԭ��ID%Type,
    Ԫ��IID_In Ӱ�񱨸��¼�.Ԫ��IID%Type,
    ����_In    Ӱ�񱨸��¼�.����%Type,
    ����_In    Ӱ�񱨸��¼�.����%Type,
    ���_In    Ӱ�񱨸��¼�.���%Type
	);
  --31.��ȡԭ��У�����𼯺�
  Procedure p_Get_Process_Kind(
    Val Out t_Refcur
	);

  ----32.��ȡԪ�ػ�����ٵ����Ƽ���
  --Procedure p_Get_Antetype_Ele_Section(
  --ԭ��ID_In  Ӱ�񱨸�ԭ���嵥.Id%Type,
  --Val     Out t_Refcur);

  --33.��ȡָ��ԭ�͵��ĵ�����
  Procedure p_Get_Doc_Process_Of_Antetype(
	Val           Out t_Refcur,
	ԭ��id_In Ӱ�񱨸涯��.ԭ��id%Type
	);

  --34. �����ֵ����ƻ�ȡ��Ӧ����
  Procedure p_Get_Dictitems_By_Title(
	Val           Out t_Refcur,
	����_In Ӱ���ֵ��嵥.����%Type
	);
  --35.������е�Ԥ�����
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	);
  --36.��ȡ���дʾ���Ϣ
  Procedure p_Get_All_Fragment(
	Val           Out t_Refcur,
	ѧ��_In Varchar2
	);

  --37.��ȡ�ʾ���Ϣ
  Procedure p_Get_Fragment_Filter(
	Val           Out t_Refcur,
	ԭ��id_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type,
    ����_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    ѧ��_In   Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Type_In   Varchar2
	);
  --38.����ԭ�ͻ�ȡ������Ƭ�α�ǩֵ
  Procedure p_Get_Label_By_Aid(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	);
  --39.��ȡ���дʾ����
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	);
  --40.��ȡ������Ӧ�����༭ʱ��
  Procedure p_Get_Data_Last_Edit_Time(
	Val           Out t_Refcur,
	Table_Name_In Varchar2
	);
  --41.����ĵ��¼�
  Procedure p_Add_Doc_Event(
    ID_In       Ӱ�񱨸��¼�.ID%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ԭ��ID_In   Ӱ�񱨸��¼�.ԭ��ID%Type,
    ���_In     Ӱ�񱨸��¼�.���%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ˵��_In     Ӱ�񱨸��¼�.˵��%Type,
    Ԫ��IID_In  Ӱ�񱨸��¼�.Ԫ��IID%Type,
    ��չ���_In Ӱ�񱨸��¼�.��չ���%Type);
  --42.�޸��ĵ��¼�
  Procedure p_Update_Doc_Event(
    Id_In       Ӱ�񱨸��¼�.Id%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ˵��_In     Ӱ�񱨸��¼�.˵��%Type,
    Ԫ��IID_In  Ӱ�񱨸��¼�.Ԫ��IID%Type,
    ��չ���_In Ӱ�񱨸��¼�.��չ���%Type);
  --43.ɾ���ĵ��¼�
  Procedure p_Delete_Doc_Event(
    Id_In Ӱ�񱨸��¼�.Id%Type
	);
  --44.ɾ������δ��ʹ�õ��ĵ��¼�
  Procedure p_Delete_Unused_Doc_Events(
    Count_Out Out Number
	);
  --45.��ȡָ��ԭ�͵��ĵ��¼�
  Procedure p_Get_Doc_Event_Of_Antetype(
	Val           Out t_Refcur,
	ԭ��ID_In       Ӱ�񱨸��¼�.ԭ��ID%Type,
	Include_Base_In Number
	);
  --46.�޸��ĵ�������
  Procedure p_Update_Doc_Process_Seqnum(
    Id_In   Ӱ�񱨸涯��.Id%Type,
	���_In Ӱ�񱨸涯��.���%Type
	);
  --47.����ĵ�����
  Procedure p_Add_Doc_Process(
    Id_In           Ӱ�񱨸涯��.Id%Type,
    ԭ��ID_In       Ӱ�񱨸涯��.ԭ��ID%Type,
    �¼�ID_In       Ӱ�񱨸涯��.�¼�ID%Type,
    ��������_In     Ӱ�񱨸涯��.��������%Type,
    ����_In         Ӱ�񱨸涯��.����%Type,
    ˵��_In         Ӱ�񱨸涯��.˵��%Type,
    �ɷ��ֹ�ִ��_In Ӱ�񱨸涯��.�ɷ��ֹ�ִ��%Type,
    ���_In         Ӱ�񱨸涯��.���%Type,
    ����_In         Ӱ�񱨸涯��.����%Type
	);
  --48.�޸��ĵ�����
  Procedure p_Update_Doc_Process(
    Id_In           Ӱ�񱨸涯��.Id%Type,
    �¼�ID_In       Ӱ�񱨸涯��.�¼�ID%Type,
    ��������_In     Ӱ�񱨸涯��.��������%Type,
    ����_In         Ӱ�񱨸涯��.����%Type,
    ˵��_In         Ӱ�񱨸涯��.˵��%Type,
    �ɷ��ֹ�ִ��_In Ӱ�񱨸涯��.�ɷ��ֹ�ִ��%Type,
    ����_In         Ӱ�񱨸涯��.����%Type
	);
  --49.��ȡԪ�ػ�����ٵ����Ƽ���
  Procedure p_Get_Antetype_Ele_Section(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ���嵥.Id%Type,
	Type_In   Varchar2
	);
  --50.ɾ���ĵ�����
  Procedure p_Del_Doc_Process(
    Id_In        Ӱ�񱨸涯��.ID%Type,
	Del_Event_In Number
	);

  --51.��ѯԪ��ֵ�����ĸ������
  Procedure p_Get_Ele_Same_Info(
	Val           Out t_Refcur,
	Id_In    Ӱ�񱨸�ֵ���嵥.Id%Type,
	Code_In  Varchar2,
	Title_In Varchar2,
	Flag_In  Varchar2
	);
  --52.������еĲ����Ϣ
  Procedure p_Get_DocPluginList(
    Val Out t_Refcur
	);
  --53.��ID�Ĳ���Ƿ�ԭ��ʹ�ù�
  Procedure p_IsExit_DocPluginByID(
	Val           Out t_Refcur,
	ID_In Varchar2
	);
  --54.������������Ϣ
  Procedure p_AddDocPlugin(
    ID_In       Ӱ�񱨸���.ID%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ˵��_In     Ӱ�񱨸���.˵��%Type,
    ��ʾ��ʽ_In Ӱ�񱨸���.��ʾ��ʽ%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    �Ƿ����_In Ӱ�񱨸���.�Ƿ����%Type
	);
  --55.�޸ı�������Ϣ
  Procedure p_EditDocPlugin(
    ID_In       Ӱ�񱨸���.ID%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ˵��_In     Ӱ�񱨸���.˵��%Type,
    ��ʾ��ʽ_In Ӱ�񱨸���.��ʾ��ʽ%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    �Ƿ����_In Ӱ�񱨸���.�Ƿ����%Type
	);
  --56.ɾ����������Ϣ
  Procedure p_DelDocPlugin(
    ID_In Ӱ�񱨸���.ID%Type
	);
  --57.�ı����Ŀ���״̬
  Procedure p_IsEnableDocPlugin(
    ID_In Ӱ�񱨸���.ID%Type
	);
  --58.ͨ��ID��ö�Ӧ�Ĳ����Ϣ
  Procedure p_GetDocPluginByID(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸���.ID%type
	);
  --59.�жϱ���������Ƿ��Ѵ���
  Procedure p_IsExitDocPlugin(
	Val           Out t_Refcur,
	ID_In   Ӱ�񱨸���.ID%Type,
    ����_In Ӱ�񱨸���.����%Type,
    ����_In Ӱ�񱨸���.����%Type
	);
  --60.ͨ��ID��ö�Ӧ��ר�ò����Ϣ
  Procedure p_GetDocSpecPluginByID(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸���.ID%Type
	);
  --61.��������б���Ϣ
  Procedure p_GetDiagnosisList(
	Val           Out t_Refcur,
	���_In Varchar2,
    ����_In Varchar2
	);
  --62.�����������б�
  Procedure p_GetDiagnosisClass(
    Val Out t_Refcur
	);
  --63.���Ӱ�񱨸�ԭ��Ӧ����Ϣ
  Procedure p_AddMedicalAntetype(
    ������ĿID_In Ӱ�񱨸�ԭ��Ӧ��.������ĿID%Type,
	Ӧ�ó���_In   Ӱ�񱨸�ԭ��Ӧ��.Ӧ�ó���%Type,
	����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	);
  --64.ɾ��ԭ��ID��Ӧ�Ĳ�������Ӧ����Ϣ
  Procedure p_DelMedicalAntetype(
    ����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	);
  --65.ͨ��ԭ��ID��ö�Ӧ�Ĳ�������Ӧ����Ϣ
  Procedure p_GetMedicalByAID(
	Val           Out t_Refcur,
	����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	);
  --66.����ԭ��IDɾ��������Ϣ
  Procedure p_DelDocProcessByAid(
    ����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	);
  --67.��ȡID��Ӧ��ԭ�͵����νṹ
  Procedure p_GetAntetypeTreeByID(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸�ԭ���嵥.ID%Type
	);
  --68.ԭ���Ƿ���ڶ�Ӧ�ı��������
  procedure p_IsExitAntetype(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�ԭ���嵥.����%Type,
    ����_In Ӱ�񱨸�ԭ���嵥.����%Type,
    ID_In  Ӱ�񱨸�ԭ���嵥.ID%Type
	);

  --69  ��ȡӰ��洢�豸
  Procedure p_GetStorageDevice(
		Val           Out t_Refcur);

End b_PACS_RptAntetype;
/

--Ӱ�񱨸�ԭ�͹���(---ʵ�ֲ���---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptAntetype Is
  --Create By Hwei;
  --2014/11/25

  --1.��ȡ�ļ�ԭ�����
  Procedure p_Get_Antetypelistkind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ����, a.����, a.���� || '-' || a.���� As ����
        From Ӱ�񱨸����� A
       Order By ����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelistkind;

  --2.�����ĵ����ͻ�ȡ�ĵ���Ϣ
  Procedure p_Get_Antetypelis_By_Kind(
	Val           Out t_Refcur,
	����_In      Ӱ�񱨸�ԭ���嵥.����%Type,
    Stop_Flag    Number,
    Condition_In Varchar2
	) As
  Begin
    Open Val For
      Select ID, ����, ����, ����, ����, �Ƿ����, ˵��, Imageindex
        From (Select Distinct ���� As ID,
                              (Select Min(b.����)
                                 From Ӱ�񱨸�ԭ���嵥 B
                                Where b.���� = a.����) As ����,
                              a.���� As ����,
                              a.���� As ����,
                              null As ����,
                              0 As �Ƿ����,
                              null As ˵��,
                              0 As Imageindex
                From Ӱ�񱨸�ԭ���嵥 A
               Where a.���� = ����_In
                 And ((a.�Ƿ���� <> 1 And Stop_Flag = 1) Or (Stop_Flag = 0))
                 And ((a.���� Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)
                 And a.���� Is Not Null
              Union
              Select RawtoHex(ID) ID,
                     a.����,
                     ���� As ����,
                     ���� || '-' || ���� As ����,
                     ����,
                     a.�Ƿ����,
                     a.˵��,
                     Decode(a.�Ƿ����, 1, 2, 1) Imageindex
                From Ӱ�񱨸�ԭ���嵥 A
               Where a.
               ���� = ����_In
                 And ((a.�Ƿ���� <> 1 And Stop_Flag = 1) Or (Stop_Flag = 0))
                 And ((a.���� Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)) A
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelis_By_Kind;

  --3.���һ���ĵ�ԭ��
  Procedure p_Add_Antetypelist(
    ID_In           Ӱ�񱨸�ԭ���嵥.ID%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
	�豸��_In		Ӱ���豸Ŀ¼.�豸��%Type,
    ˵��_In         Ӱ�񱨸�ԭ���嵥.˵��%Type,
    �ɷ�����ҳ��_In Ӱ�񱨸�ԭ���嵥.�ɷ�����ҳ��%Type,
    �ɷ����ø�ʽ_In Ӱ�񱨸�ԭ���嵥.�ɷ����ø�ʽ%Type,
	�ɷ���д���_In Ӱ�񱨸�ԭ���嵥.�ɷ���д���%Type,
    �Ƿ����_In     Ӱ�񱨸�ԭ���嵥.�Ƿ����%Type,
    ������_In       Ӱ�񱨸�ԭ���嵥.������%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����ѡ��_In     Ӱ�񱨸�ԭ���嵥.����ѡ��%Type,
	�ʾ����ʱ��_In Ӱ�񱨸�ԭ���嵥.�ʾ����ʱ��%Type,
	�������ʱ��_In Ӱ�񱨸�ԭ���嵥.�������ʱ��%Type,
    ר�ò��_In     Ӱ�񱨸�ԭ���嵥.ר�ò��%Type,
    Copy_ID_In      Ӱ�񱨸�ԭ���嵥.ID%Type,
    Only_Head_In    Varchar2,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type
	) As
    x_Str Xmltype;
  Begin
    Begin
      If Copy_ID_In Is Null or Copy_ID_In = 0 Then
        x_Str := ����_In;
      Else
        Select Decode(Only_Head_In,
                      1,
                      Deletexml(a.����, '/zlxml/document/node()'),
                      a.����)
          Into x_Str
          From Ӱ�񱨸�ԭ���嵥 A
         Where a.id = Copy_ID_In;
      End If;
    Exception
      When Others Then
        x_Str := ����_In;
    End;
  
    Insert Into Ӱ�񱨸�ԭ���嵥
      (ID,
       ����,
       ����,
       ����,
	   �豸��,
       ˵��,
       �ɷ�����ҳ��,
       �ɷ����ø�ʽ,
	   �ɷ���д���,
       �Ƿ����,
       ������,
       ����ʱ��,
       ����,
       ����ѡ��,
	   �ʾ����ʱ��,
	   �������ʱ��,
       ר�ò��,
       ����)
    Values
      (ID_In,
       ����_In,
       ����_In,
       ����_In,
	   �豸��_In,
       ˵��_In,
       �ɷ�����ҳ��_In,
       �ɷ����ø�ʽ_In,
	   �ɷ���д���_In,
       �Ƿ����_In,
       ������_In,
       sysdate,
       x_Str,
       ����ѡ��_In,
	   �ʾ����ʱ��_In,
	   �������ʱ��_In,
       ר�ò��_In,
       ����_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Antetypelist;

  --4.�޸�һ���ĵ�ԭ��
  Procedure p_Edit_Antetypelist(
    ID_In           Ӱ�񱨸�ԭ���嵥.ID%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
	�豸��_In		Ӱ���豸Ŀ¼.�豸��%Type,
    ˵��_In         Ӱ�񱨸�ԭ���嵥.˵��%Type,
    �ɷ�����ҳ��_In Ӱ�񱨸�ԭ���嵥.�ɷ�����ҳ��%Type,
    �ɷ����ø�ʽ_In Ӱ�񱨸�ԭ���嵥.�ɷ����ø�ʽ%Type,
	�ɷ���д���_In Ӱ�񱨸�ԭ���嵥.�ɷ���д���%Type,
    �Ƿ����_In     Ӱ�񱨸�ԭ���嵥.�Ƿ����%Type,
    �޸���_In       Ӱ�񱨸�ԭ���嵥.�޸���%Type,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type,
    ����ѡ��_In     Ӱ�񱨸�ԭ���嵥.����ѡ��%Type,
	�ʾ����ʱ��_In Ӱ�񱨸�ԭ���嵥.�ʾ����ʱ��%Type,
	�������ʱ��_In Ӱ�񱨸�ԭ���嵥.�������ʱ��%Type,
    ר�ò��_In     Ӱ�񱨸�ԭ���嵥.ר�ò��%Type,
    Copy_ID_In      Ӱ�񱨸�ԭ���嵥.ID%Type,
    Only_Head_In    Varchar2,
    ����_In         Ӱ�񱨸�ԭ���嵥.����%Type
	) As
    x_Str     Xmltype;
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.Id)
      Into n_Num
      From Ӱ�񱨸�ԭ���嵥 A
     Where (a.���� = ����_In Or a.���� = ����_In)
       And ID <> ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]������ͬ���ĵ�����������ƣ���������д��[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Copy_ID_In Is Null or Copy_ID_In = 0 Then
      x_Str := ����_In;
    Else
      Select Decode(Only_Head_In,
                    1,
                    Deletexml(a.����, '/zlxml/document/node()'),
                    a.����)
        Into x_Str
        From Ӱ�񱨸�ԭ���嵥 A
       Where a.id = Copy_ID_In;
    End If;
  
    Update Ӱ�񱨸�ԭ���嵥
       Set ����         = ����_In,
           ����         = ����_In,
           ����         = ����_In,
		   �豸��		= �豸��_In,
           ˵��         = ˵��_In,
           �ɷ�����ҳ�� = �ɷ�����ҳ��_In,
           �ɷ����ø�ʽ = �ɷ����ø�ʽ_In,
		   �ɷ���д��� = �ɷ���д���_In,
           �Ƿ����     = NVL(�Ƿ����_In, �Ƿ����),
           �޸���       = �޸���_In,
           �޸�ʱ��     = sysdate,
           ����         = x_Str,
           ����ѡ��     = ����ѡ��_In,
		   �ʾ����ʱ�� =�ʾ����ʱ��_In,
		   �������ʱ�� =�������ʱ��_In,
           ר�ò��     = ר�ò��_In,
           ����         = ����_In
     Where ID = ID_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Antetypelist;

  --5.ɾ��һ���ļ�ԭ��
  Procedure p_Del_Antetypelist(
    ID_In Ӱ�񱨸�ԭ���嵥.Id%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(ID) Into n_Num From Ӱ�񱨸��¼ A Where a.ԭ��id = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]��ԭ���Ѿ����ĵ�ʹ�ã�������ɾ����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(ԭ��ID)
      Into n_Num
      From Ӱ�񱨸�ԭ��Ƭ��
     Where Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]���ĵ��´��ڴʾ������������ɾ����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(ID)
      Into n_Num
      From Ӱ�񱨸淶���嵥
     Where Ӱ�񱨸淶���嵥.ԭ��ID = ID_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]�����Դ�ԭ�ͽ����ķ�����Ϣ��������ɾ����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Delete From Ӱ�񱨸�ԭ���嵥 C Where c.Id = ID_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Antetypelist;

  --6.����ID��ȡ�ļ�ԭ��
  Procedure p_Get_Antetypelist_By_Id(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸�ԭ���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select rawtohex(a.ID) ID,
             a.����,
             a.����,
             a.����,
			 a.�豸��,
             a.˵��,
             a.�ɷ�����ҳ��,
             a.�ɷ����ø�ʽ,
			 a.�ɷ���д���,
             Extractvalue(b.Column_Value, '/root/print_hf_mode') Printhfmode,
             Extractvalue(b.Column_Value, '/root/print_follow_pages') Printfollowpages,
             Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit,
             (Nvl(a.����ѡ��, XmlType('<NULL/>'))).GetClobVal() as ����ѡ��,
			 a.�ʾ����ʱ��,
			 a.�������ʱ��,
             a.�Ƿ����,
             (Nvl(a.ר�ò��, XmlType('<NULL/>'))).GetClobVal() as ר�ò��,
             a.������,
             a.����ʱ��,
             a.�޸���,
             a.�޸�ʱ��,
             a.����
        From Ӱ�񱨸�ԭ���嵥 A,
             Table(Xmlsequence(Extract(a.����ѡ��, '/root'))) B
       Where a.Id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_By_Id;

  --7.��ȡԭ��XML����
  Procedure p_Get_Antetypelist_Content(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸�ԭ���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select (Nvl(a.����, XmlType('<ZLXML/>'))).GetClobVal() As ����
        From Ӱ�񱨸�ԭ���嵥 A
       Where a.Id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_Content;

  --8.ͣ�û������ļ�ԭ��
  Procedure p_Stop_Antetypelist(
    ID_In Ӱ�񱨸�ԭ���嵥.Id%Type
	) As
  Begin
    Update Ӱ�񱨸�ԭ���嵥
       Set �Ƿ���� = Decode(�Ƿ����, 1, 0, 0, 1)
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Stop_Antetypelist;

  --9.�����ĵ�������Ϣ
  Procedure p_Add_Doc_Kind(
    ����_In Ӱ�񱨸�����.����%Type,
    ����_In Ӱ�񱨸�����.����%Type,
    ˵��_In Ӱ�񱨸�����.˵��%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.����)
      Into n_Num
      From Ӱ�񱨸����� A
     Where a.���� = ����_In
        Or a.���� = ����_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]����ı���������Ʋ�����ͬ��[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If ����_In Is Null Or ����_In Is Null Then
      v_Err_Msg := '[ZLSOFT]����ı���������Ʋ���Ϊ�գ�[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Insert Into Ӱ�񱨸�����
      (����, ����, ˵��)
    Values
      (����_In, ����_In, ˵��_In);
  
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Kind;

  --10.ɾ���ĵ�������Ϣ
  Procedure p_Del_Doc_Kind As
  Begin
    Delete From Ӱ�񱨸�����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Doc_Kind;

  --11.��ȡԤ�������Ϣ
  Procedure p_Get_Pre_Outline(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID, ����, ����, ˵��, ���༭ʱ��
        From Ӱ�񱨸�Ԥ����� A
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Pre_Outline;

  --12.���Ԥ�������Ϣ
  Procedure p_Add_Pre_Outline(
    ID_In   Ӱ�񱨸�Ԥ�����.ID%Type,
    ����_In Ӱ�񱨸�Ԥ�����.����%Type,
    ����_In Ӱ�񱨸�Ԥ�����.����%Type,
    ˵��_In Ӱ�񱨸�Ԥ�����.˵��%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(200);
    Err_Item Exception;
  Begin
    Select Count(a.����)
      Into n_Num
      From Ӱ�񱨸�Ԥ����� A
     Where a.���� = ����_In
        Or a.���� = ����_In;
  
    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]����ı���������Ʋ�����ͬ��[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If ����_In Is Null Or ����_In Is Null Then
      v_Err_Msg := '[ZLSOFT]����ı���������Ʋ���Ϊ�գ�[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Insert Into Ӱ�񱨸�Ԥ�����
      (ID, ����, ����, ˵��, ���༭ʱ��)
    Values
      (ID_In, ����_In, ����_In, ˵��_In, sysdate);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Pre_Outline;

  --13.ɾ��Ԥ�������Ϣ
  Procedure p_Del_Pre_Outline As
  Begin
    Delete From Ӱ�񱨸�Ԥ�����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Pre_Outline;

  --14.��ȡ�������ĵ�ԭ����Ϣ
  Procedure p_Output_Antetypelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select '���' As ���,
             b.���� As ID,
             Null As ����,
             b.���� As ��������,
             b.���� As ����,
             b.���� As ����,
             b.˵�� As ˵��,
             Null As �ɷ�����ҳ��,
             Null As �ɷ����ø�ʽ,
             Null As �Ƿ����,
             Null As ������,
             Null As ����ʱ��,
             Null As �޸���,
             Null As �޸�ʱ��,
             Null As ����
        From Ӱ�񱨸����� B
      Union All
      Select 'ԭ��' ���,
             RawToHex(a.Id) ID,
             a.����,
             b.���� ��������,
             a.����,
             a.����,
             a.˵��,
             a.�ɷ�����ҳ��,
             a.�ɷ����ø�ʽ,
             a.�Ƿ����,
             a.������,
             a.����ʱ��,
             a.�޸���,
             a.�޸�ʱ��,
             Null As ����
        From Ӱ�񱨸�ԭ���嵥 A, Ӱ�񱨸����� B
       Where a.���� = b.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Output_Antetypelist;

  --15.���ԭ��Ƭ��
  Procedure p_Add_Antetype_Fragments(
    ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type,
	Ƭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.Ƭ��ID%Type) As
  Begin
    Insert Into Ӱ�񱨸�ԭ��Ƭ��
      (ԭ��ID, Ƭ��ID)
    Values
      (ԭ��ID_In, Ƭ��ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Antetype_Fragments;

  --16.ɾ��ԭ��Ƭ��
  Procedure p_Del_Antetype_Fragments(
    ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸�ԭ��Ƭ�� Where ԭ��ID = ԭ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Antetype_Fragments;

  --17.��ȡԭ��Ƭ��
  Procedure p_Get_Antetype_Fragments(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(Ƭ��ID) Ƭ��ID
        From Ӱ�񱨸�ԭ��Ƭ�� A
       Where a.ԭ��ID = ԭ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Fragments;

  --18.��ȡĳ��ԭ�͹�����ĳ��Ƭ�η���
  Procedure p_Get_Antetype_f_Byaidfid(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type,
	Ƭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.Ƭ��ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(Ƭ��ID) Ƭ��ID
        From Ӱ�񱨸�ԭ��Ƭ�� A
       Where a.ԭ��ID = ԭ��ID_In
         And a.Ƭ��ID = Ƭ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_f_Byaidfid;

  --19.�����ĵ�ԭ��XML����
  Procedure p_Edit_Antetypelist_Content(
    ID_In     Ӱ�񱨸�ԭ���嵥.Id%Type,
	����_In   Ӱ�񱨸�ԭ���嵥.����%Type,
	�޸���_In Ӱ�񱨸�ԭ���嵥.�޸���%Type
	) As
  Begin
    Update Ӱ�񱨸�ԭ���嵥
       Set ���� = ����_In, �޸��� = �޸���_In, �޸�ʱ�� = Sysdate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Antetypelist_Content;

  --20.��ȡ����ԭ��
  Procedure p_Get_All_Antetype_Lists(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID,
             a.����,
             ���� || '-' || ���� As ����,
             ����,
             a.����,
             a.�Ƿ����,
             a.˵��,
             Decode(a.�Ƿ����, 1, 2, 1) Imageindex,
             (Nvl(a.����, XmlType('<ZLXML/>'))).GetClobVal() As ����
        From Ӱ�񱨸�ԭ���嵥 A
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Antetype_Lists;

  --21.��ȡ�Ѿ������˹�����ԭ��Ƭ��������Ϣ
  Procedure p_Get_Antetype_Fragments_Info(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(ID) ID,
             a.����,
             a.����,
             a.���� || '-' || a.���� ����,
             a.˵��
        From Ӱ�񱨸�Ƭ���嵥 A
       Where a.Id In (Select b.Ƭ��id
                        From Ӱ�񱨸�ԭ��Ƭ�� B
                       Where b.ԭ��id = ԭ��ID_In)
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Fragments_Info;

  --22.��ȡѡ����������Ķ�������
  Procedure p_Get_Selected_Fragments(
	Val           Out t_Refcur,
	ԭ��ID_In Varchar2
	) As
    v_Sql  Varchar2(4000);
    v_Aids Varchar2(4000);
    v_Msg  Varchar2(4000);
    Err Exception;
  Begin
    For Myrow In (Select RawtoHex(a.Ƭ��id) ID
                    From Ӱ�񱨸�ԭ��Ƭ�� A
                   Where a.ԭ��id = ԭ��ID_In) Loop
      If v_Aids Is Null Then
        v_Aids := '''' || Myrow.Id || '''';
      Else
        v_Aids := v_Aids || ',''' || Myrow.Id || '''';
      End If;
    End Loop;
  
    If v_Aids Is Null Then
      If Substr(ԭ��ID_In, 0, 1) <> '''' Then
        v_Aids := '''' || ԭ��ID_In || '''';
      Else
        v_Aids := ԭ��ID_In;
      End If;
    End If;
  
    v_Sql := 'Select Distinct  RawtoHex(a.id) ID,  RawtoHex(a.�ϼ�ID) �ϼ�ID , a.����, a.���� || ''-'' || a.���� ����,Decode(a.�ڵ�����, 0, 0, 1) �ڵ�����
      From Ӱ�񱨸�Ƭ���嵥 A
      Start With a.Id In (' || v_Aids || ')
      Connect By Prior a.Id = a.�ϼ�ID
      Order By a.����';
  
    Open Val For v_Sql;
  Exception
    When Err Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Selected_Fragments;

  --23.��ȡ�ܸ��Ƶ�ԭ������
  Procedure p_Get_Copy_Antetype(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�ԭ���嵥.����%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(a.id) ID, a.���� || '-' || a.���� ����
        From Ӱ�񱨸�ԭ���嵥 A
       Where a.���� = ����_In
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Copy_Antetype;

  --24.��ȡԭ�͵ķ�����Ϣ
  Procedure p_Get_Antetype_Category(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�ԭ���嵥.����%Type
	) As
  Begin
    Open Val For
      Select Distinct a.���� As ����
        From Ӱ�񱨸�ԭ���嵥 A
       Where a.���� = ����_In
         and a.���� Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Category;

  --25.����ԭ��ͬ���������
  Procedure p_Synchronous_Sample(
    ԭ��ID_In Ӱ�񱨸�ԭ���嵥.Id%Type
	) As
    x_Content Xmltype;
    x_Result  Xmltype;
    Cursor c_Antetype Is
      Select Extractvalue(c.Column_Value, '/section/@iid') Iid,
             Extractvalue(c.Column_Value, '/section/@title') Title,
             c.Column_Value As Content
        From Ӱ�񱨸�ԭ���嵥 A,
             Table(Xmlsequence(Extract(a.����, '/zlxml//section'))) C
       Where a.Id = ԭ��ID_In;
    n_i               Number;
    n_j               Number;
    n_Count           Number;
    x_Subdocuments    Xmltype;
    x_Docparameters   Xmltype;
    x_Antetypecontent Xmltype;
    v_Textstyleno     Varchar2(10);
    v_Parastyleno     Varchar2(10);
    x_Acontent        Xmltype;
  Begin
    For Mysample In (Select b.id, b.����
                       From Ӱ�񱨸淶���嵥 B
                      Where b.ԭ��id = ԭ��ID_In) Loop
      x_Content := Mysample.����;
      n_i       := 1;
      If x_Content Is Null Then
        Select a.����
          Into x_Result
          From Ӱ�񱨸�ԭ���嵥 A
         Where a.Id = ԭ��ID_In;
      Else
        Begin
          Select Extractvalue(c.Column_Value, '/section/@textstyleno') Textstyleno,
                 Extractvalue(c.Column_Value, '/section/@parastyleno') Parastyleno
            Into v_Textstyleno, v_Parastyleno
            From Table(Xmlsequence(Extract(x_Content, '/zlxml//section'))) C
           Where Rownum = 1;
        Exception
          When Others Then
            v_Textstyleno := '1';
            v_Parastyleno := '1';
        End;
      
        For Myantetype In c_Antetype Loop
          For I In 1 .. 1 Loop
            If n_i <> 1 Or n_Count <> 0 Or n_Count Is Null Then
              Select Count(*)
                Into n_Count
                From Table(Xmlsequence(Extract(x_Content, '/zlxml//section'))) C;
            End If;
            If n_Count < n_i Then
              Select Updatexml(Myantetype.Content,
                               '//section/@textstyleno',
                               v_Textstyleno)
                Into x_Acontent
                From Dual;
              Select Updatexml(x_Acontent,
                               '//section/@parastyleno',
                               v_Parastyleno)
                Into x_Acontent
                From Dual;
              Select Appendchildxml(x_Content,
                                    '/zlxml/document',
                                    x_Acontent)
                Into x_Content
                From Dual;
              Exit;
            End If;
            n_j := 1;
            For Mysample In (Select Extractvalue(c.Column_Value,
                                                 '/section/@iid') Iid,
                                    Extractvalue(c.Column_Value,
                                                 '/section/@title') Title
                               From Table(Xmlsequence(Extract(x_Content,
                                                              '/zlxml//section'))) C) Loop
              If n_i = n_j Then
                If Myantetype.Iid <> Mysample.Iid Then
                  Select Updatexml(Myantetype.Content,
                                   '//section/@textstyleno',
                                   v_Textstyleno)
                    Into x_Acontent
                    From Dual;
                  Select Updatexml(x_Acontent,
                                   '//section/@parastyleno',
                                   v_Parastyleno)
                    Into x_Acontent
                    From Dual;
                  Select Deletexml(x_Content,
                                   '//section[@iid="' || Myantetype.Iid || '"]')
                    Into x_Content
                    From Dual;
                  Select Insertxmlbefore(x_Content,
                                         '//section[@iid="' || Mysample.Iid || '"]',
                                         x_Acontent)
                    Into x_Content
                    From Dual;
                  n_j := n_j + 1;
                  Exit;
                Else
                  n_j := n_j + 1;
                  Exit;
                End If;
              End If;
              n_j := n_j + 1;
            End Loop;
            n_i := n_i + 1;
          End Loop;
        End Loop;
        x_Result := x_Content;
        For Mysample2 In (Select Iid
                            From (Select Extractvalue(c.Column_Value,
                                                      '/section/@iid') Iid
                                    From Table(Xmlsequence(Extract(x_Content,
                                                                   '/zlxml//section'))) C) C
                           Where c.Iid Not In
                                 (Select Extractvalue(c.Column_Value,
                                                      '/section/@iid') Iid
                                    From Ӱ�񱨸�ԭ���嵥 A,
                                         Table(Xmlsequence(Extract(a.����,
                                                                   '/zlxml//section'))) C
                                   Where a.Id = ԭ��ID_In)) Loop
          Select Deletexml(x_Result,
                           '//section[@iid="' || Mysample2.Iid || '"]')
            Into x_Result
            From Dual;
        End Loop;
      End If;
    
      Update Ӱ�񱨸淶���嵥 X
         Set x.���� = x_Result
       Where x.Id = Mysample.Id;
    End Loop;
  
    Select a.����
      Into x_Antetypecontent
      From Ӱ�񱨸�ԭ���嵥 A
     Where a.Id = ԭ��ID_In;
    Select Extract(x_Antetypecontent, 'zlxml/subdocuments')
      Into x_Subdocuments
      From Dual;
    Select Extract(x_Antetypecontent, 'zlxml/docparameters')
      Into x_Docparameters
      From Dual;
    Update Ӱ�񱨸淶���嵥
       Set ���� = Updatexml(����, '/zlxml/subdocuments', x_Subdocuments)
     Where ԭ��ID = ԭ��ID_In;
    Update Ӱ�񱨸淶���嵥
       Set ���� = Updatexml(����, '/zlxml/docparameters', x_Docparameters)
     Where ԭ��ID = ԭ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Synchronous_Sample;

  --26.��ȡ������ԭ���б�
  Procedure p_Get_Out_Antetypelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select ID,
             ����,
             ����,
             Parentid,
             ����,
             �Ƿ����,
             ˵��,
             Imageindex,
             ����
        From (Select a.���� As ID,
                     a.���� As ����,
                     a.���� As ����,
                     '' As Parentid,
                     '-1' As ����,
                     0 As �Ƿ����,
                     a.˵�� As ˵��,
                     4 As Imageindex,
                     a.���� ����
                From Ӱ�񱨸����� A
              Union
              Select Distinct a.���� || '-' || a.���� As ID,
                              (Select Min(����)
                                 From Ӱ�񱨸�ԭ���嵥 B
                                Where b.���� = a.����) As ����,
                              Max(a.����) As ����,
                              a.���� As Parentid,
                              '0' As ����,
                              0 As �Ƿ����,
                              '' As ˵��,
                              4 As Imageindex,
                              a.����
                From Ӱ�񱨸�ԭ���嵥 A
               Where a.���� Is Not Null
               Group By a.����, a.����
              Union
              Select RawTohex(ID),
                     a.����,
                     ���� || '-' || ���� As ����,
                     Decode(a.����, Null, a.����, a.���� || '-' || a.����) Parentid,
                     a.���� As ����,
                     a.�Ƿ����,
                     a.˵��,
                     Decode(a.�Ƿ����, 1, 1, 0, 2),
                     a.����
                From Ӱ�񱨸�ԭ���嵥 A) A
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Out_Antetypelist;

  --27.ͨ�������ȡԭ��������Ϣ
  Procedure p_Get_Antetype_Kind_By_Code(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�����.����%Type
	) As
  Begin
    Open Val For
      Select a.����, a.����, a.˵��
        From Ӱ�񱨸����� A
       Where a.���� = ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Kind_By_Code;
  --28.��ȡ�¼���Ϣ���������̶��¼�
  Procedure p_Get_Doc_Event(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawtoHex(a.id) ID,
             a.����,
             a.ԭ��id,
             a.���,
             a.����,
             a.˵��,
             a.Ԫ��iid,
             a.��չ���
        From Ӱ�񱨸��¼� A
       Where a.���� <> 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Event;

  --29.��ȡ����ԭ�͵������ظ���Ϣ
  Procedure p_Get_Antetypelist_Same_Info(
	Val           Out t_Refcur,
	Tablename_In Varchar2,
	ID_In        Ӱ�񱨸�ԭ���嵥.Id%Type,
	����_In      Varchar2,
	����_In      Varchar2
	) As
    n_Num    Number;
    v_Result Varchar2(100);
    v_Sql    Varchar2(4000);
  Begin
    If ID_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where id=' ||
               ID_In;
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        v_Result := 'ID�ظ�';
      End If;
    End If;
    If ����_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where ����=''' ||
               ����_In || '''';
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        If v_Result Is Not Null Then
          v_Result := v_Result || ',�����ظ�';
        Else
          v_Result := '�����ظ�';
        End If;
      End If;
    End If;
    If ����_In Is Not Null Then
      v_Sql := 'select count(*) from ' || Tablename_In || ' where ����=''' ||
               ����_In || '''';
      Execute Immediate v_Sql
        Into n_Num;
      If n_Num > 0 Then
        If v_Result Is Not Null Then
          v_Result := v_Result || ',�����ظ�';
        Else
          v_Result := '�����ظ�';
        End If;
      End If;
    End If;
    Open Val For
      Select v_Result Result From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetypelist_Same_Info;

  --30.��ȡ�¼��ظ�����Ϣ
  Procedure p_Event_Same_Info(
	Val           Out t_Refcur,
	ID_In      Ӱ�񱨸��¼�.Id%Type,
    ԭ��ID_In  Ӱ�񱨸��¼�.ԭ��ID%Type,
    Ԫ��IID_In Ӱ�񱨸��¼�.Ԫ��IID%Type,
    ����_In    Ӱ�񱨸��¼�.����%Type,
    ����_In    Ӱ�񱨸��¼�.����%Type,
    ���_In    Ӱ�񱨸��¼�.���%Type
	) As
    v_Same_Antetype Varchar2(50);
    n_Same_Id       Number;
    n_Same_Title    Number;
    n_Same_Seqnum   Number;
    n_Maxnum        Number;
  Begin
    Select Count(*)
      Into n_Same_Title
      From Ӱ�񱨸��¼� A
     Where a.ԭ��ID = ԭ��ID_In
       And a.���� = ����_In
       And a.���� = ����_In;
    Select Count(*)
      Into n_Same_Seqnum
      From Ӱ�񱨸��¼� A
     Where a.ԭ��ID = ԭ��ID_In
       And a.���� = ����_In
       And a.��� = ���_In;
    Begin
      Select a.Id
        Into v_Same_Antetype
        From Ӱ�񱨸��¼� A
       Where a.ԭ��ID = ԭ��ID_In
         And a.Ԫ��IID = Ԫ��IID_In;
    Exception
      When Others Then
        v_Same_Antetype := '';
    End;
  
    Select Count(*) Into n_Same_Id From Ӱ�񱨸��¼� A Where a.Id = ID_In;
    Select Max(a.���) Into n_Maxnum From Ӱ�񱨸��¼� A;
  
    Open Val For
      Select v_Same_Antetype As Sameaid,
             n_Same_Id       As Sameid,
             n_Same_Title    As Sametitle,
             n_Same_Seqnum   As Sameseqnum,
             n_Maxnum        As Maxnum
        From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Event_Same_Info;

  --31.��ȡԭ��У�����𼯺�
  Procedure p_Get_Process_Kind(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Distinct ��������
        From (Select Extractvalue(c.Column_Value, '/step/kind') As ��������
                From Ӱ�񱨸涯�� A,
                     Table(Xmlsequence(Extract(a.����, '/root/step'))) C) B
       Where b.�������� Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Process_Kind;


  --33.��ȡָ��ԭ�͵��ĵ�����
  Procedure p_Get_Doc_Process_Of_Antetype(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸涯��.ԭ��id%Type
	) As
  Begin
    Open Val For
      Select RawtoHex(p.id) ID,
             p.����,
             p.��������,
             p.���,
             p.˵��,
             p.�ɷ��ֹ�ִ��,
             (Nvl(p.����, XmlType('<NULL/>'))).GetClobVal() As ����, --Nvl(p.����,'<NULL/>') As ����,
             RawtoHex(p.�¼�ID) �¼�ID,
             0 Is_Event
        From Ӱ�񱨸涯�� P
       Where p.ԭ��ID = ԭ��ID_In
      Union All
      Select RawtoHex(e.id) ID,
             e.����,
             e.����,
             e.���,
             e.˵��,
             Null,
             (XmlType('<Null/>')).GetClobVal() As ����, --(Null,'<NULL/>') As ����,
             Null,
             1
        From Ӱ�񱨸��¼� E
       Where e.Id In (Select RawtoHex(�¼�ID) �¼�ID
                        From Ӱ�񱨸涯��
                       Where ԭ��ID = ԭ��ID_In)
       Order By Is_Event, ��������, ���;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Process_Of_Antetype;

  --34. �����ֵ����ƻ�ȡ��Ӧ����
  Procedure p_Get_Dictitems_By_Title(
	Val           Out t_Refcur,
	����_In Ӱ���ֵ��嵥.����%Type
	) As
  Begin
    Open Val For
      Select a.���, a.����, Rawtohex(a.�ֵ�id) As �ֵ�ID
        From Ӱ���ֵ����� A
       Where a.�ֵ�id In (Select id From Ӱ���ֵ��嵥 b Where b.���� = ����_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Dictitems_By_Title;

  --35.������е�Ԥ�����
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(ID) ID, a.����, a.����
        From Ӱ�񱨸�Ԥ����� a
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Phr_Onlines;

  --36.��ȡ���дʾ���Ϣ
  Procedure p_Get_All_Fragment(
	Val           Out t_Refcur,
	ѧ��_In Varchar2
	) As
  Begin
    If ѧ��_In <> '' Then
      Open Val For
        Select RawToHex(a.id) ID,
               RawToHex(a.�ϼ�id) �ϼ�id,
               a.����,
               a.����,
               a.˵��,
               a.�ڵ�����,
               (Nvl(a.���, XmlType('<NULL/>'))).GetClobVal() As ���,
               a.ѧ��,
               a.��ǩ,
               a.�Ƿ�˽��,
               a.����
          From Ӱ�񱨸�Ƭ���嵥 A
         Where (a.ѧ�� In
               (Select /*+rule*/
                  Column_Value As Lable
                   From Table(b_PACS_RptPublic.f_Str2list(ѧ��_In, ','))
                 Intersect
                 Select /*+rule*/
                  Column_Value As Lable
                   From Table(b_PACS_RptPublic.f_Str2list(a.ѧ��, ','))) And
               a.�ڵ����� <> 0)
            Or a.�ڵ����� = 0
            Or a.ѧ�� Is Null
         Order By a.����, a.�ϼ�id;
    Else
      Open Val For
        Select RawToHex(a.id) ID,
               RawToHex(a.�ϼ�id) �ϼ�id,
               a.����,
               a.����,
               a.˵��,
               a.�ڵ�����,
               (Nvl(a.���, XmlType('<NULL/>'))).GetClobVal() As ���,
               a.ѧ��,
               a.��ǩ,
               a.�Ƿ�˽��,
               a.����
          From Ӱ�񱨸�Ƭ���嵥 A
         Order By a.�ϼ�id, a.�ڵ�����, a.����, a.����;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Fragment;

  --37. ��ȡ�ʾ���Ϣ
  Procedure p_Get_Fragment_Filter(
	Val           Out t_Refcur,
	ԭ��id_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type,
    ����_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    ѧ��_In   Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Type_In   Varchar2
	) As
  Begin
    If Type_In = '1' Then
      Open Val For
        Select Rawtohex(b.Id) ID,
               Rawtohex(b.�ϼ�id) �ϼ�id,
               b.����,
               b.����,
               b.˵��,
               b.�ڵ�����,
               (Nvl(b.���, XmlType('<NULL/>'))).GetClobVal() As ���,
               b.ѧ��,
               b.��ǩ,
               b.�Ƿ�˽��,
               b.����,
               b.���༭ʱ��
          From Ӱ�񱨸�ԭ��Ƭ�� A, Ӱ�񱨸�Ƭ���嵥 B
         Where a.Ƭ��id = b.id
           And a.ԭ��id = ԭ��id_In;
    Else
      Open Val For
        Select /*+ rule*/
         Rawtohex(b.Id) ID,
         Rawtohex(b.�ϼ�id) �ϼ�id,
         b.����,
         b.����,
         b.˵��,
         b.�ڵ�����,
         (Nvl(b.���, XmlType('<NULL/>'))).GetClobVal() As ���,
         b.ѧ��,
         b.��ǩ,
         b.�Ƿ�˽��,
         b.����,
         b.���༭ʱ��
          From Ӱ�񱨸�Ƭ���嵥 B
         Where b.�ϼ�id = ԭ��id_In
           And (b.�Ƿ�˽�� = 0 Or (b.�Ƿ�˽�� = 1 And b.���� = ����_In))
           And (b.ѧ�� Is Null Or
               (b.ѧ�� Is Not Null And
               b_PACS_RptPublic.f_If_Intersect(b.ѧ��, ѧ��_In) > 0));
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Fragment_Filter;

  --38.����ԭ�ͻ�ȡ������Ƭ�α�ǩֵ
  Procedure p_Get_Label_By_Aid(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ��Ƭ��.ԭ��ID%Type
	) As
  Begin
    Open Val For
      Select Distinct b.��ǩ
        From Ӱ�񱨸�Ƭ���嵥 B
       Start With b.�ϼ�id In (Select a.Ƭ��id
                               From Ӱ�񱨸�ԭ��Ƭ�� A
                              Where a.ԭ��id = ԭ��ID_In)
      Connect By Prior b.Id = b.�ϼ�id;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Label_By_Aid;

  --39.��ȡ���дʾ����
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) ID,
             Rawtohex(a.�ϼ�id) �ϼ�id,
             a.����,
             a.����,
             a.˵��,
             a.�ڵ�����
        From Ӱ�񱨸�Ƭ���嵥 A
       Where a.�ڵ����� = 0
       Start With �ϼ�id Is Null
      Connect By Prior id = �ϼ�id
       Order By ����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_All_Fragment_Class;

  --40.��ȡ������Ӧ�����༭ʱ��
  Procedure p_Get_Data_Last_Edit_Time(
	Val           Out t_Refcur,
	Table_Name_In Varchar2
	) As
    v_sql Varchar2(4000);
  Begin
    v_sql := 'select max(���༭ʱ��) maxvalue from ' || Table_Name_In;
    Open val For v_sql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Data_Last_Edit_Time;

  --41.����ĵ��¼�
  Procedure p_Add_Doc_Event(
    ID_In       Ӱ�񱨸��¼�.ID%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ԭ��ID_In   Ӱ�񱨸��¼�.ԭ��ID%Type,
    ���_In     Ӱ�񱨸��¼�.���%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ˵��_In     Ӱ�񱨸��¼�.˵��%Type,
    Ԫ��IID_In  Ӱ�񱨸��¼�.Ԫ��IID%Type,
    ��չ���_In Ӱ�񱨸��¼�.��չ���%Type
	) As
    n_Seq_Num  Ӱ�񱨸��¼�.���%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(*)
      Into n_Is_Exist
      From Ӱ�񱨸��¼�
     Where ԭ��ID = ԭ��ID_In
       And ���� = ����_In
       And ���� = ����_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]ԭ�����Ѵ�����ͬ�������¼�[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If (���_In Is Null Or ���_In = 0) Then
      Select Nvl(Max(���), 0) + 1 Into n_Seq_Num From Ӱ�񱨸��¼�;
    Else
      Select Count(*)
        Into n_Is_Exist
        From Ӱ�񱨸��¼�
       Where ԭ��ID = ԭ��ID_In
         And ���� = ����_In
         And ��� = ���_In;
      If n_Is_Exist > 0 Then
        v_Err_Msg := '[ZLSOFT]ԭ�����Ѵ�����ͬ��ŵ��¼�[ZLSOFT]';
        Raise Err_Item;
      End If;
      n_Seq_Num := ���_In;
    End If;
  
    Insert Into Ӱ�񱨸��¼�
      (ID, ����, ԭ��ID, ���, ����, ˵��, Ԫ��IID, ��չ���)
    Values
      (ID_In,
       ����_In,
       ԭ��ID_In,
       n_Seq_Num,
       ����_In,
       ˵��_In,
       Ԫ��IID_In,
       ��չ���_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Event;

  --42.�޸��ĵ��¼�
  Procedure p_Update_Doc_Event(
    Id_In       Ӱ�񱨸��¼�.Id%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ����_In     Ӱ�񱨸��¼�.����%Type,
    ˵��_In     Ӱ�񱨸��¼�.˵��%Type,
    Ԫ��IID_In  Ӱ�񱨸��¼�.Ԫ��IID%Type,
    ��չ���_In Ӱ�񱨸��¼�.��չ���%Type
	) As
    r_Aid      Ӱ�񱨸��¼�.ԭ��ID%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select ԭ��ID Into r_Aid From Ӱ�񱨸��¼� Where ID = Id_In;
  
    Select Count(*)
      Into n_Is_Exist
      From Ӱ�񱨸��¼�
     Where ԭ��ID = r_Aid
       And ���� = ����_In
       And ���� = ����_In
       And ID <> Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]ԭ���ϴ�����ͬ�������¼�[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Update Ӱ�񱨸��¼�
       Set ����     = ����_In,
           ����     = ����_In,
           ˵��     = ˵��_In,
           Ԫ��IID  = Ԫ��IID_In,
           ��չ��� = ��չ���_In
     Where ID = Id_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Event;

  --43.ɾ���ĵ��¼�
  Procedure p_Delete_Doc_Event(
    Id_In Ӱ�񱨸��¼�.Id%Type
	) As
    n_Kind     Ӱ�񱨸��¼�.����%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select ���� Into n_Kind From Ӱ�񱨸��¼� Where ID = Id_In;
  
    If n_Kind = 1 Then
      v_Err_Msg := '[ZLSOFT]������ɾ���̶��¼�[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Count(*) Into n_Is_Exist From Ӱ�񱨸涯�� Where �¼�ID = Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]�¼��Ѿ���ʹ��,���ܱ�ɾ����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Delete From Ӱ�񱨸��¼� Where ID = Id_In;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Delete_Doc_Event;

  --44.ɾ������δ��ʹ�õ��ĵ��¼�
  Procedure p_Delete_Unused_Doc_Events(
    Count_Out Out Number
	) As
  Begin
    Delete From Ӱ�񱨸��¼�
     Where ���� <> 1
       And ID Not In
           (Select �¼�ID From Ӱ�񱨸涯�� Where �¼�ID Is Not Null);
    Count_Out := Sql%RowCount;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Delete_Unused_Doc_Events;

  --45.��ȡָ��ԭ�͵��ĵ��¼�
  Procedure p_Get_Doc_Event_Of_Antetype(
	Val           Out t_Refcur,
	ԭ��ID_In       Ӱ�񱨸��¼�.ԭ��ID%Type,
	Include_Base_In Number
	) As
  Begin
    If Include_Base_In = 1 Then
      Open Val For
        Select Rawtohex(t.Id) ID,
               t.����,
               t.����,
               t.˵��,
               t.Ԫ��iid,
               t.��չ���,
               Nvl(p.Used_Count, 0) Used_Count
          From Ӱ�񱨸��¼� T,
               (Select Count(*) Used_Count, Max(�¼�ID) �¼�ID
                  From Ӱ�񱨸涯��
                 Where �¼�ID Is Not Null
                 Group By �¼�ID) P
         Where (t.���� = 1 Or t.ԭ��id = ԭ��ID_In)
           And t.Id = p.�¼�ID(+)
         Order By t.���;
    Else
      Open Val For
        Select Rawtohex(t.Id) ID,
               t.����,
               t.����,
               t.˵��,
               t.Ԫ��iid,
               t.��չ���,
               Nvl(p.Used_Count, 0) Used_Count
          From Ӱ�񱨸��¼� T,
               (Select Count(*) Used_Count, Max(�¼�ID) �¼�ID
                  From Ӱ�񱨸涯��
                 Where �¼�ID Is Not Null
                 Group By �¼�ID) P
         Where t.ԭ��id = ԭ��ID_In
           And t.���� <> 1
           And t.Id = p.�¼�ID(+)
         Order By t.���;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Event_Of_Antetype;

  --46.�޸��ĵ�������
  Procedure p_Update_Doc_Process_Seqnum(
    Id_In   Ӱ�񱨸涯��.Id%Type,
	���_In Ӱ�񱨸涯��.���%Type) As
  Begin
    Update Ӱ�񱨸涯�� Set ��� = ���_In Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Process_Seqnum;

  --47.����ĵ�����
  Procedure p_Add_Doc_Process(
    Id_In           Ӱ�񱨸涯��.Id%Type,
    ԭ��ID_In       Ӱ�񱨸涯��.ԭ��ID%Type,
    �¼�ID_In       Ӱ�񱨸涯��.�¼�ID%Type,
    ��������_In     Ӱ�񱨸涯��.��������%Type,
    ����_In         Ӱ�񱨸涯��.����%Type,
    ˵��_In         Ӱ�񱨸涯��.˵��%Type,
    �ɷ��ֹ�ִ��_In Ӱ�񱨸涯��.�ɷ��ֹ�ִ��%Type,
    ���_In         Ӱ�񱨸涯��.���%Type,
    ����_In         Ӱ�񱨸涯��.����%Type
	) As
    n_Seq_Num  Ӱ�񱨸涯��.���%Type;
    n_Is_Exist Number(1) := 0;
    v_Err_Msg  Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(*)
      Into n_Is_Exist
      From Ӱ�񱨸涯��
     Where ԭ��ID = ԭ��ID_In
       And ���� = ����_In;
    If (���_In Is Null Or ���_In = 0) Then
      If (�¼�ID_In Is Null) Then
        Select Nvl(Max(���), 0) + 1
          Into n_Seq_Num
          From Ӱ�񱨸涯��
         Where ԭ��ID = ԭ��ID_In
           And �¼�ID Is Null;
      Else
        Select Nvl(Max(���), 0) + 1
          Into n_Seq_Num
          From Ӱ�񱨸涯��
         Where ԭ��ID = ԭ��ID_In
           And �¼�ID = �¼�ID_In;
      End If;
    Else
      n_Seq_Num := ���_In;
    End If;
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]ԭ���ϴ�����ͬ�����Ķ���[ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into Ӱ�񱨸涯��
      (ID, ԭ��ID, �¼�ID, ��������, ����, ˵��, �ɷ��ֹ�ִ��, ���, ����)
    Values
      (Id_In,
       ԭ��ID_In,
       �¼�ID_In,
       ��������_In,
       ����_In,
       ˵��_In,
       �ɷ��ֹ�ִ��_In,
       n_Seq_Num,
       ����_In);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Add_Doc_Process;

  --48.�޸��ĵ�����
  Procedure p_Update_Doc_Process(
    Id_In           Ӱ�񱨸涯��.Id%Type,
    �¼�ID_In       Ӱ�񱨸涯��.�¼�ID%Type,
    ��������_In     Ӱ�񱨸涯��.��������%Type,
    ����_In         Ӱ�񱨸涯��.����%Type,
    ˵��_In         Ӱ�񱨸涯��.˵��%Type,
    �ɷ��ֹ�ִ��_In Ӱ�񱨸涯��.�ɷ��ֹ�ִ��%Type,
    ����_In         Ӱ�񱨸涯��.����%Type
	) As
    r_Aid          Ӱ�񱨸��¼�.ԭ��ID%Type;
    r_Old_Event_Id Ӱ�񱨸涯��.�¼�ID%Type;
    n_Seq_Num      Ӱ�񱨸��¼�.���%Type;
    n_Is_Exist     Number(1) := 0;
    v_Err_Msg      Varchar2(100);
    Err_Item Exception;
  Begin
    Select ԭ��ID Into r_Aid From Ӱ�񱨸涯�� Where ID = Id_In;
    If (�¼�ID_In Is Not Null) Then
      Select Count(*)
        Into n_Is_Exist
        From Ӱ�񱨸��¼�
       Where (ԭ��ID Is Null Or ԭ��ID = r_Aid)
         And ID = �¼�ID_In;
    
      If n_Is_Exist = 0 Then
        v_Err_Msg := '[ZLSOFT]�������¼�������[ZLSOFT]';
        Raise Err_Item;
      End If;
    
    End If;
  
    Select Count(*)
      Into n_Is_Exist
      From Ӱ�񱨸涯��
     Where ԭ��ID = r_Aid
       And ���� = ����_In
       And ID <> Id_In;
  
    If n_Is_Exist > 0 Then
      v_Err_Msg := '[ZLSOFT]ԭ���ϴ�����ͬ�����Ķ���[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If (r_Old_Event_Id <> �¼�ID_In Or
       (�¼�ID_In Is Null And r_Old_Event_Id Is Not Null)) Then
      If (�¼�ID_In Is Null) Then
        Select Nvl(Max(���), 0) + 1
          Into n_Seq_Num
          From Ӱ�񱨸涯��
         Where ԭ��ID = r_Aid
           And �¼�ID Is Null;
      Else
        Select Nvl(Max(���), 0) + 1
          Into n_Seq_Num
          From Ӱ�񱨸涯��
         Where ԭ��ID = r_Aid
           And �¼�ID = �¼�ID_In;
      End If;
    Else
      n_Seq_Num := 0;
    End If;
  
    If n_Seq_Num > 0 Then
      Update Ӱ�񱨸涯��
         Set �¼�id       = �¼�ID_In,
             ��������     = ��������_In,
             ����         = ����_In,
             ˵��         = ˵��_In,
             �ɷ��ֹ�ִ�� = �ɷ��ֹ�ִ��_In,
             ����         = ����_In,
             ���         = n_Seq_Num
       Where ID = Id_In;
    Else
      Update Ӱ�񱨸涯��
         Set �¼�id       = �¼�ID_In,
             ��������     = ��������_In,
             ����         = ����_In,
             ˵��         = ˵��_In,
             �ɷ��ֹ�ִ�� = �ɷ��ֹ�ִ��_In,
             ����         = ����_In
       Where ID = Id_In;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Doc_Process;

  --49.��ȡԪ�ػ�����ٵ����Ƽ���
  Procedure p_Get_Antetype_Ele_Section(
	Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸�ԭ���嵥.Id%Type,
	Type_In   Varchar2
	) As
    c_Content Clob;
  Begin
    /*Select To_Clob(a.����)*/
    Select a.����.getclobval()
      Into c_Content
      From Ӱ�񱨸�ԭ���嵥 A
     Where a.Id = ԭ��ID_In;
  
    If Type_In = '1' Then
      Open Val For
        Select Distinct Name
          From (Select Extractvalue(c.Column_Value, '/*/@title') As Name
                  From Table(Xmlsequence(Extract(Xmltype(c_Content),
                                                 '/zlxml/document//element[@sid and @title]|/zlxml/document//e_list[@sid and @title]|/zlxml/document//e_enum[@sid and @title]|/zlxml/document//e_etree[@sid and @title]|/zlxml/document//e_utree[@sid and @title]'))) C) A
         Where a.Name Is Not Null;
    Else
      Open Val For
        Select Distinct Name
          From (Select Extractvalue(c.Column_Value, '/section/@title') As Name
                  From Table(Xmlsequence(Extract(Xmltype(c_Content),
                                                 '//section'))) C) A
         Where a.Name Is Not Null;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Antetype_Ele_Section;

  --50.ɾ���ĵ�����
  Procedure p_Del_Doc_Process(Id_In        Ӱ�񱨸涯��.ID%Type,
                              Del_Event_In Number) As
    r_Event_Id   Ӱ�񱨸涯��.�¼�ID%Type := Null;
    n_Event_Kind Ӱ�񱨸��¼�.����%Type;
    n_Is_Exist   Number(1) := 0;
  Begin
    If Del_Event_In = 1 Then
      Select Max(e.Id), Max(e.����)
        Into r_Event_Id, n_Event_Kind
        From Ӱ�񱨸涯�� P, Ӱ�񱨸��¼� E
       Where p.Id = Id_In
         And p.�¼�id = e.Id;
    End If;
  
    Delete From Ӱ�񱨸涯�� Where ID = Id_In;
  
    If Del_Event_In = 1 Then
      If (r_Event_Id Is Not Null And n_Event_Kind <> 1) Then
        Select Count(*)
          Into n_Is_Exist
          From Ӱ�񱨸涯��
         Where �¼�id = r_Event_Id;
        If n_Is_Exist = 0 Then
          Delete From Ӱ�񱨸��¼�
           Where ID = r_Event_Id
             And ���� <> 1;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Doc_Process;

  --51.��ѯԪ��ֵ�����ĸ������
  Procedure p_Get_Ele_Same_Info(
	Val           Out t_Refcur,
	Id_In    Ӱ�񱨸�ֵ���嵥.Id%Type,
	Code_In  Varchar2,
	Title_In Varchar2,
	Flag_In  Varchar2
	) As
    v_Result  Varchar2(50);
    v_Id      Varchar2(50);
    v_Code_Id Varchar2(50);
    n_Num     Number;
  Begin
    If Flag_In = 1 Then
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�Ԫ�ط��� A
       Where a.Id = Id_In;
    
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From Ӱ�񱨸�Ԫ�ط��� A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID�ظ�';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�Ԫ�ط��� A
       Where a.���� = Code_In;
    
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From Ӱ�񱨸�Ԫ�ط��� A
         Where a.���� = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',�����ظ�';
          Else
            v_Result := '�����ظ�';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�Ԫ�ط��� A
       Where a.���� = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From Ӱ�񱨸�Ԫ�ط��� A
         Where a.���� = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',�����ظ�';
          Else
            v_Result := '�����ظ�';
          End If;
        End If;
      End If;
    
    End If;
  
    If Flag_In = 2 Then
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�Ԫ���嵥 A
       Where a.Id = Id_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From Ӱ�񱨸�Ԫ���嵥 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID�ظ�';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�Ԫ���嵥 A
       Where a.���� = Code_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From Ӱ�񱨸�Ԫ���嵥 A
         Where a.���� = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',�����ظ�';
          Else
            v_Result := '�����ظ�';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�Ԫ���嵥 A
       Where a.���� = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From Ӱ�񱨸�Ԫ���嵥 A
         Where a.���� = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',�����ظ�';
          Else
            v_Result := '�����ظ�';
          End If;
        End If;
      End If;
    
    End If;
  
    If Flag_In = 3 Then
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�ֵ���嵥 A
       Where a.Id = Id_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From Ӱ�񱨸�ֵ���嵥 A
         Where a.Id = Id_In;
        If v_Id Is Not Null Then
          v_Result := 'ID�ظ�';
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�ֵ���嵥 A
       Where a.���� = Code_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Code_Id
          From Ӱ�񱨸�ֵ���嵥 A
         Where a.���� = Code_In;
        If v_Code_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',�����ظ�';
          Else
            v_Result := '�����ظ�';
          End If;
        End If;
      End If;
    
      Select Count(ID)
        Into n_Num
        From Ӱ�񱨸�ֵ���嵥 A
       Where a.���� = Title_In;
      If n_Num > 0 Then
        Select Rawtohex(a.Id)
          Into v_Id
          From Ӱ�񱨸�ֵ���嵥 A
         Where a.���� = Title_In;
        If v_Id Is Not Null Then
          If v_Result Is Not Null Then
            v_Result := v_Result || ',�����ظ�';
          Else
            v_Result := '�����ظ�';
          End If;
        End If;
      End If;
    
    End If;
  
    Open Val For
      Select v_Result As Result, v_Id As ID, v_Code_Id As Codesameid
        From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ele_Same_Info;

  --52.������еĲ����Ϣ
  Procedure p_Get_DocPluginList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             ����,
             ����,
             ˵��,
             ��ʾ��ʽ,
             ����,
             Decode(��ʾ��ʽ, '1', 'Ƕ��ʽ', '����ʽ') ��ʾ��ʽII,
             Decode(����, '1', 'ר�ò��', '������') ����II,
             ����,
             ����,
             �Ƿ����,
             Decode(�Ƿ����, '1', 'ͣ��', '����') IsEnable
        From Ӱ�񱨸���;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_DocPluginList;

  --53.��ID�Ĳ���Ƿ�ԭ��ʹ�ù�
  Procedure p_IsExit_DocPluginByID(
    Val           Out t_Refcur,
	ID_In Varchar2
	) As
    CURSOR C_EVENT Is
      Select t.ר�ò��.getclobval() ר�ò�� From Ӱ�񱨸�ԭ���嵥 t;
    anum Int := 0;
    sult Varchar2(6666);
  Begin
    For temp In C_EVENT Loop
      If instr(temp.ר�ò��, ID_In) > 0 Then
        anum := anum + 1;
      End If;
    End Loop;
    Open Val For
      Select anum From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsExit_DocPluginByID;

  --54.������������Ϣ
  Procedure p_AddDocPlugin(
    ID_In       Ӱ�񱨸���.ID%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ˵��_In     Ӱ�񱨸���.˵��%Type,
    ��ʾ��ʽ_In Ӱ�񱨸���.��ʾ��ʽ%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    �Ƿ����_In Ӱ�񱨸���.�Ƿ����%Type
	) As
  Begin
    Insert Into Ӱ�񱨸���
      (ID, ����, ����, ˵��, ��ʾ��ʽ, ����, ����, ����, �Ƿ����)
    Values
      (ID_In,
       ����_In,
       ����_In,
       ˵��_In,
       ��ʾ��ʽ_In,
       ����_In,
       ����_In,
       ����_In,
       �Ƿ����_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddDocPlugin;

  --55.�޸ı�������Ϣ
  Procedure p_EditDocPlugin(
    ID_In       Ӱ�񱨸���.ID%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ˵��_In     Ӱ�񱨸���.˵��%Type,
    ��ʾ��ʽ_In Ӱ�񱨸���.��ʾ��ʽ%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    ����_In     Ӱ�񱨸���.����%Type,
    �Ƿ����_In Ӱ�񱨸���.�Ƿ����%Type
	) As
  Begin
    Update Ӱ�񱨸���
       Set ����     = ����_In,
           ����     = ����_In,
           ˵��     = ˵��_In,
           ��ʾ��ʽ = ��ʾ��ʽ_In,
           ����     = ����_In,
           ����     = ����_In,
           ����     = ����_In,
           �Ƿ���� = �Ƿ����_In
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditDocPlugin;

  --56.ɾ����������Ϣ
  Procedure p_DelDocPlugin(
    ID_In Ӱ�񱨸���.ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸��� Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelDocPlugin;

  --57.�ı����Ŀ���״̬
  Procedure p_IsEnableDocPlugin(
    ID_In Ӱ�񱨸���.ID%Type
	) As
  Begin
    Update Ӱ�񱨸��� a
       Set �Ƿ���� = Decode(a.�Ƿ����, 1, 0, 1)
     Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsEnableDocPlugin;

  --58.ͨ��ID��ö�Ӧ�Ĳ����Ϣ
  Procedure p_GetDocPluginByID(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸���.ID%type
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             ����,
             ����,
             ˵��,
             ��ʾ��ʽ,
             ����,
             ����,
             ����,
             �Ƿ����
        From Ӱ�񱨸���
       Where id = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDocPluginByID;

  --59.�жϱ���������Ƿ��Ѵ���
  Procedure p_IsExitDocPlugin(
	Val           Out t_Refcur,
	ID_In   Ӱ�񱨸���.ID%Type,
	����_In Ӱ�񱨸���.����%Type,
	����_In Ӱ�񱨸���.����%Type
	) As
  Begin
    Open Val For
      Select Count(id)
        From Ӱ�񱨸��� a
       Where (a.���� = ����_In Or a.���� = ����_In)
         and a.id <> ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_IsExitDocPlugin;

  --60.ͨ��ID��ö�Ӧ��ר�ò����Ϣ
  Procedure p_GetDocSpecPluginByID(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸���.ID%Type
	) As
  Begin
    Open Val For
      Select Rawtohex(Id) ID,
             ����,
             ����,
             ˵��,
             ��ʾ��ʽ,
             ����,
             ����,
             ����,
             �Ƿ����
        From Ӱ�񱨸���
       Where id = ID_In
         And ���� = 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDocSpecPluginByID;

  --61.��������б���Ϣ
  Procedure p_GetDiagnosisList(
	Val           Out t_Refcur,
	���_In Varchar2,
	����_In Varchar2
	) As
  Begin
    Open Val For
      Select to_char(a.id) ID,
             a.����,
             a.����,
             (Select b.���� From ������Ŀ��� b Where b.���� = a.���) ���
        From ������ĿĿ¼ a
       Where (a.id In (Select t.������Ŀid From Ӱ������Ŀ t) And a.��� = ���_In)
         And (a.���� Like ����_In || '%' Or a.���� Like ����_In || '%');
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDiagnosisList;

  --62.�����������б�
  Procedure p_GetDiagnosisClass(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select t.����, t.����, t.���� From ������Ŀ��� t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDiagnosisClass;

  --63.���Ӱ�񱨸�ԭ��Ӧ����Ϣ
  Procedure p_AddMedicalAntetype(
    ������ĿID_In Ӱ�񱨸�ԭ��Ӧ��.������ĿID%Type,
    Ӧ�ó���_In   Ӱ�񱨸�ԭ��Ӧ��.Ӧ�ó���%Type,
    ����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	) As
  Begin
    Insert Into Ӱ�񱨸�ԭ��Ӧ��
      (������ĿID, Ӧ�ó���, ����ԭ��ID)
    Values
      (������ĿID_In, Ӧ�ó���_In, ����ԭ��ID_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddMedicalAntetype;

  --64.ɾ��ԭ��ID��Ӧ�Ĳ�������Ӧ����Ϣ
  Procedure p_DelMedicalAntetype(
    ����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸�ԭ��Ӧ�� Where ����ԭ��ID = ����ԭ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelMedicalAntetype;

  --65.ͨ��ԭ��ID��ö�Ӧ�Ĳ�������Ӧ����Ϣ
  Procedure p_GetMedicalByAID(
	Val           Out t_Refcur,
	����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	) As
  Begin
    Open Val For
      Select id,
             x.����,
             x.����,
             x.���,
             Sum(x.����) ����,
             Sum(x.סԺ) סԺ,
             Sum(x.����) ����,
             Sum(x.���) ���
        From (Select id,
                     ����,
                     ����,
                     ���,
                     Decode(Ӧ�ó���, '1', 1, 0) as ����,
                     Decode(Ӧ�ó���, '2', 1, 0) as סԺ,
                     Decode(Ӧ�ó���, '3', 1, 0) as ����,
                     Decode(Ӧ�ó���, '4', 1, 0) as ���
                From (Select to_Char(a.������Ŀid) ID,
                             (Select b.����
                                From ������ĿĿ¼ b
                               Where b.id = a.������Ŀid) as ����,
                             (Select b.����
                                From ������ĿĿ¼ b
                               Where b.id = a.������Ŀid) as ����,
                             (Select c.����
                                From ������Ŀ��� c
                               Where c.���� = (Select b.���
                                               From ������ĿĿ¼ b
                                              Where b.id = a.������Ŀid)) As ���,
                             a.Ӧ�ó���
                        From Ӱ�񱨸�ԭ��Ӧ�� a
                       Where a.����ԭ��id = ����ԭ��ID_In)) x
       Group By x.id, x.����, x.����, x.���;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetMedicalByAID;

  --66.����ԭ��IDɾ��������Ϣ
  Procedure p_DelDocProcessByAid(
    ����ԭ��ID_In Ӱ�񱨸�ԭ��Ӧ��.����ԭ��ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸涯�� t Where t.ԭ��id = ����ԭ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;
  --67.��ȡID��Ӧ��ԭ�͵����νṹ
  Procedure p_GetAntetypeTreeByID(
	Val           Out t_Refcur,
	ID_In Ӱ�񱨸�ԭ���嵥.ID%Type
	) As
  Begin
    Open Val For
      Select ID, ����, ����, ����, ����, �Ƿ����, ˵��, Imageindex
        From (Select Distinct ���� As ID,
                              (Select Min(b.����)
                                 From Ӱ�񱨸�ԭ���嵥 B
                                Where b.���� = a.����) As ����,
                              a.���� As ����,
                              a.���� As ����,
                              null As ����,
                              0 As �Ƿ����,
                              null As ˵��,
                              0 As Imageindex
                From Ӱ�񱨸�ԭ���嵥 A
               Where a.id = ID_In
              Union
              Select RawtoHex(ID) ID,
                     a.����,
                     a.���� As ����,
                     ���� || '-' || ���� As ����,
                     ����,
                     a.�Ƿ����,
                     a.˵��,
                     Decode(a.�Ƿ����, 1, 2, 1) Imageindex
                From Ӱ�񱨸�ԭ���嵥 A
               Where a.id = ID_In) A
       Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetAntetypeTreeByID;

  --68.ԭ���Ƿ���ڶ�Ӧ�ı��������
  procedure p_IsExitAntetype(
	Val           Out t_Refcur,
	����_In Ӱ�񱨸�ԭ���嵥.����%Type,
	����_In Ӱ�񱨸�ԭ���嵥.����%Type,
	ID_In  Ӱ�񱨸�ԭ���嵥.ID%Type
	) As
  begin
    Open Val For
      Select Count(*) AS num
        From Ӱ�񱨸�ԭ���嵥 t
       where (t.���� = ����_In
          or t.���� = ����_In) and t.id<>ID_In;
  End p_IsExitAntetype;

  --69. ��ȡӰ��洢�豸
  Procedure p_GetStorageDevice(
	Val           Out t_Refcur
	) Is 
  Begin 
	Open Val For
		Select �豸��||' - '||�豸�� As �洢�豸, �豸��, IP��ַ, FTPĿ¼, FTP�û���, FTP����, ����Ŀ¼�û���, ����Ŀ¼����, ����Ŀ¼  
		From Ӱ���豸Ŀ¼ Where ���� = 1;
	Exception
	  When Others Then
	  Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStorageDevice;
End b_PACS_RptAntetype;
/





CREATE OR REPLACE Package b_Pacs_RptSampleList Is
  Type t_Refcur Is Ref Cursor;

  -- Author  : SEEKING
  -- Created : 2014/10/30 10:05:38
  -- Purpose : ���Ĺ���

  --���ҷ��ĵ�ԭ�����

  Procedure p_Get_Sample_List_Type(
    Val Out t_Refcur,
    Type_In Varchar2,
    Kind_In varchar2
	);

  --�����ĵ�ԭ��
  Procedure p_Add_Sample_List(
    Id_In       Ӱ�񱨸淶���嵥.Id%Type,
	Aid_In      Ӱ�񱨸淶���嵥.ԭ��id%Type,
	Seq_Num_In  Ӱ�񱨸淶���嵥.���%Type,
	Title_In    Ӱ�񱨸淶���嵥.����%Type,
	Note_In     Ӱ�񱨸淶���嵥.˵��%Type,
	Content_In  Ӱ�񱨸淶���嵥.����%Type,
	Subject_In  Ӱ�񱨸淶���嵥.ѧ��%Type,
	Label_In    Ӱ�񱨸淶���嵥.��ǩ%Type,
	Private_In  Ӱ�񱨸淶���嵥.�Ƿ�˽��%Type,
	Author_In   Ӱ�񱨸淶���嵥.����%Type,
	Lasttime_In Ӱ�񱨸淶���嵥.���༭ʱ��%Type
	);

  --�༭������Ϣ
  Procedure p_Edit_Sample_List(
    Id_In       Ӱ�񱨸淶���嵥.Id%Type,
    Aid_In      Ӱ�񱨸淶���嵥.ԭ��id%Type,
    Seq_Num_In  Ӱ�񱨸淶���嵥.���%Type,
    Title_In    Ӱ�񱨸淶���嵥.����%Type,
    Note_In     Ӱ�񱨸淶���嵥.˵��%Type,
    Content_In  Ӱ�񱨸淶���嵥.����%Type,
    Subject_In  Ӱ�񱨸淶���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸淶���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸淶���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸淶���嵥.����%Type,
    Lasttime_In Ӱ�񱨸淶���嵥.���༭ʱ��%Type
	);
  --ɾ���ĵ�����
  Procedure p_Del_Sample_List(
    Id_In Ӱ�񱨸淶���嵥.Id%Type
	);

  --ͨ��ԭ��ID�����Ӧ�ķ�����Ϣ
  Procedure p_Get_Samplelist_By_Aid(
    Val Out t_Refcur,
    Antetypelist_Id_In Ӱ�񱨸淶���嵥.ԭ��id%Type,
    Author_In          Ӱ�񱨸淶���嵥.����%Type,
    Subjects_In        Varchar2
	);

  --ͨ������id��ȡ������
  Procedure p_Get_Samplelist_By_Kind(
    Val Out t_Refcur,
    Kind_In      Varchar2,
    Condition_In Varchar2,
    Author_In    Ӱ�񱨸淶���嵥.����%Type,
    Subjects_In  Varchar2
	);

  --ͨ��ID������Ӧ�ķ�����Ϣ
  Procedure p_Get_Samplelist_By_Id(
    Val Out t_Refcur,                                   
    Id_In Ӱ�񱨸淶���嵥.Id%Type
	);

  --��ѯ��Ӧԭ���µ�������
  Procedure p_Get_Samplelist_Maxseqnum(
    Val Out t_Refcur,
	Aid_In Ӱ�񱨸淶���嵥.ԭ��id%Type
	);

  --��ȡ����XML��Ϣ
  Procedure p_Get_Samplexml(
    Val Out t_Refcur,     
	Id_In Ӱ�񱨸淶���嵥.Id%Type
	);

  --�޸ķ���XML��Ϣ
  Procedure p_Edit_Samplexml(
    Id_In      Ӱ�񱨸淶���嵥.Id%Type,
	Content_In Ӱ�񱨸淶���嵥.����%Type
	);

  --�����ķ����б�
  Procedure p_Output_Samplelist(
    Val Out t_Refcur
	);

  --�Ƿ������Ӧ��ԭ�����
  Procedure p_If_Exist_Antetypelist(
    Val Out t_Refcur,
	Title_In Ӱ�񱨸�ԭ���嵥.����%Type
	);

  --ͬһ��������Ƿ������ͬ���Ƶķ���
  Procedure p_If_Exist_Samplelist(
    Val Out t_Refcur,
    Type_In  Ӱ�񱨸�ԭ���嵥.����%Type,
    Title_In Ӱ�񱨸淶���嵥.����%Type
	);

  --ͨ������ID��÷��Ķ�Ӧ�����νṹ
  Procedure p_Get_SamplelistTree_By_Id(
    Val Out t_Refcur,
    Id_In Ӱ�񱨸淶���嵥.Id%Type
	);

End b_Pacs_RptSampleList;
/

CREATE OR REPLACE Package Body b_Pacs_RptSampleList Is

  ------------------------------------------------------------------------
  --���Ĺ���
  ------------------------------------------------------------------------

  --���ҷ��ĵ�ԭ�����

  Procedure p_Get_Sample_List_Type(
    Val Out t_Refcur,
    Type_In Varchar2,
    Kind_In Varchar2
	) As
  Begin
    If Type_In = '1' Then
      Open Val For
        Select Rawtohex(a.Id) ID, a.���� || '-' || a.���� ����
          From Ӱ�񱨸�ԭ���嵥 A
         Where a.Id In (Select Distinct b.ԭ��id From Ӱ�񱨸淶���嵥 B)
           and a.���� = Kind_In;
    Else
      Open Val For
        Select Rawtohex(a.Id) ID, a.���� || '-' || a.���� ����
          From Ӱ�񱨸�ԭ���嵥 A
         where a.���� = Kind_In
         Order By a.����;
    End If;
  
  End p_Get_Sample_List_Type;
  --�����ĵ�ԭ��
  Procedure p_Add_Sample_List(
    Id_In       Ӱ�񱨸淶���嵥.Id%Type,
	Aid_In      Ӱ�񱨸淶���嵥.ԭ��id%Type,
	Seq_Num_In  Ӱ�񱨸淶���嵥.���%Type,
	Title_In    Ӱ�񱨸淶���嵥.����%Type,
	Note_In     Ӱ�񱨸淶���嵥.˵��%Type,
	Content_In  Ӱ�񱨸淶���嵥.����%Type,
	Subject_In  Ӱ�񱨸淶���嵥.ѧ��%Type,
	Label_In    Ӱ�񱨸淶���嵥.��ǩ%Type,
	Private_In  Ӱ�񱨸淶���嵥.�Ƿ�˽��%Type,
	Author_In   Ӱ�񱨸淶���嵥.����%Type,
	Lasttime_In Ӱ�񱨸淶���嵥.���༭ʱ��%Type
	) As
    n_Num Number;
    v_Msg Varchar2(200);
    Err Exception;
  Begin
    Select Count(a.Id)
      Into n_Num
      From Ӱ�񱨸淶���嵥 A
     Where a.ԭ��id = Hextoraw(Aid_In)
       And a.���� = Title_In;
  
    If n_Num > 0 Then
      v_Msg := '[ZLSOFT]��ͬһ��ԭ���µķ������Ʋ�����ͬ��[ZLSOFT]';
      Raise Err;
    End If;
  
    Insert Into Ӱ�񱨸淶���嵥
      (ID,
       ԭ��id,
       ���,
       ����,
       ˵��,
       ����,
       ѧ��,
       ��ǩ,
       �Ƿ�˽��,
       ����,
       ���༭ʱ��)
    Values
      (Hextoraw(Id_In),
       Hextoraw(Aid_In),
       Seq_Num_In,
       Title_In,
       Note_In,
       Content_In,
       Subject_In,
       Label_In,
       Private_In,
       Author_In,
       Sysdate);
  
    --������Ӷ��ڸ÷��ĵĴ���,ҳüҳ�ţ�ҳ������
    Update Ӱ�񱨸淶���嵥
       Set ���� = Updatexml(����,
                          '/zlxml/subdocuments',
                          (Select Extract(����, 'zlxml/subdocuments')
                             From Ӱ�񱨸�ԭ���嵥
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --����ҳüҳ��
  
    Update Ӱ�񱨸淶���嵥
       Set ���� = Updatexml(����,
                          '/zlxml/docparameters',
                          (Select Extract(����, 'zlxml/docparameters')
                             From Ӱ�񱨸�ԭ���嵥
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --����ҳ������
  
  Exception
    When Err Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Sample_List;

  --�༭������Ϣ
  Procedure p_Edit_Sample_List(
    Id_In       Ӱ�񱨸淶���嵥.Id%Type,
    Aid_In      Ӱ�񱨸淶���嵥.ԭ��id%Type,
    Seq_Num_In  Ӱ�񱨸淶���嵥.���%Type,
    Title_In    Ӱ�񱨸淶���嵥.����%Type,
    Note_In     Ӱ�񱨸淶���嵥.˵��%Type,
    Content_In  Ӱ�񱨸淶���嵥.����%Type,
    Subject_In  Ӱ�񱨸淶���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸淶���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸淶���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸淶���嵥.����%Type,
    Lasttime_In Ӱ�񱨸淶���嵥.���༭ʱ��%Type
	) As
    n_Num Number;
    v_Msg Varchar2(200);
    Err Exception;
  Begin
    Select Count(a.Id)
      Into n_Num
      From Ӱ�񱨸淶���嵥 A
     Where a.ԭ��id = Hextoraw(Aid_In)
       And a.���� = Title_In
       And a.Id <> Hextoraw(Id_In);
  
    If n_Num > 0 Then
      v_Msg := '[ZLSOFT]��ͬһ��ԭ���µķ������Ʋ�����ͬ��[ZLSOFT]';
      Raise Err;
    End If;
    Update Ӱ�񱨸淶���嵥
       Set ԭ��id       = Hextoraw(Aid_In),
           ���         = Decode(Seq_Num_In, 0, ���, Seq_Num_In),
           ����         = Title_In,
           ˵��         = Note_In,
           ����         = Content_In,
           ѧ��         = Subject_In,
           ��ǩ         = Label_In,
           �Ƿ�˽��     = Private_In,
           ����         = Author_In,
           ���༭ʱ�� = Sysdate
     Where ID = Hextoraw(Id_In);
  
    --������Ӷ��ڸ÷��ĵĴ���,ҳüҳ�ţ�ҳ������
    Update Ӱ�񱨸淶���嵥
       Set ���� = Updatexml(����,
                          '/zlxml/subdocuments',
                          (Select Extract(����, 'zlxml/subdocuments')
                             From Ӱ�񱨸�ԭ���嵥
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --����ҳüҳ��
  
    Update Ӱ�񱨸淶���嵥
       Set ���� = Updatexml(����,
                          '/zlxml/docparameters',
                          (Select Extract(����, 'zlxml/docparameters')
                             From Ӱ�񱨸�ԭ���嵥
                            Where ID = Hextoraw(Aid_In)))
     Where ID = Hextoraw(Id_In); --����ҳ������
  
  Exception
    When Err Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Sample_List;

  --ɾ���ĵ�����
  Procedure p_Del_Sample_List(
    Id_In Ӱ�񱨸淶���嵥.Id%Type
	) As
  Begin
    Delete From Ӱ�񱨸淶���嵥
     Where Ӱ�񱨸淶���嵥.Id = Hextoraw(Id_In);
  End p_Del_Sample_List;

  --ͨ��ԭ��ID��ȡ��ԭ���µķ����б�
  Procedure p_Get_Samplelist_By_Aid(
    Val Out t_Refcur,
    Antetypelist_Id_In Ӱ�񱨸淶���嵥.ԭ��id%Type,
    Author_In          Ӱ�񱨸淶���嵥.����%Type,
    Subjects_In        Varchar2
	) As
  Begin
    Open Val For
      Select /*+rule*/
       Rawtohex(a.Id) As ID,
       a.����,
       a.����,
       a.˵��,
       a.ѧ��,
       a.��� Seqnum,
       a.��ǩ,
       a.�Ƿ�˽��
        From Ӱ�񱨸淶���嵥 A
       Where a.ԭ��id = Hextoraw(Antetypelist_Id_In)
         And (a.���� = Author_In Or (a.ѧ�� Is Null And a.�Ƿ�˽�� = 0) Or
             Subjects_In Is Null Or
             (a.ѧ�� Is Not Null And
             b_pacs_rptpublic.f_If_Intersect(a.ѧ��, Subjects_In) > 0 And
             a.�Ƿ�˽�� = 0));
  End p_Get_Samplelist_By_Aid;

  --ͨ������id��ȡ������
  Procedure p_Get_Samplelist_By_Kind(
    Val Out t_Refcur,
    Kind_In      Varchar2,
    Condition_In Varchar2,
    Author_In    Ӱ�񱨸淶���嵥.����%Type,
    Subjects_In  Varchar2
	) As
  Begin
  
    --���һ������ԭ����Ϣ�ķ������νṹ
    Open Val For
      Select a.���� As ID,
             a.���� As ����,
             '' ˵��,
             '' As �ϼ�id,
             ' category' As Type,
             '' As ����,
             '' As ѧ��,
             Null �޸�ʱ��,
             '' As ��ǩ,
             0 As Private,
             0 As Imgindex
        From Ӱ�񱨸�ԭ���嵥 A
       Where a.���� = Kind_In
         And Exists
       (Select ID From Ӱ�񱨸淶���嵥 C Where c.ԭ��id = a.Id)
         And a.���� Is Not Null
      Union
      Select m.*
        From (Select Rawtohex(b.Id) As ID,
                     b.����,
                     b.˵��,
                     b.���� �ϼ�id,
                     'antetype' As Type,
                     '' As ����,
                     '' As ѧ��,
                     Null �޸�ʱ��,
                     '' As ��ǩl,
                     0 As Private,
                     0 As Imgindex
                From Ӱ�񱨸�ԭ���嵥 B
               Where b.���� = Kind_In
                 And Exists
               (Select ID From Ӱ�񱨸淶���嵥 C Where c.ԭ��id = b.Id)
               Order By b.����) M
      Union All
      Select n.*
        From (Select /*+rule*/
               Rawtohex(a.Id) As ID,
               a.����,
               a.˵��,
               Rawtohex(a.ԭ��id) As �ϼ�id,
               'sample' As Type,
               a.����,
               a.ѧ��,
               a.���༭ʱ�� As �޸�ʱ��,
               a.��ǩ,
               a.�Ƿ�˽�� As Private,
               Decode(a.�Ƿ�˽��, 1, 2, 1) As Imgindex
                From Ӱ�񱨸淶���嵥 A, Ӱ�񱨸�ԭ���嵥 C
               Where a.ԭ��id = c.Id
                 And c.���� = Kind_In
                 And ((a.���� Like '%' || Condition_In || '%' And
                     Condition_In Is Not Null) Or Condition_In Is Null)
                 And (a.���� = Author_In Or (a.ѧ�� Is Null And a.�Ƿ�˽�� = 0) Or
                     Subjects_In Is Null Or
                     (a.ѧ�� Is Not Null And
                     b_pacs_rptpublic.f_If_Intersect(a.ѧ��, Subjects_In) > 0 And
                     a.�Ƿ�˽�� = 0))
               Order By a.���, a.����) N;
  
  End p_Get_Samplelist_By_Kind;

  --ͨ��ID������Ӧ�ķ�����Ϣ
  Procedure p_Get_Samplelist_By_Id(
    Val Out t_Refcur,
	Id_In Ӱ�񱨸淶���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID,
             Rawtohex(a.ԭ��id) As ԭ��id,
             a.���,
             a.����,
             a.˵��,
             a.ѧ��,
             a.��ǩ,
             a.�Ƿ�˽��,
             a.����,
             a.���༭ʱ�� Lasttime
        From Ӱ�񱨸淶���嵥 A
       Where a.Id = Hextoraw(Id_In);
  
  End p_Get_Samplelist_By_Id;

  --��ѯ��Ӧԭ���µ�������
  Procedure p_Get_Samplelist_Maxseqnum(
    Val Out t_Refcur,
    Aid_In Ӱ�񱨸淶���嵥.ԭ��id%Type
	) As
  Begin
    Open Val For
      Select Nvl(Max(a.���), 0) + 1 As Num
        From Ӱ�񱨸淶���嵥 A
       Where a.ԭ��id = Hextoraw(Aid_In);
  End p_Get_Samplelist_Maxseqnum;

  --��ȡ����XML��Ϣ
  Procedure p_Get_Samplexml(
    Val Out t_Refcur,
	Id_In Ӱ�񱨸淶���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select A.����.getclobval() ���� From Ӱ�񱨸淶���嵥 A Where a.Id = Id_In;
  End p_Get_Samplexml;

  --�޸ķ���XML��Ϣ
  Procedure p_Edit_Samplexml(
    Id_In      Ӱ�񱨸淶���嵥.Id%Type,
	Content_In Ӱ�񱨸淶���嵥.����%Type
	) As
  Begin
    Update Ӱ�񱨸淶���嵥
       Set ���� = Content_In
     Where ID = Hextoraw(Id_In);
  End p_Edit_Samplexml;

  --�����ķ����б�
  Procedure p_Output_Samplelist(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select *
        From (Select Rawtohex(a.Id) As ID,
                     Null As Aid,
                     a.���� As Antetypename,
                     Null As Code,
                     a.���� As Title,
                     a.˵�� As Note,
                     Null As Content,
                     Null As Subject,
                     Null As Label,
                     Null As Private,
                     Null As Author,
                     Null As Lasttime,
                     '' As Flag,
                     1 As Image,
                     'antetype' As Type
                From Ӱ�񱨸�ԭ���嵥 A
               Where Exists (Select b.Id
                        From Ӱ�񱨸淶���嵥 B
                       Where b.ԭ��id = a.Id)
               Order By a.����)
      Union All
      Select Rawtohex(c.Id) As ID,
             Rawtohex(c.ԭ��id) As Aid,
             d.����,
             c.���,
             c.����,
             c.˵��,
             '',
             c.ѧ��,
             c.��ǩ,
             c.�Ƿ�˽��,
             c.����,
             c.���༭ʱ��,
             '',
             0,
             'sample'
        From Ӱ�񱨸淶���嵥 C, Ӱ�񱨸�ԭ���嵥 D
       Where c.ԭ��id = d.Id;
  
  End p_Output_Samplelist;

  --�Ƿ������Ӧ��ԭ�����
  Procedure p_If_Exist_Antetypelist(
    Val Out t_Refcur,
	Title_In Ӱ�񱨸�ԭ���嵥.����%Type
	) As
  Begin
    Open Val For
      Select Count(a.Id) Num
        From Ӱ�񱨸�ԭ���嵥 A
       Where a.���� = Title_In;
  End p_If_Exist_Antetypelist;

  --ͬһ��������Ƿ������ͬ���Ƶķ���
  Procedure p_If_Exist_Samplelist(
    Val Out t_Refcur,
    Type_In  Ӱ�񱨸�ԭ���嵥.����%Type,
    Title_In Ӱ�񱨸淶���嵥.����%Type
	) As
  Begin
    Open Val For
      Select Count(a.Id) Num, Max(a.Id) ID
        From Ӱ�񱨸淶���嵥 A, Ӱ�񱨸�ԭ���嵥 B
       Where a.ԭ��id = b.Id
         And a.���� = Title_In
         And b.���� = Type_In;
  End p_If_Exist_Samplelist;
  
  --ͨ������ID��÷��Ķ�Ӧ�����νṹ
  Procedure p_Get_SamplelistTree_By_Id(
    Val Out t_Refcur,
    Id_In Ӱ�񱨸淶���嵥.Id%Type
	) As
  Begin
    --'EE7CD4A510B045A9BBE6D8CC7DB6EE30'
    Open Val For
      Select RawToHex(t.ID) as ID,
             t.����,
             T.˵��,
             RawToHex(T.ԭ��ID) ԭ��ID,
             'sample' ����,
             t.����,
             t.ѧ��,
             t.���༭ʱ��,
             t.��ǩ,
             t.�Ƿ�˽��,
             2 IMGINDEX
        From Ӱ�񱨸淶���嵥 t
       Where t.id = Id_In
      Union All
      Select RawToHex(x.id) as ID,
             x.����,
             x.˵��,
             x.���� ԭ��ID,
             'antetype' ����,
             null ����,
             null ѧ��,
             null ���༭ʱ��,
             null ��ǩ,
             0 �Ƿ�˽��,
             0 IMGINDEX
        From Ӱ�񱨸�ԭ���嵥 x
       Where x.id = (Select t.ԭ��id
                       From Ӱ�񱨸淶���嵥 t
                      where t.id = Id_In
                        and rownum <= 1)
      Union All
      Select x.���� as ID,
             x.���� ����,
             null ˵��,
             null ԭ��ID,
             'category' ����,
             null ����,
             null ѧ��,
             null ���༭ʱ��,
             null ��ǩ,
             0 �Ƿ�˽��,
             0 IMGINDEX
        From Ӱ�񱨸�ԭ���嵥 x
       Where x.ID = (Select t.ԭ��id
                       From Ӱ�񱨸淶���嵥 t
                      where t.id = Id_In
                        and rownum <= 1);
  End p_Get_SamplelistTree_By_Id;
End b_Pacs_RptSampleList;
/



--Ӱ�񱨸�ҵ��(---���岿��---)***************************************************
Create Or Replace Package b_Pacs_Rptmanage Is
  Type t_Refcur Is Ref Cursor;

  --1������������
  Procedure p_Edit_Doc_Lockinfo
  (
    ����_Id_In Ӱ�񱨸��¼.Id%Type,
    ������_In  Ӱ�񱨸��¼.������%Type
  );

  --2��������������
  Procedure p_Edit_Doc_Evaluatrptquality
  (
    ����id_In   Ӱ�񱨸��¼.Id%Type,
    �����ȼ�_In Ӱ�񱨸��¼.��������%Type
  );

  --3������������
  Procedure p_Edit_Doc_Evaluatresult
  (
    ����id_In   Ӱ�񱨸��¼.Id%Type,
    �����_In Ӱ�񱨸��¼.�������%Type
  );

  --4�����淢��/����
  Procedure p_Edit_Doc_Reportrelease
  (
    ����id_In     Ӱ�񱨸��¼.Id%Type,
    ��ǰ������_In Ӱ�񱨸��¼.���淢����%Type
  );

  --5���������޸ı���
  Procedure p_Ӱ�񱨸��¼_����
  (
    ԭ��id_In     Ӱ�񱨸��¼.ԭ��id%Type,
    ��������_In   Ӱ�񱨸��¼.��������%Type,
    ��¼��_In     Ӱ�񱨸��¼.��¼��%Type,
    ���༭��_In Ӱ�񱨸��¼.���༭��%Type,
    Id_In         Ӱ�񱨸��¼.Id%Type,
    ҽ��id_In     Ӱ�񱨸��¼.ҽ��id%Type
  );

  --6����ȡ��д���ĵ�����
  Procedure p_Get_Doc_Content
  (
    Val      Out t_Refcur,
    Docid_In Ӱ�񱨸��¼.Id%Type
  );

  --7�����ñ����ӡ������Ϣ
  Procedure p_Checkrejectsignature
  (
    Signdate_In Date,
    ����id_In   Ӱ�񱨸������¼.����id%Type,
    ������_In   Ӱ�񱨸������¼.������%Type,
    ����˵��_In Ӱ�񱨸������¼.����˵��%Type,
    Val         Out Sys_Refcursor
  );

  --8����ѯ��Ӧԭ���µ�������
  Procedure p_Get_Samplelist_Maxseqnum
  (
    Val       Out t_Refcur,
    ԭ��id_In Ӱ�񱨸淶���嵥.ԭ��id%Type
  );

  --9��ɾ���ĵ�����
  Procedure p_Del_Ӱ�񱨸淶���嵥(Id_In Ӱ�񱨸淶���嵥.Id%Type);

  --10������ĵ��Ĳ�����־
  Procedure p_Ӱ�񱨸������¼_Add
  (
    Id_In       Ӱ�񱨸������¼.Id%Type,
    ����id_In   Ӱ�񱨸������¼.����id%Type,
    ������_In   Ӱ�񱨸������¼.������%Type,
    ��������_In Ӱ�񱨸������¼.��������%Type
  );

  --11��ɾ������
  Procedure p_Ӱ�񱨸��¼_ɾ��(����_Id_In Ӱ�񱨸��¼.Id%Type);

  --12����ȡǩ������
  Procedure p_Get_Sysconfigsignature
  (
    Val       Out t_Refcur,
    ����id_In In ���ű�.Id%Type
  );

  --13����ȡ�˻�ǩ��ӡ��
  Procedure p_Get_Personsignimg
  (
    Val   Out t_Refcur,
    Id_In In ��Ա��.Id%Type
  );

  --14����ȡǩ����֤����Ϣ
  Procedure p_Get_Signcertinfo
  (
    Val       Out t_Refcur,
    ֤��id_In ��Ա֤���¼.Id%Type
  );

  --15�����±���״̬
  Procedure p_Update_Reportstate
  (
    ����id_In   Ӱ�񱨸��¼.Id%Type,
    ����״̬_In Ӱ�񱨸��¼.����״̬%Type,
    �����_In   Ӱ�񱨸��¼.��������%Type
  );

  --16����ȡ����״̬
  Procedure p_Get_Reportstate
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�񱨸��¼.Id%Type
  );

  --17�����沵��
  Procedure p_Reject_Report
  (
    ҽ��id_In   Ӱ�񱨸沵��.ҽ��id%Type,
    ����id_In   Ӱ�񱨸沵��.��鱨��id%Type,
    ��������_In Ӱ�񱨸沵��.��������%Type,
    ����ʱ��_In Ӱ�񱨸沵��.����ʱ��%Type,
    ������_In   Ӱ�񱨸沵��.������%Type,
    ��������_In Ӱ�񱨸��¼.��������%Type,
    ����״̬_In Ӱ�񱨸��¼.����״̬%Type
  );

  --17.1���������沵��
  Procedure p_Reject_Cancel
  (
    Id_In       Ӱ�񱨸沵��.Id%Type,
    ����id_In   Ӱ�񱨸沵��.��鱨��id%Type,
    ����״̬_In Ӱ�񱨸��¼.����״̬%Type
  );

  --18����ȡ���沵����Ϣ
  Procedure p_Get_Rejectinfo
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�񱨸沵��.��鱨��id%Type
  );

  --19����ȡԭ�Ͷ���
  Procedure p_Get_Doc_Process
  (
    Val       Out t_Refcur,
    ԭ��id_In Ӱ�񱨸涯��.ԭ��id%Type
  );

  --20��ͨ��ѧ��ɸѡ�����Ӧ�ķ�����Ϣ
  Procedure p_Get_Samplelist_By_Conditions
  (
    Val          Out t_Refcur,
    ԭ��id_In    Varchar2,
    ѧ��_In      Varchar2,
    Condition_In Varchar2, --����ɸѡ
    ����_In      Varchar2
  );

  --21��ͨ������ID��ȡ��������
  Procedure p_Get_��������_By_Id
  (
    Val   Out t_Refcur,
    Id_In ���ű�.Id%Type
  );

  --22����ȡ����Ԥ�����
  Procedure p_Get_Allpreoutlines(Val Out t_Refcur);

  --23����ȡ�ĵ�����
  Procedure p_Get_Reporttitle_By_Id
  (
    Val   Out t_Refcur,
    Id_In Ӱ�񱨸��¼.Id%Type
  );

  --24����ȡ����������
  Procedure p_Get_����������_By_Id
  (
    Val   Out t_Refcur,
    Id_In Ӱ�񱨸��¼.Id%Type
  );

  --25��ͨ��ҽ��ID��ȡ�����б�
  Procedure p_Get_Ӱ�񱨸��¼_By_ҽ��id
  (
    Val       Out t_Refcur,
    ҽ��id_In Ӱ�񱨸��¼.ҽ��id%Type
  );

  --26����ѯӰ�����̲���ֵ
  Procedure p_Get_Ӱ�����̲���ֵ
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�����̲���.����id%Type
  );

  --27������ҽ��ID����ѯ��Ӧ��ԭ���б�
  Procedure p_Get_Ӱ��ԭ���б�_By_ҽ��id
  (
    Val     Out t_Refcur,
    ҽ��_In Ӱ�����¼.ҽ��id%Type
  );

  --28�����ݱ���ID��ѯ��ӡ��¼
  Procedure p_Get_Reportprintlog_By_����id
  (
    Val     Out Sys_Refcursor,
    ����_In Ӱ�񱨸������¼.����id%Type
  );

  --29������ҽ��ID��ѯ���淢���б�
  Procedure p_Get_Reportreleaselist
  (
    Val     Out t_Refcur,
    ҽ��_In Ӱ�񱨸��¼.ҽ��id%Type
  );

  --30�����ݱ���ID��ѯ���ؼ�¼����
  Procedure p_Get_Rejectedcount
  (
    Val     Out t_Refcur,
    ����_In Ӱ�񱨸沵��.��鱨��id%Type
  );

  --31������ҽ��ID��ѯ���涯����Ҫ��һЩID��
  Procedure p_Get_Docprocess_Ids
  (
    Val     Out t_Refcur,
    ҽ��_In ����ҽ����¼.Id%Type
  );

  --32������ҽ��ID�ͱ���ID��ѯ�����һЩ����
  Procedure p_Get_Docinfo
  (
    Val       Out t_Refcur,
    ҽ��id_In Ӱ�����¼.ҽ��id%Type,
    ����id_In Ӱ�񱨸��¼.Id%Type
  );

  --33����ѯһ���������ͬԭ��ID�ı�������
  Procedure p_Get_Sameantetypedoccounts
  (
    Val       Out t_Refcur,
    ҽ��id_In Ӱ�񱨸��¼.ҽ��id%Type,
    ԭ��id_In Ӱ�񱨸��¼.ԭ��id%Type
  );

  --34����ȡ����ͼ�洢��Ϣ
  Procedure p_Get_Docimagesaveinof_By_Id
  (
    Val   Out t_Refcur,
    Id_In Ӱ�񱨸��¼.Id%Type
  );

  --35���޸�ԭ��ʹ�ô���
  Procedure p_Update_Antetypeusecount(Id_In Ӱ�񱨸�ԭ���嵥.Id%Type);

  --36������Ӱ����ͼ��ı���ͼ���
  Procedure p_Update_Rptimage
  (
    Uid_In        Ӱ����ͼ��.ͼ��uid%Type,
    Actiontype_In Number
  );

  --37����ȡ��ӡ������Ϣ
  Procedure p_Get_Printcontrol
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�񱨸��¼.Id%Type
  );

End b_Pacs_Rptmanage;

/

--Ӱ�񱨸�ҵ��(---ʵ�ֲ���---)***************************************************

Create Or Replace Package Body b_Pacs_Rptmanage Is

  --1������������
  Procedure p_Edit_Doc_Lockinfo
  (
    ����_Id_In Ӱ�񱨸��¼.Id%Type,
    ������_In  Ӱ�񱨸��¼.������%Type
  ) Is
  Begin
  
    --  ����IDΪ�գ���������С�������_In�����������ı��
    If ����_Id_In Is Null Then
      Update Ӱ�񱨸��¼ a Set a.������ = '' Where a.������ = ������_In;
    Else
      Update Ӱ�񱨸��¼ a Set a.������ = ������_In Where a.Id = ����_Id_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Lockinfo;

  --2��������������
  Procedure p_Edit_Doc_Evaluatrptquality
  (
    ����id_In   Ӱ�񱨸��¼.Id%Type,
    �����ȼ�_In Ӱ�񱨸��¼.��������%Type
  ) Is
  Begin
    Update Ӱ�񱨸��¼ Set �������� = �����ȼ�_In Where Id = ����id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Evaluatrptquality;

  --3������������
  Procedure p_Edit_Doc_Evaluatresult
  (
    ����id_In   Ӱ�񱨸��¼.Id%Type,
    �����_In Ӱ�񱨸��¼.�������%Type
  ) Is
  Begin
    Update Ӱ�񱨸��¼ Set ������� = �����_In Where Id = ����id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Evaluatresult;

  --4�����淢��/����
  Procedure p_Edit_Doc_Reportrelease
  (
    ����id_In     Ӱ�񱨸��¼.Id%Type,
    ��ǰ������_In Ӱ�񱨸��¼.���淢����%Type
  ) Is
    v_���淢�� Ӱ�񱨸��¼.���淢��%Type;
  Begin
  
    Begin
      Select Nvl(���淢��, 0) Into v_���淢�� From Ӱ�񱨸��¼ Where Id = ����id_In;
    Exception
      When Others Then
        v_���淢�� := 0;
    End;
  
    Update Ӱ�񱨸��¼
    Set ���淢�� = Decode(v_���淢��, 0, 1, 0), ���淢���� = Decode(v_���淢��, 0, ��ǰ������_In, '')
    Where Id = ����id_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Reportrelease;

  --5���������޸ı���
  Procedure p_Ӱ�񱨸��¼_����
  (
    ԭ��id_In     Ӱ�񱨸��¼.ԭ��id%Type,
    ��������_In   Ӱ�񱨸��¼.��������%Type,
    ��¼��_In     Ӱ�񱨸��¼.��¼��%Type,
    ���༭��_In Ӱ�񱨸��¼.���༭��%Type,
    Id_In         Ӱ�񱨸��¼.Id%Type,
    ҽ��id_In     Ӱ�񱨸��¼.ҽ��id%Type
  ) As
    --ԭ��ID_In ԭ��ID
    --�����ĵ���д��¼
    --1 ������������
    --2 �����ĵ���д��¼��״̬
    --3 ����༭��־
    --4 �����ĵ�����
    v_����id    Ӱ�񱨸��¼.Id%Type;
    v_ԭ������  Ӱ�񱨸�ԭ���嵥.����%Type;
    v_�豸��    Ӱ�񱨸�ԭ���嵥.�豸��%Type;
    v_�������  Number;
    x_Editlog   Xmltype;
    Cur_Time    Date;
    To_Editlist t_Editlist;
    Tn_Editlist t_Editlist;
    v_Msg       Varchar2(200);
    v_New       Number;
    Err_Custom Exception;
    v_Result Ӱ�񱨸��¼.������%Type;
    v_����id Ӱ�񱨸������¼.Id%Type;
  
    Function Elist_Filter(Source_t t_Editlist) Return t_Editlist Is
      Target_t t_Editlist := t_Editlist();
    Begin
    
      --�Զ����ĵ���˵���������ֻ�ǽ� Source_t���ձ༭ʱ����������
      For Rs In (Select /*+rule*/
                  *
                 From Table(Cast(Source_t As t_Editlist)) a
                 Order By a.�༭ʱ��) Loop
        Target_t.Extend;
        Target_t(Target_t.Count) := t_Edits(Rs.�༭��, Rs.�༭ʱ��, Rs.ǩ��, Rs.��ǩ��);
      End Loop;
      Return Target_t;
    End;
  
    Function Build_Editlog
    (
      Tn_Edit t_Editlist,
      To_Edit t_Editlist,
      v_Did   Ӱ�񱨸��¼.Id%Type
    ) Return Xmltype Is
      --Tn_Edit ���α�����±༭��¼��To_Edit�ϴα���ľɱ༭��¼
      --�����α༭��¼����ϳ�һ���༭��¼
    
      x_Return Xmltype;
      r_Saveid Raw(16);
      n_Class  Number;
      --n_Class �༭��־�еĲ������ 1-������2-ɾ����3-�༭��4-ǩ����5-�󶩡�6-��ǩ��7-��ǩ
      v_Signor  Ӱ�񱨸��¼.������%Type;
      v_Adjunct Ӱ�񱨸��¼.������%Type;
      Tns_Edit  t_Editlist;
      Tos_Edit  t_Editlist;
    
      Function Atitle(ԭ��id Ӱ�񱨸�ԭ���嵥.Id%Type) Return Varchar2 Is
        v_ԭ������ Ӱ�񱨸�ԭ���嵥.����%Type;
      Begin
        --����ԭ��ID������ԭ������
        If ԭ��id Is Null Then
          Return Null;
        Else
          Select ���� Into v_ԭ������ From Ӱ�񱨸�ԭ���嵥 Where Id = ԭ��id;
          Return v_ԭ������;
        End If;
      End;
    
    Begin
      x_Return := Xmltype('<root></root>');
      If v_Did Is Null Then
        --�����������ĵ��������ĵ���null����
        Select Sys_Guid() Into r_Saveid From Dual;
      
        --PACS����û�����ĵ����������湹��XML����䱣���ɸ�EMR��ͬ�������v_Subiid��ֵΪ��
        Tns_Edit := Elist_Filter(Tn_Edit);
        Select Decode(Tns_Edit(Tns_Edit.Count).ǩ��, 0, 1, 4) Into n_Class From Dual;
        Select Appendchildxml(x_Return,
                               '/root',
                               Xmlelement("operate",
                                          Xmlforest(r_Saveid As "saving_id",
                                                    n_Class As "class",
                                                    To_Char(Cur_Time, 'yyyy-mm-dd hh24:mi:ss') As "cur_time",
                                                    ���༭��_In As "operator",
                                                    Decode(n_Class, 4, Tns_Edit(Tns_Edit.Count).�༭��, '') As "signer",
                                                    '' As Adjunct)))
        Into x_Return
        From Dual;
      Else
        --�����������ĵ���
        Select Sys_Guid() Into r_Saveid From Dual;
      
        v_Signor  := '';
        v_Adjunct := '';
        Tns_Edit  := Elist_Filter(Tn_Edit);
        Tos_Edit  := Elist_Filter(To_Edit);
        If Tns_Edit(Tns_Edit.Count).ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).��ǩ�� = 0 Then
          --���һ����ǩ��
          If Tos_Edit.Count = 0 Then
            --�������ĵ�ֱ��ǩ��
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).�༭�� Is Null Then
            --֮ǰûǩ��
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).�༭ʱ�� > Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�����ͨǩ��
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).�༭ʱ�� < Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�������ǩ��
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).�༭ʱ�� = Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�ޱ仯
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count).ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).��ǩ�� = 1 Then
          --��ǩ��
          If Tos_Edit(Tos_Edit.Count).��ǩ�� = 0 Then
            --֮ǰû��ǩ����������ǩ��������
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).��ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).�༭ʱ�� > Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�����ǩ
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).��ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).�༭ʱ�� < Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --���������ǩ
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).��ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).�༭ʱ�� = Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�ޱ仯
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count).�༭�� Is Null And Tos_Edit.Count = 0 Then
          n_Class := 1;
        Elsif Tns_Edit(Tns_Edit.Count).�༭�� Is Null And Tos_Edit(Tos_Edit.Count).ǩ�� = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
        Elsif Tns_Edit(Tns_Edit.Count).�༭�� Is Null And Tos_Edit(Tos_Edit.Count).�༭�� Is Null Then
          n_Class := 3;
        Elsif Tns_Edit(Tns_Edit.Count).��ǩ�� = 0 And Tos_Edit(Tos_Edit.Count).��ǩ�� = 0 Then
          n_Class := 5;
        Elsif Tns_Edit(Tns_Edit.Count).��ǩ�� = 0 And Tos_Edit(Tos_Edit.Count).��ǩ�� = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
        End If;
      
        If n_Class <> -1 Then
          Select Appendchildxml(x_Return,
                                 '/root',
                                 Xmlelement("operate",
                                            Xmlforest(r_Saveid As "saving_id",
                                                      n_Class As "class",
                                                      To_Char(Cur_Time, 'yyyy-mm-dd hh24:mi:ss') As "cur_time",
                                                      ���༭��_In As "operator",
                                                      Decode(n_Class, 4, v_Signor, 6, v_Signor, '') As "signer",
                                                      v_Adjunct As Adjunct)))
          Into x_Return
          From Dual;
        End If;
      
      End If;
      Return x_Return;
    End Build_Editlog;
  
    Function Get_Nextrptnum
    (
      Antetypename Ӱ�񱨸�ԭ���嵥.����%Type,
      Order_Id     Ӱ�񱨸��¼.ҽ��id%Type
    ) Return Number Is
      v_���  Number;
      v_Count Number;
      v_Num   Number;
    Begin
    
      v_Count := 0;
      v_Num   := 1;
      Loop
        Select Count(*) + v_Num Into v_��� From Ӱ�񱨸��¼ Where ҽ��id = Order_Id;
        Select Count(*)
        Into v_Count
        From Ӱ�񱨸��¼
        Where ҽ��id = Order_Id
        And �ĵ����� = Antetypename || '_' || v_���;
      
        If v_Count = 0 Then
          Exit;
        End If;
      
        v_Num := v_Num + 1;
      End Loop;
    
      Return v_���;
    End;
  
  Begin
  
    Select ����, �豸��, Sysdate Into v_ԭ������, v_�豸��, Cur_Time From Ӱ�񱨸�ԭ���嵥 Where Id = ԭ��id_In;
  
    --------------------1 �����ĵ���д��¼��״̬--------------------
    --��ȡ�ĵ���ǩ���ͱ༭���������޸ģ���¼
    Tn_Editlist := b_Pacs_Rptpublic.f_Geteditlist(��������_In);
  
    --------------------2 ����༭��־--------------------
    Select Count(*) Into v_New From Ӱ�񱨸��¼ Where Id = Id_In;
  
    v_����id := Id_In;
    Select Zlpub_Pacs_ȡ�������byxml(��������_In, '������') Into v_Result From Dual;
    If v_New = 0 Then
      --��������
      To_Editlist := t_Editlist();
      x_Editlog   := Build_Editlog(Tn_Editlist, To_Editlist, Null);
    
      --ȡ�������
      v_������� := Get_Nextrptnum(v_ԭ������, ҽ��id_In);
    
      Insert Into Ӱ�񱨸��¼
        (Id, ԭ��id, �ĵ�����, ��������, ����ʱ��, ������, ����״̬, ���༭ʱ��, ���༭��, �༭��־, ҽ��id, ��¼��, ������, �豸��)
      Values
        (v_����id, ԭ��id_In, v_ԭ������ || '_' || v_�������, ��������_In, Cur_Time, ���༭��_In, 1, Cur_Time, ���༭��_In, x_Editlog,
         ҽ��id_In, ��¼��_In, v_Result, v_�豸��);
      Insert Into ����ҽ������ (ҽ��id, ��鱨��id) Values (ҽ��id_In, v_����id);
    
      Select Sys_Guid() Into v_����id From Dual;
      Insert Into Ӱ�񱨸������¼
        (Id, ����id, ҽ��id, �ĵ�����, ������, ����ʱ��, ��������)
      Values
        (v_����id, v_����id, ҽ��id_In, v_ԭ������ || '_' || v_�������, ���༭��_In, Sysdate, 6);
    
    Else
      --��ȡ�ļ�ԭʼ�༭��¼,�����ڸ���֮ǰ��ȡ
      Select b_Pacs_Rptpublic.f_Geteditlist(��������) Into To_Editlist From Ӱ�񱨸��¼ Where Id = v_����id;
    
      x_Editlog := Build_Editlog(Tn_Editlist, To_Editlist, v_����id);
      Select Appendchildxml(�༭��־, '/root', Extract(x_Editlog, '/root/*'))
      Into x_Editlog
      From Ӱ�񱨸��¼
      Where Id = v_����id;
    
      Update Ӱ�񱨸��¼
      Set �������� = ��������_In, ���༭ʱ�� = Cur_Time, ���༭�� = ���༭��_In, �༭��־ = x_Editlog, ��¼�� = ��¼��_In, ������ = v_Result
      Where Id = v_����id;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Ӱ�񱨸��¼_����;

  --6����ȡ��д���ĵ�����
  Procedure p_Get_Doc_Content
  (
    Val      Out t_Refcur,
    Docid_In Ӱ�񱨸��¼.Id%Type
  ) As
  Begin
    Open Val For
      Select (Nvl(a.��������, Xmltype('<ZLXML/>'))).Getclobval() As �������� From Ӱ�񱨸��¼ a Where a.Id = Docid_In;
  End;

  --7�����ñ����ӡ������Ϣ
  Procedure p_Checkrejectsignature
  (
    Signdate_In Date,
    ����id_In   Ӱ�񱨸������¼.����id%Type,
    ������_In   Ӱ�񱨸������¼.������%Type,
    ����˵��_In Ӱ�񱨸������¼.����˵��%Type,
    Val         Out Sys_Refcursor
  ) As
  Begin
    Open Val For
      Select ������, ����ʱ��
      From Ӱ�񱨸������¼
      Where ����id = ����id_In
      And �������� = 1
      And ����ʱ�� >= Signdate_In
      And ����ʱ�� Is Null
      Order By ����ʱ�� Asc;
    --���ϴ�ӡ��¼
    Update Ӱ�񱨸������¼ b
    Set ������ = ������_In, ����ʱ�� = Sysdate, b.����˵�� = ����˵��_In
    Where ����id = ����id_In
    And �������� = 1
    And ����ʱ�� >= Signdate_In;
  
  End p_Checkrejectsignature;

  --8����ѯ��Ӧԭ���µ�������
  Procedure p_Get_Samplelist_Maxseqnum
  (
    Val       Out t_Refcur,
    ԭ��id_In Ӱ�񱨸淶���嵥.ԭ��id%Type
  ) As
  Begin
    Open Val For
      Select Nvl(Max(a.���), 0) + 1 As Num From Ӱ�񱨸淶���嵥 a Where a.ԭ��id = ԭ��id_In;
  End;

  --9��ɾ���ĵ�����
  Procedure p_Del_Ӱ�񱨸淶���嵥(Id_In Ӱ�񱨸淶���嵥.Id%Type) As
  Begin
    Delete From Ӱ�񱨸淶���嵥 Where Id = Id_In;
  End;

  --10������ĵ��Ĳ�����־
  Procedure p_Ӱ�񱨸������¼_Add
  (
    Id_In       Ӱ�񱨸������¼.Id%Type,
    ����id_In   Ӱ�񱨸������¼.����id%Type,
    ������_In   Ӱ�񱨸������¼.������%Type,
    ��������_In Ӱ�񱨸������¼.��������%Type
  ) As
    n_ҽ��id   Ӱ�񱨸������¼.ҽ��id%Type;
    n_�ĵ����� Ӱ�񱨸��¼.�ĵ�����%Type;
  Begin
  
    Begin
      Select ҽ��id, �ĵ����� Into n_ҽ��id, n_�ĵ����� From Ӱ�񱨸��¼ Where Id = ����id_In;
    Exception
      When Others Then
        Null;
    End;
    If n_ҽ��id Is Not Null Then
      Insert Into Ӱ�񱨸������¼
        (Id, ����id, ҽ��id, �ĵ�����, ������, ����ʱ��, ��������)
      Values
        (Id_In, ����id_In, n_ҽ��id, n_�ĵ�����, ������_In, Sysdate, ��������_In);
      If ��������_In = 1 Then
        Update Ӱ�񱨸��¼ Set �����ӡ = 1 Where Id = ����id_In;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --11��ɾ������
  Procedure p_Ӱ�񱨸��¼_ɾ��(����_Id_In Ӱ�񱨸��¼.Id%Type) As
  Begin
  
    Delete From Ӱ�񱨸��¼ Where Ӱ�񱨸��¼.Id = Hextoraw(����_Id_In);
  
    Delete From ����ҽ������ Where ��鱨��id = Hextoraw(����_Id_In);
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Ӱ�񱨸��¼_ɾ��;

  --12����ȡǩ������
  Procedure p_Get_Sysconfigsignature
  (
    Val       Out t_Refcur,
    ����id_In In ���ű�.Id%Type
  ) Is
  Begin
    --�����û�, ģ���,����
    Open Val For
      Select Zl_Fun_Getsignpar(7, ����id_In) As ǩ������ From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --13����ȡ�˻�ǩ��ӡ��
  Procedure p_Get_Personsignimg
  (
    Val   Out t_Refcur,
    Id_In In ��Ա��.Id%Type
  ) Is
    --v_sql Varchar2(1000);
    --n_count Number(5);
  Begin
    --Select Count(*) Into n_Count From user_tables Where table_name =Upper('Ӱ��ǩ��ͼƬ');
  
    --If n_Count > 0 Then
    --   v_sql := 'Truncate Table Ӱ��ǩ��ͼƬ';
    --   Execute Immediate v_sql;
  
    --   v_sql := 'Insert Into Ӱ��ǩ��ͼƬ Select a.id, to_lob(a.ǩ��ͼƬ) as ǩ��ͼƬ From ��Ա�� a Where a.ID=' || ID_In;
    --   Execute Immediate v_sql;
    --Else
    --   v_sql := 'Create GLOBAL TEMPORARY TABLE Ӱ��ǩ��ͼƬ ON COMMIT PRESERVE ROWS AS Select a.id, to_lob(a.ǩ��ͼƬ) as ǩ��ͼƬ From ��Ա�� a Where a.ID=' || ID_In;
    --   Execute Immediate v_sql;
    --End If;
  
    --v_sql := 'Select ǩ��ͼƬ From Ӱ��ǩ��ͼƬ Where Id=:ID';
  
    ----�����û�, ģ���,����
    --Open  Val For v_sql Using ID_In;
  
    Open Val For
      Select ǩ��ͼƬ From ��Ա�� Where Id = Id_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --14����ȡǩ����֤����Ϣ
  Procedure p_Get_Signcertinfo
  (
    Val       Out t_Refcur,
    ֤��id_In ��Ա֤���¼.Id%Type
  ) Is
  Begin
    Open Val For
      Select Id, Certdn, Certsn, Signcert, Enccert From ��Ա֤���¼ Where Id = ֤��id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --15�����±���״̬
  Procedure p_Update_Reportstate
  (
    ����id_In   Ӱ�񱨸��¼.Id%Type,
    ����״̬_In Ӱ�񱨸��¼.����״̬%Type,
    �����_In   Ӱ�񱨸��¼.��������%Type
  ) Is
  Begin
    --����״̬1-δǩ����2-����ϣ�3-����ˣ�4-������5-��ϲ��أ�6-��˲���
    --�������״̬��1-δǩ����2-�����;5-��ϲ��أ���ʱ��û������˵�
    If (����״̬_In = 1) Or (����״̬_In = 2) Or (����״̬_In = 5) Then
      Update Ӱ�񱨸��¼ Set ����״̬ = ����״̬_In, �������� = Null, ������ʱ�� = Null Where Id = ����id_In;
    Elsif (����״̬_In = 3) Or (����״̬_In = 4) Then
      Update Ӱ�񱨸��¼
      Set ����״̬ = ����״̬_In, �������� = �����_In, ������ʱ�� = Sysdate
      Where Id = ����id_In;
    Else
      Update Ӱ�񱨸��¼ Set ����״̬ = ����״̬_In Where Id = ����id_In;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --16����ȡ����״̬
  Procedure p_Get_Reportstate
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�񱨸��¼.Id%Type
  ) Is
  Begin
    Open Val For
      Select ����״̬ From Ӱ�񱨸��¼ Where Id = ����id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --17�����沵��
  Procedure p_Reject_Report
  (
    ҽ��id_In   Ӱ�񱨸沵��.ҽ��id%Type,
    ����id_In   Ӱ�񱨸沵��.��鱨��id%Type,
    ��������_In Ӱ�񱨸沵��.��������%Type,
    ����ʱ��_In Ӱ�񱨸沵��.����ʱ��%Type,
    ������_In   Ӱ�񱨸沵��.������%Type,
    ��������_In Ӱ�񱨸��¼.��������%Type,
    ����״̬_In Ӱ�񱨸��¼.����״̬%Type
  ) Is
  Begin
    Insert Into Ӱ�񱨸沵��
      (Id, ҽ��id, ��鱨��id, ��������, ����ʱ��, ������)
    Values
      (Ӱ�񱨸沵��_Id.Nextval, ҽ��id_In, ����id_In, ��������_In, ����ʱ��_In, ������_In);
  
    Update Ӱ�񱨸��¼ Set ����״̬ = ����״̬_In, �������� = ��������_In Where Id = ����id_In;
  
    --Update ����ҽ������ Set ִ�й���=-1 Where ҽ��ID= ҽ��ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --17.1���������沵��
  Procedure p_Reject_Cancel
  (
    Id_In       Ӱ�񱨸沵��.Id%Type,
    ����id_In   Ӱ�񱨸沵��.��鱨��id%Type,
    ����״̬_In Ӱ�񱨸��¼.����״̬%Type
  ) Is
  Begin
    Update Ӱ�񱨸沵�� Set �Ƿ��� = 1 Where Id = Id_In;
    Update Ӱ�񱨸��¼ Set ����״̬ = ����״̬_In, �������� = '' Where Id = ����id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --18����ȡ���沵����Ϣ
  Procedure p_Get_Rejectinfo
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�񱨸沵��.��鱨��id%Type
  ) Is
  Begin
    Open Val For
      Select a.Id, a.��������, a.����ʱ��, a.������, Nvl(a.�Ƿ���, 0) As ����״̬, b.����״̬
      From Ӱ�񱨸沵�� a, Ӱ�񱨸��¼ b
      Where a.��鱨��id = ����id_In
      And a.��鱨��id = b.Id
      Order By ����ʱ��;
  End;

  --19����ȡԭ�Ͷ���
  Procedure p_Get_Doc_Process
  (
    Val       Out t_Refcur,
    ԭ��id_In Ӱ�񱨸涯��.ԭ��id%Type
  ) As
  Begin
    Open Val For
      Select Rawtohex(p.Id) Id, p.���� As ��������, e.���� As �¼�����, e.���� As �¼�����, e.Ԫ��iid As Ԫ��iid, p.��������, p.���, p.˵��, p.�ɷ��ֹ�ִ��,
             (Nvl(p.����, Xmltype('<NULL/>'))).Getclobval() As ����, Rawtohex(p.�¼�id) �¼�id
      From Ӱ�񱨸涯�� p, Ӱ�񱨸��¼� e
      Where p.�¼�id = e.Id(+)
      And p.ԭ��id = ԭ��id_In
      Order By ��������, ���;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Process;

  --20��ͨ��ѧ��ɸѡ�����Ӧ�ķ�����Ϣ
  Procedure p_Get_Samplelist_By_Conditions
  (
    Val          Out t_Refcur,
    ԭ��id_In    Varchar2,
    ѧ��_In      Varchar2,
    Condition_In Varchar2, --����ɸѡ
    ����_In      Varchar2
  ) As
  Begin
  
    Open Val For
      Select /*+ rule*/
       Rawtohex(a.Id) Id, a.����, a.����, a.˵��, Nvl2(a.˵��, a.˵�� || '����:' || a.����, '����:' || a.����) Content, a.��ǩ, a.ѧ��
      From Ӱ�񱨸淶���嵥 a
      Where a.ԭ��id = Hextoraw(ԭ��id_In)
      And ((a.ѧ�� Is Null And a.�Ƿ�˽�� = 0) Or ѧ��_In Is Null Or a.���� = ����_In Or
            (a.ѧ�� Is Not Null And b_Pacs_Rptpublic.f_If_Intersect(a.ѧ��, ѧ��_In) > 0 And a.�Ƿ�˽�� = 0))
      And (Condition_In Is Null Or
            (a.��ǩ Is Not Null And Condition_In Is Not Null And b_Pacs_Rptpublic.f_If_Intersect(a.��ǩ, Condition_In) > 0))
      Order By a.���;
  
  End p_Get_Samplelist_By_Conditions;

  --21��ͨ������ID��ȡ��������
  Procedure p_Get_��������_By_Id
  (
    Val   Out t_Refcur,
    Id_In ���ű�.Id%Type
  ) Is
  Begin
    Open Val For
      Select ���� From ���ű� Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_��������_By_Id;

  --22����ȡ����Ԥ�����
  Procedure p_Get_Allpreoutlines(Val Out t_Refcur) Is
  Begin
    Open Val For
      Select Rawtohex(Id) Id, a.����, a.���� From Ӱ�񱨸�Ԥ����� a Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Allpreoutlines;

  --23����ȡ�ĵ�����
  Procedure p_Get_Reporttitle_By_Id
  (
    Val   Out t_Refcur,
    Id_In Ӱ�񱨸��¼.Id%Type
  ) Is
  Begin
    Open Val For
      Select �ĵ����� From Ӱ�񱨸��¼ Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Reporttitle_By_Id;

  --24����ȡ����������
  Procedure p_Get_����������_By_Id
  (
    Val   Out t_Refcur,
    Id_In Ӱ�񱨸��¼.Id%Type
  ) Is
  Begin
    Open Val For
      Select ������ From Ӱ�񱨸��¼ Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_����������_By_Id;

  --25��ͨ��ҽ��ID��ȡ�����б�
  Procedure p_Get_Ӱ�񱨸��¼_By_ҽ��id
  (
    Val       Out t_Refcur,
    ҽ��id_In Ӱ�񱨸��¼.ҽ��id%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(Id) As Reportid, Rawtohex(ԭ��id) As Antetypeid, ҽ��id As Orderid, �ĵ����� As Reportname,
             ����ʱ�� As Reportdate,
             Decode(Nvl(����״̬, 0), 1, '�༭��', 2, '�����', 3, '�����', 4, '������', 5, '��ϲ���', '��˲���') As Reportstate,
             ������ As Createuser, ������ʱ�� As Examineydate, �������� As Examineyuser,
             Decode(Nvl(�������, 0), 1, '����', '') As Resultpositive, Nvl(��������, 0) As Innerquality, ' ' As Reportquality,
             Decode(Nvl(�����ӡ, 0), 0, 'δ��ӡ', '�Ѵ�ӡ') As Reportprint,
             Decode(Nvl(���淢��, 0), 0, 'δ����', '�ѷ���') As Reportrelease, ��¼�� As Recdoctor, ������ As RecLocker, ' ' As Locked
      From Ӱ�񱨸��¼
      Where ҽ��id = ҽ��id_In
      Order By Reportdate Desc;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ӱ�񱨸��¼_By_ҽ��id;

  --26����ѯӰ�����̲���ֵ
  Procedure p_Get_Ӱ�����̲���ֵ
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�����̲���.����id%Type
  ) Is
  Begin
    Open Val For
      Select ������, ����ֵ From Ӱ�����̲��� Where ����id = ����id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ӱ�����̲���ֵ;

  --27������ҽ��ID����ѯ��Ӧ��ԭ���б�
  Procedure p_Get_Ӱ��ԭ���б�_By_ҽ��id
  (
    Val     Out t_Refcur,
    ҽ��_In Ӱ�����¼.ҽ��id%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(c.Id) As Antetypeid, c.���� As Antetypename, c.˵��
      From ����ҽ����¼ a, Ӱ�񱨸�ԭ��Ӧ�� b, Ӱ�񱨸�ԭ���嵥 c
      Where a.Id = ҽ��_In
      And a.������Ŀid = b.������Ŀid
      And b.����ԭ��id = c.Id
      And a.������Դ = b.Ӧ�ó���
      Order By c.ʹ�ô��� Desc;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ӱ��ԭ���б�_By_ҽ��id;

  --28�����ݱ���ID��ѯ��ӡ��¼
  Procedure p_Get_Reportprintlog_By_����id
  (
    Val     Out Sys_Refcursor,
    ����_In Ӱ�񱨸������¼.����id%Type
  ) Is
  Begin
    Open Val For
      Select c.�ĵ�����, b.������, To_Char(b.����ʱ��, 'yyyy-MM-dd HH24:mi') ��ӡʱ��, b.������,
             To_Char(b.����ʱ��, 'yyyy-MM-dd HH24:mi') ����ʱ��, b.����˵��
      From Ӱ�񱨸������¼ b, Ӱ�񱨸��¼ c
      Where c.Id = ����_In
      And b.����id = c.Id
      And �������� = 1
      Order By b.����ʱ��;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Reportprintlog_By_����id;

  --29������ҽ��ID��ѯ���淢���б�
  Procedure p_Get_Reportreleaselist
  (
    Val     Out t_Refcur,
    ҽ��_In Ӱ�񱨸��¼.ҽ��id%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(Id) As ����id, �ĵ����� As ��������, ���༭ʱ�� As ��������, Decode(Nvl(���淢��, 0), 0, 'δ����', '�ѷ���') As ���淢��
      From Ӱ�񱨸��¼
      Where ����״̬ Between 2 And 4
      And ҽ��id = ҽ��_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Reportreleaselist;

  --30�����ݱ���ID��ѯ���ؼ�¼����
  Procedure p_Get_Rejectedcount
  (
    Val     Out t_Refcur,
    ����_In Ӱ�񱨸沵��.��鱨��id%Type
  ) Is
  Begin
    Open Val For
      Select Count(*) As �������� From Ӱ�񱨸沵�� Where ��鱨��id = ����_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Rejectedcount;

  --31������ҽ��ID��ѯ���涯����Ҫ��һЩID��
  Procedure p_Get_Docprocess_Ids
  (
    Val     Out t_Refcur,
    ҽ��_In ����ҽ����¼.Id%Type
  ) Is
  Begin
    Open Val For
      Select Id As ҽ��id, ��ҳid, �Һŵ� From ����ҽ����¼ Where Id = ҽ��_In;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Docprocess_Ids;

  --32������ҽ��ID�ͱ���ID��ѯ�����һЩ����
  Procedure p_Get_Docinfo
  (
    Val       Out t_Refcur,
    ҽ��id_In Ӱ�����¼.ҽ��id%Type,
    ����id_In Ӱ�񱨸��¼.Id%Type
  ) Is
  Begin
    If ����id_In Is Null Then
      Open Val For
        Select ִ�п���id, '������' As ������ From Ӱ�����¼ Where ҽ��id = ҽ��id_In;
    Else
      Open Val For
        Select ִ�п���id, ������
        From Ӱ�����¼ a, Ӱ�񱨸��¼ b
        Where a.ҽ��id = b.ҽ��id
        And b.Id = ����id_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Docinfo;

  --33����ѯһ���������ͬԭ��ID�ı�������
  Procedure p_Get_Sameantetypedoccounts
  (
    Val       Out t_Refcur,
    ҽ��id_In Ӱ�񱨸��¼.ҽ��id%Type,
    ԭ��id_In Ӱ�񱨸��¼.ԭ��id%Type
  ) Is
  Begin
    Open Val For
      Select Count(Id) As Doccounts
      From Ӱ�񱨸��¼
      Where ҽ��id = ҽ��id_In
      And ԭ��id = ԭ��id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Sameantetypedoccounts;

  --34����ȡ����ͼ�洢��Ϣ
  Procedure p_Get_Docimagesaveinof_By_Id
  (
    Val   Out t_Refcur,
    Id_In Ӱ�񱨸��¼.Id%Type
  ) Is
  Begin
    Open Val For
      Select �豸��, ����ʱ�� From Ӱ�񱨸��¼ Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Docimagesaveinof_By_Id;

  --35���޸�ԭ��ʹ�ô���
  Procedure p_Update_Antetypeusecount(Id_In Ӱ�񱨸�ԭ���嵥.Id%Type) Is
  Begin
    Update Ӱ�񱨸�ԭ���嵥 Set ʹ�ô��� = ʹ�ô��� + 1 Where Id = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Antetypeusecount;

  --36������Ӱ����ͼ��ı���ͼ���
  Procedure p_Update_Rptimage
  (
    Uid_In        Ӱ����ͼ��.ͼ��uid%Type,
    Actiontype_In Number
  ) Is
    v_Sql Varchar2(4000);
    No_Column Exception;
    Pragma Exception_Init(No_Column, -00904);
  Begin
    If Actiontype_In = 1 Then
      v_Sql := 'Update Ӱ����ͼ�� Set ����ͼ = Nvl(����ͼ, 0) + 1 Where ͼ��uid = :1';
    Else
      v_Sql := 'Update Ӱ����ͼ��
      Set ����ͼ = Decode(����ͼ, Null, Null, Decode(Nvl(����ͼ, 0) - 1, 0, Null, Nvl(����ͼ, 0) - 1))
      Where ͼ��uid = :1';
    End If;
    Execute Immediate v_Sql
      Using Uid_In;
  Exception
    When No_Column Then
      --���ݴ���10.36������ ����ͼ �ֶΣ������ 103996
      Null;
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Rptimage;

  --37����ȡ��ӡ������Ϣ
  Procedure p_Get_Printcontrol
  (
    Val       Out t_Refcur,
    ����id_In Ӱ�񱨸��¼.Id%Type
  ) Is
    v_����     Number;
    v_��ӡ���� Number;
  Begin
  
    Select Nvl(Decode(a.����, 1, 1, b.������־), 0) As ����
    Into v_����
    From ���˹Һż�¼ a, ����ҽ����¼ b
    Where a.No(+) = b.�Һŵ�
    And b.Id = (Select c.ҽ��id From Ӱ�񱨸��¼ c Where c.Id = ����id_In);
  
    Select Nvl(Extractvalue(b.Column_Value, '/root/print_limit'), 0) Printlimit
    Into v_��ӡ����
    From Ӱ�񱨸�ԭ���嵥 a, Table(Xmlsequence(Extract(a.����ѡ��, '/root'))) b, Ӱ�񱨸��¼ c
    Where a.Id = c.ԭ��id
    And c.Id = ����id_In;
  
    Open Val For
      Select v_���� As ����, v_��ӡ���� As ��ӡ���� From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Printcontrol;

End b_Pacs_Rptmanage;

/


--Ӱ�񱨸�������(---���岿��---)***************************************************
Create Or Replace Package b_Pacs_Rptpluginoriginal Is
  Type t_Refcur Is Ref Cursor;

  -- 1����    �ܣ���ȡ��ʷ�����¼
  Procedure p_Getreporthistory
  (
    Val                   Out t_Refcur,
    ҽ��id_In             In ����ҽ����¼.Id%Type,
    ��Աid_In             In ������Ա.��Աid%Type,
    ��ǰ����id_In         In ������Ա.����id%Type,
    �鿴��������ʷ����_In In Number := 0
  );

  --2����    �ܣ���ȡ��Ӧ��������
  Procedure p_Getreportcontent
  (
    Val           Out t_Refcur,
    ����id_In     In Varchar2,
    Editortype_In Number := 0 --0:PACS����༭����1--���Ӳ����༭����2--�����ĵ��༭��
  );

  --3����    �ܣ�����ҽ��ID��ȡ�����Ϣ
  Procedure p_Getstudyinfobyadviceid
  (
    Val       Out t_Refcur,
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type
  );

  --4����    �ܣ���ȡ����ͼ������
  Procedure p_Getreportimagecount
  (
    Val         Out t_Refcur,
    ��ѯ����_In In Varchar2
  );

  --5����    �ܣ���ȡ����ͼ������
  Procedure p_Getreportimagedata
  (
    Val         Out t_Refcur,
    ��ѯ����_In In Varchar2,
    ��ʼλ��_In In Number,
    ����λ��_In In Number
  );

  --6����    �ܣ���ȡԤ��ͼ������
  Procedure p_Getstudyimagecount
  (
    Val         Out t_Refcur,
    ��ѯ����_In In Varchar2,
    �Ƿ���ʱ_In In Number := 0
  );

  --7����    �ܣ���ȡԤ��ͼ������
  Procedure p_Getstudyimagedata
  (
    Val         Out t_Refcur,
    ��ѯ��ʽ_In In Varchar2,
    ��ѯ����_In In Varchar2,
    ��ʼλ��_In In Number,
    ����λ��_In In Number,
    �Ƿ���ʱ_In In Number
  );

  --8�����ܣ���ȡ��ʱͼ������
  Procedure p_Get_Tempimageseries
  (
    Val         Out t_Refcur,
    ʱ�䷶Χ_In In Number,
    ����_In     In Ӱ����ʱ��¼.����%Type := Null
  );

  --9������;��ȡͼ��ע
  Procedure p_Get_Normalnote(Val Out t_Refcur);

  --10�����ܣ����볣��ͼ��ע
  Procedure p_Insert_Normalnote
  (
    Note_In In Ӱ���ֵ�����.����%Type,
    Code_In Ӱ���ֵ�����.����%Type
  );

  --11�����ܣ��޸ĳ���ͼ��ע
  Procedure p_Edit_Normalnote
  (
    Note_In In Ӱ���ֵ�����.����%Type,
    Num_In  Ӱ���ֵ�����.���%Type
  );

  --12�����ܣ�ɾ������ͼ��ע
  Procedure p_Del_Normalnote(Num_In Ӱ���ֵ�����.���%Type);

  --13�����ܣ���ȡ��ע����һ������
  Procedure p_Get_Normalnum(Val Out t_Refcur);
  --14�����ܣ���ȡ���ID
  Procedure p_Get_Plugid
  (
    Val     Out t_Refcur,
    ����_In In Ӱ�񱨸���.����%Type
  );

  --15�����ܣ�����༭���������
  Procedure p_Setfontparam
  (
    Font_In Nvarchar2,
    User_In Nvarchar2
  );

  --16�����ܣ���ȡ�༭���������
  Procedure p_Getfontparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  );

  --17�����ܣ�����༭���������
  Procedure p_Setformparam
  (
    Form_In Nvarchar2,
    User_In Nvarchar2
  );

  --18�����ܣ���ȡ�༭���������
  Procedure p_Getformparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  );

  --19�����ܣ�����ͼ��UID��ȡ�����Ϣ
  Procedure p_Getstudyinfobyimageuid
  (
    Val        Out t_Refcur,
    ҽ��id_In  In Ӱ�����¼.ҽ��id%Type,
    ͼ��uid_In In Ӱ����ͼ��.ͼ��uid%Type
  );

  --20�����ܣ����ݼ��UID��ȡFTP��Ϣ
  Procedure p_Getftpinfobystudyuid
  (
    Val        Out t_Refcur,
    ���uid_In In Ӱ�����¼.���uid%Type
  );

  --21�����ܣ����ݿ���ID��ȡFTP��Ϣ
  Procedure p_Getftpinfobydeptid
  (
    Val       Out t_Refcur,
    ����id_In In Ӱ�����̲���.����id%Type
  );

  --22�����ܣ�����ҽ��ID��ȡFTP��Ϣ
  Procedure p_Getftpinfobyadvicetid
  (
    Val       Out t_Refcur,
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type
  );

  --23�����ܣ���ȡ���UID
  Procedure p_Getstudyuid
  (
    Val        Out t_Refcur,
    ���uid_In In Ӱ�����¼.���uid%Type
  );

  --24�����ܣ���ȡ����UID
  Procedure p_Getseriesuid
  (
    Val        Out t_Refcur,
    ����uid_In In Ӱ��������.����uid%Type
  );

  --25�����ܣ������豸�Ż�ȡ�豸��Ϣ
  Procedure p_Getdeviceinfo
  (
    Val       Out t_Refcur,
    �豸��_In In Ӱ���豸Ŀ¼.�豸��%Type
  );

  --26����ȡҽ��վ�洢�豸��
  Procedure p_Getdeviceidbyadviceid
  (
    Val       Out t_Refcur,
    ҽ��id_In In ����ҽ������.ҽ��id%Type
  );
End b_Pacs_Rptpluginoriginal;

/

--Ӱ�񱨸淶�Ĺ���(---ʵ�ֲ���---)***************************************************
Create Or Replace Package Body b_Pacs_Rptpluginoriginal Is

  --1����    �ܣ���ȡ��ʷ�����¼
  Procedure p_Getreporthistory
  (
    Val                   Out t_Refcur,
    ҽ��id_In             In ����ҽ����¼.Id%Type,
    ��Աid_In             In ������Ա.��Աid%Type,
    ��ǰ����id_In         In ������Ա.����id%Type,
    �鿴��������ʷ����_In In Number := 0
  ) Is
    Strsql     Varchar2(4000);
    Strsqlback Varchar2(4000);
    Strfilter  Varchar2(400);
  Begin
    If �鿴��������ʷ����_In = 1 Then
      Strfilter := ' ';
    Else
      Strfilter := ' And c.ִ�п���id+0 in (select ����id from ������Ա where ��Աid = ' || ��Աid_In ||
                   ' union all select to_Number(' || ��ǰ����id_In || ') from dual) ';
    End If;
  
    Strsql := 'Select 2 as ��������, f.����' || '||''-''||' || 'f.���� As ��������, c.Id As ҽ��id, a.Ӱ����� as ���,b.������ as ������,' ||
              'to_char(b.����ʱ��,''yyyy-mm-dd hh24:mi:ss'') as ����ʱ��,b.�ĵ����� ��������, c.ҽ������, TO_CHAR(RAWTOHEX(b.id)) ����ID ' ||
              'From Ӱ�����¼ A, Ӱ�񱨸��¼ B, ����ҽ����¼ C, Ӱ�����¼ D, ����ҽ����¼ E, ���ű� F ' ||
              'Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And e.Id =' || ҽ��id_In ||
              ' And c.ִ�п���ID = F.ID And b.ҽ��id = c.Id And ' ||
              '(c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null ' || Strfilter || ' union all ' ||
              'Select 1 as ��������, g.����' || '||''-''||' || 'g.���� As ��������, c.Id As ҽ��id, a.Ӱ����� as ���, a.������, ' ||
              'to_char(f.����ʱ��,''yyyy-mm-dd hh24:mi:ss'') as ����ʱ��, a.Ӱ�����||''����'' ��������, c.ҽ������,TO_CHAR( b.����id) as ����ID ' ||
              'From Ӱ�����¼ A, ����ҽ������ B, ����ҽ����¼ C, Ӱ�����¼ D, ����ҽ����¼ E, ���Ӳ�����¼ F, ���ű� G ' ||
              'Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And e.Id = ' || ҽ��id_In ||
              ' And c.ִ�п���ID = g.ID And b.ҽ��id = c.Id And b.����ID Is Not Null And ' ||
              '(c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null And b.����id = f.id ' || Strfilter;
  
    Strsqlback := Strsql;
    Strsqlback := Replace(Strsqlback, 'Ӱ�����¼', 'HӰ�����¼');
    Strsqlback := Replace(Strsqlback, '����ҽ������', 'H����ҽ������');
    Strsqlback := Replace(Strsqlback, '����ҽ����¼', 'H����ҽ����¼');
  
    Strsql := Strsql || ' UNION ALL ' || Strsqlback || ' Order By ����ʱ�� Asc';
  
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreporthistory;

  --2����    �ܣ���ȡ��Ӧ��������
  Procedure p_Getreportcontent
  (
    Val           Out t_Refcur,
    ����id_In     Varchar2,
    Editortype_In Number := 0 --0:���Ӳ����༭����1--PACS����༭����2--�����ĵ��༭��
  ) Is
    Strsql Varchar2(1000);
  Begin
    If Editortype_In = 1 Then
      Strsql := 'Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b ' || 'Where a.�ļ�id = ' ||
                ����id_In || ' And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ֹ��=0 ';
    Elsif Editortype_In = 0 Then
      Strsql := 'select ���� from ���Ӳ�����ʽ where �ļ�ID=' || ����id_In;
    Else
      Strsql := 'Select �������� As ���� From Ӱ�񱨸��¼ Where ID=HexToRaw(''' || ����id_In || ''')';
    End If;
  
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreportcontent;

  --3����    �ܣ�����ҽ��ID��ȡ�����Ϣ
  Procedure p_Getstudyinfobyadviceid
  (
    Val       Out t_Refcur,
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type
  ) Is
    Strsql Varchar2(100);
  Begin
    Strsql := 'Select ���UID,����ͼ��,��������,����,����,�Ա�,���� from Ӱ�����¼ where ҽ��ID =' || ҽ��id_In;
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyinfobyadviceid;

  --4����    �ܣ���ȡ����ͼ������
  Procedure p_Getreportimagecount
  (
    Val         Out t_Refcur,
    ��ѯ����_In In Varchar2
  ) Is
  Begin
    Open Val For
      Select Count(b.Column_Value) ����ֵ
      From Ӱ�����¼ a, Table(Cast(f_Str2list(Replace(a.����ͼ��, ';', ',')) As Zltools.t_Strlist)) b
      Where ҽ��id = ��ѯ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreportimagecount;

  --5����    �ܣ���ȡ����ͼ������
  Procedure p_Getreportimagedata
  (
    Val         Out t_Refcur,
    ��ѯ����_In In Varchar2,
    ��ʼλ��_In In Number,
    ����λ��_In In Number
  ) Is
  Begin
    Open Val For
      Select *
      From (Select Rownum As ˳���, Rownum As ͼ���, b.Ftp�û��� As User1, b.Ftp���� As Pwd1, b.Ip��ַ As Host1,
                    '/' || b.FtpĿ¼ || '/' As Root1,
                    Decode(a.��������, Null, '', To_Char(a.��������, 'YYYYMMDD') || '/') || a.���uid || '/' ||
                     Replace(Trim(d.Column_Value), '.jpg', '') As Url, b.�豸�� As �豸��1, c.Ftp�û��� As User2, c.Ftp���� As Pwd2,
                    c.Ip��ַ As Host2, '/' || c.FtpĿ¼ || '/' As Root2, c.�豸�� As �豸��2,
                    Replace(Trim(d.Column_Value), '.jpg', '') As ͼ��uid, a.���uid, '' ����uid, 0 ��̬ͼ, '' ��������, '' �ɼ�ʱ��,
                    '' ¼�Ƴ���, '' ����ͼ
             From Ӱ�����¼ a, Ӱ���豸Ŀ¼ b, Ӱ���豸Ŀ¼ c, Table(Cast(f_Str2list(Replace(a.����ͼ��, ';', ',')) As Zltools.t_Strlist)) d
             Where a.λ��һ = b.�豸��(+)
             And a.λ�ö� = c.�豸��(+)
             And a.ҽ��id = ��ѯ����_In)
      Where ˳��� >= ��ʼλ��_In
      And ˳��� <= ����λ��_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getreportimagedata;

  --6����    �ܣ���ȡԤ��ͼ������
  Procedure p_Getstudyimagecount
  (
    Val         Out t_Refcur,
    ��ѯ����_In In Varchar2,
    �Ƿ���ʱ_In In Number := 0
  ) Is
    Strsql Varchar2(2000);
  Begin
    If �Ƿ���ʱ_In = 0 Then
      Strsql := 'select T1.����ֵ+T2.����ֵ as ����ֵ from ' || '(select count(1) as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c ' ||
                'where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=''' || ��ѯ����_In || ''') T1,' ||
                '(select count(1) as ����ֵ from HӰ����ͼ�� a, HӰ�������� b, Ӱ�����¼ c ' ||
                'where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=''' || ��ѯ����_In || ''') T2';
    Else
      Strsql := 'select count(1)  as ����ֵ from Ӱ����ʱͼ��  where  ����UID=''' || ��ѯ����_In || '''';
    End If;
  
    Open Val For Strsql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyimagecount;

  --7����    �ܣ���ȡԤ��ͼ������
  Procedure p_Getstudyimagedata
  (
    Val         Out t_Refcur,
    ��ѯ��ʽ_In In Varchar2,
    ��ѯ����_In In Varchar2,
    ��ʼλ��_In In Number,
    ����λ��_In In Number,
    �Ƿ���ʱ_In In Number
  ) Is
    Strsql    Varchar2(2000);
    Strsql2   Varchar2(2000);
    Strfilter Varchar2(100);
    No_Column Exception;
    Pragma Exception_Init(No_Column, -00904);
  Begin
    If ��ѯ��ʽ_In = 0 Then
      Strfilter := 'and c.ҽ��ID=''' || ��ѯ����_In || '''';
    Elsif ��ѯ��ʽ_In = 1 Then
      Strfilter := 'and B.����UID=''' || ��ѯ����_In || '''';
    Else
      Strfilter := 'and A.ͼ��UID=''' || ��ѯ����_In || '''';
    End If;
  
    Strsql := 'Select * from (Select rownum as ˳���, T.* from(' ||
              'Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,''/''||D.FtpĿ¼||''/'' As Root1,' ||
              'Decode(C.��������,Null,'''',to_Char(C.��������,''YYYYMMDD'')||''/'')||C.���UID||''/''||A.ͼ��UID As URL,d.�豸�� as �豸��1,' ||
              'E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,''/''||E.FtpĿ¼||''/'' As Root2,' ||
              'e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ���,A.����ͼ ' ||
              'From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E ' ||
              'Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) ' || Strfilter || ' ' ||
              'Order by ����UID, ͼ���) T ) ' || 'Where ˳���>=' || ��ʼλ��_In || ' and ˳���<=' || ����λ��_In || '';
  
    Strsql2 := 'Select * from (Select rownum as ˳���, T.* from(' ||
               'Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,''/''||D.FtpĿ¼||''/'' As Root1,' ||
               'Decode(C.��������,Null,'''',to_Char(C.��������,''YYYYMMDD'')||''/'')||C.���UID||''/''||A.ͼ��UID As URL,d.�豸�� as �豸��1,' ||
               'E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,''/''||E.FtpĿ¼||''/'' As Root2,' ||
               'e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� ' ||
               'From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E ' ||
               'Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) ' || Strfilter || ' ' ||
               'Order by ����UID, ͼ���) T ) ' || 'Where ˳���>=' || ��ʼλ��_In || ' and ˳���<=' || ����λ��_In || '';
  
    If �Ƿ���ʱ_In = 1 Then
      Strsql  := Replace(Strsql, 'Ӱ����', 'Ӱ����ʱ');
      Strsql2 := Replace(Strsql2, 'Ӱ����', 'Ӱ����ʱ');
    End If;
  
    Begin
      Open Val For Strsql;
    Exception
      When No_Column Then
        --���ݴ���10.36������ ����ͼ �ֶΣ������ 103996
        Open Val For Strsql2;
    End;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyimagedata;

  --8�����ܣ���ȡ��ʱͼ������
  Procedure p_Get_Tempimageseries
  (
    Val         Out t_Refcur,
    ʱ�䷶Χ_In In Number,
    ����_In     In Ӱ����ʱ��¼.����%Type := Null
  ) As
  Begin
    If ����_In Is Null Then
      Open Val For
        Select b.����uid, a.����, a.���� As ���, a.��������
        From Ӱ����ʱ��¼ a, Ӱ����ʱ���� b
        Where a.���uid = b.���uid
        And a.�������� Between Sysdate - ʱ�䷶Χ_In And Sysdate
        Order By ���;
    Else
      Open Val For
        Select b.����uid, a.����, a.���� As ���, a.��������
        From Ӱ����ʱ��¼ a, Ӱ����ʱ���� b
        Where a.���uid = b.���uid
        And a.�������� Between Sysdate - ʱ�䷶Χ_In And Sysdate
        And a.���� = ����_In
        Order By ���;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --9�����ܣ���ȡͼ��ע
  Procedure p_Get_Normalnote(Val Out t_Refcur) As
  Begin
    Open Val For
      Select b.��� As ���, b.���� As ����
      From Ӱ���ֵ��嵥 a, Ӱ���ֵ����� b
      Where a.Id = b.�ֵ�id
      And a.���� = 'Ӱ��ͼ��ע';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --10�����ܣ����볣��ͼ��ע
  Procedure p_Insert_Normalnote
  (
    Note_In In Ӱ���ֵ�����.����%Type,
    Code_In Ӱ���ֵ�����.����%Type
  ) As
    n_Num         Number;
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From Ӱ���ֵ��嵥 Where ˵�� = 'Ӱ��ͼ��ע';
    Select Decode(Max(To_Number(���)), Null, 0, Max(To_Number(���)))
    Into n_Num
    From Ӱ���ֵ�����
    Where �ֵ�id = Dictionary_Id;
    n_Num := n_Num + 1;
    Insert Into Ӱ���ֵ�����
      (�ֵ�id, ���, ����, ˵��)
    Values
      (Dictionary_Id, To_Char(n_Num), Note_In, 'Ӱ��ͼ��ע');
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Insert_Normalnote;

  --11�����ܣ��޸ĳ���ͼ��ע
  Procedure p_Edit_Normalnote
  (
    Note_In In Ӱ���ֵ�����.����%Type,
    Num_In  Ӱ���ֵ�����.���%Type
  ) As
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From Ӱ���ֵ��嵥 Where ˵�� = 'Ӱ��ͼ��ע';
    Update Ӱ���ֵ����� t
    Set t.���� = Note_In
    Where t.�ֵ�id = Dictionary_Id
    And t.��� = Num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Normalnote;

  --12�����ܣ�ɾ������ͼ��ע
  Procedure p_Del_Normalnote(Num_In Ӱ���ֵ�����.���%Type) As
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From Ӱ���ֵ��嵥 Where ˵�� = 'Ӱ��ͼ��ע';
    Delete Ӱ���ֵ����� t
    Where t.�ֵ�id = Dictionary_Id
    And t.��� = Num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Normalnote;

  --13�����ܣ���ȡ��ע����һ������
  Procedure p_Get_Normalnum(Val Out t_Refcur) As
    Dictionary_Id Varchar2(36);
  Begin
    Select Id Into Dictionary_Id From Ӱ���ֵ��嵥 Where ˵�� = 'Ӱ��ͼ��ע';
    Open Val For
      Select Decode(Max(To_Number(���)), Null, 1, Max(To_Number(���) + 1)) ���
      From Ӱ���ֵ����� t
      Where t.�ֵ�id = Dictionary_Id;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Normalnum;

  --14�����ܣ���ȡ���ID
  Procedure p_Get_Plugid
  (
    Val     Out t_Refcur,
    ����_In In Ӱ�񱨸���.����%Type
  ) Is
  Begin
    Open Val For
      Select Rawtohex(Id) Id From Ӱ�񱨸��� Where ���� = ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Plugid;

  --15�����ܣ�����༭���������
  Procedure p_Setfontparam
  (
    Font_In Nvarchar2,
    User_In Nvarchar2
  ) As
    m_Id     Nvarchar2(36);
    Numcount Int;
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From Ӱ�����˵��
    Where ģ�� = 'ImageEditor'
    And ������ = '��������';
    Select Count(*)
    Into Numcount
    From Ӱ�����ȡֵ t
    Where t.����id = m_Id
    And t.������ʶ = User_In;
    If Numcount > 0 Then
      Update Ӱ�����ȡֵ a
      Set a.����ֵ = Font_In
      Where a.������ʶ = User_In
      And a.����id = m_Id;
    Else
      Insert Into Ӱ�����ȡֵ a (Id, ����id, ������ʶ, ����ֵ) Values (Sys_Guid(), m_Id, User_In, Font_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Setfontparam;

  --16�����ܣ���ȡ�༭���������
  Procedure p_Getfontparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  ) As
    m_Id Nvarchar2(36);
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From Ӱ�����˵��
    Where ģ�� = 'ImageEditor'
    And ������ = '��������';
    Open Val For
      Select a.����ֵ
      From Ӱ�����ȡֵ a
      Where a.����id = m_Id
      And a.������ʶ = User_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getfontparam;

  --17�����ܣ�����༭���������
  Procedure p_Setformparam
  (
    Form_In Nvarchar2,
    User_In Nvarchar2
  ) As
    m_Id     Nvarchar2(36);
    Numcount Int;
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From Ӱ�����˵��
    Where ģ�� = 'ImageEditor'
    And ������ = '��������';
    Select Count(*)
    Into Numcount
    From Ӱ�����ȡֵ t
    Where t.����id = m_Id
    And t.������ʶ = User_In;
    If Numcount > 0 Then
      Update Ӱ�����ȡֵ a
      Set a.����ֵ = Form_In
      Where a.������ʶ = User_In
      And a.����id = m_Id;
    Else
      Insert Into Ӱ�����ȡֵ a (Id, ����id, ������ʶ, ����ֵ) Values (Sys_Guid(), m_Id, User_In, Form_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Setformparam;

  --18�����ܣ���ȡ�༭���������
  Procedure p_Getformparam
  (
    Val     Out t_Refcur,
    User_In Nvarchar2
  ) As
    m_Id Nvarchar2(36);
  Begin
    Select Rawtohex(Id)
    Into m_Id
    From Ӱ�����˵��
    Where ģ�� = 'ImageEditor'
    And ������ = '��������';
    Open Val For
      Select a.����ֵ
      From Ӱ�����ȡֵ a
      Where a.����id = m_Id
      And a.������ʶ = User_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getformparam;

  --19�����ܣ�����ͼ��UID��ȡ�����Ϣ
  Procedure p_Getstudyinfobyimageuid
  (
    Val        Out t_Refcur,
    ҽ��id_In  In Ӱ�����¼.ҽ��id%Type,
    ͼ��uid_In In Ӱ����ͼ��.ͼ��uid%Type
  ) As
  Begin
    Open Val For
      Select d.���uid
      From Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c, Ӱ����ʱ���� d
      Where c.ҽ��id = ҽ��id_In
      And a.ͼ��uid = ͼ��uid_In
      And a.����uid = b.����uid
      And b.���uid = c.���uid
      And a.����uid = d.����uid;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyinfobyimageuid;

  --20�����ܣ����ݼ��UID��ȡFTP��Ϣ
  Procedure p_Getftpinfobystudyuid
  (
    Val        Out t_Refcur,
    ���uid_In In Ӱ�����¼.���uid%Type
  ) As
  Begin
    Open Val For
      Select d.Ftp�û��� As Ftpuser, d.Ftp���� As Ftppwd, c.λ��һ, c.λ�ö�, c.λ����, c.��������, d.Ip��ַ As Host,
             '/' || d.FtpĿ¼ || '/' As Root,
             Decode(c.��������, Null, '', To_Char(c.��������, 'YYYYMMDD') || '/') || c.���uid As Url
      From Ӱ�����¼ c, Ӱ���豸Ŀ¼ d
      Where Decode(c.λ��һ, Null, c.λ�ö�, c.λ��һ) = d.�豸��(+)
      And c.���uid = ���uid_In
      Union All
      Select d.Ftp�û��� As Ftpuser, d.Ftp���� As Ftppwd, c.λ��һ, c.λ�ö�, c.λ����, c.��������, d.Ip��ַ As Host,
             '/' || d.FtpĿ¼ || '/' As Root,
             Decode(c.��������, Null, '', To_Char(c.��������, 'YYYYMMDD') || '/') || c.���uid As Url
      From Ӱ����ʱ��¼ c, Ӱ���豸Ŀ¼ d
      Where Decode(c.λ��һ, Null, c.λ�ö�, c.λ��һ) = d.�豸��(+)
      And c.���uid = ���uid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getftpinfobystudyuid;

  --21�����ܣ����ݿ���ID��ȡFTP��Ϣ
  Procedure p_Getftpinfobydeptid
  (
    Val       Out t_Refcur,
    ����id_In In Ӱ�����̲���.����id%Type
  ) As
  Begin
    Open Val For
      Select a.�豸��, a.Ip��ַ, a.Ftp�û���, a.Ftp����
      From Ӱ���豸Ŀ¼ a, Ӱ�����̲��� b
      Where a.�豸�� = b.����ֵ
      And b.������ = '�洢�豸��'
      And b.����id = ����id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getftpinfobydeptid;

  --22�����ܣ�����ҽ��ID��ȡFTP��Ϣ
  Procedure p_Getftpinfobyadvicetid
  (
    Val       Out t_Refcur,
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type
  ) As
  Begin
    Open Val For
      Select a.�豸��, a.Ip��ַ, a.Ftp�û���, a.Ftp����
      From Ӱ���豸Ŀ¼ a, Ӱ�����¼ b
      Where b.λ��һ = a.�豸��(+)
      And b.ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getftpinfobyadvicetid;

  --23�����ܣ���ȡ���UID
  Procedure p_Getstudyuid
  (
    Val        Out t_Refcur,
    ���uid_In In Ӱ�����¼.���uid%Type
  ) As
  Begin
    Open Val For
      Select ���uid
      From Ӱ�����¼
      Where ���uid = ���uid_In
      Union All
      Select ���uid
      From Ӱ����ʱ��¼
      Where ���uid = ���uid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getstudyuid;

  --24�����ܣ���ȡ����UID
  Procedure p_Getseriesuid
  (
    Val        Out t_Refcur,
    ����uid_In In Ӱ��������.����uid%Type
  ) As
  Begin
    Open Val For
      Select ����uid
      From Ӱ��������
      Where ����uid = ����uid_In
      Union All
      Select ����uid
      From Ӱ����ʱ����
      Where ����uid = ����uid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getseriesuid;

  --25�����ܣ������豸�Ż�ȡ�豸��Ϣ
  Procedure p_Getdeviceinfo
  (
    Val       Out t_Refcur,
    �豸��_In In Ӱ���豸Ŀ¼.�豸��%Type
  ) As
  Begin
    Open Val For
      Select �豸��, �豸��, '/' || Decode(FtpĿ¼, Null, '', FtpĿ¼ || '/') As Url, Ftp�û���, Ftp����, Ip��ַ
      From Ӱ���豸Ŀ¼
      Where ���� = 1
      And �豸�� = �豸��_In
      And Nvl(״̬, 0) = 1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getdeviceinfo;

  --26����ȡҽ��վ�洢�豸��
  Procedure p_Getdeviceidbyadviceid
  (
    Val       Out t_Refcur,
    ҽ��id_In In ����ҽ������.ҽ��id%Type
  ) As
  Begin
    Open Val For
      Select d.����ֵ
      From ҽ��ִ�з��� a, ����ҽ������ b, Ӱ��dicom����� c, Ӱ��dicom������� d
      Where a.����id = b.ִ�в���id
      And a.ִ�м� = b.ִ�м�
      And a.����豸 = c.�豸��
      And c.������ = 'ͼ�����'
      And c.����id = d.����id
      And d.�������� = '�洢�豸'
      And b.ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Getdeviceidbyadviceid;
End b_Pacs_Rptpluginoriginal;

/



--Ӱ�񱨸淶�Ĺ���(---���岿��---)***************************************************
Create Or Replace Package b_PACS_RptPluginCustom Is
  Type t_Refcur Is Ref Cursor;
-- ��    �ܣ��÷���ֻ������ʾ...
  Procedure Demo1;

end b_PACS_RptPluginCustom ;
/

--Ӱ�񱨸淶�Ĺ���(---ʵ�ֲ���---)***************************************************
Create Or Replace Package Body b_PACS_RptPluginCustom  Is
-- ��    �ܣ��÷���ֻ������ʾ...
  Procedure Demo1
  Is
  Begin
      --TODO:
      Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Demo1;

End b_PACS_RptPluginCustom;
/


--XWPACS�ӿڰ�
Create Or Replace Package b_XINWANGInterface Is
  Type t_Refcur Is Ref Cursor;
  -- 1 PACS״̬�ı���Ϣ
  Procedure PacsStatusChange
  (
    ״̬ID_In   In Number,
    ҽ��ID_In   In Ӱ�����¼.ҽ��ID%Type,
    Ӱ�����_In In Ӱ�����¼.Ӱ�����%Type,
    ����_In   In Ӱ�����¼.����%Type,
    ����ʱ��    In Date,
    ִ����      In Varchar2,
    ��Ƭ��С    In Varchar2,
    ���UID_In  In Ӱ�����¼.���UID%Type := Null
  );
  -- 2 ȡ��ͼ�����
  Procedure PacsUnmatchImage(ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type);
  -- 3 ��д����ͼ�Ĵ洢�豸
  Procedure PacsSetFTPDeviceNo
  (
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type,
    �豸��_In In Ӱ�����¼.λ��һ%Type
  );
  -- 4 ����ͼ����
  Procedure UpdateImgCount
  (
    ҽ��ID_In Ӱ�����¼.ҽ��ID%Type,
    ͼ����_In Number
  );

End b_XINWANGInterface;

/

Create Or Replace Package Body b_XINWANGInterface Is

  -- 1 PACS״̬�ı���Ϣ
  Procedure PacsStatusChange
  (
    ״̬ID_In   In Number,
    ҽ��ID_In   In Ӱ�����¼.ҽ��ID%Type,
    Ӱ�����_In In Ӱ�����¼.Ӱ�����%Type,
    ����_In   In Ӱ�����¼.����%Type,
    ����ʱ��    In Date,
    ִ����      In Varchar2,
    ��Ƭ��С    In Varchar2,
    ���UID_In  In Ӱ�����¼.���UID%Type := Null
  ) Is
    Strsql     Varchar2(2000);
    v_StudyUID Ӱ�����¼.���UID%Type;
  
    Cursor c_Advice Is
      Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or (���id = ҽ��id_In And ������� In ('F', 'G', 'D'));
  
  Begin
    --״̬ID_In:1-ƥ��ɹ�;2-ƥ��ʧ��;3-�¼�飨�յ���һ��ͼ��;4-�յ�ÿһ��ͼ��;
    --     5-ɾ�����;6-��Ƭ��ӡ�ɹ���7-���µ��ӽ�Ƭ״̬��8-ͼ��ת�Ƶ���ƽ̨��9-ͼ�����ƽ̨����
  
    If ���UID_In Is Null Then
      v_StudyUID := ҽ��ID_In;
    Else
      v_StudyUID := ���UID_In;
    End If;
  
    If ״̬ID_In = 1 Then
      --ͼ��ƥ��ɹ�
    
      --��дӰ�����¼��� ���UID���������ڵȣ����ǲ���д���м���ı�,���UID��д���UID_In��StudyUID��
      Update Ӱ�����¼
      Set ���UID = v_StudyUID, �������� = Decode(����ʱ��, Null, Sysdate, ����ʱ��), ͼ��λ�� = 1
      Where ҽ��ID = ҽ��ID_In;
    
      --����ҽ��ִ��״̬
      For r_Advice In c_Advice Loop
        Update ����ҽ������
        Set ִ��״̬ = 3, ִ�й��� = Decode(Sign(ִ�й��� - 2), 1, ִ�й���, 3)
        Where ҽ��id = r_Advice.id;
      End Loop;
    Elsif ״̬ID_In = 2 Then
      Strsql := 'dd';
    Elsif ״̬ID_In = 3 Then
      -- 3-�¼�飨�յ���һ��ͼ�񣩣���ʱ������
      Strsql := 'dd';
    Elsif ״̬ID_In = 4 Then
      --  4-�յ�ÿһ��ͼ�� ����ʱ������
      Strsql := 'dd';
    Elsif ״̬ID_In = 5 Then
      -- 5-ɾ�����
      -- ɾ��Ӱ�����¼���ж�Ӧ�ļ��UID���������ڵ�
      Update Ӱ�����¼
      Set ���UID = Null, λ��һ = Null, λ�ö� = Null, λ���� = Null, ����ͼ�� = Null, �������� = Null
      Where ҽ��ID = ҽ��ID_IN;
    Elsif ״̬ID_In = 6 Then
      -- 6-��Ƭ��ӡ�ɹ�
      --��¼��Ƭ��С����ӡ�˵�
    
      --һ��ҽ����ӡһ�Ż��߶��Ž�Ƭ�������ÿ�Ž�Ƭ����һ���̣����IDΪ��
      Insert Into ��Ƭ��ӡ��¼
        (ID, ���id, ҽ��id, ��Ƭ��С, ��ӡ��, ��ӡʱ��)
      Values
        (��Ƭ��ӡ��¼_Id.Nextval, Null, ҽ��ID_In, ��Ƭ��С, ִ����, Decode(����ʱ��, Null, Sysdate, ����ʱ��));
      Update Ӱ�����¼ Set �Ƿ��ӡ = 1 Where ҽ��ID = ҽ��ID_In;
    Elsif ״̬ID_In = 7 Then
      --���µ��ӽ�Ƭ״̬
      Update Ӱ�����¼ Set �Ƿ���ӽ�Ƭ = 1 Where ҽ��ID = ҽ��ID_In;
    Elsif ״̬ID_In = 8 Then
      -- ͼ��ת�Ƶ���ƽ̨
      Update Ӱ�����¼ Set ���UID = ���UID_In, ͼ��λ�� = 2 Where ҽ��ID = ҽ��ID_In;
    Elsif ״̬ID_In = 9 Then
      -- ͼ�����ƽ̨����
      Update Ӱ�����¼ Set ͼ��λ�� = 1 Where ҽ��ID = ҽ��ID_In;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End PacsStatusChange;

  -- 2 PACSͼ��ȡ������
  Procedure PacsUnmatchImage(ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type) Is
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    v_���ͺ�   ����ҽ������.���ͺ�%Type;
  Begin
    --����Ӱ�����¼���״̬
    Update Ӱ�����¼ Set ���UID = Null, �������� = Null, ͼ��λ�� = Null, λ��һ = Null Where ҽ��ID = ҽ��ID_In;
  
    --���� Zl_Ӱ����_State �ı�����̵�״̬
    Select ִ�й���, ���ͺ� Into v_ִ�й���, v_���ͺ� From ����ҽ������ Where ҽ��ID = ҽ��ID_In;
  
    --���ִ�й�����3���򽫹����޸ĳ�2
    If v_ִ�й��� = 3 Then
      Zl_Ӱ����_State(ҽ��ID_In, v_���ͺ�, 2);
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End PacsUnmatchImage;

  -- 3 ��д����ͼ�Ĵ洢�豸
  Procedure PacsSetFTPDeviceNo
  (
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type,
    �豸��_In In Ӱ�����¼.λ��һ%Type
  ) Is
  Begin
    --����Ӱ�����¼���״̬
    Update Ӱ�����¼ Set λ��һ = �豸��_In Where ҽ��ID = ҽ��ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End PacsSetFTPDeviceNo;

  -- 4 ����ͼ������
  Procedure UpdateImgCount
  (
    ҽ��ID_IN Ӱ�����¼.ҽ��ID%Type,
    ͼ����_In Number
  ) Is
  Begin
    Update Ӱ�����¼ Set ͼ������ = ͼ����_In Where ҽ��ID = ҽ��ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End UpdateImgCount;

End b_XINWANGInterface;
/

Create Or Replace Package b_Emergency_Rating Is
  --��ʹ�ּ�����
  --�������1�������ı�
  --������ʽ��b_emergency_rating.Is_Pain_num_rating(4,5)
  --����ֵ���ͣ�varchar2   ���ؽ����ʽ ���˵ȼ�:����������  ���磺2:9:�ض���ʹ
  Function Is_Pain_Num_Rating(Describe Varchar2) Return Varchar2;
  --�������ַּ�����
  --�������3�����۷�Ӧָ��id��ָ�������� �� ���Է�Ӧָ��id��ָ�������� �ܻ��Ӧָ��id��ָ��������
  --������ʽ�� b_emergency_rating.Is_coma_rating('1:�����̼�','2:��������','3:ʹ�̼�����')
  --����ֵ���ͣ�varchar2   ���ؽ����ʽ ���˵ȼ�:�ܷ���������  ���磺3:11:�ж���ʶ�ϰ�
  Function Is_Coma_Rating
  (
    Open_Reaction     Varchar2,
    Language_Reaction Varchar2,
    Activity_Reaction Varchar2
  ) Return Varchar2;
  --�жϿ͹�����Ϊ��ͯ���ǳ��˺���
  Function Is_Judgement_Function
  (
    Agenum  Number,
    Ageunit Varchar2
  ) Return Varchar2;
  --�͹����۷ּ��������˺Ͷ�ͯ����
  --�������3������ �����䵥λ �� ָ��id��ָ�����������ɶ����
  --������ʽ�� b_emergency_rating.Is_objective_rating(5,'��','11:9,6:100,4:20')
  --����ֵ���ͣ�varchar2   ���ؽ����ʽ ���˵ȼ� 1
  Function Is_Objective_Rating
  (
    Agenum           Number,
    Ageunit          Varchar2,
    Indexid_Describe Varchar2
  ) Return Varchar2;
End b_Emergency_Rating;
/
Create Or Replace Package Body b_Emergency_Rating Is

  Function Is_Pain_Num_Rating(Describe Varchar2) --��ʹ�ȼ�����
   Return Varchar2 As
    State_Level  Varchar2(10); --���˼���
    Score_Result Varchar2(100); --���ֽ������
    Score        Number; --����
  Begin
    Select ָ������ֵ Into Score From �������ַ������� Where ָ�������� = Describe;
  
    Select Min(���鼶��), Min(���ֽ������)
    Into State_Level, Score_Result
    From �������ַ����ּ�
    Where ����� = 2 And Score > ��ֵ���� And ����id = 4 Or ����� = 3 And ��ֵ���� < Score And ����id = 4 Or
          ����� = 6 And Score Between ��ֵ���� And ��ֵ���� And ����id = 4 Or ����� = 1 And ��ֵ���� = Score And ����id = 4 Or
          ����� = 4 And ��ֵ���� >= Score And ����id = 4 Or ����� = 5 And Score <= ��ֵ���� And ����id = 4;
    Return State_Level || ':' || Score || ':' || Score_Result;
  Exception
    --�쳣��������
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Pain_Num_Rating;

  Function Is_Coma_Rating( --���Եȼ�����
                          Open_Reaction     Varchar2,
                          Language_Reaction Varchar2,
                          Activity_Reaction Varchar2) Return Varchar2 As
  
    Coma_Score_All Number; --�����ܷ���
    Coma_Level     Varchar2(10); --���Եȼ�
    Score_Result   Varchar2(100); --���ֽ������
  
    Coma_Id1    Varchar2(10); --����-����ָ��ID
    Coma_Text1  Varchar2(100); --����-��������
    Coma_Score1 Number; --����-���۷���
  
    Coma_Id2    Varchar2(10); --����-����ָ��ID
    Coma_Text2  Varchar2(100); --����-��������
    Coma_Score2 Number; --����-���Է���
  
    Coma_Id3    Varchar2(10); --����-�ָ��ID
    Coma_Text3  Varchar2(100); --����-�����
    Coma_Score3 Number; --����-�����
  Begin
    Select C1, C2 Into Coma_Id1, Coma_Text1 From Table(f_Str2list2(Open_Reaction));
    Select C1, C2 Into Coma_Id2, Coma_Text2 From Table(f_Str2list2(Language_Reaction));
    Select C1, C2 Into Coma_Id3, Coma_Text3 From Table(f_Str2list2(Activity_Reaction));
  
    Select ָ������ֵ
    Into Coma_Score1
    From �������ַ�������
    Where ����id = 3 And ָ�������� = Coma_Text1 And ָ��id = Coma_Id1;
  
    Select ָ������ֵ
    Into Coma_Score2
    From �������ַ�������
    Where ����id = 3 And ָ�������� = Coma_Text2 And ָ��id = Coma_Id2;
  
    Select ָ������ֵ
    Into Coma_Score3
    From �������ַ�������
    Where ����id = 3 And ָ�������� = Coma_Text3 And ָ��id = Coma_Id3;
    Coma_Score_All := Coma_Score1 + Coma_Score2 + Coma_Score3;
  
    Select Min(���鼶��), Min(���ֽ������)
    Into Coma_Level, Score_Result
    From �������ַ����ּ�
    Where ����� = 2 And Coma_Score_All > ��ֵ���� And ����id = 3 Or ����� = 3 And ��ֵ���� < Coma_Score_All And ����id = 3 Or
          ����� = 6 And Coma_Score_All Between ��ֵ���� And ��ֵ���� And ����id = 3 Or
          ����� = 1 And ��ֵ���� = Coma_Score_All And ����id = 3 Or ����� = 4 And ��ֵ���� >= Coma_Score_All And ����id = 3 Or
          ����� = 5 And Coma_Score_All <= ��ֵ���� And ����id = 3;
    Return Coma_Level || ':' || Coma_Score_All || ':' || Score_Result;
  
  Exception
    --�쳣��������
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Coma_Rating;

  Function Is_Judgement_Function
  (
    Agenum  Number,
    Ageunit Varchar2
  ) Return Varchar2 As
    --�жϳ��˻��ͯ����
    Children_Age Varchar2(100);
  Begin
    If Ageunit Is Null Then
      Return '1'; --����
    End If;
  
    If Ageunit = '��' Then
      Select ����ֵ Into Children_Age From zlParameters Where ������ = '��ͯ����綨����';
    
      If Agenum <= To_Number(Children_Age) Then
        Return '2'; --��ͯ
      Else
        Return '1'; --����
      End If;
    End If;
    Return '2'; --���䵥λ��Ϊ�귵�ض�ͯ0-1��
  Exception
    --�쳣��������
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Judgement_Function;

  Function Is_Objective_Rating --�͹��������ֹ���
  (
    Agenum           Number,
    Ageunit          Varchar2,
    Indexid_Describe Varchar2
  ) Return Varchar2 As
    Person        Varchar2(2); --��ͯ���߳���
    o_Indexid     t_Numlist; --ָ��ID
    o_Describe    t_Numlist; --����ָ�����
    Level_Max     Number; --�������ֵ
    Illness_Level Number; --���鼶��
    Age_Id        Number; --��ͯ����id
  Begin
    Select b_Emergency_Rating.Is_Judgement_Function(Agenum, Ageunit) Into Person From Dual;
    If Person = '1' Then
      --���˵Ĺ���
      Select Max(���鼶��) Into Level_Max From �������ַ�������;
      Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
      For I In 1 .. o_Indexid.Count Loop
        Select Min(���鼶��)
        Into Illness_Level
        From �������ַ�������
        Where ����� = 2 And o_Describe(I) > ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 1 Or
              ����� = 3 And o_Describe(I) < ָ��ֵ���� And ����id = 1 And ָ��id = o_Indexid(I) Or
              ����� = 6 And o_Describe(I) >= ָ��ֵ���� And o_Describe(I) <= ָ��ֵ���� And ����id = 1 And ָ��id = o_Indexid(I) Or
              ����� = 1 And ָ��ֵ���� = o_Describe(I) And ����id = 1 And ָ��id = o_Indexid(I) Or
              ����� = 4 And o_Describe(I) >= ָ��ֵ���� And ����id = 1 And ָ��id = o_Indexid(I) Or
              ����� = 5 And o_Describe(I) <= ָ��ֵ���� And ����id = 1 And ָ��id = o_Indexid(I);
        If Illness_Level < Level_Max Then
          Level_Max := Illness_Level;
        End If;
      End Loop;
      Return Level_Max;
    End If;
  
    If Person = '2' Then
      --��ͯ����
      Select Max(���鼶��) Into Level_Max From �������ַ�������;
      --�����߼����ݴ�������͵�λ��ָ��id�ҵ���Ӧ������id����������id��ָ��id��ָ��ֵ�ҵ�����
      --���û���ҵ���������������Ѱ��û������ֵ�ļ����������Ϊ�ս��ҵ�����С���𸳸���
      If Ageunit = '��' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From ��������ָ������
          Where ����� = 2 And Agenum > �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 3 And Agenum < �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' And ���䵥λ = '��' Or
                ����� = 6 And Agenum Between �������� And �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 1 And �������� = Agenum And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 4 And Agenum >= �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 5 And Agenum <= �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��';
        
          Select Min(���鼶��)
          Into Illness_Level
          From �������ַ�������
          Where ����� = 2 And o_Describe(I) > ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 3 And o_Describe(I) < ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 6 And o_Describe(I) Between ָ��ֵ���� And ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And
                ����id = 2 Or ����� = 1 And ָ��ֵ���� = o_Describe(I) And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 4 And o_Describe(I) >= ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 5 And o_Describe(I) <= ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2;
          If Illness_Level Is Null Then
            Select Nvl(Min(���鼶��), Level_Max)
            Into Illness_Level
            From �������ַ�������
            Where ����� = 2 And o_Describe(I) > ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 3 And o_Describe(I) < ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 6 And o_Describe(I) Between ָ��ֵ���� And ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And
                  ָ������id Is Null Or
                  ����� = 1 And ָ��ֵ���� = o_Describe(I) And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 4 And o_Describe(I) >= ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 5 And o_Describe(I) <= ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null;
          
          End If;
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
      If Ageunit = '��' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From ��������ָ������
          Where ����� = 2 And Agenum > �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 3 And Agenum < �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' And ���䵥λ = '��' Or
                ����� = 6 And Agenum Between �������� And �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 1 And �������� = Agenum And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 4 And Agenum >= �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 5 And Agenum <= �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��';
          Select Min(���鼶��)
          Into Illness_Level
          From �������ַ�������
          Where ����� = 2 And o_Describe(I) > ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 3 And o_Describe(I) < ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 6 And o_Describe(I) Between ָ��ֵ���� And ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And
                ����id = 2 Or ����� = 1 And ָ��ֵ���� = o_Describe(I) And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 4 And o_Describe(I) >= ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 5 And o_Describe(I) <= ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2;
        
          If Illness_Level Is Null Then
            Select Nvl(Min(���鼶��), Level_Max)
            Into Illness_Level
            From �������ַ�������
            Where ����� = 2 And o_Describe(I) > ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 3 And o_Describe(I) < ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 6 And o_Describe(I) Between ָ��ֵ���� And ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And
                  ָ������id Is Null Or
                  ����� = 1 And ָ��ֵ���� = o_Describe(I) And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 4 And o_Describe(I) >= ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 5 And o_Describe(I) <= ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null;
          
          End If;
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
      If Ageunit = '��' Then
        Select C1, C2 Bulk Collect Into o_Indexid, o_Describe From Table(f_Str2list2(Indexid_Describe));
        For I In 1 .. o_Indexid.Count Loop
          Select Max(ID)
          Into Age_Id
          From ��������ָ������
          Where ����� = 2 And Agenum > �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 3 And Agenum < �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' And ���䵥λ = '��' Or
                ����� = 6 And Agenum Between �������� And �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 1 And �������� = Agenum And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 4 And Agenum >= �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��' Or
                ����� = 5 And Agenum <= �������� And ָ��id = o_Indexid(I) And ���䵥λ = '��';
          Select Min(���鼶��)
          Into Illness_Level
          From �������ַ�������
          Where ����� = 2 And o_Describe(I) > ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 3 And o_Describe(I) < ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 6 And o_Describe(I) Between ָ��ֵ���� And ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And
                ����id = 2 Or ����� = 1 And ָ��ֵ���� = o_Describe(I) And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 4 And o_Describe(I) >= ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2 Or
                ����� = 5 And o_Describe(I) <= ָ��ֵ���� And ָ��id = o_Indexid(I) And ָ������id = Age_Id And ����id = 2;
          If Illness_Level Is Null Then
            Select Nvl(Min(���鼶��), Level_Max)
            Into Illness_Level
            From �������ַ�������
            Where ����� = 2 And o_Describe(I) > ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 3 And o_Describe(I) < ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 6 And o_Describe(I) Between ָ��ֵ���� And ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And
                  ָ������id Is Null Or
                  ����� = 1 And ָ��ֵ���� = o_Describe(I) And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 4 And o_Describe(I) >= ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null Or
                  ����� = 5 And o_Describe(I) <= ָ��ֵ���� And ָ��id = o_Indexid(I) And ����id = 2 And ָ������id Is Null;
          End If;
        
          If Illness_Level < Level_Max Then
            Level_Max := Illness_Level;
          End If;
        
        End Loop;
        Return Level_Max;
      End If;
    
    End If;
    Return Level_Max;
  Exception
    --�쳣��������
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Is_Objective_Rating;

End b_Emergency_Rating;
/

Create Or Replace Package Pkg_Pretriage_Dql As
  -----------------------------------------------------
  --��ȡ������������
  -----------------------------------------------------
  Procedure Get_Nationality
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ�����б�
  -----------------------------------------------------
  Procedure Get_Reportlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --������֤¼���Ƿ���ȷ
  -----------------------------------------------------
  Procedure Checkidcard
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --�������¼���Ƿ���ȷ
  -----------------------------------------------------
  Procedure Checkage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --ͨ��ҽ���Ŷ�ȡ������Ϣ�б��嵥
  -----------------------------------------------------
  Procedure Get_Patlistbymedical
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --ͨ�����֤�Ŷ�ȡ������Ϣ�б��嵥
  -----------------------------------------------------
  Procedure Get_Patlistbyidcard
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --ͨ����������ƥ�䲡����Ϣ�б��嵥
  -----------------------------------------------------
  Procedure Get_Patlistbyname
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���µľ���״̬
  -----------------------------------------------------
  Procedure Getvisitstate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���˷���������Ϣ
  -----------------------------------------------------
  Procedure Load_Levelinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���˷���ָ����Ϣ
  -----------------------------------------------------
  Procedure Load_Rulesinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���˷����¼����
  -----------------------------------------------------
  Procedure Load_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ�������˷�����Ϣ
  -----------------------------------------------------
  Procedure Get_Patidetail
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --��ȡ�����б��嵥
  -----------------------------------------------------
  Procedure Get_Patlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ʹ�ּ�����
  -----------------------------------------------------
  Procedure Get_Pain_Num_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�������ַּ�����
  -----------------------------------------------------
  Procedure Get_Coma_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --�͹����۷ּ��������˺Ͷ�ͯ����
  -----------------------------------------------------
  Procedure Get_Objective_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��ͯ��������
  -----------------------------------------------------
  Procedure Get_Childmaxage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --��ȡ����ȼ�
  -----------------------------------------------------
  Procedure Get_Level
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ�������
  -----------------------------------------------------
  Procedure Get_Dept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ�˹���������
  -----------------------------------------------------
  Procedure Get_Rules
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --���ݳ������ڷ�������
  -----------------------------------------------------
  Procedure Get_Datetoage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --��ȡ�Ա��������
  -----------------------------------------------------
  Procedure Get_Sexbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ�����������
  -----------------------------------------------------
  Procedure Get_Nationbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ��������ָ��
  -----------------------------------------------------
  Procedure Get_Scorebase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
  -----------------------------------------------------
  --��ȡ��������
  -----------------------------------------------------
  Procedure Get_Paticc
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ������Դ
  -----------------------------------------------------
  Procedure Get_Patifrom
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ������ʶ״̬
  -----------------------------------------------------
  Procedure Get_Patistate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ������ͬ��Ա
  -----------------------------------------------------
  Procedure Get_Entourage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���ﳣ������ʷ
  -----------------------------------------------------
  Procedure Get_Dishistory
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );

  -----------------------------------------------------
  --��ȡ���ݿ�ϵͳʱ��
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

End Pkg_Pretriage_Dql;
/
Create Or Replace Package Body Pkg_Pretriage_Dql As
  -----------------------------------------------------
  --��ȡ������������
  -----------------------------------------------------
  Procedure Get_Nationality
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����, ����, ����, Nvl(ȱʡ��־, 0) As ȱʡ From ���� Order By Nvl(ȱʡ��־, 0) Desc, ����;
  End Get_Nationality;

  -----------------------------------------------------
  --��ȡ�����б�
  -----------------------------------------------------
  Procedure Get_Reportlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ��־, ϵͳ, ���, ����, Nvl(�Ƿ�ͣ��, 0) �Ƿ�ͣ��
      From (Select 1 As ��־, a.ϵͳ, a.���, a.����, a.�Ƿ�ͣ��
             From zlReports A, zlPrograms B
             Where a.ϵͳ = b.ϵͳ And a.����id = b.��� And Not Upper(a.���) Like '%BILL%' And Upper(b.����) <> Upper('zl9Report') And
                   b.ϵͳ = 100 And b.��� = 1244
             Union All
             Select Decode(a.ϵͳ, Null, 2, 1) As ��־, a.ϵͳ, a.���, a.����, a.�Ƿ�ͣ��
             From zlReports A, zlRPTPuts B, zlPrograms C
             Where a.Id = b.����id And b.ϵͳ = c.ϵͳ And b.����id = c.��� And (Not Upper(a.���) Like '%BILL%' Or a.ϵͳ Is Null) And
                   c.ϵͳ = 100 And c.��� = 1244)
      Where Instr(',ZL1_REPORT_1244_1,ZL1_REPORT_1244_2,', ',' || ��� || ',') = 0 And Nvl(�Ƿ�ͣ��, 0) = 0
      Order By ��־, ���;
  End Get_Reportlist;

  -----------------------------------------------------
  --������֤¼���Ƿ���ȷ
  -----------------------------------------------------
  Procedure Checkidcard
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --��Σ�Json_In:��ʽ
    --  input
    --    idcard        C 1 ¼������֤��
    -- ����ֵ���̶���ʽXML��
    --<OUTPUT>
    --       <BIRTHDAY></BIRTHDAY>                //��������
    --       <SEX></SEX>                  //�Ա�
    --       <AGE></AGE>                //����
    --     <MSG></MSG>         //�մ�-���֤����Ч(�ɴ����֤���л�ȡ�������ں��Ա�)���ǿմ�-���ش�����Ϣ
    --</OUTPUT>
  
    Jsonobj  Pljson;
    j_In     Pljson;
    v_¼���� Varchar2(50);
  Begin
    j_In       := Pljson(Input_In);
    Jsonobj    := j_In.Get_Pljson('input');
    v_¼����   := Jsonobj.Get_String('idcard');
    Output_Out := Zl_Fun_Checkidcard(v_¼����);
  End Checkidcard;

  -----------------------------------------------------
  --�������¼���Ƿ���ȷ
  -----------------------------------------------------
  Procedure Checkage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --��Σ�Json_In:��ʽ
    --  input
    --    age        C 1 ����
    Jsonobj Pljson;
    j_In    Pljson;
    v_����  Varchar2(50);
  Begin
    j_In       := Pljson(Input_In);
    Jsonobj    := j_In.Get_Pljson('input');
    v_����     := Jsonobj.Get_String('age');
    Output_Out := Zl_Age_Check(v_����);
  End Checkage;

  -----------------------------------------------------
  --ͨ��ҽ���Ŷ�ȡ������Ϣ�б��嵥
  -----------------------------------------------------
  Procedure Get_Patlistbymedical
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    Jsonobj    Pljson;
    j_In       Pljson;
    v_ҽ����   Varchar2(200);
    v_ҽ������ Varchar2(200);
  Begin
    j_In       := Pljson(Input_In);
    Jsonobj    := j_In.Get_Pljson('input');
    v_ҽ����   := Jsonobj.Get_String('ҽ����');
    v_ҽ������ := Jsonobj.Get_String('ҽ������');
    Open Output_Out For
      Select /*+Rule */
      Distinct a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��, a.�ֻ���, a.ҽ����,
               b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
      From ������Ϣ A, ������� B
      Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.ҽ���� = v_ҽ���� And b.���� = v_ҽ������
      Order By ����id Desc;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbymedical;

  -----------------------------------------------------
  --ͨ�����֤�Ŷ�ȡ������Ϣ�б��嵥
  -----------------------------------------------------
  Procedure Get_Patlistbyidcard
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    Jsonobj    Pljson;
    j_In    Pljson;
    v_���֤�� Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_���֤�� := Jsonobj.Get_String('���֤��');
    Open Output_Out For
      Select /*+Rule */
      Distinct a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��, a.�ֻ���, a.ҽ����,
               b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
      From ������Ϣ A, ������� B
      Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.���֤�� = v_���֤��
      Order By ����id Desc;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbyidcard;

  -----------------------------------------------------
  --ͨ����������ƥ�䲡����Ϣ�б��嵥
  -----------------------------------------------------
  Procedure Get_Patlistbyname
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    Jsonobj    Pljson;
    j_In    Pljson;
    v_����     Varchar2(200);
    n_¼����   Varchar2(200);
    n_��ȷ���� Number;
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_����  := Jsonobj.Get_String('��������');
  
    If v_���� Is Not Null Then
      If Substr(v_����, 1, 1) = '-' Then
        n_��ȷ���� := 1;
        v_����     := Substr(v_����, 2);
        n_¼����   := Zl_To_Number(v_����);
      End If;
    End If;
    If n_��ȷ���� = 1 Then
      Open Output_Out For
        Select /*+Rule */
         1 As ����id, a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��, a.�ֻ���,
         a.ҽ����, b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
        From ������Ϣ A, ������� B
        Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.���֤�� = v_����
        Union All
        Select 1 /*+Rule */ As ����id, a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����,
               a.���֤��, a.�ֻ���, a.ҽ����, b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
        From ������Ϣ A, ������� B
        Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.����id = n_¼����
        Union All
        Select /*+Rule */
         1 As ����id, a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��, a.�ֻ���,
         a.ҽ����, b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
        From ������Ϣ A, ������� B
        Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.����� = n_¼����
        Union All
        Select /*+Rule */
         1 As ����id, a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��, a.�ֻ���,
         a.ҽ����, b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
        From ������Ϣ A, ������� B
        Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.סԺ�� = n_¼����
        Union All
        Select /*+Rule */
         1 As ����id, a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��, a.�ֻ���,
         a.ҽ����, b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
        From ������Ϣ A, ������� B
        Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.ҽ���� = v_����
        Union All
        Select 0 As ����id, -null, '[�²���]', Null, Null, -null, Null, Null, Null, Null, Null, Null, Null, Null,
               To_Date(Null)
        From Dual
        Order By ����id;
    Else
      Open Output_Out For
        Select 1 As ����id, a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��,
               a.�ֻ���, a.ҽ����, b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
        From ������Ϣ A, ������� B
        Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And (a.���֤�� = v_����)
        Union All
        Select 1 As ����id, a.*
        From (Select /*+Rule */
               Distinct a.����id, a.����, a.�Ա�, a.����, a.�����, To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.����, a.����, a.���֤��, a.�ֻ���,
                        a.ҽ����, b.���� As �������, a.��ͥ��ַ, Nvl(����ʱ��, �Ǽ�ʱ��) As ����ʱ��
               From ������Ϣ A, ������� B
               Where a.���� = b.���(+) And a.ͣ��ʱ�� Is Null And a.���� Like v_���� || '%'
               Order By ����ʱ�� Desc) A
        Where Rownum < 101
        Union All
        Select 0 As ����id, -null, '[�²���]', Null, Null, -null, Null, Null, Null, Null, Null, Null, Null, Null,
               To_Date(Null)
        From Dual
        Order By ����id;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Patlistbyname;

  -----------------------------------------------------
  --��ȡ���µľ���״̬
  -----------------------------------------------------
  Procedure Getvisitstate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    ��������¼.Id%Type; --����ID
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select Max(Decode(Nvl(c.ִ��״̬, 0), 0, 0, 1)) As ����״̬
      From ��������¼ A, ��������¼ B, ���˹Һż�¼ C
      Where a.Id = b.����id And a.�Һ�id = c.Id(+) And b.Id = n_Id;
  End Getvisitstate;

  -----------------------------------------------------
  --��ȡ���˷���������Ϣ
  -----------------------------------------------------
  Procedure Load_Levelinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    ��������¼.Id%Type; --����ID
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select ID, ����id, ����id, ���ַ�����ֵ, ���ֽ������, ���鼶�� From ���ﲡ������ Where ����id = n_Id;
  End Load_Levelinfo;

  -----------------------------------------------------
  --��ȡ���˷���ָ����Ϣ
  -----------------------------------------------------
  Procedure Load_Rulesinfo
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    ��������¼.Id%Type; --����ID
  Begin
    j_In := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select a.����id, b.����id, a.ָ��id, a.ָ�����ı�
      From ���ﲡ������ָ�� A, ���ﲡ������ B
      Where a.����id = b.Id And b.����id = n_Id;
  End Load_Rulesinfo;

  -----------------------------------------------------
  --��ȡ���˷����¼����
  -----------------------------------------------------
  Procedure Load_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    Jsonobj Pljson;
    j_In    Pljson;
    n_Id    ��������¼.Id%Type; --����ID
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
    Open Output_Out For
      Select a.����id, Nvl(e.����, a.����) As ����, Nvl(e.�Ա�, a.�Ա�) As �Ա�, Nvl(e.����, a.����) As ����,
             To_Char(a.��������, 'yyyy-MM-dd') As ��������, a.���֤��, a.����, a.����, a.��ͥ��ַ, d.���� As �������, a.ҽ����, a.�ֻ���,
             b. ID As ����id, b. ����id, b. ��������, b. ������ֵ, b. ���䵥λ, b. �Һ�id, b. ���鼶��,
             To_Char(b. ��Ժʱ��, 'yyyy-MM-dd HH24:mi') As ��Ժʱ��, b. ����, b. �Ƿ�������Ա, b. ��ͬ��Ա, b. ������Դ, b. ������ʷ, b. ��ʶ״̬,
             b. �Ƿ��������, b. ������������, b. �Ƿ񸴺���, b. ��ע, b. �Ǽ��� As ����Ǽ���, b. �Ǽ�ʱ�� As ����Ǽ�ʱ��, c.�޸�˵��, c.Id As ����id, c.�������,
             c.�Զ����鼶��, c.�������id, c.�����������, c.����ѹ, c.����ѹ, c.����, c.ָ�����Ͷ�, c.����, c.Ѫ��, c.Ѫ��,
             To_Char(c.��������ʱ��, 'yyyy-MM-dd HH24:mi') As ��������ʱ��, c.�Ǽ���, c.�Ǽ�ʱ��, c.�˹����鼶��, c.�˹�����˵��, c.����Ƶ��, b. �Ƿ���ɫͨ��
      From ������Ϣ A, ��������¼ B, ��������¼ C, ������� D, ���˹Һż�¼ E
      Where a.����id = b.����id And b.Id = c.����id And b.�Һ�id = e.Id(+) And a.���� = d.���(+) And c.Id = n_Id;
  End Load_Pretriage;

  -----------------------------------------------------
  --��ȡ�������˷�����Ϣ
  -----------------------------------------------------
  Procedure Get_Patidetail
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    Jsonobj   Pljson;
    j_In    Pljson;
    n_Id      ��������¼.����id%Type;
    n_Max��� ��������¼.�������%Type;
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
  
    Select Max(�������) Into n_Max��� From ��������¼ Where ����id = n_Id;
  
    Open Output_Out For
      Select a.Id ����id, a.�������, a.�Զ����鼶�� || '��' As �Զ����鼶��, a.�˹����鼶�� || '��' As �˹����鼶��,
             '��' || a.������� || '�η���    �Զ�������' || a.�Զ����鼶�� || '����' ||
              Decode(a.�˹����鼶��, '', '', '    �˹�������' || a.�˹����鼶�� || '����') ||
              Decode(n_Max���, a.�������,
                     Decode(Nvl(b.���鼶�� || '', '0'), Nvl(b.���ﲡ�鼶�� || '', '0'), '',
                             '    �޶����鼶��' || Nvl(b.���鼶�� || '', '0') || '����')) || '    ����ʱ�䣺' ||
              To_Char(a.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:mi') As �������
      From ��������¼ A, ��������¼ B
      Where a.����id = b.Id And ����id = n_Id
      Order By ������� Desc;
  End Get_Patidetail;

  -----------------------------------------------------
  --��ȡ�����б��嵥
  -----------------------------------------------------
  Procedure Get_Patlist
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    Jsonobj    Pljson;
    j_In    Pljson;
    d_��ʼʱ�� ��������¼.�Ǽ�ʱ��%Type;
    d_����ʱ�� ��������¼.�Ǽ�ʱ��%Type;
    v_����״̬ Varchar2(10);
    n_�ѳ�ʱ   Number(2); -- =1 �������ѳ�ʱ����
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    d_��ʼʱ�� := To_Date(Jsonobj.Get_String('begin'), 'yyyy-mm-dd hh24:mi:ss');
    d_����ʱ�� := To_Date(Jsonobj.Get_String('end'), 'yyyy-mm-dd hh24:mi:ss');
    v_����״̬ := Jsonobj.Get_String('state');
    n_�ѳ�ʱ   := Nvl(To_Number(Jsonobj.Get_String('timeout')), 0);
  
    If n_�ѳ�ʱ = 1 Then
    
      Open Output_Out For
        Select b.����id, b.Id �������, Nvl(d.����, e.����) As ����, Nvl(d.�Ա�, e.�Ա�) As �Ա�, Nvl(d.����, e.����) As ����,
               To_Char(b.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:mi') As �Ǽ�ʱ��, b.�Ǽ��� ���ﻤʿ, b.���鼶�� || '��' As ���鼶��,
               Decode(Nvl(d.ִ��״̬, 0), 0, 0, 1) As ����״̬, b.�Ƿ���ɫͨ��
        From ��������¼ B, ���ﲡ�鼶�� C, ���˹Һż�¼ D, ������Ϣ E
        Where b.����id = e.����id And b.�Һ�id = d.Id(+) And b.���鼶�� = c.��� And b.�Ǽ�ʱ�� >= d_��ʼʱ�� And b.�Ǽ�ʱ�� < d_����ʱ�� And
              Decode(Nvl(d.ִ��״̬, 0), 0, 0, 1) In
              (Select Column_Value From Table(Cast(f_Str2list(v_����״̬) As t_Strlist))) And
              (c.�ٴ�����ʱ�� Is Not Null And (b.�Ǽ�ʱ�� + (Nvl(c.�ٴ�����ʱ��, 0) / 24 / 60)) < Sysdate);
    Else
      Open Output_Out For
        Select b.����id, b.Id �������, Nvl(d.����, e.����) As ����, Nvl(d.�Ա�, e.�Ա�) As �Ա�, Nvl(d.����, e.����) As ����,
               To_Char(b.�Ǽ�ʱ��, 'yyyy-MM-dd HH24:mi') As �Ǽ�ʱ��, b.�Ǽ��� ���ﻤʿ, b.���鼶�� || '��' As ���鼶��,
               Decode(Nvl(d.ִ��״̬, 0), 0, 0, 1) As ����״̬, b.�Ƿ���ɫͨ��
        From ��������¼ B, ���˹Һż�¼ D, ������Ϣ E
        Where b.����id = e.����id And b.�Һ�id = d.Id(+) And b.�Ǽ�ʱ�� >= d_��ʼʱ�� And b.�Ǽ�ʱ�� < d_����ʱ�� And
              Decode(Nvl(d.ִ��״̬, 0), 0, 0, 1) In
              (Select Column_Value From Table(Cast(f_Str2list(v_����״̬) As t_Strlist)));
    End If;
  End Get_Patlist;
  -----------------------------------------------------
  --��ʹ�ּ�����
  -----------------------------------------------------
  Procedure Get_Pain_Num_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    --  input
    --    pain        C 1 ��ʹ�ȼ�
    Jsonobj    Pljson;
    j_In    Pljson;
    v_��ʹ�ȼ� Varchar2(200);
    v_Out      Varchar2(200);
    v_���˵ȼ� Varchar2(200);
    v_���˷��� Varchar2(200);
    v_����     Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_��ʹ�ȼ� := Jsonobj.Get_String('pain');
  
    Select b_Emergency_Rating.Is_Pain_Num_Rating(v_��ʹ�ȼ�) Into v_Out From Dual;
  
    v_���˵ȼ� := Substr(v_Out, 1, Instr(v_Out, ':', 1, 1) - 1);
    v_���˷��� := Substr(v_Out, Instr(v_Out, ':', 1, 1) + 1, Instr(v_Out, ':', 1, 2) - Instr(v_Out, ':', 1, 1) - 1);
    v_����     := Substr(v_Out, Instr(v_Out, ':', 1, 2) + 1);
  
    Open Output_Out For
      Select v_���˵ȼ� As ���˵ȼ�, v_���˷��� As ���˷���, v_���� As ���� From Dual;
  End Get_Pain_Num_Rating;

  -----------------------------------------------------
  --�������ַּ�����
  -----------------------------------------------------
  Procedure Get_Coma_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    --  input
    --    open_reaction            C 1 ���۷�Ӧָ��id��ָ��������
    --    language_reaction        C 1 ���Է�Ӧָ��id��ָ��������
    --    activity_reaction        C 1 ���Ӧָ��id��ָ��������
    Jsonobj    Pljson;
    j_In    Pljson;
    v_���۷�Ӧ Varchar2(200);
    v_���Է�Ӧ Varchar2(200);
    v_���Ӧ Varchar2(200);
    v_Out      Varchar2(200);
  
    v_���˵ȼ� Varchar2(200);
    v_���˷��� Varchar2(200);
    v_����     Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    v_���۷�Ӧ := Jsonobj.Get_String('open_reaction');
    v_���Է�Ӧ := Jsonobj.Get_String('language_reaction');
    v_���Ӧ := Jsonobj.Get_String('activity_reaction');
  
    Select b_Emergency_Rating.Is_Coma_Rating(v_���۷�Ӧ, v_���Է�Ӧ, v_���Ӧ) Into v_Out From Dual;
  
    v_���˵ȼ� := Substr(v_Out, 1, Instr(v_Out, ':', 1, 1) - 1);
    v_���˷��� := Substr(v_Out, Instr(v_Out, ':', 1, 1) + 1, Instr(v_Out, ':', 1, 2) - Instr(v_Out, ':', 1, 1) - 1);
    v_����     := Substr(v_Out, Instr(v_Out, ':', 1, 2) + 1);
  
    Open Output_Out For
      Select v_���˵ȼ� As ���˵ȼ�, v_���˷��� As ���˷���, v_���� As ���� From Dual;
  End Get_Coma_Rating;

  -----------------------------------------------------
  --�͹����۷ּ��������˺Ͷ�ͯ����
  -----------------------------------------------------
  Procedure Get_Objective_Rating
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    --  input
    --    agenum                   C 1 ����
    --    ageunit                  C 1 ���䵥λ
    --    indexid_describe         C 1 ָ��id��ָ�����������ɶ����
    Jsonobj    Pljson;
    j_In    Pljson;
    n_����     Number;
    v_���䵥λ Varchar2(200);
    v_ָ����Ϣ Varchar2(200);
    v_Out      Varchar2(200);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_����     := Nvl(Zl_To_Number(Jsonobj.Get_String('agenum')), 0);
    v_���䵥λ := Jsonobj.Get_String('ageunit');
    v_ָ����Ϣ := Jsonobj.Get_String('indexid_describe');
  
    Select b_Emergency_Rating.Is_Objective_Rating(n_����, v_���䵥λ, v_ָ����Ϣ) Into v_Out From Dual;
  
    Open Output_Out For
      Select v_Out As ���˵ȼ� From Dual;
  End Get_Objective_Rating;
  -----------------------------------------------------
  --��ȡ��ͯ��������
  -----------------------------------------------------
  Procedure Get_Childmaxage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Output_Out := Nvl(zl_GetSysParameter('��ͯ����綨����'), 0);
  End Get_Childmaxage;

  -----------------------------------------------------
  --��ȡ����ȼ�
  -----------------------------------------------------
  Procedure Get_Level
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.���, a.����, a.���س̶�, a.�ٴ�����ʱ��, a.���߱�ʶ��ɫ, Null As ȱʡ
      From ���ﲡ�鼶�� A
      Order By a.���;
  End Get_Level;

  -----------------------------------------------------
  --��ȡ�������
  -----------------------------------------------------
  Procedure Get_Dept
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.Id, a.����, a.����, a.����, Null As ȱʡ
      From ���ű� A, �ٴ����� B
      Where a.Id = b.����id And b.�������� = '20' And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)
      Order By a.����;
  End Get_Dept;

  -----------------------------------------------------
  --��ȡ�˹���������
  -----------------------------------------------------
  Procedure Get_Rules
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select a.Id, a.����, a.ָ������, a.������Ⱥ, a.���鼶�� From �����˹��������� A Order By ID, ���鼶��;
  End Get_Rules;

  -----------------------------------------------------
  --���ݳ������ڷ�������
  -----------------------------------------------------
  Procedure Get_Datetoage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    --��Σ�Json_In:��ʽ
    --  input
    --    birthday        C 1 �������� yyyy-mm-dd
    Jsonobj    Pljson;
    j_In    Pljson;
    d_�������� Date;
    v_����     Varchar2(50);
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    d_�������� := To_Date(Jsonobj.Get_String('birthday'), 'yyyy-mm-dd hh24:mi:ss');
    Select Zl_Age_Calc(0, d_��������, Sysdate) Into v_���� From Dual;
  
    Output_Out := v_����;
  End Get_Datetoage;

  -----------------------------------------------------
  --��ȡ�Ա��������
  -----------------------------------------------------
  Procedure Get_Sexbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����, ����, ����, Nvl(ȱʡ��־, 0) As ȱʡ From �Ա� Order By Nvl(ȱʡ��־, 0) Desc, ����;
  End Get_Sexbase;

  -----------------------------------------------------
  --��ȡ�����������
  -----------------------------------------------------
  Procedure Get_Nationbase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����, ����, ����, Nvl(ȱʡ��־, 0) As ȱʡ From ���� Order By Nvl(ȱʡ��־, 0) Desc, ����;
  End Get_Nationbase;

  -----------------------------------------------------
  --��ȡ��������ָ��
  -----------------------------------------------------
  Procedure Get_Scorebase
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ID, ָ������, ֵ��Χ, ����id, ֵ��λ From ��������ָ�� Order By ID;
  End Get_Scorebase;

  -----------------------------------------------------
  --��ȡ��������
  -----------------------------------------------------
  Procedure Get_Paticc
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select b.���� ����, a.����, a.����, a.����
      From ���ﳣ������ A, ���ﳣ������ B
      Where a.�ϼ� = b.���� And a.�ϼ� Is Not Null And b.�ϼ� Is Null
      Order By b.����;
  End Get_Paticc;

  -----------------------------------------------------
  --��ȡ������Դ
  -----------------------------------------------------
  Procedure Get_Patifrom
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����, ����, ȱʡ��־ As ȱʡ From ���ﲡ����Դ Order By Nvl(ȱʡ��־, 0) Desc, ����;
  End Get_Patifrom;

  -----------------------------------------------------
  --��ȡ������ʶ״̬
  -----------------------------------------------------
  Procedure Get_Patistate
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����, ����, ȱʡ��־ As ȱʡ From ������ʶ״̬ Order By Nvl(ȱʡ��־, 0) Desc, ����;
  End Get_Patistate;

  -----------------------------------------------------
  --��ȡ������ͬ��Ա
  -----------------------------------------------------
  Procedure Get_Entourage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����, ����, ȱʡ��־ As ȱʡ From ������ͬ��Ա Order By Nvl(ȱʡ��־, 0) Desc, ����;
  End Get_Entourage;

  -----------------------------------------------------
  --��ȡ���ﳣ������ʷ
  -----------------------------------------------------
  Procedure Get_Dishistory
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
  Begin
    Open Output_Out For
      Select ����, ����, 0 As ȱʡ From ���ﳣ������ʷ Order By ����;
  End Get_Dishistory;

  -----------------------------------------------------
  --��ȡ���ݿ�ϵͳʱ��
  -----------------------------------------------------
  Procedure Get_Now_Time
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Output_Out := To_Char(Sysdate, 'yyyy-MM-dd HH24:mi');
  End Get_Now_Time;
End Pkg_Pretriage_Dql;
/



--145003:��͢��,2019-10-15,����ģ�鼱��Ԥ����﹤��վ
Create Or Replace Package Pkg_Pretriage_Dml As

  -----------------------------------------------------
  --������˾����¼����ɫͨ��״̬
  -----------------------------------------------------
  Procedure Change_Greenchannel
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );
  -----------------------------------------------------
  --ɾ�����˾����¼
  -----------------------------------------------------
  Procedure Del_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --����Һ���������
  -----------------------------------------------------
  Procedure Register_Unlock
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --�������µĹҺŰ���
  -----------------------------------------------------
  Procedure Register_Update
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  );

  -----------------------------------------------------
  --���没�˷�����Ϣ
  -----------------------------------------------------
  Procedure Save_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  );
End Pkg_Pretriage_Dml;
/
Create Or Replace Package Body Pkg_Pretriage_Dml As
  -----------------------------------------------------
  --������˾����¼����ɫͨ��״̬
  -----------------------------------------------------
  Procedure Change_Greenchannel
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) As
    --���ܣ���ǻ�ȡ��������ɫͨ��
    Jsonobj Pljson;
    j_In    Pljson;
  
    n_Id           ��������¼.Id%Type; --����ID
    n_�Ƿ���ɫͨ�� ��������¼.�Ƿ���ɫͨ��%Type; --�Ƿ���ɫͨ��
    n_�Һ�id       ��������¼.�Һ�id %Type;
  Begin
    j_In           := Pljson(Input_In);
    Jsonobj        := j_In.Get_Pljson('input');
    n_Id           := To_Number(Jsonobj.Get_String('id'));
    n_�Ƿ���ɫͨ�� := To_Number(Jsonobj.Get_String('�Ƿ���ɫͨ��'));
  
    Select Max(�Һ�id) Into n_�Һ�id From ��������¼ Where ID = n_Id;
  
    Zl_������ɫͨ��_Edit(n_�Һ�id, n_�Ƿ���ɫͨ��);
    Output_Out := '�ɹ�';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Change_Greenchannel;

  -----------------------------------------------------
  --ɾ�����˾����¼
  -----------------------------------------------------
  Procedure Del_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    Jsonobj  Pljson;
    j_In     Pljson;
    n_Id     ��������¼.Id%Type; --����ID
    n_����id ��������¼.����id%Type;
    n_�Һ�id ��������¼.�Һ�id%Type;
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
    n_Id    := To_Number(Jsonobj.Get_String('id'));
  
    Delete From ��������¼ Where ID = n_Id Return ����id, �Һ�id Into n_����id, n_�Һ�id;
  
    Zl_Emergencyregistdel(n_�Һ�id);
  
    Delete From ������Ϣ�ӱ�
    Where ����id = n_����id And ����id = n_�Һ�id And ��Ϣ�� In ('����', '����', '����', '����ѹ', '����ѹ', 'Ѫ��');
  
    Output_Out := '�ɹ�';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Del_Pretriage;

  -----------------------------------------------------
  --����Һ���������
  -----------------------------------------------------
  Procedure Register_Unlock
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
    v_��Ա���� Varchar2(200);
    v_Temp     Varchar2(4000);
  Begin
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    Zl_�Һ����״̬_Lock(2, v_��Ա����);
    Output_Out := '�ɹ�';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Register_Unlock;

  -----------------------------------------------------
  --�������µĹҺŰ���
  -----------------------------------------------------
  Procedure Register_Update
  (
    Input_In   In Clob,
    Output_Out Out Varchar2
  ) Is
  Begin
    Zl_�ҺŰ���_Autoupdate();
    Output_Out := '�ɹ�';
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Register_Update;

  -----------------------------------------------------
  --���没�˷�����Ϣ
  -----------------------------------------------------
  Procedure Save_Pretriage
  (
    Input_In   In Clob,
    Output_Out Out Sys_Refcursor
  ) Is
    --��Σ�Json_In:��ʽ
    Jsonobj Pljson;
    j_In    Pljson;
    n_Type  Number; --1  ����,2  �޸�
  
    --������Ϣ
    v_����     ������Ϣ.����%Type;
    v_�Ա�     ������Ϣ.�Ա�%Type;
    d_�������� ������Ϣ.��������%Type;
    v_���֤�� ������Ϣ.���֤��%Type;
    v_��ϵ�绰 ������Ϣ.��ϵ�˵绰%Type;
    v_����     ������Ϣ.����%Type;
    v_����     ������Ϣ.����%Type;
    v_ҽ������ ������Ϣ.ҽ����%Type;
    v_������� �������.����%Type;
    v_��ͥ��ַ ������Ϣ.��ͥ��ַ%Type;
  
    n_����id ��������¼.Id%Type;
    n_����id ��������¼.����id%Type;
    n_�Һ�id ��������¼.�Һ�id%Type;
    n_����id ��������¼.Id%Type;
  
    --�����¼
    v_�������� ��������¼.��������%Type;
    n_������ֵ ��������¼.������ֵ%Type;
    v_���䵥λ ��������¼.���䵥λ%Type;
  
    d_��Ժʱ��     ��������¼.��Ժʱ��%Type;
    n_�Ƿ�������Ա ��������¼.�Ƿ�������Ա%Type;
    n_�Ƿ񸴺���   ��������¼.�Ƿ񸴺���%Type;
    n_�Ƿ���ɫͨ�� ��������¼.�Ƿ���ɫͨ��%Type;
  
    n_�Ƿ�������� ��������¼.�Ƿ��������%Type;
    n_������������ ��������¼.������������%Type;
    v_������Դ     ��������¼.������Դ%Type;
    v_��ͬ��Ա     ��������¼.��ͬ��Ա%Type;
    v_��ʶ״̬     ��������¼.��ʶ״̬%Type;
    v_������ʷ     ��������¼.������ʷ%Type;
    v_����         ��������¼.����%Type;
    n_���鼶��     ��������¼.���鼶��%Type;
    v_�Ǽ���       ��������¼.�Ǽ���%Type;
    d_�Ǽ�ʱ��     ��������¼.�Ǽ�ʱ��%Type;
    v_��ע         ��������¼.��ע%Type;
  
    --�����¼
    n_������� ��������¼.�������%Type;
  
    n_�������id   ��������¼.�������id%Type;
    v_����������� ��������¼.�����������%Type;
  
    d_��������ʱ�� ��������¼.��������ʱ��%Type;
    n_����ѹ       ��������¼.����ѹ%Type;
    n_����ѹ       ��������¼.����ѹ%Type;
    n_Ѫ��         ��������¼.Ѫ��%Type;
    n_ָ�����Ͷ�   ��������¼.ָ�����Ͷ�%Type;
    n_����         ��������¼.����%Type;
    n_Ѫ��         ��������¼.Ѫ��%Type;
    n_����         ��������¼.����%Type;
    n_����Ƶ��     ��������¼.����Ƶ��%Type;
  
    n_�Զ����鼶��  ��������¼.�Զ����鼶��%Type;
    n_�˹����鼶��  ��������¼.�˹����鼶��%Type;
    v_�˹�����˵��  ��������¼.�˹�����˵��%Type;
    v_�޸�˵��      ��������¼.�޸�˵��%Type;
    v_վ��          Varchar2(10);
    n_�������idold ��������¼.�������id%Type;
  
    d_Now Date;
  
    n_�����     Number(18);
    n_����       Number(5);
    v_�Ǽ��˱�� Varchar2(6);
    n_Count      Number(5);
  
    n_����id       Number(18);
    n_����id       Number(18);
    n_���ַ�����ֵ Number(5);
    v_���ֽ������ Varchar2(100);
    n_���ֵȼ�     Number(1);
  
    Jsonlist����ָ�� Pljson_List;
    Jsonlist�������� Pljson_List;
    Jsonlistitem     Pljson;
    Jsonlistitemָ�� Pljson;
  
    n_Edittmp Number(5); --0  ����  1  �޸�
  Begin
    j_In    := Pljson(Input_In);
    Jsonobj := j_In.Get_Pljson('input');
  
    n_Type           := Jsonobj.Get_String('type');
    n_����id         := Jsonobj.Get_String('����id');
    n_����id         := Nvl(To_Number(Jsonobj.Get_String('����id')), 0);
    n_�����         := Nvl(To_Number(Jsonobj.Get_String('�����')), 0);
    v_����           := Jsonobj.Get_String('����');
    v_�Ա�           := Jsonobj.Get_String('�Ա�');
    d_��������       := To_Date(Jsonobj.Get_String('��������'), 'yyyy-mm-dd');
    v_���֤��       := Jsonobj.Get_String('���֤��');
    v_��ϵ�绰       := Jsonobj.Get_String('��ϵ�绰');
    v_����           := Jsonobj.Get_String('����');
    v_ҽ������       := Jsonobj.Get_String('ҽ������');
    v_�������       := Jsonobj.Get_String('�������');
    v_��ͥ��ַ       := Jsonobj.Get_String('��ͥ��ַ');
    v_��������       := Jsonobj.Get_String('��������');
    n_������ֵ       := To_Number(Jsonobj.Get_String('������ֵ'));
    v_���䵥λ       := Jsonobj.Get_String('���䵥λ');
    d_��Ժʱ��       := To_Date(Jsonobj.Get_String('��Ժʱ��'), 'yyyy-mm-dd hh24:mi:ss');
    n_�Ƿ�������Ա   := To_Number(Jsonobj.Get_String('�Ƿ�������Ա'));
    n_�Ƿ񸴺���     := To_Number(Jsonobj.Get_String('�Ƿ񸴺���'));
    n_�Ƿ���ɫͨ��   := To_Number(Jsonobj.Get_String('�Ƿ���ɫͨ��'));
    n_�Ƿ��������   := To_Number(Jsonobj.Get_String('�Ƿ��������'));
    n_������������   := To_Number(Jsonobj.Get_String('������������'));
    v_������Դ       := Jsonobj.Get_String('������Դ');
    v_��ͬ��Ա       := Jsonobj.Get_String('��ͬ��Ա');
    v_��ʶ״̬       := Jsonobj.Get_String('��ʶ״̬');
    v_������ʷ       := Jsonobj.Get_String('������ʷ');
    v_����           := Jsonobj.Get_String('����');
    n_���鼶��       := To_Number(Jsonobj.Get_String('���鼶��'));
    v_�Ǽ���         := Jsonobj.Get_String('�Ǽ���');
    v_��ע           := Jsonobj.Get_String('��ע');
    n_�������id     := To_Number(Jsonobj.Get_String('�������id'));
    v_�����������   := Jsonobj.Get_String('�����������');
    d_��������ʱ��   := To_Date(Jsonobj.Get_String('��������ʱ��'), 'yyyy-mm-dd hh24:mi:ss');
    n_����ѹ         := To_Number(Jsonobj.Get_String('����ѹ'));
    n_����ѹ         := To_Number(Jsonobj.Get_String('����ѹ'));
    n_Ѫ��           := To_Number(Jsonobj.Get_String('Ѫ��'));
    n_ָ�����Ͷ�     := To_Number(Jsonobj.Get_String('ָ�����Ͷ�'));
    n_����           := To_Number(Jsonobj.Get_String('����'));
    n_Ѫ��           := To_Number(Jsonobj.Get_String('Ѫ��'));
    n_����           := To_Number(Jsonobj.Get_String('����'));
    n_����Ƶ��       := To_Number(Jsonobj.Get_String('����Ƶ��'));
    n_�Զ����鼶��   := To_Number(Jsonobj.Get_String('�Զ����鼶��'));
    n_�˹����鼶��   := To_Number(Jsonobj.Get_String('�˹����鼶��'));
    v_�˹�����˵��   := Jsonobj.Get_String('�˹�����˵��');
    v_�޸�˵��       := Jsonobj.Get_String('�޸�˵��');
    v_�Ǽ��˱��     := Jsonobj.Get_String('�Ǽ��˱��');
    v_վ��           := Jsonobj.Get_String('վ��');
    v_����           := Jsonobj.Get_String('����');
    Jsonlist����ָ�� := Jsonobj.Get_Pljson_List('����ָ��');
    Jsonlist�������� := Jsonobj.Get_Pljson_List('��������');
  
    n_Edittmp := 0;
  
    --��ȡ�Ǽ��˱��
    If v_�Ǽ��˱�� Is Null Then
      Select Max(���) Into v_�Ǽ��˱�� From ��Ա�� Where ���� = v_�Ǽ���;
    End If;
    --��ȡ�������
    If v_������� Is Not Null Then
      Select Max(���) Into n_���� From ������� Where ���� = v_�������;
    End If;
  
    Select Sysdate Into d_Now From Dual;
    d_�Ǽ�ʱ�� := d_Now;
  
    --����ʱ���²���
    If n_Type = 1 Then
      Select ��������¼_Id.Nextval Into n_����id From Dual;
    End If;
  
    --����ID�������²���
    Select ��������¼_Id.Nextval Into n_����id From Dual;
  
    --���������
    --�ȴ����������Ϣ
    If n_Type = 1 Then
      If n_����id > 0 Then
        n_Edittmp := 1;
      Else
        If v_���֤�� Is Not Null And v_���� = '�й�' Then
          n_Count := Nvl(zl_GetSysParameter(279), 0);
          If n_Count = 1 Then
            Select Max(����id) Into n_����id From ������Ϣ Where ���֤�� = v_���֤��;
            If n_����id > 0 Then
              n_Edittmp := 1;
            End If;
          End If;
        End If;
      End If;
    
      If n_Edittmp = 0 Then
        Select ������Ϣ_Id.Nextval Into n_����id From Dual;
        n_����� := Nextno(3);
        Zl_������Ϣ_Insert(n_����id, n_�����, Null, Null, v_����, v_�Ա�, v_��������, d_��������, Null, v_���֤��, Null, Null, v_����, v_����,
                       Null, Null, v_��ͥ��ַ, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null,
                       Null, n_����, Sysdate, Null, Null, v_�Ǽ��˱��, v_�Ǽ���, v_ҽ������, Null, Null, Null, Null, Null, Null, Null,
                       v_��ϵ�绰);
      Else
        If n_����� = 0 Then
          Select Nvl(Max(�����), 0) Into n_����� From ������Ϣ Where ����id = n_����id;
          If n_����� = 0 Then
            n_����� := Nextno(3);
          End If;
        End If;
        Update ������Ϣ
        Set ����� = n_�����, ���� = Nvl(v_����, ����), �Ա� = Nvl(v_�Ա�, �Ա�), ���� = Nvl(v_��������, ����), �������� = Nvl(d_��������, ��������),
            ���֤�� = Nvl(v_���֤��, ���֤��), ���� = Nvl(v_����, ����), ��ͥ��ַ = Nvl(v_��ͥ��ַ, ��ͥ��ַ), ���� = Nvl(n_����, ����),
            ҽ���� = Nvl(v_ҽ������, ҽ����), �ֻ��� = Nvl(v_��ϵ�绰, �ֻ���), ���� = Nvl(v_����, ����)
        Where ����id = n_����id;
      End If;
    Else
      If n_����id Is Not Null Then
        Select Max(Nvl(����id, 0)), Max(Nvl(�Һ�id, 0)), Max(Nvl(�������id, 0))
        Into n_����id, n_�Һ�id, n_�������idold
        From ��������¼
        Where ID = n_����id;
      
        --�޸Ĳ���������Ϣ
        /*Select Max(�����) Into n_����� From ������Ϣ Where ����id = n_����id;
        Update ������Ϣ
        Set ����� = n_�����, ���� = Nvl(v_����, ����), �Ա� = Nvl(v_�Ա�, �Ա�), ���� = Nvl(v_��������, ����), �������� = Nvl(d_��������, ��������),
            ���֤�� = Nvl(v_���֤��, ���֤��), ���� = Nvl(v_����, ����), ��ͥ��ַ = Nvl(v_��ͥ��ַ, ��ͥ��ַ), ���� = Nvl(n_����, ����),
            ҽ���� = Nvl(v_ҽ������, ҽ����), �ֻ��� = Nvl(v_��ϵ�绰, �ֻ���)
        Where ����id = n_����id;*/
      End If;
    End If;
  
    If n_Type = 1 Then
      --����Һ�id
    
      n_�Һ�id := Zl_Emergencyregist(n_����id, n_�������id, v_վ��, n_�Ƿ���ɫͨ��);
    
      Insert Into ��������¼
        (ID, ����id, ��������, ������ֵ, ���䵥λ, �Һ�id, ���鼶��, ��Ժʱ��, ����, �Ƿ�������Ա, ��ͬ��Ա, ������Դ, ������ʷ, ��ʶ״̬, �Ƿ��������, ������������, �Ƿ񸴺���, ��ע,
         �Ǽ���, �Ǽ�ʱ��, ���ﲡ�鼶��, �Ƿ���ɫͨ��, �������id)
      Values
        (n_����id, n_����id, v_��������, n_������ֵ, v_���䵥λ, n_�Һ�id, n_���鼶��, d_��Ժʱ��, v_����, n_�Ƿ�������Ա, v_��ͬ��Ա, v_������Դ, v_������ʷ, v_��ʶ״̬,
         n_�Ƿ��������, n_������������, n_�Ƿ񸴺���, v_��ע, v_�Ǽ���, d_�Ǽ�ʱ��, n_���鼶��, n_�Ƿ���ɫͨ��, n_�������id);
    Else
      If n_�������idold <> n_�������id Then
        Zl_Emergencyregistredo(n_�Һ�id, n_�������id, v_վ��);
      End If;
      Update ��������¼
      Set �������� = v_��������, ������ֵ = n_������ֵ, ���䵥λ = v_���䵥λ, �Һ�id = n_�Һ�id, ���鼶�� = n_���鼶��, ��Ժʱ�� = d_��Ժʱ��, ���� = v_����,
          �Ƿ�������Ա = n_�Ƿ�������Ա, ��ͬ��Ա = v_��ͬ��Ա, ������Դ = v_������Դ, ������ʷ = v_������ʷ, ��ʶ״̬ = v_��ʶ״̬, �Ƿ�������� = n_�Ƿ��������,
          ������������ = n_������������, �Ƿ񸴺��� = n_�Ƿ񸴺���, ��ע = v_��ע, ���ﲡ�鼶�� = n_���鼶��, �Ƿ���ɫͨ�� = n_�Ƿ���ɫͨ��, �Ǽ�ʱ�� = d_�Ǽ�ʱ��,
          �������id = n_�������id
      Where ID = n_����id;
    End If;
  
    If n_Type = 1 Then
      n_������� := 1;
    Else
      Select Max(�������) + 1 Into n_������� From ��������¼ Where ����id = n_����id;
    End If;
  
    Insert Into ��������¼
      (ID, ����id, �������, �Զ����鼶��, �������id, �����������, ����ѹ, ����ѹ, ����, ָ�����Ͷ�, ����, Ѫ��, Ѫ��, ��������ʱ��, �Ǽ���, �Ǽ�ʱ��, �˹����鼶��, �˹�����˵��, ����Ƶ��,
       �޸�˵��)
    Values
      (n_����id, n_����id, n_�������, n_�Զ����鼶��, n_�������id, v_�����������, n_����ѹ, n_����ѹ, n_����, n_ָ�����Ͷ�, n_����, n_Ѫ��, n_Ѫ��, d_��������ʱ��,
       v_�Ǽ���, d_�Ǽ�ʱ��, n_�˹����鼶��, v_�˹�����˵��, n_����Ƶ��, v_�޸�˵��);
  
    Delete From ������Ϣ�ӱ�
    Where ����id = n_����id And ����id = n_�Һ�id And ��Ϣ�� In ('����', '����', '����', '����ѹ', '����ѹ', 'Ѫ��');
  
    If n_���� Is Not Null Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
        Select n_����id, n_�Һ�id, '����', To_Char(n_����) From Dual;
    End If;
  
    If n_����Ƶ�� Is Not Null Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
        Select n_����id, n_�Һ�id, '����', To_Char(n_����Ƶ��) From Dual;
    End If;
  
    If n_���� Is Not Null Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
        Select n_����id, n_�Һ�id, '����', To_Char(n_����) From Dual;
    End If;
  
    If n_����ѹ Is Not Null Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
        Select n_����id, n_�Һ�id, '����ѹ', To_Char(n_����ѹ) From Dual;
    End If;
  
    If n_����ѹ Is Not Null Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
        Select n_����id, n_�Һ�id, '����ѹ', To_Char(n_����ѹ) From Dual;
    End If;
  
    If n_Ѫ�� Is Not Null Then
      Insert Into ������Ϣ�ӱ�
        (����id, ����id, ��Ϣ��, ��Ϣֵ)
        Select n_����id, n_�Һ�id, 'Ѫ��', To_Char(n_Ѫ��) From Dual;
    End If;
  
    For I In 1 .. Jsonlist��������.Count Loop
      Jsonlistitem   := Pljson();
      Jsonlistitem   := Pljson(Jsonlist��������.Get(I));
      n_����id       := To_Number(Jsonlistitem.Get_String('����ID'));
      n_���ַ�����ֵ := To_Number(Jsonlistitem.Get_String('���ַ�����ֵ'));
      v_���ֽ������ := Jsonlistitem.Get_String('���ֽ������');
      n_���ֵȼ�     := To_Number(Jsonlistitem.Get_String('���ֵȼ�'));
      Select ���ﲡ������_Id.Nextval Into n_����id From Dual;
    
      Insert Into ���ﲡ������
        (ID, ����id, ����id, ���ַ�����ֵ, ���ֽ������, ���鼶��)
      Values
        (n_����id, n_����id, n_����id, n_���ַ�����ֵ, v_���ֽ������, n_���ֵȼ�);
    
      For I In 1 .. Jsonlist����ָ��.Count Loop
        Jsonlistitemָ�� := Pljson();
        Jsonlistitemָ�� := Pljson(Jsonlist����ָ��.Get(I));
        If n_����id = To_Number(Jsonlistitemָ��.Get_String('����ID')) Then
          Insert Into ���ﲡ������ָ��
            (����id, ָ��id, ָ�����ı�)
          Values
            (n_����id, To_Number(Jsonlistitemָ��.Get_String('ָ��ID')), Jsonlistitemָ��.Get_String('ָ�����ı�'));
        End If;
      End Loop;
    End Loop;
  
    Open Output_Out For
      Select n_����id As ����id, n_����id As ����id, n_����id As ����id From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Pretriage;

End Pkg_Pretriage_Dml;
/
