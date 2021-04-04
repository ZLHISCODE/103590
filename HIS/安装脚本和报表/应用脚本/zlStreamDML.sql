------------------------------------------------------------------------------------------------------------------------------------------

Create Or Replace Procedure ���˷��û���_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End ���˷��û���%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old ���˷��û���%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In ���˷��û���%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update ���˷��û���
    Set Ӧ�ս�� = Nvl(Ӧ�ս��, 0) + Nvl(r_Dat.Ӧ�ս��, 0) * n_Sign,
        ʵ�ս�� = Nvl(ʵ�ս��, 0) + Nvl(r_Dat.ʵ�ս��, 0) * n_Sign,
        ���ʽ�� = Nvl(���ʽ��, 0) + Nvl(r_Dat.���ʽ��, 0) * n_Sign
    Where ���� = r_Dat.���� And Nvl(���˲���id, 0) = Nvl(r_Dat.���˲���id, 0) And
          Nvl(���˿���id, 0) = Nvl(r_Dat.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Dat.��������id, 0) And
          Nvl(ִ�в���id, 0) = Nvl(r_Dat.ִ�в���id, 0) And Nvl(������Ŀid, 0) = Nvl(r_Dat.������Ŀid, 0) And
          ��Դ;�� = r_Dat.��Դ;�� And ���ʷ��� = r_Dat.���ʷ���;
    If Sql%Rowcount = 0 Then
      Insert Into ���˷��û���
        (����, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���ʷ���, Ӧ�ս��, ʵ�ս��,
         ���ʽ��)
      Values
        (r_Dat.����, r_Dat.���˲���id, r_Dat.���˿���id, r_Dat.��������id, r_Dat.ִ�в���id, r_Dat.������Ŀid,
         r_Dat.��Դ;��, r_Dat.���ʷ���, Nvl(r_Dat.Ӧ�ս��, 0) * n_Sign, Nvl(r_Dat.ʵ�ս��, 0) * n_Sign,
         Nvl(r_Dat.���ʽ��, 0) * n_Sign);
    End If;
    Delete ���˷��û���
    Where ���� = r_Dat.���� And Nvl(���˲���id, 0) = Nvl(r_Dat.���˲���id, 0) And
          Nvl(���˿���id, 0) = Nvl(r_Dat.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Dat.��������id, 0) And
          Nvl(ִ�в���id, 0) = Nvl(r_Dat.ִ�в���id, 0) And Nvl(������Ŀid, 0) = Nvl(r_Dat.������Ŀid, 0) And
          ��Դ;�� = r_Dat.��Դ;�� And ���ʷ��� = r_Dat.���ʷ��� And Nvl(Ӧ�ս��, 0) = 0 And Nvl(ʵ�ս��, 0) = 0 And
          Nvl(���ʽ��, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
    i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getdate(r_End.����);
    i_Result := l_Lcr_App.Get_Value('new', '���˲���ID', 'Y').Getnumber(r_End.���˲���id);
    i_Result := l_Lcr_App.Get_Value('new', '���˿���ID', 'Y').Getnumber(r_End.���˿���id);
    i_Result := l_Lcr_App.Get_Value('new', '��������ID', 'Y').Getnumber(r_End.��������id);
    i_Result := l_Lcr_App.Get_Value('new', 'ִ�в���ID', 'Y').Getnumber(r_End.ִ�в���id);
    i_Result := l_Lcr_App.Get_Value('new', '������ĿID', 'Y').Getnumber(r_End.������Ŀid);
    i_Result := l_Lcr_App.Get_Value('new', '��Դ;��', 'Y').Getnumber(r_End.��Դ;��);
    i_Result := l_Lcr_App.Get_Value('new', '���ʷ���', 'Y').Getnumber(r_End.���ʷ���);
    i_Result := l_Lcr_App.Get_Value('new', 'Ӧ�ս��', 'Y').Getnumber(r_End.Ӧ�ս��);
    i_Result := l_Lcr_App.Get_Value('new', 'ʵ�ս��', 'Y').Getnumber(r_End.ʵ�ս��);
    i_Result := l_Lcr_App.Get_Value('new', '���ʽ��', 'Y').Getnumber(r_End.���ʽ��);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getdate(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '���˲���ID', 'N').Getnumber(r_Old.���˲���id);
    i_Result := l_Lcr_App.Get_Value('old', '���˿���ID', 'N').Getnumber(r_Old.���˿���id);
    i_Result := l_Lcr_App.Get_Value('old', '��������ID', 'N').Getnumber(r_Old.��������id);
    i_Result := l_Lcr_App.Get_Value('old', 'ִ�в���ID', 'N').Getnumber(r_Old.ִ�в���id);
    i_Result := l_Lcr_App.Get_Value('old', '������ĿID', 'N').Getnumber(r_Old.������Ŀid);
    i_Result := l_Lcr_App.Get_Value('old', '��Դ;��', 'N').Getnumber(r_Old.��Դ;��);
    i_Result := l_Lcr_App.Get_Value('old', '���ʷ���', 'N').Getnumber(r_Old.���ʷ���);
    i_Result := l_Lcr_App.Get_Value('old', 'Ӧ�ս��', 'N').Getnumber(r_Old.Ӧ�ս��);
    i_Result := l_Lcr_App.Get_Value('old', 'ʵ�ս��', 'N').Getnumber(r_Old.ʵ�ս��);
    i_Result := l_Lcr_App.Get_Value('old', '���ʽ��', 'N').Getnumber(r_Old.���ʽ��);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End ���˷��û���_Dml;
/

Create Or Replace Procedure ���˹ҺŻ���_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End ���˹ҺŻ���%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old ���˹ҺŻ���%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In ���˹ҺŻ���%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update ���˹ҺŻ���
    Set �ѹ��� = Nvl(�ѹ���, 0) + Nvl(r_Dat.�ѹ���, 0) * n_Sign, ��Լ�� = Nvl(��Լ��, 0) + Nvl(r_Dat.��Լ��, 0) * n_Sign
    Where ���� = r_Dat.���� And ����id = Nvl(r_Dat.����id, 0) And ��Ŀid = Nvl(r_Dat.��Ŀid, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Dat.ҽ������, '-') And Nvl(ҽ��id, 0) = Nvl(r_Dat.ҽ��id, 0);
    If Sql%Rowcount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, �ѹ���, ��Լ��)
      Values
        (r_Dat.����, r_Dat.����id, r_Dat.��Ŀid, r_Dat.ҽ������, r_Dat.ҽ��id, Nvl(r_Dat.�ѹ���, 0) * n_Sign,
         Nvl(r_Dat.��Լ��, 0) * n_Sign);
    End If;
    Delete ���˹ҺŻ���
    Where ���� = r_Dat.���� And ����id = Nvl(r_Dat.����id, 0) And ��Ŀid = Nvl(r_Dat.��Ŀid, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Dat.ҽ������, '-') And Nvl(ҽ��id, 0) = Nvl(r_Dat.ҽ��id, 0) And
          Nvl(�ѹ���, 0) = 0 And Nvl(��Լ��, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
   i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getdate(r_End.����);
   i_Result := l_Lcr_App.Get_Value('new', '����ID', 'Y').Getnumber(r_End.����id);
   i_Result := l_Lcr_App.Get_Value('new', '��ĿID', 'Y').Getnumber(r_End.��Ŀid);
   i_Result := l_Lcr_App.Get_Value('new', 'ҽ������', 'Y').Getvarchar2(r_End.ҽ������);
   i_Result := l_Lcr_App.Get_Value('new', 'ҽ��ID', 'Y').Getnumber(r_End.ҽ��id);
   i_Result := l_Lcr_App.Get_Value('new', '�ѹ���', 'Y').Getnumber(r_End.�ѹ���);
   i_Result := l_Lcr_App.Get_Value('new', '��Լ��', 'Y').Getnumber(r_End.��Լ��);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getdate(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '����ID', 'N').Getnumber(r_Old.����id);
    i_Result := l_Lcr_App.Get_Value('old', '��ĿID', 'N').Getnumber(r_Old.��Ŀid);
    i_Result := l_Lcr_App.Get_Value('old', 'ҽ������', 'N').Getvarchar2(r_Old.ҽ������);
    i_Result := l_Lcr_App.Get_Value('old', 'ҽ��ID', 'N').Getnumber(r_Old.ҽ��id);
    i_Result := l_Lcr_App.Get_Value('old', '�ѹ���', 'N').Getnumber(r_Old.�ѹ���);
    i_Result := l_Lcr_App.Get_Value('old', '��Լ��', 'N').Getnumber(r_Old.��Լ��);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End ���˹ҺŻ���_Dml;
/

Create Or Replace Procedure ����δ�����_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End ����δ�����%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old ����δ�����%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In ����δ�����%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update ����δ�����
    Set ��� = ��� + Nvl(r_Dat.���, 0) * n_Sign
    Where ����id = r_Dat.����id And Nvl(��ҳid, 0) = Nvl(r_Dat.��ҳid, 0) And
          Nvl(���˲���id, 0) = Nvl(r_Dat.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Dat.���˿���id, 0) And
          Nvl(��������id, 0) = Nvl(r_Dat.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Dat.ִ�в���id, 0) And
          Nvl(������Ŀid, 0) = Nvl(r_Dat.������Ŀid, 0) And ��Դ;�� = r_Dat.��Դ;��;
    If Sql%Rowcount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (r_Dat.����id, r_Dat.��ҳid, r_Dat.���˲���id, r_Dat.���˿���id, r_Dat.��������id, r_Dat.ִ�в���id,
         r_Dat.������Ŀid, r_Dat.��Դ;��, Nvl(r_Dat.���, 0) * n_Sign);
    End If;
    Delete ����δ�����
    Where ����id = r_Dat.����id And Nvl(��ҳid, 0) = Nvl(r_Dat.��ҳid, 0) And
          Nvl(���˲���id, 0) = Nvl(r_Dat.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Dat.���˿���id, 0) And
          Nvl(��������id, 0) = Nvl(r_Dat.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Dat.ִ�в���id, 0) And
          Nvl(������Ŀid, 0) = Nvl(r_Dat.������Ŀid, 0) And ��Դ;�� = r_Dat.��Դ;�� And Nvl(���, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
   i_Result := l_Lcr_App.Get_Value('new', '����ID', 'Y').Getnumber(r_End.����id);
   i_Result := l_Lcr_App.Get_Value('new', '��ҳID', 'Y').Getnumber(r_End.��ҳid);
   i_Result := l_Lcr_App.Get_Value('new', '���˲���ID', 'Y').Getnumber(r_End.���˲���id);
   i_Result := l_Lcr_App.Get_Value('new', '���˿���ID', 'Y').Getnumber(r_End.���˿���id);
   i_Result := l_Lcr_App.Get_Value('new', '��������ID', 'Y').Getnumber(r_End.��������id);
   i_Result := l_Lcr_App.Get_Value('new', 'ִ�в���ID', 'Y').Getnumber(r_End.ִ�в���id);
   i_Result := l_Lcr_App.Get_Value('new', '������ĿID', 'Y').Getnumber(r_End.������Ŀid);
   i_Result := l_Lcr_App.Get_Value('new', '��Դ;��', 'Y').Getnumber(r_End.��Դ;��);
   i_Result := l_Lcr_App.Get_Value('new', '���', 'Y').Getnumber(r_End.���);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '����ID', 'N').Getnumber(r_Old.����id);
    i_Result := l_Lcr_App.Get_Value('old', '��ҳID', 'N').Getnumber(r_Old.��ҳid);
    i_Result := l_Lcr_App.Get_Value('old', '���˲���ID', 'N').Getnumber(r_Old.���˲���id);
    i_Result := l_Lcr_App.Get_Value('old', '���˿���ID', 'N').Getnumber(r_Old.���˿���id);
    i_Result := l_Lcr_App.Get_Value('old', '��������ID', 'N').Getnumber(r_Old.��������id);
    i_Result := l_Lcr_App.Get_Value('old', 'ִ�в���ID', 'N').Getnumber(r_Old.ִ�в���id);
    i_Result := l_Lcr_App.Get_Value('old', '������ĿID', 'N').Getnumber(r_Old.������Ŀid);
    i_Result := l_Lcr_App.Get_Value('old', '��Դ;��', 'N').Getnumber(r_Old.��Դ;��);
    i_Result := l_Lcr_App.Get_Value('old', '���', 'N').Getnumber(r_Old.���);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End ����δ�����_Dml;
/

Create Or Replace Procedure �������_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End �������%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old �������%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In �������%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Dat.Ԥ�����, 0) * n_Sign,
        ������� = Nvl(�������, 0) + Nvl(r_Dat.�������, 0) * n_Sign
    Where ����id = r_Dat.����id And ���� = r_Dat.����;
    If Sql%Rowcount = 0 Then
      Insert Into �������
        (����id, ����, Ԥ�����, �������)
      Values
        (r_Dat.����id, r_Dat.����, Nvl(r_Dat.Ԥ�����, 0) * n_Sign, Nvl(r_Dat.�������, 0) * n_Sign);
    End If;
    Delete �������
    Where ����id = r_Dat.����id And ���� = r_Dat.���� And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();


  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '����ID', 'Y').Getnumber(r_End.����id);
   i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getnumber(r_End.����);
   i_Result := l_Lcr_App.Get_Value('new', 'Ԥ�����', 'Y').Getnumber(r_End.Ԥ�����);
   i_Result := l_Lcr_App.Get_Value('new', '�������', 'Y').Getnumber(r_End.�������);	
  end if;
  
  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '����ID', 'N').Getnumber(r_Old.����id);
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getnumber(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', 'Ԥ�����', 'N').Getnumber(r_Old.Ԥ�����);
    i_Result := l_Lcr_App.Get_Value('old', '�������', 'N').Getnumber(r_Old.�������);
  END IF;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then     
    Apply_To(r_Old, -1);
  End If;
End �������_Dml;
/

Create Or Replace Procedure ��Ա�ɿ����_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End ��Ա�ɿ����%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old ��Ա�ɿ����%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In ��Ա�ɿ����%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(r_Dat.���, 0) * n_Sign
    Where �տ�Ա = r_Dat.�տ�Ա And ���㷽ʽ = r_Dat.���㷽ʽ And ���� = r_Dat.����;
    If Sql%Rowcount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (r_Dat.�տ�Ա, r_Dat.���㷽ʽ, r_Dat.����, Nvl(r_Dat.���, 0) * n_Sign);
    End If;
    Delete ��Ա�ɿ����
    Where �տ�Ա = r_Dat.�տ�Ա And ���㷽ʽ = r_Dat.���㷽ʽ And ���� = r_Dat.���� And Nvl(���, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '�տ�Ա', 'Y').Getvarchar2(r_End.�տ�Ա);
   i_Result := l_Lcr_App.Get_Value('new', '���㷽ʽ', 'Y').Getvarchar2(r_End.���㷽ʽ);
   i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getnumber(r_End.����);
   i_Result := l_Lcr_App.Get_Value('new', '���', 'Y').Getnumber(r_End.���);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '�տ�Ա', 'N').Getvarchar2(r_Old.�տ�Ա);
    i_Result := l_Lcr_App.Get_Value('old', '���㷽ʽ', 'N').Getvarchar2(r_Old.���㷽ʽ);
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getnumber(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '���', 'N').Getnumber(r_Old.���);
  END IF;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End ��Ա�ɿ����_Dml;
/

--2007/07/05 ���˺�:����ҩƷ���\ҩƷ�շ�����\Ӧ������DML����

--��ص�DML����
Create Or Replace Procedure ҩƷ���_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End ҩƷ���%RowType; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old ҩƷ���%RowType; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In ҩƷ���%RowType,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + Nvl(r_Dat.��������, 0) * n_Sign, ʵ������ = Nvl(ʵ������, 0) + Nvl(r_Dat.ʵ������, 0) * n_Sign,
        ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(r_Dat.ʵ�ʽ��, 0) * n_Sign, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(r_Dat.ʵ�ʲ��, 0) * n_Sign,
        �ϴβɹ��� = Nvl(r_Dat.�ϴβɹ���, �ϴβɹ���), �ϴβ��� = Nvl(r_Dat.�ϴβ���, �ϴβ���), �ϴ����� = Nvl(r_Dat.�ϴ�����, �ϴ�����),
        �ϴ��������� = Nvl(r_Dat.�ϴ���������, �ϴ���������), ���ۼ� = Nvl(r_Dat.���ۼ�, ���ۼ�), �ϴο��� = Nvl(r_Dat.�ϴο���, �ϴο���)
    Where �ⷿid = r_Dat.�ⷿid And ҩƷid = r_Dat.ҩƷid And Nvl(����, 0) = Nvl(r_Dat.����, 0) And ���� = r_Dat.����;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, Ч��, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, ��׼�ĺ�, ���ۼ�, �ϴο���)
      Values
        (r_Dat.�ⷿid, r_Dat.ҩƷid, r_Dat.����, r_Dat.Ч��, r_Dat.����, Nvl(r_Dat.��������, 0) * n_Sign, Nvl(r_Dat.ʵ������, 0) * n_Sign,
         Nvl(r_Dat.ʵ�ʽ��, 0) * n_Sign, Nvl(r_Dat.ʵ�ʲ��, 0) * n_Sign, r_Dat.�ϴι�Ӧ��id, r_Dat.�ϴβɹ���, r_Dat.�ϴ�����, r_Dat.�ϴ���������,
         r_Dat.�ϴβ���, r_Dat.���Ч��, r_Dat.��׼�ĺ�, r_Dat.���ۼ�, r_Dat.�ϴο���);
    End If;
    Delete ҩƷ���
    Where �ⷿid = r_Dat.�ⷿid And ҩƷid = r_Dat.ҩƷid And Nvl(����, 0) = Nvl(r_Dat.����, 0) And ���� = r_Dat.���� And
          Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  If v_Cmd_Type = 'INSERT' Or v_Cmd_Type = 'UPDATE' Then
    i_Result := l_Lcr_App.Get_Value('new', '�ⷿID', 'Y').Getnumber(r_End.�ⷿid);
    i_Result := l_Lcr_App.Get_Value('new', 'ҩƷID', 'Y').Getnumber(r_End.ҩƷid);
    i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getnumber(r_End.����);
    i_Result := l_Lcr_App.Get_Value('new', 'Ч��', 'Y').Getdate(r_End.Ч��);
    i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getnumber(r_End.����);
    i_Result := l_Lcr_App.Get_Value('new', '��������', 'Y').Getnumber(r_End.��������);
    i_Result := l_Lcr_App.Get_Value('new', 'ʵ������', 'Y').Getnumber(r_End.ʵ������);
    i_Result := l_Lcr_App.Get_Value('new', 'ʵ�ʽ��', 'Y').Getnumber(r_End.ʵ�ʽ��);
    i_Result := l_Lcr_App.Get_Value('new', 'ʵ�ʲ��', 'Y').Getnumber(r_End.ʵ�ʲ��);
    i_Result := l_Lcr_App.Get_Value('new', '�ϴι�Ӧ��ID', 'Y').Getnumber(r_End.�ϴι�Ӧ��id);
    i_Result := l_Lcr_App.Get_Value('new', '�ϴβɹ���', 'Y').Getnumber(r_End.�ϴβɹ���);
    i_Result := l_Lcr_App.Get_Value('new', '�ϴ�����', 'Y').Getvarchar2(r_End.�ϴ�����);
    i_Result := l_Lcr_App.Get_Value('new', '�ϴ���������', 'Y').Getdate(r_End.�ϴ���������);
    i_Result := l_Lcr_App.Get_Value('new', '�ϴβ���', 'Y').Getvarchar2(r_End.�ϴβ���);
    i_Result := l_Lcr_App.Get_Value('new', '���Ч��', 'Y').Getdate(r_End.���Ч��);
    i_Result := l_Lcr_App.Get_Value('new', '��׼�ĺ�', 'Y').Getvarchar2(r_End.��׼�ĺ�);
    i_Result := l_Lcr_App.Get_Value('new', '���ۼ�', 'Y').Getnumber(r_End.���ۼ�);
    i_Result := l_Lcr_App.Get_Value('new', '�ϴο���', 'Y').Getnumber(r_End.�ϴο���);
  End If;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '�ⷿID', 'N').Getnumber(r_Old.�ⷿid);
    i_Result := l_Lcr_App.Get_Value('old', 'ҩƷID', 'N').Getnumber(r_Old.ҩƷid);
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getnumber(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', 'Ч��', 'N').Getdate(r_Old.Ч��);
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getnumber(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '��������', 'N').Getnumber(r_Old.��������);
    i_Result := l_Lcr_App.Get_Value('old', 'ʵ������', 'N').Getnumber(r_Old.ʵ������);
    i_Result := l_Lcr_App.Get_Value('old', 'ʵ�ʽ��', 'N').Getnumber(r_Old.ʵ�ʽ��);
    i_Result := l_Lcr_App.Get_Value('old', 'ʵ�ʲ��', 'N').Getnumber(r_Old.ʵ�ʲ��);
    i_Result := l_Lcr_App.Get_Value('old', '�ϴι�Ӧ��ID', 'N').Getnumber(r_Old.�ϴι�Ӧ��id);
    i_Result := l_Lcr_App.Get_Value('old', '�ϴβɹ���', 'N').Getnumber(r_Old.�ϴβɹ���);
    i_Result := l_Lcr_App.Get_Value('old', '�ϴ�����', 'N').Getvarchar2(r_Old.�ϴ�����);
    i_Result := l_Lcr_App.Get_Value('old', '�ϴ���������', 'N').Getdate(r_Old.�ϴ���������);
    i_Result := l_Lcr_App.Get_Value('old', '�ϴβ���', 'N').Getvarchar2(r_Old.�ϴβ���);
    i_Result := l_Lcr_App.Get_Value('old', '���Ч��', 'N').Getdate(r_Old.���Ч��);
    i_Result := l_Lcr_App.Get_Value('old', '��׼�ĺ�', 'N').Getvarchar2(r_Old.��׼�ĺ�);
    i_Result := l_Lcr_App.Get_Value('old', '���ۼ�', 'N').Getnumber(r_Old.���ۼ�);
    i_Result := l_Lcr_App.Get_Value('old', '�ϴο���', 'N').Getnumber(r_Old.�ϴο���);
  End If;
  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End ҩƷ���_Dml;
/


Create Or Replace Procedure ҩƷ�շ�����_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End ҩƷ�շ�����%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old ҩƷ�շ�����%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In ҩƷ�շ�����%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update ҩƷ�շ�����
    Set ���� = Nvl(����, 0) + Nvl(r_Dat.����, 0) * n_Sign, ��� = Nvl(���, 0) + Nvl(r_Dat.���, 0) * n_Sign,
        ��� = Nvl(���, 0) + Nvl(r_Dat.���, 0) * n_Sign
    Where ���� = r_Dat.���� And Nvl(�ⷿid, 0) = Nvl(r_Dat.�ⷿid, 0) And Nvl(ҩƷid, 0) = Nvl(r_Dat.ҩƷid, 0) And
          Nvl(���id, 0) = Nvl(r_Dat.���id, 0) And ���� = r_Dat.����;
    If Sql%Rowcount = 0 Then
      Insert Into ҩƷ�շ�����
        (����, �ⷿid, ҩƷid, ���id, ����, ����, ���, ���)
      Values
        (r_Dat.����, r_Dat.�ⷿid, r_Dat.ҩƷid, r_Dat.���id, r_Dat.����, Nvl(r_Dat.����, 0) * n_Sign,
         Nvl(r_Dat.���, 0) * n_Sign, Nvl(r_Dat.���, 0) * n_Sign);
    End If;
    Delete ҩƷ�շ�����
    Where ���� = r_Dat.���� And Nvl(�ⷿid, 0) = Nvl(r_Dat.�ⷿid, 0) And Nvl(ҩƷid, 0) = Nvl(r_Dat.ҩƷid, 0) And
          Nvl(���id, 0) = Nvl(r_Dat.���id, 0) And ���� = r_Dat.���� And Nvl(����, 0) = 0 And Nvl(���, 0) = 0 And
          Nvl(���, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  if v_Cmd_Type = 'INSERT' or v_Cmd_Type = 'UPDATE' then
   i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getdate(r_End.����);
   i_Result := l_Lcr_App.Get_Value('new', '�ⷿID', 'Y').Getnumber(r_End.�ⷿid);
   i_Result := l_Lcr_App.Get_Value('new', 'ҩƷID', 'Y').Getnumber(r_End.ҩƷid);
   i_Result := l_Lcr_App.Get_Value('new', '���ID', 'Y').Getnumber(r_End.���id);
   i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getnumber(r_End.����);
   i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getnumber(r_End.����);
   i_Result := l_Lcr_App.Get_Value('new', '���', 'Y').Getnumber(r_End.���);
   i_Result := l_Lcr_App.Get_Value('new', '���', 'Y').Getnumber(r_End.���);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getdate(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '�ⷿID', 'N').Getnumber(r_Old.�ⷿid);
    i_Result := l_Lcr_App.Get_Value('old', 'ҩƷID', 'N').Getnumber(r_Old.ҩƷid);
    i_Result := l_Lcr_App.Get_Value('old', '���ID', 'N').Getnumber(r_Old.���id);
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getnumber(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getnumber(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '���', 'N').Getnumber(r_Old.���);
    i_Result := l_Lcr_App.Get_Value('old', '���', 'N').Getnumber(r_Old.���);
  end if;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End ҩƷ�շ�����_Dml;
/


Create Or Replace Procedure Ӧ�����_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End Ӧ�����%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old Ӧ�����%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In Ӧ�����%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update Ӧ�����
    Set ��� = Nvl(���, 0) + Nvl(r_Dat.���, 0) * n_Sign
    Where ��λid = r_Dat.��λid And ���� = r_Dat.����;
    If Sql%Rowcount = 0 Then
      Insert Into Ӧ����� (��λid, ����, ���) Values (r_Dat.��λid, r_Dat.����, Nvl(r_Dat.���, 0) * n_Sign);
    End If;
    Delete Ӧ����� Where ��λid = r_Dat.��λid And ���� = r_Dat.���� And Nvl(���, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '��λID', 'Y').Getnumber(r_End.��λid);
   i_Result := l_Lcr_App.Get_Value('new', '����', 'Y').Getnumber(r_End.����);
   i_Result := l_Lcr_App.Get_Value('new', '���', 'Y').Getnumber(r_End.���);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then 
    i_Result := l_Lcr_App.Get_Value('old', '��λID', 'Y').Getnumber(r_Old.��λid);
    i_Result := l_Lcr_App.Get_Value('old', '����', 'N').Getnumber(r_Old.����);
    i_Result := l_Lcr_App.Get_Value('old', '���', 'N').Getnumber(r_Old.���);
  END IF;  

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End Ӧ�����_Dml;
/

------------------------------------------------------------------------------------------------------------------------------------------

--��ҵ�죺ҩƷ����DML
Create Or Replace Procedure ҩƷ����_Dml(Any_In In Anydata) Is
  l_Lcr_App  Sys.Lcr$_Row_Record;
  i_Result   Pls_Integer;
  v_Cmd_Type Varchar2(30);

  r_End ҩƷ����%Rowtype; --��¼�������̬��������¼״̬���޸ĺ����״̬��ɾ��ǰ��״̬
  r_Old ҩƷ����%Rowtype; --��¼�������̬����¼��ԭ״̬���������޸Ĳ���

  ---------------------
  --Ӧ�ô����ӹ���
  Procedure Apply_To
  (
    r_Dat  In ҩƷ����%Rowtype,
    n_Sign In Number := 1 --����1,���ӣ�-1,����
  ) Is
  Begin
    Update ҩƷ����
    Set �������� = Nvl(��������, 0) + Nvl(r_Dat.��������, 0) * n_Sign,
        ʵ������ = Nvl(ʵ������, 0) + Nvl(r_Dat.ʵ������, 0) * n_Sign,
        ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(r_Dat.ʵ�ʽ��, 0) * n_Sign
    Where �ڼ� = r_Dat.�ڼ� And ����id = r_Dat.����id And �ⷿid = r_Dat.�ⷿid And ҩƷid = r_Dat.ҩƷid;
    If Sql%Rowcount = 0 Then
      Insert Into ҩƷ����
        (�ڼ�, ����id, �ⷿid, ҩƷid, ��������, ʵ������, ʵ�ʽ��)
      Values
        (r_Dat.�ڼ�, r_Dat.����id, r_Dat.�ⷿid, r_Dat.ҩƷid, Nvl(r_Dat.��������, 0) * n_Sign,
         Nvl(r_Dat.ʵ������, 0) * n_Sign, Nvl(r_Dat.ʵ�ʽ��, 0) * n_Sign);
    End If;
    Delete ҩƷ����
    Where �ڼ� = r_Dat.�ڼ� And ����id = r_Dat.����id And �ⷿid = r_Dat.�ⷿid And ҩƷid = r_Dat.ҩƷid And
          Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0;
  End Apply_To;

  ---------------------
  --������
  ---------------------
Begin
  i_Result   := Any_In.Getobject(l_Lcr_App);
  v_Cmd_Type := l_Lcr_App.Get_Command_Type();

  IF v_Cmd_Type = 'INSERT' OR v_Cmd_Type = 'UPDATE' THEN
   i_Result := l_Lcr_App.Get_Value('new', '�ڼ�', 'Y').Getvarchar2(r_End.�ڼ�);
   i_Result := l_Lcr_App.Get_Value('new', '����ID', 'Y').Getnumber(r_End.����id);
   i_Result := l_Lcr_App.Get_Value('new', '�ⷿID', 'Y').Getnumber(r_End.�ⷿid);
   i_Result := l_Lcr_App.Get_Value('new', 'ҩƷID', 'Y').Getnumber(r_End.ҩƷid);
   i_Result := l_Lcr_App.Get_Value('new', '��������', 'Y').Getnumber(r_End.��������);
   i_Result := l_Lcr_App.Get_Value('new', 'ʵ������', 'Y').Getnumber(r_End.ʵ������);
   i_Result := l_Lcr_App.Get_Value('new', 'ʵ�ʽ��', 'Y').Getnumber(r_End.ʵ�ʽ��);
  end if;

  If v_Cmd_Type = 'UPDATE' Or v_Cmd_Type = 'DELETE' Then
    i_Result := l_Lcr_App.Get_Value('old', '�ڼ�', 'Y').Getvarchar2(r_Old.�ڼ�);
    i_Result := l_Lcr_App.Get_Value('old', '����ID', 'Y').Getnumber(r_Old.����id);
    i_Result := l_Lcr_App.Get_Value('old', '�ⷿID', 'N').Getnumber(r_Old.�ⷿid);
    i_Result := l_Lcr_App.Get_Value('old', 'ҩƷID', 'N').Getnumber(r_Old.ҩƷid);
    i_Result := l_Lcr_App.Get_Value('old', '��������', 'N').Getnumber(r_Old.��������);
    i_Result := l_Lcr_App.Get_Value('old', 'ʵ������', 'N').Getnumber(r_Old.ʵ������);
    i_Result := l_Lcr_App.Get_Value('old', 'ʵ�ʽ��', 'N').Getnumber(r_Old.ʵ�ʽ��);
  END IF;

  If v_Cmd_Type = 'INSERT' Then
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'UPDATE' Then
    Apply_To(r_Old, -1);
    Apply_To(r_End, 1);
  Elsif v_Cmd_Type = 'DELETE' Then
    Apply_To(r_Old, -1);
  End If;
End ҩƷ����_Dml;
/

