--������Ժ��ҩ���ӿڣ��ڡ�Zl_ҩƷ�շ���¼_������ҩ���е���
--v0307

Create Or Replace Procedure Zl_Getpackerdetail(�շ�_Id In Number) Is
  v_��ˮ��       Varchar2(20);
  v_���         Number(2);
  v_С���־     Number(2);
  v_������       Number(2);
  v_����״̬     Number(2);
  v_�����־     Number(1);
  v_�������     Number(1);
  v_ִ������     Number(10);
  v_No           Varchar2(10);
  v_Tmp          Varchar2(100);
  v_Fields       Varchar2(100);
  v_Time         Varchar2(100);
  v_ʱ�䷽������ Varchar2(100);
  v_ʱ�䷽��     Varchar2(100);
  v_Count        Number(2) := 1;
  Cursor c_�շ���¼ Is
    Select /*+rule */
     c.����ʱ��, b.No, b.���, d.���id, a.��ʶ��, a.����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 9) As �Ա�, f.���� As ��������, f.���� As ��������, a.����,
     b.�������, Trunc(c.�״�ʱ��) As �״�����, To_Char(c.�״�ʱ��, 'hh24') || '00' As �״�ִ��ʱ��, To_Char(c.ĩ��ʱ��, 'hh24') || '00' As ĩ��ִ��ʱ��,
     g.���� As ҩƷ����, g.���� As ҩƷ����, h.���㵥λ, d.��������, i.����ϵ��, i.סԺ��λ, b.�÷�, d.ִ��ʱ�䷽��, d.Ƶ�ʼ��, d.�����λ, d.��ʼִ��ʱ��, c.�״�ʱ��,
     c.ĩ��ʱ��, d.ҽ������
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ������ C, ����ҽ����¼ D, ���ű� F, �շ���ĿĿ¼ G, ������ĿĿ¼ H, ҩƷ��� I, ҩƷ���� L
    Where a.Id = b.����id And a.ҽ����� = c.ҽ��id And c.ҽ��id = d.Id And b.No = c.No And b.���� = 9 And b.Id = �շ�_Id And
          a.���˲���id = f.Id And b.ҩƷid = g.Id And h.Id = i.ҩ��id And b.ҩƷid = i.ҩƷid And i.ҩ��id = l.ҩ��id And d.ҽ����Ч = 0 And
          b.������� Is Not Null And c.�״�ʱ�� Is Not Null And c.ĩ��ʱ�� Is Not Null And b.�÷� = '�ڷ�' And Mod(b.��¼״̬, 3) = 1 And
          d.�����λ = '��' And a.���˲���id <> 1448 And l.ҩƷ���� In ('���ͽ��Ҽ�', 'Ƭ��', '���Ҽ�', '����Ƭ��', '��ɢƬ', '����Ƭ��')
    Order By c.No, b.���;

  v_�շ���¼ c_�շ���¼%RowType;

  --����ҽ��ִ������
  Function Zl_Getexedays
  (
    ��ʼִ��ʱ��_In In Date,
    �״�ִ��ʱ��_In In Date,
    ĩ��ִ��ʱ��_In In Date,
    Ƶ�ʼ��_In     In Number,
    �����λ_In     In Varchar2,
    ִ��ʱ�䷽��_In In Varchar2
  ) Return Number As
    v_Exe         Varchar2(200);
    v_Exetime     Date;
    v_Lastexetime Date;
    v_Tmptime     Varchar2(50);
    v_Split       Number;
    v_Plan        Varchar2(20);
    v_Stop        Boolean;
    v_Cycle       Number;
    v_Adddate     Number;
    v_Addhour     Number;
    v_Addminute   Number;
    v_Execount    Number;
  Begin
    v_Exetime     := ��ʼִ��ʱ��_In;
    v_Lastexetime := �״�ִ��ʱ��_In;
    v_Split       := Ƶ�ʼ��_In;
    v_Stop        := False;
    v_Cycle       := 0;
    v_Execount    := 0;
    While Not v_Stop Loop
      v_Exe := ִ��ʱ�䷽��_In || '-';
    
      --��ִ��ʱ�䷽��ѭ�� 
      While v_Exe Is Not Null Loop
        v_Plan := Substr(v_Exe, 1, Instr(v_Exe, '-') - 1);
        v_Exe  := Replace('-' || v_Exe, '-' || v_Plan || '-');
      
        If �����λ_In = '��' Then
          --�����λ��"��"ʱ��ʼ�հ���ʼִ��ʱ�������㵱ǰִ��ʱ�� 
          v_Exetime := ��ʼִ��ʱ��_In;
        
          If Ƶ�ʼ��_In = 1 Then
            --ʱ������1ʱ����ʾ����ִ�� 
            v_Adddate := 0;
          Elsif Ƶ�ʼ��_In = 7 Then
            --���Ϊ�� 
          
            v_Adddate := To_Number(Substr(v_Plan, 1, Instr(v_Plan, '/') - 1)) - 1;
          
            --ȡִ��ʱ�� 
            v_Plan := Substr(v_Plan, Instr(v_Plan, '/') + 1);
          Else
            --ȡҪ���ӵ����� 
            v_Adddate := To_Number(Substr(v_Plan, 1, Instr(v_Plan, '/') - 1)) - 1;
          
            --ȡִ��ʱ�� 
            v_Plan := Substr(v_Plan, Instr(v_Plan, '/') + 1);
          End If;
        Elsif �����λ_In = 'Сʱ' Then
          v_Adddate := 0;
        
          If Instr(v_Plan, ':') = 0 Then
            v_Addhour   := To_Number(v_Plan) - 1;
            v_Addminute := 0;
          Else
            v_Addhour   := To_Number(Substr(v_Plan, 1, Instr(v_Plan, ':') - 1)) - 1;
            v_Addminute := To_Number(Substr(v_Plan, Instr(v_Plan, ':') + 1));
          End If;
        End If;
      
        If �����λ_In = '��' Then
          If Ƶ�ʼ��_In = 7 Then
            --ȡ��ʼִ��ʱ�������ܵ���һ���� 
            v_Exetime := Trunc(v_Exetime, 'iw');
          End If;
        
          --���㵱ǰִ��ʱ�� 
          v_Tmptime := To_Char(v_Exetime, 'YYYY-MM-DD') ||
                       Substr(To_Char(To_Date(v_Plan, 'hh24:mi:ss'), 'YYYY-MM-DD HH24:MI:SS'), 11);
          v_Exetime := To_Date(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS') + v_Split * v_Cycle + v_Adddate;
        Elsif �����λ_In = 'Сʱ' Then
          --�����λ��"ʱ"ʱ����ǰִ��ʱ���ǰ��ӿ�ʼִ��ʱ���𣬰����Сʱ����ִ�з����ۼ� 
          v_Exetime := ��ʼִ��ʱ��_In + v_Split * v_Cycle / 24 + v_Addhour / 24 + v_Addminute / 24 / 60;
        End If;
      
        --��ǰִ��ʱ�����ĩ��ִ��ʱ���˳�ѭ�� 
        If v_Exetime > ĩ��ִ��ʱ��_In Then
          v_Stop := True;
          If Trunc(v_Exetime) > Trunc(v_Lastexetime) Then
            v_Execount := v_Execount + 1;
          End If;
          Exit;
        Elsif v_Exetime >= �״�ִ��ʱ��_In Then
          If Trunc(v_Exetime) > Trunc(v_Lastexetime) Then
            v_Execount    := v_Execount + 1;
            v_Lastexetime := v_Exetime;
          End If;
        End If;
      End Loop;
    
      --ѭ������ 
      v_Cycle := v_Cycle + 1;
    End Loop;
  
    --����ִ������
    If v_Execount = 0 Then
      v_Execount := 1;
    End If;
    Return(v_Execount);
  End;
Begin
  v_С���־ := 1;
  v_������   := 1;
  v_����״̬ := 0;
  v_�����־ := 2;
  v_������� := 0;
  v_���     := 0;

  For v_�շ���¼ In c_�շ���¼ Loop
    If v_No = v_�շ���¼.No Then
      v_��� := v_��� + 1;
    Else
      v_���   := v_�շ���¼.���;
      v_No     := v_�շ���¼.No;
      v_��ˮ�� := To_Char(v_�շ���¼.����ʱ��, 'yymmdd') || v_No;
    End If;
  
    v_Tmp := v_�շ���¼.ִ��ʱ�䷽�� || '-';
  
    While v_Tmp Is Not Null Loop
      v_Fields := Substr(v_Tmp, 1, Instr(v_Tmp, '-') - 1);
    
      v_Time := v_Fields;
      If Instr(v_Fields, ',') > 0 Then
        v_Time := Substr(v_Fields, Instr(v_Fields, ',') + 1);
      End If;
      If Instr(v_Fields, '/') > 0 Then
        v_Time := Substr(v_Fields, Instr(v_Fields, '/') + 1);
      End If;
    
      If v_ʱ�䷽�� Is Null Then
        v_ʱ�䷽��     := To_Char(To_Date(v_Time, 'hh24'), 'hh24');
        v_ʱ�䷽������ := To_Char(To_Date(v_Time, 'hh24'), 'hh24') || '00';
      Else
        v_ʱ�䷽��     := v_ʱ�䷽�� || '-' || To_Char(To_Date(v_Time, 'hh24'), 'hh24');
        v_ʱ�䷽������ := v_ʱ�䷽������ || '-' || To_Char(To_Date(v_Time, 'hh24'), 'hh24') || '00';
      End If;
    
      v_Tmp := Replace('-' || v_Tmp, '-' || v_Fields || '-');
    
      /*      --�̶�ʱ��
      If v_Count = 1 Then
        If Zl_To_Number(v_Fields) < 12 Then
          v_ʱ�䷽��              := '��';
          v_ʱ�䷽������          := '0800';
          v_�շ���¼.�״�ִ��ʱ�� := '0800';
          v_�շ���¼.ĩ��ִ��ʱ�� := '0800';
        Else
          v_ʱ�䷽��              := '��';
          v_ʱ�䷽������          := '1800';
          v_�շ���¼.�״�ִ��ʱ�� := '1800';
          v_�շ���¼.ĩ��ִ��ʱ�� := '1800';
        
        End If;
      Elsif v_Count = 2 Then
        v_ʱ�䷽��              := '��-��';
        v_ʱ�䷽������          := '0800-1800';
        v_�շ���¼.�״�ִ��ʱ�� := '0800';
        v_�շ���¼.ĩ��ִ��ʱ�� := '1800';
      Elsif v_Count = 3 Then
        v_ʱ�䷽��              := '��-��-��';
        v_ʱ�䷽������          := '0800-1300-1800';
        v_�շ���¼.�״�ִ��ʱ�� := '0800';
        v_�շ���¼.ĩ��ִ��ʱ�� := '1800';
      End If;*/
    
      v_Count := v_Count + 1;
    End Loop;
  
    v_ִ������ := Zl_Getexedays(v_�շ���¼.��ʼִ��ʱ��, v_�շ���¼.�״�ʱ��, v_�շ���¼.ĩ��ʱ��, v_�շ���¼.Ƶ�ʼ��, v_�շ���¼.�����λ, v_�շ���¼.ִ��ʱ�䷽��);
    Insert Into Prescriptiondate
      (Prescriptionno, Seqno, Group_No, Machineno, Procflg, Patientid, Patientname, Sex, Ioflg, Wardcd, Wardname, Bedno,
       Prescriptiondate, Takedate, Taketime, Lasttime, Presc_Class, Drugcd, Drugname, Dispenseddose, Prescriptiondose,
       Prescriptionunit, Dispensedunit, Amount_Per_Package, Dispense_Days, Freq_Desc, Freq_Desc_Detail, Makerectime,
       Freq_Desc_Detail_Code, Explanation)
    Values
      (v_��ˮ��, v_���, v_С���־, v_������, v_����״̬, v_�շ���¼.��ʶ��, v_�շ���¼.����, v_�շ���¼.�Ա�, v_�����־, v_�շ���¼.��������, v_�շ���¼.��������,
       v_�շ���¼.����, v_�շ���¼.�������, v_�շ���¼.�״�����, v_�շ���¼.�״�ִ��ʱ��, v_�շ���¼.ĩ��ִ��ʱ��, v_�������, v_�շ���¼.ҩƷ����,
       v_�շ���¼.ҩƷ���� || '/' || v_�շ���¼.�������� || v_�շ���¼.���㵥λ, v_�շ���¼.�������� / v_�շ���¼.����ϵ��, v_�շ���¼.��������, v_�շ���¼.���㵥λ,
       v_�շ���¼.סԺ��λ, v_�շ���¼.����ϵ��, v_ִ������, v_�շ���¼.�÷�, v_ʱ�䷽��, v_�շ���¼.�������, v_ʱ�䷽������, v_�շ���¼.ҽ������);
  End Loop;
  Commit;
End Zl_Getpackerdetail;
/
