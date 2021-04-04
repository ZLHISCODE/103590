--大连三院包药机接口，在“Zl_药品收发记录_批量发药”中调用
--v0307

Create Or Replace Procedure Zl_Getpackerdetail(收发_Id In Number) Is
  v_流水号       Varchar2(20);
  v_序号         Number(2);
  v_小组标志     Number(2);
  v_机器号       Number(2);
  v_处理状态     Number(2);
  v_门诊标志     Number(1);
  v_紧急类别     Number(1);
  v_执行天数     Number(10);
  v_No           Varchar2(10);
  v_Tmp          Varchar2(100);
  v_Fields       Varchar2(100);
  v_Time         Varchar2(100);
  v_时间方案编码 Varchar2(100);
  v_时间方案     Varchar2(100);
  v_Count        Number(2) := 1;
  Cursor c_收发记录 Is
    Select /*+rule */
     c.发送时间, b.No, b.序号, d.相关id, a.标识号, a.姓名, Decode(a.性别, '男', 1, '女', 2, 9) As 性别, f.编码 As 病区编码, f.名称 As 病区名称, a.床号,
     b.审核日期, Trunc(c.首次时间) As 首次日期, To_Char(c.首次时间, 'hh24') || '00' As 首次执行时间, To_Char(c.末次时间, 'hh24') || '00' As 末次执行时间,
     g.编码 As 药品编码, g.名称 As 药品名称, h.计算单位, d.单次用量, i.剂量系数, i.住院单位, b.用法, d.执行时间方案, d.频率间隔, d.间隔单位, d.开始执行时间, c.首次时间,
     c.末次时间, d.医生嘱托
    From 住院费用记录 A, 药品收发记录 B, 病人医嘱发送 C, 病人医嘱记录 D, 部门表 F, 收费项目目录 G, 诊疗项目目录 H, 药品规格 I, 药品特性 L
    Where a.Id = b.费用id And a.医嘱序号 = c.医嘱id And c.医嘱id = d.Id And b.No = c.No And b.单据 = 9 And b.Id = 收发_Id And
          a.病人病区id = f.Id And b.药品id = g.Id And h.Id = i.药名id And b.药品id = i.药品id And i.药名id = l.药名id And d.医嘱期效 = 0 And
          b.审核日期 Is Not Null And c.首次时间 Is Not Null And c.末次时间 Is Not Null And b.用法 = '口服' And Mod(b.记录状态, 3) = 1 And
          d.间隔单位 = '天' And a.病人病区id <> 1448 And l.药品剂型 In ('缓释胶囊剂', '片剂', '胶囊剂', '控释片剂', '分散片', '缓释片剂')
    Order By c.No, b.序号;

  v_收发记录 c_收发记录%RowType;

  --计算医嘱执行天数
  Function Zl_Getexedays
  (
    开始执行时间_In In Date,
    首次执行时间_In In Date,
    末次执行时间_In In Date,
    频率间隔_In     In Number,
    间隔单位_In     In Varchar2,
    执行时间方案_In In Varchar2
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
    v_Exetime     := 开始执行时间_In;
    v_Lastexetime := 首次执行时间_In;
    v_Split       := 频率间隔_In;
    v_Stop        := False;
    v_Cycle       := 0;
    v_Execount    := 0;
    While Not v_Stop Loop
      v_Exe := 执行时间方案_In || '-';
    
      --按执行时间方案循环 
      While v_Exe Is Not Null Loop
        v_Plan := Substr(v_Exe, 1, Instr(v_Exe, '-') - 1);
        v_Exe  := Replace('-' || v_Exe, '-' || v_Plan || '-');
      
        If 间隔单位_In = '天' Then
          --间隔单位是"天"时，始终按开始执行时间来推算当前执行时间 
          v_Exetime := 开始执行时间_In;
        
          If 频率间隔_In = 1 Then
            --时间间隔是1时，表示当天执行 
            v_Adddate := 0;
          Elsif 频率间隔_In = 7 Then
            --间隔为周 
          
            v_Adddate := To_Number(Substr(v_Plan, 1, Instr(v_Plan, '/') - 1)) - 1;
          
            --取执行时点 
            v_Plan := Substr(v_Plan, Instr(v_Plan, '/') + 1);
          Else
            --取要增加的天数 
            v_Adddate := To_Number(Substr(v_Plan, 1, Instr(v_Plan, '/') - 1)) - 1;
          
            --取执行时点 
            v_Plan := Substr(v_Plan, Instr(v_Plan, '/') + 1);
          End If;
        Elsif 间隔单位_In = '小时' Then
          v_Adddate := 0;
        
          If Instr(v_Plan, ':') = 0 Then
            v_Addhour   := To_Number(v_Plan) - 1;
            v_Addminute := 0;
          Else
            v_Addhour   := To_Number(Substr(v_Plan, 1, Instr(v_Plan, ':') - 1)) - 1;
            v_Addminute := To_Number(Substr(v_Plan, Instr(v_Plan, ':') + 1));
          End If;
        End If;
      
        If 间隔单位_In = '天' Then
          If 频率间隔_In = 7 Then
            --取开始执行时间所在周的周一日期 
            v_Exetime := Trunc(v_Exetime, 'iw');
          End If;
        
          --计算当前执行时间 
          v_Tmptime := To_Char(v_Exetime, 'YYYY-MM-DD') ||
                       Substr(To_Char(To_Date(v_Plan, 'hh24:mi:ss'), 'YYYY-MM-DD HH24:MI:SS'), 11);
          v_Exetime := To_Date(v_Tmptime, 'YYYY-MM-DD HH24:MI:SS') + v_Split * v_Cycle + v_Adddate;
        Elsif 间隔单位_In = '小时' Then
          --间隔单位是"时"时，当前执行时间是按从开始执行时间起，按间隔小时数及执行方案累加 
          v_Exetime := 开始执行时间_In + v_Split * v_Cycle / 24 + v_Addhour / 24 + v_Addminute / 24 / 60;
        End If;
      
        --当前执行时间大于末次执行时，退出循环 
        If v_Exetime > 末次执行时间_In Then
          v_Stop := True;
          If Trunc(v_Exetime) > Trunc(v_Lastexetime) Then
            v_Execount := v_Execount + 1;
          End If;
          Exit;
        Elsif v_Exetime >= 首次执行时间_In Then
          If Trunc(v_Exetime) > Trunc(v_Lastexetime) Then
            v_Execount    := v_Execount + 1;
            v_Lastexetime := v_Exetime;
          End If;
        End If;
      End Loop;
    
      --循环次数 
      v_Cycle := v_Cycle + 1;
    End Loop;
  
    --返回执行天数
    If v_Execount = 0 Then
      v_Execount := 1;
    End If;
    Return(v_Execount);
  End;
Begin
  v_小组标志 := 1;
  v_机器号   := 1;
  v_处理状态 := 0;
  v_门诊标志 := 2;
  v_紧急类别 := 0;
  v_序号     := 0;

  For v_收发记录 In c_收发记录 Loop
    If v_No = v_收发记录.No Then
      v_序号 := v_序号 + 1;
    Else
      v_序号   := v_收发记录.序号;
      v_No     := v_收发记录.No;
      v_流水号 := To_Char(v_收发记录.发送时间, 'yymmdd') || v_No;
    End If;
  
    v_Tmp := v_收发记录.执行时间方案 || '-';
  
    While v_Tmp Is Not Null Loop
      v_Fields := Substr(v_Tmp, 1, Instr(v_Tmp, '-') - 1);
    
      v_Time := v_Fields;
      If Instr(v_Fields, ',') > 0 Then
        v_Time := Substr(v_Fields, Instr(v_Fields, ',') + 1);
      End If;
      If Instr(v_Fields, '/') > 0 Then
        v_Time := Substr(v_Fields, Instr(v_Fields, '/') + 1);
      End If;
    
      If v_时间方案 Is Null Then
        v_时间方案     := To_Char(To_Date(v_Time, 'hh24'), 'hh24');
        v_时间方案编码 := To_Char(To_Date(v_Time, 'hh24'), 'hh24') || '00';
      Else
        v_时间方案     := v_时间方案 || '-' || To_Char(To_Date(v_Time, 'hh24'), 'hh24');
        v_时间方案编码 := v_时间方案编码 || '-' || To_Char(To_Date(v_Time, 'hh24'), 'hh24') || '00';
      End If;
    
      v_Tmp := Replace('-' || v_Tmp, '-' || v_Fields || '-');
    
      /*      --固定时间
      If v_Count = 1 Then
        If Zl_To_Number(v_Fields) < 12 Then
          v_时间方案              := '早';
          v_时间方案编码          := '0800';
          v_收发记录.首次执行时间 := '0800';
          v_收发记录.末次执行时间 := '0800';
        Else
          v_时间方案              := '晚';
          v_时间方案编码          := '1800';
          v_收发记录.首次执行时间 := '1800';
          v_收发记录.末次执行时间 := '1800';
        
        End If;
      Elsif v_Count = 2 Then
        v_时间方案              := '早-晚';
        v_时间方案编码          := '0800-1800';
        v_收发记录.首次执行时间 := '0800';
        v_收发记录.末次执行时间 := '1800';
      Elsif v_Count = 3 Then
        v_时间方案              := '早-中-晚';
        v_时间方案编码          := '0800-1300-1800';
        v_收发记录.首次执行时间 := '0800';
        v_收发记录.末次执行时间 := '1800';
      End If;*/
    
      v_Count := v_Count + 1;
    End Loop;
  
    v_执行天数 := Zl_Getexedays(v_收发记录.开始执行时间, v_收发记录.首次时间, v_收发记录.末次时间, v_收发记录.频率间隔, v_收发记录.间隔单位, v_收发记录.执行时间方案);
    Insert Into Prescriptiondate
      (Prescriptionno, Seqno, Group_No, Machineno, Procflg, Patientid, Patientname, Sex, Ioflg, Wardcd, Wardname, Bedno,
       Prescriptiondate, Takedate, Taketime, Lasttime, Presc_Class, Drugcd, Drugname, Dispenseddose, Prescriptiondose,
       Prescriptionunit, Dispensedunit, Amount_Per_Package, Dispense_Days, Freq_Desc, Freq_Desc_Detail, Makerectime,
       Freq_Desc_Detail_Code, Explanation)
    Values
      (v_流水号, v_序号, v_小组标志, v_机器号, v_处理状态, v_收发记录.标识号, v_收发记录.姓名, v_收发记录.性别, v_门诊标志, v_收发记录.病区编码, v_收发记录.病区名称,
       v_收发记录.床号, v_收发记录.审核日期, v_收发记录.首次日期, v_收发记录.首次执行时间, v_收发记录.末次执行时间, v_紧急类别, v_收发记录.药品编码,
       v_收发记录.药品名称 || '/' || v_收发记录.单次用量 || v_收发记录.计算单位, v_收发记录.单次用量 / v_收发记录.剂量系数, v_收发记录.单次用量, v_收发记录.计算单位,
       v_收发记录.住院单位, v_收发记录.剂量系数, v_执行天数, v_收发记录.用法, v_时间方案, v_收发记录.审核日期, v_时间方案编码, v_收发记录.医生嘱托);
  End Loop;
  Commit;
End Zl_Getpackerdetail;
/
