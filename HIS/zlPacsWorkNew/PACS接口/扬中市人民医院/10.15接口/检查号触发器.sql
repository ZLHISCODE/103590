---解决病理科要求的，在填写“病理编号”的时候，同时修改“检查号”
create or replace trigger zl_病人医嘱发送_报告ID
  after update on 病人医嘱发送  
  for each row
Declare
  v_内容 Number(18);
Begin
  If :New.报告id Is Null Then
    Return;
  End If;

  Begin
    Select To_Number(内容)
    Into v_内容
    From 病人病历内容 A, 病人病历文本段 B
    Where A.ID = B.病历id And A.病历记录id = :New.报告id And A.标题文本 = '病理编号' And Rownum = 1;
  Exception
    When Others Then
      v_内容 := Null;
  End;
  If v_内容 Is Null Then
    Return;
  End If;

  Update 影像检查记录 Set 检查号 = v_内容 Where 医嘱id = :Old.医嘱id;
End Zl_病人医嘱发送_报告id;
/
