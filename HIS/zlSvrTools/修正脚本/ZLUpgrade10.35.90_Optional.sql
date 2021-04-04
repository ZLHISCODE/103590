----10.35.80---��10.35.90
--000000:����,2017-12-07,ɾ���ڼ��ZLPERIODS
Create Or Replace Procedure Zltools.Zl_Optional_�ڼ��_ɾ�� Is
  n_Count Number;
Begin
  Select Count(1) Into n_count From all_tables Where owner = 'ZLTOOLS' And table_name = 'ZLPERIODS';
  If n_count > 0 Then
    Execute Immediate 'DROP TABLE ZLTOOLS.ZLPERIODS';
    Execute Immediate 'DROP PUBLIC SYNONYM ZLPERIODS';
  End If;
  Execute Immediate 'DELETE FROM ZLTABLES WHERE ����=' || Chr(39) || 'ZLPERIODS' || Chr(39);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Optional_�ڼ��_ɾ��;
/

--117980:����,2017-12-26,ɾ������ģ����_bak�ֶ�
Create Or Replace Procedure Zltools.Zl_Optional_��־��_ɾ���ֶ� Is
Begin
  If ZL_CheckObject(2,'����ģ����_bak','Zlauditlog') = 1 then
    Execute Immediate 'ALTER TABLE Zlauditlog DROP COLUMN ����ģ����_bak';
  End if;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Optional_��־��_ɾ���ֶ�;
/