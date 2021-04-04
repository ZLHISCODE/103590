--conn sys/his@ORA_SOL

Declare
  v_Count Number;
  v_Path  Varchar2(255);
Begin
  --����Ƿ���ڸñ�ռ�
  Select Count(*) Into v_Count From Dba_Tablespaces Where Tablespace_Name = 'ZLSOL_DATA';

  If Nvl(v_Count, 0) = 0 Then
    --Ѱ�һ�׼��ռ䣬��SYSTEM�ı�ռ�Ϊ��׼��Ѱ���ļ�·����ע��·������Winϵͳ��Linuxϵͳ
    Select Substr(File_Name, 1,
                   Decode(Instr(File_Name, '\', -1), 0, Instr(File_Name, '/', -1), Instr(File_Name, '\', -1)))
    Into v_Path
    From Dba_Data_Files
    Where Tablespace_Name = 'SYSTEM' And Rownum < 2;
  
    Execute Immediate 'Create Tablespace ZLSOL_DATA Datafile ''' || v_Path ||
                      'ZLSOL_DATA.DBF'' SIZE 50M REUSE AUTOEXTEND ON NEXT 50M  EXTENT MANAGEMENT LOCAL AUTOALLOCATE';
  End If;
End;
/

create user ZLSOL identified by zlsoft;
alter user ZLSOL Default Tablespace ZLSOL_DATA;
alter user ZLSOL Temporary Tablespace TEMP;
Grant Connect,Resource,UNLIMITED TABLESPACE,CREATE VIEW,create database link,CREATE DIMENSION,CREATE JOB,CREATE MATERIALIZED VIEW, CREATE SYNONYM to ZLSOL;

--DBLink���ڲ���ϵͳ��ZLHISͬ����������¼

