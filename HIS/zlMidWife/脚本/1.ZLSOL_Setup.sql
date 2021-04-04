--conn sys/his@ORA_SOL

Declare
  v_Count Number;
  v_Path  Varchar2(255);
Begin
  --检查是否存在该表空间
  Select Count(*) Into v_Count From Dba_Tablespaces Where Tablespace_Name = 'ZLSOL_DATA';

  If Nvl(v_Count, 0) = 0 Then
    --寻找基准表空间，以SYSTEM的表空间为基准，寻找文件路径，注意路径区分Win系统与Linux系统
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

--DBLink用于产程系统向ZLHIS同步新生儿记录

