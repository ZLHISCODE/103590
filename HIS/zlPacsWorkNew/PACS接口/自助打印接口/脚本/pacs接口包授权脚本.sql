-- Create the user 
create user PACSUSER
  identified by "HIS"
  default tablespace USERS
  temporary tablespace TEMP
  profile DEFAULT;
-- Grant/Revoke object privileges 
grant execute on b_PacsInterface to PACSUSER;
-- Grant/Revoke role privileges 
grant connect to PACSUSER;
