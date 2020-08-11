/* ----------------------- SnoFlo Excel Integration ------------------------
    Run this script to set up permission for user that will use SnoFlo Snowflake Excel Integration
	The role should have the privileges to create a role and grant the appropriate privileges on the DB objects
	Before running this scripts update the following placeholders with values specific to your environment
	<role>, <database>, <schema>, <warehouse>
	
*/

-------------- Update this section -----------------------------------
-- User name and role that the user will login with 
set currentUserRole = '<role>';
 -- database and schema where you will create the SnoFlo Excel Integration stored procedures
set databaseName = '<database>';
set schemaName = '<schema>';
--warehouse that you will use
set warehouseName = '<warehouse>';
----------------------------------------------------------------
set roleName = 'ExcelAnalyst';
set databaseAndSchema = concat($databaseName,'.',$schemaName);

create role IF NOT EXISTS IDENTIFIER($roleName);
grant role IDENTIFIER($roleName) to role IDENTIFIER($currentUserRole);
grant usage on database IDENTIFIER($databaseName)  to role IDENTIFIER($roleName);
grant usage on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant create stage on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant create table on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);

grant usage on future procedures in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on future functions in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on all procedures in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on all functions in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);

grant usage on warehouse IDENTIFIER($warehouseName) to IDENTIFIER($roleName);

