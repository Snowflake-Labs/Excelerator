/* ----------------------- Snowflake Excel Integration ------------------------
    This script is only needed when updating Snowflake from Excel. 
	An alternative to running this script, is to grant the user's role the 'Create Stage' and 'Create Table' privileges on the schema they will login to.

	This script allows for update from Excel without changing the user's role directly. Instead, a new role is created, 'ExcelAnalyst',  and assigned to the user's role.
	To allow multiple roles to have update access, run this script one time for the first role and then assign the 'ExcelAnalyst' role manually to the other roles.
	
	To run this script, use either role ACCOUNTADMIN or SECURITYADMIN
	This script creates a role named 'ExcelAnalyst' that will be granted to the role specified below as 'existingUserRole'. This is the role that has been granted to the Excel user.
	Before running this scripts update the following placeholders with values specific to your environment
	<existingUserRole>, <database>, <schema>, <warehouse>
	
*/

-------------- Update this section -----------------------------------
-- User name and role that the user will login with 
set existingUserRole = '<existingUserRole>';
 -- database and schema where you will create the Snowflake Excel Integration stored procedures
set databaseName = '<database>';
set schemaName = '<schema>';
--warehouse that you will use
set warehouseName = '<warehouse>';
----------------------------------------------------------------
set roleName = 'ExcelAnalyst';
set databaseAndSchema = concat($databaseName,'.',$schemaName);

create role IF NOT EXISTS IDENTIFIER($roleName);
grant role IDENTIFIER($roleName) to role IDENTIFIER($existingUserRole);

grant create stage on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
-- For rollback functionality
grant create table on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);

------------------------------------ For advanced feature: Auto-generate Data Types ------------------------------------ 
-- Grant access to the DB and Schema where stored procs are created
grant usage on database IDENTIFIER($databaseName)  to role IDENTIFIER($roleName);
grant usage on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
-- Grant access to the stored procs
grant usage on future procedures in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on future functions in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on all procedures in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on all functions in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);



