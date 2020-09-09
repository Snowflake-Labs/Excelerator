/* ----------------------- Snowflake Excel Integration ------------------------
    Run this script to set up permission for user that will use Snowflake Excelerator
	To run this script, use a role with elvated permissions such as ACCOUNTADMIN
	This scripts create a role named ExcelAnalyst that will be granted to another role passed in: <parentrole>. This is the role that the user can login with.
	Before running this scripts update the following placeholders with values specific to your environment
	<role>, <database>, <schema>, <warehouse>
	
	Certain privileges are needed for certain capabilities 
	This script does not deal with table level privileges, only schema and database level. Ideally the ParenRole will have the table privileges
	Here are the privileges needed for each capability:
	Query    			
		Database: 	USAGE
		Schema:		USAGE, CREATE STAGE - Unless a Stage is provided in the login, then the user will need USAGE to that Stage
		Table: 		SELECT
	Upload (everything in Query plus)
		Schema: 	CREATE TABLE
		Table: 		INSERT, UPDATE, TRUNCATE			
	Rollback (Schema level privileges above plus)
		Table: 		Ownership	 
	
	For advanced feature: Auto-generate Data Types
		usage on all procedures
		usage on future procedures
		usage on all functions
		usage on future functions
	
*/

-------------- Update this section -----------------------------------
-- User name and role that the user will login with 
set parentUserRole = '<parentrole>';
 -- database and schema where you will create the Snowflake Excel Integration stored procedures
set databaseName = '<database>';
set schemaName = '<schema>';
--warehouse that you will use
set warehouseName = '<warehouse>';
----------------------------------------------------------------
set roleName = 'ExcelAnalyst';
set databaseAndSchema = concat($databaseName,'.',$schemaName);

create role IF NOT EXISTS IDENTIFIER($roleName);
grant role IDENTIFIER($roleName) to role IDENTIFIER($parentUserRole);
grant usage on database IDENTIFIER($databaseName)  to role IDENTIFIER($roleName);
grant usage on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant create stage on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant create table on schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);

grant usage on warehouse IDENTIFIER($warehouseName) to IDENTIFIER($roleName);

-- For advanced feature: Auto-generate Data Types 
grant usage on future procedures in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on future functions in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on all procedures in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);
grant usage on all functions in schema IDENTIFIER($databaseAndSchema) to role IDENTIFIER($roleName);



