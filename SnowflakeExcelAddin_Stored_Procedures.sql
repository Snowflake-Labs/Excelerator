/*
  Created July 2020 - Steve Segal
  These procs load and potentially create a table from a file in a stage
  If the table doesn't exist it will be Created. Any new columns will be added. 
  All data types will be derived from the data as best as possible unless the datatype is passed in the file in the first row
  The options for loading are: Merge. Append, Truncate & Load, Recreate table and Load
  The load is based on the column header of the file so order doesn't matter.
  It can be run on a directory of files or a single file
  
  Versions
  1.1.0 - Initial release
  1.1.1 - Added new Timestamp conversion type for Ryan T. WHEN DATE_STR RLIKE '\\d{8}\\s\\d{6}.\\d{3}' THEN TO_VARCHAR(TO_TIMESTAMP(DATE_STR, 'YYYYMMDD HH24MISS.FF'),'YYYY-MM-DD HH24:MI:SS.FF') 
  2.1.2 - Bug fix - issue when data types are provided but one is missing. Needed to set skip header = 2
*/


--**************************  get_stored_proc_version_number  **************************
create or replace procedure get_stored_proc_version_number ()
  returns string
  language javascript strict
  execute as caller
  as
// Workbook has to be less than or equal to the first digit and greater than or equal to the second
// Update the first digit for new features that new workbooks rely on
// Update the second digit if it breaks workbooks previously available
// 3rd digit is minor release and does not cause the need to upgrade
$$
return "2.1.2";
$$;
--**************************  create_table_from_file_and_load  **************************

create or replace procedure create_table_from_file_and_load (TABLE_NAME STRING, STAGE STRING, PATH STRING, COPY_TYPE STRING, PRIMARY_KEYS STRING,NUMBER_OF_COLUMNS DOUBLE)
  returns string
  language javascript strict
  execute as caller
  as
$$
// This proc creates a table based on a file and then uploads the data using the Copy command. The file must have a header row wich will be used to generate the column names
// TABLE_NAME: Name of the table to create
// STAGE: Name of the stage Example: MY_S3_STAGE
// PATH: Path to the directory or file. Does not include the stage. Example: 2020/March/file123.csv
// COPY_TYPE: 'RecreateTable','Truncate', 'Merge', 'Append', 'CreateTableOnly'
// Returns the create statement that was executed and the number of rows uploaded
var newLine = "\r\n"
var returnValue = 'Table exists, not creating. ';
var tableExists = true;
var backupTableCreated = false;
var backupTable;

try{
  try{
       snowflake.execute({ sqlText: `desc table ${TABLE_NAME};`});
    }
    catch(err){
      tableExists=false;
      if (COPY_TYPE!='CreateTableOnly'){
        COPY_TYPE = 'RecreateTable';
      }
      returnValue = 'Table does not exist. ';
    }

  if(tableExists){
    try{
        //Create backup just in case something fails. It needs to be random because this could be run concurrently
        backupTable = 'backuptable'+Math.floor((Math.random() * 10000) + 1);
        snowflake.execute({ sqlText: `create or replace table ${backupTable} clone ${TABLE_NAME};`});
        returnValue = `Creating backup table ${backupTable}. `;
        backupTableCreated = true;
      }
    catch(err){
        Throw (`Error backing up table ${TABLE_NAME}. ` + newLine + err )
      }
  }
  
  //Copy into table by column name which will add new columns
  if(COPY_TYPE=='Truncate'){      // *Truncate table    
    snowflake.execute({ sqlText: `truncate table ${TABLE_NAME};`});
  }
 
  stmt1 = snowflake.execute({ sqlText: `call create_table_from_file_and_load_work('${TABLE_NAME}','${STAGE}','${PATH}','${COPY_TYPE}','${PRIMARY_KEYS}',${NUMBER_OF_COLUMNS});`});
  stmt1.next();
  ret = stmt1.getColumnValue(1);
  if(ret.includes("ERROR:")){
    returnValue = ret;
  }
  else{
    returnValue = ret + newLine + returnValue;
  }
}
catch(err){
    returnValue = "ERROR:" + newLine + err + newLine + `Executing create_table_from_file_and_load ('${TABLE_NAME}', '${STAGE}', '${PATH}', '${COPY_TYPE}', '${PRIMARY_KEYS}')` + newLine + returnValue ;
    if(backupTableCreated){
    try{
        snowflake.execute({ sqlText: `alter table ${TABLE_NAME} swap with ${backupTable};`});
        }
        catch(err){
            returnValue += " Unable to restore backup, possibly beacuse the role doesn't have write access to the table. ";
        }
    }
}
if(backupTableCreated){
    snowflake.execute({ sqlText: `drop table ${backupTable};`});
}
return returnValue;

$$;


--**************************  create_table_from_file_and_load_work  **************************

create or replace procedure create_table_from_file_and_load_work (TABLE_NAME STRING, STAGE STRING, PATH STRING, COPY_TYPE STRING, PRIMARY_KEYS STRING,NUMBER_OF_COLUMNS DOUBLE)
  returns string
  language javascript strict
  execute as caller
  as
$$
// This proc creates a merge statment for the file in the stage and the table. If there are new columns in the file, it is added with an 'Alter' statement
// TABLE_NAME: Name of the table to create
// STAGE: Name of the stage Example: MY_S3_STAGE
// PATH: Path to the directory or file. Does not include the stage. Example: 2020/March/file123.csv
// PRIMARY_KEYS: comma separated list of the ordinal positions of primary keys for the table ex "1,2"
//NUMBER_OF_COLUMNS: The number of columns in the file. If set to -1 then the max is used
// Returns the create statement that was executed and the number of rows uploaded
// Example: merge_file_into_table('tableabc1','MY_INTERNAL_STAGE','tableabc1.csv','1,2');
 
var stageFullPath = "@"+STAGE+"/"+PATH; 
var alterStageHeader0 =`alter stage ${STAGE} set FILE_FORMAT = (FIELD_OPTIONALLY_ENCLOSED_BY = '"', SKIP_HEADER=0);`
var alterStageHeader1 =`alter stage ${STAGE} set FILE_FORMAT = (FIELD_OPTIONALLY_ENCLOSED_BY = '"', SKIP_HEADER=1);`
var alterStageHeader2 =`alter stage ${STAGE} set FILE_FORMAT = (FIELD_OPTIONALLY_ENCLOSED_BY = '"', SKIP_HEADER=2);`
var emptyFile = true;
var columnName =" ";
var TableCols = "";
var FileCols = "";
var FileColsNotConverted = "";
var stmt="";
var fileAlias = "fileSource";
var tableAlias ="t";
var matchByClause = "";
var matchByClauseIfChanged = "";
var updateClause = "";
var sql="";
var dbName ="";
var schemaName = "(select upper(current_schema()))";
var justTABLE_NAME = TABLE_NAME;
var uploadStmt="";
var newLine = "\r\n";
var columnDelimiter ='~';
var numOfBatchColumns=1001;
var dollar = '$';

var arrDataTypes = ["TEXT", "STRING", "VARCHAR","INTEGER", "NUMBER","DOUBLE", "TIMESTAMP", "DATE","FLOAT","VARIANT","ARRAY","BOOLEAN","OBJECT","TIME"];
var bIsDatatypeSupplied = false
var dataTypeFromFile = "";
var firstColumnFromFirstowFromFile = ""
var stmtDataTypesFromFile;
var stmtColumnNamesFromFile;
var arrayColunmNamesFromFile;
var arrayDataTypesFromFile ;
var ordinal = 0;
var returnVal="";
var result="";

try{
   //Check if params exists
  if(!TABLE_NAME || !STAGE || !PATH ||!COPY_TYPE){
    return `Error: At least one paramter is null: merge_file_into_table (TABLE_NAME=${TABLE_NAME}, STAGE=${STAGE}, PATH=${PATH}, COPY_TYPE=${COPY_TYPE})`
  }
      COPY_TYPE = COPY_TYPE.toUpperCase();
      snowflake.execute({ sqlText: alterStageHeader0}); // We are first retreiving headers
//First get the headers from the file
//Figure out how many columns to get from the file
      var i;      
      if(NUMBER_OF_COLUMNS>-1){
        numOfBatchColumns=NUMBER_OF_COLUMNS;
      }
      else
      {
        // Since we don't know how many columns in the file, I'm getting an estimate by pulling some columns and finding the first one that is null
        sqlStmt = `select iff($100 is null,'100',iff($300 is null,'300',iff($600 is null,'600',iff($900 is null,'900',iff($1200 is null,'1200',iff($1500 is null,'15000','')))))) from ${stageFullPath}  limit 1 offset 0`;
        var stmtGetFirstRow = snowflake.execute({ sqlText: sqlStmt});      
        stmtGetFirstRow.next();      
        numOfBatchColumns = stmtGetFirstRow.getColumnValue(1);          
      }
// Create the select statemnt: select concat($1,',',$2) ...      
      selectStmt="";
      for (i = 1; i <= numOfBatchColumns; i++) {
          selectStmt+=",'"+columnDelimiter+"',IfNULL($"+i+",'')";
      }
//Get create SQL and execute to get Col name - select ($1,$2,.. from filename limit 1)
      sqlStmt = `select concat(${selectStmt.substring(5)}) from ${stageFullPath} limit 1;`;
      var stmtGetFirstRow = snowflake.execute({ sqlText: sqlStmt});
      stmtGetFirstRow.next();
      var val = stmtGetFirstRow.getColumnValue(1);
      var arrayFirstRow = val.split(columnDelimiter);

      firstColumnFromFirstowFromFile = arrayFirstRow[0].toUpperCase();
// Check to see if the first col is a datatype. If not then its' the column headers
      if (firstColumnFromFirstowFromFile.includes("(") || arrDataTypes.includes(firstColumnFromFirstowFromFile) || firstColumnFromFirstowFromFile=="" ) {
          bIsDatatypeSupplied = true;
          dataTypeFromFile = firstColumnFromFirstowFromFile;
          arrayDataTypesFromFile = arrayFirstRow; 
// Since the data type is supplied then get the next row which are the headers
          var stmtColumnNamesFromFile = snowflake.execute({ sqlText: `select concat(${selectStmt.substring(5)}) from ${stageFullPath} limit 1 offset 1;`});
          stmtColumnNamesFromFile.next();
          val = stmtColumnNamesFromFile.getColumnValue(1);
          arrayColunmNamesFromFile = val.split(columnDelimiter);       
          columnName = arrayColunmNamesFromFile[0];
          if (columnName ==""){
              throw ("The first column does not have a name.")
          } 
      }
      else // Datatypes are not supplied so set the results and query statement to the column header variables
      {
          arrayColunmNamesFromFile = arrayFirstRow;
          columnName = firstColumnFromFirstowFromFile;
      }

  if(COPY_TYPE == "RECREATETABLE" || COPY_TYPE == "CREATETABLEONLY"){
// Get data types for each column 
       if (bIsDatatypeSupplied) {
            sql = alterStageHeader2      
       }
       else{
            sql = alterStageHeader1       
       }       
      snowflake.execute({ sqlText: sql});  
      while (columnName)
      {    
        //If the column name is null then we are done
        if(columnName)
        {
          emptyFile = false;
        if(dataTypeFromFile==""){
            //Get the data type from get_datatype proc
            result = snowflake.execute({ sqlText: `call get_datatype(${ordinal+1},'${stageFullPath}');`});
            result.next();
          dataType = result.getColumnValue(1);
          }
          else{
            dataType = dataTypeFromFile;
          }

          TableCols +=  "," + columnName + " " + dataType;
          ordinal+=1;
          columnName = arrayColunmNamesFromFile[ordinal];
          if (bIsDatatypeSupplied) {
            dataTypeFromFile =  arrayDataTypesFromFile[ordinal]
           }
        }
      }
       if(emptyFile)
          {Throw ("No data in first cell")}
       else{
          try{
            TableCols = TableCols.substring(1); //Remove leading comma
            var finalCreate = `create or replace table ${TABLE_NAME} (${TableCols});`     
            snowflake.execute({ sqlText: finalCreate});
            returnVal  = finalCreate;
          }
          catch(err){
              throw("Attempted to execute SQL: "+finalCreate+" Error:"+ newLine +err);
          }
        } 
      }      
      //Debug return finalCreate
       if (COPY_TYPE == "CREATETABLEONLY"){
        return returnVal
       }
  //*******************************************    End create   *************************************//  

   if(bIsDatatypeSupplied){
     snowflake.execute({ sqlText: alterStageHeader2}); 
   }   
   else{
    snowflake.execute({ sqlText: alterStageHeader1}); 
   }
    TableCols = ""; // This needs to be reset 
    dataType = "";
      //create array of keys for Match by
    var arrayKeys = PRIMARY_KEYS.split(",");
      //Creat map of existing columns
    var arrColsFromTable=[];
    var mapColDatType = new Map();
    sqlStmt = `desc table ${TABLE_NAME};`
    var stmtColumnsFromTable = snowflake.execute({ sqlText: sqlStmt});
    var colName;
    while (stmtColumnsFromTable.next())
    {
      colName = stmtColumnsFromTable.getColumnValue(1).toUpperCase()
      arrColsFromTable.push(colName);
      mapColDatType[colName] = stmtColumnsFromTable.getColumnValue(2)
    }
    columnName = arrayColunmNamesFromFile[0];
    ordinal = 0;
    //In the loop, execute a query to get the column name, then get the data type
    while (columnName)
    {
    //Does column exist in table? If not create an alter statement
        if(!arrColsFromTable.includes(columnName.toUpperCase()))  { 
            if (bIsDatatypeSupplied){
                dataType = arrayDataTypesFromFile[ordinal].toUpperCase();  
            }
            result="";
            if(dataType==""){  
               //Get the data type from get_datatype proc since it didn't come from the file
              result = snowflake.execute({ sqlText: `call get_datatype(${ordinal+1},'${stageFullPath}');`});
              result.next();
              dataType = result.getColumnValue(1);
            }
              // Create column
              sqlstmt = `ALTER TABLE ${TABLE_NAME} ADD COLUMN ${columnName} ${dataType};`;
              snowflake.execute({ sqlText:sqlstmt});
              result +=" "+sqlstmt;
        }
        else{
          dataType = mapColDatType[columnName.toUpperCase()];
        }
        emptyFile = false;
        ordinalForSelect = ordinal + 1; // have to do this becuase the arrays start at 0 but the columns start at 1
        FileColsNotConverted += ",$" + ordinalForSelect;
        if (dataType=="DATE"){            
            FileCols += ",convert_to_date($" + ordinalForSelect+")";
        }
        else{
          if (dataType.substring(0,9) == "TIMESTAMP"){
              FileCols += ",convert_to_datetime($" + ordinalForSelect+")";
          } 
          else{
              if (dataType.substring(0,4)=="TIME"){
                  FileCols += ",convert_to_time($" + ordinalForSelect+")";
              }
              else{
                   FileCols += ",$" + ordinalForSelect;  
              }
          } 
        }

          TableCols +=  ", " + columnName; 
          if(COPY_TYPE=='MERGE')
          {
            if (arrayKeys.includes(ordinalForSelect.toString())){
                matchByClause += " and " + fileAlias +".$"+ ordinalForSelect + "=" + tableAlias + "." + columnName ;  
               }
            else{
                    columnWithCast = `${fileAlias}.${dollar}${ordinalForSelect}`
                    if (dataType=="BOOLEAN"){
                        columnWithCast = `TO_BOOLEAN(${columnWithCast})`
                    }
                    if (dataType.substring(0,6)=="NUMBER"){
                        columnWithCast = `TO_NUMBER(${columnWithCast})`
                    }
                  updateClause += ", " + columnName + " = "+ fileAlias + ".$"+ ordinalForSelect;
                   matchByClauseIfChanged +=  ` and ${columnWithCast} = ${tableAlias}.${columnName} and `;
                   matchByClauseIfChanged +=  ` not(${tableAlias}.${columnName} is Null and ${fileAlias}.${dollar}${ordinalForSelect} is not null ) and `;
                   matchByClauseIfChanged +=  ` not(${tableAlias}.${columnName} is not Null and ${fileAlias}.${dollar}${ordinalForSelect} is null )`;   
                }
          }
        ordinal+=1;
        columnName = arrayColunmNamesFromFile[ordinal];
        dataType="";
        if (bIsDatatypeSupplied){
            dataType = arrayDataTypesFromFile[ordinal];
        }
    }
    if(emptyFile)
      {return "No data in first cell"}
    else{
      TableCols = TableCols.substring(1); //Remove leading comma
      FileCols = FileCols.substring(1); //Remove leading comma
      FileColsNotConverted = FileColsNotConverted.substring(1); //Remove leading comma
      if(COPY_TYPE=='MERGE')
      {
          matchByClause = matchByClause.substring(5); // Remove the leading " and "
          matchByClauseIfChanged = matchByClauseIfChanged.substring(5); // Remove the leading " and "          
          updateClause = updateClause.substring(1); //Remove leading comma

          uploadStmt = `merge into ${TABLE_NAME} ${tableAlias} using (select ${FileCols} from ${stageFullPath}) as ${fileAlias} on ${matchByClause} `
          if (updateClause!=""){
              uploadStmt += `WHEN MATCHED AND NOT(${matchByClauseIfChanged}) THEN UPDATE SET ${updateClause} `
             // uploadStmt += `WHEN MATCHED  THEN UPDATE SET ${updateClause} `
          }
          uploadStmt += `WHEN NOT MATCHED THEN INSERT (${TableCols}) VALUES (${FileColsNotConverted});`     
       }
       else
       {
          uploadStmt = `insert into ${TABLE_NAME} (${TableCols})  (select ${FileCols} from ${stageFullPath});` 
       }
      try{
      //**************** This is the final Insert/Merge *************//////////////
        stmt = snowflake.execute({ sqlText: uploadStmt});
        stmt.next(); 
      }
      catch(err){
        if(err.message.includes("is not recognized")) {
          return("ERROR: Datatype incorrectly set. "+err)
        }
        if(err.message.includes("Duplicate row detected during DML action")) {
          return("ERROR: Merge keys are not defined properly. There was a duplicate key detected in the file.")
        }
        throw(err)
      }
       
      if(COPY_TYPE=='MERGE')
      {
      var rowsUpdated = 0;
        if (updateClause!=""){
            rowsUpdated = stmt.getColumnValue(2);
        }        
          returnVal = `Rows Inserted: ${stmt.getColumnValue(1)}, Rows Updated: ${rowsUpdated}` + newLine + returnVal;
      }
      else
      {
        returnVal = `Rows Inserted: ${stmt.getColumnValue(1)}` + newLine + returnVal;
      }
      return returnVal + "Satement executed: " + uploadStmt;
    };
}
catch(err){

    throw(err)
}
  
 $$;


--**************************  get_datatype  **************************

create or replace procedure get_datatype(ORDINAL DOUBLE, STAGE STRING)
  returns string
  language javascript strict
  as
$$
// ORDINAL: the position of a column in a file 
// STAGE: The fully qualified path to the file or directory including stage. Example: MY_S3_STAGE/2020
// Returns the data type of the column

var datatype = "string";
var found = false; 
var dollar = '$';

//Check for number

try{    
  snowflake.execute({ sqlText: `select count(*) from ${STAGE} where $` + ORDINAL + " = 123.123 ;" });
  // get the length and subtract the position of the deimcal. If there is no decimal the position to be null. If the the math is null then set the final value to 0
  var stmt = snowflake.execute({ sqlText: `select IFNULL(max(LENGTH(${dollar}${ORDINAL})- NULLIFZERO(POSITION( '.' IN ${dollar}${ORDINAL} ))),0) from ${STAGE};`});
  stmt.next(); 
  var scale = stmt.getColumnValue(1);
 
  return `NUMBER(38,${scale})`
}
catch(err)
{
    found = false;
}
  
// Checking for any of the date time types. First see if it's over 12 chars. If it is it can't be a date or time but it could be a DateTime
  var stmt = snowflake.execute({ sqlText: "select max(Len($" + ORDINAL + ")) from "+STAGE+";"});
  stmt.next();
  var length = stmt.getColumnValue(1);
  if(length>12) //check to see if it's a DateTime
    {
    try{
        snowflake.execute({ sqlText: "select convert_to_datetime($" + ORDINAL + ") from "+STAGE+";"});
        return "TIMESTAMP";
        }
        catch(err)
    {
    found = false;
    }
  }
//check if it's a Date
try{
    snowflake.execute({ sqlText: "select convert_to_date($" + ORDINAL + ") from "+STAGE+";"});
    return "DATE";
}
catch(err)
{
found = false;
}
//Check if it's a Time
try{
    snowflake.execute({ sqlText: "select convert_to_time($" + ORDINAL + ") from "+STAGE+";"});
    return "TIME";
}
catch(err)
  {
  found = false;
  }

// Check for boolean
try{    
     snowflake.execute({ sqlText:"select  cast($" + ORDINAL + " as boolean)from "+STAGE+";"});
     return "BOOLEAN"
   }
catch(err)
  {
    found = false;
  }
  
  return datatype;
$$;

--**************************  convert_to_date  **************************

create or replace function convert_to_date(DATE_STR string)
  returns date
  language sql
  as
$$
select CASE 
         WHEN DATE_STR RLIKE '\\d{4}-\\d{2}-\\d{2}' THEN TO_DATE(DATE_STR, 'YYYY-MM-DD')
         WHEN DATE_STR RLIKE '\\d{4}/\\d{2}/\\d{2}' THEN TO_DATE(DATE_STR, 'YYYY/MM/DD')
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{2}' THEN TO_DATE(DATE_STR, 'MM/DD/YY')
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{4}' THEN TO_DATE(DATE_STR, 'MM/DD/YYYY')
         WHEN DATE_STR RLIKE '\\w+\\s\\d{1,2},\\s\\d{4}' THEN TO_DATE(DATE_STR, 'MON DD, YYYY')
         WHEN DATE_STR RLIKE '\\d{1,2}-\\w+-\\d{4}' THEN TO_DATE(DATE_STR, 'DD-MON-YYYY')     
         WHEN DATE_STR RLIKE '\\d{1,2}-\\w+-\\d{2}' THEN TO_DATE(DATE_STR, 'DD-MON-YY')
         WHEN DATE_STR RLIKE '\\d{2}/\\d{6}' THEN TO_DATE(DATE_STR, 'MM/DDYYYY')
         WHEN DATE_STR RLIKE '\\d{8}' THEN TO_DATE(DATE_STR, 'YYYYMMDD') 
         ELSE  DATE_STR
       END AS DATE_STR_TO_DATE
$$;

--**************************  convert_to_datetime  **************************

create or replace function convert_to_datetime(DATE_STR string)
  returns timestamp
  language sql
  as
$$
select CASE 
         //Timestamp down to the Second
         WHEN DATE_STR RLIKE '\\d{4}-\\d{2}-\\d{2}\\s\\d{2}:\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YYYY-MM-DD HH24:MI:SS')
         WHEN DATE_STR RLIKE '\\d{4}/\\d{2}/\\d{2}\\s\\d{2}:\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YYYY/MM/DD HH24:MI:SS')
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{4}\\s\\d{2}:\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'MM/DD/YYYY HH24:MI:SS')
         WHEN DATE_STR RLIKE '\\d{1,2}-\\d{1,2}-\\d{4}\\s\\d{2}:\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'MM-DD-YYYY HH24:MI:SS')
          //Timestamp down to the Second and 2 digit year
         WHEN DATE_STR RLIKE '\\d{2}-\\d{2}-\\d{2}\\s\\d{2}:\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YY-MM-DD HH24:MI:SS')
         WHEN DATE_STR RLIKE '\\d{2}/\\d{2}/\\d{2}\\s\\d{2}:\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YY/MM/DD HH24:MI:SS')
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{2}\\s\\d{2}:\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'MM/DD/YY HH24:MI:SS') 
         //Timestamp down to the Minute
         WHEN DATE_STR RLIKE '\\d{4}-\\d{2}-\\d{2}\\s\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YYYY-MM-DD HH24:MI')
         WHEN DATE_STR RLIKE '\\d{4}/\\d{2}/\\d{2}\\s\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YYYY/MM/DD HH24:MI')
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{4}\\s\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'MM/DD/YYYY HH24:MI')
         //Timestamp down to the Minute and 2 digit year
         WHEN DATE_STR RLIKE '\\d{2}-\\d{2}-\\d{2}\\s\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YY-MM-DD HH24:MI')
         WHEN DATE_STR RLIKE '\\d{2}/\\d{2}/\\d{2}\\s\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'YY/MM/DD HH24:MI')
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{2}\\s\\d{2}:\\d{2}' THEN TO_TIMESTAMP(DATE_STR, 'MM/DD/YY HH24:MI')
         //HH12
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{2}\\s\\d{1,2}:\\d{1,2}\\s\\w{2}' THEN TO_TIMESTAMP(DATE_STR, 'MM/DD/YY HH12:MI AM')
         WHEN DATE_STR RLIKE '\\d{1,2}/\\d{1,2}/\\d{4}\\s\\d{1,2}:\\d{1,2}:\\d{1,2}\\s\\w{2}' THEN TO_TIMESTAMP(DATE_STR, 'MM/DD/YYYY HH12:MI:SS AM') 
         //Special Case example - '20200809 042319.000'
         WHEN DATE_STR RLIKE '\\d{8}\\s\\d{6}.\\d{3}' THEN TO_TIMESTAMP(DATE_STR, 'YYYYMMDD HH24MISS.FF')
         ELSE DATE_STR
       END AS DATE_STR_TO_DATE
$$;
--**************************  convert_to_time  **************************

create or replace function convert_to_time(DATE_STR string)
  returns time
  language sql
  as
$$
select CASE 
          WHEN DATE_STR RLIKE '\\d{1,2}:\\d{2}:\\d{2}' THEN TO_TIME(DATE_STR, 'HH24:MI:SS')
          WHEN DATE_STR RLIKE '\\d{1,2}:\\d{2}:\\d{2}\\s\\w{2}' THEN TO_TIME(DATE_STR, 'HH12:MI:SS AM')
          WHEN DATE_STR RLIKE '\\d{1,2}:\\d{2}' THEN TO_TIME(DATE_STR, 'HH24:MI')
          WHEN DATE_STR RLIKE '\\d{1,2}:\\d{2}\\s\\w{2}' THEN TO_TIME(DATE_STR, 'HH12:MI AM')
          ELSE DATE_STR
       END AS DATE_STR_TO_DATE
$$;

----------------------------------------------------------

create or replace procedure createStage(STAGENAME STRING)
  returns string
  language javascript strict
  execute as caller
  as
$$
var sqlString = `create or replace stage ${STAGENAME} file_format = (type=csv, SKIP_HEADER=0)`;
try{
   snowflake.execute({ sqlText: sqlString});
   return `Stage created: ${STAGENAME}`
}    
catch(err){
    throw(`Error creating ${STAGENAME}. ${err}`)
}

$$;

----------------------------------------------------------

create or replace procedure dropStage(STAGENAME STRING)
  returns string
  language javascript strict
  execute as caller
  as
$$
var sqlString = `drop stage ${STAGENAME}`;
try{
   snowflake.execute({ sqlText: sqlString});
    return `Stage dropped: ${STAGENAME}`
}    
catch(err){
    throw(`Error dropping ${STAGENAME}. ${err}`)
}

$$;

create or replace procedure rollbackTableWithOffset(TABLE_NAME STRING, OFFSETVALUE DOUBLE)
  returns string
  language javascript strict
  execute as caller
  as
$$
try{
  snowflake.execute({ sqlText: `create or repalce table ${TABLE_NAME}_temp clone ${TABLE_NAME} at (offset =>${OFFSETVALUE});`});
  snowflake.execute({ sqlText: `alter table ${TABLE_NAME} swap with ${TABLE_NAME}_temp;`});
  snowflake.execute({ sqlText: `drop table ${TABLE_NAME}_temp;`});
  return `Rolled back ${TABLE_NAME} as of ${OFFSETVALUE} seconds ago.`
}    
catch(err){
    throw(`Error rolling back ${TABLE_NAME}. ${err}`)
}

$$;