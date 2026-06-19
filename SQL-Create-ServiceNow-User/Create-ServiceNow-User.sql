/*
================================================================================
  Script   : Create-ServiceNow-User.sql
  Purpose  : Creates a Windows-authenticated SQL Server login and database user
             for the local machine account <MachineName>\servicenow.
             The account is added to two databases:
               - master  (for server-level access)
               - msdb    (for SQL Agent / maintenance plan visibility)
  Usage    : Execute on each SQL Server instance where the ServiceNow
             discovery/integration user is required. The script derives the
             machine name dynamically, so no manual substitution is needed.
  Author   : M. Stam
  Date     : 2026-06-19
  Notes    : - Requires sysadmin or securityadmin + db_accessadmin rights.
             - The Windows account <MachineName>\servicenow must exist on the
               host before this script is run.
             - Run on each named instance separately (instances are not looped
               automatically — re-execute via SSMS against each instance).
================================================================================
*/

USE [master];
GO

DECLARE @server_name NVARCHAR(128);
DECLARE @sql        NVARCHAR(MAX);

SELECT @server_name = CAST(SERVERPROPERTY('MachineName') AS NVARCHAR(128));
PRINT @server_name;

-- Create Windows login
SET @sql = N'CREATE LOGIN [' + @server_name + N'\servicenow] FROM WINDOWS WITH DEFAULT_DATABASE=[master];';
EXEC sp_executesql @sql;

-- Create user in master
SET @sql = N'CREATE USER [' + @server_name + N'\servicenow] FOR LOGIN [' + @server_name + N'\servicenow];';
EXEC sp_executesql @sql;

SET @sql = N'ALTER USER [' + @server_name + N'\servicenow] WITH DEFAULT_SCHEMA=[dbo];';
EXEC sp_executesql @sql;
GO

-- Create user in msdb
USE [msdb];
GO

DECLARE @server_name NVARCHAR(128);
DECLARE @sql        NVARCHAR(MAX);

SELECT @server_name = CAST(SERVERPROPERTY('MachineName') AS NVARCHAR(128));

SET @sql = N'CREATE USER [' + @server_name + N'\servicenow] FOR LOGIN [' + @server_name + N'\servicenow];';
EXEC sp_executesql @sql;

SET @sql = N'ALTER USER [' + @server_name + N'\servicenow] WITH DEFAULT_SCHEMA=[dbo];';
EXEC sp_executesql @sql;
GO


