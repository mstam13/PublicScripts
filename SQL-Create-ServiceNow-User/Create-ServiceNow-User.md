# Create-ServiceNow-User

## Synopsis

Creates a Windows-authenticated SQL Server login and database user for the
local machine account `<MachineName>\servicenow`.

## Description

This script provisions the ServiceNow integration account on a SQL Server
instance. It dynamically derives the machine name from `SERVERPROPERTY('MachineName')`,
so no manual substitution of server names is required.

The account is created in two databases:

| Database | Purpose |
| --- | --- |
| `master` | Server-level login and user for general connectivity |
| `msdb` | User for SQL Agent / maintenance plan visibility |

## Prerequisites

- The Windows local account `<MachineName>\servicenow` must exist on the host
  **before** this script is run.
- The executing account must hold at minimum:
  - `securityadmin` fixed server role (to create logins), **and**
  - `db_accessadmin` in `master` and `msdb` (to create users), **or**
  - `sysadmin` fixed server role.

## Usage

Execute against each SQL Server instance where ServiceNow discovery or
integration is required. The script does **not** loop over multiple instances
automatically; run it once per instance from SSMS or `sqlcmd`.

```text
-- Using sqlcmd (replace SERVER\INSTANCE and adjust -E / -U/-P as needed)
sqlcmd -S SERVER\INSTANCE -E -i Create-ServiceNow-User.sql
```

## What the Script Does

1. Switches to the `master` database.
2. Reads the machine name via `SERVERPROPERTY('MachineName')`.
3. Creates the Windows login `<MachineName>\servicenow` with `master` as the
   default database.
4. Creates a database user for that login in `master` with default schema `dbo`.
5. Switches to the `msdb` database.
6. Creates a database user for the same login in `msdb` with default schema `dbo`.

## Notes

- Named instances each have their own independent login/user store; run the
  script separately against every instance (e.g. `SERVER\INST1`, `SERVER\INST2`).
- The `DEFAULT_SCHEMA` is set to `dbo` in both databases.
- No roles or permissions beyond basic user creation are granted by this script;
  grant additional rights separately as required by the ServiceNow integration.

## Version History

| Version | Date | Author | Change |
| --- | --- | --- | --- |
| 1.0 | 2026-06-19 | M. Stam | Initial version |
