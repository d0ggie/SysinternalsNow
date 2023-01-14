Sysinternals Now
================

This simple PowerShell script fetches the latest Sysinternals advanced system utilities.  It uses no externals components and requires a relatively old PowerShell version.  This is done so that the script is usable with Windows Server family that unfortunately lacks (an easy) access to many of the other package distribution systems.

Utilities are stored to `Content/` and related server provided information to `Content-Cache/` as simple JSON files.  The most interesting detail is HTTP ETag (Entity Tag) that is used (if available) when performing the HTTP query.

For more details please use `Get-Help` to read embedded comment-based help and/or study the script itself.
