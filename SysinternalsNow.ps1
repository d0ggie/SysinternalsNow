#Requires -Version 5.1

# =============================================================================
#   SYSINTERNALS NOW
# -----------------------------------------------------------------------------

# Public domain.  My apologies for the syntax and long lines.

<#
.SYNOPSIS
Sysinternals Now fetches the latest Sysinternals utilities from Sysinternals Live service.

.DESCRIPTION
Sysinternals Now fetches and copies all Sysinternals utilities that are available on remote Sysinternals Live service to the local computer.  The Sysinternals Live service is essentially a simple web page, accessed using a normal HTTP query, that points to individual files of the Sysinternals utilities.

It should be noted that the Sysinternals Live service also contains non-troubleshooting tools that are not present e.g. in the Sysinternals Suite package.  These are not excluded.  However, for some reason, the service also lists files that are seemingly inaccessible.  These files are excluded by default.

Individual utility files are stored inside `Content` directory relative to the current working directory.  The directory can be changed by supplying an optional `-Content` parameter or an optional `SysinternalsNow__Content` environment variable.  The command line paramter has higher presedence than the environment variable, should the both be present.  The name of this directory cannot be empty.  To effectively use an empty directory, please e.g. supply `.` as the value.

When an individual file is fetched to local computer the file last modified timestamp is chaged to match that of the remote.  If the local timestamp is newer that of the remote the local file is not updated by default.  This is to avoid any accidents.  To override this behavior please supply `-IgnoreLocalTimestamp` switch parameter.

To minimize unnecessary bandwidth usage when a file is fetched a corresponding JSON file containing cache information is created.  The JSON file contains the last modofied timestamp, HTTP ETag (Entity Tag) and the Live URI.  These files are stored inside `ContentCache` directory.  The name of this directory cannot be empty.  Like above, this directory can be supplied using an optional `-ContentCache` parameter or an optional `SysinternalsNow__ContentCache` environment variable.  The presedence, should the both values be present, is simlar to above as well.

.LINK
https://learn.microsoft.com/sysinternals/
.LINK
https://live.sysinternals.com/

.INPUTS
None.

.OUTPUTS
Sysinternals utilties stored to `Content` directory and corresponding cache files stored to `ContentCache` directory.
#>

[CmdletBinding()]
Param (
    # URI of the Sysinternals Live service.
    [ValidateNotNullOrEmpty()]
  [String] $LiveURI = 'https://live.sysinternals.com/',

    # Name of the directory in which to put contents.
  [String] $Content,
    # Name of the directory in which to put content caching information.
  [String] $ContentCache,

    # Ignore local file timestamp.
  [Switch] $IgnoreLocalTimestamp = $False
)

Set-StrictMode -Version 3.0

Function Get-Environment-Variable
  {
  Param (
    [Parameter(Mandatory)]
      [ValidateNotNull()]
    [String] $Name,
    [String] $DefaultValue
  )

  $EnvironmentValue = [System.Environment]::GetEnvironmentVariable($Name)
  If ($EnvironmentValue)
    {
    Write-Debug ('Environment override "{0}" = "{1}".' -f `
      $Name, $EnvironmentValue)
    $Value = $EnvironmentValue
    }
  Else
    {
    $Value = $DefaultValue
    }
  Return $Value
  }

If (-Not $Content)
  {
$Content = Get-Environment-Variable `
  -Name 'SysinternalsNow__Content' -DefaultValue 'Content'
  }
If (-Not $ContentCache)
  {
$ContentCache = Get-Environment-Variable `
  -Name 'SysinternalsNow__ContentCache' -DefaultValue 'Content-Cache'
  }

Function Get-Forward-Arguments
  {
  Param (
    [Parameter(Mandatory)]
      [ValidateNotNull()]
    [Management.Automation.CmdletInfo] $Command,
    [Parameter(ValueFromRemainingArguments)]
    $ArgumentsRemaining
  )

  $ArgumentsRemaining | ForEach-Object -Begin {

  $ArgumentsForward = @{
    }
  $ParameterName = $Null

  } -Process {

  If ($_ -Match '^-(?<ParameterName>[a-z]+):?$')
    {
  $ParameterName = $Matches.ParameterName
  If ($Command.Parameters.ContainsKey($ParameterName))
    {
    If ($Command.Parameters.Item($ParameterName).SwitchParameter)
      {
      $ArgumentsForward.Item($ParameterName) = $True
      }
    }
  Else
    {
  Write-Warning ('"{0}" does not have a parameter named "{1}".' -f `
    $Command.Name, $ParameterName)
    }

    }
  ElseIf ($ParameterName)
    {
  $ArgumentsForward.Item($ParameterName) = $_
    }
  }

  Return $ArgumentsForward
  }

Function New-ParentDirectoryItem
  {
  [CmdletBinding()]
  [OutputType([Void])]
  Param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [String] $Path
  )

  [ValidateNotNullOrEmpty()]
  $Path__ParentDirectory = Split-Path -Path $Path -Parent

  If (-Not (Test-Path -Path $Path__ParentDirectory))
    {
  New-Item -ItemType Directory -Path $Path__ParentDirectory | Out-Null
    }

  Return
  }

Function Has-Property
  {
  [CmdletBinding()]
  [OutputType([Bool])]
  Param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [Object] $Object,
    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $Name
  )

  Return ([PSCustomObject] $Object).PSObject.Properties[$Name] -Ne $Null
  }

Function Get-Property
  {
  [CmdletBinding()]
  [OutputType([String])]
  Param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [Object] $Object,
    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $Name
  )

  If ($Object | Has-Property -Name $Name)
    {
  Return $Object.$Name
    }
  Else
    {
  Return $Null
    }
  }

Function Get-ItemPropertyCache__NotNull
  {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  Param (
    [Parameter(Mandatory, ValueFromPipeline)]
      [ValidateNotNullOrEmpty()]
    [Object] $FileInfo
  )

  $FileInfo__Out = [PSCustomObject] @{
    }
  @( 'ModifiedTime', 'ETag', 'FileURI' ) | %{
    If ($FileInfo | Has-Property -Name $_)
      {
  $FileInfo__Out | Add-Member -NotePropertyName $_ -NotePropertyValue $FileInfo.$_
      }
    }

  Return $FileInfo__Out;
  }

Function Get-ItemPropertyCache
  {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  Param (
    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $FileInfo__Storage
  )

  If (Test-Path -Path $FileInfo__Storage)
    {
  $FileInfo = Get-Content -Raw -Path $FileInfo__Storage | ConvertFrom-Json | Get-ItemPropertyCache__NotNull
    }
  Else
    {
  $FileInfo = [PSCustomObject] @{
    }
  }

  Return $FileInfo
  }

Filter Filter-WebResponseHeader__ETag
  {
  If ($_.Headers.ContainsKey('ETag') -And $_.Headers['ETag'] -Match '"(?<ETag>[!#-~]+)"')
    {
  Return $Matches.ETag
    }

  Return $Null
  }

Filter Filter-WebResponseHeader__Last-Modified
  {
  If ($_.Headers.ContainsKey('Last-Modified'))
    {
    $LastModified = ([DateTime] $_.Headers['Last-Modified']).ToUniversalTime()
    If ($LastModified -Gt (Get-Date -Year 1999 -Day 1 -Month 1))
      {
  Return $LastModified
      }
    }

  Return $Null
  }

Function Invoke-WebRequest__Silent
  {
# Note: If progress indication is enabled there is a huge performance penalty.
  $ProgressPreference = 'SilentlyContinue'
  $Args__Forward = Get-Forward-Arguments -Command (Get-Command -Name Invoke-WebRequest) @Args
  Return Invoke-WebRequest @Args__Forward
  }

Function Invoke-WebRequest-ETag__Impl
  {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  Param (
    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $FileURI,
    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $FilePath__Storage,
    [String] $FileInfo__ETag
  )

  $Request = @{
    URI = $FileURI
    PassThru = $True
    OutFile = $FilePath__Storage
  }

  If ((Test-Path -Path $FilePath__Storage) -And ($FileInfo__ETag -And $FileInfo__ETag -Match '^[!#-~]+$'))
    {
  Try
    {
  $FileResponse = Invoke-WebRequest__Silent @Request `
    -Headers @{ 'If-None-Match' = ( '"{0}"' -f $FileInfo__ETag ) }
    }
  Catch [System.Net.WebException]
    {
  $FileResponse__StatusCode = $_.Exception.Response.StatusCode
  If ($FileResponse__StatusCode -Eq [System.Net.HttpStatusCode]::NotModified)
    {
    Write-Debug ('Request "{0}" Entity-Tag "{1}" matched.' -f `
  $FileURI, $FileInfo__ETag)
    Return $Null
    }
  Else
    {
    Throw $_.Exception
    }
    }
    }
  Else {
  $FileResponse = Invoke-WebRequest__Silent @Request
    }

  If ($FileResponse)
    {
  $FileInfo = [PSCustomObject] @{
    ModifiedTime = $FileResponse | Filter-WebResponseHeader__Last-Modified;
    ETag = $FileResponse | Filter-WebResponseHeader__ETag;
    FileURI = $FileURI;
  }
    }
  Else
    {
  $FileInfo = $Null
  Write-Debug ('Request "{0}" failed miserably.' -f `
    $FileURI)
    }

  Return $FileInfo
  }

Function Invoke-WebRequest-ETag
  {
  [CmdletBinding()]
  [OutPutType([Void])]
  Param (
    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $FilePath,
      [ValidateNotNullOrEmpty()]
    [String] $FilePath__Storage = ( `
      Join-Path -Path ("./{0}" -f $Content) -ChildPath ( `
        Split-Path -Path $FilePath -Leaf)),

    [String] $FileInfo__Storage = ( `
      Join-Path -Path ("./{0}" -f $ContentCache) -ChildPath ( `
        Split-Path -Path $FilePath -Leaf)) + ".json",

    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $FileURI
  )

  @( $FilePath__Storage, $FileInfo__Storage ) | %{
    New-ParentDirectoryItem $_
    }

  $FileInfo__L = Get-ItemPropertyCache -FileInfo__Storage $FileInfo__Storage

  If (($FileInfo__L | Has-Property -Name 'ModifiedTime') -And (Test-Path -Path $FilePath__Storage))
    {
    $FileInfo__F__ModifiedTime = (Get-ItemPropertyValue -Path $FilePath__Storage -Name 'LastWriteTime').ToUniversalTime()
# Note: It is possible that there are multiple content distribution servers and
# file timestamps are not identical between these.  Use a small deadband.
    $FileInfo__F__ModifiedTime__Deadband = New-TimeSpan -Minutes 30
    $FileInfo__L__ModifiedTime = $FileInfo__L | Get-Property -Name 'ModifiedTime'
    If (-Not $IgnoreLocalTimestamp -And (($FileInfo__F__ModifiedTime - $FileInfo__F__ModifiedTime__Deadband) -Gt $FileInfo__L__ModifiedTime))
      {
      Write-Output ('Ignored "{0}": Local timestamp newer "{1}" than remote "{2}".  Use -IgnoreLocalTimestamp to override.' -f `
        $FilePath, $FileInfo__F__ModifiedTime, $FileInfo__L__ModifiedTime)
      Return
      }
    }

  $FileInfo__R = Invoke-WebRequest-ETag__Impl `
    -FileURI $FileURI `
    -FilePath__Storage $FilePath__Storage `
    -FileInfo__ETag ($FileInfo__L | Get-Property -Name 'ETag')
  If (-Not $FileInfo__R)
    {
  Return
    }

  If (($FileInfo__R | Has-Property -Name 'ModifiedTime') -And (Test-Path -Path $FilePath__Storage))
    {
  Set-ItemProperty -Path $FilePath__Storage -Name 'LastWriteTime' -Value ($FileInfo__R | Get-Property -Name 'ModifiedTime')
    }

  $FileInfo__R | Get-ItemPropertyCache__NotNull | ConvertTo-Json -Compress | Set-Content -Path $FileInfo__Storage -Force

  Return
  }

Function Get-AbsoluteURI
  {
  [OutputType([String])]
  Param (
    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $URIString__Host,

    [Parameter(Mandatory)]
      [ValidateNotNullOrEmpty()]
    [String] $URIString__Path
  )

  Return [Uri]::new([Uri]::new($URIString__Host), $URIString__Path).AbsoluteUri
  }

Function Get-SysinternalsSuite-List
  {
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  Param (
    [String[]] $FileName__IgnorePattern = @(
      'about_this_site.txt',
      'ctrl2cap.*.sys',
      'dmon.sys',
      'portmon.cnt',
      '*.html'
    )
  )

  $Live = Invoke-WebRequest__Silent -UseBasicParsing -URI $LiveURI
  $FileList = @(
    )

  ForEach ($Link in $Live.Links)
    {

  $LinkHref = $Link.Href
  If (-Not ($LinkHref -Match '^(?:[^/]*/)?(?<FileName>\w+(?:\.\w+)+)$'))
    {
    Write-Debug ('Ignored "{0}".' -f `
      $LinkHref)
    Continue
    }

  $FileName = $Matches.FileName
  If ($FileName__IgnorePattern__Matched = $FileName__IgnorePattern | Where-Object { $FileName -Like $_ })
    {
    Write-Debug ('Ignored "{0}": Explicit ignore-pattern "{1}" matched.' -f `
      $FileName, $FileName__IgnorePattern__Matched)
    Continue
    }

  $FileList += [PSCustomObject] @{
    FileURI = Get-AbsoluteUri -URIString__Host $LiveURI -URIString__Path $LinkHref;
    FileName = $FileName;
    }

    }

  return $FileList
  }

Function Update-SysinternalsSuite
  {
  [CmdletBinding()]
  [OutputType([Void])]
  Param (
    [Object[]] $FileList = (Get-SysinternalsSuite-List)
  )

  If ($Filelist)
    {
  $FileList | ForEach-Object -Begin {

    $FileList__N = 1

  } -Process {

  Write-Progress `
    -Activity ('Updating {0} of {1}' -f $FileList__N, $FileList.Count) `
    -Status ('{0} ...' -f $_.FileName) `
    -PercentComplete (($FileList__N / $FileList.Count) * 100)

  Invoke-WebRequest-ETag `
    -FileURI $_.FileURI -FilePath $_.FileName
    $FileList__N++

  } -End {

  Write-Progress `
    -Activity 'Completed' `
    -Completed

  }
    }

  Return
  }

Update-SysinternalsSuite
