function Get-DirSize {
<#
.Synopsis
  Gets a list of directories and sizes.
.Description
  This function recursively walks the directory tree and returns the size of 
  each directory found.
.Parameter path
  The path of the root folder to start scanning.
.Example
  # Get the largest folder under the user profile
  PS> (Get-DirSize $env:userprofile | sort Size)[-2]
.ReturnValue
  An object with Folder and Size properties.
#>
  param([Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$path)
  BEGIN {}
 
  PROCESS{
    $size = 0
    $folders = @()
  
    foreach ($file in (Get-ChildItem $path -Force -ea SilentlyContinue)) {
      if ($file.PSIsContainer) {
        $subfolders = @(Get-DirSize $file.FullName)
        $size += $subfolders[-1].Size
        $folders += $subfolders
      } else {
        $size += $file.Length
      }
    }
  
    $object = New-Object -TypeName PSObject
    $object | Add-Member -MemberType NoteProperty -Name Folder `
                         -Value (Get-Item $path).FullName
    $object | Add-Member -MemberType NoteProperty -Name Size -Value $size
    $folders += $object
    Write-Output $folders
  }
  
  END {}
}