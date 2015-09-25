$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$global:reg = $wd + "\setform.reg"


$outlookver = (get-itemproperty -literalpath HKLM:\SOFTWARE\Classes\Outlook.Application\CurVer).'(default)'
if($outlookver -like "*15*")
    {

        $makereg = (Get-Content $reg |  Foreach-Object {$_ -replace "14.0","15.0" } ) | set-content $reg -encoding ascii 
      

    }
    
  #IMPORT REGISTRY FILE
  cmd /c reg import $reg 2>&1 > $null  
