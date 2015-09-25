#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$filename = "salesforce_for_outlook.zip"
$global:SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
$global:DESTINATION = $ENV:TMP + "\" + $filename
    $global:UNZIPTO=$env:tmp + "\salesforce_for_outlook"

    remove-Item $UNZIPTO -recurse -ErrorAction SilentlyContinue
    start-sleep 5
    New-Item $UNZIPTO -Type Directory -Erroraction Silentlycontinue



function do_download($src, $dst , $zipto , $dozip)
    {
       IF(TEST-PATH  $dst){Remove-Item $dst}

        $DOWNLOAD = New-Object System.Net.WebClient
        $DOWNLOAD= $DOWNLOAD.DownloadFile($src,$dst)
        
        #EXTRACT THE FILES
        if($dozip -eq "yes")
        {
            $shell = new-object -com shell.application
            $zip = $shell.NameSpace(“$dst”)
            foreach($item in $zip.items())
            {
            $shell.Namespace(“$zipto”).copyhere($item, 0x14)
            }
        }


    }
    
function exe_paths()
    {
        
        if(!(Test-Path "C:\Program Files (x86)"))
        {
            $global:exe  = $env:tmp + "\salesforce_for_outlook\install\setup.msi"
            $global:pia  = $env:tmp + "\salesforce_for_outlook\install\o2010pia.msi"
            $global:dotnet = $env:tmp + "\salesforce_for_outlook\install\dotNetFx40_Full_x86_x64.exe"
             $global:vc = $env:tmp + "\salesforce_for_outlook\install\vcredist_x86.exe"
             $global:vstore = $env:tmp + "\salesforce_for_outlook\install\vstor_redist.exe"
            
          }
          else
         {
             $global:exe = $env:tmp + "\salesforce_for_outlook\install\setup.x64.msi"
             $global:pia  = $env:tmp + "\salesforce_for_outlook\install\o2010pia.msi"
             $global:dotnet = $env:tmp + "\salesforce_for_outlook\install\dotNetFx40_Full_x86_x64.exe"
              $global:vc = $env:tmp + "\salesforce_for_outlook\install\vcredist_x86.exe"
              $global:vstore = $env:tmp + "\salesforce_for_outlook\install\vstor_redist.exe"
          }
          
    }




function install_salesforce()
{
   
    
     $commandvs = "$vstore /q /norestart"
    cmd /c $commandvs
    
    $commandvc = "$vc /install /quiet /norestart"
    cmd /c $commandvc
    
    $commanddotnet = "$dotnet /q /norestart"
    cmd /c $commanddotnet
    
    $commandpia = "$pia /quiet /norestart"
    cmd /c $commandpia
    
     cmd /c  taskkill /im outlook.exe 2>&1 > $null
    start-sleep 10
    
    $commandexe = "$exe /quiet /norestart "
    cmd /c $command

}

do_download $SOURCE $DESTINATION $UNZIPTO "YES"
exe_paths
install_salesforce
