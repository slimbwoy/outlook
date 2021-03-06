clear
#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$filename = "salesforce_for_outlook.exe"
$global:SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
$global:DESTINATION = "C:\WINDOWS\TEMP\" + $filename
    $global:UNZIPTO= "C:\WINDOWS\TEMP\salesforce_for_outlook.exe"

    remove-Item $UNZIPTO -recurse -ErrorAction SilentlyContinue
    start-sleep 5
    #New-Item $UNZIPTO -Type Directory -Erroraction Silentlycontinue



function do_download($src, $dst)
    {
       IF(TEST-PATH  $dst){Remove-Item $dst}

        $DOWNLOAD = New-Object System.Net.WebClient
        $DOWNLOAD= $DOWNLOAD.DownloadFile($src,$dst)
        
        #EXTRACT THE FILES
       $global:CMDUNZIP = "C:\WINDOWS\TEMP\UNZIP.CMD"
       echo "$UNZIPTO" | out-file $CMDUNZIP -Encoding ASCII
       start-process cmd -Argumentlist "/c $CMDUNZIP" -workingdirectory C:\windows\temp -nonewwindow -wait
     


    }
       
function exe_paths()
    {
        
        if(!(Test-Path "C:\Program Files (x86)"))
        {
            $global:exe  = "C:\WINDOWS\TEMP\salesforce_for_outlook\setup.msi"
            $global:pia  = "C:\WINDOWS\TEMP\salesforce_for_outlook\o2010pia.msi"
            $global:dotnet = "C:\WINDOWS\TEMP\salesforce_for_outlook\dotNetFx40_Full_x86_x64.exe"
             $global:vc = "C:\WINDOWS\TEMP\salesforce_for_outlook\vcredist_x86.exe"
             $global:vstore = "C:\WINDOWS\TEMP\salesforce_for_outlook\vstor_redist.exe"
            
          }
          else
         {
             $global:exe = "C:\WINDOWS\TEMP\salesforce_for_outlook\setup.x64.msi"
             $global:pia  = "C:\WINDOWS\TEMP\salesforce_for_outlook\o2010pia.msi"
             $global:dotnet = "C:\WINDOWS\TEMP\salesforce_for_outlook\dotNetFx40_Full_x86_x64.exe"
              $global:vc = "C:\WINDOWS\TEMP\salesforce_for_outlook\vcredist_x86.exe"
              $global:vstore = "C:\WINDOWS\TEMP\salesforce_for_outlook\vstor_redist.exe"
          }
          
    }




function install_salesforce()
{
   
   
    #

    #write-host $user
        #------------CRYPTOGRAPHY STUFF-----######
  $enccode = "76492d1116743f0423413b16050a5345MgB8AFkASwBGADYANAB4AHkAdwB5AEkAYgBsAFIAVwA2AEYAegAvAFIATABvAHcAPQA9AHwAMQBiADcANQA5ADEAMABiAGYAYwA5ADQAYwA4ADcANgBhADgANwBlAGMANAAwAGMAMQA3ADcAOQBhADIAZQBiAA=="

        $key = (29,198,19,130,110,209,124,187,56,144,7,99,79,240,71,85)
        $username = "$hostname\administrator"
        $EncryptedPW = $enccode
        $SecureString = ConvertTo-SecureString -String $EncryptedPW -Key $Key
        $loccred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$SecureString
        #------------END CRYPTO--------------#####
     
               
          # Start-Process powershell -Credential $loccred -ArgumentList '-noprofile -command &{Start-Process C:\windows\temp\salesforce_for_outlook\do_install.cmd -verb runas }'
          cmd /c C:\windows\temp\salesforce_for_outlook\do_install.cmd
          
}

#do_download $SOURCE $DESTINATION
exe_paths
install_salesforce
