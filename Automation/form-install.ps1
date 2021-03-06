<# ----
SYNOPSIS :
    This script will install a custom user form in microsoft outlook. This form was requested by Salesforce for integration with microsoft outlook"
    
    Written by the Team at comptuer Network systems : wwww.compnetsys.com
    
  Mckin S
  Melvin F
  
  
 ----#>  
    
 

#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$filename = "graebel_outlook_form.zip"
$global:SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
$global:DESTINATION = $ENV:TMP + "\" + $filename
    $global:UNZIPTO=$env:tmp
    $global:FILEPATH=$env:tmp + "\graebel_outlook_form"
    $global:reg = $FILEPATH + "\setform.reg"

    remove-Item $FILEPATH -recurse -ErrorAction SilentlyContinue
    start-sleep 5
   # New-Item $UNZIPTO -Type Directory



function get_form()
    {
    
       WRITE-HOST "`r`nDOWNLOADING LATEST FROM ----->`r`n"
       IF(TEST-PATH  $DESTINATION){Remove-Item $DESTINATION}
         
        #Invoke-WebRequest $source -OutFile $destination

        $DOWNLOAD = New-Object System.Net.WebClient
        $DOWNLOAD= $DOWNLOAD.DownloadFile($SOURCE,$DESTINATION)
        
        #EXTRACT THE FILES
        
        $shell = new-object -com shell.application
        $zip = $shell.NameSpace(“$DESTINATION”)
        foreach($item in $zip.items())
        {
        $shell.Namespace(“$UNZIPTO”).copyhere($item, 0x14)
        }


        WRITE-HOST "`tDOWNLOAD COMPLETED"
    }


    function install_form()
    
    {
        WRITE-HOST "`r`nINSTALLING FORM ----->"
        
        #DETERMING VERSION OF OUTLOOK INSTALLED
        $outlookver = (get-itemproperty -literalpath HKLM:\SOFTWARE\Classes\Outlook.Application\CurVer).'(default)'
        if($outlookver -like "*15*")
         {

              $makereg = (Get-Content $reg |  Foreach-Object {$_ -replace "14.0","15.0" } ) | set-content $reg -encoding ascii 
              

          }
            
        #IMPORT REGISTRY FILE
        cmd /c reg import $reg 2>&1 > $null  
          

        $oftfile= $FILEPATH + "\Appointments.oft"
        $vbfile = $FILEPATH + "\install_form.vbs"
        $setformlocation = (get-content $vbfile | foreach-object{$_ -replace "formlocation","$oftfile"} ) | set-content $vbfile -encoding ascii

        #IMPORT FORM INTO OUTLOOK
        cmd /c c:\windows\system32\cscript.exe $vbfile
        
        WRITE-HOST "`tFORM INSTALL COMPLETED. THIS WINDOW WILL CLOSE IN 20 SECONDS`r`n"  
     }
        
        
        
 clear
   
 get_form
 install_form
 
 start-sleep 20
 EXIT
 
 
 