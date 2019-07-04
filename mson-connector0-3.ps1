<#
#### requires ps-version 3.0 ####
<#
.SYNOPSIS
Autoconnect with credential files to Microsoft Online Services 
.DESCRIPTION
Uses credential files and encryption files to connect to O365 services like AzureAD,Exchange Online,MSonline 
.PARAMETER CredFilePath
path to the credential file (securestring encrypted with aes file)
.PARAMETER KeyFilePath
path to the encryption file (aes file used to encrypt credential file)
.PARAMETER ConnectionTypes 
specifies the microsoft service you want to connect only "ExOn","AzureAD","MSOL","MSTeams","MSSkype" is allowed
.PARAMETER Username
specifies the user you want to connect to O365 
.INPUTS
credential file and keyfile 
.OUTPUTS
none
.NOTES
   Version:        0.3
   Author:         Alexander Koehler
   Creation Date:  Monday, April 29th 2019, 7:26:03 pm
   File: mson-connector.ps1
   Copyright (c) 2019 blog.it-koehler.com
HISTORY:
Date      	          By	Comments
----------	          ---	----------------------------------------------------------
2019-06-23-11-37-am	 AK	    adding service skype, check module import, get commands after connections
2019-06-06-06-52-pm	 AK	    online service teams added
2019-06-06-06-42-pm	 AK	    comments added
2019-05-08-06-53-pm	 AK	    adding check for modules

.LINK
   https://blog.it-koehler.com/en/

.COMPONENT
 Required Modules: AzureAD (Install-Module AzureAD), MSOnline (Install-Module MSOnline), MSTeams (Install-Module Microsoft Teams), Skype Powershell
 see also https://docs.microsoft.com/de-de/office365/enterprise/powershell/connect-to-office-365-powershell
 For Skype there are several requirements, see https://docs.microsoft.com/en-us/office365/enterprise/powershell/manage-skype-for-business-online-with-office-365-powershell


.LICENSE
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the Software), to deal
in the Software without restriction, including without limitation the rights
to use copy, modify, merge, publish, distribute sublicense and /or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 
.EXAMPLE
.\_mson-connector-0-3.ps1 -CredFilePath .\cred.txt -KeyfilePath .\keyfile.key -ConnectionTypes MSOL -Username adminuserO365@domain.com
#

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#>


[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        
            if(($_ | Test-path -PathType Leaf)){
                Write-Verbose "Checking if file exists: $encryptfile and it is a file"
            }
            else {
                throw "File does not exist in path: $CredFilePath, please provide another path and provide the filename example C:\temp\aes.key!"
            }
            if(($_ | Test-Path -PathType Container)){
                throw "The path argument has to be a file. Folder paths are not allowed."
            }
            else {
                
                Write-Verbose "Checking if its not a folder: $encryptfile"
            }
            return $true

    })]
    [System.IO.FileInfo]$CredFilePath,    
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if(($_ | Test-path -PathType Leaf)){
            Write-Verbose "Checking if file exists: $KeyfilePath and it is a file"
        }
        else {
            throw "File does not exist in path: $KeyfilePath, please provide another path and provide the filename example C:\temp\aes.key!"
        }
        if(($_ | Test-Path -PathType Container)){
            throw "The path argument has to be a file. Folder paths are not allowed."
        }
        else {
            
            Write-Verbose "Checking if its not a folder: $KeyfilePath"
        }
        return $true
    })]
    [System.IO.FileInfo]$KeyfilePath,
    [Parameter(Mandatory = $true)]
    [ValidateSet("ExOn","AzureAD","MSOL","MSTeams","MSSkype")]
    [String]$ConnectionTypes,
    [Parameter(Mandatory = $true)]
    [String]$Username
    )
#---------------------------------------------------------[Functions]--------------------------------------------------------
#get password from files to secure string 
function Get-AESEncryptCreds {
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory= $true)]
        [string]
        $credfile,
        [Parameter(Mandatory= $true)]
        [string]
        $encryptfile
    )
    Write-Verbose "Getting content from file: $encryptfile"
    $key = (Get-Content "$encryptfile")
    Write-Verbose "Encypting $credfile with encryption file $encryptfile to secure string."  
    #converting password to secure string
    $script:password= (Get-Content "$credfile" | ConvertTo-SecureString -Key $key)
    #see also https://techibee.com/powershell/convert-system-security-securestring-to-plain-text-using-powershell/2599
}
#checking module if its installed
function Get-InstalledModule {
    [CmdletBinding()]
    param(
      [Parameter(Mandatory = $true)]
      [string]$modulename
    )
    
    Write-Verbose "Checking if module $modulename is installed correctly"
    if (Get-Module -ListAvailable -Name $modulename) {
      $Script:moduleavailable = $true
      Write-Verbose "Module $modulename found successfully!"
    } 
    else {
      Write-Verbose "Module $modulename not found!"
      throw "Module $modulename is not installed or does not exist, please install and retry.
      In an administrative Powershell console try typing: Install-Module $modulename"
    }
  
  }
#checking if module is imported, if not load it
  function Get-ImportedModule {
    [CmdletBinding()]
    param(
      [Parameter(Mandatory = $true)]
      [string]$modulename
    )
         #check if module is imported, otherwise try to import it
         if (Get-Module -Name $modulename) {
            Write-Verbose "Module $modulename already loaded"
            Write-Verbose "Getting cmdlets from module"
            #write output to variable to get all cmdlets
            $global:commands = Get-Command -Module $modulename | Format-Table -AutoSize -Wrap
            Write-Verbose "Cmdlets stored in variable commands"

        }
        else {
            Write-Verbose "Module found but not imported, import starting"
            Import-Module $modulename -force
            Write-Verbose "Module $modulename loaded successfully"
            #write output to variable to get all cmdlets
            Write-Verbose "Getting cmdlets from module"
            $global:commands = Get-Command -Module $modulename | Format-Table -AutoSize -Wrap
            Write-Verbose "Cmdlets stored in variable commands"
           
        }
  }
#connection to exchange server online
function Connect-ServiceExOn {
[CmdletBinding()]
    param (
        [Parameter(Mandatory= $true)]
        [string]$user,
        [Parameter(Mandatory= $true)]
        [securestring]$pw)
        $credential = New-Object System.Management.Automation.PsCredential("$user",$pw)   
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection
        Import-PSSession $Session -DisableNameChecking -Verbose
    }
 # connection to azure ad     
    function Connect-ServiceAzureAD {
    [CmdletBinding()]
        param (
            [Parameter(Mandatory= $true)]
            [string]$user,
            [Parameter(Mandatory= $true)]
            [securestring]$pw)
            $credential = New-Object System.Management.Automation.PsCredential("$user",$pw)   
            Connect-AzureAD -Credential $credential -Verbose
    } 
#connection to office 365 powershell
     function Connect-ServiceMSOnline {
        [CmdletBinding()]
            param (
                [Parameter(Mandatory= $true)]
                [string]$user,
                [Parameter(Mandatory= $true)]
                [securestring]$pw)
                $credential = New-Object System.Management.Automation.PsCredential("$user",$pw)   
                Connect-MsolService -Credential $credential -Verbose
        }  
#connection to ms teams 
     function Connect-ServiceMSTeams {
        [CmdletBinding()]
            param (
                [Parameter(Mandatory= $true)]
                [string]$user,
                [Parameter(Mandatory= $true)]
                [securestring]$pw)
                Write-Verbose "Connecting to Microsoft Teams with Connect-MicrosoftTeams cmdlet"
                $credential = New-Object System.Management.Automation.PsCredential("$user",$pw)   
                Write-Verbose "Using credential $credential"
                Connect-MicrosoftTeams -Credential $credential -Verbose
        }  

    #connection to skype business 
     function Connect-ServiceSkype {
        [CmdletBinding()]
            param (
                [Parameter(Mandatory= $true)]
                [string]$user,
                [Parameter(Mandatory= $true)]
                [securestring]$pw)
                Write-Verbose "Connection to Skype Business Online, if you can not connect please verify you've all requirements
                https://docs.microsoft.com/en-us/office365/enterprise/powershell/manage-skype-for-business-online-with-office-365-powershell"
                $credential = New-Object System.Management.Automation.PsCredential("$user",$pw)   
                $sfbSession = New-CsOnlineSession -Credential $credential
                Import-PSSession $sfbSession
        }  

# convert password to secure string (call function)  
Get-AESEncryptCreds -credfile $CredFilePath -encryptfile $KeyfilePath
#find connection type and call the right function
switch -Regex ($connectiontypes) {
            #exchange online connection
            "ExOn" {Connect-ServiceExOn -user $username -pw $password}
            #azure ad connection 
            "AzureAD" {
                #check if the module is installed
                $modulename = "AzureAD"
                Get-InstalledModule -modulename $modulename
                #check if the module is import, if not it 
                Get-ImportedModule -modulename $modulename
                #connect to azure ad
                Connect-ServiceAzureAD -user $username -pw $password
                Write-Output " "
                Write-Output  "To see all cmdlets from module $modulename type in "'$commands'""
                }
            "MSOL"{
                #check if the module is installed
                $modulename = "MSOnline"
                Get-InstalledModule -modulename $modulename
                #check if the module is import, if not it 
                Get-ImportedModule -modulename $modulename
                 #connect to office 365 powershell
                Connect-ServiceMSOnline -user $username -pw $password
                Write-Output " "
                Write-Output  "To see all cmdlets from module $modulename type in "'$commands'""
            }
            "MSTeams"{
                #check if the module is installed
                $modulename = "MicrosoftTeams"
                Get-InstalledModule -modulename $modulename
                #check if the module is import, if not it 
                Get-ImportedModule -modulename $modulename
                #connect to teams powershell
                Connect-ServiceMSTeams -user $username -pw $password
                Write-Output " "
                Write-Output  "To see all cmdlets from module $modulename type in "'$commands'""
            }
            "MSSkype"{
                #check if the module is installed
                $modulename = "SkypeOnlineConnector"
                Get-InstalledModule -modulename $modulename
                #check if the module is import, if not it 
                Get-ImportedModule -modulename $modulename
                Connect-ServiceSkype -user $username -pw $password
               
            }
        }




        
