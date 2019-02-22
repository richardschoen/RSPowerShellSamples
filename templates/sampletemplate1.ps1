##-------------------------------------------------------------------------------
## Desc: Put description of script here
##
## Parameters:
## $parm1 - parm 1 description
## $parm2 - parm 2 description
## $ExitType - 
##   ENVIRONMENT=Use Environment.Exit. Only use when calling from DOS or ShellExec. Otherwise it will kill the 
##   calling application. This will send back the OS return code so our caller can capture that. Also STDOUT.
##   EXIT=Use standard PowerShell exit with appropriate exit code.  
##   RETURN= ?????? TODO - Define and test diff between all return styles.          
##
## Returns: ExitCode value. o=Success, 99=Errors. Ticker list written to output file.
##
## Selenium Article References:
## Selenium PowerShell: https://github.com/adamdriscoll/selenium-powershell
##
## Misc PowerShell Articles
## https://steemit.com/utopian-io/@haig/how-to-create-a-word-document-with-powershell-using-a-powershell-script-to-automate-microsoft-word
## https://weblogs.asp.net/soever/returning-an-exit-code-from-a-powershell-script
## https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/
## http://maxtblog.com/2011/09/creating-your-own-exitcode-in-powershell-and-use-it-in-ssis-package/
##
## Read CSV Line by Line
## https://social.technet.microsoft.com/Forums/windows/en-US/aa20cabb-3196-472d-9594-aaa0ec10fa03/powershell-read-csv-file-line-by-line-and-only-first-header?forum=winserverpowershell
## https://www.petri.com/powershell-import-csv-cmdlet-parse-comma-delimited-csv-text-file
##-------------------------------------------------------------------------------

## Defined parameters
param(
[string]$parm1="I am parm 1",
[string]$parm2="I am parm 1",
[string]$ExitType="ENVIRONMENT"
)

## Include any modules here if needed
##Import-Module (Join-Path $PSScriptRoot "Selenium.psm1")

## Initialize work variables
$exitcode=0

##-------------------------------------------------------------------------------
## Let's try to do our work now and nicely handle errors
## This script should always end normally with an appropriate exit code
##-------------------------------------------------------------------------------
try {

        ##Output the current stock symbol element to Console/STDOUT in case the caller wants to use STDOUT
        ##Note: Comment the following line out if you don't want data writted to the Console/STDOUT
        Write-Output "I am the output from your template script."

        ## Exit now - success
        Write-Output "ExitCode: $exitcode"
        ## Output a completion message and blank stacktrace since we succeeded.
        Write-Output ("Script completed successfully.") 
        Write-Output "StackTrace:"
        ## Exit the script
        exit $exitcode

        ## Causes caller to exit. Only when you need to send a DOS return code to caller
        if ($ExitType.ToUpper() -eq "ENVIRONMENT") { 
           [Environment]::Exit($exitcode)
        } ## Causes standard powershell exit (returns 0=success and 1=failure. OS code does not return.)
        elseif ($ExitType.ToUpper() -eq "EXIT") { 
           exit $exitcode
        }   
    	elseif ($ExitType.ToUpper() -eq "RETURN") { 
	       return
        }
        else { 
          ## Need to determine default exit method ?
        }

}
##-------------------------------------------------------------------------------
## Catch and handle any errors and return useful info via console
##-------------------------------------------------------------------------------
catch [System.Exception] {
	$exitcode=99
	Write-Output "ExitCode: $exitcode"
	Write-Output ("Message:" + $_.Exception.Message +  " Pos msg:" + $_.InvocationInfo.PositionMessage.ToString())
	##Write-Output ("Message:" + $_.Exception.Message + " ScriptName:" + $_.InvocationInfo.ScriptName.ToString() + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString() + " Pos msg:" + $_.InvocationInfo.PositionMessage.ToString())
    Write-Output "StackTrace: $_.Exception.StackTrace" 
    Write-Output $

	## Causes caller to exit. Only when you need to send a DOS return code to caller
	if ($ExitType.ToUpper() -eq "ENVIRONMENT") { 
	   [Environment]::Exit($exitcode)
    } ## Causes standard powershell exit (returns 0=success and 1=failure. OS code does not return.)
	elseif ($ExitType.ToUpper() -eq "EXIT") { 
	   exit $exitcode
    }
	elseif ($ExitType.ToUpper() -eq "RETURN") { 
	   return
    }
    else { 
      ## Need to determine default exit method ?
    }
		
    }
