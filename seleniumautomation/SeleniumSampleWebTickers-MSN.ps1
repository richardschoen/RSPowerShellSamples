##-------------------------------------------------------------------------------
## Desc: Read list of stock ticker symbols from a CSV file and go to MSN Money 
##       web site and get the current stock price for each ticker and write back
##       to a delimited output file so it can be consumed by the calling program/task.
##
## Parameters:
## $inputfile - Stock ticker CSV file
## $outputfile - Pipe delimited output file. (Use pipe since amount could have commas.)
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
[string]$inputfilename="C:\temp\AMSeleniumPs\StockPrices.csv",
[string]$outputfilename="c:\temp\tickeroutput.txt",
[string]$ExitType="ENVIRONMENT"
)

## Include the Selenium PowerShell module
Import-Module (Join-Path $PSScriptRoot "Selenium.psm1")

## Initialize variables
$curelement = ""
$Driver=$null
$curticker = ""      
$curexchange=""
$exitcode=0

##-------------------------------------------------------------------------------
## Let's try to do our work now and nicely handle errors
## This script should always end normally with an appropriate exit code
##-------------------------------------------------------------------------------
try {

        ## Read ticker sumbols from CSV
        $csvtickers = Import-Csv $inputfilename
        
        ## Output CSV Header record to new output file. This replaces the file 
        Out-File -FilePath $outputfilename -InputObject "Symbol|Exchange|LastPrice|StartTime|EndTime" -Encoding ASCII

        ## Launch new Chrome session
        $Driver = Start-SeChrome

        ## Navigat to MSN money site
        Enter-SeUrl "https://www.msn.com/en-us/money/markets" -Driver $Driver

        ## Process each ticker symbol and get current price from MSN site
        ForEach ($csvticker in $csvtickers){

                ## Capture transaction start time in hh:mm:ss format
                $date = Get-Date 
                $starttime = '{0:hh:mm:ss}' -f $Date
                ##$starttime = '{0:MM/dd/yyyy hh:mm:ss}' -f $Date

                ## Navigate to selected ticker symbol
                Enter-SeUrl "https://www.msn.com/en-us/money/stockdetails/fi-126.1.$($csvticker.Symbol).$($csvticker.Exchange)?symbol=$($csvticker.Symbol)&form=PRFIMQ" -Driver $Driver

                ## Find current price HTML element by its CSS class name
                $Element = Find-SeElement -Driver $Driver -ClassName "current-price"

                ## Bail out if no element found
                If ($Element -eq $null) {
                    throw "Unable to find element"
                }				

                ## Capture transaction end time in hh:mm:ss format
                $date = Get-Date 
                $endtime = '{0:hh:mm:ss}' -f $Date
                ##$endtime = '{0:MM/dd/yyyy hh:mm:ss}' -f $Date

                ##Output the current stock symbol element to Console/STDOUT in case the caller wants to use STDOUT
                ##Note: Comment the following line out if you don't want data writted to the Console/STDOUT
                Write-Output "$($csvticker.Symbol)|$($csvticker.Exchange)|$($Element.Text)!$($starttime)|$($endtime)"
                
                ##Output/append the current stock symbol delimited record delimited output file 
                Out-File -Append -FilePath $outputfilename -InputObject "$($csvticker.Symbol)|$($csvticker.Exchange)|$($Element.Text)|$($starttime)|$($endtime)" -Encoding ASCII

        }

        ## Stop the Selenium Web Driver. We are done with current browser session.
        If ($Driver -ne $null) {
           Stop-SeDriver $Driver
        }

        ## Exit now - success
        Write-Output "ExitCode: $exitcode"
        Write-Output ("Browser automation completed successfully.")
        Write-Output "StackTrace:"

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

    ## Stop the Selenium Web Driver which should also close browser if it is loaded
    If ($Driver -ne $null) {
       Stop-SeDriver -Driver $Driver 
    }

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
