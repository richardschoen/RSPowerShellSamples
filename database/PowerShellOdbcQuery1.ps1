##-------------------------------------------------------------------------------
## Desc: Read database using selected ODBC driver, data source and SQL query.
##       Output data to delimited file from query results.
##
## Parameters:
## $Delimiter - Delimited record delimiter. Default = |
## $RemoveDblQuote - Remove double quotes around data which is PowerShell default. Y=Remove dbl quotes, N=Do not remove double quotes.
## $OutputFile - Output file. Required
##
## $ExitType - ENVIRONMENT=Use Environment.Exit. Only use when calling from DOS or ShellExec. Otherwise it will kill the 
## calling application. EXIT=Use standard PowerShell exit with appropriate exit code.            
##
## $ConnectionString - ODBC connection string. Uses System.Data.ODBC.ODBCConnection. Required.
## Visit: http://www.connectionstrings.com if you need to find a specific ODBC connection string sample.
##        Our example defaults to using the IBMi Access ODBC driver.
##
## $Sql - SQL selection string - Required.
##
## Returns:
##
## Article References:
## https://weblogs.asp.net/soever/returning-an-exit-code-from-a-powershell-script
## https://blogs.technet.microsoft.com/heyscriptingguy/2011/11/02/remove-unwanted-quotation-marks-from-csv-files-by-using-powershell/
## http://maxtblog.com/2011/09/creating-your-own-exitcode-in-powershell-and-use-it-in-ssis-package/
##-------------------------------------------------------------------------------

## Defined parameters
param(
[string]$Delimiter="COMMA",
[string]$RemoveDblQuote="Y",
[string]$RemoveFirstLine="N",
[string]$OutputFile="c:\temp\outputdelim.txt",
[string]$ExitType="EXIT",
[string]$ConnectionString="Driver={Client Access ODBC Driver (32-bit)};System=1.1.1.1;Uid=USER1;Pwd=PASS1;",
[string]$Sql="select * from qiws.qcustcdt"
)

##-------------------------------------------------------------------------------
## Let's try to do our work now and nicely handle errors
## This script should always end normally with an appropriate exit code
##-------------------------------------------------------------------------------
try {

## Init initial work variables
$exitcode=0

## Bail out if no output file passed
if ($OutputFile.Trim() -eq "") {
    throw "No output file specified."
}

## Replace special characters in SQL statement
##$query = $Sql.Replace("&sqt;","'") 
##$query = $query.Replace("&pct;","%") 

## Display parameters passsed
Write-Output "Exporting Data to Delimited File"
Write-Output "Delimiter: $Delimiter"
Write-Output "RemoveDblQuote: $RemoveDblQuote"
Write-Output "RemoveFirstLine: $RemoveFirstLine"
Write-Output "OutputFile: $OutputFile"
Write-Output "ExitType: $ExitType"
Write-Output "ConnectionString: $ConnectionString"
Write-Output "Sql: $Sql"

## Create database connection and open it
$database = ""
$connection = New-Object System.Data.ODBC.ODBCConnection
$connection.ConnectionString = $ConnectionString
$connection.Open()

## Create database command using SQL query 
$command = $connection.CreateCommand()
$command.CommandText = $Sql

## Execute the query to a Data Reader object (forward only)
$result = $command.ExecuteReader()

## Create Data Table object and load it from the Data Reader.
$table = new-object "System.Data.DataTable"
$table.Load($result)

## Set appropriate delimiter
if ($Delimiter.ToUpper() -eq "PIPE") { 
   $Delimiter = "|"
} 
elseif ($ExitType.ToUpper() -eq "COMMA") { 
   $Delimiter = ","
}
else { ## Default to comma
   $Delimiter = ","
}   

# Export the Data Table to delimited output file
$table | Export-Csv -NoTypeInformation -Path "$OutputFile" -Delimiter "$Delimiter"

## Close the database connection. We're done.
$connection.Close()

##Remove double quotes from delimited file data if switch passed
if ($RemoveDblQuote.Trim().ToUpper()  -eq "Y") {
  (Get-Content $outputfile) | % {$_ -replace '"', ""} | out-file -FilePath $outputfile -Force -Encoding ascii
}

##Remove first line from delimited file data if switch passed
if ($RemoveFirstLine.Trim().ToUpper()  -eq "Y") {
  (Get-Content $outputfile | Select-Object -Skip 1) | Set-Content $outputfile -Force -Encoding ascii
}

Write-Output "ExitCode: $exitcode"
Write-Output ("Message: Database export to file " +  $outputfile + "  was successful.")
Write-Output "StackTrace:"
exit $exitcode

## Causes caller to exit. Only when you need to send a DOS return code to caller
if ($ExitType.ToUpper() -eq "ENVIRONMENT") { 
   [Environment]::Exit($exitcode)
} ## Causes standard powershell exit 
elseif ($ExitType.ToUpper() -eq "EXIT") { 
   exit $exitcode
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
	Write-Output ("Message:" + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString())
    Write-Output "StackTrace: $_.Exception.StackTrace" 

	## Causes caller to exit. Only when you need to send a DOS return code to caller
	if ($ExitType.ToUpper() -eq "ENVIRONMENT") { 
	   [Environment]::Exit($exitcode)
    } ## Causes standard powershell exit 
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



