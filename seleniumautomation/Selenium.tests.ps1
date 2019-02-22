Import-Module (Join-Path $PSScriptRoot "Selenium.psm1")

Describe "Get-SeCookie" {
    Context "Should get cookies from google" {
        $Driver = Start-SeChrome
        Enter-SeUrl -Driver $Driver -Url "http://www.google.com"

        Get-SeCookie $Driver
    }
}