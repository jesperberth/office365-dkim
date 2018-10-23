# Menu driven DKIM tool for Office 365
# - Creates CSV file with CNAME for DNS 
# - Test DNS Settings 
# - Enables DKIM no Domains
#
# Author: Jesper Berth, Arrow ECS, jesper.berth@arrow.com - 9. October 2018
# Version 1.0.1

$createdate = get-date -Format ddMMyyyyhhmmss
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$outfile = "$DesktopPath\dkim-DNS-$createdate.csv"
$topline = "domain,hostname,cname"
#$id = "3"

function Show-Menu
{
    param (
        [string]$Title = 'Office 365 DKIM'
    )
    Clear-Host
    Write-Host "=========== $Title ============`n"
    Write-Host "1: Create DKIM CSV File"
    Write-Host "2: Test DKIM DNS"
    Write-Host "3: Enable DKIM"
    Write-Host "==============================="
    Write-Host "L: Login to Office 365"
    Write-Host "X: Logout off Office 365"
    Write-Host "Q: Press 'Q' to quit."
    Write-Host "==============================="
}
function Login{
    
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking
    #Connect-azuread -Credential $UserCredential
    #Connect-MsolService -Credential $UserCredential
    Clear-Host
    write-host -ForegroundColor Green "Logon granted.."
}

function Logout{
    write-host "Logout Office 365 Powershell"
    $session = get-pssession
    $id = $session.id
    Remove-PSSession -id $id
}
function Add-DKIM-DNS-CSV{

$dkim = Get-DkimSigningConfig 

Write-Output $topline >> $outfile

    For ($i=0; $i  -lt $dkim.length; $i++){
    $domain = $dkim[$i].domain
    $cname1 = $dkim[$i].Selector1Cname
    $cname2 = $dkim[$i].Selector2Cname
    $enabled = $dkim[$i].enabled
    $hostname1 = "selector1._domainkey"
    $hostname2 = "selector2._domainkey"

        if ($enabled -eq $false){
        Write-host  "Adding domain: $domain" 
           
        Write-Output "$domain,$hostname1,$cname1" >> $outfile
    
        Write-Output "$domain,$hostname2,$cname2" >> $outfile 
        }
    }
}## Add-DKIM-DNS-CSV

function Test-DKIM-DNS{

    $dkim = Get-DkimSigningConfig

    For ($i=0; $i  -lt $dkim.length; $i++){
    
            $domain = $dkim[$i].domain
            $cname1 = $dkim[$i].Selector1Cname
            $cname2 = $dkim[$i].Selector2Cname
            $hostname1 = "selector1._domainkey."+"$domain"
            $hostname2 = "selector2._domainkey."+"$domain"

        if($domain.Contains("onmicrosoft.com") -eq $false ){
            write-host "Resolving domain: $domain"
                try
                {
                $name1 = Resolve-DnsName -Type CNAME $hostname1 -ErrorAction SilentlyContinue
                }
                Catch
                {
                write-host -ForegroundColor Yellow "Could Not resolve: $hostname1"
                }
                try
                {
                $name2 = Resolve-DnsName -Type CNAME $hostname2 -ErrorAction SilentlyContinue
                }
                catch
                {
                write-host -ForegroundColor Yellow "Could Not resolve: $hostname2"
                }
            if ($name1.NameHost -eq $cname1){
                write-host -ForegroundColor Green "CNAME 1 for DKIM ok: $cname1" 
            }
            else {
                Write-Host -ForegroundColor Red "CNAME 1 for DKIM Not resolveble: $cname1"
            }
            if ($name2.NameHost -eq $cname2){
                write-host -ForegroundColor Green "CNAME 2 for DKIM ok: $cname2" 
            }
            else {
            Write-Host -ForegroundColor Red "CNAME 2 for DKIM Not resolveble: $cname2"
            }
        }# Check Domain
    }
}
function Enable-DKIM{

$dkim = Get-DkimSigningConfig 

    For ($i=0; $i  -lt $dkim.length; $i++){
    $domain = $dkim[$i].domain
    $enabled = $dkim[$i].enabled

        if ($enabled -eq $false){
        Write-host  "Enable DKIM for domain: $domain" 
        $Readhost = Read-Host " (y/n) " 
            Switch ($ReadHost) 
            { 
            Y {Write-host "Yes, Enable DKIM"; Enable-DKIM-doit $domain} 
            N {Write-Host "No, Skip DKIM"; } 
            } 
      
        }
    }
}## Enable-DKIM

function Enable-DKIM-doit{
    Param([string]$domain)

    write-host "Enabling DKIM for domain: $domain"
    Set-dkimsigningconfig -identity $domain -enabled $true
} ## Enable-DKIM-doit

do
 {
     Show-Menu
     $selection = Read-Host "Please make a selection"
     switch ($selection)
     {
         '1' {
             Clear-Host            
             Add-DKIM-DNS-CSV
         } '2' {
             Clear-Host
             Test-DKIM-DNS
         } '3' {
             Clear-Host
             Enable-DKIM
         }'L' {
             Clear-Host
             Login
         }'X' {
             Clear-Host
             Logout
      }
     }
     pause
 }
 until ($selection -eq 'q')