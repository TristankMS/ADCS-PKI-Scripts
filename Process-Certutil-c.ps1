<#
Disclaimer
The sample scripts are not supported under any Microsoft standard support 
program or service. 
The sample scripts are provided AS IS without warranty of any kind. Microsoft
further disclaims all implied warranties including, without limitation, any 
implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample scripts and 
documentation remains with you. In no event shall Microsoft, its authors, or 
anyone else involved in the creation, production, or delivery of the scripts be
liable for any damages whatsoever (including, without limitation, damages for 
loss of business profits, business interruption, loss of business information, 
or other pecuniary loss) arising out of the use of or inability to use the 
sample scripts or documentation, even if Microsoft has been advised of the 
possibility of such damages.
 -----------------------
 Process-Certutil-C.ps1
-----------------------
# Was originally Process-Certutil-B for Butchered, but now it's newer
# so it works with ADCSA 5.5+ Collectors and paged result sets
# Original by Russ Tomkins, butchered by TristanK

  .SYNOPSIS
  Imports the output of a CA database query using "certutil.exe -view"" and converts it into a PowerShell XML object or a CSV file.

  .DESCRIPTION
  Allows an administrator to work with a Powershell Object or CSV version of the Certutil CA Database view command.
  
  An example command that generates the output for all issued certificates that will expire in the next 90 days is
  certutil.exe -view -restrict "NotAfter<=now+90:00,Disposition=20"

  Imports the certutil.exe dump output and outputs the contents to xml.
  Process-CertUtil.ps1 -InputFile .\CertUtilExport.txt -OutputFile .\CertutilExport.xml
  
  Imports the certutil.exe dump output and outputs those that expire in the next 60 days to xml.
  Process-CertUtil.ps1 -InputFile .\CertUtilExport.txt -OutputFile .\CertutilExport.xml -DaysTilExpiry 60
  
  Imports the certutil.exe dump output and outputs those that expire in the next 30 days to a CSV file
  Process-CertUtil.ps1 -InputFile .\CertUtilExport.txt -ExportFile .\CertutilExport.CSV -DaysTilExpiry 30

  .EXAMPLE
  Imports the certutil.exe dump output and outputs the contents to xml.
  Process-CertUtil.ps1 -InputFile .\CertUtilExport.txt -OutputFile .\CertutilExport.xml
  .EXAMPLE
  Imports the certutil.exe dump output and outputs those that expire in the next 90 days to xml.
  Process-CertUtil.ps1 -InputFile .\CertUtilExport.txt -OutputFile .\CertutilExport.xml -DaysTilExpiry 90
  .EXAMPLE
  Imports the certutil.exe dump output and outputs those that expire in the next 30 days to a CSV file
  Process-CertUtil.ps1 -InputFile .\CertUtilExport.txt -ExportFile .\CertutilExport.CSV -DaysTilExpiry 90
  
  .PARAMETER InputFile
  The certutil.exe -dump output file file to be processed. This file will be record of the stdout ">" the certutil.exe command executed against a CA database.
  .PARAMETER OutputFile
  The resulting processed powershell object as an XML object
  .PARAMETER ExpiresIn
  The resulting processed powershell object
  .PARAMETER ExportFile
  The CSV export of the PowerShell object 
  .PARAMETER CAName
  Optional name for the CA Issuer
  #>

[CmdletBinding()]
    Param (
	# Input File
    [Parameter(Mandatory=$true,ValueFromPipeline=$True,Position=0)]
	[ValidateNotNullOrEmpty()]
	[String]$InputFile,
	
    # Output File
    [Parameter(Mandatory=$True,Position=1,ParameterSetName = "Output")]
	[string]$OutputFile,

   	# Export File
    [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,Position=1,ParameterSetName = "Export")]
	[ValidateNotNullOrEmpty()]
	[String]$ExportFile,

	# DaysTilExpiry
    [parameter(Mandatory=$false,Position=2,ParameterSetName = "Output")]
    [Parameter(Mandatory=$false,Position=2,ParameterSetName = "Export")]
    [Int]$DaysTilExpiry,
    
    [Parameter(Mandatory=$false)]
	[String]$CAName,

    [Parameter(Mandatory=$false)]
	[Bool]$Append = $false
    )

# Preparation
$Now = Get-Date
$Rows = @()
$InputFile = (Get-Item $InputFile).FullName
Write-Host "Reading Input File " $InputFile
Switch($PSCmdlet.ParameterSetName)
    {
		"Export"    
        {
            if($Append -ne $True){
                if(Test-Path $ExportFile){
                    Remove-Item $ExportFile -Force
                }
            }
        }
		"Output"
		{
			if(Test-Path $OutputFile){
				Remove-Item $OutputFile -Force
			}
		}

    }

$rowfakeID=1

# Loop through the input file, one line at a time.
ForEach ($Line in [System.IO.File]::ReadLines("$InputFile")) {
    
    # Last line of ADCS Collector files... cos the Certutil command line might simply be an error or a midpoint
    If($Line -Match "_ADCS_ROW_COUNT:"){Break}
    
    # Look for the word Row in othe output
    If($Line -Match "Row \d"){
        #If we have a RowID populated on a Row custom object, finalise the custom object. Don't add it to our custom object though if a it's greater the the days til expiry Value
        If($Row.RowID -ne $Null){

           If($DaysTilExpiry){                
                If($Row.DaysTilExpiry -le $DaysTilExpiry -and $Row.DaysTilExpiry -gt 0){
                        $Rows += $Row
                }
            }
            Else{$Rows += $Row}
        }
        
        
        # Create a new row for the next record
        $Row = "" | Select-Object Host,RowID,RequestID,RequestSubmitted,Requester,Disposition,Serial,Subject-CN,ValidFrom,ValidTo,EKU,CDP,AIA,AKI,SKI,Subject-C,Subject-O,Subject-OU,Subject-DN,Subject-Email,SAN,Template,Template-Major,Template-Minor,BinaryCert,DaysTilExpiry,RequestStatusCode
        $Row.Host = $CAName
        $Row.RowID = $rowfakeID #($Line.Replace("Row ","")).Replace(":","")
        $Row.EKU = 'Empty'
        $Row.SAN = 'Empty'
        $Row.CDP = 'Empty'
        $Row.AIA = 'Empty'
        
        $rowfakeID++
        if($rowfakeID % 100 -eq 0){
            "$rowfakeID"
        }

        if($rowfakeID % 4000 -eq 0){
                
                # Export-As-We-Go
                Switch($PSCmdlet.ParameterSetName){
                    "Export"    
                        {
                        Write-Host "Exporting CSV records"
                        $Rows | Export-CSV $ExportFile -NoTypeInformation -Append
                        Write-Host "Checkpoint records written to file $ExportFile"
                        $Rows = @()
                        }
                }
        }
    }


        # Prepare the Field and Values
        $Line = $Line.Trim()
        $arrLine = $Line.Split(":",2)
        $Field = $arrLIne[0].Trim()
        $Value = $arrLIne[1]
        $Value = $Value -Replace '"',''
        $Value = $Value.Trim()

        # Well known single lines
        Switch($Field) {
            "Request ID" {$Row.RequestID = $Value}
            "Request Disposition" {$Row.Disposition = $Value}
            "Requester Name" {$Row.Requester = $Value}
			"Request Submission Date" {$Row.RequestSubmitted = $Value}
            "Serial Number" {$Row.Serial = $Value}
            "Certificate Effective Date" {$Row.ValidFrom = $Value -as [datetime]}
            "Issued Common Name" {$Row."Subject-CN" = $Value}
            "Issued Country/Region" {$Row."Subject-C" = $Value}
            "Issued Organization" {$Row."Subject-O" = $Value}
            "Issued Organization Unit" {$Row."Subject-OU" = $Value}
            "Issued Distinguished Name" {$Row."Subject-DN" = $Value}
            "Issued Email Address" {$Row."Subject-Email" = $Value}
            "Issued Subject Key Identifier" {$Row.SKI = $Value}
            # note that this will cause a breakage when trying to go cross-region (eg interpreting an en-US date from anywhere else)
            # so rem it out if needed
            "Certificate Expiration Date" {$Row.ValidTo = $Value}  #{$Row.ValidTo =  $Value-as [datetime]
                                           #$Row.DaysTilExpiry = (([Decimal]::Round((New-TimeSpan $Now $Row.ValidTo).TotalDays))*-1)}   
            
            # and specifically because it's not done in the same way as a "whole" export, we need a template catcher
            "Certificate Template" {$Row.Template = $Value}
            "Request Status Code" {$Row.RequestStatusCode = $Value}
            # OID stripping is an exercise for the reader
        }    
    
        # Process the Multi Line Values if we identified them on the last loop
        Switch ($NextSection){
            "Template" {
                If($Line -match "Template="){
                    $Row.Template = $Line.Split("=",2)[1]}
                If($Line -match "Major Version Number="){
                    $Row."Template-Major" = $Line.Split("=",2)[1]}
                If($Line -match "Minor Version Number="){
                    $Row."Template-Minor" = $Line.Split("=",2)[1]}
                } 
                "AKI"{
                    If(($Line -match "KeyID")){$Row.AKI = $Line}
                }       
                "EKU"{ 
                    If ($Line -ne ''){
                        If($Row.EKU -eq 'Empty'){
                            $Row.EKU = $Line}
                        Else {$Row.EKU = $Row.EKU + "|" + $Line}
                    }
                }
                "SAN"{
                    If(($Line -match "Principal Name=") -or ($Line -match "DNS Name=")){
                        If($Row.SAN -eq 'Empty'){
                            $Row.SAN = $Line}
                        Else {$Row.SAN = $Row.SAN + "|" + $Line}
                    }
                }
                "CDP"{ 
                    If(($Line -match "URL")){
                        If($Row.CDP -eq 'Empty'){
                            $Row.CDP = $Line}
                        Else {$Row.CDP = $Row.CDP + "|" + $Line}
                    }
                }
                "AIA"{ 
                    If(($Line -match "URL")){
                        If($Row.AIA -eq 'Empty'){
                            $Row.AIA = $Line}
                        Else {$Row.AIA = $Row.AIA + "|" + $Line}
                    }
                }
            "BinaryCert"{ 
                    If ($Line -ne ''){
                        $Row.BinaryCert = $Row.BinaryCert + $Line}
                }
            }
            Switch($Field) {
                "Authority Key Identifier" {$NextSection = "AKI"}
                "Subject Alternative Name" {$NextSection = "SAN"}
                "Enhanced Key Usage" {$NextSection = "EKU"}
                "CRL Distribution Points" {$NextSection = "CDP"}
                "1.3.6.1.5.5.7.1.1" {$NextSection = "AIA"}
                "1.3.6.1.4.1.311.21.7" {$NextSection = "Template"}
                "-----BEGIN CERTIFICATE-----" {$NextSection = "BinaryCert";$Row.BinaryCert="-----BEGIN CERTIFICATE-----"}
                "" {$NextSection=$Null}
            }    

        }
"Completed $rowfakeID"

# Finished processing lines
# Add the final row if we reached the end or just stopped receiving "row" lines.
If($DaysTilExpiry){                
    If($Row.DaysTilExpiry -le $DaysTilExpiry -and $Row.DaysTilExpiry -gt 0){
        $Rows += $Row
    }
}
Else{$Rows += $Row}

# Output depending on our purpose
Switch($PSCmdlet.ParameterSetName){
    "Output"    {Write-Host "Running in Output Mode"
                $Rows | Export-Clixml $OutputFile
                Write-Host $Rows.Count " written to output file $OutputFile"}
    "Export"    {Write-Host "Running in CSV Export Mode"
                $Rows | Export-CSV $ExportFile -NoTypeInformation -Append
                Write-Host $Rows.Count " written to file $ExportFile"}
}
# ==============================================    
# End of Main Script
# ==============================================
