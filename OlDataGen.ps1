#
# Outlook Data Generator 
# Version 1.0
#
# Copyright (c) 2017 Hernán Di Pietro
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy of 
# this software and associated documentation files (the "Software"), to deal in 
# the Software without restriction, including without limitation the rights to use, 
# copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, 
# and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial 
# portions of the Software. THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, 
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS 
# FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS 
# BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
# TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
# DEALINGS IN THE SOFTWARE.

Param (
    [Parameter(Mandatory=$True)] [int] $ItemCount = 1,
    [int] $SubjectWordCountMin  = 1,
    [int] $SubjectWordCountMax  = 8,
    [int] $BodyWordCountMin     = 30,
    [int] $BodyWordCountMax     = 200,
    [string] $WordSourceFile    = '.\words.txt',
    [string] $AddressSourceFile = '.\addr.txt'
    )

     
Write-Host -ForegroundColor Green "-----------------------------------------------------------------------"
Write-Host -ForegroundColor Green "Outlook Data Generator V1.0      "
Write-Host -ForegroundColor Green "Copyright 2017 Hernan Di Pietro  " 
Write-Host -ForegroundColor Green " " 
Write-Host -ForegroundColor Green "This utility is MIT-Licensed. "
Write-Host -ForegroundColor Green "Visit https://github.com/hernandp/OutlookDataGenerator for details     "
Write-Host -ForegroundColor Green "-----------------------------------------------------------------------"
Write-Host "Setting up..."

# Setup Outlook interop API and objects
# -------------------------------------

Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop

$Outlook      = New-Object -ComObject Outlook.Application -ErrorAction Stop
$ns           = $Outlook.GetNamespace("MAPI") 
$olFolders    = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
$inbox        = $ns.GetDefaultFolder($olFolders::olFolderInbox);

# Read address and word source files
# ----------------------------------
Write-Host "Reading word source file $WordSourceFile..."
$wordArray = Get-Content -Path $WordSourceFile -ErrorAction Stop

Write-Host "Reading address source file $AddressSourceFile..."
$addrArray = Get-Content -Path $AddressSourceFile -ErrorAction Stop

#
# Create mail items 
#
##########################################################################

Write-Host "Creating mail items..."

$totalSizeBytes = 0;
for ($i=0; $i -le $ItemCount; $i++)
{
    if (($i % 100) -eq 0)
    {
       Write-Host -NoNewline(".")
    }

    if (($i % 1000) -eq 0)
    {
       Write-Host -NoNewline $i
    }

    $numWords = Get-Random -Minimum $SubjectWordCountMin -Maximum $SubjectWordCountMax

    $subject = "";
    $text = "";

    for($j=0; $j -le $numWords; $j++)
    {
        $w = Get-Random -Minimum 0 -Maximum $wordArray.Length
        $subject += " " + $wordArray[$w]
    }
        
    # Mail text
    $numWords = Get-Random -Minimum $BodyWordCountMin -Maximum $BodyWordCountMax

    for($k=0; $k -le $numWords; $k++)
    {
        $w = Get-Random -Minimum 0 -Maximum $wordArray.Length
        $text += " " + $wordArray[$w]
    }


    $mailItem = $inbox.Items.Add(0) # 0 is MailItem
    $mailItem.Subject = $subject
    $mailItem.Body = $text
    $fromAddrIdx = Get-Random -Minimum 0 -Maximum $addrArray.Length
    $toAddrIdx = Get-Random -Minimum 0 -Maximum $addrArray.Length
    $mailItem.Sender = $addrArray[$fromAddrIdx]
    $mailItem.To = $addrArray[$toAddrIdx]
    $mailItem.Save()

    $totalSizeBytes += $mailItem.Size
}

#
# print summary
#
##########################################################################
$totalSizeK = ($totalSizeBytes / 1024) -as [Int32]
Write-Host " "
Write-Host "Summary: "
Write-Host -ForegroundColor Cyan "Generated ~ $totalSizeK KBytes in $ItemCount item(s)"
