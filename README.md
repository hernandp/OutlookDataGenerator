# Outlook Data Generator

This small utility uses the Outlook Interop API through Powershell for generating random mail items. This is useful when developing addins and/or Outlook and Office interoperability solutions.

I expect OlDataGen to run on reasonably modern versions of Outlook (>2010).

Current version generates plain-body messages. Content is formed randomly with words chosen from a file (WORDS.TXT is provided). 

Addresses are also gathered from an external file (see ADDR.TXT).

## Parameters

| Parameter     | Explanation                | Default value |
| ------------- | ---------------------------| --------------|
| `-ItemCount`  | Sets the number of items to generate |  N/A |
| `-WordSourceFile` | External file containing words to generate body content | `.\words.txt`|
| `-AddressSourceFile` | External file containing addresses for To/From fields | `.\addr.txt`|
| `-SubjectWordCountMin` | Minimum number of words in subject | 1 | 
| `-SubjectWordCountMax` | Maximum number of words in subject | 8 |
| `-BodyWordCountMin` | Minimum number of words in message body | 30 |
| `-BodyWordCountMax` | Maximum number of words in message body | 200 |

## Sample Output
```
PS C:\Users\Hernan\Documents> .\OlDataGen.ps1 -ItemCount 3000
-----------------------------------------------------------------------
Outlook Data Generator V1.0      
Copyright 2017 Hernan Di Pietro  
 
This utility is MIT-Licensed. 
Visit https://github.com/hernandp/OutlookDataGenerator for details     
-----------------------------------------------------------------------
Setting up...
Reading word source file .\words.txt...
Reading address source file .\addr.txt...
Creating mail items...
.0..........1000..........2000..........3000 
Summary: 
Generated ~ 16053 KBytes in 3000 item(s)
```
