
# File: DOCX to PDF Converter.ps1
# Author: Clint Kline
# Last Modified: 9/16/2022
# Purpose: convert docx files to pdf files. the idea for this file is to copy it to your desktop and click it(or open with powershell initially). 
#     from there a file dialog opens up to choose a docx file to convert. The file is converted and the new pdf verion is saved next to the original docx file. 


function docxtopdf() {
    #Browsing file

    Write-host(">> Choose .docx file to convert: ")
    Start-Sleep(1)

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    # titled 'Choose a DOCX to convert"
    $FileBrowser.title = "Choose a DOCX to convert"
    # that filters docx files.
    $FileBrowser.filter = "docx (*.docx)| *.docx"
    # open file dialog
    [void]$FileBrowser.ShowDialog()
    # assign selected file to a variable
    $InputFile = $FileBrowser.FileName

    write-host("`n>> `"" + $InputFile + "`" is being converted to PDF format...")

    # create a word object
    $Word = NEW-OBJECT -COMOBJECT WORD.APPLICATION
    # get the current permissions for the docx file
    $ACL = Get-ACL -Path $InputFile
    # get the current username
    $user = $env:UserName
    # set the permissions of the docx file to 'full control'
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($user,"FullControl","Allow")
    $ACL | Set-Acl -Path $InputFile
    Write-Host("`n>> Full control permissions granted to file...`n")
    # $ACL = Get-ACL -Path $InputFile

    # open a Word document, filename from the directory
    $Doc=$Word.Documents.Open($InputFile)

    # Swap out DOCX with PDF in the Filename
    $Name=$Doc.Fullname.replace("docx","pdf")

    # Save this File as a PDF in Word 2010/2013
    $Doc.saveas([ref] $Name, [ref] 17)  

    Write-Host("`n>> DOCX has been converted.`n")
    Start-Sleep(1)
    
    $Doc.close()
    exit    
}

docxtopdf
