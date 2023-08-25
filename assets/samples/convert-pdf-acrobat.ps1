function convert-pdf

{

<#

  .SYNOPSIS

  Convert PDF file(s) into PowerPoint presentation(s) containing non-editable graphics on each slide.

  .DESCRIPTION

  Function converts PDF file(s) into PowerPoint presentation(s), optionally saving result(s) to network share
  location(s).

  Function takes an array of files, testing each file is a PDF type, then converts each PDF to PowerPoint. This is managed by converting each page of the PDF to a graphic file; content is non-editable. For each graphic, a new slide is created in the PowerPoint presentation and the graphic is added.

  Optionally, the new PowerPoint presentation(s) is moved to a specified network share location(s). Otherwise, it is saved to the current directory.

  .NOTES

  File Name        : Convert-PDF.ps1
  Author           : Justin Townsend

  Create Date      : 23/12/2016
  Purpose / Change : Initial version

  Prerequisite     : Acrobat Standard (v11.0)

  .LINK

  https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-5.1

  .EXAMPLE

  Convert-PDF -pdfs 'C:\test.pdf'

  .EXAMPLE

  Convert-PDF -pdfs 'C:\test.pdf' -outPaths 'C:\'

  .EXAMPLE

  Convert-PDF -pdfs 'C:\test_1.pdf' 'C:\test_2.pdf' -outPaths '<NETWORK SHARE LOCATION>'

  .EXAMPLE

  Convert-PDF -pdfs 'C:\test_1.pdf' 'C:\test_2.pdf' -outPaths '<NETWORK SHARE LOCATIONS>' -infoTbl

  .PARAMETER pdfs

  PDF file(s) for processing (accepts array).

  .PARAMETER outPaths

  Target location(s) to output file converted file(s) (accepts array).

  .PARAMETER infoTbl

  Switch, if invoked, produces a *.csv with useful information for performance logging.

#>

[cmdletbinding()]

param ([Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage='Specify the location(s) of PDF files.',
                   ValueFromPipeline=$true)]
                   [ValidateScript({Test-Path $_ -Pathtype 'Leaf'})]
                   [string[]] $pdfs
                   ,
        [Parameter(Position=1,
                   HelpMessage='Specify the location for output.',
                   ValueFromPipeline=$true)]
                   [ValidateScript({Test-Path $_ -Pathtype 'Container'})]
                   [string[]] $outPaths = [environment]::CurrentDirectory
                   ,
        [Parameter(Position=2,
                   HelpMessage='Information Classification.',
                   ValueFromPipeLine=$true)]
                   [ValidateSet('HIGHLY RESTRICTED','RESTRICTED','INTERNAL','PUBLIC','Non - Company')]
                   [string] $ConfidentialityClass = "RESTRICTED"
                   ,
        [Parameter(Position=3,
                   HelpMessage='CSV table for process monitoring.')]
                   [switch] $infoTbl
                   )

:pdfloop foreach ($pdf in $pdfs)

{

   $pdf = get-childitem $pdf
   $out_dir = $pdf.DirectoryName
   $out_dir = $out_dir + "\" + $pdf.Basename
   $out_dir += "_PROC"
   $out_file = $out_dir + "\" + $pdf.Basename

   new-item $out_dir -type directory -force

   # Adobe Acrobat Standard (convert to graphic files)
   $adobeApp = New-Object -ComObject AcroExch.AVDoc;
   $adobeApp.Open($pdf.Fullname, "") | Out-Null;
   $pdfDoc = $adobeApp.GetPDDoc();
   $pdfJSObject = $pdfDoc.GetJSObject();

   $TypeExt="jpeg";
   $closeDocParam = $true;
   $T = $pdfJSObject.GetType();

   $T.InvokeMember("SaveAs",
     [Reflection.BindingFlags]::InvokeMethod -bor `
     [Reflection.BindingFlags]::Public       -bor `
     [Reflection.BindingFlags]::Instance,
     $null,
     $pdfJSObject,
     @([IO.Path]::ChangeExtension($out_file, $TypeExt), ("com.adobe.acrobat."+$TypeExt)));

   $T.InvokeMember("closeDoc",
     [Reflection.BindingFlags]::InvokeMethod -bor `
     [Reflection.BindingFlags]::Public       -bor `
     [Reflection.BindingFlags]::Instance,
     $null,
     $pdfJSObject,
     $closeDocParam) | Out-Null;
     $pdfDoc.Close()  | Out-Null;
     $adobeApp.Close(1) | Out-Null;

   # Microsoft PowerPoint (create presentation with graphic files)
   Add-type -AssemblyName office
   $msoappPPT = New-Object -ComObject powerpoint.application
   $msoappPPT.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
   $slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]
   $out_ppt = $pdf.DirectoryName + "\" + $pdf.Basename
   $pptPres = $msoappPPT.Presentations.add()

   # PageSetup.SlideSize late binding 
   $pptPres.PageSetup.SlideSize = 3
   get-childitem -path $out_dir | ForEach-Object { `
     $pic = $_.fullname
     $add_slide = $pptPres.Slides.Add($pptPres.Slides.Count + 1, 15);
     $add_slide.layout = $slideType::ppLayoutBlank;
     $add_slide.HeadersFooters.Footer.Visible = 1;
     $add_slide.HeadersFooters.Footer.Text = $ConfidentialityClass
     $add_slide.Shapes.AddPicture($pic, $false, $true, 0, 0, -1, -1);
   }

   $pptPres.SaveAs($out_ppt)
   $pptPres.Close()
   $msoappPPT.quit()
   $msoappPPT = $null

   # Produce a *.csv with useful information for performance logging.
   if $infoTbl {
   get-childItem $pdf | Select-Object BaseName, Name, Extension, Length, CreationTimeUtc, LastWriteTimeUtc, DirectoryName | Export-Csv perf.csv -NoTypeInformation
   get-childItem -Path $out_dir -Recurse |`
     foreach{
       $fBasename = $_.BaseName
       $fName = $_.Name
       $fExtension = $_.Extension
       $fLength = $_.Length
       $fCreationTimeUtc = $_.CreationTimeUtc
       $fLastWriteTimeUtc = $_.LastWriteTimeUtc
       $fDirectoryName = $_.DirectoryName
       $Path | Select-Object `
           @{n="Basename";e={$fBasename}},`
           @{n="Name";e={$fName}},`
           @{n="Extension";e={$fExtension}},`
           @{n="Length";e={$fLength}},`
           @{n="CreationTimeUtc";e={$fCreationTimeUtc}},`
           @{n="LastWriteTimeUtc";e={$fLastWriteTimeUtc}},`
           @{n="DirectoryName";e={$fDirectoryName}}`
       }| Export-Csv -Append perf.csv -NoTypeInformation

     $ppt_info = $out_ppt + ".pptx"

     get-childitem $ppt_info | Select-Object BaseName, Name, Extension, Length, CreationTimeUtc, LastWriteTimeUtc, DirectoryName | Export-Csv -Append perf.csv -NoTypeInformation

    }

}

   Remove-Item $out_dir -recurse

}