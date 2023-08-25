function Convert-NuPDF

{

<#

  .SYNOPSIS

  Convert PDF file(s) into PowerPoint presentation(s) containing non-editable graphics on each slide.

 .DESCRIPTION

  Function converts PDF file(s) into PowerPoint presentation(s), optionally saving result(s) to network share
  location(s).

  Function takes an array of files, testing each file is a PDF type, then converts each PDF to PowerPoint. This is managed by converting each page of the PDF to a graphic file; content is non-editable. For each graphic, a new slide is created in the PowerPoint presentation and the graphic is added.

  Optionally, the new PowerPoint presentation(s) is moved to specified network share location(s). Otherwise, it is saved to the current directory.

  .NOTES

  File Name        : Convert-PDF.ps1
  Author           : Justin Townsend

  Create Date      : 23/12/2016
  Purpose / Change : Initial version

  Prerequisite     : Nuance Power PDF Advanced (v1.2)

  .LINK

  https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-5.1

  .EXAMPLE

  Convert-NuPDF -pdfs "C:\test.pdf"

  .EXAMPLE

  Convert-NuPDF -pdfs "C:\test_1.pdf", "C:\test_2.pdf" -ConfidentialityClass "HIGHLY RESTRICTED"

  .EXAMPLE

  Convert-NuPDF -pdfs "C:\test_1.pdf", "C:\test_2.pdf" -ConfidentialityClass "INTERNAL" -infoTbl

  .EXAMPLE

  Convert-NuPDF -pdfs "C:\test_1.pdf", "C:\test_2.pdf" -infoTbl

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
                  [ValidateScript({ foreach ($pdf in $_) { if (![bool]($pdf -like '*.pdf')) { throw "$($pdf) is an invalid PDF file!" } } return $true })]
                  [string[]] $pdfs
                   ,
        [Parameter(Position=1,
                   HelpMessage='Information Classification.',
                   ValueFromPipeLine=$true)]
                   [ValidateSet("HIGHLY RESTRICTED","RESTRICTED","INTERNAL","PUBLIC","Non - Company")]
                   [string] $ConfidentialityClass = "RESTRICTED"
                   ,
        [Parameter(Position=2,
                   HelpMessage='CSV table for process monitoring.')]
                   [switch]$infoTbl
                   )

:pdfloop foreach ($pdf in $pdfs)

{

   # Nuance Batch Converter (convert to graphic files)
   $pdf = get-childitem $pdf
   $outExt="jpg"
   $out_dir = $pdf.DirectoryName
   $out_dir = $out_dir + "\" + $pdf.Basename
   $out_dir += "_PROC"
   $out_file = $out_dir + "\" + $pdf.Basename + "." + $outExt

   new-item $out_dir -type directory -force
   & "C:\Program Files\Nuance\Power PDF\batchconverter" -I"$pdf" -O"$out_file" -TTIF -CcJpegMax -Q

   # Microsoft PowerPoint (create presentation with graphic files)
   Add-type -AssemblyName Office
   Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

   $msoappPPT = New-Object -ComObject powerpoint.application
   $msoappPPT.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
   $slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]
   $slideSize = "microsoft.office.interop.powerpoint.ppSlideSizeType" -as [type]
   $msoSendToBack = 1;

   $out_ppt = $pdf.DirectoryName + "\" + $pdf.Basename
   $pptPres = $msoappPPT.Presentations
   $pptPres = $pptPres.add()
   $pptPresPS = $pptPres.PageSetup
   $pptPresPS_SS = $pptPresPS.slideSize
   $pptPresPS_SS = $slideSize::ppSlideSizeA4Paper;

   # get-childitem -path $out_dir | sort-object -Property CreationTime;
   get-childitem -path $out_dir | sort-object -Property CreationTime | ForEach-Object { `
     $pic = $_.fullname
     $ppt_slides = $pptPres.Slides
       $add_slide = $ppt_slides.Add($pptPres.Slides.Count + 1, 15)
         $slide_layout = $add_slide.layout
         $slide_layout = $slideType::ppLayoutBlank
         $slide_HF = $add_slide.HeadersFooters
         $slide_HF_F = $slide_HF.Footer
         $slide_HF_F_vis = $slide_HF_F.Visible
         $slide_HF_F_vis = [Microsoft.Office.Core.MsoTriState]::msoTrue
         $slide_HF_F_txt = $slide_HF_F.text
         $slide_HF_F_txt = $ConfidentialityClass;
       $add_shapes = $add_slide.Shapes
       $add_shape_Rng = $add_shapes.Range("Footer Placeholder 2").Left = -100;
     $shape = $add_shapes.AddPicture($pic, $false, $true, 0, 0, -1, -1);
     $shape.ZOrder($msoSendToBack);
   }

   $pptPres.SaveAs($out_ppt)
   $pptPres.Close()
   $msoappPPT.quit()
   $msoappPPT = $null;

}

# Produce a *.csv with useful information for performance logging.
if (($infoTbl -eq $true))
{
   $perf_csv = $pdf.DirectoryName + "\" + "perf.csv"
   get-childItem $pdf | select-object BaseName, Name, Extension, Length, CreationTimeUtc, LastWriteTimeUtc, DirectoryName |`
   export-csv $perf_csv -NoTypeInformation
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
     }| export-csv -Append $perf_csv -NoTypeInformation

   $ppt_info = $out_ppt + ".pptx"
   get-childitem $ppt_info |`
   select-object BaseName, Name, Extension, Length, CreationTimeUtc, LastWriteTimeUtc, DirectoryName |`
   export-csv -Append $perf_csv -NoTypeInformation
}

  Remove-Item $out_dir -recurse

}