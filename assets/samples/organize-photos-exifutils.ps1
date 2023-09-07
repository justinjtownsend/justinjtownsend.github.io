Param ( [Parameter(Mandatory=$True,Position=1)] [string]$psdir )
$erroractionpreference = "SilentlyContinue"
Clear-Host
 
<#
 
.SYNOPSIS
    Script is designed to organise pictures into folders using the date-taken.
 
.DESCRIPTION
    Script organises pictures into folders using the date the photo was taken. It requires installation of EXIFutils. The script extracts the date-taken from the EXIF properties of the picture and uses this date to create a folder. The picture is MOVED once the new folder is created.
   
    EXIFutils v3.0 (non-licensed version) was used when building this script. exiflist is called for each file processed due to the limitations set for the non-licensed version of EXIFutils.
   
    The author takes no responsibility for correct functioning of the script with any other versions of EXIFutils.
   
    ***It is STRONGLY recommended you back-up your pictures before using this script.***
   
.NOTES
    File Name    :    JT_EXIFutils_date-taken.ps1
    Author       :    Justin Townsend
    Prerequisite :    EXIFutils v3.0
 
.LINK
    EXIFutils available at:
    http://www.hugsan.com/EXIFutils/
   
.SETTINGS
   Setting parameters, processing directories and other script properties.

   1. Input Parameters (see line 1)
   2. Photo log file
   3. Logging function (fn_plog)
   4. Picture processing directories

 #>

# 2. Photo log file.
$stamp = get-date -format yyyy.MMM.dd.HH.mm.ss
$plog = "myphotos_" + $stamp + ".log"
new-item -path . -name $plog -type file
 
# 3. Logging function
function fn_plog ([string]$msg)
{
     $logstamp = get-date -format "yyyy.MMM.dd HH:mm:ss.fff"
     "$logstamp :>  $msg" | out-file $plog -Append
}
 
fn_plog "Photo log created."

# 4. Picture processing directories (source-parent, target-parent, dump)
# $ptdir = $psdir
# $ptdir += "\Photos"

$ptdir = "C:\Users\Justin\Pictures\Photos"

 if ( test-path $ptdir ) # Target directory
 {
     fn_plog "$ptdir present."
 }
else
 {
     new-item $ptdir -type directory
     fn_plog "$ptdir created."
 }

$ddir = $ptdir
$ddir += "\1_UNPROCESSED"

 if ( test-path $ddir ) #Dump directory
 {
     fn_plog "$ddir present."
 }
else
 {
     new-item $ddir -type directory
     fn_plog "$ddir created."
 }
 
<#
   Picture processing.
  
   1. If date-taken is <NULL> or not present, move file to dump directory.
   2. If date-taken is present, create folder and move picture file.

#>
 
    foreach ($pic in get-childitem -path $psdir -include *.JPG, *.jpeg, *.jpg -recurse | get-itemproperty | select-object -ExpandProperty fullname)
   
    # Pictures Date Taken
    # { exiflist /o l /f date-taken $pic }
           
    {
     $PSName = $pic | get-itemproperty | select-object -ExpandProperty PSChildName
     $fdir = $ddir + "\" + $PSName
     
     $erroractionpreference = "SilentlyContinue"
     $dt = exiflist /o l /f date-taken $pic
     $cmake = exiflist /o l /f make $pic
     $cmodel = exiflist /o l /f model $pic
     
     $obj = new-object PSObject   
     $obj | add-member Noteproperty Filename $PSNAME
     $obj | add-member Noteproperty Camera_Make $cmake
     $obj | add-member Noteproperty Camera_Model $cmodel
     $obj | add-member Noteproperty Date_Taken $dt
     
     write-output "Processing..."
     write-output $obj
     $erroractionpreference = "SilentlyContinue"
     
        if ( $dt )
        {
         $fn = exiflist /o l /f file-name $pic
         $y = $dt.substring(0,4)
         $d = $dt.substring(0,10) -replace ":", "_"
 
         $y = $ptdir + "\" + $y
         $d = $y + "\" + $d
       
            if ( test-path $d )
            {
             move-item $pic $d
             fn_plog "INFO $fn : Moved to $d."
            }
            elseif ( test-path $y )
            {
             new-item $d -type directory #test
             move-item $pic $d
             fn_plog "INFO $fn : $d created."
             fn_plog "INFO $fn : Moved to $d."
            }
            else
            {
             new-item $y -type directory
             new-item $d -type directory
             move-item $pic $d
             fn_plog "INFO $fn : $y created, $d created."
             fn_plog "INFO $fn : Moved to $d."
            }
        }       
        elseif ( test-path $fdir )
        {
         $stamp = get-date -format yyyy.MMM.dd.HH.mm.ss
         $fdir += "_$stamp"
         move-item $pic $fdir
         fn_plog "ERROR $PSName : EXIF date-taken not available. Moving to $ddir."
         fn_plog "ERROR $PSName : Moved to $fdir because $PSName already exists!"
        }
        else
        {
         fn_plog "ERROR $PSName : EXIF date-taken not available. Moving to $ddir."
         $fdir = $ddir + "\" + $PSName
         move-item $pic $ddir
         fn_plog "ERROR $PSName : Moved to $ddir."
          
        }
    }
 
# Success Rate
$err_cnt = @(get-childitem -recurse -path $ddir).Count
$all_cnt = @(get-childitem -recurse -path $ptdir).Count
$succ_cnt = $all_cnt - $err_cnt
$succ_rate = ($succ_cnt / $all_cnt) * 100
   
fn_plog "Successes        : $succ_cnt"
fn_plog "Failures         : $err_cnt"
fn_plog "Success Rate (%) : $succ_rate"
 
fn_plog "$err_cnt pictures couldn't be processed. You can find them in $ddir."
fn_plog "Consider moving them to folders using other information (e.g. date-created, date-modified)."

$p_obj = new-object PSObject
$p_obj | add-member Noteproperty Files_Processed $all_cnt
$p_obj | add-member Noteproperty Errors $err_cnt
$p_obj | add-member Noteproperty Successes $succ_cnt
$p_obj | add-member Noteproperty Success_Rate $succ_rate

write-output "***** SUCCESS RATE *****"
write-output $p_obj

remove-item $p_obj