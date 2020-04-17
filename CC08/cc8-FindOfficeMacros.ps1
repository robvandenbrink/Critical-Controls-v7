# this script checks word and excel documents for zoneid, presence of macros and zero-byte file
# the file owner, last write date and last access date is also collected
# This can easily be extended to include project, visio and other office files
# updated source code will be located at: https://github.com/robvandenbrink

$targetlist = @()
$filelist = @()
$resultslist = @()

### input data ###
# file extensions of interest - update as needed
$exts = "xls","xlsx","doc","docx","docm","dotm","xlm","xlsxm"

# the share to enumerate - use the knowledge of your environment to make an effective choice here
# or as Indiana Jones was told "choose wisely"
# if the user is running this, to check that users' temp directory:
# $targetshare = $env:temp
# or if you are targeting a user or department share, specify the
# full path to the share
# $targetshare = "\\some\fully\qualified\unc"
# in any case, update this variable to best suite your organization and situation:

$targetshare = "L:\testing"


# add a trailing backslash if not present in the share defined above
if ($targetshare.substring($targetshare.length -1) -ne "\") { $targetshare += "\" }

# collect all of the filenames that match the identified extensions
# this can take a while in a large environment
foreach ($ext in $exts) {
    $fullpath = $targetshare + "*." + $ext
    $targetfiles = get-childitem -Path $fullpath -Recurse -file -Force
    $filelist += $targetfiles
    }

# yes, this loops multiple times, so is less efficient time-wise, but is more efficient in how
# many filenames are collected.  
# The alternative is to make one pass, collect all the filenames and winnow the list down from there
# if you prefer that option, it would look something like:
#
# $allfiles = get-childitem -targetshare -recurse -file -force
# foreach ($ext in $exts) {
#    $tfiles = $targetfiles | where-object { $_.name -like "."+$ext }
#    $filelist += $tfiles
#    }
# or if you want to do it in one line, you can pipe the get-childitem statment into a where-object command,
# with your various extensions hard-coded (hard-coding anything is $bad)


# with the targetfile list collected, loop through and collect:
#            which zone did the file come from?
#            does the file contain macros?
#            when was the file created?
#            when was the file last accessed?
#            is the file password protected?  (another common IoC for malware, but real people do this too)
#            and who saved the file? (who is the file owner)
#
# this opens each file in the matching MS Office application, so it can take a while as well
# be sure to open the various office apps **once**, then open each file in turn, collect the data,
# then close that file before proceeding to the next one.
# be sure to close the office app when done
#

# Open the office apps.  Set them both to run in the background
$objExcel = New-Object -ComObject Excel.Application
$objWord = New-Object -ComObject Word.Application
$objExcel.visible = $false
# Disable macro execution, either using the value or string method
# thanks to our anonymous reader for pointing this out
# also disable alerts (note that this does not apply to alerts due to macros)
$objExcel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
$objExcel.DisplayAlerts = $false
$objWord.visible = $false
$objWord.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
$objWord.DisplayAlerts = $false

##########
# Vars for file open
# XLImportFormat is set to 5, don't convert anything.  This isn't used, but is needed to
# test for password-protected files (you can't skip variables as you open files)
$ConfirmConversions = $false
$UpdateLinks = 0
$ReadOnly = $true
$AddToRecentFiles = $false
$XLImportFormat = 5

foreach ($indfile in $filelist) {
    $f = $indfile.fullname
    $ext = $indfile.extension

    # zero out critical values for each loop
    $hasmacro = $false
    $hasxl4macro = $false
    $zone = 0
    $pwdprotected = $false
    $zerosize = $false
 
    # zero size?
    if ($indfile.length -eq 0)  { $zerosize = $true }

    # collect alt datastream info (zone)
    $b = (get-content $f -stream Zone.Identifier -erroraction 'silentlycontinue' )

    if ( $b.length -gt 0 ) { $zone = ($b -match "ZoneId").split("=")[1] }


    # EXCEL SECTION

    # skip zero byte files, but record them - possibly AV caught these during a file save
    # also check for and skip pwd protected files. Record them as *potential* malware

    if(( $ext.substring(0,3).tolower() -eq ".xl") -and (-not $zerosize)) {
        # collect excel specific info (are there macros?)
        # full path is required to open the file
        # echo the filename with path, just so we can monitor progress
        # and be sure the script is still running :-)
        write-host $f

        # is it password protected?
        try {
            $WorkBook = $objExcel.Workbooks.Open($f,$UpdateLinks,$ReadOnly,$XLImportFormat,"a")
            }
        catch {
            $pwdprotected = $true
            }
        }
    $error.clear()

    # check the file if we are able to, then close it:
    if((-not $pwdprotected) -and (-not $zerosize)) {
        # excel macros?
        $hasmacro = $workbook.hasvbproject
        $hasxl4macro = $objExcel.Excel4MacroSheets.count + $objExcel.Excel4IntlMacroSheets.count
        $WorkBook.close($false)
        }
    }

    # WORD SECTION
    if(( $ext.substring(0,3).tolower() -eq ".do") -and (-not $zerosize)) {
        # collect word specific info (are there macros?)
        #full path is required to open the file
        write-host $f
        # is it password protected? (a dummy password will trigger

        # the error condition if a password exists, no error if no pwd
        try {
            $Doc = $objWord.documents.Open($f,$ConfirmConversions,$ReadOnly,$AddToRecentFiles,"a")
            } catch {
            $pwdprotected = $true
            }
        $error.clear()

        if((-not $pwdprotected) -and (-not $zerosize)) {
            # word macros?
            $hasmacro = $doc.hasvbproject
            $Doc.close($false)
            }
        }

    # add all info to the list
    $tempobj = [pscustomobject]@{
                fname = $f
                zone = $zone
                macro = ($hasmacro -or $hasxl4macro)
                PwdProtected = $pwdprotected
                ZeroSize = $zerosize
                LastAccessTime = $indfile.LastAccessTime.tostring()
                Owner = $indfile.GetAccessControl().Owner
                }
    # add short wait for RPC bug in some Windows / Office versions
    sleep(1)
    $resultslist += $tempobj
    }

# Close out the two apps
$objWord.quit()
$objExcel.quit()

# macros from the internet
$resultslist  | Where { ($_.zone -eq 3) -and ($_.macro -eq $true) } | out-gridview
pause

# all macros
$resultslist | where { $_.macro -eq $true } | out-gridview
pause

# Zero sized files?
$resultslist | where { $_.zerosize -eq $true } | out-gridview
pause
