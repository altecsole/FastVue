<# 
 - Jim Williams, March 2021

 - Name: Get-FastVueYouTubeTitles.ps1

 - Description:
    Moved from Barracuda Web Filter to Sophos XG, so updated to work with Sophos Fastvue repoter

    FastVue Reporter does a great job of reporting web usage from Sophos XG. One of the best features for our school
    is the ability to run an Activity report to get information on YouTube video usage. Unfortunately, the report doesn't
    include the YouTube video title so this makes it difficult to get a quick overview of what's being watched.

 - Usage: 
    Reads FastVue activity report for YouTube exported as a CSV. It uses a regular expression
    to extract the video URL and user information which are comma separated.
    
    The YouTube URL and watcher are then shown in a grid view for further filtering if reqd. Highlight
    the records you want to process further and hit OK. 

    The URL is then used to get the page tile, which is the same as the video title.

    The result opens in a Grid-View. Data can be filtered, and the results highlighted and exported to CSV.

 - Arguments: -FastVueFile - a FastVue activity report for YouTube videos exported as a CSV file. This is OPTIONAL and
                            will prompt for a file if not included on the command line.
              -Unique - False by default, use this argument if you just want unique video titles

 - NOTE: -Computer domain account
            Sometime the Fastvue reports shows the user as the local machine domain account. To remove this from the 
            results, replace .domain.local with your domain on lines 199 and 223.
         -Video length
            You can tweak the length of videos returned by editing the time in lines 199 and 223. Please note, this is a string
            value, so "00:00:0*" will return videos greater than 10 seconds, "00:0*" for greater than 10 minutes. It would 
            be possible to split the string for more accuracy, but it's not something I need at the moment.

 - UPDATES: 7th Nov 2016 - Found that the format of the Activity report has changed so that ', youtube.com' is no longer
            shown at the end of the line that contains the video url and watcher. Ammended regular expression can commented
            out the old version in case it changes again.

            15th Nov 2016 - Modified the Get-YouTubeTitle function so that it replaces any html char codes in the title
            with the correct chars - this happens a lot!

            2nd Mar 2021 - Updated to work with Sophos XG Fastvue Reporter. 
                         - Changed output based on -Unique:
                           - If Unique; just get video title and number of times watched
                           - If nott unique, get title for each video
                         - Code previoulsy got title for each video. Tweaked so that any titles already retrieved, 
                           are reported from local results, rather than retrieving from the Internet.
                           

######################################################################################################################
#>

param(
    [string]$FastVueFile, # Fastvue activity report for YouTube exported to CSV - OPTIONAL
    [switch]$unique = $false # Get all results by default
    
)


######### Functions ############

Function Get-CSVFile($initialDirectory) {   
    # Use Open File Dialog to get filepath
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.Title = "Select CSV file"
    $OpenFileDialog.filter = "CSV Files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Request-UserToSave {
    # Prompt user for save
    $title = "Save Results"
    $message = "Do you want to save your selections to CSV?"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
        "Save to CSV."

    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
        "Don't bother."

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
    $result
}

Function Save-FileName($initialDirectory) {   
    # Use File Save Dialog to save file
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV Files (*.csv)| *.csv"
    $OpenFileDialog.AddExtension = $true
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}


function Get-YouTubeTitle {
    # Use Invoke-WebRequste to download page into a variable and then 
    # use a regular expresion to get the data inside the title tags.
    param([string] $vidURL)
    $data = Invoke-WebRequest $vidURL
    $title = [regex] '(?im)(?<=<title>)([\S\s]*?)(?= - YouTube</title>)' 
    $result = $title.Match($data).value.trim()
    @(
        @("&#33;", "!"),
        @("&#34;", '"'),
        @("&#35;", "#"),
        @("&#36;", "$"),
        @("&#37;", "%"),
        @("&#38;", "&"),
        @("&#39;", "'"),
        @("&#40;", "("),
        @("&#41;", ")"),
        @("&#42;", "*"),
        @("&#43;", "+"),
        @("&#44;", ","),
        @("&#45;", "-"),
        @("&#46;", "."),
        @("&#47;", "/"),
        @("&#58;", ":"),
        @("&#59;", ";"),
        @("&#60;", "<"),
        @("&#61;", "="),
        @("&#62;", ">"),
        @("&#63;", "?"),
        @("&#64;", "@"),
        @("&#91;", "["),
        @("&#92;", "\"),
        @("&#93;", "]"),
        @("&#94;", "^"),
        @("&#95;", "_"),
        @("&#96;", '`'),
        @("&#123;", "{"),
        @("&#124;", "|"),
        @("&#125;", "}"),
        @("&#126;", "~"),
        @("&quot;", '"'),
        @("&amp;", "&"),
        @("&lt;", "<"),
        @("&gt;", ">")
    ) | ForEach-Object { $result = $result -replace $_[0], $_[1] }
    If ($result -eq "") {
        $result = "NO TITLE - POSSIBLE PRIVATE VIDEO!"
    }
    $result  
}


############ SCRIPT BODY #########
 
$Error.Clear()

if ($FastVueFile) {
    # Check if file passed as an argument
    $input_path = $FastVueFile
}
else {
    $FastVueFile = Get-CSVFile ".\" # if not, use function to get the file
    if ($FastVueFile) {
        $input_path = $FastVueFile
    }
    else {
        write-host "No FastVue CSV file to run against. Exiting..."
        break
    }
}

# Regex for getting the YouTube video URL and user from the FastVue CSV file

# $regex = [regex] '(?i)(?<= )(?=https)([\S\s]*?)(?=")'
$regex = [regex] '(?i)(?=https://www\.youtube\.com/watch)([\S\s]*?)(?=(?:)(?m:$))'
# all instances of the pattern are captured
$tempData = select-string -Path $input_path -Pattern $regex -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }

#$tempData | Out-File C:\Working\Powershell\Fastvue\letmesee.txt


$fastVueResults = @()
# initialise array with dummy object, so we don't need to check for this in the foreach loop
if ($unique) {
    $fastVueResults += New-Object -TypeName PSObject -Pro ([ordered]@{
            YouTube_URL = "placeholder"
            Watch_Count = 0
        })
    ForEach ($entry in $tempData) {
        $arr = $entry.Split(",").Trim('"', " ")

        if (($arr[1] -notlike "*.domain.local" -and $arr[1] -notlike "Video Hosting") -and ($arr[3] -notlike "00:00:0*")) {
            #check if we've already seen this video URL
            $isUnique = $fastVueResults | Where-Object { $_.YouTube_URL -eq $arr[0] }
            if ($null -eq $isUnique) {
                # URL not in array - add
                $fastVueResults += New-Object -TypeName PSObject -Pro ([ordered]@{
                        YouTube_URL = $arr[0]
                        Watch_Count = 1
                    })
            }
            else {
                # URL in list - update watch count
                $isUnique.Watch_Count = $isUnique.Watch_Count + 1
            }

        }
    }
    # remove placeholder object that was created at initialisation, and sort output by Watch_Count
    $fastVueData = $fastVueResults | Where-Object { $_.YouTube_URL -ne "placeholder" } | Sort-Object -Property Watch_Count -Descending
}
else {
    ForEach ($entry in $tempData) {
        $arr = $entry.Split(",").Trim('"', " ")

        if (($arr[1] -notlike "*.domain.local" -and $arr[1] -notlike "Video Hosting") -and ($arr[3] -notlike "00:00:0*")) {
            # Add each pair of data as an object
            $fastVueResults += New-Object -TypeName PSObject -Pro ([ordered]@{
                    YouTube_URL         = $arr[0]
                    Watcher             = $arr[1]

                })
        }
    }
    $fastVueData = $fastVueResults | Sort-Object -Property Watcher
}


# Present results in Grid View so user can filter if reqd.
$subSet = $fastVueData | Out-GridView -PassThru -Title "Select video entries to get the title of.."

# Array of objects for each entry
$youtubeData = @()

# initialise video count to show progress
$count = 1
$total = $subSet.Count

# Used at end of script to display summary of results
$retFromWeb = 0
$retFromHash = 0

# Hashtable of URL, Title. Speeds up script if not run as unique. First entry is got from web, with subsequent from table
$titleHashTable = @{}
#$titleHashTable.Add("https://www.youtube.com/watch?v=55IKKyFBW0Y", "W2S plays SIDEMEN geoguessr - GUERNSEY edition!")
$hashChecker = ""

if ($unique) {
    # unique URL values, all titles got from web
    foreach ($vid in $subSet) {
        if ($count % 20 -eq 0) { Write-Output "Video $count of $total.." }
        #get video title
        $youtubeData += New-Object -TypeName PSObject -Property ([ordered]@{
                YouTube_URL = $vid.YouTube_URL
                Watch_Count = $vid.Watch_Count
                Video_Title = Get-YouTubeTitle $vid.YouTube_URL
            })
        $retFromWeb++
        $count++
    }
}
else {
    # non-unique. Some titles from web, some from hashtable
    foreach ($vid in $subSet) {
        if ($count % 20 -eq 0) { Write-Output "Video $count of $total.." }
        #get video title
        $vidTitle = ""
        #check if video title has alredy been retrieved
        $hashChecker = $titleHashTable[$vid.YouTube_URL] 
        if ($null -eq $hashChecker) {
            # no, get video title
            $vidTitle = Get-YouTubeTitle $vid.YouTube_URL
            # add to hashtable
            $titleHashTable.Add($vid.YouTube_URL, $vidTitle)
            $retFromWeb++
        }
        else {
            # get video title
            $vidTitle = $hashChecker
            $retFromHash++
        }
        $youtubeData += New-Object -TypeName PSObject -Property ([ordered]@{

                YouTube_URL   = $vid.YouTube_URL
                Watcher       = $vid.Watcher
                Video_Title   = $vidTitle
                Browsing_Time = $vid.Browsing_Time
            })
    
        $count++
    }    
}

$resultsSelected = $youtubeData | Out-GridView -PassThru -Title "YouTube video title results"
# If items seleced, prompt to save to CSV file
if ($resultsSelected) {
    $returnValue = Request-UserToSave
    switch ($returnValue) {
        0 { $resultsSelected | Export-Csv -Path (Save-Filename -initialDirectory .\) -Encoding UTF8 -NoClobber -NoTypeInformation }
        1 { "" }
    }
}
write-host "Complete!"
If ($unique) {
    write-host "Unique search: " $retFromWeb " results returnd from the web. "
}
else {
    write-host "Non Unique search: " $retFromWeb " results returnd from the web. " $retFromHash " results returnd from hashtable."
}
