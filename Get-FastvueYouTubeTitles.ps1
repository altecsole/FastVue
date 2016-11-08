<# 
 - Jim Williams, Oct 2016

 - Name: Get-FastVueYouTubeTitles.ps1

 - Description:
    Work with Barracuda Web Filter logs exported to FastVue Reporter - http://www.fastvue.co/barracuda/

    FastVue Reporter does a great job of reporting web usage from Barracuda Web Filter. One of the best features for our school
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

 - UPDATES: 7th Nov 2016 - Found that the format of the Activity report has changed so that ', youtube.com' is no longer
            shown at the end of the line that contains the video url and watcher. Ammended regular expression can commented
            out the old version in case it changes again.

######################################################################################################################
#>

param(
    [string]$FastVueFile, # Fastvue activity report for YouTube exported to CSV - OPTIONAL
    [switch]$unique = $false # Get all results by default
    
)


######### Functions ############

Function Get-CSVFile($initialDirectory)
{   
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

function Prompt-UserToSave{
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

Function Save-FileName($initialDirectory)
{   
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


function Get-YouTubeTitle{
    # Use Invoke-WebRequste to download page into a variable and then 
    # use a regular expresion to get the data inside the title tags.
    param($vidURL)
    $data = Invoke-WebRequest $vidURL
    $title = [regex] '(?im)(?<=<title>)([\S\s]*?)(?= - YouTube</title>)' 
    $title.Match($data).value.trim()
}


############ SCRIPT BODY #########
 
$Error.Clear()

if($FastVueFile){ # Check if file passed as an argument
    $input_path = $FastVueFile
}
else{
    $FastVueFile = Get-CSVFile ".\" # if not, use function to get the file
    if($FastVueFile){
        $input_path = $FastVueFile
    }
    else{
        write-host "No FastVue CSV file to run against. Exiting..."
        break
    }
}

# Regex for getting the YouTube video URL and user from the FastVue CSV file
<#  
    Old version Regex that pulled URL and watcher from the Activity file.
    The latest version of the file does not seem to append 'youtube.com' to the
    end of the line.

    $regex = [regex] '(?i)(?<= )(?=https)([\S\s]*?)(?=, youtube.com)'

#>

$regex = [regex] '(?i)(?<= )(?=https)([\S\s]*?)(?=")'
# all instances of the pattern are captured
$tempData = select-string -Path $input_path -Pattern $regex -AllMatches | % { $_.Matches } | % { $_.Value }

$fastVueResults = @()
# split csv data into seperate URL and User info
ForEach($entry in $tempData){
    $arr = $entry -split ','

    # Add each pair of data as an object
    $fastVueResults += New-Object -TypeName PSObject -Pro ( [ordered]@{
        YouTube_URL = $arr[0].Trim()
        Watcher = $arr[1].Trim()
    })
}

if($unique){
    # We now want to sort the objects and just get the unique video titles
    $fastVueData = $fastVueResults | Sort-Object -Property YouTube_URL -Unique 
}
else{
    # Get all entries
    $fastVueData = $fastVueResults | Sort-Object -Property Watcher
}

# Present results in Grid View so user can filter if reqd.
$subSet = $fastVueData | Out-GridView -PassThru -Title "Select video entries to get the title of.."

$youtubeData = @()
$count = 1
$total = $subSet.Count

foreach($vid in $subSet){
    if($count % 20 -eq 0){Write-Output "Video $count of $total.."    }
    #get video title
    $youtubeData += New-Object -TypeName PSObject -Property ([ordered]@{
        YouTube_URL = $vid.YouTube_URL
        Watcher = $vid.Watcher
        Video_Title = Get-YouTubeTitle $vid.YouTube_URL
    })
    
    $count++
}


$resultsSelected = $youtubeData | Out-GridView -PassThru -Title "YouTube video title results"
# If items seleced, prompt to save to CSV file
if($resultsSelected){
    $returnValue = Prompt-UserToSave
    switch ($returnValue){
        0 {$resultsSelected | Export-Csv -Path (Save-Filename -initialDirectory .\) -Encoding UTF8 -NoClobber -NoTypeInformation}
        1 {""}
    }
}