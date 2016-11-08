# FastVue YouTube Activity Report - Video Titles

We use FastVue Reporter to get great reports of web usage from our Barracuda Web Filter. One of the really good reports that we use is an Activity report on YouTube access. This shows the YouTube URL with the watcher and times, but doesn't include the title of the video. The activity report does a great job of showing the top users who are watching YouTube, but we really wanted to know the title of the video to put some context on it.

This Powershell script parses the Activity report (exported as a CSV file) for the YouTube URL and the user. This information is then presented in a Grid-View where you can filter if required.

Once you're happy with the results, select all the records you want and click OK. The script will then do a web request to get the page title from the URL (which is the same as the video title) and presents this information in another Grid-View. Highlight the entries you want and click OK to be prompted to save the results as a .CSV file.

## Notes
1. You don't need to specify the FastVue file at the command line - you will be prompted for it.
2. The default is to get all URLs; if you're only interested in unique URLs you can specify -Unique on the command line.
3. If you open the resultant .CSV file in Excel you will probably want to have the URLs as clickable. The simplest way to do this is to create a new column and use the 'hyperlink' function. If I'm sharing, I then hide column A and save as a spreadsheet.
