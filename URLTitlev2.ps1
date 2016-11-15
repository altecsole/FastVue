Function Get-Title { 
    param([string] $url) 
    $wc = New-Object System.Net.WebClient 
    $data = $wc.downloadstring($url) 
    $title = [regex] '(?im)(?<=<title>)([\S\s]*?)(?= - YouTube</title>)'
    write-output $title.Match($data).value.trim() 
}

function Get-YouTubeTitle{
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
        @("&#126;", "~")
    ) | ForEach-Object {$result = $result -replace $_[0], $_[1]}
    $result  
}



Function Replace-HTMLCharCode{
    param([string] $htmlString)
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
    ) | ForEach-Object {$htmlString = $htmlString -replace $_[0], $_[1]}
    Write-Output $htmlString

}

#$string = Get-YouTubeTitle 'https://www.youtube.com/watch?v=2VWDJX2jouc'
$string = "Hello, &#96;this&#96; is Jim's &#34;computer&#34; &quot;How you &amp; like this? &lt;&gt;"
$newTest = Replace-HTMLCharCode $string
Write-Host $newTest
