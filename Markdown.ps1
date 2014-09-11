$Script:scriptDir = Split-Path $MyInvocation.MyCommand.Definition -Parent

<#
.Synopsis
    Markdownテキストのプレビュー
.DESCRIPTION
    Markdown形式で記述されたテキストファイルの内容をHTMLに変換し、ブラウザでプレビュー表示します。
.EXAMPLE
    Start-Markdown .\markdownText.md
#>
function Start-Markdown
{
    Param(
        # Markdownテキストファイルのパス
        [Parameter(Mandatory=$true)]
        [string]$Path,
        # スタイルシート名。 HighlightJSで使用可能なスタイル名(.cssは含まない)を指定します。
        [string]$StyleSheet = 'default'
    )

    # HTMLテンプレート
    $html=
@'
<!DOCTYPE html>
<html>
    <head>
        <link rel="stylesheet" href="http://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.2/styles/{0}.min.css">
		<script src="http://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.2/highlight.min.js"></script>
        <script>hljs.initHighlightingOnLoad();</script>
        <title>Markdown Preview</title>
    </head>
    <body>
    {1}
    </body>
</html>
'@
    # フォルダ、ファイルのパス設定
    if(-not(Test-Path $Path))
    {
        new-item $Path -ItemType file -Force
    }
    $sourceDir = Split-Path (Get-ChildItem $Path).FullName -Parent
    $previewDir = "$scriptDir\Html"
    if(-not(Test-Path $previewDir))
    {
        New-Item $previewDir -ItemType directory
    }
    $previewPath = "$previewDir\MarkdownPreview.html"
    
    # MarkdownSharpのコンパイル ＆ Markdown オブジェクトの生成
    Add-Type -Path "$scriptDir\CSharp\Markdown.cs" -ReferencedAssemblies System.Configuration
    $markdown = New-Object MarkdownSharp.Markdown
    
    # Internet Exploroler を起動
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.StatusBar = $false
    $ie.AddressBar = $false
    $ie.MenuBar = $false
    $ie.ToolBar = $false
    
    $rawText = Get-Content $Path -raw
    $html -f $StyleSheet, $markdown.Transform($rawText) | Out-File $previewPath -Encoding utf8
    
    $ie.Navigate($previewPath)
    $ie.Visible = $true
    
    # Markdownテキストファイルの変更を監視する File Watcherを生成
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $sourceDir
    $watcher.Filter = Split-Path $Path -Leaf
    $watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite

    # Markdownテキストファイルを、その拡張子に関連付けされたアプリケーション(エディタ)で起動する
    start $Path

    # ファイルが変更されるのを監視し、変更されたらMarkdownテキストをHTMLに変換してInternet Explorerでプレビューする
    while($ie.ReadyState -ne $null)
    {
        $result = $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Changed, 5000)
        if(-not $result.TimedOut)
        {
            $rawText = Get-Content $Path -raw
            $html -f $StyleSheet, $markdown.Transform($rawText) | Out-File $previewPath -Encoding utf8
            $ie.Refresh()
        }
        Start-Sleep -Milliseconds 100
    }
}