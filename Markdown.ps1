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
        [string]$Path
    )

    # HTMLテンプレート
    $html=
@'
<!DOCTYPE html>
<html>
    <head>
        <title>Markdown Preview</title>
    </head>
    <body>
    {0}
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
    if(-not(Test-Path $previewDir -PathType Container))
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
    $html -f $markdown.Transform($rawText) | Out-File $previewPath -Encoding utf8
    
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
            $html -f $markdown.Transform($rawText) | Out-File $previewPath -Encoding utf8
            $ie.Refresh()
        }
        Start-Sleep -Milliseconds 100
    }
}