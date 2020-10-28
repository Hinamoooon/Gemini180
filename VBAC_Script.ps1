echo "VBAソースコードを出力します。"
echo "VBAC in Ariawase v0.6.0"

cd $PSScriptRoot

#Sourceフォルダに内容物ある時、確認
if ( test-path .\Source\* ){
    $title = "*** 実行確認 ***"
    $message = "Sourceフォルダの内容を上書きします。実行してよろしいですか？"

    $objYes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","実行"
    $objNo = New-Object System.Management.Automation.Host.ChoiceDescription "&No","実行終了"
    $objOptions = [System.Management.Automation.Host.ChoiceDescription[]]($objYes, $objNo)
    $resultVal = $host.ui.PromptForChoice($title, $message, $objOptions, 1)

    if ($resultVal -ne 0) { exit } else {
        remove-item .\Source\*
    }
} elseif (test-path .\Source) {
        echo " "
        } else {
        New-Item Source -ItemType Directory
}
        


#VBAC.wsfを持ってくる
Copy-Item C:\Users\Tatsuki\Documents\Program\vbacGUI\vbac.wsf $PSScriptRoot

#コードをエクスポートしたいexcelファイルをbinフォルダへコピー
New-Item bin -ItemType Directory
Copy-Item *.xlsm bin

cscript vbac.wsf decombine

$FolderName = get-childitem -Name src
echo $FolderName
$FolderPath = Join-Path .\src $FolderName
echo $FolderPath

Push-Location $FolderPath
move-Item *.bas ..\..\Source

#current directoryを$PSScriptRootへ戻す
Pop-Location

Remove-Item bin -Recurse
Remove-Item src -Recurse
remove-item vbac.wsf