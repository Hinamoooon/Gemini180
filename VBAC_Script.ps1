echo "VBA�\�[�X�R�[�h���o�͂��܂��B"
echo "VBAC in Ariawase v0.6.0"

cd $PSScriptRoot

#Source�t�H���_�ɓ��e�����鎞�A�m�F
if ( test-path .\Source\* ){
    $title = "*** ���s�m�F ***"
    $message = "Source�t�H���_�̓��e���㏑�����܂��B���s���Ă�낵���ł����H"

    $objYes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","���s"
    $objNo = New-Object System.Management.Automation.Host.ChoiceDescription "&No","���s�I��"
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
        


#VBAC.wsf�������Ă���
Copy-Item C:\Users\Tatsuki\Documents\Program\vbacGUI\vbac.wsf $PSScriptRoot

#�R�[�h���G�N�X�|�[�g������excel�t�@�C����bin�t�H���_�փR�s�[
New-Item bin -ItemType Directory
Copy-Item *.xlsm bin

cscript vbac.wsf decombine

$FolderName = get-childitem -Name src
echo $FolderName
$FolderPath = Join-Path .\src $FolderName
echo $FolderPath

Push-Location $FolderPath
move-Item *.bas ..\..\Source

#current directory��$PSScriptRoot�֖߂�
Pop-Location

Remove-Item bin -Recurse
Remove-Item src -Recurse
remove-item vbac.wsf