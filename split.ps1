$targetFolder = "C:\Users\���c���V\Desktop\�e�X�g"
#$inputfile = ".\testdata.txt"
$outPath = ".\testfol\"
$extension = ".txt"

$i=0

$itemList = Get-ChildItem $targetFolder;
$fileCount = (Get-ChildItem -Recurse $targetFolder | Where-Object { ! $_.PsIsContainer }).Count
$p=1
foreach($item in $itemList) {
	$inputfile = $item.FullName
	$i=0
    [int]$progress = $p / [int]$fileCount * 100
    [string]$joinedStr = -join('�i����:' , [string]$progress , '%')
    Write-Host $joinedStr 
    Get-Content -ReadCount 20 $inputfile -Encoding UTF8 | ForEach-Object {

        $outfile = $outPath + $item.Name.substring(0, $item.Name.length - 4)  + "_" + + $i + $extension
        $i ++
        # ��������̕ҏW���� ��ŕҏW���@������������
        Out-File $outfile -Encoding UTF8 -InputObject $_
    }

    $p ++
} 