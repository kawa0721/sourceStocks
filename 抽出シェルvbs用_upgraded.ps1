#$targetFolder = "C:\Users\川喜田将之\Desktop\test"
#$inputfile = ".\testdata.txt"
$outPath = ".\testfol\"

$i=0

#$itemList = Get-ChildItem $targetFolder;
#$fileCount = (Get-ChildItem -Recurse $targetFolder | Where-Object { ! $_.PsIsContainer }).Count
$p=1

#$inputfile = $item.FullName
$i=0

Write-Host '実行中'
echo $Args

$CSV = Get-Content $Args | ConvertFrom-CSV -header `
ログ出力時刻, `
ログレベル, `
トランザクションID, `
ロギング種別, `
出力ホスト名, `
インタフェースID, `
プロセスID, `
URL, `
要求元IPアドレス, `
要求元キー, `
要求元ユーザ／装置, `
セッションID, `
原トランザクションID, `
ログ順序番号, `
基点時刻, `
処理時間ミリ秒, `
処理結果, `
HTTPステータス, `
エラー通番, `
エラーコード, `
サブエラーコード, `
エラーメッセージID, `
要求入力データ, `
取引テーブル更新内容, `
会員属性変更履歴テーブル更新内容, `
操作ログテーブル更新内容, `
応答出力データ `
    -Delimiter "`t"

$CSV =$($CSV | Where-Object {$_.ロギング種別 -eq 'end'}) | Select-Object インタフェースID, 基点時刻, 処理結果, 要求元IPアドレス, 要求入力データ, 応答出力データ, @{Name='共通会員IDハッシュ';Expression={''}}, @{Name='ウォレットシステムID';Expression={''}}, @{Name='端末識別番号';Expression={''}}

foreach($csvdata in $CSV) {
    $csvdata.ウォレットシステムID = $($csvdata |Select-String -Pattern "wid:\d{16}" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value.substring($_.Value.length -16, 16)})
    If($csvdata.ウォレットシステムID.Length -eq 0){
        $csvdata.ウォレットシステムID = 'NGのため出力無し'
    }
    $csvdata.端末識別番号 = $($csvdata |Select-String -Pattern "tid:[A-Za-z0-9._%-]{32}" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value.substring($_.Value.length -32, 32)})
    $csvdata.共通会員IDハッシュ=$($csvdata |Select-String -Pattern "omnilog:[A-Za-z0-9._%-]+_[A-Za-z0-9]{128}" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value.substring($_.Value.length -128, 128)})
    $csvdata
} 

#$CSV | Export-Csv -Encoding UTF8 -NoTypeInformation -Force -Path ".\テスト\エクスポートtest2.csv"

$filename = $($Args |Select-String -Pattern '\b[a-zA-Z]+\d+\.txt$' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value.substring(0,$_.Value.length -4) })

$db = "C:\Users\川喜田将之\Desktop\testfol\" + $filename + ".accdb"
 
If(-not(Test-Path -Path $db))
{
    #DBファイルが存在しない場合作成する
    $application = New-Object -ComObject Access.Application
    $application.NewCurrentDataBase($db,12)
    $application.CloseCurrentDataBase()
    $application.Quit()
 
    #テーブル作成
    $connection = New-Object -ComObject ADODB.Connection
    $connection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source = $db")
 
    $table = "テーブル1"
    $fields = "共通会員IDハッシュ LongCHAR, ウォレットシステムID Text, 端末識別番号 Text, インターフェースID TEXT, 処理日時 Text, 要求元IPアドレス Text, 実行結果 Text"
    $command = "Create Table $table `($fields`)"
    $connection.Execute($command)
    $connection.Close()  
}

$connection = New-Object -ComObject ADODB.Connection
$connection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$db")
$recordset = New-Object -ComObject ADODB.Connection
$sql = "SELECT * FROM テーブル1"
$recordset.open($sql, $connection, 3, 3)
$connection.BeginTrans()

foreach($data in $CSV) {

    $a = $data.共通会員IDハッシュ
    $b = $data.ウォレットシステムID
    $c = $data.端末識別番号
    $d = $data.インタフェースID
    $e = $data.基点時刻
    $f = $data.要求元IPアドレス
    $g = $data.処理結果

    $recordset.AddNew()
    $recordset.Fields.Item("共通会員IDハッシュ").value = $a
    $recordset.Fields.Item("共通会員IDハッシュ").value = $b
    $recordset.Fields.Item("共通会員IDハッシュ").value = $c
    $recordset.Fields.Item("共通会員IDハッシュ").value = $d
    $recordset.update()

    # $connection = New-Object -ComObject ADODB.Connection
    # $connection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$db")

    # $insCmd = "Insert into テーブル1 `(共通会員IDハッシュ, ウォレットシステムID, 端末識別番号, インターフェースID, 処理日時, 要求元IPアドレス, 実行結果`) `
    # Values `('$a', '$b', '$c', '$d', '$e', '$f', '$g'`)"
    # $connection.Execute($insCmd)
}

$connection.CommitTrans()
$connection.Close()  

echo $Args
echo $filename

#[int]$progress = $p / [int]$fileCount * 100
#[string]$joinedStr = -join('進捗率:' , [string]$progress , '%')
#Write-Host $joinedStr 
#$p ++



