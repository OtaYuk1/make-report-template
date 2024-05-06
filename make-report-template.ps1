Add-Type -AssemblyName System.Windows.Forms

#フォームのオブジェクトを作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "Input Form"    #フォームのタイトル
$form.Size = New-Object System.Drawing.Size(300,300)    #フォームのウィンドウサイズ
$form.StartPosition = "CenterScreen"    #画面中央に表示

#タイトルの入力欄
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10,20)
$titleLabel.Size = New-Object System.Drawing.Size(280,20)
$titleLabel.Text = "Title:"
$form.Controls.Add($titleLabel)

$titleTextBox = New-Object System.Windows.Forms.TextBox
$titleTextBox.Location = New-Object System.Drawing.Point(10,40)
$titleTextBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($titleTextBox)

#学生番号の入力欄
$numberLabel = New-Object System.Windows.Forms.Label
$numberLabel.Location = New-Object System.Drawing.Point(10,70)
$numberLabel.Size = New-Object System.Drawing.Size(280,20)
$numberLabel.Text = "Number:"
$form.Controls.Add($numberLabel)

$numberTextBox = New-Object System.Windows.Forms.TextBox
$numberTextBox.Location = New-Object System.Drawing.Point(10,90)
$numberTextBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($numberTextBox)

#名前の入力欄
$authorLabel = New-Object System.Windows.Forms.Label
$authorLabel.Location = New-Object System.Drawing.Point(10,120)
$authorLabel.Size = New-Object System.Drawing.Size(280,20)
$authorLabel.Text = "Author:"
$form.Controls.Add($authorLabel)

$authorTextBox = New-Object System.Windows.Forms.TextBox
$authorTextBox.Location = New-Object System.Drawing.Point(10,140)
$authorTextBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($authorTextBox)

#実行ボタン
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(85,190)
$button.Size = New-Object System.Drawing.Size(120,30)
$button.Text = "Run"

$button.Add_Click({
    #フォームの入力値を代入
    $title = $titleTextBox.Text
    $number = $numberTextBox.Text
    $author = $authorTextBox.Text   

    #ファイル名に使用できない文字の場合が入力されていないかの確認
    #ファイル名に使用できない文字を含む配列を取得
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $invalidCharFound = $false

    #各入力文字列の検証
    $title, $number, $author | ForEach-Object {
        $_.ToCharArray() | ForEach-Object {
            if ($invalidChars -contains $_) {
                $invalidCharFound = $true
            }
        }
    }
    #使用できない文字が含まれていた場合，エラーメッセージを表示，スクリプトの実行を終了
    if ($invalidCharFound) {
        [System.Windows.Forms.MessageBox]::Show("Error: Invalid characters in input", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return  
    }   
    # Wordがインストールされていない場合，エラーメッセージを表示，スクリプトの実行を終了
    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: Microsoft Word is not installed on this system.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return  
    }

    #新しいドキュメントを作成
    $doc = $word.Documents.Add()

    #フッターにページ番号を挿入
    $footer = $doc.Sections.Item(1).Footers.Item(1).Range
    $footer.Fields.Add($footer, 33) 
    $footer.ParagraphFormat.Alignment = 1    #中央揃え

    #ヘッダーを設定
    $header = $doc.Sections.Item(1).Headers.Item(1).Range
    $header.ParagraphFormat.Alignment = 2    #右詰め
    $header.Text = "$title`n$number  $author"

    $fileName = "${title}_${number}_${author}.docx"
    $filePath = Join-Path -Path $pwd -ChildPath $fileName

    #ファイルが既に存在する場合，エラーメッセージを表示，スクリプトの実行を終了
    if (Test-Path $filePath) {
        [System.Windows.Forms.MessageBox]::Show("Error: Duplicate filename", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $word.Quit()
        return
    }
    #ファイルを保存,閉じる
    $doc.SaveAs([string] $filePath) 
    $doc.Close()

    #Wordを終了
    $word.Quit()

    Write-Host "File created: $filePath"

    #フォームを閉じる
    $form.Close()
})
$form.Controls.Add($button)

#フォーム表示
$form.ShowDialog()
