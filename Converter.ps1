PowerShell.exe -WindowStyle Hidden {
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form                            = New-Object system.Windows.Forms.Form
$form.ClientSize                 = '388,153'
$form.text                       = "Elongate"
$form.TopMost                    = $true
$form.BackColor                  = "White"
$form.FormBorderStyle            = "FixedSingle"

$fileLabel                       = New-Object system.Windows.Forms.Label
$fileLabel.text                  = "File name:"
$fileLabel.AutoSize              = $true
$fileLabel.width                 = 25
$fileLabel.height                = 10
$fileLabel.location              = New-Object System.Drawing.Point(9,10)
$fileLabel.Font                  = 'Microsoft Sans Serif,9'

$filePath                        = New-Object system.Windows.Forms.TextBox
$filePath.multiline              = $false
$filePath.width                  = 239
$filePath.height                 = 25
$filePath.location               = New-Object System.Drawing.Point(77,9)
$filePath.Font                   = 'Microsoft Sans Serif,9'

$browseFile                      = New-Object system.Windows.Forms.Button
$browseFile.text                 = "Browse"
$browseFile.width                = 60
$browseFile.height               = 24
$browseFile.BackColor            = "Transparent"
$browseFile.location             = New-Object System.Drawing.Point(322,7)
$browseFile.Font                 = 'Microsoft Sans Serif,9'

$savePathLabel                   = New-Object system.Windows.Forms.Label
$savePathLabel.text              = "Save to:"
$savePathLabel.AutoSize          = $true
$savePathLabel.width             = 25
$savePathLabel.height            = 10
$savePathLabel.location          = New-Object System.Drawing.Point(9,40)
$savePathLabel.Font              = 'Microsoft Sans Serif,9'

$savePath                        = New-Object system.Windows.Forms.TextBox
$savePath.multiline              = $false
$savePath.width                  = 239
$savePath.height                 = 20
$savePath.location               = New-Object System.Drawing.Point(77,38)
$savePath.Font                   = 'Microsoft Sans Serif,9'

$browseSavePath                  = New-Object system.Windows.Forms.Button
$browseSavePath.text             = "Browse"
$browseSavePath.width            = 60
$browseSavePath.height           = 24
$browseSavePath.BackColor        = "Transparent"
$browseSavePath.location         = New-Object System.Drawing.Point(322,36)
$browseSavePath.Font             = 'Microsoft Sans Serif,9'

$yearLabel                       = New-Object system.Windows.Forms.Label
$yearLabel.text                  = "Year:"
$yearLabel.AutoSize              = $true
$yearLabel.width                 = 25
$yearLabel.height                = 10
$yearLabel.location              = New-Object System.Drawing.Point(9,98)
$yearLabel.Font                  = 'Microsoft Sans Serif,9'

$windowLabel                     = New-Object system.Windows.Forms.Label
$windowLabel.text                = "Window:"
$windowLabel.AutoSize            = $true
$windowLabel.width               = 25
$windowLabel.height              = 10
$windowLabel.location            = New-Object System.Drawing.Point(9,127)
$windowLabel.Font                = 'Microsoft Sans Serif,9'

$testLabel                       = New-Object system.Windows.Forms.Label
$testLabel.text                  = "Test:"
$testLabel.AutoSize              = $true
$testLabel.width                 = 25
$testLabel.height                = 10
$testLabel.location              = New-Object System.Drawing.Point(9,69)
$testLabel.Font                  = 'Microsoft Sans Serif,9'

$year                            = New-Object system.Windows.Forms.TextBox
$year.multiline                  = $false
$year.width                      = 46
$year.height                     = 20
$year.location                   = New-Object System.Drawing.Point(77,96)
$year.Font                       = 'Microsoft Sans Serif,9'

$window                          = New-Object system.Windows.Forms.TextBox
$window.multiline                = $false
$window.width                    = 46
$window.height                   = 20
$window.location                 = New-Object System.Drawing.Point(77,125)
$window.Font                     = 'Microsoft Sans Serif,9'

$test                            = New-Object system.Windows.Forms.TextBox
$test.multiline                  = $false
$test.width                      = 175
$test.height                     = 20
$test.location                   = New-Object System.Drawing.Point(77,67)
$test.Font                       = 'Microsoft Sans Serif,9'

$cancel                          = New-Object system.Windows.Forms.Button
$cancel.text                     = "Cancel"
$cancel.width                    = 60
$cancel.height                   = 24
$cancel.BackColor                = "Transparent"
$cancel.location                 = New-Object System.Drawing.Point(322,123)
$cancel.Font                     = 'Microsoft Sans Serif,9'

$run                             = New-Object system.Windows.Forms.Button
$run.text                        = "Run"
$run.width                       = 60
$run.height                      = 24
$run.BackColor                   = "Transparent"
$run.location                    = New-Object System.Drawing.Point(322,66)
$run.Font                        = 'Microsoft Sans Serif,9'

$fileBrowser                     = New-Object System.Windows.Forms.OpenFileDialog
$fileBrowser.Multiselect         = $false
$fileBrowser.Filter              = 'CSV (Comma delimited)(*.csv)|*.csv'

$savePathBrowser                 = New-Object System.Windows.Forms.FolderBrowserDialog

$form.controls.AddRange(@($fileLabel,$filePath,$browseFile,$savePathLabel,$savePath,$browseSavePath,
    $yearLabel,$windowLabel,$testLabel,$test,$year,$window,$run,$cancel))

$browseFile.Add_Click({
    $fileBrowser.ShowDialog()
    $filePath.Text = $fileBrowser.FileName
})

$browseSavePath.Add_Click({
    $savePathBrowser.ShowDialog()
    $savePath.Text = $savePathBrowser.SelectedPath
})

$run.Add_Click({

    $csv = Import-Csv -Path $filePath.Text | Where-Object { $_."Email Address" -ne "" }
    $questions = New-Object System.Collections.ArrayList
    $entries = New-Object System.Collections.ArrayList
    $saveFile = $savePath.Text + "\" + [System.IO.Path]::GetFileNameWithoutExtension($filePath.Text) + "_FORMATTED.csv"

    foreach($question in $csv[0].PSObject.Properties.Name) {
        if($question -match "Question") {
            $questions.Add($question)
        }
    }

    foreach($line in $csv) {
        $sid     = $($line."Email Address").Substring(0, 10)
        $score   = $($line."Percent").Trim("%")
        $teacher = $($line."TEACHER Last Name")
        $school  = $($line."SCHOOL")

        foreach($question in $questions) {
            $answer = $($line.$question)
            
            $entries.Add(
                [PSCustomObject]@{
                    "Year"              = $year.Text
                    "Window"            = $window.Text
                    "Student ID"        = $sid
                    "School"            = $school
                    "Teacher"           = $teacher
                    "Question ID"       = $question.Trim("Question ")
                    "Question Score"    = $null
                    "Selection"         = $answer
                    "Correct"           = $null
                    "Total Score"       = $score
                    "Standard Number"   = $null
                    "Standard Category" = $null
                    "Test"              = $test.Text
                }
            )
        }
    }

    $entries | Export-Csv -NoTypeInformation -Path $saveFile -Force
    $form.Close()
})          

$form.ShowDialog()
}