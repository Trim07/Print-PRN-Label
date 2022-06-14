


Function File ($InitialDirectory)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Please Select File"
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.filter = “All files (*.*) | *.*”

    If ($OpenFileDialog.ShowDialog() -eq "OK") 
    {
        If((Split-Path -Path $OpenFileDialog.FileName -Leaf).Split(".")[1] -ne "prn")
        {
            PopUp 'Arquivo incorreto, selecione o arquivo com extensão .prn!' "error"
        }else
        {
            $textBox.Text = $OpenFileDialog.FileName
        }
        
    }
}


Function PrintLabel ()
{
    Param
        (
            [Parameter()] [string] $file,
            [Parameter()] [string] $printer
        )

    if(([string]::IsNullOrEmpty($file)) -or ([string]::IsNullOrEmpty($printer)))
    {
        PopUp 'Preencha os campos corretamente!' "error"
    }
    else
    {
        cmd.exe /c copy /B $file \\$(hostname)\$printer
        PopUp 'Arquivo enviado para a impressão!' "info"
    }
}


Function PopUp ()
{
    Param
    (
        [Parameter(Mandatory = $true)] [string] $msg,
        [Parameter(Mandatory = $true)] [string] $type
    )
    $wshell = New-Object -ComObject Wscript.Shell
    if($type -eq "info"){$wshell.Popup($msg,0,"Impressão de etiquetas",0x0 + 0x40)}
    if($type -eq "error"){$wshell.Popup($msg,0,"Impressão de etiquetas",0x0 + 0x10)}
    if($type -eq "question"){$wshell.Popup($msg,0,"Impressão de etiquetas",0x0 + 0x20)}
    if($type -eq "exclamation"){$wshell.Popup($msg,0,"Impressão de etiquetas",0x0 + 0x30)}
    
}


Add-Type -assembly System.Windows.Forms

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = "Impressão de etiquetas"

$main_form.Width = 550
$main_form.Height = 150
$main_form.AutoSize = $true

#form-1

$label = New-Object System.Windows.Forms.Label
$label.Text = 'Arquivo de etiqueta'
$label.Location = New-Object System.Drawing.Point(10,20)
$label.AutoSize = $true

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(130,20)
$textBox.Size = New-Object System.Drawing.Size(260,20)

$selectFileButton = New-Object System.Windows.Forms.Button
$selectFileButton.Location = New-Object System.Drawing.Size(400,20)
$selectFileButton.Size = New-Object System.Drawing.Size(120,20)
$selectFileButton.Text = "Selecionar arquivo"
$selectFileButton.Add_Click({File('C:\')})

$printerLabelButton = New-Object System.Windows.Forms.Label
$printerLabelButton.Location = New-Object System.Drawing.Size(10,60)
$printerLabelButton.Text = "Selecione a impressora"
$printerLabelButton.AutoSize = $true

$printerList = New-Object System.Windows.Forms.ComboBox
$printerList.Size = New-Object System.Drawing.Size(245,20)
foreach($printer in Get-Printer -ComputerName $(hostname))
{
    [void]$printerList.Items.Add($printer.Name);
}
$printerList.Location = New-Object System.Drawing.Point(145,60)

$printLabelButton = New-Object System.Windows.Forms.Button
$printLabelButton.Location = New-Object System.Drawing.Size(400,60)
$printLabelButton.Size = New-Object System.Drawing.Size(120,20)
$printLabelButton.Text = "Imprimir"
$printLabelButton.Add_Click({PrintLabel $textBox.Text $printerList.Text})


$main_form.Controls.Add($label)
$main_form.Controls.Add($textBox)
$main_form.Controls.Add($selectFileButton)
$main_form.Controls.Add($printerLabelButton)
$main_form.Controls.Add($printerList)
$main_form.Controls.Add($printLabelButton)

# run app
[void]$main_form.ShowDialog()
