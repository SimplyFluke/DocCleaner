[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | out-null

# Function to prevent a powershell window from popping up when running the script.
Function Hide-PowerShellWindow()
		{
			[CmdletBinding()]
			param (
				[IntPtr]$Handle=$(Get-Process -id $PID).MainWindowHandle
			)
			$WindowDisplay = @"
			using System;
			using System.Runtime.InteropServices;
			namespace Window
			{
				public class Display
				{
					[DllImport("user32.dll")]
					private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
					public static bool Hide(IntPtr hWnd)
					{
					return ShowWindowAsync(hWnd, 0);
					}
				}
			}
"@
			Try
			{
			Add-Type -TypeDefinition $WindowDisplay
			[Window.Display]::Hide($Handle)
			}
			Catch
			{
			}
		}
		#EndRegion
[Void]$(Hide-PowerShellWindow)

function CleanFile($file)
{
    # Open Word-doc in a hidden window
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($file)

    # Select all text and delete it
    $word.Selection.WholeStory()
    $word.Selection.Delete()

    # Save and close the document
    $doc.Save()
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
}

#Create UI
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Document Cleaner'
$main_form.Width = 350
$main_form.Height = 130

$main_form.MinimumSize = New-Object System.Drawing.Size(350,150)
$main_form.MaximumSize = New-Object System.Drawing.Size(350,150)

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Select document to clean:"
$Label.Font = [System.Drawing.Font]::new("Verdana", 10)
$Label.Width = 200
$Label.Location  = New-Object System.Drawing.Point(10,12)
$main_form.Controls.Add($Label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Width = 200
$textBox.Text = "brrrrr"
$textBox.ReadOnly = $true
$main_form.Controls.Add($textBox)

$fileSelectButton = New-Object System.Windows.Forms.Button
$fileSelectButton.Location = New-Object System.Drawing.Point(220,40)
$fileSelectButton.Text = "Select"
$main_form.Controls.Add($fileSelectButton)

$fileCleanButton = New-Object System.Windows.Forms.Button
$fileCleanButton.Location = New-Object System.Drawing.Point(20,80)
$fileCleanButton.Text = "Clean!"
$main_form.Controls.Add($fileCleanButton)


$fileSelectButton.Add_Click({
    $oneDrivePath = [System.Environment]::GetFolderPath("UserProfile") + "\OneDrive" # Set default path to the OneDrive-folder when opening filedialog
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ # Limit filetypes to docx and xlsx
        InitialDirectory = $oneDrivePath
        Filter           = 'Documents (*.docx)|*.docx|SpreadSheet (*.xlsx)|*.xlsx'
    }
    
    if ($FileBrowser.ShowDialog() -eq "OK"){
        $textBox.Text = $FileBrowser.FileName # Chuck filename in the textbox field once a file is selected
    }
})

$fileCleanButton.Add_Click({
    CleanFile ($textBox.Text)
})


$main_form.ShowDialog()
