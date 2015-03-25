<# 

Jeff Bond
March 2015

Script for converting RMF data in excel spreadsheet to Word formatted SSP

Notes: Full word text formatting is not available in Excel and any additional formatting (lists, indentation, etc) will be added to document manually
    Based on NIST 800-53r4

    Must complete Generic SSP Workbook before running this script, instruction for that on on the first tab of the Excel document

    DO NOT USE CLIPBOARD WHILE SCRIPT IS RUNNING

#>

Function Get-FileName($initialDirectory,$title) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.title = $title
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowHelp = $false
    $OpenFileDialog.ShowDialog() | Out-Null
    
    Return $OpenFileDialog.filename
}

Function OpenExcelBook($FileName) {
    $Excel= New-Object -ComObject Excel.Application
    Return $Excel.workbooks.open($Filename)
}

Function SearchAWord($Document,$findtext,$replacewithtext) { 
    $FindReplace=$Document.ActiveWindow.Selection.Find
    $matchCase = $false;
    $matchWholeWord = $true;
    $matchWildCards = $false;
    $matchSoundsLike = $false;
    $matchAllWordForms = $false;
    $forward = $true;
    $format = $false;
    $matchKashida = $false;
    $matchDiacritics = $false;
    $matchAlefHamza = $false;
    $matchControl = $false;
    $read_only = $false;
    $visible = $true;
    $replace = 2;
    $wrap = 1;
    
    $FindReplace.Execute2007($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl) | Out-Null
}

$wdStory = 6
$wdMove = 0
$introw = 1

Write-Host "`nDO NOT USE CLIPBOARD WHILE SCRIPT IS RUNNING`n`n"

Write-Host "Select Output Word Document (Will be overwritten)"
$SSP_File = Get-FileName -initialDirectory "c:\" -title "Word Output File"

Write-Host "Select Source Excel Workbook"
$Excel_Book = Get-FileName -initialDirectory "c:\" -title "Excel Source File"

$Word = New-Object –ComObject Word.Application
$Document = $Word.Documents.Add()
$Selection = $Word.Selection
$Selection.TypeParagraph()
$Selection.TypeParagraph()

$Workbook = OpenExcelBook –Filename $Excel_Book
$WSssp = $Workbook.Worksheets.Item(5) # SSP Text Worksheet
$Worksheet = $Workbook.Worksheets.Item(2) # Control Template Worksheet
$WSnist = $Workbook.Worksheets.Item(4) # Selected NIST Control Text Worksheet

Write-Host -NoNewline "Working..."

Do {
    Write-Host -NoNewline "."

    #Paste control template
    $Template = $Worksheet.Range("A1:B2").Copy()
    $Selection.PasteSpecial()

    $Cnt_Num = $WSnist.Cells.Item($intRow, 1).Value()
    $WSnist.Cells.Item($intRow, 2).Value() | clip

    # Replace placeholders with NIST data
    SearchAWord -Document $Document -findtext '##Control_Num##' -replacewithtext $Cnt_Num
    SearchAWord -Document $Document -findtext '##Control_Text##' -replacewithtext "^c" # Paste from clipboard workaround for 255 char paste limit

    $SSP_Text = $WSssp.Cells.Item($intRow, 2).Value()
    # Paste control explanation
    $Selection.EndKey($wdStory, $wdMove) | Out-Null
    $Selection.TypeText($SSP_Text)
    $Selection.TypeParagraph()
    $Selection.TypeParagraph()

    $introw++
} While (($WSnist.Cells.Item(($intRow), 1).Value()) -ne $null)

Write-Host -NoNewline "Complete"

$Document.Saveas([REF]$SSP_File)
$Document.Close()
$Word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null