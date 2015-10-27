Param(
    [string]$Manifest = (Get-Item -Path ".\" -Verbose).FullName + '\Manifest.xml',
    [string]$OutputPath = (Get-Item -Path ".\" -Verbose).FullName + '\',
    [string]$ExcelPath = (Get-Item -Path ".\" -Verbose).FullName + '\Manifest.xlsx',
    [string]$ExportPath = (Get-Item -Path ".\" -Verbose).FullName + '\',
    [switch]$WriteExcel,
    [switch]$CopyFiles
)

function WriteSite([System.Xml.XmlElement]$spObject, [__ComObject]$sheet, [int]$rowCount, [bool]$writeExcel, [bool]$writeFile, [string]$outputPath) 
{
    if($writeExcel)
    {
        $sheet.Cells.Item($rowCount, 1) = $spObject.ObjectType;
        $sheet.Cells.Item($rowCount, 2) = $spObject.Id;    
    }
}

function WriteWeb([System.Xml.XmlElement]$spObject, [__ComObject]$sheet, [int]$rowCount, [bool]$writeExcel, [bool]$writeFile, [string]$outputPath) 
{
    if($writeExcel)
    {
        $sheet.Cells.Item($rowCount, 1) = $spObject.ObjectType;
        $sheet.Cells.Item($rowCount, 2) = $spObject.Id;    
        $sheet.Cells.Item($rowCount, 3) = $spObject.ParentId;    
        $sheet.Cells.Item($rowCount, 4) = $spObject.ParentWebId;    
        $sheet.Cells.Item($rowCount, 5) = $spObject.Web.Title;
        $sheet.Cells.Item($rowCount, 6) = $spObject.Url;
    }

    if($writeFile)
    {
        New-Item -ItemType Directory -Force -Path ($outputPath + $spObject.Url) | Out-Null;
    }
}

function WriteFolder([System.Xml.XmlElement]$spObject, [__ComObject]$sheet, [int]$rowCount, [bool]$writeExcel, [bool]$writeFile, [string]$outputPath) 
{
    if($writeExcel)
    {
        $sheet.Cells.Item($rowCount, 1) = $spObject.ObjectType;
        $sheet.Cells.Item($rowCount, 2) = $spObject.Id;    
        $sheet.Cells.Item($rowCount, 3) = $spObject.ParentId;    
        $sheet.Cells.Item($rowCount, 4) = $spObject.ParentWebId;    
        $sheet.Cells.Item($rowCount, 5) = $spObject.Folder.Url;
        $sheet.Cells.Item($rowCount, 6) = $spObject.Url;
    }

    if($writeFile)
    {
        New-Item -ItemType Directory -Force -Path ($outputPath + $spObject.Url) | Out-Null;
    }
}

function WriteFile([System.Xml.XmlElement]$spObject, [__ComObject]$sheet, [int]$rowCount, [bool]$writeExcel, [bool]$writeFile, [string]$outputPath, [string]$exportPath) 
{
    if($writeExcel)
    {
        $sheet.Cells.Item($rowCount, 1) = $spObject.ObjectType;
        $sheet.Cells.Item($rowCount, 2) = $spObject.Id;    
        $sheet.Cells.Item($rowCount, 3) = $spObject.ParentId;    
        $sheet.Cells.Item($rowCount, 4) = $spObject.ParentWebId;    
        $sheet.Cells.Item($rowCount, 5) = $spObject.File.Name;
        $sheet.Cells.Item($rowCount, 6) = $spObject.Url;
        $sheet.Cells.Item($rowCount, 7) = $spObject.File.FileValue;
    }

    if ($writeFile)
    {
        if ((Test-Path ($outputPath + $spObject.Url)) -eq $False) 
        {
            New-Item -ItemType File -Path ($outputPath + $spObject.Url) -Force | Out-Null;
        } 

        Copy-Item -Path ($exportPath + $spObject.File.FileValue) -Destination ($outputPath + $spObject.Url) -Force | Out-Null;
    }
}

if ($WriteExcel)
{
    # Configure Excel
    $excel = new-Object -comobject Excel.Application;
    $excel.visible = $False;

    # Create Workbook
    $book = $excel.Workbooks.Add();
    $sheet = $book.Worksheets.Item(1);

    # Create Header
    $sheet.Cells.Item(1,1) = "Object Type";
    $sheet.Cells.Item(1,2) = "Id";
    $sheet.Cells.Item(1,3) = "Parent Id";
    $sheet.Cells.Item(1,4) = "Parent Web Id";
    $sheet.Cells.Item(1,5) = "Name";
    $sheet.Cells.Item(1,6) = "Url";
    $sheet.Cells.Item(1,7) = "Export File";

    # Create Table in Excel
    $table = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes);
    $table.Name = "TableData";
    $table.TableStyle = "TableStyleMedium6";
}

# Configure XML Manifest
Write-Progress -Activity "Processing..." -Status "Reading Manifest" -PercentComplete 0
[xml]$xmlDoc = Get-Content $Manifest;
Write-Progress -Activity "Processing..." -Status "Reading Manifest" -PercentComplete 100

# Set the row count to the first row after header
$rowCount = 2;

# Process Site Collections
Write-Progress -Activity "Processing..." -Status "Site Collection" -PercentComplete 0;
$nodes = @($xmlDoc.SPObjects.SPObject | Where {$_.ObjectType -eq 'SPSite'});
$nodeCount = 0;
foreach($spObject in $nodes)
{
    WriteSite -spObject $spObject -sheet $sheet -rowCount $rowCount -writeExcel $WriteExcel;
    $rowCount++;
    $nodeCount++;
    Write-Progress -Activity "Processing..." -Status "Site Collection" -PercentComplete ($nodeCount/$nodes.Count*100);
}

# Process Sites
Write-Progress -Activity "Processing..." -Status "Sites" -PercentComplete 0;
$nodes = @($xmlDoc.SPObjects.SPObject | Where {$_.ObjectType -eq 'SPWeb'});
$nodeCount = 0;
foreach($spObject in $nodes)
{
    WriteWeb -spObject $spObject -sheet $sheet -rowCount $rowCount -writeExcel $WriteExcel -writeFile $CopyFiles -outputPath $OutputPath;
    $rowCount++;
    $nodeCount++;
    Write-Progress -Activity "Processing..." -Status "Sites" -PercentComplete ($nodeCount/$nodes.Count*100);
}

# Process Folders
Write-Progress -Activity "Processing..." -Status "Folders" -PercentComplete 0;
$nodes = @($xmlDoc.SPObjects.SPObject | Where {$_.ObjectType -eq 'SPFolder'});
$nodeCount = 0;
foreach($spObject in $nodes)
{
    WriteFolder -spObject $spObject -sheet $sheet -rowCount $rowCount -writeExcel $WriteExcel -writeFile $CopyFiles -outputPath $OutputPath;
    $rowCount++;
    $nodeCount++;
    Write-Progress -Activity "Processing..." -Status "Folders" -PercentComplete ($nodeCount/$nodes.Count*100);
}

# Process Files
Write-Progress -Activity "Processing..." -Status "Files" -PercentComplete 0;
$nodes = @($xmlDoc.SPObjects.SPObject | Where {$_.ObjectType -eq 'SPFile'});
$nodeCount = 0;
foreach($spObject in $nodes)
{
    WriteFile -spObject $spObject -sheet $sheet -rowCount $rowCount -writeExcel $WriteExcel -writeFile $CopyFiles -outputPath $OutputPath -exportPath $ExportPath;
    $rowCount++;
    $nodeCount++;
    Write-Progress -Activity "Processing..." -Status "Files" -PercentComplete ($nodeCount/$nodes.Count*100);
}

if ($WriteExcel)
{
    # Save and Close Workbook
    $excel.ActiveWorkbook.SaveAs($ExcelPath);
    $book.Close();
    $excel.Quit();
}

[GC]::Collect();


