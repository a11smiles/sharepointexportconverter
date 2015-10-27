# SharePoint Export Converter
Converts a SharePoint site collection export to a directory/file structure based on the Manifest.xml file.

## Background
When SharePoint exports a site collection, the resulting file, while having a .cmp or .bak file extension, is nothing more than a .cab file.  The export file contains a few .xml configuration files and many .dat files.  The .dat files are the original files from the SharePoint site collection.  One of the configuration files is the Manifest.xml.  The Manifest.xml has a list of the site collection's content including a list of the files and their "conversion" names (i.e. 00000001.dat). 

## Usage
The export converter has a few options allowing you to either simply create a list of artifacts in an Excel file (NOTE: Excel must be installed for the COM libraries to be available) or actually converting the files from the export.

##### Default Usage
This assumes that the Manifest.xml along with the export files (*.dat) are in the current, working directory.
```powershell
> .\BuildExport.ps1 
```
##### Different Path for the Manifest.xml
In the case that your Manifest.xml is in a different location than the current directory.
```powershell
> .\BuildExport.ps1 -Manifest 'C:\path_to_manifest\Manifest.xml'
```
##### Write Excel File
Create an Excel file listing the contents of the Manifest.xml.  By default, the resulting Manifest.xlsx will be written to the current directory.  However, you can specify an alternate path by using the optional `-ExcelPath` parameter. 
```powershell
> .\BuildExport.ps1 -WriteExcel -ExcelPath 'C:\path_to_excel\Manifest.xslx'
```
##### Convert and Copy Files
Convert and copy all of the files from the export based on the Manifest.xml.  By default, the resulting files will be written to the current directory.  However, you can specify an alternate path by uting the optional `-OutputPath` parameter.  Additionally, if your .dat files are in a different directory than the current one, then use the `-ExportPath` parameter and specify and alternate path.
```powershell
> .\BuildExport.ps1 -CopyFiles -OutputPath 'C:\path_copy_files_to\' -ExportPath 'C:\path_to_exported_dat_files\'
```
## Warranty
This script is provided without warranty.  Please use at your own discretion.  (And, please feel free to contribute, if you'd like.)	