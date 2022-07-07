# tnsnamesParser
Code to parse an Oracle tnsnames.ora file

The script is provided as is without any error checking. Feel free to modify it.

Usage:
```powershell
Parse-Tnsnames -PathToTnsnamesFile <Full Path to the Tnsnames.ora file> -RegexFolder <Full Path to the folder containing the text files with the regex patterns> -RegexFilePrefix <Prefix for the regex text files> -OutputFolder <Folder to store the output Excel Files>
```
> Requires ImportExcel module. Get it here: https://github.com/dfinke/ImportExcel
