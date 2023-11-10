@echo off
setlocal enabledelayedexpansion

:: Define source and destination directories
set "sourceDir=1. Site Info"
set "drawingsDir=1. Site Info\1. Drawings"
set "emailsDir=1. Site Info\3. Emails"

:: Create directories if they don't exist
if not exist "1. Site Info\1. Drawings" mkdir "1. Site Info\1. Drawings"
if not exist "1. Site Info\3. Emails" mkdir "1. Site Info\3. Emails"
if not exist "2. Reports" mkdir "2. Reports"

:: Move PDF files to "1. Site Info\1. Drawings"
for %%F in (*.pdf) do (
    move "%%F" "1. Site Info\1. Drawings\"
)

:: Move Excel files to "1. Site Info\1. Drawings"
for %%F in (*.xlsx *.csv *.xls) do (
    move "%%F" "1. Site Info\1. Drawings\"
)

:: Move MSG files to "1. Site Info\3. Emails"
for %%F in (*.msg) do (
    move "%%F" "1. Site Info\3. Emails\"
)

:: Move RSD files to "2. Reports"
for %%F in (*.rsd) do (
    move "%%F" "2. Reports\"
)

:: Loop through DOCX files containing "PRD" or "EMEG" in their filenames
for %%F in (*PRD*.docx *EMEG*.docx) do (
    :: Move each matching file to "\2. Reports\"
    move "%%F" "2. Reports\"
    echo Moved "%%F" to "2. Reports\"
)

:: Loop through PNG files containing "MicrosoftTeams" in their filenames
for %%F in (*image*.png *MicrosoftTeams*.png) do (
    :: Extract the file extension (in case it's not .png)
    set "ext=%%~xF"

    :: Rename the file to "nearmap" with a unique identifier
    set "newName=nearmap_%%~nF!ext!"
    ren "%%F" "!newName!"

    :: Move the renamed file to "1. Site Info\"
    move "!newName!" "1. Site Info\"
    echo Renamed and moved "%%F" to "1. Site Info\!newName!"
)


:: Check if any files were moved
if exist "1. Site Info\*.png" (
    echo PNG files containing "MicrosoftTeams" have been moved.
) else (
    echo No PNG files containing "MicrosoftTeams" found in the current directory.
)


:: Check if "2. Photos" folder exists in the current directory
if exist "2. Photos" (
    :: Move the folder to "1. Site Info\"
    move "2. Photos" "1. Site Info\"
    echo "2. Photos" folder moved successfully to "1. Site Info\"
) else (
    echo "2. Photos" folder does not exist in the current directory.
)


:: Move PDF files to "\1. Site Info\1. Drawings"
move "%sourceDir%\*.pdf" "%drawingsDir%"

:: Move MSG files to "\1. Site Info\3. Emails"
move "%sourceDir%\*.msg" "%emailsDir%"


echo Files moved successfully.

endlocal
