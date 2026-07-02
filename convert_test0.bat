@echo off
echo Starting conversion process...
echo.

echo Converting to Word (DOCX)...
bin\convert_md.exe test0.md word test0.docx

echo.
echo Converting to HTML...
bin\convert_md.exe test0.md html test0.html

echo.
echo Converting to PDF...
bin\convert_md.exe test0.md pdf test0.pdf

echo.
echo Conversion complete!
