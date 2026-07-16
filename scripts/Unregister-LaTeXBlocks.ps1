$key='HKCU:\Software\Microsoft\Office\Word\Addins\LaTeXBlocks.Word.AddIn'
if(Test-Path $key){Remove-Item $key -Recurse -Force}
Write-Output 'LaTeX Blocks unregistered.'
