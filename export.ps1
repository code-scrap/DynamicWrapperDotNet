$manifest =  '<?xml version="1.0" encoding="UTF-16" standalone="yes"?> <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"> <assemblyIdentity type="win32" name="Export" version="0.0.0.0"/> <file name="Export.dll"> <comClass  description="Export Class" clsid="{185FAAFF-9A8A-41B4-809A-CA6EEAA95D61}" threadingModel="Both" progid="DynamicWrapperDotNet"/> </file>  </assembly>'
$env:tmp = 'C:\Tools\TNT'
$ax = New-Object -ComObject Microsoft.Windows.ActCtx
$ax.ManifestText = $manifest
$mdo = $ax.CreateObject("DynamicWrapperDotNet" )
$mdo.getValue3()
$mdo | gm
