new ActiveXObject('WScript.Shell').Environment('Process')('TMP') = 'C:\\Tools\\TNT';
// You could add a way to drop this dynamically 	
var manifest = '<?xml version="1.0" encoding="UTF-16" standalone="yes"?> <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"> 	<assemblyIdentity type="win32" name="Export" version="0.0.0.0"/> 	<file name="Export.dll"> 	<comClass  description="Export Class" clsid="{185FAAFF-9A8A-41B4-809A-CA6EEAA95D61}" threadingModel="Both" progid="DynamicWrapperDotNet"/> 	</file>  </assembly>';
var ax = new ActiveXObject("Microsoft.Windows.ActCtx");
ax.ManifestText = manifest;	

var mdo = ax.CreateObject("DynamicWrapperDotNet");
var s = mdo.getValue1("a");
WScript.StdOut.WriteLine(s);
var t = mdo.getValue1("b");
var s = mdo.getValue2();
mdo.getValue3();
WScript.StdOut.WriteLine(s);
WScript.StdOut.WriteLine(t);

