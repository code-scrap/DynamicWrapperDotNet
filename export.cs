using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;

/*
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /target:library AllTheThings.cs
"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\ildasm.exe" AllTheThings.dll  /out=AllTheThings.il
C:\Windows\Microsoft.NET\Framework\v4.0.30319\ilasm.exe /DLL /X64 AllTheThings.il
*/



[Guid("F35D5D5D-4A3C-4042-AC35-CE0C57AF8383")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
[ComVisible(true)]
public interface IDynamicWrapperDotNet
{
	string getValue1(string sParameter);
	void getValue3();
	string getValue2();
}

//https://gist.github.com/jjeffery/1568627
[Guid("00000001-0000-0000-c000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[ComImport]
internal interface IClassFactory
{
	void CreateInstance([MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);
	void LockServer(bool fLock);
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
[Guid("185FAAFF-9A8A-41B4-809A-CA6EEAA95D61")]
[ProgId("DynamicWrapperDotNet")]
public class DynamicWrapperDotNet : MarshalByRefObject, IDynamicWrapperDotNet, IClassFactory
{
	
	
	public string getValue1(string sParameter)
	{
	 switch (sParameter)
	 {
		 case "a":
		 return "A was chosen";

		 case "b":
		 return "B was chosen";

		 case "c":
		 return "C was chosen";

		 default:
		 return "Other";
		}	
	}
	
	[ComVisible(true)]
	public string getValue2()
	{
	return "From VBS String Function";
	}

	[ComVisible(true)]
	public void getValue3()
	{
			System.Windows.Forms.MessageBox.Show("Hey From My Assembly");

	}
	
	public static uint DllGetClassObject(Guid rclsid, Guid riid, out IntPtr ppv)
	{
		// We will expport this	
		ppv = IntPtr.Zero;

		try
		{
			if (riid.CompareTo(Guid.Parse("00000001-0000-0000-c000-000000000046")) == 0)
			{
				//Call to DllClassObject is requesting IClassFactory.
				var instance = new DynamicWrapperDotNet();
				IntPtr iUnk = Marshal.GetIUnknownForObject(instance);
				//return instance;
				Marshal.QueryInterface(iUnk, ref riid, out ppv);
				return 0;
			}
			else
				return 0x80040111; //CLASS_E_CLASSNOTAVAILABLE
		}
		catch
		{
			return 0x80040111; //CLASS_E_CLASSNOTAVAILABLE
		}        
	}

	public void CreateInstance([MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject)
	{
		IntPtr ppv = IntPtr.Zero;

		//http://stackoverflow.com/a/13355702/864414
		AppDomainSetup domaininfo = new AppDomainSetup();
		domaininfo.ApplicationBase = @"C:\Tools\TNT";
		var curDomEvidence = AppDomain.CurrentDomain.Evidence;
		AppDomain newDomain = AppDomain.CreateDomain("MyDomain", curDomEvidence, domaininfo);

		Type type = typeof(DynamicWrapperDotNet);
		var instance = newDomain.CreateInstanceAndUnwrap(
			   type.Assembly.FullName,
			   type.FullName);

		ppvObject = instance;
	}

	public void LockServer(bool fLock)
	{
		//Do nothing
	}
	
	
	
}
