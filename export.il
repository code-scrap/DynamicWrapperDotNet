
//  Microsoft (R) .NET Framework IL Disassembler.  Version 4.8.3928.0
//  Copyright (c) Microsoft Corporation.  All rights reserved.



// Metadata version: v4.0.30319
.assembly extern mscorlib
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly extern System.Windows.Forms
{
  .publickeytoken = (B7 7A 5C 56 19 34 E0 89 )                         // .z\V.4..
  .ver 4:0:0:0
}
.assembly export
{
  .custom instance void [mscorlib]System.Runtime.CompilerServices.RuntimeCompatibilityAttribute::.ctor() = ( 01 00 01 00 54 02 16 57 72 61 70 4E 6F 6E 45 78   // ....T..WrapNonEx
                                                                                                             63 65 70 74 69 6F 6E 54 68 72 6F 77 73 01 )       // ceptionThrows.
  .custom instance void [mscorlib]System.Runtime.CompilerServices.CompilationRelaxationsAttribute::.ctor(int32) = ( 01 00 08 00 00 00 00 00 ) 
  .hash algorithm 0x00008004
  .ver 0:0:0:0
}
.module export.dll
// MVID: {6E521C71-3623-4805-AB4F-995DFDB3CCF7}
.imagebase 0x10000000
.file alignment 0x00000200
.stackreserve 0x00100000
.subsystem 0x0003       // WINDOWS_CUI
.corflags 0x00000001    //  ILONLY
// Image base: 0x04D30000


// =============== CLASS MEMBERS DECLARATION ===================

.class interface public abstract auto ansi IDynamicWrapperDotNet
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComVisibleAttribute::.ctor(bool) = ( 01 00 01 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 02 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 46 33 35 44 35 44 35 44 2D 34 41 33 43   // ..$F35D5D5D-4A3C
                                                                                                  2D 34 30 34 32 2D 41 43 33 35 2D 43 45 30 43 35   // -4042-AC35-CE0C5
                                                                                                  37 41 46 38 33 38 33 00 00 )                      // 7AF8383..
  .method public hidebysig newslot abstract virtual 
          instance string  getValue1(string sParameter) cil managed
  {
  } // end of method IDynamicWrapperDotNet::getValue1

  .method public hidebysig newslot abstract virtual 
          instance void  getValue3() cil managed
  {
  } // end of method IDynamicWrapperDotNet::getValue3

  .method public hidebysig newslot abstract virtual 
          instance string  getValue2() cil managed
  {
  } // end of method IDynamicWrapperDotNet::getValue2

} // end of class IDynamicWrapperDotNet

.class interface private abstract auto ansi import IClassFactory
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 30 30 30 30 30 30 30 31 2D 30 30 30 30   // ..$00000001-0000
                                                                                                  2D 30 30 30 30 2D 63 30 30 30 2D 30 30 30 30 30   // -0000-c000-00000
                                                                                                  30 30 30 30 30 34 36 00 00 )                      // 0000046..
  .custom instance void [mscorlib]System.Runtime.InteropServices.InterfaceTypeAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ComInterfaceType) = ( 01 00 01 00 00 00 00 00 ) 
  .method public hidebysig newslot abstract virtual 
          instance void  CreateInstance(object  marshal( iunknown ) pUnkOuter,
                                        valuetype [mscorlib]System.Guid& riid,
                                        [out] object&  marshal( iunknown ) ppvObject) cil managed
  {
  } // end of method IClassFactory::CreateInstance

  .method public hidebysig newslot abstract virtual 
          instance void  LockServer(bool fLock) cil managed
  {
  } // end of method IClassFactory::LockServer

} // end of class IClassFactory

.class public auto ansi beforefieldinit DynamicWrapperDotNet
       extends [mscorlib]System.MarshalByRefObject
       implements IDynamicWrapperDotNet,
                  IClassFactory
{
  .custom instance void [mscorlib]System.Runtime.InteropServices.ClassInterfaceAttribute::.ctor(valuetype [mscorlib]System.Runtime.InteropServices.ClassInterfaceType) = ( 01 00 00 00 00 00 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.ComVisibleAttribute::.ctor(bool) = ( 01 00 01 00 00 ) 
  .custom instance void [mscorlib]System.Runtime.InteropServices.GuidAttribute::.ctor(string) = ( 01 00 24 31 38 35 46 41 41 46 46 2D 39 41 38 41   // ..$185FAAFF-9A8A
                                                                                                  2D 34 31 42 34 2D 38 30 39 41 2D 43 41 36 45 45   // -41B4-809A-CA6EE
                                                                                                  41 41 39 35 44 36 31 00 00 )                      // AA95D61..
  .custom instance void [mscorlib]System.Runtime.InteropServices.ProgIdAttribute::.ctor(string) = ( 01 00 14 44 79 6E 61 6D 69 63 57 72 61 70 70 65   // ...DynamicWrappe
                                                                                                    72 44 6F 74 4E 65 74 00 00 )                      // rDotNet..
  .method public hidebysig newslot virtual final 
          instance string  getValue1(string sParameter) cil managed
  {
    // Code size       81 (0x51)
    .maxstack  2
    .locals init (string V_0,
             string V_1)
    IL_0000:  nop
    IL_0001:  ldarg.1
    IL_0002:  stloc.1
    IL_0003:  ldloc.1
    IL_0004:  brfalse.s  IL_0047

    IL_0006:  ldloc.1
    IL_0007:  ldstr      "a"
    IL_000c:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                   string)
    IL_0011:  brtrue.s   IL_002f

    IL_0013:  ldloc.1
    IL_0014:  ldstr      "b"
    IL_0019:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                   string)
    IL_001e:  brtrue.s   IL_0037

    IL_0020:  ldloc.1
    IL_0021:  ldstr      "c"
    IL_0026:  call       bool [mscorlib]System.String::op_Equality(string,
                                                                   string)
    IL_002b:  brtrue.s   IL_003f

    IL_002d:  br.s       IL_0047

    IL_002f:  ldstr      "A was chosen"
    IL_0034:  stloc.0
    IL_0035:  br.s       IL_004f

    IL_0037:  ldstr      "B was chosen"
    IL_003c:  stloc.0
    IL_003d:  br.s       IL_004f

    IL_003f:  ldstr      "C was chosen"
    IL_0044:  stloc.0
    IL_0045:  br.s       IL_004f

    IL_0047:  ldstr      "Other"
    IL_004c:  stloc.0
    IL_004d:  br.s       IL_004f

    IL_004f:  ldloc.0
    IL_0050:  ret
  } // end of method DynamicWrapperDotNet::getValue1

  .method public hidebysig newslot virtual final 
          instance string  getValue2() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.ComVisibleAttribute::.ctor(bool) = ( 01 00 01 00 00 ) 
    // Code size       11 (0xb)
    .maxstack  1
    .locals init (string V_0)
    IL_0000:  nop
    IL_0001:  ldstr      "From VBS String Function"
    IL_0006:  stloc.0
    IL_0007:  br.s       IL_0009

    IL_0009:  ldloc.0
    IL_000a:  ret
  } // end of method DynamicWrapperDotNet::getValue2

  .method public hidebysig newslot virtual final 
          instance void  getValue3() cil managed
  {
    .custom instance void [mscorlib]System.Runtime.InteropServices.ComVisibleAttribute::.ctor(bool) = ( 01 00 01 00 00 ) 
    // Code size       13 (0xd)
    .maxstack  8
    IL_0000:  nop
    IL_0001:  ldstr      "Hey From My Assembly"
    IL_0006:  call       valuetype [System.Windows.Forms]System.Windows.Forms.DialogResult [System.Windows.Forms]System.Windows.Forms.MessageBox::Show(string)
    IL_000b:  pop
    IL_000c:  ret
  } // end of method DynamicWrapperDotNet::getValue3

  .method public hidebysig static uint32 
          DllGetClassObject(valuetype [mscorlib]System.Guid rclsid,
                            valuetype [mscorlib]System.Guid riid,
                            [out] native int& ppv) cil managed
  {
    // Code size       89 (0x59)
	.export[0]
    .maxstack  3
    .locals init (class DynamicWrapperDotNet V_0,
             native int V_1,
             uint32 V_2,
             bool V_3)
    IL_0000:  nop
    IL_0001:  ldarg.2
    IL_0002:  ldsfld     native int [mscorlib]System.IntPtr::Zero
    IL_0007:  stobj      [mscorlib]System.IntPtr
    .try
    {
      IL_000c:  nop
      IL_000d:  ldarga.s   riid
      IL_000f:  ldstr      "00000001-0000-0000-c000-000000000046"
      IL_0014:  call       valuetype [mscorlib]System.Guid [mscorlib]System.Guid::Parse(string)
      IL_0019:  call       instance int32 [mscorlib]System.Guid::CompareTo(valuetype [mscorlib]System.Guid)
      IL_001e:  ldc.i4.0
      IL_001f:  ceq
      IL_0021:  ldc.i4.0
      IL_0022:  ceq
      IL_0024:  stloc.3
      IL_0025:  ldloc.3
      IL_0026:  brtrue.s   IL_0044

      IL_0028:  nop
      IL_0029:  newobj     instance void DynamicWrapperDotNet::.ctor()
      IL_002e:  stloc.0
      IL_002f:  ldloc.0
      IL_0030:  call       native int [mscorlib]System.Runtime.InteropServices.Marshal::GetIUnknownForObject(object)
      IL_0035:  stloc.1
      IL_0036:  ldloc.1
      IL_0037:  ldarga.s   riid
      IL_0039:  ldarg.2
      IL_003a:  call       int32 [mscorlib]System.Runtime.InteropServices.Marshal::QueryInterface(native int,
                                                                                                  valuetype [mscorlib]System.Guid&,
                                                                                                  native int&)
      IL_003f:  pop
      IL_0040:  ldc.i4.0
      IL_0041:  stloc.2
      IL_0042:  leave.s    IL_0056

      IL_0044:  ldc.i4     0x80040111
      IL_0049:  stloc.2
      IL_004a:  leave.s    IL_0056

    }  // end .try
    catch [mscorlib]System.Object 
    {
      IL_004c:  pop
      IL_004d:  nop
      IL_004e:  ldc.i4     0x80040111
      IL_0053:  stloc.2
      IL_0054:  leave.s    IL_0056

    }  // end handler
    IL_0056:  nop
    IL_0057:  ldloc.2
    IL_0058:  ret
  } // end of method DynamicWrapperDotNet::DllGetClassObject

  .method public hidebysig newslot virtual final 
          instance void  CreateInstance(object  marshal( iunknown ) pUnkOuter,
                                        valuetype [mscorlib]System.Guid& riid,
                                        [out] object&  marshal( iunknown ) ppvObject) cil managed
  {
    // Code size       93 (0x5d)
    .maxstack  3
    .locals init (native int V_0,
             class [mscorlib]System.AppDomainSetup V_1,
             class [mscorlib]System.Security.Policy.Evidence V_2,
             class [mscorlib]System.AppDomain V_3,
             class [mscorlib]System.Type V_4,
             object V_5)
    IL_0000:  nop
    IL_0001:  ldsfld     native int [mscorlib]System.IntPtr::Zero
    IL_0006:  stloc.0
    IL_0007:  newobj     instance void [mscorlib]System.AppDomainSetup::.ctor()
    IL_000c:  stloc.1
    IL_000d:  ldloc.1
    IL_000e:  ldstr      "C:\\Tools\\TNT"
    IL_0013:  callvirt   instance void [mscorlib]System.AppDomainSetup::set_ApplicationBase(string)
    IL_0018:  nop
    IL_0019:  call       class [mscorlib]System.AppDomain [mscorlib]System.AppDomain::get_CurrentDomain()
    IL_001e:  callvirt   instance class [mscorlib]System.Security.Policy.Evidence [mscorlib]System.AppDomain::get_Evidence()
    IL_0023:  stloc.2
    IL_0024:  ldstr      "MyDomain"
    IL_0029:  ldloc.2
    IL_002a:  ldloc.1
    IL_002b:  call       class [mscorlib]System.AppDomain [mscorlib]System.AppDomain::CreateDomain(string,
                                                                                                   class [mscorlib]System.Security.Policy.Evidence,
                                                                                                   class [mscorlib]System.AppDomainSetup)
    IL_0030:  stloc.3
    IL_0031:  ldtoken    DynamicWrapperDotNet
    IL_0036:  call       class [mscorlib]System.Type [mscorlib]System.Type::GetTypeFromHandle(valuetype [mscorlib]System.RuntimeTypeHandle)
    IL_003b:  stloc.s    V_4
    IL_003d:  ldloc.3
    IL_003e:  ldloc.s    V_4
    IL_0040:  callvirt   instance class [mscorlib]System.Reflection.Assembly [mscorlib]System.Type::get_Assembly()
    IL_0045:  callvirt   instance string [mscorlib]System.Reflection.Assembly::get_FullName()
    IL_004a:  ldloc.s    V_4
    IL_004c:  callvirt   instance string [mscorlib]System.Type::get_FullName()
    IL_0051:  callvirt   instance object [mscorlib]System.AppDomain::CreateInstanceAndUnwrap(string,
                                                                                             string)
    IL_0056:  stloc.s    V_5
    IL_0058:  ldarg.3
    IL_0059:  ldloc.s    V_5
    IL_005b:  stind.ref
    IL_005c:  ret
  } // end of method DynamicWrapperDotNet::CreateInstance

  .method public hidebysig newslot virtual final 
          instance void  LockServer(bool fLock) cil managed
  {
    // Code size       2 (0x2)
    .maxstack  8
    IL_0000:  nop
    IL_0001:  ret
  } // end of method DynamicWrapperDotNet::LockServer

  .method public hidebysig specialname rtspecialname 
          instance void  .ctor() cil managed
  {
    // Code size       7 (0x7)
    .maxstack  8
    IL_0000:  ldarg.0
    IL_0001:  call       instance void [mscorlib]System.MarshalByRefObject::.ctor()
    IL_0006:  ret
  } // end of method DynamicWrapperDotNet::.ctor

} // end of class DynamicWrapperDotNet


// =============================================================

// *********** DISASSEMBLY COMPLETE ***********************
// WARNING: Created Win32 resource file Export.res
