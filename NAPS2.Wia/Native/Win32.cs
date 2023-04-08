using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace NAPS2.Wia.Native;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
#endif
internal class Win32
{
    [DllImport("propsys.dll", PreserveSig = true)]
    public static extern uint PropVariantGetElementCount(in PROPVARIANT propvar);

    [DllImport("propsys.dll", PreserveSig = false)]
    public static extern void PropVariantGetInt32Elem(in PROPVARIANT propvar, uint iElem, out int pnVal);

    [DllImport("propsys.dll", PreserveSig = false)]
    public static extern void PropVariantToInt32(in PROPVARIANT propvarIn, out int plRet);

    [DllImport("propsys.dll", PreserveSig = false)]
    public static extern void InitPropVariantFromInt32(int lVal, out PROPVARIANT ppropvar);

    [DllImport("propsys.dll", PreserveSig = false)]
    public static extern void PropVariantToBSTR(in PROPVARIANT propvar,
        [MarshalAs(UnmanagedType.BStr)]
        out string pbstrOut);

    [DllImport("propsys.dll", PreserveSig = false)]
    public static extern void PropVariantToInt32VectorAlloc(
        in PROPVARIANT propvar,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)]
        out int[] pprgn,
        out uint pcElem
    );

    [DllImport("ole32.dll", PreserveSig = false)]
    public static extern void PropVariantClear(in PROPVARIANT pvar);

    [DllImport("shlwapi.dll", PreserveSig = true)]
    public static extern IStream SHCreateMemStream(
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
        byte[]? pInit, uint cbInit);
}