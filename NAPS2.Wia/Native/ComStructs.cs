using System.Runtime.InteropServices;

namespace NAPS2.Wia.Native;

[StructLayout(LayoutKind.Sequential)]
internal struct PROPSPEC
{
    public uint ulKind;
    public IntPtr unionmember;
}

[StructLayout(LayoutKind.Sequential)]
internal struct STATPROPSTG
{
    [MarshalAs(UnmanagedType.LPWStr)]
    public string lpwstrName;
    public uint propid;
    public ushort vt;
}

[StructLayout(LayoutKind.Sequential)]
internal struct STATPROPSETSTG
{
    public Guid fmtid;
    public Guid clsid;
    public uint grfFlags;
    public System.Runtime.InteropServices.ComTypes.FILETIME mtime;
    public System.Runtime.InteropServices.ComTypes.FILETIME ctime;
    public System.Runtime.InteropServices.ComTypes.FILETIME atime;
    public uint dwOSVersion;
}

[StructLayout(LayoutKind.Sequential)]
internal struct WIA_FORMAT_INFO
{
    public Guid guidFormatID;
    public int lTymed;
}

[StructLayout(LayoutKind.Sequential)]
internal struct WIA_DEV_CAP
{
    public Guid guid;
    public uint ulFlags;
    [MarshalAs(UnmanagedType.BStr)]
    public string bstrName;
    [MarshalAs(UnmanagedType.BStr)]
    public string bstrDescription;
    [MarshalAs(UnmanagedType.BStr)]
    public string bstrIcon;
    [MarshalAs(UnmanagedType.BStr)]
    public string bstrCommandline;
}

[StructLayout(LayoutKind.Sequential)]
internal struct WiaTransferParams
{
    public int lMessage;
    public int lPercentComplete;
    public ulong ulTransferredBytes;
    public int hrErrorStatus;
}

[StructLayout(LayoutKind.Sequential)]
internal struct WIA_DATA_TRANSFER_INFO
{
    public uint ulSize;
    public uint ulSection;
    public uint ulBufferSize;
    [MarshalAs(UnmanagedType.Bool)]
    public bool bDoubleBuffer;
    public uint ulReserved1;
    public uint ulReserved2;
    public uint ulReserved3;
}

[StructLayout(LayoutKind.Sequential)]
internal struct WIA_EXTENDED_TRANSFER_INFO
{
    public uint ulSize;
    public uint ulMinBufferSize;
    public uint ulOptimalBufferSize;
    public uint ulMaxBufferSize;
    public uint ulNumBuffers;
}