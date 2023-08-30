using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace NAPS2.Wia.Native;

[ComImport, Guid("00000139-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IEnumSTATPROPSTG
{
    [PreserveSig]
    int Next(
        uint celt,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        STATPROPSTG[] rgelt,
        out uint pceltFetched
    );

    void Skip(uint celt);

    void Reset();

    IEnumSTATPROPSTG Clone();
}

[ComImport, Guid("0000000c-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IStreamPointerRead
{
    unsafe void Read(byte* pv, int cb, IntPtr pcbRead);
}

[ComImport, Guid("98B5E8A0-29CC-491a-AAC0-E6DB4FDCCEB6")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaPropertyStorage
{
    void ReadMultiple(
        uint cpspec,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        PROPSPEC[] rgpspec,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        PROPVARIANT[] rgpropvar
    );

    void WriteMultiple(
        uint cpspec,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        PROPSPEC[] rgpspec,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        PROPVARIANT[] rgpropvar,
        uint propidNameFirst
    );

    void DeleteMultiple(
        uint cpspec,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        PROPSPEC[] rgpspec
    );

    void ReadPropertyNames(
        uint cpropid,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        uint[] rgpropid,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.LPWStr)]
        string[] rglpwstrName
    );

    void WritePropertyNames(
        uint cpropid,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        uint[] rgpropid,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0, ArraySubType = UnmanagedType.LPWStr)]
        string[] rglpwstrName
    );

    void DeletePropertyNames(
        uint cpropid,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        uint[] rgpropid
    );

    void Commit(uint grfCommitFlags);
    void Revert();
    IEnumSTATPROPSTG? Enum();

    void SetTimes(
        in System.Runtime.InteropServices.ComTypes.FILETIME pctime,
        in System.Runtime.InteropServices.ComTypes.FILETIME patime,
        in System.Runtime.InteropServices.ComTypes.FILETIME pmtime
    );

    void SetClass(in Guid clsid);
    void Stat(out STATPROPSETSTG pstatpsstg);

    void GetPropertyAttributes(
        uint cpspec,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        PROPSPEC[] rgpspec,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        uint[] rgflags,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        PROPVARIANT[] rgpropvar
    );

    uint GetCount();
    void GetPropertyStream(out Guid pCompatibilityId, out IStream ppIStream);
    void SetPropertyStream(in Guid pCompatibilityId, IStream pIStream);
}

[ComImport, Guid("27d4eaaf-28a6-4ca5-9aab-e678168b9527")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaTransferCallback
{
    [PreserveSig]
    int TransferCallback(int lFlags, in WiaTransferParams pWiaTransferParams);

    void GetNextStream(int lFlags, string bstrItemName, string bstrFullItemName, out IStream ppDestination);
}

[ComImport, Guid("81BEFC5B-656D-44f1-B24C-D41D51B4DC81")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IEnumWIA_FORMAT_INFO
{
    [PreserveSig]
    int Next(
        uint celt,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        WIA_FORMAT_INFO[] rgelt,
        out uint pceltFetched
    );

    void Skip(uint celt);
    void Reset();
    IEnumWIA_FORMAT_INFO Clone();
    uint GetCount();
}

[ComImport, Guid("a558a866-a5b0-11d2-a08f-00c04f72dc3c")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaDataCallback
{
    [PreserveSig]
    int BandedDataCallback(
        int lMessage,
        int lStatus,
        int lPercentComplete,
        int lOffset,
        int lLength,
        int lReserved,
        int lResLength,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)]
        byte[] pbBuffer
    );
}

[ComImport, Guid("a6cef998-a5b0-11d2-a08f-00c04f72dc3c")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaDataTransfer
{
    [PreserveSig]
    int idtGetData(ref STGMEDIUM pMedium, IWiaDataCallback pIWiaDataCallback);

    [PreserveSig]
    int idtGetBandedData(in WIA_DATA_TRANSFER_INFO pWiaDataTransInfo, IWiaDataCallback pIWiaDataCallback);

    void idtQueryGetData(in WIA_FORMAT_INFO pfe);
    IEnumWIA_FORMAT_INFO idtEnumWIA_FORMAT_INFO();
    void idtGetExtendedTransferInfo(out WIA_EXTENDED_TRANSFER_INFO pExtendedTransferInfo);
};

[ComImport, Guid("c39d6942-2f4e-4d04-92fe-4ef4d3a1de5a")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaTransfer
{
    void Download(int lFlags, IWiaTransferCallback pIWiaTransferCallback);
    void Upload(int lFlags, IStream pSource, IWiaTransferCallback pIWiaTransferCallback);
    void Cancel();
    IEnumWIA_FORMAT_INFO EnumWIA_FORMAT_INFO();
}

[ComImport, Guid("59970AF4-CD0D-44d9-AB24-52295630E582")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IEnumWiaItem2
{
    [PreserveSig]
    int Next(
        uint cElt,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        IWiaItem2[] ppIWiaItem2,
        out uint pcEltFetched
    );

    void Skip(uint cElt);
    void Reset();
    IEnumWiaItem2 Clone();
    uint GetCount();
}

[ComImport, Guid("95C2B4FD-33F2-4d86-AD40-9431F0DF08F7")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaPreview
{
    void GetNewPreview(
        int lFlags,
        IWiaItem2 pWiaItem2,
        IWiaTransferCallback pWiaTransferCallback
    );

    void UpdatePreview(
        int lFlags,
        IWiaItem2 pChildWiaItem2,
        IWiaTransferCallback pWiaTransferCallback
    );

    void DetectRegions(int lFlags);
    void Clear();
}

[ComImport, Guid("1fcc4287-aca6-11d2-a093-00c04f72dc3c")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IEnumWIA_DEV_CAPS
{
    [PreserveSig]
    int Next(
        uint celt,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        WIA_DEV_CAP[] rgelt,
        out uint pceltFetched
    );

    void Skip(uint celt);
    void Reset();
    IEnumWIA_DEV_CAPS Clone();
    uint GetCount();
}

[ComImport, Guid("5e8383fc-3391-11d2-9a33-00c04fa36145")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IEnumWiaItem
{
    [PreserveSig]
    int Next(
        uint celt,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        IWiaItem[] ppIWiaItem,
        out uint pceltFetched
    );

    void Skip(uint celt);
    void Reset();
    IEnumWiaItem Clone();
    uint GetCount();
}

[ComImport, Guid("4db1ad10-3391-11d2-9a33-00c04fa36145")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaItem
{
    int GetItemType();
    void AnalyzeItem(int lFlags);
    IEnumWiaItem EnumChildItems();
    void DeleteItem(int lFlags);
    IWiaItem CreateChildItem(int lFlags, string bstrItemName, string bstrFullItemName);
    IEnumWIA_DEV_CAPS EnumRegisterEventInfo(int lFlags, in Guid pEventGUID);
    IWiaItem FindItemByName(int lFlags, string bstrFullItemName);

    [PreserveSig]
    int DeviceDlg(IntPtr hwndParent, int lFlags, int lIntent, out int plItemCount,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 3)]
        out IWiaItem[]? ppIWiaItem);

    IWiaItem DeviceCommand(int lFlags, in Guid pCmdGUID);
    IWiaItem GetRootItem();
    IEnumWIA_DEV_CAPS EnumDeviceCapabilities(int lFlags);
    string DumpItemData();
    string DumpDrvItemData();
    string DumpTreeItemData();

    void Diagnostic(
        uint ulSize,
        [Out, MarshalAs(UnmanagedType.LPArray)]
        byte[] pBuffer
    );
}

[ComImport, Guid("6CBA0075-1287-407d-9B77-CF0E030435CC")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaItem2
{
    IWiaItem2 CreateChildItem(int lItemFlags, int lCreationFlags, string bstrItemName);

    void DeleteItem(int lFlags);
    IEnumWiaItem2 EnumChildItems(IntPtr pCategoryGUID);
    IWiaItem2 FindItemByName(int lFlags, string bstrFullItemName);
    void GetItemCategory(out Guid pItemCategoryGUID);
    int GetItemType();

    void DeviceDlg(
        int lFlags,
        IntPtr hwndParent,
        string bstrFolderName,
        string bstrFilename,
        out int plNumFiles,
        [Out, MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 4)]
        string[] ppbstrFilePaths,
        ref IWiaItem2 ppItem
    );

    void DeviceCommand(int lFlags, in Guid pCmdGUID, out IWiaItem2 ppIWiaItem2);
    IEnumWIA_DEV_CAPS EnumDeviceCapabilities(int lFlags);
    bool CheckExtension(int lFlags, string bstrName, in Guid riidExtensionInterface);

    void GetExtension(int lFlags, string bstrName, in Guid riidExtensionInterface,
        [MarshalAs(UnmanagedType.IUnknown)]
        out object ppOut);

    IWiaItem2 GetParentItem();
    IWiaItem2 GetRootItem();
    IWiaPreview GetPreviewComponent(int lFlags);
    IEnumWIA_DEV_CAPS EnumRegisterEventInfo(int lFlags, in Guid pEventGUID);
    void Diagnostic(uint ulSize, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] byte[] pBuffer);
}

[ComImport, Guid("5e38b83c-8cf1-11d1-bf92-0060081ed811")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IEnumWIA_DEV_INFO
{
    [PreserveSig]
    int Next(
        uint celt,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)]
        IWiaPropertyStorage[] rgelt,
        out uint pceltFetched
    );

    void Skip(uint celt);
    void Reset();
    IEnumWIA_DEV_INFO Clone();
    uint GetCount();
}

[ComImport, Guid("ae6287b0-0084-11d2-973b-00a0c9068f2e")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaEventCallback
{
    void ImageEventCallback(
        in Guid pEventGUID,
        string bstrEventDescription,
        string bstrDeviceID,
        string bstrDeviceDescription,
        uint dwDeviceType,
        string bstrFullItemName,
        ref uint pulEventType,
        uint ulReserved
    );
}

[ComImport, Guid("5eb2502a-8cf1-11d1-bf92-0060081ed811")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaDevMgr
{
    IEnumWIA_DEV_INFO EnumDeviceInfo(int lFlag);
    IWiaItem CreateDevice(string bstrDeviceID);

    IWiaItem SelectDeviceDlg(
        IntPtr hwndParent,
        int lDeviceType,
        int lFlags,
        ref string pbstrDeviceID
    );

    string SelectDeviceDlgID(
        IntPtr hwndParent,
        int lDeviceType,
        int lFlags
    );

    void GetImageDlg(
        IntPtr hwndParent,
        int lDeviceType,
        int lFlags,
        int lIntent,
        IWiaItem? pItemRoot,
        string bstrFilename,
        ref Guid pguidFormat
    );

    void RegisterEventCallbackProgram(
        int lFlags,
        string bstrDeviceID,
        in Guid pEventGUID,
        string bstrCommandline,
        string bstrName,
        string bstrDescription,
        string bstrIcon
    );

    void RegisterEventCallbackInterface(
        int lFlags,
        string bstrDeviceID,
        in Guid pEventGUID,
        IWiaEventCallback pIWiaEventCallback,
        [MarshalAs(UnmanagedType.IUnknown)]
        out object pEventObject
    );

    void RegisterEventCallbackCLSID(
        int lFlags,
        string bstrDeviceID,
        in Guid pEventGUID,
        in Guid pClsID,
        string bstrName,
        string bstrDescription,
        string bstrIcon
    );

    void AddDeviceDlg(IntPtr hwndParent, int lFlags);
}

[ComImport, Guid("a1f4e726-8cf1-11d1-bf92-0060081ed811")]
internal class WiaDevMgr
{
}

[ComImport, Guid("79C07CF1-CBDD-41ee-8EC3-F00080CADA7A")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IWiaDevMgr2
{
    IEnumWIA_DEV_INFO EnumDeviceInfo(int lFlags);
    IWiaItem2 CreateDevice(int lFlags, string bstrDeviceID);
    IWiaItem2 SelectDeviceDlg(IntPtr hwndParent, int lDeviceType, int lFlags, ref string pbstrDeviceID);
    string SelectDeviceDlgID(IntPtr hwndParent, int lDeviceType, int lFlags);

    void RegisterEventCallbackInterface(
        int lFlags,
        string bstrDeviceID,
        in Guid pEventGUID,
        IWiaEventCallback pIWiaEventCallback,
        [MarshalAs(UnmanagedType.IUnknown)]
        out object pEventObject
    );

    void RegisterEventCallbackProgram(
        int lFlags,
        string bstrDeviceID,
        in Guid pEventGUID,
        string bstrFullAppName,
        string bstrCommandLineArg,
        string bstrName,
        string bstrDescription,
        string bstrIcon
    );

    void RegisterEventCallbackCLSID(
        int lFlags,
        string bstrDeviceID,
        in Guid pEventGUID,
        in Guid pClsID,
        string bstrName,
        string bstrDescription,
        string bstrIcon
    );

    [PreserveSig]
    int GetImageDlg(
        int lFlags,
        string bstrDeviceID,
        IntPtr hwndParent,
        string bstrFolderName,
        string bstrFilename,
        ref int plNumFiles,
        [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 5, ArraySubType = UnmanagedType.BStr)]
        ref string[] ppbstrFilePaths,
        ref IWiaItem2? ppItem
    );
}

[ComImport, Guid("B6C292BC-7C88-41ee-8B54-8EC92617E599")]
internal class WiaDevMgr2
{
};