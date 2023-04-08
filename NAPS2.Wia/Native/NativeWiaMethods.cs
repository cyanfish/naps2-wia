using System.Runtime.InteropServices;

namespace NAPS2.Wia.Native;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
#endif
internal static class NativeWiaMethods
{
    private static IntPtr ToPtr(object? comObject)
    {
        if (comObject == null) return IntPtr.Zero;
        return Marshal.GetIUnknownForObject(comObject);
    }

    private static object ToObjectNonNull(IntPtr ptr)
    {
        if (ptr == IntPtr.Zero) throw new ArgumentException();
        return Marshal.GetObjectForIUnknown(ptr);
    }

    private static object? ToObject(IntPtr ptr)
    {
        if (ptr == IntPtr.Zero) return null;
        return Marshal.GetObjectForIUnknown(ptr);
    }

    public static uint GetDeviceManager1(out IntPtr deviceManager)
    {
        try
        {
            deviceManager = ToPtr((IWiaDevMgr) new WiaDevMgr());
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            deviceManager = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static uint GetDeviceManager2(out IntPtr deviceManager)
    {
        try
        {
            deviceManager = ToPtr((IWiaDevMgr2) new WiaDevMgr2());
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            deviceManager = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static uint GetDevice1(IntPtr deviceManagerPtr, string deviceId, out IntPtr device)
    {
        var deviceManager = (IWiaDevMgr) ToObjectNonNull(deviceManagerPtr);
        try
        {
            device = ToPtr(deviceManager.CreateDevice(deviceId));
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            device = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static uint GetDevice2(IntPtr deviceManagerPtr, string deviceId, out IntPtr device)
    {
        var deviceManager = (IWiaDevMgr2) ToObjectNonNull(deviceManagerPtr);
        try
        {
            device = ToPtr(deviceManager.CreateDevice(0, deviceId));
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            device = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static uint EnumerateDevices1(IntPtr deviceManagerPtr, Action<IntPtr> callback)
    {
        var deviceManager = (IWiaDevMgr) ToObjectNonNull(deviceManagerPtr);
        const int WIA_DEVINFO_ENUM_LOCAL = 0x00000010;

        try
        {
            IEnumWIA_DEV_INFO enumDevInfo = deviceManager.EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL);
            var hr = Hresult.S_OK;
            do
            {
                var propStorage = new IWiaPropertyStorage[1];
                hr = (uint) enumDevInfo.Next(1, propStorage, out var fetched);
                if (hr == Hresult.S_OK)
                {
                    if (fetched != 1)
                        break;

                    callback(ToPtr(propStorage[0]));
                }
            } while (hr == Hresult.S_OK);

            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint EnumerateDevices2(IntPtr deviceManagerPtr, Action<IntPtr> callback)
    {
        var deviceManager = (IWiaDevMgr2) ToObjectNonNull(deviceManagerPtr);
        const int WIA_DEVINFO_ENUM_LOCAL = 0x00000010;

        try
        {
            IEnumWIA_DEV_INFO enumDevInfo = deviceManager.EnumDeviceInfo(WIA_DEVINFO_ENUM_LOCAL);
            var hr = Hresult.S_OK;
            do
            {
                var propStorage = new IWiaPropertyStorage[1];
                hr = (uint) enumDevInfo.Next(1, propStorage, out var fetched);
                if (hr == Hresult.S_OK)
                {
                    if (fetched != 1)
                        break;

                    callback(ToPtr(propStorage[0]));
                }
            } while (hr == Hresult.S_OK);

            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint EnumerateItems1(IntPtr itemPtr, Action<IntPtr> childCallback)
    {
        var item = (IWiaItem) ToObjectNonNull(itemPtr);
        const int WiaItemTypeFolder = 0x00000004;
        const int WiaItemTypeHasAttachments = 0x00008000;

        try
        {
            int itemType = item.GetItemType();
            if (((itemType & WiaItemTypeFolder) == WiaItemTypeFolder) ||
                ((itemType & WiaItemTypeHasAttachments) == WiaItemTypeHasAttachments))
            {
                IEnumWiaItem enumerator = item.EnumChildItems();
                if (enumerator != null)
                {
                    uint hr;
                    do
                    {
                        var items = new IWiaItem[1];
                        hr = (uint) enumerator.Next(1, items, out var fetched);
                        if (hr == Hresult.S_OK)
                        {
                            if (fetched != 1)
                                break;

                            childCallback(ToPtr(items[0]));
                        }
                    } while (hr == Hresult.S_OK);
                }
            }

            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint EnumerateItems2(IntPtr itemPtr, Action<IntPtr> childCallback)
    {
        var item = (IWiaItem2) ToObjectNonNull(itemPtr);
        const int WiaItemTypeFolder = 0x00000004;
        const int WiaItemTypeHasAttachments = 0x00008000;

        try
        {
            int itemType = item.GetItemType();
            if (((itemType & WiaItemTypeFolder) == WiaItemTypeFolder) ||
                ((itemType & WiaItemTypeHasAttachments) == WiaItemTypeHasAttachments))
            {
                IEnumWiaItem2 enumerator = item.EnumChildItems(IntPtr.Zero);
                if (enumerator != null)
                {
                    uint hr;
                    do
                    {
                        var items = new IWiaItem2[1];
                        hr = (uint) enumerator.Next(1, items, out var fetched);
                        if (hr == Hresult.S_OK)
                        {
                            if (fetched != 1)
                                break;

                            childCallback(ToPtr(items[0]));
                        }
                    } while (hr == Hresult.S_OK);
                }
            }

            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint GetItemPropertyStorage(IntPtr item, out IntPtr propStorage)
    {
        try
        {
            propStorage = ToPtr(item);
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            propStorage = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public delegate void EnumPropertyCallback(int propId, [MarshalAs(UnmanagedType.LPWStr)] string propName,
        ushort propType);

    public static uint EnumerateProperties(IntPtr propStoragePtr, EnumPropertyCallback func)
    {
        var propStorage = (IWiaPropertyStorage) ToObjectNonNull(propStoragePtr);
        try
        {
            IEnumSTATPROPSTG? enumProps = propStorage.Enum();
            if (enumProps != null)
            {
                uint hr;
                do
                {
                    var props = new STATPROPSTG[1];
                    hr = (uint) enumProps.Next(1, props, out var fetched);
                    if (hr == Hresult.S_OK)
                    {
                        if (fetched != 1)
                            break;

                        func((int) props[0].propid, props[0].lpwstrName, props[0].vt);
                    }
                } while (hr == Hresult.S_OK);
            }
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint GetPropertyBstr(IntPtr propStoragePtr, int propId, out string value)
    {
        var propStorage = (IWiaPropertyStorage) ToObjectNonNull(propStoragePtr);
        var propSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint) propId) } };
        var propVariant = new PROPVARIANT[1];
        try
        {
            try
            {
                propStorage.ReadMultiple(1, propSpec, propVariant);
                Win32.PropVariantToBSTR(propVariant[0], out value);

                return Hresult.S_OK;
            }
            finally
            {
                Win32.PropVariantClear(propVariant[0]);
            }
        }
        catch (COMException ex)
        {
            value = String.Empty;
            return (uint) ex.HResult;
        }
    }

    public static uint GetPropertyInt(IntPtr propStoragePtr, int propId, out int value)
    {
        var propStorage = (IWiaPropertyStorage) ToObjectNonNull(propStoragePtr);
        var propSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint) propId) } };
        var propVariant = new PROPVARIANT[1];
        try
        {
            try
            {
                propStorage.ReadMultiple(1, propSpec, propVariant);
                Win32.PropVariantToInt32(propVariant[0], out value);

                return Hresult.S_OK;
            }
            finally
            {
                Win32.PropVariantClear(propVariant[0]);
            }
        }
        catch (COMException ex)
        {
            value = 0;
            return (uint) ex.HResult;
        }
    }

    public static uint SetPropertyInt(IntPtr propStoragePtr, int propId, int value)
    {
        var propStorage = (IWiaPropertyStorage) ToObjectNonNull(propStoragePtr);
        const int WIA_IPA_FIRST = 4098;

        try
        {
            var propSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint) propId) } };
            var propVariants = new PROPVARIANT[1];
            try
            {
                propVariants[0].Init(value);
                propStorage.WriteMultiple(1, propSpec, propVariants, WIA_IPA_FIRST);
            }
            finally
            {
                Win32.PropVariantClear(propVariants[0]);
            }
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint GetPropertyAttributes(
        IntPtr propStoragePtr,
        int propId,
        out int flags,
        out int min,
        out int nom,
        out int max,
        out int step,
        out int numElems,
        out int[]? elems)
    {
        var propStorage = (IWiaPropertyStorage) ToObjectNonNull(propStoragePtr);
        const uint WIA_PROP_RANGE = 0x10;
        const uint WIA_PROP_LIST = 0x20;

        const uint WIA_RANGE_MIN = 0;
        const uint WIA_RANGE_NOM = 1;
        const uint WIA_RANGE_MAX = 2;
        const uint WIA_RANGE_STEP = 3;

        const uint WIA_LIST_NOM = 1;

        var propSpec = new PROPSPEC[1] { new PROPSPEC { ulKind = 1, unionmember = new IntPtr((uint) propId) } };
        var propFlags = new uint[1];
        var propVariant = new PROPVARIANT[1];

        elems = null;
        flags = 0;
        max = 0;
        min = 0;
        nom = 0;
        numElems = 0;
        step = 0;
        try
        {
            try
            {
                propStorage.GetPropertyAttributes(1, propSpec, propFlags, propVariant);

                flags = (int) propFlags[0];
                if ((flags & WIA_PROP_RANGE) == WIA_PROP_RANGE)
                {
                    Win32.PropVariantGetInt32Elem(propVariant[0], WIA_RANGE_MIN, out min);
                    Win32.PropVariantGetInt32Elem(propVariant[0], WIA_RANGE_NOM, out nom);
                    Win32.PropVariantGetInt32Elem(propVariant[0], WIA_RANGE_MAX, out max);
                    Win32.PropVariantGetInt32Elem(propVariant[0], WIA_RANGE_STEP, out step);
                }
                if ((flags & WIA_PROP_LIST) == WIA_PROP_LIST)
                {
                    numElems = (int) Win32.PropVariantGetElementCount(propVariant[0]);
                    Win32.PropVariantGetInt32Elem(propVariant[0], WIA_LIST_NOM, out nom);
                    Win32.PropVariantToInt32VectorAlloc(propVariant[0], out elems, out _);
                }
            }
            finally
            {
                Win32.PropVariantClear(propVariant[0]);
            }
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint StartTransfer1(IntPtr itemPtr, out IntPtr transfer)
    {
        var item = (IWiaItem) ToObjectNonNull(itemPtr);
        const uint WIA_IPA_FIRST = 4098;
        const uint WIA_IPA_FORMAT = 4106;
        const uint WIA_IPA_TYMED = 4108;

        const int TYMED_CALLBACK = 128;

        var WiaImgFmt_BMP = new Guid(0xb96b3cab, 0x0728, 0x11d3, 0x9d, 0x7b, 0x00, 0x00, 0xf8, 0x1e, 0xf3, 0x2e);

        try
        {
            var propSpec = new PROPSPEC[2]
            {
                new PROPSPEC { ulKind = 1, unionmember = new IntPtr(WIA_IPA_FORMAT) },
                new PROPSPEC { ulKind = 1, unionmember = new IntPtr(WIA_IPA_TYMED) },
            };
            var propVariant = new PROPVARIANT[2];
            propVariant[0].Init(WiaImgFmt_BMP);
            propVariant[1].Init(TYMED_CALLBACK);

            var propStorage = (IWiaPropertyStorage) item;
            propStorage.WriteMultiple(2, propSpec, propVariant, WIA_IPA_FIRST);
            transfer = ToPtr((IWiaDataTransfer) item);
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            transfer = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static uint StartTransfer2(IntPtr itemPtr, out IntPtr transfer)
    {
        var item = (IWiaItem2) ToObjectNonNull(itemPtr);
        try
        {
            transfer = ToPtr((IWiaTransfer) item);
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            transfer = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static unsafe uint Download1(IntPtr transferPtr, TransferStatusCallback func)
    {
        var transfer = (IWiaDataTransfer) ToObjectNonNull(transferPtr);
        var callbackClass = new WiaTransferCallback1(func);
        /*STGMEDIUM StgMedium = { 0 };
        StgMedium.tymed = TYMED_ISTREAM;
        StgMedium.pstm*/
        var wiaDataTransferInfo = new WIA_DATA_TRANSFER_INFO
        {
            ulSize = (uint) sizeof(WIA_DATA_TRANSFER_INFO),
            ulBufferSize = 131072 * 2,
            bDoubleBuffer = true
        };
        return (uint) transfer.idtGetBandedData(wiaDataTransferInfo, callbackClass);
    }

    public static uint Download2(IntPtr transferPtr, TransferStatusCallback func)
    {
        var transfer = (IWiaTransfer) ToObjectNonNull(transferPtr);
        try
        {
            var callbackClass = new WiaTransferCallback2(func);
            transfer.Download(0, callbackClass);
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint SelectDevice1(IntPtr deviceManagerPtr, IntPtr hwnd, int deviceType, int flags,
        out string deviceId, out IntPtr device)
    {
        var deviceManager = (IWiaDevMgr) ToObjectNonNull(deviceManagerPtr);
        try
        {
            deviceId = String.Empty;
            device = ToPtr(deviceManager.SelectDeviceDlg(hwnd, deviceType, flags,
                ref deviceId));
            return device == IntPtr.Zero ? 1 : Hresult.S_OK;
        }
        catch (COMException ex)
        {
            deviceId = String.Empty;
            device = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static uint SelectDevice2(IntPtr deviceManagerPtr, IntPtr hwnd, int deviceType, int flags,
        out string deviceId, out IntPtr device)
    {
        var deviceManager = (IWiaDevMgr2) ToObjectNonNull(deviceManagerPtr);
        try
        {
            deviceId = String.Empty;
            device = ToPtr(deviceManager.SelectDeviceDlg(hwnd, deviceType, flags,
                ref deviceId));
            return device == IntPtr.Zero ? 1 : Hresult.S_OK;
        }
        catch (COMException ex)
        {
            deviceId = String.Empty;
            device = IntPtr.Zero;
            return (uint) ex.HResult;
        }
    }

    public static uint GetImage1(IntPtr deviceManagerPtr, IntPtr hwnd, int deviceType, int flags, int intent,
        [MarshalAs(UnmanagedType.BStr)]
        string filePath, IntPtr itemPtr)
    {
        var deviceManager = (IWiaDevMgr) ToObjectNonNull(deviceManagerPtr);
        var item = (IWiaItem?) ToObject(itemPtr);
        var format = new Guid(0xb96b3cae, 0x0728, 0x11d3, 0x9d, 0x7b, 0x00, 0x00, 0xf8, 0x1e, 0xf3, 0x2e);
        try
        {
            deviceManager.GetImageDlg(hwnd, deviceType, flags, intent, item, filePath, format);
            return Hresult.S_OK;
        }
        catch (COMException ex)
        {
            return (uint) ex.HResult;
        }
    }

    public static uint GetImage2(
        IntPtr deviceManagerPtr,
        int flags,
        string deviceId,
        IntPtr hwnd,
        string folder,
        string fileName, ref int numFiles,
        ref string[] filePaths,
        ref IntPtr itemPtr)
    {
        var deviceManager = (IWiaDevMgr2) ToObjectNonNull(deviceManagerPtr);
        var item = (IWiaItem2?) ToObject(itemPtr);
        var result = (uint) deviceManager.GetImageDlg(flags, deviceId, hwnd, folder, fileName, ref numFiles,
            ref filePaths, ref item);
        itemPtr = ToPtr(item);
        return result;
    }

    public static uint ConfigureDevice1(IntPtr devicePtr, IntPtr hwnd, int flags, int intent, out int itemCount,
        out IntPtr[]? itemPtrs)
    {
        var device = (IWiaItem) ToObjectNonNull(devicePtr);
        var result = (uint) device.DeviceDlg(hwnd, flags, intent, out itemCount, out IWiaItem[]? items);
        itemPtrs = items?.Select(ToPtr).ToArray();
        return result;
    }
}