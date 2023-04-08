using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace NAPS2.Wia.Native;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
#endif
internal class WiaTransferCallback2 : IWiaTransferCallback
{
    private const int WIA_TRANSFER_MSG_STATUS = 0x00001;
    private const int WIA_TRANSFER_MSG_END_OF_STREAM = 0x00002;
    private const int WIA_TRANSFER_MSG_END_OF_TRANSFER = 0x00003;
    private const int WIA_TRANSFER_MSG_DEVICE_STATUS = 0x00005;
    private const int WIA_TRANSFER_MSG_NEW_PAGE = 0x00006;

    private readonly Queue<IStream> _streams = new();
    private readonly TransferStatusCallback _statusCallback;

    internal WiaTransferCallback2(TransferStatusCallback statusCallback)
    {
        _statusCallback = statusCallback;
    }

    public int TransferCallback(int lFlags, in WiaTransferParams pWiaTransferParams)
    {
        int hr = (int) Hresult.S_OK;
        var stream = pWiaTransferParams.lMessage == WIA_TRANSFER_MSG_END_OF_STREAM
            ? _streams.Dequeue()
            : null;

        if (!_statusCallback(
                pWiaTransferParams.lMessage,
                pWiaTransferParams.lPercentComplete,
                pWiaTransferParams.ulTransferredBytes,
                (uint) pWiaTransferParams.hrErrorStatus,
                stream))
        {
            hr = 1;
        }

        if (stream != null)
        {
            Marshal.ReleaseComObject(stream);
        }

        return hr;
    }

    public void GetNextStream(
        int lFlags,
        string bstrItemName,
        string bstrFullItemName,
        out IStream ppDestination)
    {
        IStream stream = Win32.SHCreateMemStream(null, 0);
        ppDestination = stream;
        _streams.Enqueue(stream);
    }
};