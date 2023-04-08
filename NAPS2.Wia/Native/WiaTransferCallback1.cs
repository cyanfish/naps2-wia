using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace NAPS2.Wia.Native;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
#endif
internal class WiaTransferCallback1 : IWiaDataCallback
{
    private const int IT_MSG_DATA_HEADER = 0x0001;
    private const int IT_MSG_DATA = 0x0002;
    private const int IT_MSG_STATUS = 0x0003;
    private const int IT_MSG_TERMINATION = 0x0004;
    private const int IT_MSG_NEW_PAGE = 0x0005;
    private const int IT_MSG_FILE_PREVIEW_DATA = 0x0006;
    private const int IT_MSG_FILE_PREVIEW_DATA_HEADER = 0x0007;

    private const int STREAM_SEEK_SET = 0;
    private const int STREAM_SEEK_CUR = 1;
    private const int STREAM_SEEK_END = 2;

    IStream? _stream;
    readonly TransferStatusCallback _statusCallback;

    internal WiaTransferCallback1(TransferStatusCallback statusCallback)
    {
        _statusCallback = statusCallback;
    }

    [PreserveSig]
    public int BandedDataCallback(
        int lMessage,
        int lStatus,
        int lPercentComplete,
        int lOffset,
        int lLength,
        int lReserved,
        int lResLength,
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)]
        byte[] pbBuffer
    )
    {
        uint hr = Hresult.S_OK;
        IStream? finishedStream = null;

        switch (lMessage)
        {
            case IT_MSG_DATA_HEADER:
                break;
            case IT_MSG_DATA:
                if (_stream == null)
                {
                    _stream = Win32.SHCreateMemStream(null, 0);
                    var empty_header = new byte[] { 0x42, 0x4D, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                    _stream.Write(empty_header, empty_header.Length, IntPtr.Zero);
                }
                _stream.Write(pbBuffer, lLength, IntPtr.Zero);
                break;
            case IT_MSG_STATUS:
                break;
            case IT_MSG_NEW_PAGE:
                finishedStream = _stream;
                _stream = null;
                break;
            case IT_MSG_TERMINATION:
                finishedStream = _stream;
                break;
        }

        if (finishedStream != null)
        {
            finishedStream.Seek(14, STREAM_SEEK_SET, IntPtr.Zero);
            var bytes = new byte[4];
            finishedStream.Read(bytes, 4, IntPtr.Zero);
            int header_bytes = BitConverter.ToInt32(bytes, 0);
            header_bytes += 14;
            finishedStream.Seek(2, STREAM_SEEK_SET, IntPtr.Zero);
            finishedStream.Stat(out var stat, 1);
            bytes = BitConverter.GetBytes(stat.cbSize);
            finishedStream.Write(bytes, 4, IntPtr.Zero);
            finishedStream.Seek(10, STREAM_SEEK_SET, IntPtr.Zero);
            bytes = BitConverter.GetBytes(header_bytes);
            finishedStream.Write(bytes, 4, IntPtr.Zero);
            finishedStream.Seek(0, STREAM_SEEK_SET, IntPtr.Zero);
            _statusCallback(
                2,
                lPercentComplete,
                (ulong) (lOffset + lLength),
                (uint) lStatus,
                finishedStream);
        }

        if (!_statusCallback(
                lMessage == IT_MSG_TERMINATION ? 3 : 1,
                lPercentComplete,
                (ulong) (lOffset + lLength),
                (uint) lStatus,
                null))
        {
            hr = 1;
        }

        return (int) hr;
    }
};