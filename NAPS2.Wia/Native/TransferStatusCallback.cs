using System.Runtime.InteropServices.ComTypes;

namespace NAPS2.Wia.Native;

internal delegate bool TransferStatusCallback(
    int msgType, int percent, ulong bytesTransferred, uint hresult, IStream? stream);