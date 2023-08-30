using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace NAPS2.Wia;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
#endif
internal class NativeStreamWrapper : Stream
{
    private readonly IStream _source;
    private readonly IStreamPointerRead _sourceReader;
    private readonly IntPtr _nativeLong;

    private long _position;

    public NativeStreamWrapper(IStream source)
    {
        _source = source;
        // ReSharper disable once SuspiciousTypeConversion.Global
        _sourceReader = (IStreamPointerRead) _source;
        _nativeLong = Marshal.AllocCoTaskMem(8);
    }

    ~NativeStreamWrapper()
    {
        Marshal.FreeCoTaskMem(_nativeLong);
    }

    public override bool CanRead => true;

    public override bool CanSeek => true;

    public override bool CanWrite => true;

    public override void Flush()
    {
        _source.Commit(0);
    }

    public override long Length
    {
        get
        {
            _source.Stat(out var stat, 1);
            return stat.cbSize;
        }
    }

    public override long Position
    {
        get => _position;
        set => Seek(value, SeekOrigin.Begin);
    }

    public override unsafe int Read(byte[] buffer, int offset, int count)
    {
        if (buffer.Length < offset + count)
        {
            throw new ArgumentException("Buffer too small");
        }
        fixed (byte* ptr = buffer)
        {
            _sourceReader.Read(ptr + offset, count, _nativeLong);
        }
        int bytesRead = Marshal.ReadInt32(_nativeLong);
        _position += bytesRead;
        return bytesRead;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        if (origin == SeekOrigin.Begin)
        {
            _position = offset;
        }
        else if (origin == SeekOrigin.Current)
        {
            _position += offset;
        }
        else
        {
            throw new NotImplementedException();
        }
        _source.Seek(offset, (int)origin, _nativeLong);
        return Marshal.ReadInt64(_nativeLong);
    }

    public override void SetLength(long value)
    {
        _source.SetSize(value);
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        if (offset != 0) throw new NotImplementedException();
        _source.Write(buffer, count, IntPtr.Zero);
        _position += count;
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        Marshal.Release(Marshal.GetIUnknownForObject(_source));
    }
}