using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NAPS2.Wia;
using NAPS2.Wia.Native;

if (!OperatingSystem.IsWindows())
{
    Console.WriteLine("WIA COM regression tests require Windows.");
    return;
}

Assert(typeof(IWiaTransfer).GetMethod("Download")!.ReturnType == typeof(int),
    "Download must return the native HRESULT");
Assert(typeof(IWiaTransfer).GetMethod("Download")!.IsDefined(typeof(PreserveSigAttribute), false),
    "Download must preserve the native HRESULT");
CheckResult(Hresult.S_OK, true);
CheckResult(1, false);
CheckError(Hresult.E_INVALIDARG);
CheckError(Hresult.E_FAIL);
CheckFormatNegotiation(null);
CheckFormatNegotiation(new ArgumentException("MemoryBMP rejected"));
CheckFormatNegotiation(new COMException("MemoryBMP rejected", unchecked((int) Hresult.E_INVALIDARG)));
CheckSetupFailure(false);
CheckSetupFailure(true);
Console.WriteLine("PASS: COM signature, download results, MemoryBMP preference, BMP fallback, setup failures");

static void CheckFormatNegotiation(Exception? rejection)
{
    var memoryBmp = new Guid("b96b3caa-0728-11d3-9d7b-0000f81ef32e");
    var bmp = new Guid("b96b3cab-0728-11d3-9d7b-0000f81ef32e");
    var storage = new StubStorage(format =>
    {
        if (format == memoryBmp && rejection != null) throw rejection;
    });
    NativeWiaMethods.ConfigureTransfer1(storage);
    Assert(storage.Formats.SequenceEqual(rejection == null ? new[] { memoryBmp } : new[] { memoryBmp, bmp }),
        "Setup must prefer MemoryBMP and retry BMP only if MemoryBMP was rejected");
}

static void CheckSetupFailure(bool programmingError)
{
    Exception failure = programmingError ? new InvalidOperationException("Not a driver format rejection")
        : new COMException("Both formats rejected", unchecked((int) Hresult.E_INVALIDARG));
    var storage = new StubStorage(_ => throw failure);
    try
    {
        NativeWiaMethods.ConfigureTransfer1(storage);
    }
    catch (Exception ex) when (ReferenceEquals(ex, failure))
    {
        Assert(storage.Formats.Count == (programmingError ? 1 : 2),
            "Only driver setup failures should retry, and no more than once");
        return;
    }
    throw new Exception("Setup failure must be propagated");
}

static void CheckResult(uint result, bool expected)
{
    using var transfer = CreateTransfer(result);
    Assert(transfer.Download() == expected, $"Unexpected result for HRESULT {result:X8}");
}

static void CheckError(uint result)
{
    using var transfer = CreateTransfer(result);
    try
    {
        transfer.Download();
    }
    catch (WiaException ex) when (ex.ErrorCode == result)
    {
        return;
    }
    throw new Exception($"Expected WiaException with HRESULT {result:X8}");
}

static WiaTransfer CreateTransfer(uint result) => new(WiaVersion.Wia20,
    Marshal.GetIUnknownForObject(new StubTransfer(unchecked((int) result))));

static void Assert(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

[ComVisible(true), ClassInterface(ClassInterfaceType.None)]
internal class StubTransfer : IWiaTransfer
{
    private readonly int _result;

    public StubTransfer(int result) => _result = result;

    public int Download(int flags, IWiaTransferCallback callback) => _result;
    public void Upload(int flags, IStream source, IWiaTransferCallback callback) => throw new NotSupportedException();
    public void Cancel() => throw new NotSupportedException();
    public IEnumWIA_FORMAT_INFO EnumWIA_FORMAT_INFO() => throw new NotSupportedException();
}

internal class StubStorage : IWiaPropertyStorage
{
    private readonly Action<Guid> _write;
    public List<Guid> Formats { get; } = new();

    public StubStorage(Action<Guid> write) => _write = write;

    public void WriteMultiple(uint count, PROPSPEC[] specs, PROPVARIANT[] values, uint first)
    {
        if (count != 2 || first != 4098 || specs[0].ulKind != 1 || specs[1].ulKind != 1 ||
            specs[0].unionmember != new IntPtr(4106) || specs[1].unionmember != new IntPtr(4108) ||
            values[0].varType != (ushort) VarEnum.VT_CLSID ||
            values[1].varType != (ushort) VarEnum.VT_I4 || values[1].intVal != 128)
            throw new InvalidOperationException("Expected an atomic FORMAT + TYMED_CALLBACK write");
        var format = Marshal.PtrToStructure<Guid>(values[0].pclsidVal);
        Formats.Add(format);
        _write(format);
    }

    public void ReadMultiple(uint count, PROPSPEC[] specs, PROPVARIANT[] values) => throw new NotSupportedException();
    public void DeleteMultiple(uint count, PROPSPEC[] specs) => throw new NotSupportedException();
    public void ReadPropertyNames(uint count, uint[] ids, string[] names) => throw new NotSupportedException();
    public void WritePropertyNames(uint count, uint[] ids, string[] names) => throw new NotSupportedException();
    public void DeletePropertyNames(uint count, uint[] ids) => throw new NotSupportedException();
    public void Commit(uint flags) => throw new NotSupportedException();
    public void Revert() => throw new NotSupportedException();
    public IEnumSTATPROPSTG? Enum() => throw new NotSupportedException();
    public void SetTimes(in FILETIME created, in FILETIME accessed, in FILETIME modified) => throw new NotSupportedException();
    public void SetClass(in Guid id) => throw new NotSupportedException();
    public void Stat(out STATPROPSETSTG stat) => throw new NotSupportedException();
    public void GetPropertyAttributes(uint count, PROPSPEC[] specs, uint[] flags, PROPVARIANT[] values) => throw new NotSupportedException();
    public uint GetCount() => throw new NotSupportedException();
    public void GetPropertyStream(out Guid id, out IStream stream) => throw new NotSupportedException();
    public void SetPropertyStream(in Guid id, IStream stream) => throw new NotSupportedException();
}
