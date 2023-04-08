namespace NAPS2.Wia;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
#endif
public class WiaDeviceInfo : NativeWiaObject, IWiaDeviceProps
{
    protected internal WiaDeviceInfo(WiaVersion version, IntPtr propStorageHandle) : base(version)
    {
        Properties = new WiaPropertyCollection(version, propStorageHandle);
    }

    public WiaPropertyCollection Properties { get; }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing)
        {
            Properties?.Dispose();
        }
    }
}