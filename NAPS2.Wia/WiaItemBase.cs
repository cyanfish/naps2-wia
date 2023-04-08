namespace NAPS2.Wia;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows7.0")]
#endif
public class WiaItemBase : NativeWiaObject
{
    private WiaPropertyCollection? _properties;

    protected internal WiaItemBase(WiaVersion version, IntPtr handle) : base(version, handle)
    {
    }

    public WiaPropertyCollection Properties
    {
        get
        {
            if (_properties == null)
            {
                _properties = new WiaPropertyCollection(Version, Handle);
            }
            return _properties;
        }
    }

    public List<WiaItem> GetSubItems()
    {
        var items = new List<WiaItem>();
        WiaException.Check(Version == WiaVersion.Wia10
            ? NativeWiaMethods.EnumerateItems1(Handle, itemHandle => items.Add(new WiaItem(Version, itemHandle)))
            : NativeWiaMethods.EnumerateItems2(Handle, itemHandle => items.Add(new WiaItem(Version, itemHandle))));
        return items;
    }

    public WiaItem? FindSubItem(string name)
    {
        return GetSubItems().FirstOrDefault(x => x.Name() == name);
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing)
        {
            _properties?.Dispose();
        }
    }
}