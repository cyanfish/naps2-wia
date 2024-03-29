﻿namespace NAPS2.Wia;

#if NET6_0_OR_GREATER
[System.Runtime.Versioning.SupportedOSPlatform("windows")]
#endif
public class WiaDevice : WiaItemBase, IWiaDeviceProps
{
    protected internal WiaDevice(WiaVersion version, IntPtr handle) : base(version, handle)
    {
    }

    public WiaItem? PromptToConfigure(IntPtr parentWindowHandle = default)
    {
        if (Version == WiaVersion.Wia20)
        {
            throw new InvalidOperationException("WIA 2.0 does not support PromptToConfigure. Use WiaDeviceManager.PromptForImage if you want to use the native WIA 2.0 UI.");
        }

        var hr = NativeWiaMethods.ConfigureDevice1(Handle, parentWindowHandle, 0, 0, out int itemCount, out IntPtr[]? items);
        if (hr == 1)
        {
            return null;
        }
        WiaException.Check(hr);
        if (items == null)
        {
            return null;
        }
        return new WiaItem(Version, items[0]);
    }
}