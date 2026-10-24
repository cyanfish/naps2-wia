# NAPS2.Wia

[![NuGet](https://img.shields.io/nuget/v/NAPS2.Wia)](https://www.nuget.org/packages/NAPS2.Wia/)

NAPS2.Wia is a standalone component that acts as a low-level wrapper around [Windows Image Acquisition (WIA)](https://docs.microsoft.com/en-us/windows/win32/wia/-wia-startpage).

Compared to the COM-based [wiaaut.dll](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wiaaut/-wiaaut-startpage), you get:
- WIA 2.0 support
- Better feeder compatibility
- Idiomatic .NET interface

If you're looking for a higher-level and easier-to-use scanning interface, check out [NAPS2.Sdk](https://github.com/cyanfish/naps2/tree/master/NAPS2.Sdk).

## Example

```c#
using var deviceManager = new WiaDeviceManager();

// Prompt the user to select a scanner
using var device = deviceManager.PromptForDevice();

// Select either "Flatbed" or "Feeder"
using var item = device.FindSubItem("Feeder");

// Scan all pages in the feeder at once
item.SetProperty(WiaPropertyId.IPS_PAGES,
                 WiaPropertyValue.ALL_PAGES);

// Enable duplex scanning
item.SetProperty(WiaPropertyId.IPS_DOCUMENT_HANDLING_SELECT,
                 WiaPropertyValue.DUPLEX);

// Set up the scan
using var transfer = item.StartTransfer();
transfer.PageScanned += (sender, args) =>
{
    using (args.Stream)
    {
        var bitmap = new Bitmap(args.Stream);
        // Do something with the image
    }
};

// Do the actual scan
transfer.Download();
```

## Known issues
- Native acquisition with WIA 1.0 only works in 32-bit processes. NAPS2.Sdk provides a 32-bit worker process that can be used from 64-bit applications. Most users shouldn't be affected since WIA 2.0 has no such issue.

## Transfer compatibility

WIA 1.0 banded transfers first request `WiaImgFmt_MEMORYBMP` with `TYMED_CALLBACK`.
If the driver rejects this setup, the library retries the previous
`WiaImgFmt_BMP` with `TYMED_CALLBACK` configuration before starting the transfer.
This retains the earlier behavior for drivers that reject MemoryBMP while
allowing legacy drivers such as the Canon CanoScan 4400F to use MemoryBMP.
See Microsoft's [format/medium example](https://learn.microsoft.com/en-us/windows-hardware/drivers/image/adding-image-format-support)
and [legacy-driver transfer documentation](https://learn.microsoft.com/en-us/windows-hardware/drivers/image/data-transfer-between-windows-vista-application-and-legacy-driver).

WIA 2.0 download HRESULTs are preserved: cancellation returns `false`, and driver
errors such as `E_INVALIDARG` become `WiaException` with the original error code.
This allows callers to handle the error or retry through WIA 1.0. It does not
repair a failure inside the Windows/driver WIA 2.0 compatibility path.

Run the Windows COM regression checks with:

```sh
dotnet run --project tests/TransferRegressionTests.csproj
```

For hardware verification, scan the same full-page color document at 300 dpi
through WIA 1.0 and WIA 2.0. On the tested 4400F, WIA 1.0 succeeds with this fix;
direct WIA 2.0 still returns `E_INVALIDARG`. NAPS2's default WIA version can then
retry using WIA 1.0. Explicitly selecting WIA 2.0 disables that automatic fallback.

## Supported Platforms
- .NET Framework 4.6.2+ (Windows)
- .NET 6.0+ (Windows)

## License
Unlike most of NAPS2.Sdk which is licensed under the LGPL, NAPS2.Wia uses the more permissive MIT license. 