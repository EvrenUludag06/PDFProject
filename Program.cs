using Newtonsoft.Json;
using PDFProject;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Previewer;

class Program
{
    static void Main()
    {
        RootObject rootObject = ReadJSON("report_data.json");
        GeneratePdf(rootObject);
    }

    public static RootObject ReadJSON(string filePath)
    {
        RootObject data = JsonConvert.DeserializeObject<RootObject>(File.ReadAllText(filePath));
        return data;
    }

    [Obsolete]
    public static void GeneratePdf(RootObject rootObject, bool generatePdf = false)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        Document doc = Document.Create(document =>
        {
            document.Page(page =>
            {
                page.Header().Element(Header);

                page.Content().Element(Content);

                page.Footer().Element(Footer);

                void Header(IContainer container)
                {
                    container.Row(row =>
                    {
                        row.ConstantItem(100)
                        .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");

                        row.RelativeItem()
                        .PaddingLeft(100)
                        .Column(column =>
                        {
                            column.Item()
                               .PaddingTop(30)
                               .Text("Product Content")
                               .FontSize(25)
                               .FontColor(Colors.Black)
                               .SemiBold();
                        });

                        row.ConstantItem(100)
                             .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");
                    });
                }

                void Content(IContainer container)
                {
                    var pageSizes = new List<(object name, object features)>()
                    {
                        // Bios
                        ("BIOSVersion" , rootObject.Bios.BIOSVersion),
                        ("BuildNumber", rootObject.Bios.BuildNumber),
                        ("Caption", rootObject.Bios.Caption),
                        ("CurrentLanguage", rootObject.Bios.CurrentLanguage),
                        ("Description", rootObject.Bios.Description),
                        ("Name", rootObject.Bios.Name),
                        ("SerialNumber", rootObject.Bios.SerialNumber),
                        ("Version", rootObject.Bios.Version),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),

                        // Cpu
                        ("Cpu Name", rootObject.Cpu.Name),
                        ("Caption", rootObject.Cpu.Caption),
                        ("Manufacturer", rootObject.Cpu.Manufacturer),
                        ("NumberOfCores", rootObject.Cpu.NumberOfCores),
                        ("NumberOfLogicalProcessors", rootObject.Cpu.NumberOfLogicalProcessors),
                        ("ProcessorId", rootObject.Cpu.ProcessorId),
                        ("SerialNumber", rootObject.Cpu.SerialNumber),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),

                        // MemoryCard
                        ("Memory Card 1", ""),
                        ("Capacity", rootObject.Memory.Cards[0].Capacity),
                        ("Manufacturer", rootObject.Memory.Cards[0].Manufacturer),
                        ("MemoryType", rootObject.Memory.Cards[0].MemoryType),
                        ("Name", rootObject.Memory.Cards[0].Name),
                        ("PartNumber", rootObject.Memory.Cards[0].PartNumber),
                        ("SerialNumber", rootObject.Memory.Cards[0].SerialNumber),
                        ("Speed", rootObject.Memory.Cards[0].Speed),

                        ("Memory Card 2", ""),
                        ("Capacity", rootObject.Memory.Cards[1].Capacity),
                        ("Manufacturer", rootObject.Memory.Cards[1].Manufacturer),
                        ("MemoryType", rootObject.Memory.Cards[1].MemoryType),
                        ("Name", rootObject.Memory.Cards[1].Name),
                        ("PartNumber", rootObject.Memory.Cards[1].PartNumber),
                        ("SerialNumber", rootObject.Memory.Cards[1].SerialNumber),
                        ("Speed", rootObject.Memory.Cards[1].Speed),

                        ("Memory Card 3", ""),
                        ("Capacity", rootObject.Memory.Cards[2].Capacity),
                        ("Manufacturer", rootObject.Memory.Cards[2].Manufacturer),
                        ("MemoryType", rootObject.Memory.Cards[2].MemoryType),
                        ("Name", rootObject.Memory.Cards[2].Name),
                        ("PartNumber", rootObject.Memory.Cards[2].PartNumber),
                        ("SerialNumber", rootObject.Memory.Cards[2].SerialNumber),
                        ("Speed", rootObject.Memory.Cards[2].Speed),

                        ("Memory Card 4", ""),
                        ("Capacity", rootObject.Memory.Cards[3].Capacity),
                        ("Manufacturer", rootObject.Memory.Cards[3].Manufacturer),
                        ("MemoryType", rootObject.Memory.Cards[3].MemoryType),
                        ("Name", rootObject.Memory.Cards[3].Name),
                        ("PartNumber", rootObject.Memory.Cards[3].PartNumber),
                        ("SerialNumber", rootObject.Memory.Cards[3].SerialNumber),
                        ("Speed", rootObject.Memory.Cards[3].Speed),
                        ("FreePhysicalMemory", rootObject.Memory.FreePhysicalMemory),
                        ("FreePhysicalMemoryFormat", rootObject.Memory.FreePhysicalMemoryFormat),
                        ("TotalPhysicalMemory", rootObject.Memory.TotalPhysicalMemory),
                        ("TotalPhysicalMemoryFormat", rootObject.Memory.TotalPhysicalMemoryFormat),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),

                        // NetworkInterface
                        ("NetworkInterface 1", ""),
                        ("Description", rootObject.NetworkInterface[0].Description),
                        ("DHCP", rootObject.NetworkInterface[0].DHCP),
                        ("Dns", rootObject.NetworkInterface[0].Dns[0]),
                        ("Gateways", rootObject.NetworkInterface[0].Gateways[0]),
                        ("Id", rootObject.NetworkInterface[0].Id),
                        ("IPAddresses", rootObject.NetworkInterface[0].IPAddresses[0]),
                        ("MACAddress", rootObject.NetworkInterface[0].MACAddress),
                        ("Name", rootObject.NetworkInterface[0].Name),
                        ("SubnetMasks", rootObject.NetworkInterface[0].SubnetMasks[0]),

                        ("NetworkInterface 2", ""),
                        ("Description", rootObject.NetworkInterface[1].Description),
                        ("DHCP", rootObject.NetworkInterface[1].DHCP),
                        ("Dns", rootObject.NetworkInterface[1].Dns[0]),
                        ("Gateways", ""),
                        ("Id", rootObject.NetworkInterface[1].Id),
                        ("IPAddresses", rootObject.NetworkInterface[1].IPAddresses[0]),
                        ("MACAddress", rootObject.NetworkInterface[1].MACAddress),
                        ("Name", rootObject.NetworkInterface[1].Name),
                        ("SubnetMasks", rootObject.NetworkInterface[1].SubnetMasks[0]),

                        // GPU
                        ("Gpu 1", ""),
                        ("Caption ", rootObject.Gpu[0].Caption),
                        ("Description ", rootObject.Gpu[0].Description),
                        ("DeviceID ", rootObject.Gpu[0].DeviceID),
                        ("DriverVersion ", rootObject.Gpu[0].DriverVersion),
                        ("Manufacturer ", rootObject.Gpu[0].Manufacturer),
                        ("Name ", rootObject.Gpu[0].Name),
                        ("PNPDeviceID ", rootObject.Gpu[0].PNPDeviceID),
                        ("SystemID ", rootObject.Gpu[0].SystemID),

                        ("Gpu 2", ""),
                        ("Caption ", rootObject.Gpu[1].Caption),
                        ("Description ", rootObject.Gpu[1].Description),
                        ("DeviceID ", rootObject.Gpu[1].DeviceID),
                        ("DriverVersion ", rootObject.Gpu[1].DriverVersion),
                        ("Manufacturer ", rootObject.Gpu[1].Manufacturer),
                        ("Name ", rootObject.Gpu[1].Name),
                        ("PNPDeviceID ", rootObject.Gpu[1].PNPDeviceID),
                        ("SystemID ", rootObject.Gpu[1].SystemID),
                        ("", ""),

                        // CaptureCardObject
                        ("Capture Card 1", ""),
                        ("Capture Card Name", rootObject.CaptureCard[0].CaptureCardName),

                        ("Item 1", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[0].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[0].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[0].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[0].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[0].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[0].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[0].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[0].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[0].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[0].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[0].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[0].Id),
                        ("Name", rootObject.CaptureCard[0].Items[0].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[0].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[0].ProductName),

                        ("Item 2", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[1].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[1].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[1].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[1].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[1].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[1].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[1].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[1].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[1].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[1].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[1].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[1].Id),
                        ("Name", rootObject.CaptureCard[0].Items[1].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[1].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[1].ProductName),

                        ("Item 3", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[2].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[2].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[2].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[2].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[2].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[2].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[2].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[2].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[2].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[2].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[2].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[2].Id),
                        ("Name", rootObject.CaptureCard[0].Items[2].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[2].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[2].ProductName),

                        ("Item 4", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[3].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[3].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[3].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[3].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[3].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[3].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[3].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[3].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[3].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[3].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[3].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[3].Id),
                        ("Name", rootObject.CaptureCard[0].Items[3].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[3].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[3].ProductName),

                        ("Item 5", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[4].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[4].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[4].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[4].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[4].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[4].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[4].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[4].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[4].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[4].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[4].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[4].Id),
                        ("Name", rootObject.CaptureCard[0].Items[4].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[4].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[4].ProductName),

                        ("Item 6", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[5].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[5].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[5].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[5].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[5].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[5].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[5].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[5].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[5].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[5].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[5].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[5].Id),
                        ("Name", rootObject.CaptureCard[0].Items[5].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[5].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[5].ProductName),

                        ("Item 7", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[6].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[6].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[6].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[6].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[6].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[6].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[6].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[6].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[6].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[6].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[6].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[6].Id),
                        ("Name", rootObject.CaptureCard[0].Items[6].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[6].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[6].ProductName),

                        ("Item 8", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[7].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[7].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[7].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[7].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[7].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[7].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[7].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[7].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[7].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[7].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[7].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[7].Id),
                        ("Name", rootObject.CaptureCard[0].Items[7].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[7].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[7].ProductName),

                        ("Item 9", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[8].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[8].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[8].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[8].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[8].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[8].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[8].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[8].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[8].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[8].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[8].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[8].Id),
                        ("Name", rootObject.CaptureCard[0].Items[8].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[8].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[8].ProductName),

                        ("Item 10", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[9].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[9].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[9].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[9].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[9].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[9].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[9].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[9].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[9].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[9].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[9].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[9].Id),
                        ("Name", rootObject.CaptureCard[0].Items[9].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[9].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[9].ProductName),

                        ("Item 11", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[10].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[10].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[10].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[10].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[10].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[10].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[10].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[10].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[10].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[10].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[10].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[10].Id),
                        ("Name", rootObject.CaptureCard[0].Items[10].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[10].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[10].ProductName),

                        ("Item 12", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[11].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[11].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[11].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[11].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[11].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[11].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[11].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[11].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[11].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[11].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[11].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[11].Id),
                        ("Name", rootObject.CaptureCard[0].Items[11].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[11].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[11].ProductName),

                        ("Item 13", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[12].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[12].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[12].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[12].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[12].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[12].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[12].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[12].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[12].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[12].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[12].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[12].Id),
                        ("Name", rootObject.CaptureCard[0].Items[12].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[12].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[12].ProductName),

                        ("Item 14", ""),
                        ("AudioDid", rootObject.CaptureCard[0].Items[13].AudioDid),
                        ("BoardIndex", rootObject.CaptureCard[0].Items[13].BoardIndex),
                        ("BoardSerialNo", rootObject.CaptureCard[0].Items[13].BoardSerialNo),
                        ("ChannelIndex", rootObject.CaptureCard[0].Items[13].ChannelIndex),
                        ("DriverVersion", rootObject.CaptureCard[0].Items[13].DriverVersion),
                        ("FamilyName", rootObject.CaptureCard[0].Items[13].FamilyName),
                        ("FamilyType", rootObject.CaptureCard[0].Items[13].FamilyType),
                        ("FirmwareId", rootObject.CaptureCard[0].Items[13].FirmwareId),
                        ("FirmwareName", rootObject.CaptureCard[0].Items[13].FirmwareName),
                        ("FirmwareVersion", rootObject.CaptureCard[0].Items[13].FirmwareVersion),
                        ("HardwareVersion", rootObject.CaptureCard[0].Items[13].HardwareVersion),
                        ("Id", rootObject.CaptureCard[0].Items[13].Id),
                        ("Name", rootObject.CaptureCard[0].Items[13].Name),
                        ("NameLabel", rootObject.CaptureCard[0].Items[13].NameLabel),
                        ("ProductName", rootObject.CaptureCard[0].Items[13].ProductName),

                        ("Capture Card 2", ""),
                        ("Capture Card Name", rootObject.CaptureCard[1].CaptureCardName),

                        ("Item 1", ""),
                        ("AudioStreamID", rootObject.CaptureCard[1].Items[0].AudioStreamID),
                        ("BoardNumber", rootObject.CaptureCard[1].Items[0].BoardNumber),
                        ("ConnectorIndex", rootObject.CaptureCard[1].Items[0].ConnectorIndex),
                        ("FirmwareVersion", rootObject.CaptureCard[1].Items[0].FirmwareVersion),
                        ("HardwareRevision", rootObject.CaptureCard[1].Items[0].HardwareRevision),
                        ("NameLabel", rootObject.CaptureCard[1].Items[0].NameLabel),
                        ("PCBNumber", rootObject.CaptureCard[1].Items[0].PCBNumber),
                        ("ProductName", rootObject.CaptureCard[1].Items[0].ProductName),
                        ("SerialNumber", rootObject.CaptureCard[1].Items[0].SerialNumber),
                        ("VideoStreamID", rootObject.CaptureCard[1].Items[0].VideoStreamID),

                        ("Item 2", ""),
                        ("AudioStreamID", rootObject.CaptureCard[1].Items[1].AudioStreamID),
                        ("BoardNumber", rootObject.CaptureCard[1].Items[1].BoardNumber),
                        ("ConnectorIndex", rootObject.CaptureCard[1].Items[1].ConnectorIndex),
                        ("FirmwareVersion", rootObject.CaptureCard[1].Items[1].FirmwareVersion),
                        ("HardwareRevision", rootObject.CaptureCard[1].Items[1].HardwareRevision),
                        ("NameLabel", rootObject.CaptureCard[1].Items[1].NameLabel),
                        ("PCBNumber", rootObject.CaptureCard[1].Items[1].PCBNumber),
                        ("ProductName", rootObject.CaptureCard[1].Items[1].ProductName),
                        ("SerialNumber", rootObject.CaptureCard[1].Items[1].SerialNumber),
                        ("VideoStreamID", rootObject.CaptureCard[1].Items[1].VideoStreamID),

                        ("Item 3", ""),
                        ("AudioStreamID", rootObject.CaptureCard[1].Items[2].AudioStreamID),
                        ("BoardNumber", rootObject.CaptureCard[1].Items[2].BoardNumber),
                        ("ConnectorIndex", rootObject.CaptureCard[1].Items[2].ConnectorIndex),
                        ("FirmwareVersion", rootObject.CaptureCard[1].Items[2].FirmwareVersion),
                        ("HardwareRevision", rootObject.CaptureCard[1].Items[2].HardwareRevision),
                        ("NameLabel", rootObject.CaptureCard[1].Items[2].NameLabel),
                        ("PCBNumber", rootObject.CaptureCard[1].Items[2].PCBNumber),
                        ("ProductName", rootObject.CaptureCard[1].Items[2].ProductName),
                        ("SerialNumber", rootObject.CaptureCard[1].Items[2].SerialNumber),
                        ("VideoStreamID", rootObject.CaptureCard[1].Items[2].VideoStreamID),

                        ("Item 4", ""),
                        ("AudioStreamID", rootObject.CaptureCard[1].Items[3].AudioStreamID),
                        ("BoardNumber", rootObject.CaptureCard[1].Items[3].BoardNumber),
                        ("ConnectorIndex", rootObject.CaptureCard[1].Items[3].ConnectorIndex),
                        ("FirmwareVersion", rootObject.CaptureCard[1].Items[3].FirmwareVersion),
                        ("HardwareRevision", rootObject.CaptureCard[1].Items[3].HardwareRevision),
                        ("NameLabel", rootObject.CaptureCard[1].Items[3].NameLabel),
                        ("PCBNumber", rootObject.CaptureCard[1].Items[3].PCBNumber),
                        ("ProductName", rootObject.CaptureCard[1].Items[3].ProductName),
                        ("SerialNumber", rootObject.CaptureCard[1].Items[3].SerialNumber),
                        ("VideoStreamID", rootObject.CaptureCard[1].Items[3].VideoStreamID),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                        ("", ""),
                    };

                    const int inchesToPoints = 72;

                    container
                    .PaddingTop(-29)
                    .Padding(25)
                    .MinimalBox()
                    .Border(2)
                    .Table(table =>
                    {
                        IContainer DefaultCellStyle(IContainer container, string backgroundColor)
                        {
                            return container
                                .Border(2)
                                .BorderColor(Colors.Black)
                                .Background(backgroundColor)
                                .PaddingVertical(8)
                                .PaddingHorizontal(8)
                                .AlignCenter()
                                .AlignMiddle();
                        }

                        table.ColumnsDefinition(columns =>
                        {
                            columns.RelativeColumn();
                            columns.ConstantColumn(377);
                        });

                        table.Header(header =>
                        {
                            header.Cell().RowSpan(1).Element(CellStyle).ExtendHorizontal().AlignLeft().Text("Name").Bold();
                            header.Cell().ColumnSpan(1).Element(CellStyle).Text("Features").Bold();

                            IContainer CellStyle(IContainer container) => DefaultCellStyle(container, Colors.Grey.Lighten2);
                        });

                        foreach (var page in pageSizes)
                        {
                            table.Cell().Element(CellStyle).ExtendHorizontal().AlignLeft().Text(page.name);

                            table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(page.features);

                            IContainer CellStyle(IContainer container) => DefaultCellStyle(container, Colors.White).ShowOnce();
                        }
                    });
                }

                void Footer(IContainer container)
                {
                    container.AlignCenter().Text(text =>
                    {
                        text.DefaultTextStyle(x => x.FontSize(15).Bold());
                        text.Span("Page Number ");
                        text.CurrentPageNumber();
                        text.Span(" of ");
                        text.TotalPages();
                    });
                }
            });
        });

        if (generatePdf)
        {
            doc.GeneratePdf("result.pdf");
            return;
        }

        doc.ShowInPreviewer();
    }
}