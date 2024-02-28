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
    public static Dictionary<string, object> GetValues(object obj)
    {
        return obj.GetType().GetProperties().ToDictionary(p => p.Name, p => p.GetValue(obj));
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
                page.Header().Element(BiosHeader);

                page.Content().Element(BiosContent);

                page.Footer().Element(Footer);

                void BiosHeader(IContainer container)
                {
                    container.Row(row =>
                    {
                        row.ConstantItem(100)
                        .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");

                        row.RelativeItem()
                        .PaddingLeft(120)
                        .Column(column =>
                        {
                            column.Item()
                               .PaddingTop(30)
                               .Text("Bios Content")
                               .FontSize(25)
                               .FontColor(Colors.Black)
                               .SemiBold();
                        });

                        row.ConstantItem(100)
                             .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");
                    });
                }

                void BiosContent(IContainer container)
                {
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

                        foreach (var item in GetValues(rootObject.Bios))
                        {
                            table.Cell().Element(CellStyle).ExtendHorizontal().AlignLeft().Text(item.Key);

                            table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(item.Value);

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
            document.Page(page =>
            {
                page.Header().Element(CpuHeader);

                page.Content().Element(CpuContent);

                page.Footer().Element(Footer);

                void CpuHeader(IContainer container)
                {
                    container.Row(row =>
                    {
                        row.ConstantItem(100)
                        .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");

                        row.RelativeItem()
                        .PaddingLeft(125)
                        .Column(column =>
                        {
                            column.Item()
                               .PaddingTop(30)
                               .Text("Cpu Content")
                               .FontSize(25)
                               .FontColor(Colors.Black)
                               .SemiBold();
                        });

                        row.ConstantItem(100)
                             .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");
                    });
                }

                void CpuContent(IContainer container)
                {

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

                        foreach (var item in GetValues(rootObject.Cpu))
                        {
                            table.Cell().Element(CellStyle).ExtendHorizontal().AlignLeft().Text(item.Key);

                            table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(item.Value);

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
            document.Page(page =>
            {
                page.Header().Element(MemoryHeader);

                page.Content().Element(MemoryContent);

                page.Footer().Element(Footer);

                void MemoryHeader(IContainer container)
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
                               .Text("Memory Content")
                               .FontSize(25)
                               .FontColor(Colors.Black)
                               .SemiBold();
                        });

                        row.ConstantItem(100)
                             .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");
                    });
                }

                void MemoryContent(IContainer container)
                {
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

                        IContainer CellStyle(IContainer container) => DefaultCellStyle(container, Colors.White).ShowOnce();

                        foreach (var card in rootObject.Memory.Cards)
                        {
                            foreach (var item in GetValues(card))
                            {
                                table.Cell().Element(CellStyle).ExtendHorizontal().AlignLeft().Text(item.Key);
                                table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(item.Value);
                            }
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
            document.Page(page =>
            {
                page.Header().Element(NetworkInterfaceHeader);

                page.Content().Element(NetworkInterfaceContent);

                page.Footer().Element(Footer);

                void NetworkInterfaceHeader(IContainer container)
                {
                    container.Row(row =>
                    {
                        row.ConstantItem(100)
                        .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");

                        row.RelativeItem()
                        .PaddingLeft(46)
                        .Column(column =>
                        {
                            column.Item()
                               .PaddingTop(30)
                               .Text("Network Interface Content")
                               .FontSize(25)
                               .FontColor(Colors.Black)
                               .SemiBold();
                        });

                        row.ConstantItem(100)
                             .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");
                    });
                }

                void NetworkInterfaceContent(IContainer container)
                {

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

                        IContainer CellStyle(IContainer container) => DefaultCellStyle(container, Colors.White).ShowOnce();

                        foreach (var networkInterface in rootObject.NetworkInterface)
                        {
                            foreach (var item in GetValues(networkInterface))
                            {
                                table.Cell().Element(CellStyle).ExtendHorizontal().AlignLeft().Text(item.Key);

                                if (item.Value is string[] stringArray)
                                {
                                    string concatenatedValues = string.Join(", ", stringArray);

                                    table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(concatenatedValues);
                                }
                                else
                                {
                                    table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(item.Value.ToString());
                                }
                            }
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
            document.Page(page =>
            {
                page.Header().Element(GpuHeader);

                page.Content().Element(GpuContent);

                page.Footer().Element(Footer);

                void GpuHeader(IContainer container)
                {
                    container.Row(row =>
                    {
                        row.ConstantItem(100)
                        .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");

                        row.RelativeItem()
                        .PaddingLeft(120)
                        .Column(column =>
                        {
                            column.Item()
                               .PaddingTop(30)
                               .Text("Gpu Content")
                               .FontSize(25)
                               .FontColor(Colors.Black)
                               .SemiBold();
                        });

                        row.ConstantItem(100)
                             .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");
                    });
                }

                void GpuContent(IContainer container)
                {

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
                            columns.ConstantColumn(380);
                        });

                        table.Header(header =>
                        {
                            header.Cell().RowSpan(1).Element(CellStyle).ExtendHorizontal().AlignLeft().Text("Name").Bold();
                            header.Cell().ColumnSpan(1).Element(CellStyle).Text("Features").Bold();

                            IContainer CellStyle(IContainer container) => DefaultCellStyle(container, Colors.Grey.Lighten2);
                        });

                        IContainer CellStyle(IContainer container) => DefaultCellStyle(container, Colors.White).ShowOnce();

                        foreach (var gpu in rootObject.Gpu)
                        {
                            foreach (var item in GetValues(gpu))
                            {
                                table.Cell().Element(CellStyle).ExtendHorizontal().AlignLeft().Text(item.Key);

                                if (item.Value is string[] stringArray)
                                {
                                    string concatenatedValues = string.Join(", ", stringArray);

                                    table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(concatenatedValues);
                                }
                                else
                                {
                                    table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(item.Value.ToString());
                                }
                            }
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
            document.Page(page =>
            {
                page.Header().Element(CaptureCardHeader);

                page.Content().Element(CaptureCardContent);

                page.Footer().Element(Footer);

                void CaptureCardHeader(IContainer container)
                {
                    container.Row(row =>
                    {
                        row.ConstantItem(100)
                        .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");

                        row.RelativeItem()
                        .PaddingLeft(75)
                        .Column(column =>
                        {
                            column.Item()
                               .PaddingTop(30)
                               .Text("Capture Card Content")
                               .FontSize(25)
                               .FontColor(Colors.Black)
                               .SemiBold();
                        });

                        row.ConstantItem(100)
                             .Image("C:\\Users\\Evren\\Desktop\\1681044504416.jpeg");
                    });
                }

                void CaptureCardContent(IContainer container)
                {

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

                        IContainer CellStyle(IContainer container) => DefaultCellStyle(container, Colors.White).ShowOnce();

                        foreach (var captureCard in rootObject.CaptureCard)
                        {
                            foreach (var item in GetValues(captureCard))
                            {
                                table.Cell().Element(CellStyle).ExtendHorizontal().AlignLeft().Text(item.Key);

                                if (item.Value is string[] stringArray)
                                {
                                    foreach (var arrayItem in stringArray)
                                    {
                                        table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(arrayItem);
                                    }
                                }
                                else
                                {
                                    table.Cell().PaddingHorizontal(5).PaddingTop(7).BorderBottom(2).Text(item.Value.ToString());
                                }
                            }
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