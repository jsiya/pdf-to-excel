using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Linq;

namespace CommandPatternPDFandEXCEL
{
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
    }

    public interface ITableActionCommand
    {
        MemoryStream Execute();
        string GetFileName();
    }

    public class ExcelFile<T> : ITableActionCommand
    {
        private readonly List<T> _list;

        public ExcelFile(List<T> list)
        {
            _list = list;
        }

        public MemoryStream Execute()
        {
            var wb = new XLWorkbook();
            var ds = new DataSet();

            ds.Tables.Add(GetTable());
            wb.Worksheets.Add(ds);

            var excelMemory = new MemoryStream();
            wb.SaveAs(excelMemory);

            return excelMemory;
        }

        public string GetFileName()
        {
            return $"{typeof(T).Name}.xlsx";
        }

        private DataTable GetTable()
        {
            var table = new DataTable();
            var type = typeof(T);

            type.GetProperties()
                .ToList()
                .ForEach(x => table.Columns.Add(x.Name, x.PropertyType));

            _list.ForEach(x =>
            {
                var values = type.GetProperties()
                                 .Select(propertyInfo => propertyInfo
                                 .GetValue(x, null))
                                 .ToArray();

                table.Rows.Add(values);
            });

            return table;
        }
    }

    public class PdfFile<T> : ITableActionCommand
    {
        private readonly List<T> _list;

        public PdfFile(List<T> list)
        {
            _list = list;
        }

        public MemoryStream Execute()
        {
            var pdfMemory = new MemoryStream();
            using (var doc = new Document())
            {
                var pdfWriter = PdfWriter.GetInstance(doc, pdfMemory);
                doc.Open();

                // tablein gorunusu
                var table = new PdfPTable(4); // 4 basliq var deye
                table.WidthPercentage = 100;

                var type = typeof(T);
                //her basliga property qoy
                var headers = type.GetProperties().Take(4).Select(x => x.Name).ToList();
                foreach (var header in headers) //burda basliqlara design
                {
                    table.AddCell(new PdfPCell(new Phrase(header, FontFactory.GetFont(FontFactory.TIMES_BOLDITALIC, 14, BaseColor.RED))) //basligin fontu
                    {
                        HorizontalAlignment = Element.ALIGN_CENTER,
                        VerticalAlignment = Element.ALIGN_MIDDLE,
                        BackgroundColor = BaseColor.PINK 
                    });
                }

                // productlar
                foreach (var item in _list)
                {
                    var values = headers.Select(header => type.GetProperty(header).GetValue(item)?.ToString() ?? "");
                    foreach (var value in values)
                    {
                        table.AddCell(new PdfPCell(new Phrase(value, FontFactory.GetFont(FontFactory.TIMES, 10)))
                        {
                            HorizontalAlignment = Element.ALIGN_CENTER,
                            VerticalAlignment = Element.ALIGN_MIDDLE
                        });
                    }
                }

                doc.Add(table);
                doc.Close();
            }

            return pdfMemory;
        }

        public string GetFileName()
        {
            return $"{typeof(T).Name}.pdf";
        }
    }


    class FileCreateInvoker
    {
        private List<ITableActionCommand> tableActionCommands = new List<ITableActionCommand>();

        public void AddCommand(ITableActionCommand tableActionCommand)
        {
            tableActionCommands.Add(tableActionCommand);
        }

        public void CreateFiles()
        {
            using (var zipMemoryStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(zipMemoryStream, ZipArchiveMode.Create, true))
                {
                    foreach (var command in tableActionCommands)
                    {
                        var commandFileName = command.GetFileName();
                        var entry = archive.CreateEntry(commandFileName);

                        using (var entryStream = entry.Open())
                        {
                            var fileData = command.Execute();
                            var fileBytes = fileData.ToArray();
                            entryStream.Write(fileBytes, 0, fileBytes.Length);
                        }
                    }
                }
                File.WriteAllBytes("output.zip", zipMemoryStream.ToArray());
            }
        }
    }


    class Program
    {
        static void Main()
        {
            var products = Enumerable.Range(1, 30).Select(index =>
                new Product
                {
                    Id = index,
                    Name = $"Product {index}",
                    Price = index + 100,
                    Stock = index
                }
            ).ToList();

            ExcelFile<Product> excelFile = new ExcelFile<Product>(products);
            PdfFile<Product> pdfFile = new PdfFile<Product>(products);

            FileCreateInvoker invoker = new FileCreateInvoker();
            invoker.AddCommand(excelFile);
            invoker.AddCommand(pdfFile);

            invoker.CreateFiles();
            Console.WriteLine("Files created successfully.");
        }
    }
}
