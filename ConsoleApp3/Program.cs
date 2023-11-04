using System;
using System.Data.SqlClient;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
namespace ConsoleApp3
{
    internal class Program
    {

        static void AddTable(Section section, Color primaryColor, string[] colSizes, string[] colHeadings, string[][] rowDatas, int fontSize = 8, string fontName = "Noto Sans")
        {

            var table = section.AddTable();

            for (int i = 0; i < colSizes.Length; i++)
            {
                table.AddColumn(colSizes[i]);
            }

            var headingRow = table.AddRow();

            var headingFont = new Font();
            headingFont.Bold = true;
            headingFont.Name = fontName;
            headingFont.Size = fontSize;
            for (int i = 0; i < colHeadings.Length && i < colSizes.Length; i++)
            {
                var headingPara = headingRow.Cells[i].AddParagraph();
                headingPara.AddFormattedText(colHeadings[i], headingFont);
                headingRow.Cells[i].Shading.Color = primaryColor;
            }

            var primaryColorLight = Color.FromArgb(60, (byte)primaryColor.R, (byte)primaryColor.G, (byte)primaryColor.B);


            var dataFont = new Font();
            dataFont.Bold = false;
            dataFont.Name = fontName;
            dataFont.Size = fontSize;

            for (int i = 0; i < rowDatas.Length; i++)
            {
                string[] rowData = rowDatas[i];
                var row = table.AddRow();


                for (int j = 0; j < colSizes.Length && j < rowData.Length; j++)
                {
                    var para = row.Cells[j].AddParagraph();
                    para.AddFormattedText(rowData[j], dataFont);
                    if ((i + 1) % 2 == 0) row.Cells[j].Shading.Color = primaryColorLight;
                }

            }

            table.Borders.Color = primaryColor;
            table.Borders.Width = 1.0;

            table.Format.Alignment = ParagraphAlignment.Center;
            table.Format.SpaceAfter = "0.2cm";
            table.Format.SpaceBefore = "0.2cm";
        }

        static void Main(string[] args)
        {
            var document = new Document();
            document.DefaultPageSetup.RightMargin = "2cm";
            document.DefaultPageSetup.LeftMargin = "2cm";

            var section = document.AddSection();


            var primaryColor = new Color(0, 123, 197);

            var footer = section.Footers.Primary.AddParagraph();
            var footerText = footer.AddText("Generated on " + DateTime.Now.ToString("dd/MM/yyyy \\a\\t HH:mm:ss UTCzzz"));
            var paragraph = section.AddParagraph();

            var f = new Font("Noto Sans");
            f.Size = 25;
            paragraph.AddFormattedText("Test Reports", f);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            //paragraph.AddText("Filter criteria: {Insert used filters while searching here}");
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();


            var renderer = new PdfDocumentRenderer();
            renderer.Document = document;

            string[] colSizes = { "2cm", "2cm" };
            string[] colHeads = { "One", "Two" };
            string[][] rowDatas = {
            new string[] {"123", "354m", "55vm"}
            ,new string[] {"123", "354m", "55vm"}
            ,new string[] {"123", "354m", "55vm"}
            ,new string[] {"123", "354m", "55vm"}
            , new string[] {"123",}
            };
            AddTable(section, primaryColor, colSizes, colHeads, rowDatas, fontSize: 12);

            renderer.PrepareRenderPages();

            renderer.RenderDocument();
            renderer.PdfDocument.Save("C:\\tmp\\hello.pdf");


        }
    }
}
