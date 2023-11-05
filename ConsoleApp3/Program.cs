using System;
using System.Data.SqlClient;
using System.Linq;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
namespace ConsoleApp3
{
    internal class Program
    {

        static void AddTable(Section section, Color primaryColor, string[] colSizes, string[] colHeadings, string[][] rowDatas, int fontSize = 8, string fontName = "Noto Sans", bool summationRow = false)
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
            var primaryColorHeading = Color.FromArgb(150, (byte)primaryColor.R, (byte)primaryColor.G, (byte)primaryColor.B);
            for (int i = 0; i < colHeadings.Length && i < colSizes.Length; i++)
            {
                var headingPara = headingRow.Cells[i].AddParagraph();
                headingPara.AddFormattedText(colHeadings[i], headingFont);
                headingRow.Cells[i].Shading.Color = primaryColorHeading;
            }
            headingRow.Borders.Bottom.Width = 2;

            var primaryColorLight = Color.FromArgb(60, (byte)primaryColor.R, (byte)primaryColor.G, (byte)primaryColor.B);


            var dataFont = new Font();
            dataFont.Bold = false;
            dataFont.Name = fontName;
            dataFont.Size = fontSize;


            for (int i = 0; i < rowDatas.Length; i++)
            {
                string[] rowData = rowDatas[i];
                var row = table.AddRow();


                if ((i + 1) % 2 == 0) row.Shading.Color = primaryColorLight;
                if (i == rowDatas.Length - 1 && summationRow)
                {
                    row.Borders.Top.Width = 2;
                }
                for (int j = 0; j < colSizes.Length && j < rowData.Length; j++)
                {
                    var para = row.Cells[j].AddParagraph();
                    para.AddFormattedText(rowData[j], i < rowDatas.Length - 1 || !summationRow ? dataFont : headingFont);
                }
            }



            table.Borders.Color = Color.FromArgb(255, 0, 0, 0);
            table.Borders.Width = 1.0;

            table.Format.Alignment = ParagraphAlignment.Center;
            table.Format.SpaceAfter = "0.2cm";
            table.Format.SpaceBefore = "0.2cm";
        }

        static void Main(string[] args)
        {
            var document = new Document();
            document.DefaultPageSetup.RightMargin = "0.8cm";
            document.DefaultPageSetup.LeftMargin = "0.8cm";
            document.DefaultPageSetup.TopMargin = "1.4cm";
            document.DefaultPageSetup.TopMargin = "1.4cm";
            document.DefaultPageSetup.PageWidth = "29.7cm";
            document.DefaultPageSetup.PageHeight = "21cm";

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


            var renderer = new PdfDocumentRenderer();
            renderer.Document = document;

            string[] colSizes = {
                "3cm",
                "1.5cm",
                "1.3cm",
                "3cm",
                "2.5cm",
                "2.5cm",
                "1.5cm",
                "1.5cm",
                "1.3cm",
                "1.3cm",
                "1.3cm",
                "1.4cm",
                "1.5cm",
                "1.5cm",
                "1.5cm",
                "1.8cm",
            };
            string[] colHeads = {
                "Client Address",
                "Type of Customer",
                "Test Lab",
                "Type of Test",
                "Name of Test",
                "Work Order Number",
                "Work Order Date",
                "Due Date",
                "No. of Material",
                "Rate Per Material",
                "Total Test Charge",
                "Test Report Number",
                "Test Report Date",
                "Receiving Date",
                "Dispatch Date",
                "Is Accreditted"
            };
            string[][] csvData = new string[][]
            {
        new string[]
        {
            "123 Mahatma Gandhi Road, Delhi, India","Regular","ABC Labs","Chemical Analysis","Chemical Test 1","WO12345","2023-11-05","2023-11-15","5","50","250","TRN123","2023-11-20","2023-11-10","2023-11-16","Accredited"
        },
        new string[]
        {
            "456 Jawaharlal Nehru Street, Mumbai, India","Priority","XYZ Labs","Metallurgical Analysis","Metal Test 1","WO54321","2023-11-06","2023-11-16","3","75","225","TRN678","2023-11-21","2023-11-11","2023-11-17","Non Accredited"
        },
        new string[]
        {
            "789 Subhas Chandra Bose Avenue, Kolkata, India","Regular","LMN Labs","Microbiological Analysis","Bacterial Test 1","WO98765","2023-11-07","2023-11-17","10","30","300","TRN999","2023-11-22","2023-11-12","2023-11-18","Accredited"
        },
        new string[]
        {
            "555 Rajendra Prasad Road, Chennai, India","Regular","ABC Labs","Chemical Analysis","Chemical Test 2","WO11111","2023-11-08","2023-11-18","7","60","420","TRN456","2023-11-23","2023-11-13","2023-11-19","Accredited"
        },
        new string[]
        {
            "777 Sardar Patel Street, Bangalore, India","Priority","XYZ Labs","Metallurgical Analysis","Metal Test 2","WO22222","2023-11-09","2023-11-19","4","80","320","TRN789","2023-11-24","2023-11-14","2023-11-20","Non Accredited"
        },
        new string[]
        {
            "999 Lal Bahadur Shastri Avenue, Hyderabad, India","Regular","LMN Labs","Microbiological Analysis","Bacterial Test 2","WO33333","2023-11-10","2023-11-20","6","40","240","TRN101","2023-11-25","2023-11-15","2023-11-21","Accredited"
        },
        new string[]
        {
            "222 Jawaharlal Nehru Road, Pune, India","Priority","ABC Labs","Chemical Analysis","Chemical Test 3","WO44444","2023-11-11","2023-11-21","8","70","560","TRN222","2023-11-26","2023-11-16","2023-11-22","Non Accredited"
        },
        new string[]
        {
            "444 Indira Gandhi Street, Ahmedabad, India","Regular","XYZ Labs","Metallurgical Analysis","Metal Test 3","WO55555","2023-11-12","2023-11-22","5","90","450","TRN333","2023-11-27","2023-11-17","2023-11-23","Accredited"
        },
        new string[]
        {
            "666 Morarji Desai Avenue, Jaipur, India","Priority","LMN Labs","Microbiological Analysis","Bacterial Test 3","WO66666","2023-11-13","2023-11-23","9","50","450","TRN444","2023-11-28","2023-11-18","2023-11-24","Non Accredited"
        },
        new string[]
        {
            "888 Atal Bihari Vajpayee Road, Chandigarh, India","Regular","ABC Labs","Chemical Analysis","Chemical Test 4","WO77777","2023-11-14","2023-11-24","6","60","360","TRN555","2023-11-29","2023-11-19","2023-11-25","Accredited"
        },
        new string[]
        {
            "111 B.R. Ambedkar Street, Lucknow, India","Priority","XYZ Labs","Metallurgical Analysis","Metal Test 4","WO88888","2023-11-15","2023-11-25","3","70","210","TRN666","2023-11-30","2023-11-20","2023-11-26","Non Accredited"
        },
        new string[]
        {
            "333 Sarojini Naidu Avenue, Bhopal, India","Regular","LMN Labs","Microbiological Analysis","Bacterial Test 4","WO99999","2023-11-16","2023-11-26","8","40","320","TRN777","2023-12-01","2023-11-21","2023-11-27","Accredited"
        },
        new string[]
        {
            "555 Rani Lakshmi Bai Street, Indore, India","Regular","ABC Labs","Chemical Analysis","Chemical Test 5","WO10101","2023-11-17","2023-11-27","4","50","200","TRN888","2023-12-02","2023-11-22","2023-11-28","Non Accredited"
        },
        new string[]
        {
            "777 Chhatrapati Shivaji Road, Kanpur, India","Priority","XYZ Labs","Metallurgical Analysis","Metal Test 5","WO11111","2023-11-18","2023-11-28","7","75","525","TRN999","2023-12-03","2023-11-23","2023-11-29","Accredited"
        },
        new string[]
        {
            "999 Bhagat Singh Avenue, Nagpur, India","Regular","LMN Labs","Microbiological Analysis","Bacterial Test 5","WO12121","2023-11-19","2023-11-29","5","60","300","TRN101","2023-12-04","2023-11-24","2023-11-30","Non Accredited"
        },
        new string[]
        {
            "222 Lala Lajpat Rai Road, Coimbatore, India","Regular","ABC Labs","Chemical Analysis","Chemical Test 6","WO13131","2023-11-20","2023-11-30","10","80","800","TRN111","2023-12-05","2023-11-25","2023-12-01","Accredited"
        },
        new string[]
        {
            "444 Kasturba Gandhi Street, Visakhapatnam, India","Priority","XYZ Labs","Metallurgical Analysis","Metal Test 6","WO14141","2023-11-21","2023-12-01","6","70","420","TRN121","2023-12-06","2023-11-26","2023-12-02","Non Accredited"
        },
        new string[]
        {
            "666 Ram Manohar Lohia Road, Madurai, India","Regular","LMN Labs","Microbiological Analysis","Bacterial Test 6","WO15151","2023-11-22","2023-12-02","3","90","270","TRN131","2023-12-07","2023-11-27","2023-12-03","Accredited"
        },
            };

            var totalNoOfMaterial = 0;
            var totalTestCharge = 0;

            for (int i = 0; i < csvData.Length; i++)
            {
                try
                {
                    var noOfMaterial = Int32.Parse(csvData[i][8]);
                    totalNoOfMaterial += noOfMaterial;
                }
                catch { }
                try
                {
                    var testCharge = Int32.Parse(csvData[i][10]);
                    totalTestCharge += testCharge;
                }
                catch { }
            }

            csvData = csvData.Append(new string[] {
            "Total", "", "", "", "", "", "", "", totalNoOfMaterial.ToString(), "", totalTestCharge.ToString(),"", "", "", "", ""
        }).ToArray();
            AddTable(section, primaryColor, colSizes, colHeads, csvData, fontSize: 8, summationRow: true);

            renderer.PrepareRenderPages();

            renderer.RenderDocument();
            renderer.PdfDocument.Save("C:\\tmp\\report.pdf");


        }
    }
}
