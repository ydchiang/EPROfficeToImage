using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Spire.Doc;
using Spire.Pdf;
using Spire.Presentation;
using Spire.Xls;
using ImageType = Spire.Doc.Documents.ImageType;

namespace EPROfficeToImage
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!Directory.Exists("Images"))
            {
                Directory.CreateDirectory("Images");
            }

            // Powerpoint 
            Presentation p = new Presentation();
            p.LoadFromFile(@"C:\MyFolder\code\EPROfficeToImage\Chap4.pptx");

            for (int i = 0; i < p.Slides.Count; i++)
            {
                p.Slides[i].SaveAsImage().Save("Images\\PPT-" + i + ".png", ImageFormat.Png);
            }

            // Office
            Document doc = new Document();
            doc.LoadFromFile(@"C:\MyFolder\code\EPROfficeToImage\UART.docx");

            for (int i = 0; i < doc.PageCount; i++)
            {
                Image img = doc.SaveToImages(i,ImageType.Bitmap);
                img.Save("Images\\Word-" + i + ".png", ImageFormat.Png);
            }

            // PDF
            PdfDocument pdf = new PdfDocument();
            pdf.LoadFromFile(@"C:\MyFolder\code\EPROfficeToImage\Serial.pdf");

            for (int i = 0; i < pdf.Pages.Count; i++)
            {
                try
                {
                    pdf.SaveAsImage(i).Save("Images\\PDF-" + i + ".png", ImageFormat.Png);
                }
                catch { }   // Catch over 3 pages exception
            }

            // Excel
            Workbook wb = new Workbook();
            wb.LoadFromFile(@"C:\MyFolder\code\EPROfficeToImage\ArduinoUno.xlsx");

            for (int i = 0; i < wb.Worksheets.Count; i++)
            {
                wb.Worksheets[i].SaveToImage("Images\\XLS-i" + ".png", ImageFormat.Png);
            }
        }
    }
}
