using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp;
using System.IO;
using System.Drawing;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace Spectrum_Test
{
    class ExportPDF
    {
        public static void export_PDF()
        {
            var doc1 = new iTextSharp.text.Document(new iTextSharp.text.Rectangle(593f, 842f), 30, 30, 20, 20);
            PdfWriter writer = PdfWriter.GetInstance(doc1, new FileStream(@"Report\" + "1234.pdf", FileMode.Create));
            //設定文字
            BaseFont bfEnglish = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
            BaseFont chBaseFont = BaseFont.CreateFont(@"fonts\chinese\BiauKai.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font myFont = new iTextSharp.text.Font(bfEnglish, 26, 1);//26級粗體
            iTextSharp.text.Font mychFont = new iTextSharp.text.Font(chBaseFont, 18, 1);//12級粗體
            doc1.Open();
            Paragraph para = new Paragraph();
            para.Alignment = Element.ALIGN_CENTER;
            Phrase a = new Phrase("Spectrochips", myFont);
            para.Add(a);
            para.Add("\n\n\n");
            PdfPTable basicTable = new PdfPTable(2);
            basicTable.AddCell(new PdfPCell(new Phrase("機台", mychFont)));
            basicTable.AddCell(new PdfPCell(new Phrase("line2,cell2")));
            para.Add(basicTable);
            doc1.Add(para);
            doc1.Close();
        }
    }
}
