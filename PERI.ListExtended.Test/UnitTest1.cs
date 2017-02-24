using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PERI.ListExtender;
using iTextSharp.text.pdf;

namespace PERI.ListExtended.Test
{
    [TestClass]
    public class UnitTest1
    {
        private class Foo
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
        }

        private List<Foo> ListFoo
        {
            get
            {
                var list = new List<Foo>();
                list.Add(new Foo { FirstName = "AAA", LastName = "BBB" });
                list.Add(new Foo { FirstName = "BBB", LastName = "CCC" });
                list.Add(new Foo { FirstName = "CCC", LastName = "DDD" });
                list.Add(new Foo { FirstName = "DDD", LastName = "EEE" });
                list.Add(new Foo { FirstName = "EEE", LastName = "FFF" });

                return list;
            }
        }

        [TestMethod]
        public void ExportToCsv()
        {
            ListFoo.ExportToCsv();           
        }

        [TestMethod]
        public void ExportToCsvAsFile()
        {
            ListFoo.ExportToCsv(@"e:\a.csv");
        }

        [TestMethod]
        public void ExportToPdfAsFile()
        {
            var layout = new PdfLayout();
            layout.Headers.Add(new PdfHeader() { Text = "Hello", FontSize = 9, Alignment = PdfPCell.ALIGN_CENTER });
            layout.Headers.Add(new PdfHeader() { Text = "This is my subheader", FontSize = 7, Alignment = PdfPCell.ALIGN_CENTER });
            layout.Headers.Add(new PdfHeader() { Text = "This is my tiny header", FontSize = 5, Alignment = PdfPCell.ALIGN_CENTER });
                        
            layout.Footers.Add(new PdfFooter() { Text = "Thank you.", FontSize = 9 });
            layout.Footers.Add(new PdfFooter() { Text = "This is my subfooter", FontSize = 7 });
            layout.Footers.Add(new PdfFooter() { Text = "This is my tiny footer", FontSize = 5 });

            ListFoo.ExportToPdf(@"e:\Sample.pdf", layout);
        }
    }
}
