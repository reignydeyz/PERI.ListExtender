using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Reflection;
using System.IO;

namespace PERI.ListExtender
{
    public static class ListExtended
    {
        /// <summary>
        /// Checker for property type
        /// Will only allow these datatypes
        /// </summary>
        /// <param name="prop"></param>
        /// <returns></returns>
        private static bool IsAllowedPropertyInfo(PropertyInfo prop)
        {
            return prop.PropertyType == typeof(String)
                    || prop.PropertyType == typeof(int)
                    || prop.PropertyType == typeof(int?)
                    || prop.PropertyType == typeof(long)
                    || prop.PropertyType == typeof(long?)
                    || prop.PropertyType == typeof(decimal)
                    || prop.PropertyType == typeof(decimal?)
                    || prop.PropertyType == (typeof(bool))
                    || prop.PropertyType == (typeof(bool?))
                    || prop.PropertyType == (typeof(DateTime))
                    || prop.PropertyType == (typeof(DateTime?))
                    || prop.PropertyType == (typeof(DateTimeOffset))
                    || prop.PropertyType == (typeof(DateTimeOffset?));
        }

        /// <summary>
        /// Gets the CSV as StringBuilder
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns>StringBuilder</returns>
        public static StringBuilder ExportToCsv<T>(this List<T> list)
        {
            StringBuilder fileContent = new StringBuilder();

            foreach (var prop in typeof(T).GetProperties())
            {
                if (!IsAllowedPropertyInfo(prop))
                    continue;

                fileContent.Append(prop.Name + ",");
            }

            fileContent.Append("\r\n");

            foreach (var rec in list)
            {
                foreach (var field in rec.GetType().GetProperties())
                {
                    fileContent.Append(field.GetValue(rec, null) + ",");
                }

                fileContent.Append("\r\n");
            }

            return fileContent;
        }

        /// <summary>
        /// Saves the CSV to a folder
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="filePath">The complete path with the desired file name</param>
        public static void ExportToCsv<T>(this List<T> list, string filePath)
        {
            var sb = list.ExportToCsv();

            File.WriteAllText(filePath, sb.ToString());
        }

        public static void ExportToPdf<T>(this List<T> list, string filePath, PdfLayout layout)
        {
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));

            if (layout.IsLandscape)
                document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());

            // Open the document  
            document.Open();

            iTextSharp.text.Font fontbold = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA_BOLD, layout.FontSize);
            iTextSharp.text.Font font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, layout.FontSize);

            PdfPTable table = new PdfPTable(typeof(T).GetProperties().Count());
                        
            // Setting columns
            List<float> l = new List<float>();
            foreach (var prop in typeof(T).GetProperties())
                l.Add(4f);

            float[] a = l.ToArray();
            table.SetWidths(a);
            table.WidthPercentage = 100;
            
            // Adding columns
            foreach (var prop in typeof(T).GetProperties())
            {
                table.AddCell(new Phrase(prop.Name, fontbold));
            }

            // Adding the data
            foreach (var rec in list)
            {
                int colindex = 0;
                foreach (var field in rec.GetType().GetProperties())
                {
                    table.AddCell(new Phrase(field.GetValue(rec, null).ToString(), font5));
                    colindex++;
                }
            }

            // Add headers
            foreach (var header in layout.Headers)
            {
                Paragraph para = new Paragraph(header.Text + Environment.NewLine, new Font(Font.FontFamily.HELVETICA, header.FontSize));
                para.Alignment = header.Alignment;
                if (header.Text != null && header.Text != "")
                    document.Add(para);
            }

            // Adjust header
            if (layout.Headers.Count > 0)
            {
                var header = layout.Headers[layout.Headers.Count - 1];
                Paragraph para = new Paragraph(Environment.NewLine, new Font(Font.FontFamily.HELVETICA, header.FontSize));
                document.Add(para);
            }

            document.Add(table);

            // Add footers
            foreach (var footer in layout.Footers)
            {
                Paragraph para = new Paragraph(footer.Text + Environment.NewLine, new Font(Font.FontFamily.HELVETICA, footer.FontSize));
                para.Alignment = footer.Alignment;
                if (footer.Text != null && footer.Text != "")
                    document.Add(para);
            }

            document.Close();
        }
    }
}
