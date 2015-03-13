using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ResMedSummaryReport
{
    public class Report
    {
        public string Name;
        public string Id;
        public string Gender;
        public string BedInfo;
        public string CreateDate;
        public string Birth;
        public string Age;
        public string Diagnosis;
        public string Suggestion;

        private static BaseFont bfSun = BaseFont.CreateFont("SIMSUN.TTC,1", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        private static Font defaultFont = new Font(bfSun, 14);

        private Document doc;
        private Document document
        {
            get
            {
                if (doc == null)
                {
                    doc = new Document(PageSize.A4);
                }
                return doc;
            }
            set
            {
                doc = value;
            }
        }

        private string FilePath;

        public void Save(string filePath)
        {
            this.FilePath = filePath;
            PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
            document.Open();

            document.Add(GenerateHeader1());
            document.Add(GenerateHeader2());
            document.Add(GenerateLineSeparate());
            document.Add(GenerateHeader3());
            document.Add(GenerateBasicInfoTable());
            AddNewLine(document);
            document.Add(GenerateDiagnosisPara());
            AddNewLine(document);
            AddNewLine(document);
            document.Add(GenerateSuggestionPara());

            document.Close();
        }

        public void Print()
        {
            var adobeexe = ConfigurationManager.AppSettings["AdobeExe"];
            Process p = new Process();
            p.StartInfo.FileName = adobeexe;
            p.StartInfo.Arguments = "/p \"" + FilePath + "\"";
            p.Start();
        }

        private IElement GenerateHeader1()
        {
            var p = new Paragraph("上海市中医药大学附属普陀医院", new Font(bfSun, 18));
            p.Alignment = Element.ALIGN_CENTER;
            return p;
        }

        private IElement GenerateHeader2()
        {
            var p = new Paragraph("上海市普陀区中心医院", new Font(bfSun, 18));
            p.Alignment = Element.ALIGN_CENTER;
            return p;
        }

        private IElement GenerateLineSeparate()
        {
            var p = new Paragraph(" ");
            p.Add(new LineSeparator());
            p.Add(new Paragraph(" "));
            return p;
        }

        private IElement GenerateHeader3()
        {
            var p = new Paragraph("睡眠监测报告", new Font(bfSun, 18, Font.BOLD));
            p.Alignment = Element.ALIGN_CENTER;
            p.Add(new Paragraph(" "));
            return p;
        }

        private IElement GenerateBasicInfoTable()
        {
            PdfPTable table = new PdfPTable(4);
            table.TotalWidth = 800f;
            table.SetWidths(new float[] { 150f, 250f, 150f, 250f });

            table.Rows.Add(new PdfPRow(new PdfPCell[]{
                 CreateLabelCell("姓名："),
                 CreateValueCell(this.Name),
                 CreateLabelCell("性别："),
                 CreateValueCell(this.Gender)
            }));
            table.Rows.Add(new PdfPRow(new PdfPCell[]{
                 CreateLabelCell("年龄："),
                 CreateValueCell(this.Age),
                 CreateLabelCell("出生日期："),
                 CreateValueCell(this.Birth)
            }));
            table.Rows.Add(new PdfPRow(new PdfPCell[]{
                 CreateLabelCell("床位信息："),
                 CreateValueCell(this.BedInfo),
                 CreateLabelCell("住院号："),
                 CreateValueCell(this.Id)
            }));
            table.Rows.Add(new PdfPRow(new PdfPCell[]{
                 CreateLabelCell("记录日期："),
                 CreateValueCell(this.CreateDate),
                 CreateLabelCell(null),
                 CreateValueCell(null)
            }));

            return table;
        }

        private IElement GenerateDiagnosisPara()
        {
            var p = new Paragraph("一、诊断", new Font(bfSun, 16, Font.BOLD));
            var pContent = new Paragraph(Diagnosis, defaultFont);
            pContent.SpacingBefore = 10f;
            pContent.IndentationLeft = 40f;
            p.Add(pContent);
            return p;
        }

        private IElement GenerateSuggestionPara()
        {
            var p = new Paragraph("二、建议", new Font(bfSun, 16, Font.BOLD));
            var pContent = new Paragraph(Suggestion, defaultFont);
            pContent.SpacingBefore = 10f;
            pContent.IndentationLeft = 40f;
            p.Add(pContent);
            return p;
        }

        private PdfPCell CreateLabelCell(string content)
        {
            var cell = new PdfPCell();
            cell.AddElement(new Chunk(content == null ? " " : content, defaultFont));
            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Border = PdfPCell.NO_BORDER;
            return cell;
        }

        private PdfPCell CreateValueCell(string content)
        {
            var cell = new PdfPCell();
            cell.AddElement(new Chunk(content == null ? " " : content, defaultFont));
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Border = PdfPCell.NO_BORDER;
            return cell;
        }

        private void AddNewLine(Document doc)
        {
            doc.Add(new Paragraph(" "));
        }
    }
}
