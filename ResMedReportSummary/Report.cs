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

        public static string FlowEvaluationPeriodKey = "流量评估周期";
        public string FlowEvaluationPeriod;

        public static string SpO2EvaluationPeriodKey = "血氧饱和度评估周期";
        public string SpO2EvaluationPeriod;

        public List<string[]> Details = new List<string[]>();

        private static BaseFont bfSun = BaseFont.CreateFont("SIMSUN.TTC,1", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        private static Font defaultFont = new Font(bfSun, 11);

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
            AddNewLine(document);

            document.Add(GenerateLineSeparate());

            document.Add(GenerateHeader3());

            document.Add(GenerateBasicInfoTable());

            document.Add(GenerateDetailAnalysisTitle());
            document.Add(GenerateDetailAnalysisPara("（一）", Enumerable.Range(0, 2)));
            document.Add(GenerateDetailAnalysisPara("（二）", Enumerable.Range(2, 3)));
            document.Add(GenerateDetailAnalysisPara("（三）", Enumerable.Range(5, 6)));
            document.Add(GenerateDetailAnalysisPara("（四）", Enumerable.Range(11, 4)));

            document.Add(GenerateDiagnosisPara());
            document.Add(GenerateSuggestionPara());

            document.Close();

            Process.Start(filePath);
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
            var p = new Paragraph("上海市中医药大学附属普陀医院", new Font(bfSun, 12));
            p.Alignment = Element.ALIGN_CENTER;
            return p;
        }

        private IElement GenerateHeader2()
        {
            var p = new Paragraph("上海市普陀区中心医院", new Font(bfSun, 12));
            p.Alignment = Element.ALIGN_CENTER;
            return p;
        }

        private IElement GenerateLineSeparate()
        {
            return new LineSeparator();
        }

        private IElement GenerateHeader3()
        {
            var p = new Paragraph("睡眠监测报告", new Font(bfSun, 14, Font.BOLD));
            p.Alignment = Element.ALIGN_CENTER;
            p.SpacingBefore = 5f;
            p.SpacingAfter = 10f;
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

        private IElement GenerateDetailAnalysisTitle()
        {
            var p = new Paragraph(string.Format("分析结果（{0}：{1}/{2}：{3}）",
                    FlowEvaluationPeriodKey,
                    FlowEvaluationPeriod,
                    SpO2EvaluationPeriodKey,
                    SpO2EvaluationPeriod),
                new Font(bfSun, 11, Font.BOLD));
            p.SpacingBefore = 10f;
            p.IndentationLeft = 54f;
            return p;
        }

        private IElement GenerateDetailAnalysisPara(string title, IEnumerable<int> range)
        {
            var p = CreateDetailAnalysisParagraghTitle(title);
            AddDetailLine(p, range);
            return p;
        }

        private IElement GenerateDiagnosisPara()
        {
            var p = CreateDetailAnalysisParagraghTitle("（五）诊断");
            var pContent = new Paragraph(Diagnosis, defaultFont);
            pContent.IndentationLeft = 30f;
            p.Add(pContent);
            return p;
        }

        private IElement GenerateSuggestionPara()
        {
            var p = CreateDetailAnalysisParagraghTitle("（六）建议");
            var pContent = new Paragraph(Suggestion, defaultFont);
            pContent.IndentationLeft = 30f;
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

        private Paragraph CreateDetailAnalysisParagraghTitle(string content)
        {
            var p = new Paragraph(content, new Font(bfSun, 12, Font.BOLD));
            p.IndentationLeft = 50f;
            p.SpacingBefore = 10f;
            return p;
        }

        private void AddNewLine(Document doc)
        {
            doc.Add(new Paragraph(" "));
        }

        private void AddDetailLine(Paragraph para, IEnumerable<int> range)
        {
            foreach (var index in range)
            {
                var p = new Paragraph(FormatDetailString(Details[index]), defaultFont);
                p.IndentationLeft = 30f;
                para.Add(p);
            }
        }

        private string FormatDetailString(string[] details)
        {
            if (string.IsNullOrWhiteSpace(details[1]))
            {
                return string.Format("{0}：{1}", details[0], details[2]);
            }
            else
            {
                return string.Format("{0}：{2}（{1}）", details[0], details[1], details[2]);
            }
        }
    }
}