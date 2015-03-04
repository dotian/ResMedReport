using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ResMedSummaryReport
{
    public partial class Main : Form
    {
        private static string ReportSavePath = Path.Combine(Environment.CurrentDirectory, "Reports");

        public Main()
        {
            InitializeComponent();
        }

        #region Event

        private void Main_Load(object sender, EventArgs e)
        {
            txtCreateDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("是否确认打印？", "确认打印", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    var report = GetReport();
                    string filePath = Path.Combine(ReportSavePath, report.Name + "_" + report.Id);
                    report.Save(filePath);
                    report.Print();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("系统报错：" + ex.Message, "系统报错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string name = txtName_Search.Text;
                string id = txtId_Search.Text;
                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(id))
                {
                    throw new Exception("姓名和住院号需至少填入一项。");
                }

                string regex = string.IsNullOrEmpty(id) ? name + "_.*" : ".*_" + id + "$";

                DirectoryInfo dir = new DirectoryInfo(ReportSavePath);
                var reports = dir.GetFiles().Where(f => Regex.IsMatch(f.Name, regex));

            }
            catch (Exception ex)
            {
                MessageBox.Show("系统报错：" + ex.Message, "系统报错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {

        }

        #endregion

        private Report GetReport()
        {
            var report = new Report();
            report.Name = txtName.Text;
            report.Id = txtId.Text;

            if (string.IsNullOrEmpty(report.Name))
            {
                throw new Exception("姓名不得为空。");
            }

            if (string.IsNullOrEmpty(report.Id))
            {
                throw new Exception("住院号不得为空。");
            }

            report.Age = txtAge.Text;
            report.Birth = txtBirth.Text;
            report.BedInfo = txtBedInfo.Text;
            report.CreateDate = txtCreateDate.Text != string.Empty ? txtCreateDate.Text : DateTime.Now.ToString("yyyy-MM-dd");
            report.Gender = txtGender.Text;

            report.Diagnosis = txtDiagnosis.Text;
            if (report.Diagnosis.Length > 400)
            {
                throw new Exception("诊断内容过长。");
            }

            report.Suggestion = txtSuggestion.Text;
            if (report.Suggestion.Length > 400)
            {
                throw new Exception("建议内容过长。");
            }

            return report;
        }

    }
}
