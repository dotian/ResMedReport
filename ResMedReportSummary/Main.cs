using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
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
        private static string strDateFormat = "yyyy-MM-dd";

        public Main()
        {
            InitializeComponent();
        }

        #region Event

        private void Main_Load(object sender, EventArgs e)
        {
            txtCreateDate.Text = DateTime.Now.ToString(strDateFormat);
            //txtName.Text = "张某某";
            //txtId.Text = "1234567";
            //txtGender.Text = "男";
            //txtBedInfo.Text = "18区40床xxxx";
            //txtAge.Text = "20";
            //txtDiagnosis.Text = "1. 重度睡眠呼吸暂停综合征。\n2. 夜间重度低氧血症。\n3. 其他";
            //txtSuggestion.Text = "1. 随访，积极治疗原发疾病。\n2. 必要时予以气道正压通气呼吸机治疗。\n3. 其他";
            //txtBirth.Text = "1967年1月1日";

            DirectoryInfo dir = new DirectoryInfo(ReportSavePath);
            if (!dir.Exists)
            {
                dir.Create();
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("是否确认打印？", "确认打印", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    var report = GetReport();
                    string filePath = Path.Combine(ReportSavePath, report.Name + "_" + report.Id + ".pdf");
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

                dgvReports.Rows.Clear();

                string regex = string.IsNullOrEmpty(id) ? name + "_.*" : ".*_" + id + ".pdf";

                DirectoryInfo dir = new DirectoryInfo(ReportSavePath);
                var reports = dir.GetFiles().Where(f => Regex.IsMatch(f.Name, regex));
                foreach (var report in reports)
                {
                    dgvReports.Rows.Add(report.Name, report.CreationTime.ToString(strDateFormat), report.FullName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("系统报错：" + ex.Message, "系统报错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                var filePath = dgvReports.SelectedRows[0].Cells["colPath"].Value.ToString();
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("系统报错：" + ex.Message, "系统报错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            report.CreateDate = txtCreateDate.Text != string.Empty ? txtCreateDate.Text : DateTime.Now.ToString(strDateFormat);
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
