using Aspose.Words;
using Campus.Report;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StudentSemesterScoreNotification
{
    public partial class MainForm : BaseForm
    {
        public enum PrintType { 學生, 班級 };

        private PrintType _PrintType;
        private int _schoolYear, _semester;
        private List<StudentRecord> _Students;

        private BackgroundWorker _BW;
        private ReportConfiguration _Config;

        public MainForm(PrintType type)
        {
            InitializeComponent();

            _PrintType = type;

            _BW = new BackgroundWorker();
            _BW.DoWork += new DoWorkEventHandler(BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_RunWorkerCompleted);
        }

        private void BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SetForm(true);

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
                return;
            }

            Document doc = e.Result as Document;

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "新版學期成績通知單";
            save.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    doc.Save(save.FileName, SaveFormat.Doc);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            Document doc;

            //讀取報表設定
            _Config = new ReportConfiguration(Global.TemplateConfigName);

            //列印設定
            string setting = _Config.GetString("列印設定", "預設範本");
            string template_base64str = _Config.GetString(setting, string.Empty);

            if (string.IsNullOrWhiteSpace(template_base64str))
                doc = new Document(new MemoryStream(Properties.Resources.Template));
            else
                doc = new Document(new MemoryStream(Convert.FromBase64String(template_base64str)));

            if (_PrintType == PrintType.學生)
                _Students = K12.Data.Student.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
            else
                _Students = K12.Data.Student.SelectByClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource).FindAll(x => x.Status == StudentRecord.StudentStatus.一般 || x.Status == StudentRecord.StudentStatus.休學);

            DataTable dt = DataSource.GetDataTable(_Students, _schoolYear, _semester);

            doc.MailMerge.Execute(dt);

            doc.MailMerge.DeleteFields();

            e.Result = doc;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            int schoolYear, semester;

            _schoolYear = int.TryParse(K12.Data.School.DefaultSchoolYear, out schoolYear) ? schoolYear : 0;
            _semester = int.TryParse(K12.Data.School.DefaultSemester, out semester) ? semester : 0;

            for (int i = -2; i <= 2; i++)
                cboSchoolYear.Items.Add(_schoolYear + i);

            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);

            cboSchoolYear.Text = _schoolYear + "";
            cboSemester.Text = _semester + "";
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            int sy,sm;

            if (!int.TryParse(cboSchoolYear.Text, out sy))
            {
                MessageBox.Show("學年度必須為整數數字");
                return;
            }

            if (!int.TryParse(cboSemester.Text, out sm))
            {
                MessageBox.Show("學期必須為整數數字");
                return;
            }

            _schoolYear = sy;
            _semester = sm;

            if (_BW.IsBusy)
                MessageBox.Show("系統忙碌中,請稍後再試");
            else
            {
                SetForm(false);
                _BW.RunWorkerAsync();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new PrintConfigForm().ShowDialog();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new SelectTypeForm().ShowDialog();
        }

        private void SetForm(bool b)
        {
            cboSchoolYear.Enabled = b;
            cboSemester.Enabled = b;
            btnConfirm.Enabled = b;
        }

    }
}
