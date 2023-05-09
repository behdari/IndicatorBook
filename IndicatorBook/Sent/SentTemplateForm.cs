using System.Diagnostics;

namespace IndicatorBook.Sent
{
    public partial class SentTemplateForm : Form
    {
        private string _fileSharingBase;

        public SentTemplateForm(string fileSharingBase)
        {
            InitializeComponent();

            radioButtonA4.Checked = true;
            _fileSharingBase = fileSharingBase;
            textBoxAddress.Text = $"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\قالبهای آماده\\[نام فایل].docx";
        }

        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            if (textBoxFileName.Text.Trim() == "")
            {
                errorProvider1.SetError(textBoxFileName, "لطفا نام فایل را وارد کنید");
                return;
            }

            try
            {
                var currentDirectory = Directory.GetCurrentDirectory();
                FileInfo fi;

                if (radioButtonA4.Checked)
                {
                    fi = new FileInfo(currentDirectory + "\\Sent\\A4-Template.docx");
                }
                else if (radioButtonA5.Checked)
                {
                    fi = new FileInfo(currentDirectory + "\\Sent\\A5-Template.docx");
                }
                else
                {
                    throw new Exception();
                }

                Directory.CreateDirectory($"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\قالبهای آماده");

                fi.CopyTo($"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\قالبهای آماده\\{textBoxFileName.Text}.docx", false); // existing file will not be overwritten

                Process.Start(Environment.GetEnvironmentVariable("WINDIR") +
        @"\explorer.exe", $"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\قالبهای آماده");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("already exists."))
                {
                    errorProvider1.SetError(textBoxAddress, "فایل با این نام قبلا وجود داشته است");
                    return;
                }

                throw;
            }

            Close();
        }

        private void textBoxFileName_TextChanged(object sender, EventArgs e)
        {
            textBoxAddress.Text = $"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\قالبهای آماده\\{textBoxFileName.Text}.docx";
        }

        private void textBoxFileName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                buttonSubmit_Click(sender, e);
            }
        }
    }
}
