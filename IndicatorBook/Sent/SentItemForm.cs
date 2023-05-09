using IndicatorBook.Classes;
using IndicatorBook.Resources;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;

namespace IndicatorBook.Sent
{
    public partial class SentItemForm : Form
    {
        public string ItemID = "";

        bool NewItem = false;
        bool EditItem = false;
        private byte[] _fileByteArray;
        private string _wordFileExtension;
        private string _fileSharingBase;

        public SentItemForm(bool newItem, bool editItem, string fileSharingBase)
        {
            InitializeComponent();
            this.tableLayoutPanel1.CellPaint += new TableLayoutCellPaintEventHandler(tableLayoutPanel1_CellPaint);

            //////////////////////
            NewItem = newItem;
            EditItem = editItem;

            if (NewItem)
                Text = "Add New Item";

            else if (EditItem)
                Text = "Edit Item";

            radioButtonHasNotAttachment.Checked = true;
            _fileSharingBase = fileSharingBase;
        }

        void tableLayoutPanel1_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
        {
            try
            {
                if (e.Row % 2 == 0)
                {
                    Graphics g = e.Graphics;
                    Rectangle r = e.CellBounds;
                    g.FillRectangle(new SolidBrush(Color.FromArgb(225, 225, 225)), r);
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                errorProvider1.Clear();

                #region add new item

                if (NewItem)
                {
                    if (textBoxPamphleteer.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxPamphleteer, "لطفا جزوه دان را وارد کنید");
                        return;
                    }

                    if (textBoxFile.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxFile, "لطفا پرونده را وارد کنید");
                        return;
                    }

                    if (textBoxNumber.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxNumber, "لطفا شماره را وارد کنید");
                        return;
                    }

                    if (textBoxTitle.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxTitle, "لطفا عنوان را وارد کنید");
                        return;
                    }

                    if (textBoxDescription.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxDescription, "لطفا توضیحات را وارد کنید");
                        return;
                    }

                    if (textBoxNextNumber.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxNextNumber, "لطفا شماره بعدی را وارد کنید");
                        return;
                    }

                    if (persianDatePickerDate.PersianDate.ToString().Trim() == "")
                    {
                        errorProvider1.SetError(persianDatePickerDate, "لطفا تاریخ را وارد کنید");
                        return;
                    }

                    if (_fileByteArray == null || _fileByteArray.Length == 0)
                    {
                        errorProvider1.SetError(buttonSelectWordFile, "لطفا فایل نامه(ها) را وارد کنید");
                        return;
                    }

                    if (SentRepository.Instance.IsExist(textBoxFile.Text, textBoxPamphleteer.Text, Convert.ToInt32(textBoxNumber.Text)))
                    {
                        errorProvider1.SetError(textBoxNumber, "این شماره تکراری می باشد.");
                        return;
                    }

                    var hasAttachment = radioButtonHasAttachment.Checked ? true
                        : radioButtonHasNotAttachment.Checked ? false
                        : throw new Exception("invalid hasAttachment value.");

                    var sentLetter = new SentLetter
                    {
                        File = textBoxFile.Text.Trim(),
                        Pamphleteer = textBoxPamphleteer.Text.Trim(),
                        Number = Convert.ToInt32(textBoxNumber.Text.Trim()),
                        Title = textBoxTitle.Text.Trim(),
                        Description = textBoxDescription.Text.Trim(),
                        NextNumber = Convert.ToInt32(textBoxNextNumber.Text.Trim()),
                        Date = persianDatePickerDate.PersianDate.ToString(),
                        HasAttachment = hasAttachment,
                        WordFile = _fileByteArray,
                        WordFileExtension = _wordFileExtension
                    };

                    SentRepository.Instance.Add(sentLetter);
                }

                #endregion

                #region edit item

                else if (EditItem)
                {
                    if (textBoxFile.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxFile, "لطفا نام خانوادگی را وارد کنید");
                        return;
                    }

                    if (!string.IsNullOrWhiteSpace(textBoxDescription.Text) &&
                        IsDuplicated(textBoxFile.Text.Trim(),
                                     textBoxPamphleteer.Text.Trim(),
                                     Convert.ToInt32(textBoxNumber.Text.Trim()),
                                     Convert.ToInt32(ItemID)))
                    {
                        errorProvider1.SetError(textBoxDescription, "این شماره تکراری (متعلق به نامه ی دیگری) می باشد.");
                        return;
                    }

                    var hasAttachment = radioButtonHasAttachment.Checked ? true
                        : radioButtonHasNotAttachment.Checked ? false
                        : throw new Exception("invalid hasAttachment value.");

                    var dbSentLetter = SentRepository.Instance.GetSentLetterById(Convert.ToInt32(ItemID));

                    var sentLetter = new SentLetter
                    {
                        Id = Convert.ToInt32(ItemID),
                        File = textBoxPamphleteer.Text.Trim(),
                        Pamphleteer = textBoxFile.Text.Trim(),
                        Number = Convert.ToInt32(textBoxNumber.Text.Trim()),
                        Title = textBoxTitle.Text.Trim(),
                        Description = textBoxDescription.Text.Trim(),
                        NextNumber = Convert.ToInt32(textBoxNextNumber.Text.Trim()),
                        HasAttachment = hasAttachment,
                        Date = persianDatePickerDate.PersianDate.ToString(),
                        WordFile = _fileByteArray ?? dbSentLetter.WordFile,
                        WordFileExtension = _wordFileExtension ?? dbSentLetter.WordFileExtension
                    };

                    SentRepository.Instance.Update(sentLetter);
                }

                #endregion

                Close();
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        private bool IsDuplicated(string file, string pamphleteer, int number, int id)
        {
            var sentLetter = SentRepository.Instance.GetSentLetter(file, pamphleteer, number);

            if (sentLetter == null)
                return false;

            var isDuplicated = sentLetter.Id != id;

            return isDuplicated;
        }

        #region

        Image ResizeImage(Image FullsizeImage, int NewWidth, int MaxHeight, bool OnlyResizeIfWider)
        {
            // Prevent using images internal thumbnail
            FullsizeImage.RotateFlip(RotateFlipType.Rotate180FlipNone);
            FullsizeImage.RotateFlip(RotateFlipType.Rotate180FlipNone);

            if (OnlyResizeIfWider)
            {
                if (FullsizeImage.Width <= NewWidth)
                {
                    NewWidth = FullsizeImage.Width;
                }
            }

            int NewHeight = FullsizeImage.Height * NewWidth / FullsizeImage.Width;
            if (NewHeight > MaxHeight)
            {
                // Resize with height instead
                NewWidth = FullsizeImage.Width * MaxHeight / FullsizeImage.Height;
                NewHeight = MaxHeight;
            }

            Image NewImage = FullsizeImage.GetThumbnailImage(NewWidth, NewHeight, null, IntPtr.Zero);

            // Clear handle to original file so that we can overwrite it if necessary
            FullsizeImage.Dispose();

            // Save resized picture
            return NewImage;
        }

        string ImageToBase64String(Image image, ImageFormat format)
        {
            MemoryStream memory = new MemoryStream();
            image.Save(memory, format);
            string base64 = Convert.ToBase64String(memory.ToArray());
            memory.Close();

            return base64;
        }

        Image ImageFromBase64String(string base64)
        {
            MemoryStream memory = new MemoryStream(Convert.FromBase64String(base64));
            Image result = Image.FromStream(memory);
            memory.Close();

            return result;
        }

        #endregion

        private void buttonSelectWordFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog opnfd = new OpenFileDialog();
            opnfd.Filter = @"Word File Or Compressed File (.docx ,.doc, .rar, .zip)|*.docx;*.doc;*.rar;*.zip";
            opnfd.InitialDirectory = _fileSharingBase;

            if (opnfd.ShowDialog() == DialogResult.OK)
            {
                var fileName = opnfd.FileName;

                _wordFileExtension = fileName.Split('.')[fileName.Split('.').Length - 1];
                _fileByteArray = ConvertToByteArray(fileName);
            }
        }

        public static byte[] ConvertToByteArray(string filePath)
        {
            byte[] fileByteArray = File.ReadAllBytes(filePath);
            return fileByteArray;
        }

        private void textBoxNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            CheckForNumberOnlyTextBox(sender, e);
        }

        private void textBoxNextNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            CheckForNumberOnlyTextBox(sender, e);
        }

        private void CheckForNumberOnlyTextBox(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void textBoxNumber_TextChanged(object sender, EventArgs e)
        {
            var nextNumber = Convert.ToInt32(textBoxNumber.Text) + 1;

            textBoxNextNumber.Text = nextNumber.ToString();
        }
    }
}
