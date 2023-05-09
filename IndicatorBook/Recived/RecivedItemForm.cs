using IndicatorBook.Classes;
using IndicatorBook.Resources;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Globalization;

namespace IndicatorBook.Sent
{
    public partial class RecivedItemForm : Form
    {
        public string ItemID = "";

        bool NewItem = false;
        bool EditItem = false;
        private int _previousRowNumber;
        private string _scanFileExtension;
        private byte[] _fileByteArray;
        private string _fileSharingBase;

        public RecivedItemForm(bool newItem, bool editItem, string fileSharingBase)
        {
            InitializeComponent();
            this.tableLayoutPanel1.CellPaint += new TableLayoutCellPaintEventHandler(tableLayoutPanel1_CellPaint);

            //////////////////////
            this.NewItem = newItem;
            this.EditItem = editItem;

            if (NewItem)
                this.Text = "Add New Item";

            else if (EditItem)
                this.Text = "Edit Item";

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
                    if (textBoxRowNumber.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxRowNumber, "لطفا شماره ردیف را وارد کنید");
                        return;
                    }

                    if (!string.IsNullOrWhiteSpace(textBoxRowNumber.Text) &&
                        RecivedRepository.Instance.IsExist(Convert.ToInt32(textBoxRowNumber.Text)))
                    {
                        errorProvider1.SetError(textBoxRowNumber, "این شماره ردیف تکراری می باشد.");
                        return;
                    }

                    if (Convert.ToInt32(textBoxRowNumber.Text) <= 0)
                    {
                        errorProvider1.SetError(textBoxRowNumber, "شماره ردیف نمی تواند 0 یا کوچکتر از 0 باشد.");
                        return;
                    }

                    if (textBoxPreviousRowNumber.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxPreviousRowNumber, "لطفا شماره ردیف قبلی را وارد کنید");
                        return;
                    }

                    if (persianDatePickerDate.PersianDate.ToString().Trim() == "")
                    {
                        errorProvider1.SetError(persianDatePickerDate, "لطفا تاریخ را وارد کنید");
                        return;
                    }

                    if (textBoxLetterOwners.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxLetterOwners, "لطفا صاحبان نامه را وارد کنید");
                        return;
                    }

                    if (textBoxDescription.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxDescription, "لطفا شرح نامه های رسیده را وارد کنید");
                        return;
                    }

                    if (textBoxRecivedLetterNumber.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxRecivedLetterNumber, "لطفا شماره نامه دریافتی را وارد کنید");
                        return;
                    }

                    if (persianDatePickerRecivedLetterDate.PersianDate.ToString().Trim() == "")
                    {
                        errorProvider1.SetError(persianDatePickerRecivedLetterDate, "لطفا تاریخ نامه دریافتی را وارد کنید");
                        return;
                    }

                    if (_fileByteArray == null || _fileByteArray.Length == 0)
                    {
                        errorProvider1.SetError(buttonSelectScanFile, "لطفا فایل اسکن نامه(ها) را وارد کنید");
                        return;
                    }

                    var hasAttachment = radioButtonHasAttachment.Checked ? true
                        : radioButtonHasNotAttachment.Checked ? false
                        : throw new Exception("invalid hasAttachment value.");

                    var recivedLetter = new RecivedLetter
                    {
                        RowNumber = Convert.ToInt32(textBoxRowNumber.Text.Trim()),
                        PreviousRowNumber = Convert.ToInt32(textBoxPreviousRowNumber.Text.Trim()),
                        Date = persianDatePickerDate.PersianDate.ToString(),
                        LetterOwners = textBoxLetterOwners.Text.Trim(),
                        Description = textBoxDescription.Text.Trim(),
                        HasAttachment = hasAttachment,
                        RecivedLetterNumber = textBoxRecivedLetterNumber.Text.Trim(),
                        RecivedLetterDate = persianDatePickerRecivedLetterDate.PersianDate.ToString(),
                        ScanFile = _fileByteArray,
                        ScanFileExtension = _scanFileExtension
                    };

                    RecivedRepository.Instance.Add(recivedLetter);
                }

                #endregion

                #region edit item

                else if (EditItem)
                {
                    if (textBoxPreviousRowNumber.Text.Trim() == "")
                    {
                        errorProvider1.SetError(textBoxPreviousRowNumber, "لطفا نام خانوادگی را وارد کنید");
                        return;
                    }

                    if (!string.IsNullOrWhiteSpace(textBoxRowNumber.Text) &&
                        IsDuplicated(Convert.ToInt32(textBoxRowNumber.Text.Trim()), Convert.ToInt32(ItemID)))
                    {
                        errorProvider1.SetError(textBoxRowNumber, "این شماره تکراری (متعلق به نامه ی دیگری) می باشد.");
                        return;
                    }

                    var hasAttachment = radioButtonHasAttachment.Checked ? true
                        : radioButtonHasNotAttachment.Checked ? false
                        : throw new Exception("invalid hasAttachment value.");

                    var dbRecivedLetter = RecivedRepository.Instance.GetRecivedLetterById(Convert.ToInt32(ItemID));

                    var recivedLetter = new RecivedLetter
                    {
                        Id = Convert.ToInt32(ItemID),
                        RowNumber = Convert.ToInt32(textBoxRowNumber.Text.Trim()),
                        PreviousRowNumber = Convert.ToInt32(textBoxPreviousRowNumber.Text.Trim()),
                        Date = persianDatePickerDate.PersianDate.ToString(),
                        LetterOwners = textBoxLetterOwners.Text.Trim(),
                        Description = textBoxDescription.Text.Trim(),
                        HasAttachment = hasAttachment,
                        RecivedLetterNumber = textBoxRecivedLetterNumber.Text.Trim(),
                        RecivedLetterDate = persianDatePickerRecivedLetterDate.PersianDate.ToString(),
                        ScanFile = _fileByteArray ?? dbRecivedLetter.ScanFile,
                        ScanFileExtension = _scanFileExtension ?? dbRecivedLetter.ScanFileExtension
                    };

                    RecivedRepository.Instance.Update(recivedLetter);
                }

                #endregion

                this.Close();
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        private bool IsDuplicated(int rowNumber, int id)
        {
            var recivedLetter = RecivedRepository.Instance.GetByRowNumber(rowNumber);

            if (recivedLetter == null)
                return false;

            var isDuplicated = recivedLetter.Id != id;

            return isDuplicated;
        }

        #region

        Image ResizeImage(Image FullsizeImage, int NewWidth, int MaxHeight, bool OnlyResizeIfWider)
        {
            // Prevent using images internal thumbnail
            FullsizeImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone);
            FullsizeImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone);

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

            System.Drawing.Image NewImage = FullsizeImage.GetThumbnailImage(NewWidth, NewHeight, null, IntPtr.Zero);

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

        private void textBoxRowNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            SetNumberOnly(sender, e);
        }

        private void textBoxPreviousRowNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            SetNumberOnly(sender, e);
        }

        private void SetNumberOnly(object sender, KeyPressEventArgs e)
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

        private void buttonSelectScanFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog opnfd = new OpenFileDialog();
            opnfd.Filter = "Image Files (*.jpg;*.jpeg;.*.gif;*.png;)|*.jpg;*.jpeg;.*.gif;*.png";
            opnfd.InitialDirectory = _fileSharingBase;

            if (opnfd.ShowDialog() == DialogResult.OK)
            {
                var fileName = opnfd.FileName;

                _scanFileExtension = fileName.Split('.')[fileName.Split('.').Length - 1];
                _fileByteArray = ConvertToByteArray(fileName);
            }
        }

        public static byte[] ConvertToByteArray(string filePath)
        {
            byte[] fileByteArray = File.ReadAllBytes(filePath);
            return fileByteArray;
        }

        private void textBoxRowNumber_TextChanged(object sender, EventArgs e)
        {
            var previousRowNumber = Convert.ToInt32(textBoxRowNumber.Text) - 1;

            textBoxPreviousRowNumber.Text = previousRowNumber.ToString();
        }
    }
}
